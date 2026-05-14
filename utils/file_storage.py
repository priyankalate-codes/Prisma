"""
File Storage Manager - Handles FILE/DB/BOTH modes
"""
import os
import tempfile
from io import BytesIO
from urllib.parse import quote

import requests
from werkzeug.utils import secure_filename

from config import BASE_UPLOAD_URL
from utils.db_manager import Settings, User


class FileStorageManager:
    """Manages file storage based on UPLOAD_FILE_MODE setting."""

    def __init__(self, upload_folder, base_upload_url=BASE_UPLOAD_URL):
        self.upload_folder = upload_folder
        self.base_upload_url = base_upload_url
        self.upload_endpoint = base_upload_url.rstrip('/')
        self.remote_root = os.path.basename(os.path.normpath(upload_folder)) or 'uploads'
        os.makedirs(upload_folder, exist_ok=True)

    def upload_file(self, file_path, file_name_to_save, target_dir):
        """Upload a file using the PHP upload endpoint."""
        if not file_path or not os.path.exists(file_path):
            raise FileNotFoundError("The file to upload does not exist.", file_path)

        with open(file_path, 'rb') as file_stream:
            response = requests.post(
                self.upload_endpoint,
                files={'file': file_stream},
                data={'targetDir': target_dir},
                timeout=30
            )

        if not response.ok:
            raise RuntimeError(f"Upload failed: {response.reason}")

        return True

    def _build_target_dir(self, user_id):
        try:
            user = User.query.get(user_id)
            if user:
                full_name = f"{user.FirstName} {user.LastName}".strip()
                safe_folder = secure_filename(full_name).strip('_')
                if safe_folder:
                    return f"{self.remote_root}/{safe_folder}"
        except Exception:
            # If the DB is temporarily unavailable, fall back to a generic folder.
            pass
        return f"{self.remote_root}/user_{user_id}"

    def _build_remote_file_url(self, target_dir, filename):
        encoded_parts = [quote(part.strip('/')) for part in target_dir.split('/') if part.strip('/')]
        encoded_name = quote(filename)
        return f"{self.base_upload_url.rstrip('/')}/{'/'.join(encoded_parts)}/{encoded_name}"

    def save_file(self, file_data, filename, user_id, job_id):
        """
        Save file based on current mode

        Args:
            file_data: bytes or BytesIO object
            filename: original filename
            user_id: user ID for folder organization
            job_id: job ID for unique naming

        Returns:
            dict with keys: server_path (str or None), db_bytes (bytes or None)
        """
        mode = Settings.get_upload_mode()

        if isinstance(file_data, BytesIO):
            file_bytes = file_data.getvalue()
        else:
            file_bytes = file_data

        result = {
            'server_path': None,
            'db_bytes': None
        }

        safe_filename = secure_filename(filename)
        unique_filename = f"{job_id}_{safe_filename}"
        target_dir = self._build_target_dir(user_id)

        if mode in ['FILE', 'BOTH']:
            temp_path = None
            try:
                with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
                    tmp_file.write(file_bytes)
                    temp_path = tmp_file.name

                self.upload_file(temp_path, unique_filename, target_dir)
                server_path = self._build_remote_file_url(target_dir, unique_filename)
                result['server_path'] = server_path
                print(f"File saved to server: {server_path}")
            except Exception as exc:
                if mode == 'FILE':
                    raise RuntimeError(f"Error uploading file: {exc}") from exc
                print(f"Warning: server upload failed, continuing with DB storage only: {exc}")
            finally:
                if temp_path and os.path.exists(temp_path):
                    os.remove(temp_path)

        if mode in ['DB', 'BOTH']:
            result['db_bytes'] = file_bytes
            print(f"File saved to database ({len(file_bytes)} bytes)")

        print(
            f"[FileStorage] Mode: {mode}, "
            f"Server: {result['server_path'] is not None}, "
            f"DB: {result['db_bytes'] is not None}"
        )
        return result

    def retrieve_file(self, server_path, db_bytes, filename):
        """
        Retrieve file with priority: Server -> Database -> Error

        Args:
            server_path: path or URL to file on server (can be None)
            db_bytes: file bytes from database (can be None)
            filename: original filename for error messages

        Returns:
            BytesIO object with file data

        Raises:
            FileNotFoundError if file not found anywhere
        """
        if server_path:
            if server_path.startswith(('http://', 'https://')):
                try:
                    response = requests.get(server_path, timeout=30)
                    if response.ok:
                        print(f"File retrieved from server: {server_path}")
                        return BytesIO(response.content)
                except requests.RequestException:
                    pass
            elif os.path.exists(server_path):
                print(f"File retrieved from server: {server_path}")
                with open(server_path, 'rb') as f:
                    return BytesIO(f.read())

        if db_bytes:
            print(f"File retrieved from database ({len(db_bytes)} bytes)")
            return BytesIO(db_bytes)

        raise FileNotFoundError(
            f"File not found: {filename} (checked server and database)"
        )
