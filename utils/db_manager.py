"""
Database Manager
"""

from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.dialects.mysql import LONGBLOB
from sqlalchemy import inspect, text
from enum import Enum
import logging

# Configure logging
logger = logging.getLogger(__name__)

db = SQLAlchemy()

# ENUMS
class JobType(str, Enum):
    TEXT = 'Text'
    FILE = 'File'
    CONVERSATIONAL = 'Conversational'

class JobStatus(str, Enum):
    PROCESSING = 'Processing'
    SUCCESS = 'Success'
    FAILED = 'Failed'

# MODELS
class Settings(db.Model):
    """System settings for configuration."""
    __tablename__ = 'Settings'

    VariableName = db.Column(db.String(100), primary_key=True, nullable=False)
    VariableValue = db.Column(db.String(255), nullable=False)
    Description = db.Column(db.Text, nullable=False)

    @staticmethod
    def get_upload_mode():
        """Get current file upload mode."""
        setting = Settings.query.filter_by(VariableName='UPLOAD_FILE_MODE').first()
        if not setting or not setting.VariableValue:
            return 'BOTH'

        value = setting.VariableValue.strip().upper()
        return value if value in {'FILE', 'DB', 'BOTH'} else 'BOTH'

    @staticmethod
    def set_upload_mode(mode):
        """Update upload mode (FILE, DB, or BOTH)."""
        normalized_mode = (mode or '').strip().upper()
        if normalized_mode not in {'FILE', 'DB', 'BOTH'}:
            raise ValueError("Mode must be FILE, DB, or BOTH")

        setting = Settings.query.filter_by(VariableName='UPLOAD_FILE_MODE').first()
        if setting:
            setting.VariableValue = normalized_mode
        else:
            setting = Settings(
                VariableName='UPLOAD_FILE_MODE',
                VariableValue=normalized_mode,
                Description='Controls where uploaded files are stored.'
            )
            db.session.add(setting)

        db.session.commit()


class User(db.Model):
    """User accounts"""
    __tablename__ = 'Users'
    
    Id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    FirstName = db.Column(db.String(100), nullable=False)
    LastName = db.Column(db.String(100), nullable=False)
    Email = db.Column(db.String(255), unique=True, nullable=False, index=True)
    Password = db.Column(db.String(255), nullable=False)
    IsActive = db.Column(db.Boolean, default=True)
    TokenVersion = db.Column(db.Integer, default=0)
    CreatedDate = db.Column(db.DateTime, default=datetime.utcnow)
    ModifiedDate = db.Column(db.DateTime, onupdate=datetime.utcnow)

    jobs = db.relationship('ProcessingJob', backref='user', lazy=True, cascade='all, delete-orphan')
    
    def set_password(self, password: str):
        self.Password = generate_password_hash(password, method='pbkdf2:sha256')
    
    def check_password(self, password: str) -> bool:
        return check_password_hash(self.Password, password)
    
    def invalidate_tokens(self):
        self.TokenVersion += 1
    
    def to_dict(self):
        return {
            'id': self.Id,
            'firstName': self.FirstName,
            'lastName': self.LastName,
            'email': self.Email,
            'isActive': self.IsActive,
            'createdDate': self.CreatedDate.isoformat() if self.CreatedDate else None
        }

class ProcessingJob(db.Model):
    """User's document processing projects (Conversations/Threads)"""
    __tablename__ = 'ProcessingJobs'
    
    Id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    JobName = db.Column(db.String(255), nullable=False)
    IsFavorite = db.Column(db.Boolean, default=False)
    UserId = db.Column(db.Integer, db.ForeignKey('Users.Id', ondelete='CASCADE'), nullable=False, index=True)
    CreatedDate = db.Column(db.DateTime, default=datetime.utcnow)
    
    history = db.relationship('ProcessingJobHistory', backref='job', lazy=True, cascade='all, delete-orphan')
    
    def get_last_activity(self):
        """Used for the sidebar to show the most recent version's summary"""
        return ProcessingJobHistory.query.filter_by(ProcessJobId=self.Id)\
                 .order_by(ProcessingJobHistory.CreatedDate.desc()).first()

    def to_dict(self):
        return {
            'id': self.Id,
            'jobName': self.JobName,
            'userId': self.UserId,
            'createdDate': self.CreatedDate.isoformat() if self.CreatedDate else None,
            'processCount': len(self.history)
        }

class ProcessingJobHistory(db.Model):
    """Individual processing instances (Messages/Versions)"""
    __tablename__ = 'ProcessingJobsHistory'
    
    Id = db.Column(db.Integer, primary_key=True)
    ProcessJobId = db.Column(db.Integer, db.ForeignKey('ProcessingJobs.Id'), nullable=False)
    JobType = db.Column(db.Enum(JobType), nullable=False)
    
    UploadFileName = db.Column(db.String(255))
    UploadFileData = db.Column(LONGBLOB, nullable=True)
    UploadFileServerPath = db.Column(db.String(500), nullable=True)
    OutputFileName = db.Column(db.String(255))
    OutputFileData = db.Column(LONGBLOB, nullable=True)
    OutputFileServerPath = db.Column(db.String(500), nullable=True)
    
    FontFamily = db.Column(db.String(50), default='Calibri')
    FontSize = db.Column(db.Integer, default=11)
    IncludeCover = db.Column(db.Boolean, default=False)
    IncludeTOC = db.Column(db.Boolean, default=False)
    
    ProcessingCount = db.Column(db.Integer, default=1, nullable=False)
    ProcessingTime = db.Column(db.Float, nullable=True) # [NEW] Persist real duration
    CreatedBy = db.Column(db.Integer, db.ForeignKey('Users.Id'), nullable=False)
    ModifiedBy = db.Column(db.Integer, db.ForeignKey('Users.Id'), nullable=False)
    
    Status = db.Column(db.Enum(JobStatus), default=JobStatus.PROCESSING)
    ErrorMessage = db.Column(db.Text)
    
    CreatedDate = db.Column(db.DateTime, default=datetime.utcnow)
    ModifiedDate = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    @property
    def UploadFilePath(self):
        """Backward-compatible alias for the server path field."""
        return self.UploadFileServerPath

    @UploadFilePath.setter
    def UploadFilePath(self, value):
        self.UploadFileServerPath = value

    @property
    def OutputFilePath(self):
        """Backward-compatible alias for the server path field."""
        return self.OutputFileServerPath

    @OutputFilePath.setter
    def OutputFilePath(self, value):
        self.OutputFileServerPath = value

    def to_dict(self, include_file_data=False):
        """Serializes history for the chat window UI"""
        job_type_val = self.JobType.value if hasattr(self.JobType, 'value') else self.JobType
        status_val = self.Status.value if hasattr(self.Status, 'value') else self.Status
        
        data = {
            'id': self.Id,
            'jobId': self.ProcessJobId,
            'type': job_type_val,
            'processingCount': self.ProcessingCount,
            'fontFamily': self.FontFamily,
            'fontSize': self.FontSize,
            'includeCover': self.IncludeCover,
            'includeTOC': self.IncludeTOC,
            'status': status_val,
            'errorMessage': self.ErrorMessage,
            'processingTime': self.ProcessingTime,
            'timestamp': self.CreatedDate.strftime("%I:%M %p") if self.CreatedDate else '',
            'createdDate': self.CreatedDate.isoformat() if self.CreatedDate else None,
            'createdBy': self.CreatedBy,
            'modifiedBy': self.ModifiedBy,
            'uploadFileName': self.UploadFileName,
            'rawText': None,
            'downloadUrl': f'/api/process/download/{self.Id}'
        }

        # Handle raw text decoding for text-based jobs
        if self.UploadFileData and job_type_val == 'Text':
            try:
                data['rawText'] = self.UploadFileData.decode('utf-8')
            except (UnicodeDecodeError, AttributeError):
                data['rawText'] = str(self.UploadFileData)

        return data

# Helper Function
def init_db(app):
    """Initialize database with Flask app"""
    db.init_app(app)
    with app.app_context():
        db.create_all()
        _ensure_runtime_columns()
        _ensure_default_settings()
        logger.info("Database tables created/updated successfully")


def _ensure_runtime_columns():
    """Best-effort schema patching for existing environments without migrations."""
    inspector = inspect(db.engine)

    if 'ProcessingJobsHistory' in inspector.get_table_names():
        columns = {col['name'] for col in inspector.get_columns('ProcessingJobsHistory')}
        ddl = []
        if 'UploadFileServerPath' not in columns:
            ddl.append("ALTER TABLE ProcessingJobsHistory ADD COLUMN UploadFileServerPath VARCHAR(500) NULL")
        if 'OutputFileServerPath' not in columns:
            ddl.append("ALTER TABLE ProcessingJobsHistory ADD COLUMN OutputFileServerPath VARCHAR(500) NULL")
        if 'CreatedBy' not in columns:
            ddl.append("ALTER TABLE ProcessingJobsHistory ADD COLUMN CreatedBy INT NULL")
        if 'ModifiedBy' not in columns:
            ddl.append("ALTER TABLE ProcessingJobsHistory ADD COLUMN ModifiedBy INT NULL")
        if ddl:
            with db.engine.begin() as conn:
                for stmt in ddl:
                    conn.execute(text(stmt))
                if 'UploadFilePath' in columns and 'UploadFileServerPath' not in columns:
                    conn.execute(
                        text(
                            "UPDATE ProcessingJobsHistory "
                            "SET UploadFileServerPath = UploadFilePath "
                            "WHERE UploadFilePath IS NOT NULL AND UploadFileServerPath IS NULL"
                        )
                    )
                if 'OutputFilePath' in columns and 'OutputFileServerPath' not in columns:
                    conn.execute(
                        text(
                            "UPDATE ProcessingJobsHistory "
                            "SET OutputFileServerPath = OutputFilePath "
                            "WHERE OutputFilePath IS NOT NULL AND OutputFileServerPath IS NULL"
                        )
                    )
                if 'CreatedBy' not in columns:
                    conn.execute(
                        text(
                            "UPDATE ProcessingJobsHistory h "
                            "JOIN ProcessingJobs j ON j.Id = h.ProcessJobId "
                            "SET h.CreatedBy = j.UserId "
                            "WHERE h.CreatedBy IS NULL"
                        )
                    )
                if 'ModifiedBy' not in columns:
                    conn.execute(
                        text(
                            "UPDATE ProcessingJobsHistory h "
                            "JOIN ProcessingJobs j ON j.Id = h.ProcessJobId "
                            "SET h.ModifiedBy = j.UserId "
                            "WHERE h.ModifiedBy IS NULL"
                        )
                    )


def _ensure_default_settings():
    """Create required default settings for first-run environments."""
    if not Settings.query.filter_by(VariableName='UPLOAD_FILE_MODE').first():
        db.session.add(
            Settings(
                VariableName='UPLOAD_FILE_MODE',
                VariableValue='Both',
                Description='Controls where uploaded files are stored.'
            )
        )
        db.session.commit()


def get_upload_file_mode(default='BOTH'):
    """Return the configured upload mode, falling back to the provided default."""
    try:
        mode = Settings.get_upload_mode()
    except Exception as exc:
        logger.warning("Falling back to default upload mode due to settings lookup error: %s", exc)
        mode = (default or 'BOTH').strip().upper()

    return mode if mode in {'FILE', 'DB', 'BOTH'} else 'BOTH'
