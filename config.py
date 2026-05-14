import os
import secrets
import urllib.parse

from dotenv import load_dotenv

load_dotenv()


class Config:
    DB_HOST = os.getenv('DB_HOST')
    DB_USER = os.getenv('DB_USER')
    DB_PASSWORD = os.getenv('DB_PASSWORD', os.getenv('DB_PASS'))
    DB_NAME = os.getenv('DB_NAME')
    BASE_UPLOAD_URL = os.getenv('BASE_UPLOAD_URL', 'https://codestechnology.com/upload')

    if not all([DB_HOST, DB_USER, DB_PASSWORD, DB_NAME]):
        raise RuntimeError("Missing database credentials in .env file.")

    if 'site4now.net' not in DB_HOST:
        raise RuntimeError("DB_HOST must point to the remote site4now.net database host.")

    safe_password = urllib.parse.quote_plus(DB_PASSWORD)
    SQLALCHEMY_DATABASE_URI = (
    f"mysql+pymysql://{DB_USER}:{safe_password}@{DB_HOST}/{DB_NAME}"
    "?connect_timeout=10&read_timeout=30&write_timeout=30"
    )

    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ENGINE_OPTIONS = {
        "pool_pre_ping": True,
        "pool_recycle": 180,      # slightly lower = safer
        "pool_size": 3,           # reduce connections
        "max_overflow": 2,        # allow small burst
        "pool_timeout": 30,       # wait before failing
    }
    SQLALCHEMY_ECHO = False

    SECRET_KEY = os.getenv('SECRET_KEY', 'dev-secret-key-2024')
    JWT_SECRET_KEY = os.getenv('JWT_SECRET_KEY', 'dev-jwt-key-2024')
    JWT_TOKEN_LOCATION = ['cookies']
    JWT_COOKIE_SECURE = False
    JWT_COOKIE_SAMESITE = os.getenv('JWT_COOKIE_SAMESITE', 'Lax')
    JWT_ACCESS_TOKEN_EXPIRES = 86400

    APP_INSTANCE_ID = os.getenv('APP_INSTANCE_ID', secrets.token_urlsafe(32))

    ADOBE_CLIENT_ID = os.getenv('ADOBE_CLIENT_ID')
    ADOBE_CLIENT_SECRET = os.getenv('ADOBE_CLIENT_SECRET')

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    TEMPLATE_DOCX = os.path.join(BASE_DIR, 'Letter_pad.docx')
    STORAGE_DIR = os.path.join(BASE_DIR, 'storage')
    UPLOAD_FOLDER = os.path.join(BASE_DIR, 'file_storage')
    ALLOWED_EXTENSIONS = {'pdf', 'txt', 'docx'}
    MAX_FILE_SIZE = 16 * 1024 * 1024


DB_HOST = Config.DB_HOST
DB_USER = Config.DB_USER
DB_PASSWORD = Config.DB_PASSWORD
DB_NAME = Config.DB_NAME
BASE_UPLOAD_URL = Config.BASE_UPLOAD_URL
SQLALCHEMY_DATABASE_URI = Config.SQLALCHEMY_DATABASE_URI
SQLALCHEMY_TRACK_MODIFICATIONS = Config.SQLALCHEMY_TRACK_MODIFICATIONS
SQLALCHEMY_ENGINE_OPTIONS = Config.SQLALCHEMY_ENGINE_OPTIONS
SQLALCHEMY_ECHO = Config.SQLALCHEMY_ECHO
SECRET_KEY = Config.SECRET_KEY
JWT_SECRET_KEY = Config.JWT_SECRET_KEY
JWT_TOKEN_LOCATION = Config.JWT_TOKEN_LOCATION
JWT_COOKIE_SECURE = Config.JWT_COOKIE_SECURE
JWT_COOKIE_SAMESITE = Config.JWT_COOKIE_SAMESITE
JWT_ACCESS_TOKEN_EXPIRES = Config.JWT_ACCESS_TOKEN_EXPIRES
APP_INSTANCE_ID = Config.APP_INSTANCE_ID
ADOBE_CLIENT_ID = Config.ADOBE_CLIENT_ID
ADOBE_CLIENT_SECRET = Config.ADOBE_CLIENT_SECRET
BASE_DIR = Config.BASE_DIR
TEMPLATE_DOCX = Config.TEMPLATE_DOCX
STORAGE_DIR = Config.STORAGE_DIR
UPLOAD_FOLDER = Config.UPLOAD_FOLDER
ALLOWED_EXTENSIONS = Config.ALLOWED_EXTENSIONS
MAX_FILE_SIZE = Config.MAX_FILE_SIZE
