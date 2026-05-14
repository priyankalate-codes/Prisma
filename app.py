"""
Main application file with JWT authentication and Session Reset
"""
from datetime import timedelta
import os
import platform
from flask import Flask, jsonify, render_template, redirect, request, url_for, make_response
from flask_jwt_extended import JWTManager, verify_jwt_in_request, get_jwt_identity, get_jwt, unset_jwt_cookies
from flask_cors import CORS
from config import (
    SQLALCHEMY_DATABASE_URI,
    SQLALCHEMY_TRACK_MODIFICATIONS,
    APP_INSTANCE_ID,
    SECRET_KEY,
    JWT_COOKIE_SECURE,
    JWT_COOKIE_SAMESITE,
    JWT_ACCESS_TOKEN_EXPIRES,
)
from utils.db_manager import init_db, db
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


   
app = Flask(__name__)

CORS(app, supports_credentials=True)

app.config['SQLALCHEMY_DATABASE_URI'] = SQLALCHEMY_DATABASE_URI
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = SQLALCHEMY_TRACK_MODIFICATIONS

# Every restart = new key = old cookies invalid = login required
app.config['SECRET_KEY'] = SECRET_KEY
app.config['JWT_SECRET_KEY'] = SECRET_KEY

app.config['JWT_TOKEN_LOCATION'] = ['cookies']
app.config['JWT_COOKIE_HTTPONLY'] = True
app.config['JWT_COOKIE_SECURE'] = JWT_COOKIE_SECURE
app.config['JWT_COOKIE_SAMESITE'] = JWT_COOKIE_SAMESITE
app.config['JWT_COOKIE_CSRF_PROTECT'] = os.getenv('JWT_COOKIE_CSRF_PROTECT', 'false').lower() == 'true'
app.config['JWT_ACCESS_TOKEN_EXPIRES'] = timedelta(seconds=JWT_ACCESS_TOKEN_EXPIRES)  # keep JWT config aligned with config.py

jwt = JWTManager(app)

init_db(app)

from route.auth import auth_bp
from route.process import process_bp
app.register_blueprint(auth_bp, url_prefix='/api/auth')
app.register_blueprint(process_bp, url_prefix='/api/process')


# ── FRONTEND ROUTES ────────────────────────────────────────

@app.route('/login')
def login_page():
    return render_template('login.html')

@app.route('/')
def index():
    try:
        verify_jwt_in_request(optional=True)
        user_id = get_jwt_identity()

        if not user_id:
            return redirect(url_for('login_page'))

        claims = get_jwt()
        token_instance = claims.get('instance_id')
        
        if token_instance != APP_INSTANCE_ID:
            logger.warning("Instance mismatch — forcing re-login")
            response = make_response(redirect(url_for('login_page')))
            unset_jwt_cookies(response)
            return response

        response = make_response(render_template('upload.html'))
        response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'
        return response

    except Exception as e:
        logger.error(f"Authentication error: {e}")
        return redirect(url_for('login_page'))


# ── API ROUTES ─────────────────────────────────────────────

@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({'success': True, 'message': 'Prisma API is running', 'version': '1.0.0'}), 200


# ── JWT ERROR HANDLERS ─────────────────────────────────────
# Bug 3 fix: handlers now redirect browser requests, return JSON for API calls

@jwt.expired_token_loader
def expired_token_callback(jwt_header, jwt_payload):
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'error': 'Token expired'}), 401
    response = make_response(redirect(url_for('login_page')))
    unset_jwt_cookies(response)
    return response

@jwt.invalid_token_loader
def invalid_token_callback(error):
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'error': 'Invalid token'}), 401
    response = make_response(redirect(url_for('login_page')))
    unset_jwt_cookies(response)
    return response

@jwt.unauthorized_loader
def missing_token_callback(error):
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'error': 'No token provided'}), 401
    return redirect(url_for('login_page'))


@app.teardown_appcontext
def shutdown_session(exception=None):
    db.session.remove()


if __name__ == '__main__':
    app.run(
        debug=os.getenv('FLASK_DEBUG', 'false').lower() == 'true',
        host='0.0.0.0',
        port=int(os.getenv('PORT', '5000'))
    )
