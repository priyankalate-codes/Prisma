"""
Custom decorators for route protection
"""
from functools import wraps
import time
from flask import jsonify
from flask_jwt_extended import jwt_required, get_jwt_identity, get_jwt
from sqlalchemy.exc import OperationalError, SQLAlchemyError
from utils.db_manager import User, db
import logging

logger = logging.getLogger(__name__)


def _auth_error(message, status_code=401):
    return jsonify({
        'success': False,
        'error': message
    }), status_code

def require_auth(fn):
    @wraps(fn)
    @jwt_required()
    def wrapper(*args, **kwargs):
        try:
            user_id = get_jwt_identity()
            claims = get_jwt()

            if user_id is None:
                return _auth_error('Session expired. Please login again.')

            try:
                user_pk = int(user_id)
            except (TypeError, ValueError):
                return _auth_error('Invalid session. Please login again.')

            user = None
            for attempt in range(3):
                try:
                    user = User.query.get(user_pk)
                    break
                except OperationalError as e:
                    if "2013" in str(e) and attempt < 2:
                        db.session.remove()
                        time.sleep(0.5)
                        continue
                    raise

            if not user:
                return _auth_error('User not found. Please login again.')
            
            if not user.IsActive:
                return _auth_error('Account disabled', 403)
            
            # IMPORTANT: Check token instance (app restart forced invalidation)
            from config import APP_INSTANCE_ID
            if claims.get('instance_id') != APP_INSTANCE_ID:
                return _auth_error('Session expired due to server restart. Please login again.')
            
            # IMPORTANT: Check token version (user-initiated logout security)
            if claims.get('token_version') != user.TokenVersion:
                return _auth_error('Session expired. Please login again.')
            
            return fn(current_user = user, *args, **kwargs)

        except SQLAlchemyError as e:
            logger.error("Authentication database error: %s", e, exc_info=True)
            return _auth_error('Authentication database unavailable. Please try again.', 503)
        except Exception as e:
            logger.error("Authentication failed: %s", e, exc_info=True)
            return _auth_error(f"Authentication failed: {str(e)}")
        
    return wrapper
