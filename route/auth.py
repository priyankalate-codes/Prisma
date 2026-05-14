"""
Authentication routes for user login, logout, and profile retrieval.
"""

import logging

from utils.decorators import require_auth
from flask import Blueprint, request, jsonify, make_response
from flask_jwt_extended import create_access_token, set_access_cookies, unset_jwt_cookies
from utils.db_manager import db, User
from config import APP_INSTANCE_ID
from datetime import timedelta

auth_bp = Blueprint('auth', __name__)
logger = logging.getLogger(__name__)

@auth_bp.route('/login', methods=['POST'])
def login():
    try:
        data = request.get_json()
        if not data or not all(k in data for k in ['email', 'password']):
            return jsonify({'success': False, 'error': 'Missing email or password'}), 400
        
        email = data['email'].lower().strip()
        password = data['password']
        user = User.query.filter_by(Email=email).first()
        
        if not user or not user.check_password(password):
            return jsonify({'success': False, 'error': 'Wrong password'}), 401
        
        if not user.IsActive:
            return jsonify({'success': False, 'error': 'Account disabled'}), 403
        
        # Claims link the token to a specific 'version' in the DB
        access_token = create_access_token(
            identity=str(user.Id),
            additional_claims={
                'token_version': user.TokenVersion,
                'instance_id': APP_INSTANCE_ID   # ← ADD THIS
            },
            expires_delta=timedelta(hours=24)
        )
                
        response = make_response(jsonify({
            'success': True,
            'user': user.to_dict()
        }))
        
        set_access_cookies(response, access_token)
        return response, 200
        
    except Exception as e:
        logger.error("Login failed: %s", e, exc_info=True)
        return jsonify({'success': False, 'error': 'Login failed. Please try again.'}), 500

@auth_bp.route('/logout', methods=['POST'])
@require_auth
def logout(current_user):
    try:      
        # Security: Invalidate tokens on the server side
        current_user.invalidate_tokens()
        db.session.commit() # Don't forget to commit the version change!
        
        response = make_response(jsonify({'success': True}))
        unset_jwt_cookies(response)
        return response, 200
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@auth_bp.route('/me', methods=['GET'])
@require_auth
def get_current_user(current_user):
    """Combines 'validate' and 'me' into one call."""
    return jsonify({
        'success': True,
        'user': current_user.to_dict()
    }), 200
