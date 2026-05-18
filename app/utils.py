from __future__ import annotations
from datetime import date
from functools import wraps
from flask import request
from . import db
from .models import User, AuditLog

DEFAULT_REMINDER_DAYS = 30

def days_until(d):
    if not d:
        return None
    return (d - date.today()).days

def is_due_soon(d, within_days=DEFAULT_REMINDER_DAYS):
    x = days_until(d)
    return x is not None and x <= within_days

def is_overdue(d):
    x = days_until(d)
    return x is not None and x < 0

def current_user() -> User | None:
    return None

def role_required(role: str):
    def deco(fn):
        @wraps(fn)
        def wrap(*args, **kwargs):
            return fn(*args, **kwargs)
        return wrap
    return deco

def audit(action: str, entity_type=None, entity_id=None, vehicle_id=None, note=None):
    ip = request.headers.get("X-Forwarded-For", request.remote_addr)
    rec = AuditLog(
        actor=None,
        action=action,
        entity_type=entity_type,
        entity_id=entity_id,
        vehicle_id=vehicle_id,
        note=note,
        ip=ip,
    )
    db.session.add(rec)
    db.session.commit()