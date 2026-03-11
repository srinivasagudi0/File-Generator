"""Simple authentication and file storage helpers for the app."""

from __future__ import annotations

import hashlib
import hmac
import os
import re
import sqlite3
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

USERNAME_PATTERN = re.compile(r'^[a-z0-9_.-]{3,32}$')
MIN_PASSWORD_LENGTH = 8
PBKDF2_ITERATIONS = 240_000


def data_root(data_dir: str | Path | None = None) -> Path:
    """Return the directory where the app keeps its data."""
    configured = ''
    if data_dir is None:
        configured = (os.getenv('FILEGEN_DATA_DIR') or '').strip()
    root = Path(data_dir or configured or (Path.cwd() / '.filegen_data')).resolve()
    root.mkdir(parents=True, exist_ok=True)
    return root


def database_path(data_dir: str | Path | None = None) -> Path:
    """Return the path to the SQLite database file."""
    return data_root(data_dir) / 'app.db'


def sanitize_file_name(name: str, default_name: str = 'output.txt') -> str:
    """Make a filename safe by keeping only supported characters."""
    raw = Path(str(name or '').strip()).name
    if not raw:
        raw = default_name

    suffix = ''.join(Path(raw).suffixes)
    stem = raw[:-len(suffix)] if suffix else raw
    safe_stem = re.sub(r'[^A-Za-z0-9._-]+', '_', stem).strip('._-')
    safe_suffix = re.sub(r'[^A-Za-z0-9.]+', '', suffix)

    if not safe_stem:
        fallback = Path(default_name)
        safe_stem = fallback.stem or 'output'
        if not safe_suffix:
            safe_suffix = fallback.suffix

    return f'{safe_stem}{safe_suffix}'


def normalize_username(username: str) -> str:
    """Trim whitespace and force lowercase before validations."""
    return str(username or '').strip().lower()


def validate_username(username: str) -> str:
    """Return an error message if the username is invalid."""
    normalized = normalize_username(username)
    if not normalized:
        return 'Username is required.'
    if not USERNAME_PATTERN.fullmatch(normalized):
        return 'Use 3-32 characters: lowercase letters, numbers, ".", "_" or "-".'
    return ''


def validate_password(password: str) -> str:
    """Ensure the password meets the minimum length requirement."""
    value = str(password or '')
    if len(value) < MIN_PASSWORD_LENGTH:
        return f'Password must be at least {MIN_PASSWORD_LENGTH} characters.'
    return ''


def _utcnow() -> str:
    """Return the current UTC time as an ISO-formatted string."""
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _connect(data_dir: str | Path | None = None) -> sqlite3.Connection:
    """Connect to the database and prepare the schema."""
    conn = sqlite3.connect(database_path(data_dir))
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys = ON')
    _ensure_schema(conn)
    return conn


def _ensure_schema(conn: sqlite3.Connection) -> None:
    """Create the tables used for users and files."""
    conn.executescript(
        '''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_salt TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL,
            display_name TEXT NOT NULL,
            file_type TEXT NOT NULL,
            storage_relpath TEXT NOT NULL UNIQUE,
            status TEXT NOT NULL DEFAULT 'active',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            deleted_at TEXT,
            FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS idx_files_user_updated
            ON files(user_id, status, updated_at DESC);
        '''
    )
    conn.commit()


def _hash_password(password: str, salt: bytes) -> str:
    """Derive a consistent hash from a password and salt."""
    digest = hashlib.pbkdf2_hmac(
        'sha256',
        str(password or '').encode('utf-8'),
        salt,
        PBKDF2_ITERATIONS,
    )
    return digest.hex()


def _user_row_to_dict(row: sqlite3.Row | None) -> dict[str, object] | None:
    """Convert a raw user row into a plain dictionary."""
    if row is None:
        return None
    return {
        'id': int(row['id']),
        'username': str(row['username']),
        'created_at': str(row['created_at']),
    }


def _file_row_to_dict(
    row: sqlite3.Row | None,
    data_dir: str | Path | None = None,
) -> dict[str, object] | None:
    """Turn a file row into a dictionary with resolved paths."""
    if row is None:
        return None
    root = data_root(data_dir)
    storage_relpath = str(row['storage_relpath'])
    storage_path = (root / storage_relpath).resolve()
    return {
        'id': int(row['id']),
        'user_id': int(row['user_id']),
        'display_name': str(row['display_name']),
        'file_type': str(row['file_type']),
        'storage_relpath': storage_relpath,
        'storage_path': str(storage_path),
        'status': str(row['status']),
        'created_at': str(row['created_at']),
        'updated_at': str(row['updated_at']),
        'deleted_at': None if row['deleted_at'] is None else str(row['deleted_at']),
    }


def create_user(
    username: str,
    password: str,
    data_dir: str | Path | None = None,
) -> tuple[dict[str, object] | None, str]:
    """Validate inputs and add a new user if there are no conflicts."""
    username_error = validate_username(username)
    if username_error:
        return None, username_error

    password_error = validate_password(password)
    if password_error:
        return None, password_error

    normalized = normalize_username(username)
    salt = os.urandom(16)
    created_at = _utcnow()

    with _connect(data_dir) as conn:
        try:
            conn.execute(
                '''
                INSERT INTO users (username, password_salt, password_hash, created_at)
                VALUES (?, ?, ?, ?)
                ''',
                (normalized, salt.hex(), _hash_password(password, salt), created_at),
            )
            conn.commit()
        except sqlite3.IntegrityError:
            return None, 'That username already exists.'

        row = conn.execute(
            'SELECT id, username, created_at FROM users WHERE username = ?',
            (normalized,),
        ).fetchone()
    return _user_row_to_dict(row), ''


def authenticate_user(
    username: str,
    password: str,
    data_dir: str | Path | None = None,
) -> tuple[dict[str, object] | None, str]:
    """Check the given credentials and return the matching user when valid."""
    normalized = normalize_username(username)
    if not normalized or not password:
        return None, 'Enter both username and password.'

    with _connect(data_dir) as conn:
        row = conn.execute('SELECT * FROM users WHERE username = ?', (normalized,)).fetchone()
        if row is None:
            return None, 'Invalid username or password.'

        salt_hex = str(row['password_salt'])
        salt = bytes.fromhex(salt_hex)
        expected_hash = str(row['password_hash'])
        actual_hash = _hash_password(password, salt)
        if not hmac.compare_digest(actual_hash, expected_hash):
            return None, 'Invalid username or password.'

    return _user_row_to_dict(row), ''


def user_files_dir(user_id: int, data_dir: str | Path | None = None) -> Path:
    """Build and ensure the directory where a user keeps their files."""
    target = data_root(data_dir) / 'users' / str(int(user_id)) / 'files'
    target.mkdir(parents=True, exist_ok=True)
    return target


def is_user_storage_path(
    user_id: int,
    path: str | Path,
    data_dir: str | Path | None = None,
) -> bool:
    """Check whether a path lives inside the user's directory."""
    candidate = Path(path).resolve()
    root = user_files_dir(user_id, data_dir).resolve()
    try:
        candidate.relative_to(root)
    except ValueError:
        return False
    return True


def allocate_storage_path(
    user_id: int,
    display_name: str,
    data_dir: str | Path | None = None,
) -> Path:
    """Choose a unique storage path for a new file owned by user_id."""
    safe_name = sanitize_file_name(display_name)
    stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S')
    target = user_files_dir(user_id, data_dir) / f'{stamp}_{uuid4().hex[:10]}_{safe_name}'
    return target.resolve()


def create_file_record(
    user_id: int,
    display_name: str,
    file_type: str,
    storage_path: str | Path,
    data_dir: str | Path | None = None,
) -> dict[str, object]:
    """Store a new file entry for the user and return its metadata."""
    safe_name = sanitize_file_name(display_name)
    target = Path(storage_path).resolve()
    if not is_user_storage_path(user_id, target, data_dir):
        raise ValueError('Storage path must stay inside the user workspace.')

    root = data_root(data_dir).resolve()
    storage_relpath = str(target.relative_to(root))
    now = _utcnow()

    with _connect(data_dir) as conn:
        cursor = conn.execute(
            '''
            INSERT INTO files (
                user_id, display_name, file_type, storage_relpath, status, created_at, updated_at, deleted_at
            )
            VALUES (?, ?, ?, ?, 'active', ?, ?, NULL)
            ''',
            (int(user_id), safe_name, str(file_type), storage_relpath, now, now),
        )
        conn.commit()
        row = conn.execute('SELECT * FROM files WHERE id = ?', (cursor.lastrowid,)).fetchone()
    record = _file_row_to_dict(row, data_dir)
    if record is None:
        raise RuntimeError('Failed to create file record.')
    return record


def update_file_record(
    user_id: int,
    file_id: int,
    data_dir: str | Path | None = None,
    *,
    display_name: str | None = None,
    file_type: str | None = None,
    storage_path: str | Path | None = None,
    status: str | None = None,
) -> dict[str, object] | None:
    """Update metadata for a stored file."""
    assignments: list[str] = ['updated_at = ?']
    values: list[object] = [_utcnow()]

    if display_name is not None:
        assignments.append('display_name = ?')
        values.append(sanitize_file_name(display_name))
    if file_type is not None:
        assignments.append('file_type = ?')
        values.append(str(file_type))
    if storage_path is not None:
        target = Path(storage_path).resolve()
        if not is_user_storage_path(user_id, target, data_dir):
            raise ValueError('Storage path must stay inside the user workspace.')
        root = data_root(data_dir).resolve()
        assignments.append('storage_relpath = ?')
        values.append(str(target.relative_to(root)))
    if status is not None:
        assignments.append('status = ?')
        values.append(str(status))
        if status == 'deleted':
            assignments.append('deleted_at = ?')
            values.append(_utcnow())
        else:
            assignments.append('deleted_at = NULL')

    values.extend([int(file_id), int(user_id)])

    with _connect(data_dir) as conn:
        conn.execute(
            f'UPDATE files SET {", ".join(assignments)} WHERE id = ? AND user_id = ?',
            tuple(values),
        )
        conn.commit()
        row = conn.execute(
            'SELECT * FROM files WHERE id = ? AND user_id = ?',
            (int(file_id), int(user_id)),
        ).fetchone()
    return _file_row_to_dict(row, data_dir)


def mark_file_deleted(
    user_id: int,
    file_id: int,
    data_dir: str | Path | None = None,
) -> dict[str, object] | None:
    """Soft-delete a file by marking it deleted."""
    return update_file_record(user_id, file_id, data_dir, status='deleted')


def get_file_record(
    user_id: int,
    file_id: int,
    data_dir: str | Path | None = None,
    *,
    include_deleted: bool = False,
) -> dict[str, object] | None:
    """Look up a single file record, optionally including deleted files."""
    query = 'SELECT * FROM files WHERE user_id = ? AND id = ?'
    params: list[object] = [int(user_id), int(file_id)]
    if not include_deleted:
        query += ' AND status = ?'
        params.append('active')

    with _connect(data_dir) as conn:
        row = conn.execute(query, tuple(params)).fetchone()
    return _file_row_to_dict(row, data_dir)


def list_file_records(
    user_id: int,
    data_dir: str | Path | None = None,
    *,
    include_deleted: bool = True,
) -> list[dict[str, object]]:
    """Return all file records for a user, sorted with active files first."""
    query = 'SELECT * FROM files WHERE user_id = ?'
    params: list[object] = [int(user_id)]
    if not include_deleted:
        query += ' AND status = ?'
        params.append('active')
    query += " ORDER BY CASE status WHEN 'active' THEN 0 ELSE 1 END, updated_at DESC, id DESC"

    records: list[dict[str, object]] = []
    with _connect(data_dir) as conn:
        rows = conn.execute(query, tuple(params)).fetchall()
        for row in rows:
            record = _file_row_to_dict(row, data_dir)
            if record is not None:
                records.append(record)
    return records
