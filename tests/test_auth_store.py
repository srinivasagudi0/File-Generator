from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from auth_store import (
    allocate_storage_path,
    authenticate_user,
    create_file_record,
    create_user,
    get_file_record,
    is_user_storage_path,
    list_file_records,
    mark_file_deleted,
    update_file_record,
)


class AuthStoreTests(unittest.TestCase):
    def test_create_and_authenticate_user(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            user, error = create_user('alice_1', 'supersecret', tmp_dir)
            self.assertEqual(error, '')
            self.assertIsNotNone(user)
            self.assertEqual(user['username'], 'alice_1')

            duplicate, duplicate_error = create_user('Alice_1', 'supersecret', tmp_dir)
            self.assertIsNone(duplicate)
            self.assertIn('already exists', duplicate_error)

            authenticated, auth_error = authenticate_user('ALICE_1', 'supersecret', tmp_dir)
            self.assertEqual(auth_error, '')
            self.assertIsNotNone(authenticated)
            self.assertEqual(authenticated['id'], user['id'])

            invalid_user, invalid_error = authenticate_user('alice_1', 'wrong-pass', tmp_dir)
            self.assertIsNone(invalid_user)
            self.assertIn('Invalid username or password.', invalid_error)

    def test_file_records_stay_in_user_storage(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            user, error = create_user('bob', 'supersecret', tmp_dir)
            self.assertEqual(error, '')
            self.assertIsNotNone(user)

            path = allocate_storage_path(int(user['id']), '../../quarterly/report.txt', tmp_dir)
            self.assertTrue(is_user_storage_path(int(user['id']), path, tmp_dir))

            target = Path(path)
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_text('hello', encoding='utf-8')

            record = create_file_record(int(user['id']), '../../quarterly/report.txt', 'txt', target, tmp_dir)
            self.assertEqual(record['display_name'], 'report.txt')
            self.assertEqual(record['file_type'], 'txt')
            self.assertEqual(len(list_file_records(int(user['id']), tmp_dir)), 1)

            updated = update_file_record(int(user['id']), int(record['id']), tmp_dir, display_name='renamed.txt')
            self.assertIsNotNone(updated)
            self.assertEqual(updated['display_name'], 'renamed.txt')

    def test_deleted_files_remain_in_history(self) -> None:
        with TemporaryDirectory() as tmp_dir:
            user, error = create_user('carol', 'supersecret', tmp_dir)
            self.assertEqual(error, '')
            self.assertIsNotNone(user)

            path = allocate_storage_path(int(user['id']), 'notes.txt', tmp_dir)
            Path(path).write_text('owned content', encoding='utf-8')
            record = create_file_record(int(user['id']), 'notes.txt', 'txt', path, tmp_dir)

            deleted = mark_file_deleted(int(user['id']), int(record['id']), tmp_dir)
            self.assertIsNotNone(deleted)
            self.assertEqual(deleted['status'], 'deleted')

            active = get_file_record(int(user['id']), int(record['id']), tmp_dir)
            self.assertIsNone(active)

            history = list_file_records(int(user['id']), tmp_dir, include_deleted=True)
            self.assertEqual(len(history), 1)
            self.assertEqual(history[0]['status'], 'deleted')


if __name__ == '__main__':
    unittest.main()
