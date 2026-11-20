import unittest

from src.auth.gsheet import SheetUnavailableError, normalize_private_key
from src.auth.models import ROLE_ADMIN, ROLE_USER, UserAccount, username_problem
from src.auth.passwords import hash_password, is_password_hash, password_problem, verify_password
from src.auth.remember_token import create_token, read_username, verify_token
from src.auth.user_store import SecretsUserBackend, UserDirectory


class PasswordTest(unittest.TestCase):
    def test_hash_round_trip_and_unique_salt(self):
        first = hash_password("Mật-khẩu-1", iterations=1_000)
        second = hash_password("Mật-khẩu-1", iterations=1_000)

        self.assertTrue(is_password_hash(first))
        self.assertNotEqual(first, second)
        self.assertTrue(verify_password("Mật-khẩu-1", first))
        self.assertFalse(verify_password("mật-khẩu-1", first))

    def test_plain_text_value_from_sheet_is_still_accepted(self):
        self.assertTrue(verify_password("abc123", "abc123"))
        self.assertFalse(verify_password("abc124", "abc123"))
        self.assertFalse(verify_password("", ""))

    def test_malformed_hash_is_rejected(self):
        self.assertFalse(verify_password("x", "pbkdf2_sha256$abc$$"))

    def test_password_rules(self):
        self.assertIsNotNone(password_problem("12345"))
        self.assertIsNotNone(password_problem(" 123456"))
        self.assertIsNone(password_problem("123456"))


class UsernameTest(unittest.TestCase):
    def test_office_codes_are_valid(self):
        for name in ("VPSG", "VPBD", "user"):
            self.assertIsNone(username_problem(name))

    def test_spaces_and_vietnamese_marks_are_rejected(self):
        self.assertIsNotNone(username_problem("VP SG"))
        self.assertIsNotNone(username_problem("Phương"))
        self.assertIsNotNone(username_problem("a"))


class RememberTokenTest(unittest.TestCase):
    SECRET = "s" * 32

    def _token(self, **overrides):
        options = dict(
            secret=self.SECRET, password_hash="hash", token_version=0, lifetime_seconds=60, now=1_000
        )
        options.update(overrides)
        return create_token("VPSG", **options)

    def test_valid_token(self):
        token = self._token()
        self.assertEqual(read_username(token), "VPSG")
        self.assertTrue(
            verify_token(token, secret=self.SECRET, password_hash="hash", token_version=0, now=1_030)
        )

    def test_token_expires(self):
        self.assertFalse(
            verify_token(self._token(), secret=self.SECRET, password_hash="hash", token_version=0, now=1_061)
        )

    def test_password_change_or_revocation_invalidates_token(self):
        token = self._token()
        self.assertFalse(verify_token(token, secret=self.SECRET, password_hash="new", token_version=0, now=1_000))
        self.assertFalse(verify_token(token, secret=self.SECRET, password_hash="hash", token_version=1, now=1_000))
        self.assertFalse(verify_token(token, secret="x" * 32, password_hash="hash", token_version=0, now=1_000))

    def test_tampered_or_garbage_token_is_rejected(self):
        token = self._token()
        encoded, signature = token.split(".")
        forged = create_token(
            "VPBD", secret="other" * 8, password_hash="hash", token_version=0, lifetime_seconds=60, now=1_000
        ).split(".")[0]
        self.assertFalse(
            verify_token(f"{forged}.{signature}", secret=self.SECRET, password_hash="hash", token_version=0, now=1_000)
        )
        for garbage in (None, "", "abc", "a.b.c", "!!!.sig"):
            self.assertIsNone(read_username(garbage))
            self.assertFalse(
                verify_token(garbage, secret=self.SECRET, password_hash="hash", token_version=0, now=1_000)
            )


class _MemoryBackend:
    writable = True

    def __init__(self, users):
        self.users = list(users)
        self.fail = False

    def read(self, *, use_cache=True):
        if self.fail:
            raise SheetUnavailableError("offline")
        return list(self.users)

    def write(self, users):
        if self.fail:
            raise SheetUnavailableError("offline")
        self.users = list(users)


def _directory():
    backend = _MemoryBackend(
        [
            UserAccount("VPSG", hash_password("admin-pass", iterations=1_000), ROLE_ADMIN),
            UserAccount("VPBD", hash_password("vpbd-pass", iterations=1_000), ROLE_USER),
            UserAccount("user", "plain-pass", ROLE_USER),
        ]
    )
    fallback = (UserAccount("VPSG", "fallback-pass", ROLE_ADMIN),)
    return UserDirectory(backend, fallback_users=fallback), backend


class UserDirectoryTest(unittest.TestCase):
    def test_authenticate(self):
        directory, _ = _directory()
        self.assertEqual(directory.authenticate("VPSG", "admin-pass").username, "VPSG")
        self.assertEqual(directory.authenticate("user", "plain-pass").username, "user")
        self.assertIsNone(directory.authenticate("VPSG", "wrong"))
        self.assertIsNone(directory.authenticate("vpsg", "admin-pass"))

    def test_add_user_validates_and_hashes(self):
        directory, backend = _directory()
        self.assertFalse(directory.add_user("vpsg", "123456", ROLE_USER).ok)
        self.assertFalse(directory.add_user("VPHN", "123", ROLE_USER).ok)
        result = directory.add_user("VPHN", "hanoi-pass", ROLE_USER)

        self.assertTrue(result.ok)
        stored = next(user for user in backend.users if user.username == "VPHN")
        self.assertTrue(is_password_hash(stored.password_hash))
        self.assertTrue(directory.authenticate("VPHN", "hanoi-pass"))

    def test_change_password_requires_current_and_bumps_token_version(self):
        directory, backend = _directory()
        self.assertEqual(directory.change_password("VPBD", "wrong", "new-pass").message, "Mật khẩu hiện tại không đúng.")
        self.assertFalse(directory.change_password("VPBD", "vpbd-pass", "vpbd-pass").ok)

        self.assertTrue(directory.change_password("VPBD", "vpbd-pass", "new-pass").ok)
        updated = directory.find("VPBD")
        self.assertEqual(updated.token_version, 1)
        self.assertTrue(directory.authenticate("VPBD", "new-pass"))
        self.assertIsNone(directory.authenticate("VPBD", "vpbd-pass"))

    def test_update_user_keeps_password_when_blank(self):
        directory, _ = _directory()
        before = directory.find("VPBD")
        self.assertTrue(directory.update_user("VPBD", new_username="VPBD", role=ROLE_ADMIN).ok)
        after = directory.find("VPBD")

        self.assertEqual(after.password_hash, before.password_hash)
        self.assertEqual(after.role, ROLE_ADMIN)
        self.assertEqual(after.token_version, before.token_version)

    def test_last_admin_is_protected(self):
        directory, _ = _directory()
        self.assertFalse(directory.update_user("VPSG", new_username="VPSG", role=ROLE_USER).ok)
        self.assertFalse(directory.delete_user("VPSG", acting_username="VPBD").ok)
        self.assertFalse(directory.delete_user("VPSG", acting_username="VPSG").ok)
        self.assertTrue(directory.delete_user("user", acting_username="VPSG").ok)
        self.assertIsNone(directory.find("user"))

    def test_revoke_remembered_bumps_version(self):
        directory, _ = _directory()
        self.assertTrue(directory.revoke_remembered("VPBD").ok)
        self.assertEqual(directory.find("VPBD").token_version, 1)

    def test_sheet_outage_falls_back_to_secrets_and_blocks_changes(self):
        directory, backend = _directory()
        backend.fail = True

        self.assertEqual([user.username for user in directory.users()], ["VPSG"])
        self.assertIsNotNone(directory.warning)
        self.assertTrue(directory.authenticate("VPSG", "fallback-pass"))
        result = directory.change_password("VPSG", "admin-pass", "another-pass")
        self.assertFalse(result.ok)
        self.assertIn("Google Sheet", result.message)

    def test_secrets_backend_is_read_only(self):
        directory = UserDirectory(SecretsUserBackend((UserAccount("VPSG", "p", ROLE_ADMIN),)))
        self.assertFalse(directory.writable)
        self.assertTrue(directory.authenticate("VPSG", "p"))
        self.assertFalse(directory.add_user("VPHN", "123456", ROLE_USER).ok)


class PrivateKeyTest(unittest.TestCase):
    def test_single_line_key_is_rebuilt(self):
        key = "-----BEGIN PRIVATE KEY-----\\nABCD EFGH\\n-----END PRIVATE KEY-----\\n"
        self.assertEqual(
            normalize_private_key(key),
            "-----BEGIN PRIVATE KEY-----\nABCDEFGH\n-----END PRIVATE KEY-----\n",
        )


if __name__ == "__main__":
    unittest.main()
