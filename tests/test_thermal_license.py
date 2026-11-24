import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
import win32crypt

import issue_thermal_card as issuer
import thermal_license as licensing


class LicenseTests(unittest.TestCase):
    machine = 'AB' * 12
    start = date(2026, 11, 23)

    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        root = Path(directory.name)
        key = Ed25519PrivateKey.generate()
        public = key.public_key().public_bytes_raw()
        for module in (issuer, licensing):
            self.enter_patch(patch.object(module, 'PUBLIC_KEY', public))
        self.enter_patch(patch.dict('os.environ', LOCALAPPDATA=str(root)))
        folder = root / 'ThermalReportIssuer'
        folder.mkdir()
        (folder / 'issuer.dpapi').write_bytes(win32crypt.CryptProtectData(
            key.private_bytes_raw(), '', None, None, None, 1))
        self.card = issuer.issue_card(self.machine, self.start)
        self.path = root / 'activation.dat'

    def enter_patch(self, context):
        context.start()
        self.addCleanup(context.stop)

    def test_dates_and_end_of_month(self):
        self.assertEqual(licensing.expiry_date(date(2023, 12, 31)), date(2024, 2, 29))
        self.assertEqual(licensing.expiry_date(date(2026, 12, 31)), date(2027, 2, 28))
        end = date(2027, 1, 23)
        for day in (self.start, date(2027, 1, 22)):
            self.assertEqual(licensing.verify_card(self.card, self.machine, day), end)
        for day in (date(2026, 11, 22), end, date(2027, 1, 24)):
            with self.assertRaises(ValueError):
                licensing.verify_card(self.card, self.machine, day)

    def test_invalid_tampered_foreign_machine(self):
        for card in ('', 'wrong', self.card.replace('2026-11-23', '2026-12-23'),
                     self.card.replace(self.machine, 'CD' * 12),
                     issuer.issue_card('CD' * 12, self.start)):
            with self.assertRaises(ValueError):
                licensing.verify_card(card, self.machine, self.start)

    def test_restart_expiry_and_renewal(self):
        first = licensing.Activation(self.machine, self.path)
        first.accept(self.card, self.start)
        self.assertNotIn(self.card.encode(), self.path.read_bytes())
        second = licensing.Activation(self.machine, self.path)
        second.load()
        second.accept(second.card, date(2027, 1, 22))
        end = date(2027, 1, 23)
        with self.assertRaisesRegex(ValueError, '到期'):
            second.accept(second.card, end)
        with self.assertRaisesRegex(ValueError, '校准'):
            second.accept(second.card, self.start)
        renewal = issuer.issue_card(self.machine, end)
        self.assertEqual(second.accept(renewal, end), date(2027, 3, 23))

    def test_damaged_cache_and_write_failure(self):
        activation = licensing.Activation(self.machine, self.path)
        self.path.write_bytes(b'bad-cache')
        with self.assertRaisesRegex(ValueError, '重新输入'):
            activation.load()
        activation.accept(self.card, self.start)
        broken = licensing.Activation(self.machine, self.path / 'child')
        with self.assertRaises(OSError):
            broken.accept(self.card, self.start)
        self.assertEqual(broken.card, '')
