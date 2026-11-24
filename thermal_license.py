"""Offline, machine-specific two-calendar-month activation for Windows."""
import base64
import calendar
import hashlib
import os
import re
import uuid
import winreg
from datetime import date
from pathlib import Path

import pywintypes
import win32crypt
from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
from PyQt5.QtCore import QIODevice, QSaveFile


PUBLIC_KEY = bytes.fromhex('dcb7e018d60809d1b5c5c7590a01f71a1f8d7806a674533045409d3ac7383419')
PREFIX = 'WS1'


def machine_code():
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Microsoft\Cryptography',
                            access=winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
            guid = uuid.UUID(winreg.QueryValueEx(key, 'MachineGuid')[0])
    except (OSError, ValueError, AttributeError) as exc:
        raise ValueError('无法读取本机机器码，请检查 Windows 设置。') from exc
    return hashlib.sha256(b'thermal-report-v4:' + guid.bytes).hexdigest()[:24].upper()


def expiry_date(issued):
    year, month_index = divmod(issued.year * 12 + issued.month - 1 + 2, 12)
    month = month_index + 1
    return date(year, month, min(issued.day, calendar.monthrange(year, month)[1]))


def verify_card(card, machine, today):
    try:
        prefix, target, issued_text, signature = card.strip().split('.')
        if prefix != PREFIX or not re.fullmatch(r'[A-F0-9]{24}', target):
            raise ValueError
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', issued_text, re.ASCII):
            raise ValueError
        issued = date.fromisoformat(issued_text)
        if not re.fullmatch(r'[A-Za-z0-9_-]{86}', signature):
            raise ValueError
        raw = base64.urlsafe_b64decode(signature + '==')
        payload = f'{prefix}.{target}.{issued_text}'.encode('ascii')
        Ed25519PublicKey.from_public_bytes(PUBLIC_KEY).verify(raw, payload)
        expiry = expiry_date(issued)
    except (InvalidSignature, ValueError, OverflowError) as exc:
        raise ValueError('卡密无效，请完整复制后重试。') from exc
    if target != machine:
        raise ValueError('这张卡密属于其他电脑，请使用本机机器码重新签发。')
    if today < issued:
        raise ValueError('尚未到签发日期，请检查电脑日期。')
    if today >= expiry:
        raise ValueError(f'卡密已于 {expiry} 到期，请输入新卡密。')
    return expiry


class Activation:
    def __init__(self, machine, path):
        self.machine = machine
        self.path = path
        self.card = ''
        self.last_seen = None

    @classmethod
    def current_user(cls):
        return cls(machine_code(), Path(os.environ['LOCALAPPDATA']) /
                   'ThermalReportV4' / 'activation.dat')

    def load(self):
        try:
            blob = self.path.read_bytes()
        except FileNotFoundError:
            return
        try:
            decoded = win32crypt.CryptUnprotectData(blob, None, None, None, 1)[1]
            seen, card = decoded.decode('ascii').split('\n', 1)
            self.last_seen = date.fromisoformat(seen)
            self.card = card
        except (ValueError, pywintypes.error) as exc:
            raise ValueError('本机授权记录无法读取，请重新输入卡密。') from exc

    def _remember(self, card, today):
        data = f'{today}\n{card}'.encode('ascii')
        try:
            blob = win32crypt.CryptProtectData(data, 'Thermal report license',
                                               None, None, None, 1)
        except pywintypes.error as exc:
            raise OSError('无法保护本机授权记录。') from exc
        self.path.parent.mkdir(parents=True, exist_ok=True)
        output = QSaveFile(str(self.path))
        if (not output.open(QIODevice.WriteOnly) or output.write(blob) != len(blob)
                or not output.commit()):
            raise OSError('无法写入本机授权记录。')
        self.card, self.last_seen = card, today

    def accept(self, card, today):
        if self.last_seen and today < self.last_seen:
            raise ValueError('电脑日期早于上次使用日期，请校准后重试。')
        card = card.strip()
        if self.card and card == self.card and self.last_seen != today:
            self._remember(card, today)
        expiry = verify_card(card, self.machine, today)
        if card != self.card or self.last_seen != today:
            self._remember(card, today)
        return expiry
