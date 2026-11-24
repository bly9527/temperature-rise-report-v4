"""Owner tool. The signing key is protected by the owner's Windows account."""
import base64
import os
import re
import sys
from datetime import date, timedelta
from pathlib import Path

import pywintypes
import win32crypt
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
from PyQt5.QtWidgets import QApplication, QDialog, QLabel, QLineEdit, QPushButton, QVBoxLayout

from thermal_gui import configure_qt_plugins
from thermal_license import PREFIX, PUBLIC_KEY, expiry_date


def issue_card(machine, issued):
    machine = machine.strip().upper()
    if not re.fullmatch(r'[A-F0-9]{24}', machine):
        raise ValueError('请粘贴软件显示的 24 位机器码。')
    try:
        path = Path(os.environ['LOCALAPPDATA']) / 'ThermalReportIssuer' / 'issuer.dpapi'
        protected_key = path.read_bytes()
        raw_key = win32crypt.CryptUnprotectData(protected_key, None, None, None, 1)[1]
        key = Ed25519PrivateKey.from_private_bytes(raw_key)
        if key.public_key().public_bytes_raw() != PUBLIC_KEY:
            raise ValueError
    except (OSError, ValueError, KeyError, pywintypes.error) as exc:
        raise ValueError('无法使用签发密钥，请在原签发电脑的 Windows 账户中运行。') from exc
    payload = f'{PREFIX}.{machine}.{issued.isoformat()}'
    signature = base64.urlsafe_b64encode(key.sign(payload.encode('ascii'))).decode('ascii')
    return f"{payload}.{signature.rstrip('=')}"


def main():
    configure_qt_plugins()
    app = QApplication(sys.argv)
    dialog = QDialog()
    dialog.setWindowTitle('温升报告生成器 · 签发卡密')
    dialog.setMinimumWidth(430)
    layout = QVBoxLayout(dialog)
    layout.addWidget(QLabel('目标电脑的机器码'))
    machine = QLineEdit()
    layout.addWidget(machine)
    issue = QPushButton('签发两个月卡密')
    layout.addWidget(issue)
    status = QLabel()
    layout.addWidget(status)
    result = QLineEdit()
    result.setReadOnly(True)
    layout.addWidget(result)
    copy = QPushButton('复制卡密')
    copy.setEnabled(False)
    copy.clicked.connect(lambda: app.clipboard().setText(result.text()))
    layout.addWidget(copy)

    def clear_result():
        result.clear()
        status.clear()
        copy.setEnabled(False)

    def issue_clicked():
        clear_result()
        today = date.today()
        try:
            card = issue_card(machine.text(), today)
        except ValueError as error:
            status.setText(str(error))
            return
        result.setText(card)
        status.setText(f'可用日期：{today} 至 {expiry_date(today) - timedelta(days=1)}（含）')
        copy.setEnabled(True)

    machine.textChanged.connect(clear_result)
    issue.clicked.connect(issue_clicked)
    dialog.exec_()


if __name__ == '__main__':
    main()
