"""Хранение ключа DaData в системном хранилище учётных данных."""

import keyring


class TokenStore:
    service = 'reinziniring-pd DaData'
    account = 'default'

    def load(self):
        return keyring.get_password(self.service, self.account) or ''

    def save(self, token):
        keyring.set_password(self.service, self.account, token)

    def delete(self):
        try:
            keyring.delete_password(self.service, self.account)
        except keyring.errors.PasswordDeleteError:
            pass
