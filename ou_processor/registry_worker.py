"""Фоновая проверка резервного реестра для локального интерфейса."""

from http.client import IncompleteRead, RemoteDisconnected
import socket
import ssl
import threading
import time
from urllib.error import HTTPError, URLError

from .registry import update


def retryable(error):
    if isinstance(error, HTTPError):
        return error.code in {408, 429, 500, 502, 503, 504}
    return isinstance(error, (URLError, TimeoutError, ConnectionError,
                              IncompleteRead, RemoteDisconnected, socket.timeout, ssl.SSLError))


def error_message(error):
    if isinstance(error, (TimeoutError, socket.timeout)) or 'timed out' in str(error).lower():
        return 'Сервер реестра не ответил вовремя (таймаут TLS или сети).'
    if isinstance(error, HTTPError):
        return f'Сервер реестра ответил с ошибкой HTTP {error.code}.'
    if isinstance(error, (URLError, ConnectionError, IncompleteRead, RemoteDisconnected, ssl.SSLError)):
        return 'Нет соединения с сервером реестра или он временно недоступен.'
    return 'Не удалось проверить резервный реестр. Проверьте сеть и свободное место.'


class RegistryWorker:
    def __init__(self, cache, updater=update, sleeper=time.sleep, delays=(2, 5)):
        self.cache, self.updater, self.sleeper = cache, updater, sleeper
        self.delays = delays
        self.max_attempts = len(delays) + 1
        self.lock = threading.Lock()
        self.state = {'phase': 'idle', 'attempt': 0, 'max_attempts': self.max_attempts,
                      'message': 'Проверка реестра ещё не запускалась.', 'retry_available': True}
        self.thread = None

    def snapshot(self):
        with self.lock:
            return dict(self.state)

    def change(self, **fields):
        with self.lock:
            self.state.update(fields)

    def start(self):
        with self.lock:
            if self.thread and self.thread.is_alive():
                return dict(self.state)
            self.state.update(phase='checking', attempt=1, retry_available=False,
                              message=f'Проверяем обновление реестра: попытка 1 из {self.max_attempts}.')
            self.thread = threading.Thread(target=self.run, name='registry-update', daemon=True)
            self.thread.start()
            return dict(self.state)

    def run(self):
        for attempt in range(1, self.max_attempts + 1):
            self.change(phase='checking', attempt=attempt, retry_available=False,
                        message=f'Проверяем и при необходимости загружаем реестр: попытка {attempt} из {self.max_attempts}.')
            try:
                report = self.updater(self.cache)
            except Exception as error:
                message = error_message(error)
                if retryable(error) and attempt < self.max_attempts:
                    delay = self.delays[attempt - 1]
                    self.change(phase='waiting', message=f'{message} Повтор {attempt + 1} из {self.max_attempts} через {delay} с.')
                    self.sleeper(delay)
                    continue
                self.change(phase='error', message=f'{message} Нажмите «Повторить загрузку реестра».',
                            retry_available=True)
                return
            self.change(phase='ready', message='Резервный реестр обновлён.' if report['status'] == 'updated'
                        else 'Резервный реестр проверен, обновлений нет.', retry_available=True)
            return
