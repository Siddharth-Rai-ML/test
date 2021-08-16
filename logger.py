import datetime
import sys


class Logger:
    def __init__(self, log_level='debug'):
        if not self._isValidLogLevel(log_level):
            raise Exception(f'{log_level} is not a valid log level')
        self.log_level = log_level

    def debug(self, msg):
        if self.log_level == 'debug':
            print(f'[{self._getTimestamp()}] [debug] {msg}')

    def info(self, msg):
        if self.log_level in ['debug', 'info']:
            print(f'[{self._getTimestamp()}] [info] {msg}')

    def error(self, msg, exit=False):
        if self.log_level == 'silent':
            return

        print(f'[{self._getTimestamp()}] [error] {msg}')
        if exit:
            sys.exit(int(exit))

    def setLevel(self, log_level):
        if not self._isValidLogLevel(log_level):
            raise Exception(f'{log_level} is not a valid log level')
        self.log_level = log_level

    def _getTimestamp(self):
        return str(datetime.datetime.now())

    def _isValidLogLevel(self, log_level):
        return log_level in ['debug', 'info', 'error', 'silent']


log = Logger('debug')
