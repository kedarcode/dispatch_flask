import os
import sys


class PathResource:
    def __init__(self, path):
        self.resource_path(path)

    @classmethod
    def resource_path(cls, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.dirname(__file__)
        return os.path.join(base_path, relative_path)
