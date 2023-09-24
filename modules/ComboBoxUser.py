import subprocess
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *


class ComboBoxUser(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.config_file = 'config.ini'
        self.file_watcher = QFileSystemWatcher()
        self.file_watcher.fileChanged.connect(self.update_users)
        self.file_watcher.addPath(self.config_file)

        self.set_users()

    def set_users(self):
        config = QSettings(self.config_file, QSettings.IniFormat)
        config.setIniCodec("UTF-8")
        names_list = config.value("Names/names")
        if names_list:
            self.addItems(names_list)

    def update_users(self):
        self.clear()
        self.set_users()

    def open_config_file(self):
        print(self.config_file)
        config_file = self.config_file
        subprocess.Popen(["notepad.exe", config_file])