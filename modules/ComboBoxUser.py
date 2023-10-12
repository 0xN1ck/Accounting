import subprocess
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *


class ComboBoxUser(QComboBox):
    """
        Класс ComboBoxUser наследуется от класса QComboBox и представляет выпадающий список пользователей.

        Метод __init__ инициализирует объект класса. Он устанавливает путь к файлу конфигурации config_file и создает
    объект file_watcher класса QFileSystemWatcher для отслеживания изменений в этом файле. При изменении файла,
    срабатывает сигнал fileChanged, который связан с методом update_users. После этого вызывается метод
    set_users для установки пользователей в выпадающем списке.

        Метод set_users открывает файл конфигурации config.ini с помощью класса QSettings и устанавливает кодек UTF-8.
    Затем получает список имен пользователей из секции "Names/names" конфигурационного файла и добавляет
    их в выпадающий список.

        Метод update_users очищает выпадающий список и вызывает метод set_users для установки обновленного списка
    пользователей.

        Метод open_config_file открывает файл конфигурации в текстовом редакторе Notepad с помощью модуля subprocess.

        Комментарий:
        Класс ComboBoxUser представляет выпадающий список пользователей, который автоматически обновляется при изменении
    файла конфигурации. Методы set_users и update_users отвечают за установку и обновление списка пользователей,
    а метод open_config_file открывает файл конфигурации для редактирования.
    """
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