from PyQt5.QtWidgets import *
from PyQt5.QtCore import *


class LineEditMarketPath(QLineEdit):
    """
        Класс LineEditMarketPath является наследником класса QLineEdit из библиотеки PyQt5.QtWidgets. Он представляет
    собой виджет для ввода и отображения пути к папке на компьютере.

        Конструктор класса принимает аргументы name_market и parent. Параметр name_market представляет собой имя рынка,
    для которого будет выбран путь к папке. Параметр parent указывает на родительский виджет.

        В конструкторе вызывается конструктор родительского класса с помощью super(). Затем инициализируются атрибуты
    market_path и name_market. Атрибут market_path представляет собой выбранный путь к папке, а name_market - имя рынка.

        Далее вызывается метод setReadOnly(True), который делает поле ввода только для чтения.

        Метод open_folder_dialog переопределяет событие нажатия мыши на поле ввода. При нажатии мыши открывается
    диалоговое окно выбора папки с помощью класса QFileDialog. При нажатии кнопки "Открыть" в диалоговое
    окно выбора папки с помощью класса QFileDialog. Устанавливается режим выбора только папок с помощью метода
    setFileMode(QFileDialog.Directory).

        Если пользователь выбирает папку и нажимает кнопку "Открыть" в диалоговом окне, выбранный путь к папке
    сохраняется в атрибуте market_path. Затем получается информация о выбранной папке с помощью класса QFileInfo и
    извлекается имя выбранной папки с помощью метода fileName().

        Если имя рынка равно 'all_market', то выбранный путь к папке устанавливается для всех рынков в родительском
    виджете. В противном случае, выбранный путь к папке устанавливается только для текущего рынка.

        Наконец, метод setText(market_name) устанавливает текст поля ввода равным имени выбранной папки.
    """
    def __init__(self, name_market, parent=None):
        super().__init__(parent)

        self.market_path = ""
        self.name_market = name_market

        self.setReadOnly(True)
        self.mousePressEvent = self.open_folder_dialog

    def open_folder_dialog(self, event):
        folder_dialog = QFileDialog()
        folder_dialog.setFileMode(QFileDialog.Directory)
        if folder_dialog.exec_():
            self.market_path = folder_dialog.selectedFiles()[0]
            market_info = QFileInfo(self.market_path)
            market_name = market_info.fileName()
            if self.name_market == 'all_market':
                for key in self.parent().parent().path_all.keys():
                    if key == 'file':
                        continue
                    self.parent().parent().path_all[key] = self.market_path
            else:
                self.parent().parent().path_all[self.name_market] = self.market_path
            self.setText(market_name)