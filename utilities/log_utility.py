from enum import Enum

from colorama import init, Fore
from tkinter import messagebox
from tkinter import Tk     # from tkinter import Tk for Python 3.x

from openpyxl.styles.builtins import title

from configs.configs import Configs


init(autoreset=True)
Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing


class LogMessageType(Enum):
    ERROR = 1
    INFO = 2
    SUCCESS = 3

class LogUtility:

    @staticmethod
    def log_error(ex: Exception):
        LogUtility.print(str(ex), LogMessageType.ERROR)

    @staticmethod
    def log_success(message: str):
        LogUtility.print(message, LogMessageType.SUCCESS)

    @staticmethod
    def log_info(message: str):
        LogUtility.print(message, LogMessageType.INFO)

    @staticmethod
    def print(message: str, message_type: LogMessageType):

        if Configs.app_is_launched_on_console_mode:
            prefix: str | None = None

            match message_type:
                case LogMessageType.ERROR:
                    prefix = Fore.RED
                case LogMessageType.INFO:
                    prefix = ""
                case LogMessageType.SUCCESS:
                    prefix = Fore.GREEN

            message = prefix + message
            print(message)

        else:
            the_title: str | None = None
            icon: str | None = None

            match message_type:
                case LogMessageType.ERROR:
                    icon = messagebox.ERROR
                    the_title = "Erreur"
                case LogMessageType.SUCCESS | LogMessageType.INFO:
                    icon = messagebox.INFO
                    the_title = "Info"

            message += "\n" * 2 + LogUtility.format_brand_message()
            messagebox.Message(message=message, icon=icon, title=the_title).show()


    @staticmethod
    def format_brand_message():
        hyphen_len = 35
        arr = ["-" * hyphen_len, " - By Nick KINGUELEOUA - "]

        arr.append(arr[0])

        return "\n".join(arr)


