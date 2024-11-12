from colorama import init, Fore, Back

init()


class LogUtility:

    @staticmethod
    def log_error(ex: Exception):
        print(Fore.RED + str(ex))

    @staticmethod
    def log_success(message: str):
        print(Fore.GREEN + message)

    @staticmethod
    def log_info(message: str):
        print(Fore.CYAN + message)
