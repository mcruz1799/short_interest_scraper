# Printing helpers
from colorama import init, Fore, Back, Style

def printg(*args):
    print_colored(Fore.GREEN, *args)

def printr(*args):
    print_colored(Fore.RED, *args)

def printb(*args):
    print_colored(Fore.BLUE, *args)

def print_colored(color, *args):
    print(color + ' '.join(map(str,args)) + Style.RESET_ALL)
