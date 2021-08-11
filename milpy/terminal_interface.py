import os
from typing import Iterable
from milpy.miscellaneous import EscapedString


def path_to_system_executable(executable: str) -> str:
    path = EscapedString(os.path.join(os.path.dirname(__file__), executable))
    os.system(f'chmod 777 {path}')
    return path


def construct_terminal_commands(command_line_arguments: Iterable[str]) -> str:
    return ' '.join(command_line_arguments)
