import os

class _ValidatePath(str):

    """
    This class ensures that an provided file exists.
    """

    def __new__(cls, content, *args, **kwargs):
        cls._raise_value_error_if_path_does_not_exist(content)
        return str.__new__(cls, content, *args, **kwargs)

    @staticmethod
    def _raise_value_error_if_path_does_not_exist(input_path: str):
        if not os.path.exists(input_path):
            raise ValueError("The input file doesn't exist.")


class _ValidateDirectory(str):

    """
    This class ensures that the a provided directory exists. If it doesn't, it creates the directory.
    """

    def __new__(cls, content, *args, **kwargs):
        cls._create_directory_if_it_does_not_exist(content)
        return str.__new__(cls, content, *args, **kwargs)

    @staticmethod
    def _create_directory_if_it_does_not_exist(path: str):
        directory = os.path.dirname(path)
        if not os.path.exists(directory):
            os.mkdir(directory)


class _EscapedString(str):
    r"""Create a string with backslashes before special characters.

    This class will place a backslash before any of the following characters:
    |&:;()<>~*@?!$#"` '. It is particularly beneficial for creating paths to
    terminal files.

    Parameters
    ----------
    content
        Any string.

    Examples
    --------
    Make the string of a movie into its terminal representation

    >>> print(_EscapedString('/path/to/file with ?$:"'))
    /path/to/file\ with\ \?\$\:\"

    """

    def __new__(cls, content: str):
        escape_str = cls._add_backslashes_before_special_characters(content)
        obj = super().__new__(cls, escape_str)
        obj._original = content
        return obj

    @staticmethod
    def _add_backslashes_before_special_characters(string: str) -> str:
        special_characters = '\|&:;()<>~*@?!$#"` ' + "'"
        for i in special_characters:
            string = string.replace(i, rf'\{i}')
        return string

    @property
    def original(self) -> str:
        return self._original