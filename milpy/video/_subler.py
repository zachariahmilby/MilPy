import os
import pandas as pd
from milpy.miscellaneous import EscapedString
from milpy.terminal_interface import construct_terminal_commands, path_to_system_executable


def path_to_subler_cli():
    return path_to_system_executable("video/anc/SublerCLI")


def format_subler_metadata_from_dictionary(metadata_dictionary):
    return ''.join([r'{%s:%s}' % (EscapedString(key),
                                  EscapedString(value).replace(',', r'\,'))
                    for key, value in metadata_dictionary.items()])


def make_temporary_video(source_filepath):
    temporary_filepath = list(os.path.splitext(source_filepath))
    temporary_filepath[-1] = '_temp' + temporary_filepath[-1]
    temporary_filepath = ''.join(temporary_filepath)
    os.system(f'mv {EscapedString(source_filepath)} {EscapedString(temporary_filepath)}')
    return EscapedString(temporary_filepath)


class SpreadsheetLoader:

    def __init__(self, path_to_spreadsheet):
        self.metadata: pd.DataFrame = pd.read_excel(path_to_spreadsheet, dtype=str)
        self.n_items = self.metadata.shape[0]

    def _get_item_metadata(self, line):
        item_df = self.metadata.iloc[line].dropna()
        keys = item_df.keys().tolist()
        values = item_df.tolist()
        return zip(keys, values)

    @staticmethod
    def _include_copyright_symbol(subler_dictionary):
        subler_dictionary['Copyright'] = f'\u00A9 {subler_dictionary["Copyright"]}'
        return subler_dictionary

    @staticmethod
    def _handbrake_dictionary_keys():
        """These are the spreadsheet components specific to Handbrake."""
        return ["Source", "Destination", "Title", "Audio", "Dimensions",
                "Crop", "Audio Bitrate", "Audio Mixdown", "Audio Track Names",
                "Subtitle", "Chapters", "Quality Factor"]

    def make_subler_dictionary(self, line):
        subler_dictionary = {key: value
                             for key, value in self._get_item_metadata(line)
                             if key not in self._handbrake_dictionary_keys()}
        return self._include_copyright_symbol(subler_dictionary)

    def make_handbrake_dictionary(self, line):
        return {key: value
                for key, value in self._get_item_metadata(line)
                if key in self._handbrake_dictionary_keys()}
