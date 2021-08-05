import os
import pandas as pd
from milpy.miscellaneous import _EscapedString, _ValidateDirectory
from milpy.video._handbrake import SourceOptions, DestinationOptions, VideoOptions, AudioOptions, PictureOptions, SubtitleOptions
from milpy.parallel_processing import get_appropriate_number_of_cores, set_processor_pool, cleanup_parallel_processing
from milpy.terminal_interface import path_to_system_executable, construct_terminal_commands
from milpy.video._spreadsheet_creation import _columns_for_kind, _make_dataframe_from_columns


class Video:

    def __init__(self, source_filepath: str):

        self.source = _EscapedString(source_filepath)
        self.source_type: str = self._check_source_type()
        self.converter_options: _VideoConverter = None
        self.destination: _EscapedString = None
        self._intermediate_file: _EscapedString = None
        self._handbrake_cli = path_to_system_executable("video/anc/HandBrakeCLI")
        self._handbrake_options = None
        self._subler_cli = path_to_system_executable("video/anc/SublerCLI")
        self._subler_options = None

    def _check_source_type(self) -> str:
        source_file_extension = os.path.splitext(self.source)[1]
        source_type = 'video'
        if source_file_extension == '.iso':
            source_type = 'disc image'
        return source_type

    def _check_if_converter_set(self):
        if self.converter_options is None:
            raise Exception('Set converter options before trying to convert a video.')

    def initialize_converter_options(self, destination_filepath):

        """
        This method sets initial parameters for video conversion. To change the defaults, access via
        `Video.converter_options.<sub_option>.<parameter>`. You can also print a summary with
        `print(Video.converter_options)`.
        """

        self.destination = _EscapedString(destination_filepath)
        _ValidateDirectory(self.destination.original)
        self._intermediate_file = _EscapedString(self.destination.original + 'temp')
        self.converter_options = _VideoConverter(self.source.original, self.destination.original)

    def convert(self):

        """
        Converts source video.
        """

        self._check_if_converter_set()
        self.converter_options.convert()

    def test_convert(self):

        """
        Tests video conversion output by quick-converting a 10-second clip.
        """

        self._check_if_converter_set()
        self.converter_options.test()

    def inspect_metadata(self):

        """
        This method prints the existing metadata in an MP4 file.
        """

        if self.source_type == "disc image":
            raise Exception('Disc images do not have metadata.')

        options = [self._subler_cli,
                   f"-source {self.source}",
                   f"-listmetadata"]
        os.system(construct_terminal_commands(options))

    def tag(self):

        """
        This method tags a video with metadata.
        """

        print(f'Tagging \"{os.path.basename(self.destination.original)}\"...')
        options = [self._subler_cli,
                   f"-source {self.source}",
                   f"-dest {self.destination}",
                   f"-metadata {self._subler_options}",
                   f"-language English"]
        os.system(construct_terminal_commands(options))

    def remove_source(self):
        os.remove(self.source.original)


class _VideoConverter:

    """
    This class sets up HandBrakeCLI command-line flags and provides a method for initiating video conversion using the
    stored parameters. These parameters can be changed after creating an instance of the class. The default behavior is
    to convert the video with the same parameters as the source.
    """

    def __init__(self, input_filepath: str, output_filepath: str):

        self.handbrake_cli = path_to_system_executable("video/anc/HandBrakeCLI")
        self.source_options = SourceOptions(input_filepath)
        self.destination_options = DestinationOptions(output_filepath)
        self.video_options = VideoOptions()
        self.audio_options = AudioOptions()
        self.picture_options = PictureOptions()
        self.subtitle_options = SubtitleOptions()
        self.options = None

    def __str__(self):
        print_string = "Video converter settings:\n"
        print_string += "-" * (len(print_string) - 2) + "\n"
        print_string += self.source_options.__str__() + "\n"
        print_string += self.destination_options.__str__() + "\n"
        print_string += self.video_options.__str__() + "\n"
        print_string += self.audio_options.__str__() + "\n"
        if self.subtitle_options.subtitles != "None":
            print_string += self.picture_options.__str__() + "\n"
            print_string += self.subtitle_options.__str__()
        else:
            print_string += self.picture_options.__str__()
        return print_string

    def _run_conversion_in_terminal(self):
        os.system(construct_terminal_commands(self.options))

    def _check_if_input_and_output_match(self):
        if self.source_options.input.original == self.destination_options.output.original:
            raise Exception('Input and output files cannot be the same! Either change the output directory or give the '
                            'input file a temporary filename.')

    def convert(self):

        self._set_options_list()
        self._check_if_input_and_output_match()
        self._run_conversion_in_terminal()

    def _set_test_video_options(self):
        self.video_options.encoder = 'x264'
        self.video_options.speed = 'ultrafast'

    def _set_options_list(self):
        self.options = [self.handbrake_cli,
                        self.source_options.construct_terminal_commands(),
                        self.destination_options.construct_terminal_commands(),
                        self.video_options.construct_terminal_commands(),
                        self.audio_options.construct_terminal_commands(),
                        self.picture_options.construct_terminal_commands(),
                        self.subtitle_options.construct_terminal_commands()]

    def _set_test_video_length(self):
        self.options.append('--start-at=seconds:0')
        self.options.append('--stop-at=seconds:10')

    def test(self):
        self._set_test_video_options()
        self._set_options_list()
        self._set_test_video_length()
        self._check_if_input_and_output_match()
        self._run_conversion_in_terminal()


class Spreadsheet:

    """
    This class allows for conversion and tagging of a video using parameters in a spreadsheet.
    """

    def __init__(self, path_to_spreadsheet):

        """

        Parameters
        ----------
        path_to_spreadsheet
            The absolute path to the spreadsheet containing the information
        """

        self.metadata = _MetadataLoader(path_to_spreadsheet)
        self.n_items = len(self.metadata.data)
        self._keys = self.metadata.get_keys()

    def get_item_metadata(self, line):
        return self.metadata.get_values(line)

    def get_keys(self):
        return self._keys

    @staticmethod
    def _include_copyright_symbol(subler_dictionary):
        subler_dictionary['Copyright'] = f'\u00A9 {subler_dictionary["Copyright"]}'
        return subler_dictionary

    @staticmethod
    def _handbrake_dictionary_keys():
        return ["Source", "Destination", "Title", "Audio", "Dimensions", "Crop", "Audio Bitrate", "Audio Mixdown",
                "Audio Notes", "Subtitle", "Chapters", "Quality Factor"]

    def _make_subler_dictionary(self, line):
        subler_dictionary = {key: value for key, value in zip(self.get_keys(), self.get_item_metadata(line))
                             if key not in self._handbrake_dictionary_keys()}
        return self._include_copyright_symbol(subler_dictionary)

    def _make_handbrake_dictionary(self, line):
        return {key: value for key, value in zip(self.get_keys(), self.get_item_metadata(line))
                if key in self._handbrake_dictionary_keys()}

    def _format_subler_metadata_from_dictionary(self, line):
        return ''.join([r'{"%s":"%s"}' % (key, value) for key, value in self._make_subler_dictionary(line).items()])

    def _initialize_converter(self, line):
        handbrake_metadata = self._make_handbrake_dictionary(line)
        video = Video(handbrake_metadata["Source"])
        video.initialize_converter_options(handbrake_metadata["Destination"])
        video.converter_options.source_options.title = handbrake_metadata["Title"]
        video.converter_options.video_options.quality = handbrake_metadata["Quality Factor"]
        video.converter_options.audio_options.audio_titles = handbrake_metadata["Audio"]
        video.converter_options.audio_options.bitrates = handbrake_metadata["Audio Bitrate"]
        video.converter_options.audio_options.mixdowns = handbrake_metadata["Audio Mixdown"]
        video.converter_options.audio_options.names = handbrake_metadata["Audio Notes"]
        video.converter_options.picture_options.width = handbrake_metadata["Dimensions"].split("x")[0]
        video.converter_options.picture_options.height = handbrake_metadata["Dimensions"].split("x")[1]
        video.converter_options.picture_options.crop = handbrake_metadata["Crop"]
        return video











    def _convert_spreadsheet_item(self, line):
        video = self._initialize_converter(line)
        video.convert()

    def _test_convert_spreadsheet_item(self, line):
        video = self._initialize_converter(line)
        video.test_convert()

    def _convert_and_tag_spreadsheet_item(self, line):
        video = self._initialize_converter(line)
        self._set_temporary_destination(video)
        video.convert()
        self._set_temporary_source(video)
        video.tag()

    def _test_convert_and_tag_spreadsheet_item(self, line):
        video = self._initialize_converter(line)
        self._set_temporary_destination(video)
        video.test_convert()
        self._set_temporary_source(video)
        video.tag()

    @staticmethod
    def _get_multiprocessing_pool():
        n_cores = get_appropriate_number_of_cores()
        pool = set_processor_pool(n_cores)
        return pool

    def test_convert_and_tag(self):
        """
        Tests the spreadsheet for filepaths and tags.
        """
        pool = self._get_multiprocessing_pool()
        [pool.apply_async(self._test_convert_and_tag_spreadsheet_item, args=(item,)).get() for item in range(self.n_items)]
        cleanup_parallel_processing(pool)

    def parallel_convert_and_tag(self):
        """
        Convert spreadsheet using parallel processing (one core per video, but multiple videos at once). This is best
        for 1080p or lower resolution.
        """
        pool = self._get_multiprocessing_pool()
        [pool.apply_async(self._convert_and_tag_spreadsheet_item, args=(i,)) for i in range(self.n_items)]
        cleanup_parallel_processing(pool)

    def serial_convert_and_tag(self):
        """
        Convert spreadsheet using serial processing (multiple cores per video, but only a single video at a time). This
        is better for 4K videos.
        """
        for i in range(self.n_items):
            self._convert_and_tag_spreadsheet_item(i)

    def _tag_spreadsheet_item(self, line):
        video = Video(self.metadata.data['Source'][line])
        video.destination = _EscapedString(self.metadata.data['Destination'][line])
        _ValidateDirectory(video.destination.original)
        video._subler_options = self._format_subler_metadata_from_dictionary(line)
        video.tag()

    def tag_only(self):
        """
        This assumes the sources listed in the spreadsheet don't need conversion, just metadata tagging.
        """
        pool = self._get_multiprocessing_pool()
        [pool.apply_async(self._tag_spreadsheet_item, args=(i,)).get() for i in range(self.n_items)]

    # add parallel and serial convert only


class _MetadataLoader:

    """This class loads a Microsoft Excel metadata spreadsheet and returns the keys and values for a particular entry."""

    def __init__(self, path_to_excel_spreadsheet: str):
        self.data = pd.read_excel(path_to_excel_spreadsheet, dtype=str)

    def get_keys(self) -> list:
        """Get a list of the column names from the spreadsheet."""
        return self.data.keys().tolist()

    def get_values(self, line: int) -> list:
        """Get the line number you want, (index starting from 0)."""
        return self.data.iloc[line].tolist()


if __name__ == "__main__":
    sheet = Spreadsheet('/Volumes/Media HD/Disc Images/Frasier/metadata.xlsx')
    sheet.test_convert_and_tag()


def make_empty_metadata_spreadsheet(save_directory: str, kind: str):

    """
    This function saves an empty Microsoft Excel metadata spreadsheet for you to fill in. Note that not all fields need
    to be filled out for each item or even an entire spreadsheet, but are the recommended fields for the video type.

    Parameters
    ----------
    save_directory
        The directory in which you want the empty spreadsheet saved.
    kind
        The kind of spreadsheet (choose between "tv" or "movie").

    Examples
    --------
    For a TV spreadsheet:

    >>> make_empty_metadata_spreadsheet("/path/to/directory", kind='TV')
    Empty spreadsheet saved to "/path/to/directory/empty_metadata_tv.xlsx."

    For a movie spreadsheet:

    >>> make_empty_metadata_spreadsheet("/path/to/directory", kind='movie')
    Empty spreadsheet saved to "/path/to/directory/empty_metadata_movie.xlsx."

    Notes
    -----
    TV Shows:
     - **Name:** The episode title.
     - **Artist:** The name of the series.
     - **Album Artist:** The name of the series.
     - **Album:** The name of the series and the season number, e.g., "Star Trek, Season 1."
     - **Genre:** Your choice, e.g., "Science Fiction" or "Situation Comedy."
     - **Release Date:** The date of first broadcast in format YYYY-MM-DD. If it's a special feature, I usually do the
       release date of the media I'm getting the video from, like the DVD or Blu-ray release date.
     - **Track #:** For episodes, I use the format #/total, so episode 7 in a season of 22 episodes would be "7/22". For
       special features, I start at 101, so special feature 3 out of 12 would be "103/112".
     - **TV Show:** The name of the series.
     - **TV Episode ID:** The internal production ID for the episode (not all series have one).
     - **TV Season:** The season number (just the number).
     - **TV Episode #:** The episode number (just the number; if multiple episodes combined into one, I guess just choose
       the number of the first episode).
     - **TV Network:** The network on which the show originall broadcast. For special features, I usually put "DVD" or
       "Blu-ray" depending on the source.
     - **Description:** The episode description.
     - **Series Description:** A general series description.
     - **Copyright:** The copyright information, e.g., "Production House, LLC. All Rights Reserved." The copyright symbol
       is added automatically behind-the-scenes.
     - **Media Kind:** 10 (this indicates a TV show)
     - **Cover Art:** The absolute path to the artwork you want for this item.
     - **Rating:** For the US, the options are listed below (they are a little hard to find). For a series from another
       country, make a sample video file, set the rating manually with the Subler GUI, then use the video_inspector()
       function to see what the correct format is.

          - us-tv|TV-Y|100|
          - us-tv|TV-Y7|200|
          - us-tv|TV-G|300|
          - us-tv|TV-PG|400|
          - us-tv|TV-14|500|
          - us-tv|TV-MA|600|
          - us-tv|Unrated|???|

     - **Cast:** The names of the main cast members, separated by commas (for names with commas, surround with quotation
       marks like, e.g., Patrick Stewart, "Patrick Stewart, Jr."
     - **Source:** The absolute path to the source, either a disc image `.iso` or a video file.
     - **Destination:** The absolute path to where you want the final video to be saved. I usually prefix the file with it's
       number, e.g., /.../07 Episode 7.mp4 for a main episode or /.../Special Features/103 Special Feature.mp4
     - **Title:** The video title, meaning which video track to convert. Discs have multiple, individual videos are just
       title "1".
     - **Dimensions:** The dimensions of the final video separated by an "x", e.g., "854x480" or "640x480" for SD.
     - **Crop:** Some videos files have the letterbox black bars encoded. Crop can remove them if you know the relative
       sizes. The format is #:#:#:#, with the numbers referring to <top:bottom:left:right>.
     - **Audio:** The audio title(s), separated by a comma. You might want multiple audio tracks (like a commentary).
     - **Audio Bitrate:** The bitrates for each of the audio titles, using 64-bits per channel. This means 64 for mono,
       128 for stereo, 384 for 5.1 surround sound, and 512 for 7.1 surround sound.
     - **Audio Mixdown:** The mixdowns for the audio titles. Options are mono, stereo, 5point1 or 7point1.
     - **Audio Notes:** What to name the audio tracks, like "Mono Audio", "Stereo Audio", "5.1 Surround Sound",
       "7.1 Surround Sound", or "Commentary with XXX and XXX". If the name has commas in it, it will need to be in
       quotation marks.
     - **HD Video:** Video definition flag.

         - "0" for 480p/576p Standard Definition
          - "1" for 720p High Definition
          - "2" for 1080p Full High Definition
          - "3" for 2160p 4K Ultra High Definition

     - **Quality Factor:** Handbrake quality factor.

          - 20±2 for 480p/576p Standard Definition
          - 21±2 for 720p High Definition
          - 22±2 for 1080p Full High Definition
          - 25±2 for 2160p 4K Ultra High Definition

     - **Subtitle:** If you want to hard-burn a subtitle track, put its number here.
     - **Chapters:** If you want to name the chapters, you need to provide the absolute path to a CSV file containing
       those names. The format should be "#,name# for each chapter in the video. For instance, the third line in the CSV
       might be "3,The Third Chapter".

    Movies:
     - **Name:** The name of the movie.
     - **Genre:** Your choice, e.g., "Science Fiction" or "Comedy."
     - **Release Date:** The date of the film in format YYYY-MM-DD. If it's a special feature, I usually do the release
       date of the media I'm getting the video from, like the DVD or Blu-ray release date.
     - **Description:** The episode description.
     - **Copyright:** The copyright information, e.g., "Production House, LLC. All Rights Reserved." The copyright symbol
       is added automatically behind-the-scenes.
     - **Media Kind:** 9 (this indicates a movie)
     - **Cover Art:** The absolute path to the artwork you want for this item.
     - **Rating:** For the US, the options are listed below (they are a little hard to find). For a series from another
       country, make a sample video file, set the rating manually with the Subler GUI, then use the video_inspector()
       function to see what the correct format is.

          - mpaa|NR|000|
          - mpaa|G|100|
          - mpaa|PG|200|
          - mpaa|PG-13|300|
          - mpaa|R|400|
          - mpaa|NC-17|500|
          - mpaa|Unrated|???|

     - **Rating Annotation:** any reasons listed for the given rating, e.g., "violence and sexual content."
     - **Cast:** The names of the main cast members, separated by commas (for names with commas, surround with quotation
       marks like, e.g., Patrick Stewart, "Patrick Stewart, Jr."
     - **Director:** The names of the director(s), same comma considerations as with cast.
     - **Producers:** The names of the producer(s), same comma considerations as with cast.
     - **Screenwriters:** The names of the screenwriter(s), same comma considerations as with cast.
     - **Source:** The absolute path to the source, either a disc image `.iso` or a video file.
     - **Destination:** The absolute path to where you want the final video to be saved. I don't usually prefix the main
       movie file, e.g., /.../Film Name.mp4, but for a special feature /.../Special Features/103 Special Feature.mp4
     - **Title:** The video title, meaning which video track to convert. Discs have multiple, individual videos are just
       title "1".
     - **Dimensions:** The dimensions of the final video separated by an "x", e.g., "854x480" or "640x480" for SD.
     - **Crop:** Some videos files have the letterbox black bars encoded. Crop can remove them if you know the relative
       sizes. The format is #:#:#:#, with the numbers referring to <top:bottom:left:right>.
     - **Audio:** The audio title(s), separated by a comma. You might want multiple audio tracks (like a commentary).
     - **Audio Bitrate:** The bitrates for each of the audio titles, using 64-bits per channel. This means 64 for mono,
       128 for stereo, 384 for 5.1 surround sound, and 512 for 7.1 surround sound.
     - **Audio Mixdown:** The mixdowns for the audio titles. Options are mono, stereo, 5point1 or 7point1.
     - **Audio Notes:** What to name the audio tracks, like "Mono Audio", "Stereo Audio", "5.1 Surround Sound",
       "7.1 Surround Sound", or "Commentary with XXX and XXX". If the name has commas in it, it will need to be in
       quotation marks. If it has quotation marks it might break entirely. Best to check in advance or be sure to use
       single quotes.
     - **HD Video:** Video definition flag.

          - "0" for 480p/576p Standard Definition
          - "1" for 720p High Definition
          - "2" for 1080p Full High Definition
          - "3" for 2160p 4K Ultra High Definition

     - **Quality Factor:** Handbrake quality factor.

          - 20±2 for 480p/576p Standard Definition
          - 21±2 for 720p High Definition
          - 22±2 for 1080p Full High Definition
          - 25±2 for 2160p 4K Ultra High Definition

     - **Subtitle:** If you want to hard-burn a subtitle track, put its number here.
     - **Chapters:** If you want to name the chapters, you need to provide the absolute path to a CSV file containing
       those names. The format should be "#,name# for each chapter in the video. For instance, the third line in the CSV
       might be "3,The Third Chapter".
    """

    kind, columns = _columns_for_kind(kind)
    save_path = os.path.join(save_directory, f"empty_metadata_{kind}.xlsx")
    _make_dataframe_from_columns(columns).to_excel(save_path, index=False)

    print(f"Empty spreadsheet saved to \"{save_path}\".")
