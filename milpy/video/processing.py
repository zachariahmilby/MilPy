import os
from pathlib import Path
from milpy.miscellaneous import EscapedString
from milpy.video._handbrake import path_to_handbrake_cli, \
    VideoConversionOptions
from milpy.video._subler import path_to_subler_cli, \
    format_subler_metadata_from_dictionary, make_temporary_video, \
    SpreadsheetLoader
from milpy.terminal_interface import construct_terminal_commands
from milpy.parallel_processing import get_multiprocessing_pool, \
    cleanup_parallel_processing
from milpy.video._spreadsheet_creation import _columns_for_kind, \
    _make_dataframe_from_columns
import subprocess


# TODO: Figure out why double quotes won't appear in HandbrakeCLI audio track
#  names.


class _VideoConverter:
    def __init__(self, converter_options: VideoConversionOptions):
        self.options = converter_options

    def _create_command_of_input_options(self):
        return f'{path_to_handbrake_cli()} ' + \
               f'{repr(self.options.source)} ' + \
               f'{repr(self.options.destination)} ' + \
               f'{repr(self.options.video)} ' + \
               f'{repr(self.options.audio)} ' + \
               f'{repr(self.options.picture)} ' + \
               f'{repr(self.options.subtitle)} '

    def _set_test_video_options(self):
        self.options.video.encoder = 'x264'
        self.options.video.speed = 'ultrafast'

    @staticmethod
    def _add_30_second_test_to_options(options: str):
        return f'{options} ' + '--start-at=seconds:0 ' + '--stop-at=seconds:30'

    def convert(self):
        command = self._create_command_of_input_options()
        os.system(command)

    def test(self):
        self._set_test_video_options()
        options = self._create_command_of_input_options()
        options = self._add_30_second_test_to_options(options)
        os.system(options)


class _File(str):

    def __new__(cls, path: str, extension: str, *args, **kwargs):

        """
        Instances of this class represent a file that exists on this computer.

        Parameters
        ----------
        path
            The absolute path to a file on this computer.
        extension
            The extension the input path must have.

        Raises
        ------
        ValueError
            Raised if the input file path points to a file that does not exist.
        TypeError
            Raised if the input file path does not have the specified extension.
        """

        cls._raise_value_error_if_path_does_not_exist(path)
        cls._raise_type_error_if_input_has_wrong_extension(path, extension)
        return super().__new__(cls, path, *args, **kwargs)

    @staticmethod
    def _raise_value_error_if_path_does_not_exist(path: str):
        if not Path(path).exists():
            message = 'The input file path does not exist.'
            raise ValueError(message)

    @staticmethod
    def _raise_type_error_if_input_has_wrong_extension(path: str, ext: str):
        if Path(path).suffix != f'.{ext}':
            message = f'The input file path is not a {ext} file.'
            raise TypeError(message)


class AVI:
    """
    Instances of this class represent an MP4 file which is accessible to
    the computer.
    """

    def __init__(self, file_path: str):
        """
        Parameters
        ----------
        file_path
            The absolute path to a AVI file.

        Raises
        ------
        ValueError
            Raised if the input file path points to a file that does not exist
            or is inaccessible to the computer.
        TypeError
            Raised if the input file path does not have a `.avi` extension.
        """
        self.file_path = _File(file_path, 'avi')
        self.terminal_file_path = EscapedString(file_path)
        self.converter_options = self._set_converter_options()

    def _set_converter_options(self):
        file_extension = Path(self.file_path).suffix
        destination = self.file_path.replace(file_extension,
                                             '_converted' + file_extension)
        converter_options = VideoConversionOptions(self.file_path, destination)
        return converter_options

    def inspect_metadata(self):
        """
        Examine any existing metadata in an MP4 file.
        """
        command = f'{path_to_subler_cli()} -source {self.terminal_file_path} -listmetadata'
        os.system(command)

    def convert(self):
        """
        Convert the video using the parameters in the `converter_options`
        attribute.
        """
        converter = _VideoConverter(self.converter_options)
        converter.convert()

    def test_convert(self):
        """
        Run a test video conversion using 10 seconds at high speed.
        """
        converter = _VideoConverter(self.converter_options)
        converter.test()

    def tag(self, metadata_dictionary: dict):
        """
        Tag the video with the keys and values in the provided dictionary.

        Parameters
        ----------
        metadata_dictionary
            The metadata tags and values.
        """
        temporary_filepath = make_temporary_video(self.file_path)
        metadata = format_subler_metadata_from_dictionary(metadata_dictionary)
        options = [path_to_subler_cli(),
                   f"-source {temporary_filepath}",
                   f"-dest {EscapedString(self.file_path)}",
                   f"-metadata {metadata}",
                   f"-language English"]
        os.system(construct_terminal_commands(options))
        os.remove(temporary_filepath.original)

    def get_number_of_chapters(self):
        cmd = f"{path_to_handbrake_cli()} --input={self.terminal_file_path} " \
              f"--title={self.converter_options.source.title} --scan"
        cp = subprocess.Popen(cmd, stderr=subprocess.PIPE,
                              stdout=subprocess.PIPE, shell=True)
        chapters = [line.decode("utf-8").replace('\n', '').strip()
                    for line in cp.stderr.readlines()
                    if ": duration " in line.decode("utf-8")
                    and '00:00:01' not in line.decode("utf-8")]
        number_of_chapters = len(chapters)
        cp.stderr.close()
        cp.stdout.close()
        cp.wait()
        return str(number_of_chapters)


class MP4:
    """
    Instances of this class represent an MP4 file which is accessible to
    the computer.
    """

    def __init__(self, file_path: str):
        """
        Parameters
        ----------
        file_path
            The absolute path to a MP4 file.

        Raises
        ------
        ValueError
            Raised if the input file path points to a file that does not exist
            or is inaccessible to the computer.
        TypeError
            Raised if the input file path does not have a `.mp4` extension.
        """
        self.file_path = _File(file_path, 'mp4')
        self.terminal_file_path = EscapedString(file_path)
        self.converter_options = self._set_converter_options()

    def _set_converter_options(self):
        file_extension = Path(self.file_path).suffix
        destination = self.file_path.replace(file_extension,
                                             '_converted' + file_extension)
        converter_options = VideoConversionOptions(self.file_path, destination)
        return converter_options

    def inspect_metadata(self):
        """
        Examine any existing metadata in an MP4 file.
        """
        command = f'{path_to_subler_cli()} -source {self.terminal_file_path} -listmetadata'
        os.system(command)

    def convert(self):
        """
        Convert the video using the parameters in the `converter_options`
        attribute.
        """
        converter = _VideoConverter(self.converter_options)
        converter.convert()

    def test_convert(self):
        """
        Run a test video conversion using 10 seconds at high speed.
        """
        converter = _VideoConverter(self.converter_options)
        converter.test()

    def tag(self, metadata_dictionary: dict):
        """
        Tag the video with the keys and values in the provided dictionary.

        Parameters
        ----------
        metadata_dictionary
            The metadata tags and values.
        """
        temporary_filepath = make_temporary_video(self.file_path)
        metadata = format_subler_metadata_from_dictionary(metadata_dictionary)
        options = [path_to_subler_cli(),
                   f"-source {temporary_filepath}",
                   f"-dest {EscapedString(self.file_path)}",
                   f"-metadata {metadata}",
                   f"-language English"]
        os.system(construct_terminal_commands(options))
        os.remove(temporary_filepath.original)

    def get_number_of_chapters(self):
        cmd = f"{path_to_handbrake_cli()} --input={self.terminal_file_path} " \
              f"--title={self.converter_options.source.title} --scan"
        cp = subprocess.Popen(cmd, stderr=subprocess.PIPE,
                              stdout=subprocess.PIPE, shell=True)
        chapters = [line.decode("utf-8").replace('\n', '').strip()
                    for line in cp.stderr.readlines()
                    if ": duration " in line.decode("utf-8")
                    and '00:00:01' not in line.decode("utf-8")]
        number_of_chapters = len(chapters)
        cp.stderr.close()
        cp.stdout.close()
        cp.wait()
        return str(number_of_chapters)


class MKV:
    """
    Instances of this class represent an MKV file which is accessible to
    the computer.
    """

    def __init__(self, file_path: str):
        """
        Parameters
        ----------
        file_path
            The absolute path to an MKV file.

        Raises
        ------
        ValueError
            Raised if the input file path points to a file that does not exist
            or is inaccessible to the computer.
        TypeError
            Raised if the input file path does not have a `.mkv` extension.
        """
        self.file_path = _File(file_path, 'mkv')
        self.terminal_file_path = EscapedString(file_path)
        self.converter_options = self._set_converter_options()

    def _set_converter_options(self):
        file_extension = Path(self.file_path).suffix
        destination = self.file_path.replace(file_extension, '_converted.mp4')
        converter_options = VideoConversionOptions(self.file_path, destination)
        return converter_options

    def convert(self):
        """
        Convert the video using the parameters in the `converter_options`
        attribute.
        """
        converter = _VideoConverter(self.converter_options)
        converter.convert()

    def test_convert(self):
        """
        Run a test video conversion using 10 seconds at high speed.
        """
        converter = _VideoConverter(self.converter_options)
        converter.test()

    def get_number_of_chapters(self):
        cmd = f"{path_to_handbrake_cli()} --input={self.terminal_file_path} " \
              f"--title={self.converter_options.source.title} --scan"
        cp = subprocess.Popen(cmd, stderr=subprocess.PIPE,
                              stdout=subprocess.PIPE, shell=True)
        chapters = [line.decode("utf-8").replace('\n', '').strip()
                    for line in cp.stderr.readlines()
                    if ": duration " in line.decode("utf-8")
                    and '00:00:01' not in line.decode("utf-8")]
        number_of_chapters = len(chapters)
        cp.stderr.close()
        cp.stdout.close()
        cp.wait()
        return str(number_of_chapters)


class ISO:
    """
    Instances of this class represent a Blu-ray or DVD disc image in ISO
    format.
    """

    def __init__(self, file_path: str):
        """
        Parameters
        ----------
        file_path
            The absolute path to an ISO file on this computer.

        Raises
        ------
        ValueError
            Raised if the input file path points to a file that does not exist
            or is inaccessible to the computer.
        TypeError
            Raised if the input file path does not have a `.iso` extension.
        """
        self.file_path = _File(file_path, 'iso')
        self.terminal_file_path = EscapedString(file_path)
        self.converter_options = self._set_converter_options()

    def _set_converter_options(self):
        file_extension = Path(self.file_path).suffix
        destination = self.file_path.replace(file_extension,
                                             ', Title 1 (Converted).mp4')
        converter_options = VideoConversionOptions(self.file_path, destination)
        return converter_options

    def convert(self):
        """
        Convert the video using the parameters in the `converter_options`
        attribute.
        """
        converter = _VideoConverter(self.converter_options)
        converter.convert()

    def test_convert(self):
        """
        Run a test video conversion using 10 seconds at high speed.
        """
        converter = _VideoConverter(self.converter_options)
        converter.test()

    def get_number_of_chapters(self):
        cmd = f"{path_to_handbrake_cli()} --input={self.terminal_file_path} " \
              f"--title={self.converter_options.source.title} --scan"
        cp = subprocess.Popen(cmd, stderr=subprocess.PIPE,
                              stdout=subprocess.PIPE, shell=True)
        chapters = [line.decode("utf-8").replace('\n', '') for line in
                    cp.stderr.readlines() if "    + " in line.decode("utf-8")
                    and ": duration " in line.decode("utf-8")
                    and '00:00:01' not in line.decode("utf-8")]
        number_of_chapters = len(chapters)
        cp.stderr.close()
        cp.stdout.close()
        cp.wait()
        return str(number_of_chapters)


class Spreadsheet:
    """
    This class provides the ability to process and tag items from a
    spreadsheet.
    """

    def __init__(self, path_to_spreadsheet: str):
        """
        Parameters
        ----------
        path_to_spreadsheet
            The absolute path to the spreadsheet. To make a blank spreadsheet,
            use the function `make_empty_metadata_spreadsheet()`.
        """
        self.spreadsheet = SpreadsheetLoader(path_to_spreadsheet)

    @staticmethod
    def _get_source_type(source_filepath: str):
        extension = os.path.splitext(source_filepath)[1]
        if extension == ".mp4":
            return MP4(source_filepath)
        elif extension == ".mkv":
            return MKV(source_filepath)
        elif extension == ".iso":
            return ISO(source_filepath)
        elif extension == ".avi":
            return AVI(source_filepath)
        else:
            raise Exception('Bad input file type.')

    def _set_source_parameters(self, item: int):
        handbrake_metadata = self.spreadsheet.make_handbrake_dictionary(item)
        video = self._get_source_type(handbrake_metadata["Source"])
        video.converter_options.source.title = handbrake_metadata["Title"]
        video.converter_options.source.quality = handbrake_metadata["Quality Factor"]
        video.converter_options.destination.output = handbrake_metadata["Destination"]
        if "Chapters" in handbrake_metadata:
            video.converter_options.destination.chapters = EscapedString(handbrake_metadata["Chapters"])
        else:
            video.converter_options.destination.chapters = video.get_number_of_chapters()
        video.converter_options.audio.audio_titles = handbrake_metadata["Audio"]
        video.converter_options.audio.bitrates = handbrake_metadata["Audio Bitrate"]
        video.converter_options.audio.mixdowns = handbrake_metadata["Audio Mixdown"]
        video.converter_options.audio.track_names = handbrake_metadata["Audio Track Names"]
        video.converter_options.picture.width = handbrake_metadata["Dimensions"].split("x")[0]
        video.converter_options.picture.height = handbrake_metadata["Dimensions"].split("x")[1]
        video.converter_options.picture.crop = handbrake_metadata["Crop"]
        if "Subtitles" in handbrake_metadata:
            video.converter_options.subtitle.subtitles = handbrake_metadata["Subtitles"]
        return video

    def _parallel_convert(self, item):
        video = self._set_source_parameters(item)
        video.convert()

    def _test_parallel_convert(self, item):
        video = self._set_source_parameters(item)
        video.test_convert()

    def _parallel_tag(self, item):
        handbrake_dictionary = self.spreadsheet.make_handbrake_dictionary(item)
        metadata_dictionary = self.spreadsheet.make_subler_dictionary(item)
        MP4(handbrake_dictionary["Destination"]).tag(metadata_dictionary)

    def _parallel_convert_and_tag(self, item):
        video = self._set_source_parameters(item)
        video.convert()
        handbrake_dictionary = self.spreadsheet.make_handbrake_dictionary(item)
        metadata_dictionary = self.spreadsheet.make_subler_dictionary(item)
        MP4(handbrake_dictionary["Destination"]).tag(metadata_dictionary)

    def _test_parallel_convert_and_tag(self, item):
        video = self._set_source_parameters(item)
        video.test_convert()
        handbrake_dictionary = self.spreadsheet.make_handbrake_dictionary(item)
        metadata_dictionary = self.spreadsheet.make_subler_dictionary(item)
        MP4(handbrake_dictionary["Destination"]).tag(metadata_dictionary)

    def serial_convert(self):
        """
        Convert the source to destination using the Handbrake parameters for
        each item. This converts the items one at a time using as many cores as
        your computer decides to allocate. I find serial conversion is better
        for a set of high-definition videos, while parallel conversion is
        better for a set of standard-definition videos.
        """
        for item in range(self.spreadsheet.n_items):
            video = self._set_source_parameters(item)
            video.convert()

    def parallel_convert(self, n_cores: int = None):
        """
        Convert the source to destination using the Handbrake parameters for
        each item. This converts multiple items at once using as many cores as
        your computer decides to allocate. I find parallel conversion is
        better for a set of standard-definition videos, while serial conversion
        is better for a set of high-definition videos.

        Parameters
        ----------
        n_cores
            If desired, the user-specified number of cores to use.
        """
        pool = get_multiprocessing_pool(n_cores)
        for item in range(self.spreadsheet.n_items):
            pool.apply_async(self._parallel_convert, args=(item,))
        cleanup_parallel_processing(pool)

    def test_convert(self):
        """
        Test conversion of a spreadsheet. This is done using parallel
        processing.
        """
        pool = get_multiprocessing_pool()
        for item in range(self.spreadsheet.n_items):
            pool.apply_async(self._test_parallel_convert, args=(item,))
        cleanup_parallel_processing(pool)

    def tag(self):
        """
        Assuming all source items are MP4 files which only need tagging, tag
        them in parallel using the Subler parameters for each item. This uses
        the "Destination" parameter as the source and destination file.
        """
        pool = get_multiprocessing_pool()
        for item in range(self.spreadsheet.n_items):
            pool.apply_async(self._parallel_tag, args=(item,))
        cleanup_parallel_processing(pool)

    def serial_convert_and_tag(self):
        """
        Convert the source to destination using the Handbrake parameters then
        tag with Subler parameters for each item. This converts the items one
        at a time using as many cores as your computer decides to allocate. I
        find serial conversion is better for a set of high-definition videos,
        while parallel conversion is better for a set of standard-definition
        videos.
        """
        for item in range(self.spreadsheet.n_items):
            handbrake_dictionary = self.spreadsheet.make_handbrake_dictionary(item)
            subler_dictionary = self.spreadsheet.make_subler_dictionary(item)
            video = self._set_source_parameters(item)
            video.convert()
            MP4(handbrake_dictionary["Destination"]).tag(subler_dictionary)

    def parallel_convert_and_tag(self, n_cores=None):
        """
        Convert the source to destination using the Handbrake parameters then
        tag with Subler parameters for each item. This converts multiple items
        at once using as many cores as your computer decides to allocate. I
        find serial conversion is better for a set of high-definition videos,
        while parallel conversion is better for a set of standard-definition
        videos.

        Parameters
        ----------
        n_cores
            If desired, the user-specified number of cores to use.
        """
        pool = get_multiprocessing_pool(n_cores)
        for item in range(self.spreadsheet.n_items):
            pool.apply_async(self._parallel_convert_and_tag, args=(item,))
        cleanup_parallel_processing(pool)

    def test_convert_and_tag(self):
        """
        Test conversion and tagging of a spreadsheet. This is done using
        parallel processing.
        """
        pool = get_multiprocessing_pool()
        for item in range(self.spreadsheet.n_items):
            pool.apply_async(self._test_parallel_convert_and_tag, args=(item,))
        cleanup_parallel_processing(pool)


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
     - **Audio Track Names:** What to name the audio tracks, like "Mono Audio", "Stereo Audio", "5.1 Surround Sound",
       "7.1 Surround Sound", or "Commentary with XXX and XXX". **NOTE:** If one of the names has a comma in it (like a list
       of people's names), it will need to be encapsulated by double quotes. Unfortunately, I haven't yet figured out
       how to include double quotes in a name, so you will have to fix those manually after conversion.
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
       those names. You will want these in escaped format, so for instance:

       | ``1,The\ First\ Chapter\,\ with\ comma``
       | ``2,The\ Second\ Chapter\,\ with\ comma``

       Otherwise, as long as there are more a single chapter, it will automatically include them.

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
     - **Audio Track Names:** What to name the audio tracks, like "Mono Audio", "Stereo Audio", "5.1 Surround Sound",
       "7.1 Surround Sound", or "Commentary with XXX and XXX". **NOTE:** If one of the names has a comma in it (like a list
       of people's names), it will need to be encapsulated by double quotes. Unfortunately, I haven't yet figured out
       how to include double quotes in a name, so you will have to fix those manually after conversion.
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
       those names. You will want these in escaped format, so for instance:

       | ``1,The\ First\ Chapter\,\ with\ comma``
       | ``2,The\ Second\ Chapter\,\ with\ comma``

       Otherwise, as long as there are more a single chapter, it will automatically include them.
    """

    kind, columns = _columns_for_kind(kind)
    save_path = os.path.join(save_directory, f"empty_metadata_{kind}.xlsx")
    _make_dataframe_from_columns(columns).to_excel(save_path, index=False)

    print(f"Empty spreadsheet saved to \"{save_path}\".")
