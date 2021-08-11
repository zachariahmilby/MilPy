import os
from pathlib import Path
from milpy.miscellaneous import EscapedString
from milpy.video._handbrake import path_to_handbrake_cli, \
    VideoConversionOptions
from milpy.video._subler import path_to_subler_cli, \
    format_subler_metadata_from_dictionary, make_temporary_video, _Spreadsheet
from milpy.terminal_interface import construct_terminal_commands
from milpy.parallel_processing import get_multiprocessing_pool, cleanup_parallel_processing


class _VideoConverter:
    def __init__(self, converter_options: VideoConversionOptions):
        self.options = converter_options

    def _create_command_of_input_options(self):
        return f'{path_to_subler_cli()} ' + \
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
    def _add_10_second_test_to_options(options: str):
        return f'{options} ' + '--start-at=seconds:0 ' + '--stop-at=seconds:10'

    def convert(self):
        command = self._create_command_of_input_options()
        os.system(command)

    def test(self):
        self._set_test_video_options()
        options = self._create_command_of_input_options()
        options = self._add_10_second_test_to_options(options)
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
        options = [path_to_subler_cli(),
                   f"-source {temporary_filepath}",
                   f"-dest {EscapedString(self.file_path)}",
                   f"-metadata {format_subler_metadata_from_dictionary(metadata_dictionary)}",
                   f"-language English"]
        os.system(construct_terminal_commands(options))
        os.remove(temporary_filepath.original)


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
        destination = self.file_path.replace(file_extension, ', Title 1 (Converted).mp4')
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
        self.spreadsheet = _Spreadsheet(path_to_spreadsheet)

    @staticmethod
    def _get_source_type(source_filepath: str):
        ext = os.path.splitext(source_filepath)
        if ext == ".mp4":
            return MP4(source_filepath)
        elif ext == ".mkv":
            return MKV(source_filepath)
        elif ext == ".iso":
            return ISO(source_filepath)
        else:
            raise Exception('Bad input file type.')

    def _set_source_parameters(self, item: int):
        handbrake_metadata = self.spreadsheet.make_handbrake_dictionary(item)
        video = _get_source_type(handbrake_metadata["Source"])
        video.converter_options.source_options.title = handbrake_metadata["Title"]
        video.converter_options.source_options.quality = handbrake_metadata["Quality Factor"]
        video.converter_options.destination_options.destination = handbrake_metadata["Destination"]
        video.converter_options.destination_options.markers = handbrake_metadata["Chapters"]
        video.converter_options.audio_options.audio_titles = handbrake_metadata["Audio"]
        video.converter_options.audio_options.bitrates = handbrake_metadata["Audio Bitrate"]
        video.converter_options.audio_options.mixdowns = handbrake_metadata["Audio Mixdown"]
        video.converter_options.audio_options.track_names = handbrake_metadata["Audio Track Names"]
        video.converter_options.picture_options.width = handbrake_metadata["Dimensions"].split("x")[0]
        video.converter_options.picture_options.height = handbrake_metadata["Dimensions"].split("x")[1]
        video.converter_options.picture_options.crop = handbrake_metadata["Crop"]
        return video

    @staticmethod
    def _parallel_convert(video):
        video.convert()

    @staticmethod
    def _test_parallel_convert(video):
        video.test_convert()

    @staticmethod
    def _parallel_tag(video: MP4, metadata_dictionary):
        video.tag(metadata_dictionary)

    @staticmethod
    def _parallel_convert_and_tag(video, metadata_dictionary):
        video.convert()
        output_video = MP4(video.converter_options.destination_options.destination)
        output_video.tag(metadata_dictionary)

    @staticmethod
    def _test_parallel_convert_and_tag(video, metadata_dictionary):
        video.test_convert()
        output_video = MP4(video.converter_options.destination_options.destination)
        output_video.tag(metadata_dictionary)

    def serial_convert_only(self):
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

    def parallel_convert_only(self):
        """
        Convert the source to destination using the Handbrake parameters for
        each item. This converts multiple items at once using as many cores as
        your computer decides to allocate. I find parallel conversion is
        better for a set of standard-definition videos, while serial conversion
        is better for a set of high-definition videos.
        """
        pool = get_multiprocessing_pool()
        for item in range(self.spreadsheet.n_items):
            video = self._set_source_parameters(item)
            pool.apply_async(self._parallel_convert, args=(video,))
        cleanup_parallel_processing(pool)

    def test_convert_only(self):
        """
        Test conversion of a spreadsheet. This is done using parallel
        processing.
        """
        pool = get_multiprocessing_pool()
        for item in range(self.spreadsheet.n_items):
            video = self._set_source_parameters(item)
            pool.apply_async(self._test_parallel_convert, args=(video,))
        cleanup_parallel_processing(pool)

    def tag_only(self):
        """
        Assuming all source items are MP4 files which only need tagging, tag
        them in parallel using the Subler parameters for each item. This uses
        the "Destination" parameter as the source and destination file.
        """
        pool = get_multiprocessing_pool()
        for item in range(self.spreadsheet.n_items):
            handbrake_dictionary = self.spreadsheet.make_handbrake_dictionary(item)
            subler_dictionary = format_subler_metadata_from_dictionary(self.spreadsheet.make_subler_dictionary(item))
            video = MP4(handbrake_dictionary["Destination"])
            pool.apply_async(self._parallel_tag, args=(video, subler_dictionary))
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
            subler_dictionary = format_subler_metadata_from_dictionary(self.spreadsheet.make_subler_dictionary(item))
            video = self._set_source_parameters(item)
            video.convert()
            video.tag(subler_dictionary)

    def parallel_convert_and_tag(self):
        """
        Convert the source to destination using the Handbrake parameters then
        tag with Subler parameters for each item. This converts multiple items
        at once using as many cores as your computer decides to allocate. I
        find serial conversion is better for a set of high-definition videos,
        while parallel conversion is better for a set of standard-definition
        videos.
        """
        pool = get_multiprocessing_pool()
        for item in range(self.spreadsheet.n_items):
            subler_dictionary = format_subler_metadata_from_dictionary(self.spreadsheet.make_subler_dictionary(item))
            video = self._set_source_parameters(item)
            pool.apply_async(self._parallel_convert_and_tag, args=(video, subler_dictionary))
        cleanup_parallel_processing(pool)

    def test_convert_and_tag(self):
        """
        Test conversion and tagging of a spreadsheet. This is done using
        parallel processing.
        """
        pool = get_multiprocessing_pool()
        for item in range(self.spreadsheet.n_items):
            subler_dictionary = format_subler_metadata_from_dictionary(self.spreadsheet.make_subler_dictionary(item))
            video = self._set_source_parameters(item)
            pool.apply_async(self._test_parallel_convert_and_tag, args=(video, subler_dictionary))
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
       "7.1 Surround Sound", or "Commentary with XXX and XXX". **NOTE:** These will need to be encapsulated by quotation marks,
       e.g., "Stereo Audio","Director's Commentary". Take care if the label has apostrophes or double quotes.
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
     - **Audio Track Names:** What to name the audio tracks, like "Mono Audio", "Stereo Audio", "5.1 Surround Sound",
       "7.1 Surround Sound", or "Commentary with XXX and XXX". **NOTE:** These will need to be encapsulated by quotation marks,
       e.g., "Stereo Audio","Director's Commentary". Take care if the label has apostrophes or double quotes.
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
