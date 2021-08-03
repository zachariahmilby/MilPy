import os
import subprocess
import sys
import glob
import numpy as np
import pandas as pd
import re
from typing import Iterable
import multiprocessing as mp


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


def _path_to_system_executable(executable: str) -> str:
    path = _EscapedString(os.path.join(os.path.dirname(__file__), executable))
    os.system(f'chmod 777 {path}')
    return path


def _construct_terminal_commands(command_line_arguments: Iterable[str]) -> str:
    return ' '.join(command_line_arguments)


class _SourceOptions:

    """
    This class stores video source options.

    See the HandBrakeCLI documentation for more information:
    https://handbrake.fr/docs/en/latest/cli/command-line-reference.html
    """

    def __init__(self, input_file: str, title="1"):

        """
        Parameters
        ----------
        input_file
            The location of the video source, either a DVD or Blu-ray disc image in `.iso` format or a video file.
        title
            The video title to convert.
        """

        self._input = _EscapedString(input_file)
        _ValidatePath(self._input.original)
        self.title = title

    def __str__(self):
        return f"Source options:\n"\
               f"   Input: {self._input.original}\n"\
               f"   Title: {self.title}"

    def __repr__(self):
        return self.construct_terminal_commands()

    def construct_terminal_commands(self) -> str:

        """
        Returns the options as a set of HandBrakeCLI flags and options.
        """

        options = [f"--input={self._input}",
                   f"--title={self.title}"]
        return _construct_terminal_commands(options)

    @property
    def input(self):
        return self._input

    @input.setter
    def input(self, value):
        self._input = _EscapedString(value)
        _ValidatePath(self._input.original)


class _DestinationOptions:

    """
    This class stores video destination options.

    See the HandBrakeCLI documentation for more information:
    https://handbrake.fr/docs/en/latest/cli/command-line-reference.html
    """

    def __init__(self, output: str, video_format="av_mp4", markers=True, optimize=True, align_av=True):

        """
        Parameters
        ----------
        output
            The absolute path with filename to where you want the converted video.
        video_format
            Video container format.
        markers
            Whether or not to add chapter markers as defined in the source.
        optimize
            Whether or not to optimize MP4 files for HTTP streaming.
        align_av
            Whether or not to add audio silence or black video frames to start of streams so that all streams start at
            exactly the same time.
        """

        self._output = _EscapedString(output)
        _ValidateDirectory(self._output.original)
        self.format = video_format
        self.markers = markers
        self.optimize = optimize
        self.align_av = align_av

    def __str__(self):
        return f"Destination options:\n"\
               f"   Output: {self._output.original}\n"\
               f"   Video format: {self.format}\n"\
               f"   Markers: {self.markers}\n"\
               f"   Optimize: {self.markers}\n"\
               f"   Align A/V: {self.align_av}"

    def __repr__(self):
        return self.construct_terminal_commands()

    def construct_terminal_commands(self):

        """
        Returns the options as a set of HandBrakeCLI flags and options.
        """

        options = [f"--output={self._output}",
                   f"--format={self.format}",
                   ]
        if self.markers:
            options.append(f"--markers")
        else:
            options.append(f"--no-markers")
        if self.optimize:
            options.append(f"--optimize")
        if self.align_av:
            options.append(f"--align-av")
        return _construct_terminal_commands(options)

    @property
    def output(self):
        return self._output

    @output.setter
    def output(self, value):
        self._output = _EscapedString(value)
        _ValidateDirectory(self._output.original)


class _VideoOptions:

    """
    This class stores video encoder options.

    See the HandBrakeCLI documentation for more information:
    https://handbrake.fr/docs/en/latest/cli/command-line-reference.html
    """

    def __init__(self, encoder="x265", speed="fast", quality="20", two_pass=False):

        """
        Parameters
        ----------
        encoder
            The video encoder.
        speed
            Adjustment to video encoding settings for a particularmspeed/efficiency tradeoff. "fast" is the default and
            is usually a good choice.
        quality
            Video quality factor. "20" is best for SD video, "22" for HD video.
        two_pass
            Whether or not to do an initial pass through the video to further optimize the conversion.
        """

        self.encoder = encoder
        self.speed = speed
        self.quality = quality
        self.two_pass = two_pass

    def __str__(self):
        return f"Video options:\n"\
               f"   Encoder: {self.encoder}\n"\
               f"   Speed (encoder preset): {self.speed}\n"\
               f"   Quality: {self.quality}\n"\
               f"   Two-pass: {self.two_pass}\n"\
               f"   Framerate: variable"

    def __repr__(self):
        return self.construct_terminal_commands()

    def construct_terminal_commands(self):

        options = [f"--encoder={self.encoder}",
                   f"--encoder-preset={self.speed}",
                   f"--quality={self.quality}",
                   f"--vfr"]
        if self.two_pass:
            options.append("--two-pass")
            options.append("--turbo")

        return _construct_terminal_commands(options)


class _AudioOptions:

    """
    This class stores audio encoder options.

    See the HandBrakeCLI documentation for more information:
    https://handbrake.fr/docs/en/latest/cli/command-line-reference.html
    """

    def __init__(self, audio_titles="1", encoder="ca_aac", bitrates=None, mixdowns=None, sample_rates=None, names='None'):

        """
        Parameters
        ----------
        audio_titles
            The audio tracks to include (for instance, "1,2" might be the main audio (track 1) and a director's
            commentary (track 2)).
        encoder
            The audio encoder.
        bitrates
            The bitrates for the audio tracks in kbps, separated by a comma. Use 64-bits per channel, so 128 kbps for
            stereo, 384 for 5.1 surround sound and 512 for 7.1 surround sound. For a 5.1 surround sound main track and a
            stereo director's commentary, this would look like "384,128".
        mixdowns
            Format(s) for audio downmixing/upmixing. For the example in bitrates, this would be "5point1,stereo".
        sample_rates
            Sample rate in kHz. A good choice is probably 48 if you want to manually define it, otherwise it will
            automatically determine an appropriate rate.
        names
            The audio track names, separated by a comma. For the example above, this might be 5.1 Surround Sound,"Commentary by Director
            XXX XXX, Producer YYY YYY and Screenwriter ZZZ ZZZ".
        """

        self.audio_titles = audio_titles
        self.encoder = encoder
        self.bitrates = bitrates
        self.mixdowns = mixdowns
        self.sample_rates = sample_rates
        self._names = _EscapedString(names)

    def __str__(self):
        print_string = f"Audio options:\n"\
                       f"   Audio titles: {self.audio_titles}\n"\
                       f"   Encoder: {self.encoder}\n"\
                       f"   Bitrate(s) (kbps): {self.bitrates}\n"\
                       f"   Mixdown(s): {self.mixdowns}\n"\
                       f"   Sample rate(s) (kHz): {self.sample_rates}\n"\
                       f"   Track name(s): {self.names.original}"
        return print_string.replace('None', 'Same as source')

    def __repr__(self):
        return self.construct_terminal_commands()

    @property
    def names(self):
        return self._names

    @names.setter
    def names(self, value):
        self._names = _EscapedString(value)

    def construct_terminal_commands(self):

        options = [f"--audio={self.audio_titles}",
                   f"--aencoder={self.encoder}"]
        if self.bitrates is not None:
            options.append(f"--ab={self.bitrates}")
        if self.mixdowns is not None:
            options.append(f"--mixdown={self.mixdowns}")
        if self.sample_rates is not None:
            options.append(f"--arate={self.sample_rates}")
        if self.names != 'None':
            options.append(f'--aname={self.names}')

        return _construct_terminal_commands(options)


class _PictureOptions:

    """
    This class stores video picture options.

    See the HandBrakeCLI documentation for more information:
    https://handbrake.fr/docs/en/latest/cli/command-line-reference.html
    """

    def __init__(self, width=None, height=None, crop="0:0:0:0"):

        """
        Parameters
        ----------
        width
            The width of the converted video in pixels.
        height
            The height of the converted video in pixels.
        crop
            How much to crop off of either side (if there are any black bars or something). Format is
            "top:bottom:left:right".
        """

        self.width = width
        self.height = height
        self.crop = crop

    def __str__(self):
        print_string = f"Picture options:\n"\
                       f"   Width: {self.width}\n"\
                       f"   Height: {self.height}\n"\
                       f"   Crop: {self.crop}\n"\
                       f"   Anamorphic: off\n"\
                       f"   Comb-detection: on\n"\
                       f"   Decomb method: bob"
        return print_string.replace('None', 'Same as source')

    def __repr__(self):
        return self.construct_terminal_commands()

    def construct_terminal_commands(self):
        options = ["--non-anamorphic",
                   "--comb-detect",
                   "--decomb=\"bob\"",
                   f"--crop={self.crop}"]
        if self.width is not None:
            options.append(f"--width={self.width}")
            options.append(f"--display-width={self.width}")
        if self.height is not None:
            options.append(f"--height={self.height}")

        return _construct_terminal_commands(options)


class _SubtitleOptions:

    """
    This class stores subtitle options. For now, you can only choose one subtitle track and it is burned-in.

    See the HandBrakeCLI documentation for more information:
    https://handbrake.fr/docs/en/latest/cli/command-line-reference.html
    """

    def __init__(self, subtitle_tracks="None"):

        """
        Parameters
        ----------
        subtitle_tracks
            The subtitle track to burn-in.
        """

        self.subtitles = subtitle_tracks

    def __str__(self):
        if self.subtitles != "None":
            print_string = f"Subtitle options:\n"\
                           f"   Track: {self.subtitles}\n"\
                           f"   Burned-in: on"
        else:
            print_string = ''
        return print_string

    def __repr__(self):
        return self.construct_terminal_commands()

    def construct_terminal_commands(self) -> str:

        options = []
        if self.subtitles != "None":
            options.append(f"--subtitle={self.subtitles}")
            options.append("--subtitle-burned")

        return _construct_terminal_commands(options)


class VideoConverter:

    """
    This class sets up HandBrakeCLI command-line flags and provides a method for initiating video conversion using the
    stored parameters. These parameters can be changed after creating an instance of the class. The default behavior is
    to convert the video with the same parameters as the source.
    """

    def __init__(self, input_filepath: str, output_filepath: str):

        """
        Parameters
        ----------
        input_filepath
            The location of the video source, either a DVD or Blu-ray disc image in `.iso` format or a video file like
            an MP4 or MKV.
        output_filepath
            The absolute path with filename of the name/extension and location of the converted video.

        Examples
        --------
        To setup a video conversion, create an instance of the VideoConverter class:

        >>> converter = VideoConverter("/path/to/disc.iso", "/path/to/output.mp4")
        >>> print(converter)
        Video converter settings:
        ------------------------
        Source options:
           Input: /path/to/disc.iso
           Title: 1
        Destination options:
           Output: /path/to/output.mp4
           Video format: av_mp4
           Markers: True
           Optimize: True
           Align A/V: True
        Video options:
           Encoder: x265
           Speed (encoder preset): fast
           Quality: 20
           Two-pass: False
           Framerate: variable
        Audio options:
           Audio titles: 1
           Encoder: ca_aac
           Bitrate(s) (kbps): Same as source
           Mixdown(s): Same as source
           Sample rate(s) (kHz): Same as source
           Track name(s): Same as source
        Picture options:
           Width: Same as source
           Height: Same as source
           Crop: 0:0:0:0
           Anamorphic: off
           Comb-detection: on
           Decomb method: bob

        To change the input title:

        >>> converter = VideoConverter("/path/to/disc.iso", "/path/to/output.mp4")
        >>> converter.source_options.title = "2"
        >>> print(converter.source_options.__str__())
        Source options:
           Input: /path/to/disc.iso
           Title: 2

        You shouldn't need to change any of the destination options. You have to set the output filepath when creating an
        instance of `VideoConverter`, and the settings `video_format`, `markers`, `optimize` and `align_av` are defaults
        appropriate for almost all situations.

        To change the video options:

        >>> converter = VideoConverter("/path/to/disc.iso", "/path/to/output.mp4")
        >>> converter.video_options.encoder = "x264"
        >>> converter.video_options.speed = "ultrafast"
        >>> converter.video_options.quality = "47"
        >>> converter.video_options.two_pass = True
        >>> print(converter.video_options.__str__())
        Video options:
           Encoder: x264
           Speed (encoder preset): ultrafast
           Quality: 47
           Two-pass: True
           Framerate: variable

        To change the audio options (note there is a method for the audio track names):

        >>> converter = VideoConverter("/path/to/disc.iso", "/path/to/output.mp4")
        >>> converter.audio_options.audio_titles = "1,2"
        >>> converter.audio_options.encoder = "ca_aac"
        >>> converter.audio_options.bitrates = "384,128"
        >>> converter.audio_options.mixdowns = "5point1,stereo"
        >>> converter.audio_options.sample_rates = "48,48"
        >>> converter.audio_options.names = "5.1 Surround Sound,Stereo"
        >>> print(converter.audio_options.__str__())
        Audio options:
           Audio titles: 1,2
           Encoder: ca_aac
           Bitrate(s) (kbps): 384,128
           Mixdown(s): 5point1,stereo
           Sample rate(s) (kHz): 48,48
           Track name(s): 5.1 Surround Sound,Stereo

        To change the picture options:

        >>> converter = VideoConverter("/path/to/disc.iso", "/path/to/output.mp4")
        >>> converter.picture_options.width = '1920'
        >>> converter.picture_options.height = '1080'
        >>> converter.picture_options.crop = '0:0:20:20'  # cropping 20 pixels from the top and bottom
        >>> print(converter.picture_options.__str__())
        Picture options:
           Width: 1920
           Height: 1080
           Crop: 0:0:20:20
           Anamorphic: off
           Comb-detection: on
           Decomb method: bob

        To add a subtitle track (for now, it has to be burned-in):

        >>> converter = VideoConverter("/path/to/disc.iso", "/path/to/output.mp4")
        >>> converter.subtitle_options.subtitles = "1"
        >>> print(converter.subtitle_options.__str__())
        Subtitle options:
           Track: 1
           Burned-in: on

        Finally, once you've set everything the way you want it, you can begin the video conversion using the method
        `convert()`.

        >>> converter = VideoConverter("/path/to/disc.iso", "/path/to/output.mp4")
        >>> converter.convert()
        """

        self.handbrake_cli = _path_to_system_executable("anc/HandBrakeCLI")
        self.source_options = _SourceOptions(input_filepath)
        self.destination_options = _DestinationOptions(output_filepath)
        self.video_options = _VideoOptions()
        self.audio_options = _AudioOptions()
        self.picture_options = _PictureOptions()
        self.subtitle_options = _SubtitleOptions()

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

    def convert(self, test: bool = False) -> None:

        """
        Initiates the video conversion with HandBrakeCLI using set input parameters.

        Parameters
        ----------
        test
            If True, does a quick, short conversion of the file to test output and tagging.
        """

        if test:
            self.video_options.encoder = 'x264'
            self.video_options.speed = 'ultrafast'

        options = [self.handbrake_cli,
                   self.source_options.construct_terminal_commands(),
                   self.destination_options.construct_terminal_commands(),
                   self.video_options.construct_terminal_commands(),
                   self.audio_options.construct_terminal_commands(),
                   self.picture_options.construct_terminal_commands(),
                   self.subtitle_options.construct_terminal_commands()]

        if test:
            options.append('--start-at=seconds:0')
            options.append('--stop-at=seconds:10')
        os.system(_construct_terminal_commands(options))


def video_inspector(source_filepath: str):

    """
    This function prints the existing metadata from an MP4 file in the console using the SublerCLI executable.

    Parameters
    ----------
    source_filepath
        The absolute path to the source file.

    Examples
    --------
    >>> video_inspector('/path/to/video file.mp4')
    """

    options = [_path_to_system_executable("anc/SublerCLI"),
               f"-source {_EscapedString(source_filepath)}",
               f"-listmetadata"]
    os.system(_construct_terminal_commands(options))


def _create_metadata_dictionaries(keys: Iterable[str], values: Iterable[str]) -> dict:

    """
    This function takes a set of metadata keys and values and converts them into a dictionaries for HandBrake and
    Subler settings.

    Parameters
    ----------
    keys
        An iterable of dictionary keys.
    values
        An iterable of values for the dictionary keys.

    Examples
    --------
    >>> metadata = _MetadataLoader("/path/to/metadata.xlsx")
    >>> line = 17
    >>> keys_list = metadata.get_keys()
    >>> values_list = metadata.get_values(line)
    >>> handbrake_dict, subler_dict = _create_metadata_dictionaries(keys_list, values_list)
    """

    excluded_keys = ["Filename", "Source", "Title", "Audio", "Dimensions", "Crop", "Audio Bitrate", "Audio Mixdown", "Audio Notes",
                     "Subtitle", "Chapters", "Quality Factor"]
    subler_dictionary = {key: value for key, value in zip(keys, values) if key not in excluded_keys}
    subler_dictionary['Copyright'] = f'\u00A9 {subler_dictionary["Copyright"]}'
    handbrake_dictionary = {key: value for key, value in zip(keys, values) if key in excluded_keys}
    return handbrake_dictionary, subler_dictionary


class _MetadataLoader:

    """
    This class loads a Microsoft Excel metadata spreadsheet and returns the keys and values for a particular entry.
    """

    def __init__(self, path_to_excel_spreadsheet: str):

        self.data = pd.read_excel(path_to_excel_spreadsheet, dtype=str)

    def get_keys(self) -> list:

        """
        Get a list of the column names from the spreadsheet.
        """

        return self.data.keys().tolist()

    def get_values(self, line: int) -> list:

        """
        Get the line number you want, (index starting from 0).
        """

        return self.data.iloc[line].tolist()


class _VideoTagger:

    """
    This class takes an MP4 file and tags it with metadata.
    """

    def __init__(self, source_filepath: str, destination_filepath: str, metadata: dict):

        """
        Parameters
        ----------
        source_filepath
            The absolute filepath to the source file (likely a temporary file).
        destination_filepath
            This is the filename and absolute filepath for the tagged video. Cannot be the same as the source filepath.
        metadata
            A dictionary of metadata tags and values.

        Examples
        --------
        Load metadata from a metadata.xlsx spreadsheet, then apply the video tags for line 17.

        >>> metadata_dataframe = _MetadataLoader("/path/to/metadata.xlsx")
        >>> line = 17
        >>> keys_list = metadata_dataframe.get_keys()
        >>> values_list = metadata_dataframe.get_values(line)
        >>> handbrake_dict, subler_dict = _create_metadata_dictionaries(keys_list, values_list)
        >>> _VideoTagger("/path/to/temp17.mp4", "/path/to/18 Output.mp4", metadata=subler_dict).tag(remove_input=True)
        """

        self.subler_cli = _path_to_system_executable("anc/SublerCLI")
        self.source = _EscapedString(source_filepath)
        self.destination = _EscapedString(destination_filepath)
        self._check_if_input_and_output_match()
        _ValidateDirectory(self.destination.original)
        self.metadata = metadata

    @staticmethod
    def _format_metadata(metadata: dict) -> str:
        return ''.join([r'{"%s":"%s"}' % (i, j) for i, j in metadata.items()])

    def _check_if_input_and_output_match(self):
        if self.source.original == self.destination.original:
            raise Exception('Input and output files cannot be the same! Either change the output directory or give the '
                            'input file a temporary filename.')

    def tag(self, remove_input=False):
        options = [self.subler_cli,
                   f"-source {self.source}",
                   f"-dest {self.destination}",
                   f"-metadata {self._format_metadata(self.metadata)}",
                   f"-language English"]
        os.system(_construct_terminal_commands(options))

        if remove_input:
            os.remove(self.source.original)


def empty_metadata_spreadsheet(save_directory: str, kind: str):

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

    >>> empty_metadata_spreadsheet("/path/to/directory", kind='TV')
    Empty spreadsheet saved to "/path/to/directory/empty_metadata_tv.xlsx."

    For a movie spreadsheet:

    >>> empty_metadata_spreadsheet("/path/to/directory", kind='movie')
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
     - **Filename:** The absolute path to where you want the final video to be saved. I usually prefix the file with it's
       number, e.g., /.../07 Episode 7.mp4 for a main episode or /.../Special Features/103 Special Feature.mp4
     - **Source:** The absolute path to the source, either a disc image `.iso` or a video file.
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

         - "0" = SD
         - "1" = 720p HD
         - "2" = 1080p HD
         - "3" = 4K UHD

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
     - **Filename:** The absolute path to where you want the final video to be saved. I don't usually prefix the main
       movie file, e.g., /.../Film Name.mp4, but for a special feature /.../Special Features/103 Special Feature.mp4
     - **Source:** The absolute path to the source, either a disc image `.iso` or a video file.
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

          - "0" = SD
          - "1" = 720p HD
          - "2" = 1080p HD
          - "3" = 4K UHD

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

    tv_columns = ['Name',
                  'Artist',
                  'Album Artist',
                  'Album',
                  'Genre',
                  'Release Date',
                  'Track #',
                  'TV Show',
                  'TV Episode ID',
                  'TV Season',
                  'TV Episode #',
                  'TV Network',
                  'Description',
                  'Series Description',
                  'Copyright',
                  'Media Kind',
                  'Cover Art',
                  'Rating',
                  'Cast',
                  'Filename',
                  'Source',
                  'Title',
                  'Audio',
                  'Dimensions',
                  'Crop',
                  'Audio Bitrate',
                  'Audio Mixdown',
                  'Audio Notes',
                  'HD Video',
                  'Quality Factor',
                  'Subtitle',
                  'Chapters']

    movie_columns = ['Name',
                     'Genre',
                     'Release Date',
                     'Description',
                     'Copyright',
                     'Media Kind',
                     'Cover Art',
                     'Rating',
                     'Rating Annotation',
                     'Cast',
                     'Director',
                     'Producers',
                     'Screenwriters',
                     'Filename',
                     'Source',
                     'Title',
                     'Audio',
                     'Dimensions',
                     'Crop',
                     'Audio Bitrate',
                     'Audio Mixdown',
                     'Audio Notes',
                     'HD Video',
                     'Quality Factor',
                     'Subtitle',
                     'Chapters']

    kind = kind.lower()
    if (kind == 'tv') or (kind == 'television'):
        kind = 'tv'
        columns = tv_columns
    elif (kind == 'movie') or (kind == 'film'):
        kind = 'movie'
        columns = movie_columns
    else:
        raise Exception('Unrecognized kind.')

    note = ['Make sure you change all cells to "Text" to avoid any Excel automated formatting shit.']
    note.extend(['']*(len(columns)-1))
    df = pd.DataFrame([note], columns=columns)
    save_path = os.path.join(save_directory, f"empty_metadata_{kind}.xlsx")
    df.to_excel(save_path, index=False)
    print(f"Empty spreadsheet saved to \"{save_path}\".")


def _convert_and_tag_spreadsheet_item(iterator: int, dataframe: pd.DataFrame, output_directory: str, test: bool):

    """
    This function actually processes an item from a spreadsheet. It's designed to be called by multiprocessing, but can
    be called in serial, too.

    Parameters
    ----------
    iterator
        The looping iterator.
    dataframe
        The loaded metadata spreadsheet.
    output_directory
        The directory where the files are to be saved.
    test
        Whether or not to do a quick test run to check file tagging and output paths.
    """

    # get the line from the metadata
    row = dataframe.data.iloc[iterator].dropna().to_dict()
    keys = row.keys()
    values = row.values()

    # location for temporary file
    temporary_file = _EscapedString(os.path.join(output_directory, f"temp{iterator}.mp4"))

    # extract metadata to dictionaries
    handbrake_metadata, subler_metadata = _create_metadata_dictionaries(keys, values)

    # setup a VideoConverter instance and set parameters
    converter = VideoConverter(handbrake_metadata["Source"], temporary_file.original)
    converter.source_options.title = handbrake_metadata["Title"]
    converter.video_options.quality = handbrake_metadata["Quality Factor"]
    converter.audio_options.audio_titles = handbrake_metadata["Audio"]
    converter.audio_options.bitrates = handbrake_metadata["Audio Bitrate"]
    converter.audio_options.mixdowns = handbrake_metadata["Audio Mixdown"]
    converter.audio_options.names = handbrake_metadata["Audio Notes"]
    converter.picture_options.width = handbrake_metadata["Dimensions"].split("x")[0]
    converter.picture_options.height = handbrake_metadata["Dimensions"].split("x")[1]
    converter.picture_options.crop = handbrake_metadata["Crop"]

    # convert the video
    converter.convert(test=test)

    # tag temporary video
    output_file = os.path.join(output_directory, handbrake_metadata["Filename"])
    _VideoTagger(temporary_file.original, output_file, metadata=subler_metadata).tag(remove_input=True)


def _get_appropriate_number_of_cores():
    n_cores = int(mp.cpu_count() / 2) - 2
    if n_cores < 1:
        n_cores = 1
    return n_cores


def _set_processor_pool(n_cores):
    """
    mp.get_context('fork') allows parallel processing without 'if __name__ == '__main__'.
    """
    return mp.get_context('fork').Pool(n_cores)


def _asynchronous_process_spreadsheet(pool, dataframe, output_directory, test):
    for i in range(len(dataframe.data)):
        args = (i, dataframe, output_directory, test)
        pool.apply_async(_convert_and_tag_spreadsheet_item, args=args)


def _cleanup_parallel_processing(pool):
    """
    No one knows what these do...but things don't work without them.
    """
    pool.close()
    pool.join()


def process_spreadsheet(path_to_spreadsheet: str, output_directory: str, parallel: bool = True, test: bool = False):

    """
    Convert and tag a metadata spreadsheet using either parallel or serial (non-parallel) processing.

    **Parallel processing:**
    This will process multiple videos simultaneously, one per core. For some reason my 6-core computer comes back
    with 12 cores available, so I've set this to use half the apparently available cores minus 2. For my 2019 MacBook Pro
    with a 6-core i7 processor, this assigns (12/2 - 2) = 4 cores to parallel processing. The two additional cores allows
    for you to continue to work on other things that aren't processor-heavy.

    **Serial processing:**
    This will use multiple cores automatically assigned by your system but will process a single line in the spreadsheet
    at a time. This will be faster for a single video file since Handbrake scales well up to 4-6 cores, but for an
    larger set of videos you're probably better off processing in parallel.

    Parameters
    ----------
    path_to_spreadsheet
        Absolute path to the spreadsheet.
    output_directory
        Output directory where you want the final tagged videos.
    parallel
        Whether or not to process in parallel.
    test
        Whether or not to do a test run, converting from 00:00:03 to 00:00:06 at low-resolution and very fast speed in
        order to test that all the outputs are tagged correctly and go where you expect them to. Useful to run before
        initiating a long full-spreadsheet process.

    Examples
    --------
    >>> process_spreadsheet('/path/to/metadata.xlsx', '/path/to/output_directory',
    ...                     parallel=True, test=False)
    """

    dataframe = _MetadataLoader(path_to_spreadsheet)
    if parallel:
        n_cores = _get_appropriate_number_of_cores()
        pool = _set_processor_pool(n_cores)
        _asynchronous_process_spreadsheet(pool, dataframe, output_directory, test)
        _cleanup_parallel_processing(pool)
    else:
        for i in range(len(dataframe.data)):
            _convert_and_tag_spreadsheet_item(i, dataframe, output_directory, test)
