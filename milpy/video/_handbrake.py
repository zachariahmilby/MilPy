from milpy.miscellaneous import EscapedString, ValidatePath, ValidateDirectory
from milpy.terminal_interface import construct_terminal_commands, path_to_system_executable


def path_to_handbrake_cli():
    return path_to_system_executable("video/anc/HandBrakeCLI")


class SourceOptions:

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

        self._input = EscapedString(input_file)
        ValidatePath(self._input.original)
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
                   f"--title={self.title}",
                   f"--previews=1:0"]
        return construct_terminal_commands(options)

    @property
    def input(self):
        return self._input

    @input.setter
    def input(self, value):
        self._input = EscapedString(value)
        ValidatePath(self._input.original)


class DestinationOptions:

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

        self._output = EscapedString(output)
        ValidateDirectory(self._output.original)
        self.format = video_format
        self.markers = markers
        self.optimize = optimize
        self.align_av = align_av

    def __str__(self):
        return f"Destination options:\n"\
               f"   Output: {self._output.original}\n"\
               f"   Video format: {self.format}\n"\
               f"   Markers: {self.markers}\n"\
               f"   Optimize: {self.optimize}\n"\
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
            options.append(f"--markers={self.markers}")
        else:
            options.append(f"--no-markers")
        if self.optimize:
            options.append(f"--optimize")
        if self.align_av:
            options.append(f"--align-av")
        return construct_terminal_commands(options)

    @property
    def output(self):
        return self._output

    @output.setter
    def output(self, value):
        self._output = EscapedString(value)
        ValidateDirectory(self._output.original)


class VideoOptions:

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

        return construct_terminal_commands(options)


class AudioOptions:

    """
    This class stores audio encoder options.

    See the HandBrakeCLI documentation for more information:
    https://handbrake.fr/docs/en/latest/cli/command-line-reference.html
    """

    def __init__(self, audio_titles="1", encoder="ca_aac", bitrates=None, mixdowns=None, sample_rates=None, track_names='None'):

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
        track_names
            The audio track names, separated by a comma. **NOTE:** These will need to be encapsulated by quotation marks,
            e.g., "Stereo Audio","Director's Commentary". Take care if the label has apostrophes or double quotes.
        """

        self.audio_titles = audio_titles
        self.encoder = encoder
        self.bitrates = bitrates
        self.mixdowns = mixdowns
        self.sample_rates = sample_rates
        self.track_names = track_names

    def __str__(self):
        print_string = f"Audio options:\n"\
                       f"   Audio titles: {self.audio_titles}\n"\
                       f"   Encoder: {self.encoder}\n"\
                       f"   Bitrate(s) (kbps): {self.bitrates}\n"\
                       f"   Mixdown(s): {self.mixdowns}\n"\
                       f"   Sample rate(s) (kHz): {self.sample_rates}\n"\
                       f"   Track name(s): {self.track_names}"
        return print_string.replace('None', 'Same as source')

    def __repr__(self):
        return self.construct_terminal_commands()

    def construct_terminal_commands(self):

        options = [f"--audio={self.audio_titles}",
                   f"--aencoder={self.encoder}"]
        if self.bitrates is not None:
            options.append(f"--ab={self.bitrates}")
        if self.mixdowns is not None:
            options.append(f"--mixdown={self.mixdowns}")
        if self.sample_rates is not None:
            options.append(f"--arate={self.sample_rates}")
        if self.track_names != 'None':
            options.append(f'--aname={self.track_names}')

        return construct_terminal_commands(options)


class PictureOptions:

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

        return construct_terminal_commands(options)


class SubtitleOptions:

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

        return construct_terminal_commands(options)


class VideoConversionOptions:
    """Instances of this class hold all options and ensure they're valid.

    Raises
    ------
    TypeError
        Raised if the inputs are not of the proper options.

    """
    def __init__(self, source, destination):
        self.source = SourceOptions(source)
        self.destination = DestinationOptions(destination)
        self.video = VideoOptions()
        self.audio = AudioOptions()
        self.picture = PictureOptions()
        self.subtitle = SubtitleOptions()

        self._raise_type_error_if_inputs_are_not_proper_type()
        self._raise_exception_if_input_and_output_filenames_match()

    def _raise_type_error_if_inputs_are_not_proper_type(self):
        self._raise_type_error_if_not_correct_option_type(
            self.source, 'source', SourceOptions)
        self._raise_type_error_if_not_correct_option_type(
            self.destination, 'destination', DestinationOptions)
        self._raise_type_error_if_not_correct_option_type(
            self.video, 'video', VideoOptions)
        self._raise_type_error_if_not_correct_option_type(
            self.audio, 'audio', AudioOptions)
        self._raise_type_error_if_not_correct_option_type(
            self.picture, 'picture', PictureOptions)
        self._raise_type_error_if_not_correct_option_type(
            self.subtitle, 'subtitle', SubtitleOptions)

    @staticmethod
    def _raise_type_error_if_not_correct_option_type(
            input_options, option_name: str, option_type):
        if not isinstance(input_options, option_type):
            message = f'{option_name} must be an instance of {option_type}, ' \
                      f'not {type(input_options)}.'
            raise TypeError(message)

    def __str__(self):
        header = 'Video converter settings:'
        delimiter = '-' * (len(header) - 2)
        return f'{header}\n' \
               f'{delimiter}\n' \
               f'{self.source}\n' \
               f'{self.destination}\n' \
               f'{self.video}\n' \
               f'{self.audio}\n' \
               f'{self.picture}\n' \
               f'{self.subtitle}'

    def _raise_exception_if_input_and_output_filenames_match(self):
        if self.source.input.original == self.destination.output.original:
            message = 'Input and output files cannot be the same! Either ' \
                      'change the output directory or give the input file a ' \
                      'temporary filename.'
            raise Exception(message)
