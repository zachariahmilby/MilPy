import pandas as pd


def _tv_spreadsheet_columns():
    return ['Name', 'Artist', 'Album Artist', 'Album', 'Genre', 'Release Date', 'Track #', 'TV Show', 'TV Episode ID',
            'TV Season', 'TV Episode #', 'TV Network', 'Description', 'Series Description', 'Copyright', 'Media Kind',
            'Cover Art', 'Rating', 'Cast', 'Source', 'Destination', 'Title', 'Audio', 'Dimensions', 'Crop',
            'Audio Bitrate', 'Audio Mixdown', 'Audio Track Names', 'HD Video', 'Quality Factor', 'Subtitles', 'Chapters']


def _movie_spreadsheet_columns():
    return ['Name', 'Genre', 'Release Date', 'Description', 'Copyright', 'Media Kind', 'Cover Art', 'Rating',
            'Rating Annotation', 'Cast', 'Director', 'Producers', 'Screenwriters', 'Source', 'Destination', 'Title',
            'Audio', 'Dimensions', 'Crop', 'Audio Bitrate', 'Audio Mixdown', 'Audio Track Names', 'HD Video',
            'Quality Factor', 'Subtitles', 'Chapters']


def _columns_for_kind(kind):

    kind = kind.lower()

    if (kind == 'tv') or (kind == 'television'):
        kind = 'tv'
        columns = _tv_spreadsheet_columns()
    elif (kind == 'movie') or (kind == 'film'):
        kind = 'movie'
        columns = _movie_spreadsheet_columns()
    else:
        raise Exception("Unrecognized kind. Try 'tv' or 'movie'.")

    return kind, columns


def _make_dataframe_from_columns(columns):
    note = ['Make sure you change all cells to "Text" to avoid any Excel automated formatting shit.']
    note.extend([''] * (len(columns) - 1))
    return pd.DataFrame([note], columns=columns)
