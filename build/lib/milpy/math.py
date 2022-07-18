"""Math functions that don't seem to exist elsewhere."""

import numpy as np
from typing import Union


def haversine(lat0: float, lon0: float, lat1: Union[float, np.ndarray], lon1: Union[float, np.ndarray]) -> Union[float, np.ndarray]:

    """
    Calculates the angle(s) between a starting point (lat0, lon0) and
    another single coordinate point or array of coordinate points.

    Parameters
    ----------
    lat0
        Starting latitude.
    lon0
        Starting longitude.
    lat1
        Ending latitude, either a single point or ndarray of points.
    lon1
        Ending longitude, either a single point or ndarray of points.

    Notes
    -----
    All inputs need to be in radians.

    Examples
    --------
    To calculate the angle between two discrete points and (10°, 5°) and (-87°, 146°):

    >>> angle = haversine(np.radians(10), np.radians(-5), np.radians(-87), np.radians(146))
    >>> angle
    102.62029119229642

    To calculate the angles between a meshgrid of latitudes and longitudes and (10°, 5°):

    >>> longitude_array, latitude_array = np.meshgrid(np.linspace(-180, 180, 3), np.linspace(-90, 90, 3))
    >>> angles = haversine(np.radians(10), np.radians(5), latitude_array, longitude_array)
    >>> angles
    array([[ 85.58252865, 126.50037948,  82.03049899],
           [121.22238399,  11.16895281, 130.98714208],
           [ 67.20066783, 106.51977943,  63.31226578]])
    >>> angles.shape
    (3, 3)
    """

    ha = np.sin((lat0 - lat1) / 2) ** 2
    hb = np.cos(lat1) * np.cos(lat0) * np.sin((lon0 - lon1) / 2) ** 2
    return np.degrees(2 * np.arcsin(np.sqrt(ha + hb)))
