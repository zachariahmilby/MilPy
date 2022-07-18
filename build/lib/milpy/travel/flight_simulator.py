"""Makes an animation of a commercial airline flight on a rotating globe with day and night."""

# TODO: try converting pytz uses to datetime.timezone?

import os
import re
import sys
import time
from datetime import datetime, timedelta

import cartopy.crs as ccrs
import ephem
import matplotlib.animation as animation
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pytz
from pyproj import Geod
from shapely.geometry.polygon import LinearRing
from fuzzywuzzy import process

from milpy.math import haversine


def _get_current_directory():
    return os.path.dirname(__file__)


def _get_line_from_database(path_to_database: str, query_item: str, query_type: str):

    """
    This function takes a relative path to one of the CSV databases, finds the queried item, and returns the row from
    the CSV for that item.

    Parameters
    ----------
    path_to_database
        Relative path to the database, e.g., "anc/database.csv."
    query_item
        What to query from the database, e.g., for an airport, something like "Los Angeles International Airport."
    query_type
        What kind of item you are querying. Options are "airline", "airport," or "aircraft."

    Returns
    -------
    The line from the CSV as a list of string items.

    Raises
    ------
    Raises exception if the item isn't found in the CSV. Will print five of the closest options given your input to
    check against, but if none of this is what you want, it provides a link to the database for you to search through
    manually.
    """

    absolute_path_to_database = os.path.join(_get_current_directory(), path_to_database)
    database = pd.read_csv(absolute_path_to_database)
    values = database["Name"].values
    ind = np.where(values == query_item)[0]
    if len(ind) == 0:
        try:
            possibilities = np.array(process.extract(query_item, values))[:, 0]
        except IndexError:
            possibilities = ["No matches found."]
        error_string = f"Could not find {query_type} \"{query_item}\". Is one of these the {query_type} you want?\n"
        for possibility in possibilities:
            error_string += f'   {possibility}\n'
        error_string += f"If not, check your spelling or the database at\n   \"{absolute_path_to_database}\".\n"
        raise ValueError(error_string)
    else:
        return database.values[ind[0]]


# Note: if used by another travel/something.py, these should go into their own file somewhere
class _Airline:

    """This class gets and stores airline data from the ancillary airlines.csv file."""

    def __init__(self, name: str):

        """
        Parameters
        ----------
        name
            The name of the airline, e.g., "Delta Air Lines" or "United Airlines."
        """

        info = _get_line_from_database('anc/airlines.csv', name, 'airline')
        self.name = info[1]
        self.alias = info[2]
        self.IATA = info[3]
        self.ICAO = info[4]
        self.callsign = info[5]

    def __str__(self):
        print_string = f'Airline Information:\n'
        print_string += f'   Name: {self.name}\n'
        print_string += f'   Alias: {self.alias}\n'
        print_string += f'   IATA Code: {self.IATA}\n'
        print_string += f'   ICAO Code: {self.ICAO}\n'
        print_string += f'   Callsign: {self.callsign}'
        return print_string


class _Airport:

    """This class gets and stores airport data from the ancillary airports.csv file."""

    def __init__(self, name):

        """
        Parameters
        ----------
        name
            The name of the airport, e.g., "Los Angeles International Airport."
        """

        info = _get_line_from_database('anc/airports.csv', name, 'airport')
        self.name = info[1]
        self.city = info[2]
        self.country = info[3]
        self.IATA = info[4]
        self.ICAO = info[5]
        self.latitude = info[6]
        self.longitude = info[7]
        self.altitude = info[8]
        self.timezone = info[11]

    def __str__(self):
        print_string = f'Airport Information:\n'
        print_string += f'   Name: {self.name}\n'
        print_string += f'   City: {self.city}\n'
        print_string += f'   Country: {self.country}\n'
        print_string += f'   IATA Code: {self.IATA}\n'
        print_string += f'   ICAO Code: {self.ICAO}\n'
        print_string += f'   Latitude: {self.latitude}°\n'
        print_string += f'   Longitude: {self.longitude}°\n'
        print_string += f'   Altitude: {int(self.altitude*3.28084):,} feet\n'
        print_string += f'   Timezone: {self.timezone}'
        return print_string


class _Aircraft:

    """This class gets and stores aircraft data from the ancillary aircraft.csv file."""

    def __init__(self, name):

        """
        Parameters
        ----------
        name
            The name of the aircraft type, e.g., "Airbus A350-900" or "Embraer 175."
        """

        info = _get_line_from_database('anc/aircraft.csv', name, 'aircraft')
        self.manufacturer = name.split(' ')[0]
        self.type = ' '.join(name.split(' ')[1:])
        self.ISO = info[1]
        self.DAFIF = info[2]

    def __str__(self):
        print_string = f'Aircraft Information:\n'
        print_string += f'   Manufacturer: {self.manufacturer}\n'
        print_string += f'   Type: {self.type}\n'
        print_string += f'   ISO Code: {self.ISO}\n'
        print_string += f'   DAFIF Code (Obsolete): {self.DAFIF}'
        return print_string


def _calculate_subsolar_position(date_time: datetime) -> (float, float):

    """
    This function calculates the coordinates of the sub-solar position for a given date/time.

    Parameters
    ----------
    date_time
        The datetime object for which you want to calculate the sub-solar position.

    Returns
    -------
    A tuple of (sub-solar latitude, sub-solar longitude) in degrees.
    """

    greenwich = ephem.Observer()
    greenwich.lat = "0"
    greenwich.lon = "0"
    greenwich.date = date_time
    sun = ephem.Sun(greenwich)
    sun.compute(greenwich.date)
    sun_lon = np.degrees(sun.ra - greenwich.sidereal_time())
    if sun_lon < -180:
        sun_lon += 360
    elif sun_lon > 180:
        sun_lon -= 360
    sun_lat = np.degrees(sun.dec)
    return sun_lat, sun_lon


class _DayNightMapCreator:

    """This class creates a cylindrical day/night map of Earth with a 12-degree-wide twilight zone."""

    def __init__(self, sub_solar_latitude, sub_solar_longitude):

        """
        Parameters
        ----------
        sub_solar_latitude
            The latitude position of the Sun in degrees.
        sub_solar_longitude
            The longitude position of the Sun in degrees.
        """

        self.sub_solar_latitude = sub_solar_latitude
        self.sub_solar_longitude = sub_solar_longitude
        self.day_map = self.convert_to_float_image(self.load_image('anc/earth_day.jpg'))
        self.night_map = self.convert_to_float_image(self.load_image('anc/earth_night.jpg'))

    @staticmethod
    def load_image(path_to_image):
        return plt.imread(os.path.join(_get_current_directory(), path_to_image))

    @staticmethod
    def convert_to_float_image(image_array):
        return image_array / 255

    def sza(self) -> np.ndarray:

        """
        This method creates an 1800x3600 array of angles from the sub-solar position. Any positions with angles below 90°
        are day time, from 90° to 102° are twilight, and above 102° are night time.
        """

        longitudes = np.linspace(np.radians(-180), np.radians(180), 3600)
        latitudes = np.linspace(np.radians(-90), np.radians(90), 1800)
        longitudes, latitudes = np.meshgrid(longitudes, latitudes)
        return haversine(np.radians(self.sub_solar_latitude), np.radians(self.sub_solar_longitude), latitudes, longitudes)

    @staticmethod
    def terminator_mask(sza_arr: np.ndarray) -> np.ndarray:

        """
        This method takes an array of solar zenith angles made by the method `self._sza()` and
        turns it into an alpha mask.

        Parameters
        ----------
        sza_arr
            An cylindrical-map-shaped array of solar zenith angles.

        Returns
        -------
        An array of multiplication factors to replicate day/night transition.
        """

        ind90 = np.where(sza_arr <= 90)
        ind108 = np.where(sza_arr > 108)
        sza_arr -= 90
        sza_arr /= 18
        sza_arr[ind90] = 0
        sza_arr[ind108] = 1
        sza_arr *= np.pi / 2
        sza_arr = np.cos(sza_arr) ** 2
        return np.flipud(np.repeat(sza_arr[:, :, None], 3, axis=2))

    def map(self) -> np.ndarray:

        """
        This method returns a cylindrical map image of Earth day and night as specified by the subsolar position with a
        12-degree-wide twilight zone.

        Returns
        -------
        The cylindrical map image with dimensions (1800, 3600, 3).
        """

        twilight_mask = self.terminator_mask(self.sza())
        return (self.day_map * twilight_mask) + (self.night_map * (1 - twilight_mask))


class _DayNightMap(np.ndarray):

    """This class creates a cylindrical day/night map of Earth with a 12-degree-wide twilight zone which acts like a
    numpy ndarray, so you can display it as-is or access sub-solar latitude and longitude as properties."""

    def __new__(cls, sub_solar_latitude, sub_solar_longitude):
        cyl_map = _DayNightMapCreator(sub_solar_latitude, sub_solar_longitude).map()
        obj = np.asarray(cyl_map).view(cls)
        obj.sub_solar_latitude = sub_solar_latitude
        obj.sub_solar_longitude = sub_solar_longitude
        return obj

    def __array_finalize__(self, obj: np.ndarray):
        if obj is None:
            return
        self.sub_solar_latitude = getattr(obj, 'sub_solar_latitude', None)
        self.sub_solar_longitude = getattr(obj, 'sub_solar_longitude', None)


class _FlightParameters:

    """This class stores specific flight parameters and calculates some information needed for the animation."""

    def __init__(self, departure_airport: _Airport, departure_time: datetime, arrival_airport: _Airport,
                 arrival_time: datetime):

        """
        Parameters
        ----------
        departure_airport
            The airport from which you are departing.
        departure_time
            The local departure date and time with minute precision.
        arrival_airport
            The airport at which you are arriving.
        arrival_time
            The local arrival date and time with minute precision.

        Examples
        --------
        >>> flight_example = Flight(departure_airport=_Airport("Los Angeles International Airport"),
        ...                         departure_time=datetime(year=2021, month=7, day=2, hour=18, minute=50),
        ...                         arrival_airport=_Airport("Denver International Airport"),
        ...                         arrival_time=datetime(year=2021, month=7, day=2, hour=22, minute=6))
        """

        self.departure_airport = departure_airport
        self.departure_time = departure_time
        self.arrival_airport = arrival_airport
        self.arrival_time = arrival_time
        self.duration = arrival_time - departure_time
        self.n_frames = int(self.duration.total_seconds() / 60) + 1
        self.flight_path = self.calculate_flight_path()
        self.camera_path = self.calculate_camera_path()

    @staticmethod
    def dateline_fix(coordinates):
        longitude = coordinates[:, 0]
        latitude = coordinates[:, 1]
        if np.size(np.where((longitude < -175) | (longitude > 175))) != 0:
            longitude[np.where(longitude < 0)] += 360
        return np.array([longitude, latitude]).T

    def calculate_flight_path(self) -> np.ndarray:

        """
        This method calculates the great-circle flight path between the two airports in steps of 1 minute based on the
        departure and arrival times.

        Returns
        -------
        An array of flight path coordinates with shape (n, 2). The first entry in axis 1 is longitude, the second is
        latitude.
        """

        geod = Geod("+ellps=WGS84")
        flight_path = np.array(geod.npts(lon1=self.departure_airport.longitude, lat1=self.departure_airport.latitude,
                                         lon2=self.arrival_airport.longitude, lat2=self.arrival_airport.latitude,
                                         npts=self.n_frames))
        return self.dateline_fix(flight_path)

    def calculate_camera_path(self) -> np.ndarray:

        """
        This method calculates the camera position (the central latitude and longitude of the map projection) for each
        frame in the animation. It is based on the previously-calculated flight path.

        Returns
        -------
        An array of camera coordinate positions with shape (n, 2). The first entry in axis 1 is longitude, the second
        is latitude.
        """

        camera_path = self.flight_path.copy()
        lons = camera_path[:, 0]
        lats = camera_path[:, 1]
        lats = np.linspace(lats[0], lats[-1], self.n_frames)  # redo latitudes to move straight from start to end
        return self.dateline_fix(np.array([lons, lats]).T)


class Flight:

    """This class initates a flight simulation."""

    def __init__(self, airline_name: str, flight_number: int, departure_airport: str, arrival_airport: str,
                 departure_time: str, arrival_time: str, aircraft: str):

        """
        Parameters
        ----------
        airline_name
            The name of the airline.
        flight_number
            The flight number.
        departure_airport
            The airport from which you are departing.
        arrival_airport
            The airport at which you are arriving.
        departure_time
            The local departure date/time.
        arrival_time
            The local arrival date/time.
        aircraft
            The name of the aircraft type.

        Raises
        ------
        ValueError
            Raised if the item isn't found in the CSV. Will print five of the closest options given your input to
            check against, but if none of this is what you want, it provides a link to the database for you to search
            through manually.

        Examples
        --------
        >>> example_flight = Flight(airline_name="United Airlines",
        ...                         flight_number=2283,
        ...                         departure_airport="Los Angeles International Airport",
        ...                         arrival_airport="Denver International Airport",
        ...                         departure_time="July 2, 2021, 6:50 pm",
        ...                         arrival_time="July 2, 2021, 10:06 pm",
        ...                         aircraft="Boeing 737 MAX 9")
        >>> print(example_flight)
        This flight simulation has the following parameters:
           Airline: United Airlines
           Origin: Los Angeles (Los Angeles International Airport/LAX)
           Destination: Denver (Denver International Airport/DEN)
           Departure Time: Fri, Jul 2, 2021, 6:50 PM
           Arrival Time: Fri, Jul 2, 2021, 10:06 PM
           Aircraft: Boeing 737 MAX 9
        """

        self.airline = _Airline(airline_name)
        self.flight_number = flight_number
        self.departure_airport = _Airport(departure_airport)
        self.arrival_airport = _Airport(arrival_airport)
        self.departure_time = self._convert_time_to_utc(departure_time, self.departure_airport.timezone)
        self.arrival_time = self._convert_time_to_utc(arrival_time, self.arrival_airport.timezone)
        self.aircraft = _Aircraft(aircraft)
        self.flight_parameters = _FlightParameters(self.departure_airport, self.departure_time,
                                                   self.arrival_airport, self.arrival_time)

    def __str__(self):
        print_string = f'This flight simulation has the following parameters:\n' \
                       f'   Airline: {self.airline.name}\n' \
                       f'   Origin: {self.departure_airport.city} ({self.departure_airport.name}/{self.departure_airport.IATA})\n' \
                       f'   Destination: {self.arrival_airport.city} ({self.arrival_airport.name}/{self.arrival_airport.IATA})\n' \
                       f'   Departure Time: {self._printable_datetime(self.departure_time, self.departure_airport.timezone)}\n' \
                       f'   Arrival Time: {self._printable_datetime(self.arrival_time, self.arrival_airport.timezone)}\n' \
                       f'   Aircraft: {self.aircraft.manufacturer} {self.aircraft.type}'
        return print_string

    @staticmethod
    def _printable_datetime(datetime_obj, timezone):
        return datetime.strftime(datetime_obj.astimezone(pytz.timezone(timezone)), '%a, %b %d, %Y, %I:%M %p').replace(' 0', ' ')

    @staticmethod
    def _highres_orthographic(projection: ccrs.Projection) -> None:

        """
        This method makes the default Cartopy orthographic projection higher resolution, so the polygon shape of the
        edge of the globe isn't as obvious.

        Parameters
        ----------
        projection
            The Cartopy orthographic projection.
        """

        r = 6378137  # radius of Earth in meters
        a = float(projection.globe.semimajor_axis or r)
        b = float(projection.globe.semiminor_axis or a)
        t = np.linspace(0, 2 * np.pi, 3601)
        coords = np.vstack([a * 0.99999 * np.cos(t), b * 0.99999 * np.sin(t)])[
                 :, ::-1]
        projection._boundary = LinearRing(coords.T)

    @staticmethod
    def _convert_time_to_utc(date_time: str, timezone: str) -> datetime:

        """
        This method takes a date and time string, makes it timezone-aware, then converts it to UTC.

        Parameters
        ----------
        date_time
            The date/time as a string with the format "Month [D]D, YYYY, [H]H:MM am/pm."
        timezone
            The timezone for the date/time string. For a list of all possible timezones, query the pytz package with
            `pytz.all_timezones`.

        Returns
        -------
        A UTC datetime object for the given date/time.
        """

        dt = datetime.strptime(date_time, '%B %d, %Y, %I:%M %p')
        tz = pytz.timezone(timezone)
        dt = tz.localize(dt)
        return dt.astimezone(pytz.utc)

    def _func_animate(self, iterator: int, t0: float, highres: bool):

        """
        This is the iterable animation function required by `matplotlib.animation.FuncAnimation`.

        Parameters
        ----------
        iterator
            The first argument of this method must be the iterator passed by `matplotlib.animation.FuncAnimation`.
        t0
            The time when the system began making the animation (for printing out progress report).
        highres
            Whether or not to display the surface map in native Cartopy resolution (which for some reason is 700x1400),
            or the higher resolution of the video itself (currently 1280x2560).
        """

        # remove any existing axes
        plt.gca().remove()

        # print progress report
        seconds = time.time() - t0
        m, s = divmod(seconds, 60)
        h, m = divmod(m, 60)
        print(f'Generating frame {iterator + 1}/{self.flight_parameters.n_frames} '
              f'({(iterator + 1) / self.flight_parameters.n_frames * 100.:.2f}%), '
              f'{int(h)}:{int(m):0>2}:{int(s):0>2} elapsed...' + ' ' * 20, end='\r')

        # create high-resolution orthographic projection and place in an axis; also define the coordinate transform
        projection = ccrs.Orthographic(central_longitude=self.flight_parameters.camera_path[iterator, 0],
                                       central_latitude=self.flight_parameters.camera_path[iterator, 1])
        self._highres_orthographic(projection)
        transform = ccrs.PlateCarree()
        ax = plt.axes([0.1, 0.1, 0.8, 0.8], projection=projection, facecolor='k')

        # define imshow parameters
        imshow_params = {'extent': [-180, 180, -90, 90], 'origin': 'upper', 'transform': transform}
        if highres:
            imshow_params['regrid_shape'] = (1280, 2560)

        # display the Earth surface image for current time in the flight
        sub_solar_latitude, sub_solar_longitude = _calculate_subsolar_position(self.departure_time + timedelta(minutes=iterator))
        surface_image = _DayNightMap(sub_solar_latitude, sub_solar_longitude)
        ax.imshow(surface_image, **imshow_params)

        # plot the red flight path and point indicating the position of the plane
        ax.plot(self.flight_parameters.flight_path[:iterator, 0], self.flight_parameters.flight_path[:iterator, 1],
                linewidth=2, color='tab:red', transform=transform)
        ax.scatter([self.flight_parameters.flight_path[iterator, 0]], [self.flight_parameters.flight_path[iterator, 1]],
                   s=15, color='tab:red', edgecolor='none', transform=transform)

        # place text
        fig = plt.gcf()
        text_params = dict(transform=fig.transFigure, color='white', ha='left', va='top', fontsize=8)
        meta_text = [fr'$\bf Origin:$ {self.departure_airport.name} ({self.departure_airport.IATA})',
                     fr'$\bf Destination:$ {self.arrival_airport.name} ({self.arrival_airport.IATA})',
                     fr"$\bf Departure\ Time:$ {self._printable_datetime(self.departure_time, self.departure_airport.timezone)}",
                     fr"$\bf Arrival\ Time:$ {self._printable_datetime(self.arrival_time, self.arrival_airport.timezone)}",
                     fr'$\bf Aircraft:$ {self.aircraft.manufacturer} {self.aircraft.type}',
                     fr'$\bf Flight\ Time\ Elapsed:$ {int(divmod(iterator, 60)[0])}h {int(divmod(iterator, 60)[1]):0>2}m',
                     ]
        title_string = f'{self.airline.name} Flight {self.flight_number} from {self.departure_airport.city} to {self.arrival_airport.city}'
        plt.text(0.025, 0.975, title_string, fontweight='bold', fontsize='10', transform=fig.transFigure, color='white', ha='left', va='top')
        [plt.text(0.025, 0.9725 - 0.02 * (j + 1), meta_text[j], **text_params) for j in range(len(meta_text))]

    def animate(self, save_directory: str, highres: bool = False):

        """
        This method generates and saves an animation of the simulated flight.

        Parameters
        ----------
        save_directory
            The absolute path to the directory where you want to save the animation.
        highres
            Whether or not you want the animation to be at better resolution.

        Examples
        --------
        >>> example_flight = Flight(airline_name="United Airlines",
        ...                         flight_number=2283,
        ...                         departure_airport="Los Angeles International Airport",
        ...                         departure_time="July 2, 2021, 6:50 pm",
        ...                         arrival_airport="Denver International Airport",
        ...                         arrival_time="July 2, 2021, 10:06 pm",
        ...                         aircraft="Boeing 737 MAX 9")
        >>> example_flight.animate('/absolute/path/to/save/directory', highres=False)
        """

        # print flight information
        print(self.__str__())

        # record starting time
        t0 = time.time()

        # make a figure and set the coordinate transform
        fig = plt.figure(figsize=(8, 8), facecolor='k', dpi=200)

        # make the animation
        ani = animation.FuncAnimation(fig, self._func_animate, frames=self.flight_parameters.n_frames,
                                      fargs=(t0, highres), interval=41.67, repeat=False)

        # make the filename and path for saving
        ext = ''
        if highres:
            ext = '_highres'
        save_name = f"{self.airline.ICAO}{self.flight_number}_{datetime.strftime(self.departure_time, '%Y-%m-%d')}{ext}.mp4"

        # save the animation
        ani.save(os.path.join(save_directory, save_name), savefig_kwargs={'facecolor': fig.get_facecolor()})
        print('\n')
