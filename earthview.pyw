import ctypes
import logging
import os
import requests
import sys
import win32gui
import win32con

from bs4 import BeautifulSoup
from enum import Enum
from traceback import format_exc
from win32com.shell import shell, shellcon

"""
CONSTANTS
"""

MAIN_PAGE_URL = "https://earthview.withgoogle.com"

# Gets the path to the GoogleEarthView folder inside the user's Pictures folder.
# Code taken from <https://stackoverflow.com/a/3858957>.
PICTURES_FOLDER = os.path.join(shell.SHGetFolderPath(0, shellcon.CSIDL_MYPICTURES, None, 0), "GoogleEarthView")

class ExitCode(Enum):
    SUCCESS = 0
    # Errors:
    UNKNOWN = 1
    REQUEST = 2
    FILE_IO = 3

"""
FUNCTIONS
"""

def get_res(url):
    # Downloads URL and checks status.
    # In case of failure, logs error and quits program.
    try:
        res = requests.get(url, timeout=1.0)
    except Exception as e:
        logging.error(f"When connecting to {url}: {format_exc(1)}")
        sys.exit(ExitCode.REQUEST)

    status = res.status_code
    if status != requests.codes.ok:
        logging.error(f"Code {status} when requesting {url}")
        sys.exit(ExitCode.REQUEST)
        
    return res


def save_res(path, res):
    # Writes a request to disk.
    # In case of failure, logs error and quits program.
    try:
        path_head = os.path.dirname(path)
        os.mkdir(path_head)
    except FileExistsError:
        pass
    except:
        logging.error(f"When creating folder {path_head}: {format_exc(1)}")
        quit(ExitCode.FILE_IO)

    try:
        with open(path, "wb") as f:
            for chunk in res.iter_content():
                f.write(chunk)
    except IOError as e:
        logging.error(f"When saving {path}: {format_exc(1)}")
        quit(ExitCode.FILE_IO)


def parse_html(html):
    return BeautifulSoup(html, "html.parser")

"""
MAIN
"""

def main():
    # Downloads the main page of Google Earth View and finds info about the image.
    main_page = parse_html(get_res(MAIN_PAGE_URL).content)

    image_url = main_page.find("img", {"class":"photo-view photo-view--1 photo-view--active"})["src"]

    image_name = main_page.find("a", {"class":"button intro__explore"})["href"][1:]

    image_extension = os.path.splitext(image_url)[-1]

    image_path = os.path.join(PICTURES_FOLDER, image_name + image_extension)

    # If the image does not yet exist on the disk, downloads and saves it.
    if not os.path.exists(image_path):
        image = get_res(image_url)
        save_res(image_path, image)
        
    # Sets the wallpaper.
    changed = win32con.SPIF_UPDATEINIFILE | win32con.SPIF_SENDCHANGE | win32con.SPIF_SENDWININICHANGE
    win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, image_path, changed)

    # Logs success.
    logging.info(f"Set {image_name} as wallpaper")


if __name__ == "__main__":
    logging.basicConfig(format="%(levelname)s - %(asctime)s: %(message)s",
                        filename="earthview.log", level=logging.INFO, datefmt="%d/%m/%Y %H:%M")
    try:
        main()
    except SystemExit as e:
        sys.exit(e.code)
    except Exception as e:
        logging.critical(f"Unknown exception: {format_exc(1)}")
        sys.exit(ExitCode.UNKNOWN)
    else:
        sys.exit(ExitCode.SUCCESS)
