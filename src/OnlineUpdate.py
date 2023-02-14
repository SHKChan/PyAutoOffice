import json
import os
import requests
import subprocess
import time
import zipfile

from MyLogger import LOGGER

GITHUB_USERNAME="SHKChan"
REPO_NAME = "PyAutoOffice"

def check_for_updates(CURRENT_VERSION: str) -> bool:
    try:
        # Get the latest release information from Github API
        url = "https://api.github.com/repos/{}/{}/releases/latest".format(GITHUB_USERNAME, REPO_NAME)
        response = requests.get(url)
        latest_release = json.loads(response.text)
        latest_version = latest_release["tag_name"]

        filename = ""
        # Compare the latest version with the current version
        if latest_version != CURRENT_VERSION:
            # Download the updated version from Github
            file_url = latest_release['assets'][0]['browser_download_url']
            r = requests.get(file_url)

            filename = REPO_NAME + "_" + latest_version + ".zip"
            # Save the updated version to a file
            with open(filename, "wb") as f:
                f.write(r.content)

            # Replace the existing version with the updated version
            # (Note: You will need to unzip the file and overwrite the existing files)
            print("A new version is available. Updating to version", latest_version)
        else:
            print("No updates available.")
            
        return filename
    except Exception as e:
        # Log the error into a file
        LOGGER.wt()
        raise Exception("Error while checking for updates: " + str(e))


def install_update(filename: str) -> None:
    try:
        # Wait for the software to close
        time.sleep(10)
        # Run the 'unzip' command using the subprocess module
        if(os.path.exists(filename)):
            with zipfile.ZipFile(filename, 'r') as zip_ref:
                zip_ref.extractall(".")
    except subprocess.CalledProcessError as e:
    # Log the error into a file
        LOGGER.wt()
        raise subprocess.CalledProcessError("Error while installing for updates: " + str(e))
        