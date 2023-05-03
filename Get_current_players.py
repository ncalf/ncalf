import requests
import re
import os
import urllib.request
import argparse
import shutil
import sqlite3
from sqlite3 import Error
import pyodbc
import time
import Update_database
import NCALF_bin
import pymysql
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC

from bs4 import BeautifulSoup
import csv

chromedriverfile = r"C:\ncalf\bin\NCALF v2020\ncalf_python_local\chromedriver.exe"


def output_to_text_file(source):
    filename = r'D:\Heners\Python training\NCALF\html.txt'
    with open(filename, 'w', encoding="utf-8") as f:
        f.write(source)


def convertToBinaryData(filename):
    # Convert digital data to binary format
    with open(filename, 'rb') as file:
        blobData = file.read()
    return blobData


def download_image_to_file(player_image, player_name):
    base_dir = r'D:\Heners\Python training\NCALF\Player Images 2022\\'

    # Open the url image, set stream to True, this will return the stream content.
    resp = requests.get(player_image, stream=True)
    # Open a local file with wb ( write binary ) permission.
    image_filename = base_dir + player_name + ".png"
    local_file = open(image_filename, 'wb')
    # Set decode_content value to True, otherwise the downloaded image file's size will be zero.
    resp.raw.decode_content = True
    # Copy the response stream raw data to local image file.
    shutil.copyfileobj(resp.raw, local_file)
    # Remove the image url response object.
    del resp
    return image_filename


def get_all_current_playersrrent_players():
    # downloads player names and details and imports them into the transition table
    driver = webdriver.Chrome(executable_path=chromedriverfile)

    URL = "https://afl.com.au"
    r = requests.get(URL)
    if r.status_code != 200:
        print("Can't download ", URL)
        exit
    team_URLs = []
    senior_player_URLs = []
    soup = BeautifulSoup(r.text, 'lxml')

    # Set the directory name where the images will be saved
    base_dir = r'D:\Heners\Python training\NCALF\Player Images 2023\\'

    # Connect to the database
    connection = NCALF_bin.create_sql_connection()

    # Get a cursor
    cursor = connection.cursor()

    # Get the URLs for each of the teams
    team_sites = soup.find_all('a', attrs={'class': 'th1-club-nav__club-link'})

    # Counter for number of clubs to process
    i = 0  # Add 1 to get the index number of the club on the AFL website

    for team_site in team_sites:

        if i > 18:
            continue

        team_URL = team_site.get('href')
        team_name = team_site.find('span', attrs={'class': 'th1-club-nav__club-name'}).get_text().strip()
        print(team_URL)

        # Get the site for the AFL team
        team_r = requests.get(team_URL)
        team_soup = BeautifulSoup(team_r.text, 'lxml')
        print("Downloading players for", team_name)

        # Get the banner headings and find the AFL one to drill down into
        banner_headings = team_soup.find_all('a', class_='navigation__link navigation__link--in-drop-down')
        for banner_heading in banner_headings:
            if banner_heading.text.strip() == 'AFL' or \
                    banner_heading.text.strip() == "Men's" or \
                    banner_heading.text.strip() == 'Player Profiles' or \
                    banner_heading.text.strip() == 'Player profiles':
                playerlist_URL = banner_heading.get('href')
                while playerlist_URL[0] == "/":
                    playerlist_URL = playerlist_URL[1:]
                playerlist_URL = team_URL + playerlist_URL
                if team_name == "North Melbourne":
                    playerlist_URL = playerlist_URL + '/players'  # Adjustment for North Melbourne website
                if team_name in ["Richmond", "St Kilda"]:
                    playerlist_URL = playerlist_URL + '/squad'  # Adjustment for Richmond website

                # Go to each player list page and get the links to the player pages
                playerlist_r = requests.get(playerlist_URL)
                player_soup = BeautifulSoup(playerlist_r.text, 'lxml')

                squad_list = player_soup.find_all('li', class_='squad-list__item')

                # Work through each player in the list
                for player in squad_list:
                    player_profile_URL = player.a.get('href')
                    if player_profile_URL is None:
                        continue

                    # Get the player name
                    player_firstname = player.find('h1',
                                                   attrs={'class': 'player-item__name'}).get_text().split()[0].strip()
                    player_surname = player.find('span', attrs={'class': 'player-item__last-name'}).get_text()
                    # Check for middle initials and ignore
                    if '.' in player_firstname:
                        player_firstname = player_firstname.split()[0]
                    player_name = player_firstname + " " + player_surname

                    player_profile_URL = team_URL + player_profile_URL[1:]

                    # Go to each individual player page and download the image and player name
                    print("Getting:", player_profile_URL)
                    driver.get(player_profile_URL)
                    time.sleep(0.8)
                    html = driver.page_source
                    player_soup = BeautifulSoup(html, 'lxml')
                    # output_to_text_file(player_soup.prettify())

                    if player_soup.find('picture', class_="player__headshot-image picture") is not None:

                        player_file_name = player_name

                        # Try to get the player image, if it isn't there, just pass through as None
                        try:
                            player_image_link = "http:" + \
                                                player_soup.find('picture', class_="player__headshot-image picture"). \
                                                    find('img',
                                                         class_='picture__img js-faded-image fade-in-on-load is-loaded'). \
                                                    get('src')
                        except:
                            player_image = None
                        else:
                            player_image_filename = download_image_to_file(player_image_link, player_file_name)
                            player_image = convertToBinaryData(player_image_filename)
                    else:
                        player_image = None

                    Sql_update_download_table = \
                        """INSERT INTO playerdownload 
                        (PlayerFirstName, PlayerSurname, Club, Photo) VALUES (%s, %s, %s, %s)"""
                    val = (player_firstname, player_surname, team_name, player_image)
                    cursor.execute(Sql_update_download_table, val)
                    connection.commit()
                    r.close()
        i += 1


def check_current_players():
    # checks that current players are in the database

    URL = "https://afl.com.au"
    r = requests.get(URL)
    if r.status_code != 200:
        print("Can't download ", URL)
        exit
    team_URLs = []
    senior_player_URLs = []
    soup = BeautifulSoup(r.text, 'lxml')

    # Set the directory name where the images will be saved
    base_dir = "D:\\Heners\\Python training\\NCALF\\Player Images\\"
    database = r"D:\Heners\ncalf\ncalfdb.db3"
    conn = Update_database.create_sql_connection(database)
    c = conn.cursor()

    # Create the directory if already not there
    # if not os.path.exists(dir_path):
    #    os.mkdir(dir_path)

    # Function to take an image url and save the image in the given directory
    # def download_image(player_image, player_name):
    #    print("[INFO] downloading {}".format(player_image))
    #    name = str(player_image.split('/')[-1])
    #    urllib.request.urlretrieve(player_image,os.path.join(base_dir, player_name,".png"))

    # Get the URLs for each of the teams
    team_sites = soup.find_all('a', attrs={'class': 'th-club-nav__club-link'})

    # Counter for number of clubs to process
    i = 0  # A dd 1 to get the index number of the club on the AFL website

    for team_site in team_sites[i:18]:
        if i > 18:
            continue

        team_URL = team_site.get('href')
        team_name = team_site.find('span', attrs={'class': 'th-club-nav__club-name'}).get_text().strip()
        print(team_URL)

        # Get the site for the AFL team
        team_r = requests.get(team_URL)
        team_soup = BeautifulSoup(team_r.text, 'lxml')
        print("Checking players for", team_name)

        # Get the banner headings and find the AFL one to drill down into
        banner_headings = team_soup.find_all('a', class_='navigation__link navigation__link--in-drop-down')
        for banner_heading in banner_headings:
            if banner_heading.text.strip() == 'AFL' or banner_heading.text.strip() == "Men's" or banner_heading.text.strip() == 'Player Profiles':
                playerlist_URL = banner_heading.get('href')
                while playerlist_URL[0] == "/":
                    playerlist_URL = playerlist_URL[1:]
                playerlist_URL = team_URL + playerlist_URL
                if i == 11: playerlist_URL = playerlist_URL + '/players'  # Adjustment for North Melbourne website
                if i == 13: playerlist_URL = playerlist_URL + '/squad'  # Adjustment for Richmond website

                # Go to each player list page and get the links to the player pages
                playerlist_r = requests.get(playerlist_URL)
                player_soup = BeautifulSoup(playerlist_r.text, 'lxml')

                squad_list = player_soup.find_all('li', class_='squad-list__item')

                # Work through each player in the list
                for player in squad_list:

                    # Get the player name
                    player_firstname = player.h1.get_text().split()[0]
                    player_surname = player.h1.get_text().split()[1]
                    player_name = player_firstname + " " + player_surname

                    Sql_select_transition_table = """SELECT FirstName, Surname 
                                                     FROM transition 
                                                     WHERE FirstName=:player_firstname 
                                                         AND Surname=:player_surname 
                                                         AND Status = 'Current'"""
                    # val = (player_firstname, player_surname,)
                    c.execute(Sql_select_transition_table,
                              {"player_firstname": player_firstname, "player_surname": player_surname})

                    num_found = c.fetchall()
                    if num_found is None:
                        # Add player to database
                        print(player_firstname, player_surname, "not found in database")
        i += 1


def main():
    # get_all_current_players()
    # check_current_players()
    None


if __name__ == "__main__": main()
