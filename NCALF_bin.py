import sys
import time
import xlwings as xw
import pymysql
# import openpyxl
import re
from bs4 import BeautifulSoup
from selenium import webdriver
import pandas as pd
from playwright.sync_api import Playwright, sync_playwright
import json
import requests
import shutil

Numrounds = 24
NumNCALFClubs = 11
HoardingPenalty = -5
season = 2023

chromedriverfile = r"C:\ncalf\bin\NCALF v2020\ncalf_python_local\chromedriver.exe"


def convert_AFL_name_to_NCALF_format(clubname):
    if clubname in ['CROWS', 'Adelaide Crows', 'Adelaide']:
        return 'Adelaide'
    elif clubname in ['LIONS', 'Brisbane Lions', 'Brisbane']:
        return 'Brisbane'
    elif clubname in ['BLUES', 'Carlton Blues', 'Carlton']:
        return 'Carlton'
    elif clubname in ['MAGPIES', 'Collingwood Magpies', 'Collingwood']:
        return 'Collingwood'
    elif clubname in ['BOMBERS', 'Essendon Bombers', 'Essendon']:
        return 'Essendon'
    elif clubname in ['BULLDOGS', 'Western Bulldogs']:
        return "Western Bulldogs"
    elif clubname in ['DOCKERS', 'Fremantle Dockers', 'Fremantle']:
        return 'Fremantle'
    elif clubname in ['CATS', 'Geelong Cats', 'Geelong']:
        return 'Geelong'
    elif clubname in ['SUNS', 'Gold Coast Suns', 'Gold Coast']:
        return 'Gold Coast'
    elif clubname in ['GIANTS', 'GWS Giants', 'Greater Western Sydney', 'GWS']:
        return 'GWS'
    elif clubname in ['HAWKS', 'Hawthorn Hawks', 'Hawthorn']:
        return 'Hawthorn'
    elif clubname in ['DEMONS', 'Melbourne Demons', 'Melbourne']:
        return 'Melbourne'
    elif clubname in ['KANGAROOS', 'North Melbourne Kangaroos', 'North Melbourne']:
        return 'North Melbourne'
    elif clubname in ['POWER', 'Port Adelaide Power', 'Port Adelaide']:
        return 'Port Adelaide'
    elif clubname in ['TIGERS', 'Richmond Tigers', 'Richmond']:
        return 'Richmond'
    elif clubname in ['SAINTS', 'St Kilda Saints', 'St. Kilda', 'St Kilda']:
        return 'St Kilda'
    elif clubname in ['SWANS', 'Sydney Swans', 'Sydney']:
        return "Sydney"
    elif clubname in ['EAGLES', 'West Coast Eagles', 'West Coast']:
        return "West Coast"
    else:
        return 'No club found'


def convert_long_AFL_name_to_short_NCALF_format(clubname):
    if clubname in ['CROWS', 'Adelaide Crows', 'Adelaide']:
        return 'Ade'
    elif clubname in ['LIONS', 'Brisbane Lions', 'Brisbane']:
        return 'Bris'
    elif clubname in ['BLUES', 'Carlton Blues', 'Carlton']:
        return 'Carl'
    elif clubname in ['MAGPIES', 'Collingwood Magpies', 'Collingwood']:
        return 'Coll'
    elif clubname in ['BOMBERS', 'Essendon Bombers', 'Essendon']:
        return 'Ess'
    elif clubname in ['BULLDOGS', 'Western Bulldogs']:
        return "WB"
    elif clubname in ['DOCKERS', 'Fremantle Dockers', 'Fremantle']:
        return 'Fre'
    elif clubname in ['CATS', 'Geelong Cats', 'Geelong']:
        return 'Geel'
    elif clubname in ['SUNS', 'Gold Coast Suns', 'Gold Coast']:
        return 'GC'
    elif clubname in ['GIANTS', 'GWS Giants', 'Greater Western Sydney', 'GWS']:
        return 'GWS'
    elif clubname in ['HAWKS', 'Hawthorn Hawks', 'Hawthorn']:
        return 'Haw'
    elif clubname in ['DEMONS', 'Melbourne Demons', 'Melbourne']:
        return 'Melb'
    elif clubname in ['KANGAROOS', 'North Melbourne Kangaroos', 'North Melbourne']:
        return 'NM'
    elif clubname in ['POWER', 'Port Adelaide Power', 'Port Adelaide']:
        return 'PA'
    elif clubname in ['TIGERS', 'Richmond Tigers', 'Richmond']:
        return 'Rich'
    elif clubname in ['SAINTS', 'St Kilda Saints', 'St. Kilda', 'St Kilda']:
        return 'StK'
    elif clubname in ['SWANS', 'Sydney Swans', 'Sydney']:
        return "Syd"
    elif clubname in ['EAGLES', 'West Coast Eagles', 'West Coast']:
        return "WC"
    else:
        return 'No club found'


def convert_ncalfclub_code_to_name(clubcode):
    if clubcode in [1, 101, 'BU']:
        return 'Barnestoneworth United'
    elif clubcode in [2, 102, 'BB1']:
        return 'Berwick Blankets'
    elif clubcode in [3, 103, 'BB2']:
        return 'Bogong Bedouin'
    elif clubcode in [4, 104, 'BB3']:
        return 'Bohemian Buffali'
    elif clubcode in [5, 105, 'GKR']:
        return 'G. K. Rovers'
    elif clubcode in [6, 106, 'JJ']:
        return 'Jancourt Jackrabbits'
    elif clubcode in [7, 107, 'KP']:
        return 'Kamarah Paddockbashers'
    elif clubcode in [8, 108, 'KC']:
        return 'Kennedy Celtics'
    elif clubcode in [9, 109, 'LH']:
        return 'Laughing Hyenas'
    elif clubcode in [10, 110, 'RR']:
        return 'Rostron Redbacks'
    elif clubcode in [11, 111, 'SS']:
        return 'Southern Squadron'
    else:
        return 'No club found'


def convert_ncalfclub_name_to_code(clubname):
    if clubname in ['Barnestoneworth United', 'BU']:
        return 1
    elif clubname in ['Berwick Blankets', 'BB1']:
        return 2
    elif clubname in ['Bogong Bedouin', 'BB2']:
        return 3
    elif clubname in ['Bohemian Buffali', 'BB3']:
        return 4
    elif clubname in ['G. K. Rovers', 'GKR']:
        return 5
    elif clubname in ['Jancourt Jackrabbits', 'JJ']:
        return 6
    elif clubname in ['Kamarah Paddockbashers', 'KP']:
        return 7
    elif clubname in ['Kennedy Celtics', 'KC']:
        return 8
    elif clubname in ['Laughing Hyenas', 'LH']:
        return 9
    elif clubname in ['Rostron Redbacks', 'RR']:
        return 10
    elif clubname in ['Southern Squadron', 'SS']:
        return 11
    else:
        return 99


def convert_position_to_index(position):
    position_list = ['BPL', 'FB', 'BPR', 'HBFL', 'CHB', 'HBFR', 'WL', 'C', 'WR', 'HFFL', 'CHF', 'HFFR', 'FPL', 'FF',
                     'FPR', 'RK', 'RR', 'R', 'INT', 'SUB']
    return position_list.index(position) + 1


def create_sql_connection():
    # create a database connection to the SQL database
    conn = None
    try:
        # Connect to the database
        # passwd = input("Enter database password:")
        passwd = 'Jo!@04067324'

        conn = pymysql.connect(host='localhost',
                               user='admin',
                               password=passwd,
                               db='ncalfdb',
                               charset='utf8mb4',
                               cursorclass=pymysql.cursors.DictCursor)
        return conn
    except Exception as e:
        print(e)
    return conn


def output_to_text_file(source):
    filename = r'c:\ncalf\bin\html.txt'
    with open(filename, 'w', encoding="utf-8") as f:
        f.write(source)


def Get_player_positions_AFL(soup):
    game = soup.find('div', attrs={'class': 'team-lineups__wrapper'})

    # Get the home and away team names
    home_team_tag = game.find('span', attrs={
        'class': re.compile('team-lineups__team-name team-lineups__team-name--home active js-team-tab')})
    home_team = home_team_tag.get_text().strip()
    away_team_tag = game.find('span', attrs={
        'class': re.compile('team-lineups__team-name team-lineups__team-name--away js-team-tab')})
    away_team = away_team_tag.get_text().strip()

    player_pos_list = []

    # Get the names and positions for each of the players
    playerrows = game.find_all('div', attrs={'class': 'team-lineups__positions-row'})

    for playerrow in playerrows:
        if playerrows.index(playerrow) < 6:
            players = playerrow.find_all('span', attrs={'class': 'team-lineups__player'})
            for player in players:
                # if the players on named on field
                if players.index(player) <= 2:
                    player_position = (playerrows.index(playerrow) * 3) + players.index(player) + 1
                    team = home_team
                else:
                    player_position = (playerrows.index(playerrow) * 3) + players.index(player) - 2
                    team = away_team
                player_firstname = player.get_text().split()[1]

                # Check for middle initials and ignore
                if player.get_text().split()[2][-1] == '.':
                    player_surname = ' '.join(player.get_text().split()[3:])
                elif len(player.get_text().split()) == 4:
                    player_surname = ' '.join(player.get_text().split()[2:])
                else:
                    player_surname = player.get_text().split()[2]

                # Remove the trailing comma
                if player_surname[-1] == ',':
                    player_surname = player_surname[:-1]

                # Add the row to the player list
                player_pos_list.append(
                    [convert_AFL_name_to_NCALF_format(team), player_firstname, player_surname, str(player_position)])

        # if the players on named on the interchange bench
        else:
            benches = playerrow.find_all('div', attrs={'class': 'team-lineups__positions-players'})
            for bench in benches:
                bench_players = bench.find_all('span', attrs={'class': 'team-lineups__player'})
                # Set the home or away team
                if benches.index(bench) == 0:
                    team = home_team
                else:
                    team = away_team
                for bench_player in bench_players:
                    player_position = 19 + bench_players.index(bench_player)
                    player_firstname = bench_player.get_text().split()[1]

                    # Check for middle initials and ignore
                    if bench_player.get_text().split()[2][-1] == '.':
                        player_surname = ' '.join(bench_player.get_text().split()[3:])
                    elif len(bench_player.get_text().split()) == 4:
                        player_surname = ' '.join(bench_player.get_text().split()[2:])
                    else:
                        player_surname = bench_player.get_text().split()[2]

                    # Remove the trailing comma
                    if player_surname[-1] == ',':
                        player_surname = player_surname[:-1]

                    # Add the row to the player list
                    player_pos_list.append([convert_AFL_name_to_NCALF_format(team), player_firstname, player_surname,
                                            str(player_position)])
    return player_pos_list


global json_stats
global json_match
global json_round


def onResponse(r):
    global json_stats
    global json_match
    global json_round

    stats_filter = "api.afl.com.au/cfs/afl/playerStats/match/"
    match_filter = "api.afl.com.au/cfs/afl/matchRoster/full/"
    round_filter = "api.afl.com.au/broadcasting/match-events?competition=1"

    if stats_filter in r.url:
        json_stats = json.loads(r.body().decode("utf-8"))

    if match_filter in r.url:
        json_match = json.loads(r.body().decode("utf-8"))

    if round_filter in r.url:
        json_round = json.loads(r.body().decode("utf-8"))


def run(playwright: Playwright, url) -> None:
    browser = playwright.chromium.launch(channel="chrome")
    context = browser.new_context()
    page = context.new_page()

    page.on("response", onResponse)
    page.goto(url)

    context.close()
    browser.close()


def Download_player_stats_AFL(season, roundno):
    # Downloads the player stats and positions from the AFL website and loads them into the database

    # Connect to the database
    conn = create_sql_connection()
    c = conn.cursor()

    # games = soup.find_all('div', attrs={'class': "match-list__match-center"})

    # Check to delete all existing stats entries for that round
    print("Deleting existing stats:")
    start_time = time.time()
    sqlstring = "SELECT PlayerSeasonID FROM stats WHERE Season = %s AND Round = %s AND " \
                "(k + m + hb + ff + fa + g + b + ho + t) > 0"
    val = (season, roundno)
    c.execute(sqlstring, val)
    existing_stats_list = c.fetchall()
    if len(existing_stats_list) > 0:
        # Delete the existing stats
        sqlstring = """UPDATE stats SET posplyd = 0, k = 0, m = 0, hb = 0, ff = 0, fa = 0, g = 0, b = 0, ho = 0, t = 0
                            WHERE Season = %s AND Round = %s"""
        val = (season, roundno)
        c.execute(sqlstring, val)
        conn.commit()
    end_time = time.time()
    print("Time taken:" + str(end_time - start_time))

    print('Downloading the player stats from the AFL website')

    # Get the round information from the network response for the round
    round_URL = "https://www.afl.com.au/fixture?Competition=1&CompSeason=52&MatchTimezone=MY_TIME&Regions=2" \
                  "&ShowBettingOdds=1&GameWeeks=" + str(roundno) + "&Teams=1&Venues=3#byround"

    with sync_playwright() as playwright:
        run(playwright, round_URL)

    roundmatch_ids = []

    for match in range(0, 6):
        roundmatch_ids.append(json_round["content"][match]["contentReference"]["id"])

    # for game in games:
    for roundmatch_id in roundmatch_ids:

        # Get the match id
        # match_url_id = game.find_next('a').get('href')
        # match_url = "https://www.afl.com.au" + match_url_id + "#player-stats"
        match_url = "https://www.afl.com.au/afl/matches/" + str(roundmatch_id) + "#player-stats"
        print("Downloading stats from:" + match_url)
        start_time = time.time()

        # Get the match and stats data from the network response for the round
        with sync_playwright() as playwright:
            run(playwright, match_url)

        teamstats = []
        for team in ['homeTeamPlayerStats', 'awayTeamPlayerStats']:
            for playerrow in json_stats[team]:
                if team == 'homeTeamPlayerStats':
                    team_name = json_match['match']['homeTeam']['name']
                if team == 'awayTeamPlayerStats':
                    team_name = json_match['match']['awayTeam']['name']

                playerstats = {'club': convert_AFL_name_to_NCALF_format(team_name),
                               'playerfirstname': playerrow['player']['player']['player']['playerName']['givenName'],
                               'playersurname': playerrow['player']['player']['player']['playerName']['surname'],
                               'posplyd': convert_position_to_index(playerrow['player']['player']['position']),
                               'kicks': str(int(playerrow['playerStats']['stats']['kicks'])),
                               'marks': str(int(playerrow['playerStats']['stats']['marks'])),
                               'handballs': str(int(playerrow['playerStats']['stats']['handballs'])),
                               'freesfor': str(int(playerrow['playerStats']['stats']['freesFor'])),
                               'freesagainst': str(int(playerrow['playerStats']['stats']['freesAgainst'])),
                               'goals': str(int(playerrow['playerStats']['stats']['goals'])),
                               'behinds': str(int(playerrow['playerStats']['stats']['behinds'])),
                               'hitouts': str(int(playerrow['playerStats']['stats']['hitouts'])),
                               'tackles': str(int(playerrow['playerStats']['stats']['tackles']))
                               }

                # Add the stats line to the team stats list
                teamstats.append(playerstats)

        end_time = time.time()
        print("Time taken:" + str(end_time - start_time))

        # Import the teamstats list to the database
        print("Import stats to database:")
        start_time = time.time()

        # Initialise the val_list variable for the final update statement
        val_list = []

        # Get all the players for the round into a list
        head_sqlstring = """SELECT S.PlayerSeasonID, P.PlayerID, P.PlayerFirstName, P.PlayerSurname, 
        P.AltPlayerFirstName, P.AltPlayerSurname, S.Club, S.Position, S.k, S.m, S.hb, S.ff, S.fa, S.g, S.b, S.ho, 
        S.t FROM stats S INNER JOIN Players P ON P.PlayerID = S.playerID WHERE S.Season = %s AND S.Round = %s """
        head_val = (season, roundno)
        c.execute(head_sqlstring, head_val)
        round_player_list = c.fetchall()

        for playerstats in teamstats:

            # Check that the player exists in the round_player_list
            found_player = [x for x in round_player_list if
                            ((x['PlayerFirstName'] == playerstats['playerfirstname'] or
                              x['AltPlayerFirstName'] == playerstats['playerfirstname']) and
                             (x['PlayerSurname'] == playerstats['playersurname'] or
                              x['AltPlayerSurname'] == playerstats['playersurname']) and
                             x['Club'] == convert_long_AFL_name_to_short_NCALF_format(playerstats['club']))]

            # If the player not found
            if len(found_player) == 0:

                # Print the player list for the club that is being processed
                print(
                    "Player list for " + playerstats['club'] + ":")

                # Get the list of players to choose from
                club_player_list = [x for x in round_player_list if
                                    (x['Club'] == convert_long_AFL_name_to_short_NCALF_format(playerstats['club']))]

                # Offer up the names of players to choose from , plus option to add a new player
                for club_player in club_player_list:
                    print(str(club_player['PlayerSeasonID']) + ' ' + club_player['PlayerFirstName'] + ' ' + club_player[
                        'PlayerSurname'] + ' (' + club_player['Club'] + ')')

                # Print the name of the player not found
                print(
                    "Player not found: " + playerstats['playerfirstname'] + ' ' + playerstats['playersurname'] + ' (' +
                    playerstats['club'] + ')')

                choice = input("Choose the player ID, press A to add a new player, or C to check for old players:")

                # If user chooses to add a new player
                if choice == "A" or choice == "a":
                    # Offer to confirm player details (firstname, surname, club) from playerstats in Rookie position
                    print("Adding new player: " + playerstats['playerfirstname'] + " " + playerstats['playersurname'] +
                          " (" + playerstats['club'] + ")")

                    # feedback = input("Add as a Rookie? (Y/N)")
                    feedback = "Y"
                    # If yes
                    if feedback == "Y" or feedback == "y":
                        # Create the new playerID for the new player
                        # Get the highest playerID from the players table and add 1
                        sqlstring = """SELECT MAX(PlayerID) AS maxid 
                                            FROM Players"""
                        c.execute(sqlstring)
                        record = c.fetchone()
                        highest_playerID = record['maxid']

                        # Add player to player table
                        sqlstring = """INSERT INTO Players (PlayerID, PlayerFirstName, PlayerSurname)
                                        VALUES (%s, %s, %s) """
                        val = (highest_playerID + 1, playerstats['playerfirstname'], playerstats['playersurname'])
                        c.execute(sqlstring, val)
                        conn.commit()

                        # Create new playerseasonID
                        # Get the highest playerseasonID from the stats table and add 1
                        sqlstring = """SELECT MAX(PlayerSeasonID) AS maxid
                                        FROM stats
                                        WHERE Season = %s"""
                        val = (season,)
                        c.execute(sqlstring, val)
                        record = c.fetchone()
                        highest_playerseasonID = record['maxid']

                        # Add player to stats table for all rounds
                        for roundnum in range(1, Numrounds + 1):
                            sqlstring = """INSERT INTO stats (Season, Round, PlayerID, PlayerSeasonID, Club, 
                            Position, posplyd, ncalfclub, k, m, hb, ff, fa, g, b, ho, t) 
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) """
                            val = (season, roundnum, int(highest_playerID) + 1, int(highest_playerseasonID) + 1,
                                   convert_long_AFL_name_to_short_NCALF_format(playerstats['club']), 'Rookie', 0, 0, 0,
                                   0, 0, 0, 0, 0, 0, 0, 0)
                            c.execute(sqlstring, val)
                        conn.commit()

                        # Save playerseasonID in choice variable for stats entry later
                        choice = int(highest_playerseasonID) + 1

                        # Update the round_player_list
                        c.execute(head_sqlstring, head_val)
                        round_player_list = c.fetchall()

                        # # Refresh the SQL player list
                        # c.execute(head_sqlstring, head_val)
                        # round_player_list = c.fetchall()
                        #
                        # # Get the list of players to choose from
                        #
                        # club_player_list = [x for x in round_player_list if (
                        #         x['Club'] == convert_long_AFL_name_to_short_NCALF_format(playerstats['club']))]
                        #
                        # # Offer up the names of players to choose from , plus option to add a new player
                        # for club_player in club_player_list:
                        #     print(str(club_player['PlayerSeasonID']) + ' ' + club_player['PlayerFirstName'] + ' ' +
                        #           club_player['PlayerSurname'] + ' (' + club_player['Club'] + ')')
                    # If no
                    else:
                        # Exit message and quit
                        print("Exiting as player position not chosen")
                        exit()

                # If checking for players that have recently retired
                if choice == "C" or choice == "c":

                    # Get a list of all historical players that share the same name
                    sqlstring = """SELECT PlayerID, CONCAT(PlayerID, ' ', PlayerFirstName, ' ', PlayerSurname) AS name
                                    FROM Players
                                    WHERE (PlayerFirstName = %s OR AltPlayerFirstName = %s) AND
                                    (PlayerSurname = %s OR AltPlayerSurname = %s)"""
                    val = (playerstats['playerfirstname'], playerstats['playerfirstname'], playerstats['playersurname'],
                           playerstats['playersurname'])
                    c.execute(sqlstring, val)
                    club_player_list = c.fetchall()

                    # Create new playerseasonID
                    # Get the highest playerseasonID from the stats table and add 1
                    sqlstring = """SELECT MAX(PlayerSeasonID) AS maxid
                                    FROM stats
                                    WHERE Season = %s"""
                    val = (season,)
                    c.execute(sqlstring, val)
                    record = c.fetchone()
                    highest_playerseasonID = record['maxid']

                    # Print the list of players that share the same name
                    for club_player in club_player_list:
                        print(club_player['name'])

                    # Ask user to choose
                    choice = input("Enter the Player ID for the correct player:")

                    # Check that they've entered a valid PlayerID from the list
                    for club_player in club_player_list:
                        if int(choice) == club_player['PlayerID']:
                            # Add player to stats table for all rounds
                            for roundnum in range(1, Numrounds + 1):
                                sqlstring = """INSERT INTO stats (Season, Round, PlayerID, PlayerSeasonID, Club, 
                                Position, posplyd, ncalfclub, k, m, hb, ff, fa, g, b, ho, t) 
                                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s) """
                                val = (season, roundnum, int(choice), int(highest_playerseasonID) + 1,
                                       convert_long_AFL_name_to_short_NCALF_format(playerstats['club']), 'Rookie', 0, 0,
                                       0, 0, 0, 0, 0, 0, 0, 0, 0)
                                c.execute(sqlstring, val)
                            conn.commit()
                        else:
                            print("Incorrect PlayerID chosen. Exiting")
                            exit()

                    # Save playerseasonID in choice variable for stats entry later
                    choice = int(highest_playerseasonID) + 1

                    # Update the round_player_list
                    c.execute(head_sqlstring, head_val)
                    round_player_list = c.fetchall()

                # If player chooses a player
                # if choice.isnumeric():
                # Get the player name
                found_player = [x for x in round_player_list if (x['PlayerSeasonID'] == int(choice))]

                # Store the data
                playerID = found_player[0]['PlayerID']
                playerseasonID = found_player[0]['PlayerSeasonID']
                selected_firstname = found_player[0]['PlayerFirstName']
                selected_surname = found_player[0]['PlayerSurname']
                alt_firstname = found_player[0]['AltPlayerFirstName']
                alt_surname = found_player[0]['AltPlayerSurname']
                # Compare the first name
                # If it is different, offer to add new player firstname to the player table as alternative name
                if not (selected_firstname == playerstats['playerfirstname']) or (
                        alt_firstname == playerstats['playerfirstname']):
                    msgstring = "The stored first name is '" + selected_firstname + \
                                "' and the selected first name is '" + \
                                playerstats['playerfirstname'] + "'\r. Do you want to save '" + \
                                playerstats[
                                    'playerfirstname'] + "' as an alternative first name in the database? (Y/N)"
                    choice = input(msgstring)
                    # If accepted, add the new player firstname to the player table as alternative name
                    if choice == "y" or choice == "Y":
                        sqlstring = "UPDATE players SET AltPlayerFirstName = %s WHERE PlayerID = %s"
                        val = (playerstats['playerfirstname'], playerID)
                        c.execute(sqlstring, val)
                        conn.commit()
                # Compare the surname
                if not ((selected_surname == playerstats['playersurname']) or (
                        alt_surname == playerstats['playersurname'])):
                    msgstring = "The stored surname is '" + selected_surname + "' and the selected surname is '" + \
                                playerstats['playersurname'] + "'.\rDo you want to save '" + playerstats[
                                    'playersurname'] + "' as an alternative surname in the database? (Y/N)"
                    choice = input(msgstring)
                    # If accepted, add the new player surname to the player table as alternative name
                    if choice == "y" or choice == "Y":
                        sqlstring = "UPDATE Players SET AltPlayerSurname = %s WHERE PlayerID = %s"
                        val = (playerstats['playersurname'], playerID)
                        c.execute(sqlstring, val)
                        conn.commit()

                # If offer accepted, update the stats table with the player stats and position information
                # If offer not accepted, provide message and exit
                # else:
                #     playerseasonID = found_player[0]['PlayerSeasonID']
            else:
                playerseasonID = found_player[0]['PlayerSeasonID']

            # Update the stats list with the player stats and position information
            val_list.append(
                [playerstats['posplyd'], playerstats['kicks'], playerstats['marks'], playerstats['handballs'],
                 playerstats['freesfor'], playerstats['freesagainst'], playerstats['goals'], playerstats['behinds'],
                 playerstats['hitouts'], playerstats['tackles'], season, roundno, playerseasonID])

        # Update the stats table from the list
        sqlstring = """UPDATE stats
                        SET posplyd = %s, k = %s, m = %s, hb = %s, ff = %s,
                            fa = %s, g = %s, b = %s, ho = %s, t = %s
                        WHERE Season = %s AND Round = %s AND PlayerSeasonID = %s"""
        c.executemany(sqlstring, val_list)
        end_time = time.time()
        print("Time taken:" + str(end_time - start_time))
    conn.commit()


def calculate_results(season, roundno):
    # Calculates the results for the current round after stats and positions are imported
    print("Calculating the results...")
    print("Connecting to the database...")
    sql_conn = create_sql_connection()
    sql_cursor = sql_conn.cursor()

    # Delete current points results
    print("Deleting current points results...")
    sqlstring = """UPDATE points SET pki = 0, pm = 0, phb = 0, pf = 0, psc = 0, pho = 0, pt = 0, hoarding = 0 
                        WHERE Season = %s AND Round = %s"""
    val = (season, roundno)
    sql_cursor.execute(sqlstring, val)
    sql_conn.commit()

    # Get the stats for the points calculations
    # Need to watch for case when fa = 0
    sqlcommand = """SELECT ncalfclub, sum(k) as sk, sum(m) as sm, sum(hb) as shb, sum(ff) as sff, sum(fa) as sfa, 
                    round(sum(ff), 1) / round(sum(fa), 1) AS fratio, sum(g) as sg, sum(b) as sb, sum(g*6 + b) as ssc, 
                    sum(ho) as sho, sum(t) as st 
                    FROM stats 
                    WHERE Season = %s AND Round = %s AND ncalfclub > 0 AND ncalfclub < %s 
                    GROUP BY ncalfclub 
                    ORDER BY ncalfclub"""
    val = (season, roundno, (NumNCALFClubs + 1))
    print("Retrieving the stats from the database...")
    sql_cursor.execute(sqlcommand, val)
    stats_totals = sql_cursor.fetchall()

    print("Calculating the points...")
    team_points_table = [[1], [2], [3], [4], [5], [6], [7], [8], [9], [10], [11]]
    for category in ['sk', 'sm', 'shb', 'fratio', 'ssc', 'sho', 'st']:
        stats_category = []
        for stats_total in stats_totals:
            stats_category.append(
                {'ncalfclub': stats_total['ncalfclub'], 'stat': float(stats_total[category]), 'points': 0})

        # Sort the list based on the stats
        sorted_stats_list = sorted(stats_category, key=lambda k: k['stat'])

        # Assign points within sorted list
        StartTeamIndex = 0
        while StartTeamIndex < NumNCALFClubs:
            points = StartTeamIndex + 1
            EndTeamIndex = StartTeamIndex
            NumSamePoints = 1
            currstat = sorted_stats_list[StartTeamIndex]['stat']
            if EndTeamIndex < NumNCALFClubs - 1:
                while sorted_stats_list[EndTeamIndex + 1]['stat'] == currstat:
                    NumSamePoints += 1
                    EndTeamIndex += 1
                    points += (EndTeamIndex + 1)
                    if EndTeamIndex == NumNCALFClubs - 1:  # If reach the end of the list, exit the while loop
                        break
            allocated_points = points / NumSamePoints
            for TeamIndex in range(StartTeamIndex, EndTeamIndex + 1):
                sorted_stats_list[TeamIndex]['points'] = allocated_points
            StartTeamIndex = EndTeamIndex + 1

        # Sort the list back into NCALF club order
        stats_list = sorted(sorted_stats_list, key=lambda k: k['ncalfclub'])
        # Update the team points table
        for stat in stats_list:
            team_points_table[stat['ncalfclub'] - 1].append(stat['points'])

    # Add the hoarding points to the table as zero
    for i in range(NumNCALFClubs):
        team_points_table[i].append(0)

    # Check for hoarding
    print("Checking for hoarding...")
    # Get the players on the reserve list who got stats in the past 3 rounds incl current round
    sqlcommand = """SELECT S.PlayerID, S.ncalfclub, P.PlayerFirstName, P.PlayerSurname, S.Club, S.ncalfclub FROM 
                        (SELECT Round, count(PlayerID) as num_rounds, PlayerID, ncalfclub, Club FROM stats 
                        WHERE Season = %s AND Round >= %s AND  (k + m + hb + ff + fa + g + b + ho + t) > 0 
                        AND ncalfclub >= 100
                        GROUP by PlayerID 
                        ORDER BY PlayerID, Round desc) as S
                    INNER JOIN
                    players P ON P.PlayerID = S.PlayerID
                    WHERE S.num_rounds > 2"""
    val = (season, roundno - 2)
    sql_cursor.execute(sqlcommand, val)
    hoarded_player_list = sql_cursor.fetchall()

    for hoarded_player in hoarded_player_list:
        # Player has been hoarded
        choice = input(
            "{} {} {} ({}) has been hoarded by the {}. "
            "Process the offence? (Y/N): ".format(hoarded_player['PlayerID'],
                                                  hoarded_player['PlayerFirstName'], hoarded_player['PlayerSurname'],
                                                  hoarded_player['Club'],
                                                  convert_ncalfclub_code_to_name(hoarded_player['ncalfclub'])))
        if choice in ['Y', 'y']:
            # Remove the player from the team's reserve list
            sqlcommand = """UPDATE stats 
                            SET ncalfclub = 0 
                            WHERE Season = %s AND Round = %s AND 
                            PlayerID = %s"""
            val = (season, roundno, hoarded_player['PlayerID'])
            sql_cursor.execute(sqlcommand, val)

            # Apply the hoarding penalty to the offending team
            team_points_table[hoarded_player['ncalfclub'] - 100 - 1][8] += HoardingPenalty

            print(
                "{} {} {} ({}) has been removed from the reserve list of the {}. "
                "A hoarding penalty of {} points has also been applied.".format(
                    hoarded_player['PlayerID'], hoarded_player['PlayerFirstName'], hoarded_player['PlayerSurname'],
                    hoarded_player['Club'], convert_ncalfclub_code_to_name(hoarded_player['ncalfclub']),
                    HoardingPenalty))

    # Update the database
    print("Updating the database with the points...")
    for team in team_points_table:
        sqlcommand = """UPDATE points 
                        SET pki = %s, pm = %s, phb = %s, pf = %s, psc = %s, pho = %s, pt = %s, hoarding = %s 
                        WHERE ncalfclub = %s AND Season = %s AND Round = %s"""
        val = (team[1], team[2], team[3], team[4], team[5], team[6], team[7], team[8], team[0], season, roundno)
        sql_cursor.execute(sqlcommand, val)

    sql_conn.commit()


def import_changes(season, roundno):
    # Imports the changes from a csv file which is output from the changes spreadsheet

    # Connect to the database
    conn = create_sql_connection()
    c = conn.cursor()

    # Bring the changes from the excel changes file into changes_list
    print('Reading the changes in from the changes spreadsheet file')
    excel_app = xw.App(visible=False)
    print('Opening the workbook')
    wb = excel_app.books.open(r'C:\Users\hener\Dropbox\NCALF\Changes\NCALF ' + str(season) + r' changes.xlsx')
    changes_sheet = wb.sheets(r'Round ' + str(roundno))
    changes = changes_sheet.range('A3').options(pd.DataFrame, header=1, index=False, expand='table').value
    # Replace the nan values in the dataframe with a ''
    changes['Player In ID'] = changes['Player In ID'].fillna('')
    changes_list = changes.values.tolist()

    print('Close the spreadsheet')
    wb.close()
    excel_app.quit()

    # Delete any changes that already exist
    print('Delete existing changes')
    change_sql_string = 'DELETE FROM changes WHERE Season = %s AND round = %s'
    change_val = (season, roundno)
    c.execute(change_sql_string, change_val)
    conn.commit()

    # Get a list of all players from the database including ncalfclubs for the previous round
    print('Get the players from the database for the previous round')
    players_sqlstring = """SELECT S.Round, S.PlayerSeasonID, PL.PlayerFirstname, PL.PlayerSurname, S.Club, S.ncalfclub,
                                S.Position
                            FROM
                                players AS PL
                                    INNER JOIN
                                stats S ON S.PlayerID = PL.PlayerID
                            WHERE
                                S.Season = %s AND S.Round = %s
                            ORDER BY S.PlayerSeasonID"""
    players_val = (season, roundno - 1)
    c.execute(players_sqlstring, players_val)
    players_list = c.fetchall()

    print('Process the changes')

    # Add a changed flag to the player list: 0 - no change, 1 - change
    for players in players_list:
        players['change'] = 0
        players['pos_change'] = 0

    # Get max playerseasonID
    seq = [x['PlayerSeasonID'] for x in players_list]
    max_playerseasonID = max(seq)

    # For each change entry
    for change in changes_list:

        # Get the relevant details from the changes_list entry
        position = change[0]
        ncalfclub = change[1]
        if change[2] == '':
            plyr_in_id = ''
        else:
            plyr_in_id = str(int(change[2]))
        plyr_in_name = change[3]
        change_type = change[4]
        plyr_out_id = str(int(change[5]))
        plyr_out_name = change[6]

        # Set the change variables
        in_from = ''
        out_to = ''

        # Check that the change is valid
        # Check the incoming player_id
        if plyr_in_id is not '':  # In case it is a sack from reserve change
            if int(plyr_in_id) > max_playerseasonID:
                print("Error: Incoming player ID: " + plyr_in_id + ' for the ' + ncalfclub + ' is out of range. '
                                                                                             'Please check. Exiting.')
                sys.exit()

        # Check the outgoing player_id
        if int(plyr_out_id) > max_playerseasonID:
            print("Error: Outgoing player ID: " + plyr_out_id + ' for the ' + ncalfclub + ' is out of range. '
                                                                                          'Please check. Exiting.')
            sys.exit()

        # Get the index of the outgoing player in the players list
        players_id_list = [x['PlayerSeasonID'] for x in players_list]
        index_out = players_id_list.index(int(plyr_out_id))

        # Configure the decision variables according to the type of change
        if change_type == 'Reserve' or change_type == 'Sack':

            # Get the index of the incoming player in the players list
            index_in = players_id_list.index(int(plyr_in_id))

            # Check that incoming player is not in a side already or already on the reserve list of the ncalf club
            if players_list[index_in]['ncalfclub'] > 0 and not \
                    players_list[index_in]['ncalfclub'] == (convert_ncalfclub_name_to_code(ncalfclub) + 100):
                print(
                    "Error: Incoming player: " + plyr_in_id + ' for the ' + ncalfclub + ' is already in a side. '
                                                                                        'Please recheck. Exiting.')
                sys.exit()

            # Check that outgoing player is actually in the current side
            if players_list[index_out]['ncalfclub'] != convert_ncalfclub_name_to_code(ncalfclub):
                print("Error: Outgoing player: " + plyr_out_id + ' for the ' + ncalfclub +
                      ' is not already in the side. Please recheck. Exiting.')
                sys.exit()

            # If incoming player is a rookie or a mixed position player, change his position
            if (players_list[index_in]['Position'] in ('ROOK', 'Rookie')) or '/' in players_list[index_in]['Position']:
                players_list[index_in]['Position'] = position
                players_list[index_in]['pos_change'] = 1

            # If the two players are in different positions then prompt error
            plyr_in_pos = players_list[index_in]['Position']
            plyr_out_pos = players_list[index_out]['Position']
            if plyr_in_pos != plyr_out_pos:
                print("Error: Incoming player: " + plyr_in_id + ' and outgoing player: ' + plyr_out_id +
                      ' for the ' + ncalfclub + ' are in different positions. Please recheck. Exiting.')
                sys.exit()

            # Check if the incoming player is on reserve (so that info can be stored in changes table)
            if players_list[index_in]['ncalfclub'] == (convert_ncalfclub_name_to_code(ncalfclub) + 100):
                in_from = 'Reserve'
            else:
                in_from = 'Pool'

            # Make the change in the players table
            if change_type == 'Reserve':
                players_list[index_in]['ncalfclub'] = convert_ncalfclub_name_to_code(ncalfclub)
                players_list[index_out]['ncalfclub'] = convert_ncalfclub_name_to_code(ncalfclub) + 100
                players_list[index_in]['change'] = 1
                players_list[index_out]['change'] = 1
                out_to = 'Reserve'
            if change_type == 'Sack':
                players_list[index_in]['ncalfclub'] = convert_ncalfclub_name_to_code(ncalfclub)
                players_list[index_out]['ncalfclub'] = 0
                players_list[index_in]['change'] = 1
                players_list[index_out]['change'] = 1
                out_to = 'Pool'

        if change_type == 'Sack from reserve':
            # Check that an incoming player is not specified for a sack from reserve change
            if plyr_in_id is not '':
                print("Error: Incoming player: " + plyr_in_id + ' for the ' + ncalfclub + ' specified for '
                                                                                          'Sack from Reserve change. '
                                                                                          'Please check. Exiting.')
                sys.exit()

            # Check that outgoing player is actually in the current side and on reserve
            if players_list[index_out]['ncalfclub'] != convert_ncalfclub_name_to_code(ncalfclub) + 100:
                print("Error: Outgoing player: " + plyr_out_id + ' for the ' + ncalfclub +
                      ' is not already on reserve in the side. Please recheck. Exiting.')
                sys.exit()

            # Make the change - sack from reserve
            players_list[index_out]['ncalfclub'] = 0
            players_list[index_out]['change'] = 1
            out_to = 'Pool'

        # Update the changes table
        change_sql_string = 'INSERT INTO changes (ncalfclub, Season, round, pos, plyrin, plyrout, infrom, outto) ' \
                            'VALUES (%s, %s, %s, %s, %s, %s, %s, %s)'
        change_val = (convert_ncalfclub_name_to_code(ncalfclub), season, roundno, position,
                      ' '.join([plyr_in_id, plyr_in_name]), ' '.join([plyr_out_id, plyr_out_name]), in_from, out_to)
        c.execute(change_sql_string, change_val)

    # Check that all teams have the right players in the right positions
    print('Checking number players in positions and on reserve')
    for club in range(1, NumNCALFClubs + 1):
        for pos in ('C', 'D', 'F', 'RK', 'OB'):
            pos_list = [x for x in players_list if x['ncalfclub'] == club and x['Position'] == pos]
            if pos == 'C' and len(pos_list) != 3:
                print('Error: Incorrect number of centres in ' + convert_ncalfclub_code_to_name(club) +
                      '. Check changes.')
                sys.exit()
            if pos == 'D' and len(pos_list) != 8:
                print('Error: Incorrect number of defenders in ' + convert_ncalfclub_code_to_name(club) +
                      '. Check changes.')
                sys.exit()
            if pos == 'F' and len(pos_list) != 8:
                print('Error: Incorrect number of forwards in ' + convert_ncalfclub_code_to_name(club) +
                      '. Check changes.')
                sys.exit()
            if pos == 'RK' and len(pos_list) != 1:
                print('Error: Incorrect number of rucks in ' + convert_ncalfclub_code_to_name(club) +
                      '. Check changes.')
                sys.exit()
            if pos == 'OB' and len(pos_list) != 2:
                print('Error: Incorrect number of onballers in ' + convert_ncalfclub_code_to_name(club) +
                      '. Check changes.')
                sys.exit()

    # Check the number of players on reserve is within limits
    for club in range(101, 100 + NumNCALFClubs + 1):
        reserve_list = [x for x in players_list if x['ncalfclub'] == club]
        if len(reserve_list) > 4:
            print('Error: Too many players on reserve in the ' + convert_ncalfclub_code_to_name(club) +
                  '. Check changes.')
            sys.exit()

    # Update the stats with the new updated contents of the players_list
    print('Update the stats table')
    changed_players_list = [x for x in players_list if x['change'] == 1 or x['ncalfclub'] > 0]
    for change in changed_players_list:
        print("Updating changes details: " + str(changed_players_list.index(change)) + " of " +
              str(len(changed_players_list)))
        change_sql_string = 'UPDATE stats SET ncalfclub = %s WHERE Season = %s AND Round = %s AND PlayerSeasonID = %s'
        change_val = (change['ncalfclub'], season, roundno, change['PlayerSeasonID'])
        c.execute(change_sql_string, change_val)

        if change['pos_change']:
            changepos_sql_string = 'UPDATE stats SET Position = %s ' \
                                   'WHERE Season = %s AND Round >= %s AND PlayerSeasonID = %s'
            changepos_val = (change['Position'], season, roundno, change['PlayerSeasonID'])
            c.execute(changepos_sql_string, changepos_val)

    # Commit all changes to the database
    print('Commit changes to database')
    conn.commit()


def output_results_to_new_excel(df_stats, df_rpoints, newroundno):
    # Opening the NCALF reports spreadsheet
    print('OUTPUT THE RESULTS TO THE NEW SPREADSHEET FORMAT')
    print('Opening Excel')
    excel_app = xw.App(visible=False)
    print('Opening the workbook')
    wb = excel_app.books.open(r'c:\ncalf\NCALF ' + str(season) + r'\NCALF ' + str(season) + r' results report.xlsx')
    stats_sheet = wb.sheets('Data')

    # print('Clear any current round stats from the spreadsheet')
    # round_stats = stats_sheet.tables('stats').api.autofilter.apply('stats', 1, '=' + str(round))
    # round_stats.delete(shift='up')
    # stats_sheet.range('stats').api.autofilter.clearCriteria()

    print('Write the results to the spreadsheet')
    # Write the stats to the spreadsheet
    # Find the stats row to start writing to
    if newroundno == 1:
        last_stats_row = 2
    else:
        last_stats_row_cell = stats_sheet.range('A2').end('down')
        last_stats_row = last_stats_row_cell.row + 1
    stats_sheet.range('A' + str(last_stats_row)).options(index=False, header=False).value = df_stats

    # Write the points to the spreadsheet
    # Convert the data frame to a list so we can change the ncalfclub data type from int to string
    points_list = df_rpoints.values.tolist()

    # Convert the ncalfclub ID to a name
    for pointsrow in points_list:
        pointsrow[1] = convert_ncalfclub_code_to_name(pointsrow[1])

    if newroundno == 1:
        last_stats_row = 2
    else:
        last_stats_row_cell = stats_sheet.range('AE2').end('down')
        last_stats_row = last_stats_row_cell.row + 1
    stats_sheet.range('AE' + str(last_stats_row)).value = points_list

    wb.sheets('Ladder').range("F1").value = "ROUND " + str(newroundno)
    wb.sheets('Team Stats').range("F1").value = "ROUND " + str(newroundno)
    wb.sheets('Available').range("H1").value = "ROUND " + str(newroundno)
    wb.sheets('Teamlists').range("B12").value = "ROUND " + str(newroundno)

    wb.save()
    wb.close()
    excel_app.quit()


def output_results_to_old_excel(stats_list, rpoints_list, changes_list, season, roundno):
    # Opening the NCALF reports spreadsheet
    print('OUTPUT THE RESULTS TO THE OLD SPREADSHEET FORMAT')
    print('Opening Excel')
    excel_app = xw.App(visible=False)
    print('Opening the workbook')
    wb = excel_app.books.open(r"c:\ncalf\bin\NCALF Reports.xlsx")

    # Connect to the database
    conn = create_sql_connection()
    c = conn.cursor()

    # Get the season ladder information summary from the database
    points_sqlstring = """SELECT ncalfclub, sum(pki) as spki, sum(pm) as spm, sum(phb) as sphb, 
                sum(pf) as spf, sum(psc) as spsc, sum(pho) as spho, sum(pt) as spt, sum(hoarding) as sh, 
                sum(pki + pm + phb + pf + psc + pho + pt + hoarding) as ptotal 
            FROM points 
            WHERE Season = %s AND Round <= %s 
            GROUP BY ncalfclub 
            ORDER BY sum(pki + pm + phb + pf + psc + pho + pt + hoarding) DESC"""
    points_val = (season, roundno)
    c.execute(points_sqlstring, points_val)
    points = c.fetchall()

    # Get the reserve players from the database
    reserve_sqlstring = """SELECT Round, ncalfclub, posplyd, PlayerSeasonID FROM stats WHERE season = %s AND
                            Round <= %s AND ncalfclub > 100"""
    reserve_val = (season, roundno)
    c.execute(reserve_sqlstring, reserve_val)
    reserve_stats = c.fetchall()

    # Prepare the season ladder for output
    ladder_list = []
    for line in points:
        ladder_list.append([convert_ncalfclub_code_to_name(line['ncalfclub']), "", "", "", "", "", "",
                            float(line['ptotal']), "", "", "", float(line['spki']), "", float(line['spm']), "",
                            float(line['sphb']), "", float(line['spf']), "", float(line['spsc']), "",
                            float(line['spho']), "", float(line['spt']), "", float(line['sh'])])

    # Prepare the round points list for output
    rpoints = []
    team_rpoints = []
    for rpoints_row in rpoints_list:
        if rpoints_row['hoarding'] < 0:
            hoarding_value = rpoints_row['hoarding']
        else:
            hoarding_value = ""

        rpoints.append([convert_ncalfclub_code_to_name(int(rpoints_row['ncalfclub'])), "", "", "", "", "",
                        float(rpoints_row['ptotal']), hoarding_value, "",
                        rpoints_row['k'], "", rpoints_row['m'], "", rpoints_row['hb'], "", rpoints_row['ff'], "",
                        rpoints_row['fa'], "", rpoints_row['g'], "", rpoints_row['b'], "", rpoints_row['ho'], "",
                        rpoints_row['t']])

    print('Printing the team results to the spreadsheet')
    # Write the season ladder to the spreadsheet
    ncalfclubs = ['BU', 'BB1', 'BB2', 'BB3', 'GKR', 'JJ', 'KP', 'KC', 'LH', 'RR', 'SS']
    for ncalfclub in ncalfclubs:
        team_sheet = wb.sheets(ncalfclub)
        # Print the page headers
        team_sheet.range('B1').value = 'NCALF ' + str(season)
        team_sheet.range('F1').value = 'ROUND ' + str(roundno)

        # Print the season ladder
        team_sheet.range('A6').value = ladder_list
        team_sheet.range('H6:Z16').number_format = '0.0'
        team_sheet.range('H6:Z16').columns.autofit()

        # Print the round points list
        team_sheet.range('A20').value = rpoints
        team_sheet.range('G20:G30').number_format = '0.0'
        team_sheet.range('G20:G30').columns.autofit()

        # Print the points for the round for each team
        trpoints = [x for x in rpoints_list if x['ncalfclub'] == convert_ncalfclub_name_to_code(ncalfclub)][0]
        team_round_points = [float(trpoints['pki']), "", float(trpoints['pm']), "", float(trpoints['phb']), "", "",
                             float(trpoints['pf']), "", "", "", float(trpoints['psc']), "", "", float(trpoints['pho']),
                             "", float(trpoints['pt'])]

        team_sheet.range('J31').value = team_round_points

        # Prepare the stats information for output
        stats = []
        # Filter the stats list for the round and ncalfclub
        team_stats = [x for x in stats_list if x['ncalfclub'] == convert_ncalfclub_name_to_code(ncalfclub)]
        for stats_row in team_stats:
            stats.append([stats_row['Pos'], stats_row['ID'], stats_row['First_Name'] + ' ' + stats_row['Surname'], "",
                          "", stats_row['Club'], stats_row['k'], stats_row['m'], stats_row['hb'], stats_row['ff'],
                          stats_row['fa'], stats_row['g'], stats_row['b'], stats_row['ho'], stats_row['t'], "",
                          stats_row['gms'], stats_row['sk'], stats_row['sm'], stats_row['shb'], stats_row['sff'],
                          stats_row['sfa'], stats_row['sg'], stats_row['sb'], stats_row['sho'], stats_row['st']])

        # Print the stats for the team
        team_sheet.range('A35').value = stats

        stats = []
        # Filter the stats list for the round and ncalfclub players on reserve
        round_reserve_players = [x for x in stats_list if
                                 x['ncalfclub'] == (convert_ncalfclub_name_to_code(ncalfclub) + 100)]
        for round_reserve_player in round_reserve_players:
            # Get the number of weeks on reserve for each player
            num_weeks_on_reserve = len([x for x in reserve_stats if
                                        x['ncalfclub'] == (convert_ncalfclub_name_to_code(ncalfclub) + 100) and
                                        x['Round'] >= (roundno - 2) and
                                        x['posplyd'] > 0 and
                                        x['PlayerSeasonID'] == round_reserve_player['ID']])
            if num_weeks_on_reserve > 0:
                nwor_text = ' (' + str(num_weeks_on_reserve) + ')'
            else:
                nwor_text = ""
            stats.append([round_reserve_player['Pos'], round_reserve_player['ID'], round_reserve_player['First_Name'] +
                          ' ' + round_reserve_player['Surname'] + nwor_text, "", "", round_reserve_player['Club'],
                          round_reserve_player['k'], round_reserve_player['m'], round_reserve_player['hb'],
                          round_reserve_player['ff'], round_reserve_player['fa'], round_reserve_player['g'],
                          round_reserve_player['b'], round_reserve_player['ho'], round_reserve_player['t'], "",
                          round_reserve_player['gms'], round_reserve_player['sk'], round_reserve_player['sm'],
                          round_reserve_player['shb'], round_reserve_player['sff'], round_reserve_player['sfa'],
                          round_reserve_player['sg'], round_reserve_player['sb'], round_reserve_player['sho'],
                          round_reserve_player['st']])

        # Print the stats for the team
        team_sheet.range('A59').value = stats

    # Prepare and output the availability list
    print('Printing the availability list')
    team_sheet = wb.sheets('Avail')

    # Print the page headers
    team_sheet.range('A1').value = 'NCALF ' + str(season)
    team_sheet.range('F1').value = 'ROUND' + str(roundno)
    avail_stats_list = [x for x in stats_list if x['posplyd'] > 0 and x['ncalfclub'] == 0]
    avail_list = []
    for stats_row in avail_stats_list:
        avail_list.append([stats_row['Pos'], stats_row['ID'], stats_row['First_Name'] + ' ' + stats_row['Surname'], "",
                           "", stats_row['Club'], stats_row['k'], stats_row['m'], stats_row['hb'], stats_row['ff'],
                           stats_row['fa'], stats_row['g'], stats_row['b'], stats_row['ho'], stats_row['t'], "",
                           stats_row['gms'], stats_row['sk'], stats_row['sm'], stats_row['shb'], stats_row['sff'],
                           stats_row['sfa'], stats_row['sg'], stats_row['sb'], stats_row['sho'], stats_row['st']])

    # Print the availability list to the spreadsheet
    team_sheet.range('A5').value = avail_list

    # Print the teamsheets to the spreadsheet
    print('Printing the teamsheets')
    team_sheet = wb.sheets('Teamlist')

    # Print the page headers
    team_sheet.range('B4').value = 'NCALF ' + str(season)
    team_sheet.range('B12').value = 'ROUND ' + str(roundno)
    teamsheet_details = [[1, 'F2', 'F24'], [2, 'J2', 'J24'], [3, 'N2', 'N24'], [4, 'R2', 'R24'], [5, 'V2', 'V24'],
                         [6, 'B32', 'B54'], [7, 'F32', 'F54'], [8, 'J32', 'J54'], [9, 'N32', 'N54'], [10, 'R32', 'R54'],
                         [11, 'V32', 'V54']]

    for teamsheet in teamsheet_details:
        print_team = []
        print_reserves = []
        teamlist = [x for x in stats_list if (x['ncalfclub'] == teamsheet[0] or x['ncalfclub'] == (teamsheet[0] + 100))
                    and x['Round'] == roundno]
        for team in teamlist:
            if team['ncalfclub'] < 100:
                print_team.append([team['ID'], team['First_Name'], team['Surname'], team['Club']])
            else:
                print_reserves.append([team['ID'], team['First_Name'], team['Surname'], team['Club']])

        # Print the team list to the spreadsheet
        team_sheet.range(teamsheet[1]).value = print_team
        team_sheet.range(teamsheet[2]).value = print_reserves

    # Print changes for the round
    print('Printing the changes for the round')
    team_sheet = wb.sheets('Changes')

    # Print the page headers
    team_sheet.range('B1').value = 'NCALF ' + str(season)
    team_sheet.range('D1').value = 'ROUND ' + str(roundno)

    changes_output_list = []
    for changes_row in changes_list:
        if changes_row['outto'] == 'Pool':
            changetype = 'Sack'
        else:
            changetype = "Reserve"
        changes_output_list.append(
            [changes_row['pos'], convert_ncalfclub_code_to_name(changes_row['ncalfclub']), changes_row['plyrin'],
             changetype, changes_row['plyrout']])

    # Print the changes list to the spreadsheet
    team_sheet.range('B5').value = changes_output_list

    print('Save and close the spreadsheet')
    wb.save(r'C:\ncalf\NCALF ' + str(season) + r'\reports\Round ' + str(roundno) + ' results.xlsx')
    wb.close()
    excel_app.quit()


def output_results_data_to_excel_dump(season, roundno):
    # Outputs the database information to a dump for the NCALF results spreadsheet

    # Connect to the database
    conn = create_sql_connection()
    c = conn.cursor()

    # Start by opening the spreadsheet and selecting the main sheet
    fn = r'c:\ncalf\NCALF ' + str(season) + r'\NCALF ' + str(season) + r' results dump.xlsx'

    stats_list = []
    rpoints_list = []

    if roundno:
        # for round in range(1,roundno+1):
        # for round in range(1,3):
        # Set the sqlstring to get the data out for the round
        stats_sqlstring = """SELECT RS.Round, RS.PlayerSeasonID AS ID, RS.PlayerFirstname AS First_Name, 
                    RS.PlayerSurname AS Surname, RS.Position AS Pos, NULL, RS.Club,   
                    RS.k, RS.m, RS.hb, RS.ff, RS.fa, RS.g, RS.b, RS.ho, RS.t, NULL, 
                    SS.gms as gms, sum(SS.sk) AS sk, sum(SS.sm) AS sm, sum(SS.shb) AS shb, sum(SS.sff) AS sff, 
                    sum(SS.sfa) AS sfa, sum(SS.sg) AS sg, sum(SS.sb) AS sb, sum(SS.sho) AS sho, sum(SS.st) AS st, 
                    RS.ncalfclub, RS.posplyd 
                    FROM 
                    (SELECT S.Round, S.PlayerSeasonID, PL.PlayerFirstname, PL.PlayerSurname, S.Club, S.ncalfclub, 
                    S.Position, S.posplyd, S.k, S.m, S.hb, S.ff, S.fa, S.g, S.b, S.ho, S.t 
                        FROM players AS PL 
                        INNER JOIN 
                        stats S ON S.PlayerID = PL.PlayerID 
                        WHERE S.Season = %s AND S.Round = %s  
                            ) 
                        AS RS 
                INNER JOIN 
                (SELECT S.PlayerSeasonID, count(CASE WHEN S.posplyd > 0 THEN 1 ELSE NULL END) as gms, sum(S.k) AS sk, 
                sum(S.m) AS sm, sum(S.hb) AS shb, sum(S.ff) AS sff, sum(S.fa) AS sfa, sum(S.g) AS sg, sum(S.b) AS sb, 
                sum(S.ho) AS sho, sum(S.t) AS st 
                    FROM players AS PL 
                            INNER JOIN  
                            stats S ON S.PlayerID = PL.PlayerID 
                    WHERE S.Season = %s AND S.Round <= %s  
                    GROUP by PL.PlayerID) AS SS 
                ON RS.PlayerSeasonID = SS.PlayerSeasonID 
                GROUP BY RS.PlayerSeasonID 
                ORDER BY RS.ncalfclub, RS.Position"""
        stats_val = (season, roundno, season, roundno)

        rpoints_sqlstring = """SELECT P.round, P.ncalfclub, 
        (P.pki + P.pm + P.phb + P.pf + P.psc + P.pho + P.pt + P.hoarding) AS ptotal, P.hoarding, P.pki, P.pm, P.phb,
                P.pf, P.psc, P.pho, P.pt, S.k, S.m, S.hb, S.ff, S.fa, S.g, S.b, S.ho, S.t
            FROM points P
            INNER JOIN
                   (SELECT ncalfclub, sum(k) AS k, sum(m) AS m, sum(hb) AS hb, sum(ff) AS ff, sum(fa) AS fa, 
                   sum(g) AS g, sum(b) AS b, sum(ho) AS ho, sum(t) AS t
                    FROM stats
                    WHERE Season = %s AND Round = %s AND 
                        ncalfclub > 0
                    GROUP BY ncalfclub
                   ) S ON P.ncalfclub = S.ncalfclub
             WHERE P.Season = %s AND P.Round = %s  
             GROUP BY P.ncalfclub"""
        rpoints_val = (season, roundno, season, roundno)

        print("Round " + str(roundno))
        c.execute(stats_sqlstring, stats_val)
        stats = c.fetchall()
        c.execute(rpoints_sqlstring, rpoints_val)
        rpoints = c.fetchall()

        for statsrow in stats:
            stats_list.append(statsrow)
        for pointsrow in rpoints:
            rpoints_list.append(pointsrow)
            # Update the ncalfclub references to club names
            clubname = convert_ncalfclub_code_to_name(int(pointsrow['ncalfclub']))
            pointsrow['ncalfclub'] = clubname

    stats_dataframe = pd.DataFrame(stats_list)
    rpoints_dataframe = pd.DataFrame(rpoints_list)

    # Write the stats to the spreadsheet
    writer_stats = pd.ExcelWriter(fn, engine='openpyxl', mode='a')
    stats_dataframe.to_excel(excel_writer=writer_stats, sheet_name='Stats R' + str(roundno), header=None, index=False)
    writer_stats.close()

    # Write the points to the spreadsheet
    writer_points = pd.ExcelWriter(fn, engine='openpyxl', mode='a')
    rpoints_dataframe.to_excel(excel_writer=writer_points, sheet_name='Points R' + str(roundno), header=None,
                               index=False)
    writer_points.close()


def output_results_reports(season, roundno):
    # Outputs the database information directly to the NCALF results spreadsheet
    print('OUTPUTTING THE RESULTS TO THE SPREADSHEETS')

    # Connect to the database
    conn = create_sql_connection()
    c = conn.cursor()

    stats_list = []
    rpoints_list = []
    changes_list = []

    if roundno:
        # Set the sqlstring to get the stats out for the round
        stats_sqlstring = """SELECT RS.Round, RS.PlayerSeasonID AS ID, RS.PlayerFirstname AS First_Name, 
                    RS.PlayerSurname AS Surname, RS.Position AS Pos, NULL, RS.Club,   
                    RS.k, RS.m, RS.hb, RS.ff, RS.fa, RS.g, RS.b, RS.ho, RS.t, NULL, 
                    SS.gms as gms, 
                    CAST(sum(SS.sk) AS UNSIGNED) AS sk, CAST(sum(SS.sm) AS UNSIGNED) AS sm, 
                    CAST(sum(SS.shb) AS UNSIGNED) AS shb, CAST(sum(SS.sff) AS UNSIGNED) AS sff, 
                    CAST(sum(SS.sfa) AS UNSIGNED) AS sfa, CAST(sum(SS.sg) AS UNSIGNED) AS sg, 
                    CAST(sum(SS.sb) AS UNSIGNED) AS sb, CAST(sum(SS.sho) AS UNSIGNED) AS sho, 
                    CAST(sum(SS.st) AS UNSIGNED) AS st, RS.ncalfclub, RS.posplyd 
                    FROM 
                    (SELECT S.Round, S.PlayerSeasonID, PL.PlayerFirstname, PL.PlayerSurname, S.Club, S.ncalfclub, 
                    S.Position, S.posplyd, S.k, S.m, S.hb, S.ff, S.fa, S.g, S.b, S.ho, S.t 
                        FROM players AS PL 
                        INNER JOIN 
                        stats S ON S.PlayerID = PL.PlayerID 
                        WHERE S.Season = %s AND S.Round = %s  
                            ) 
                        AS RS 
                INNER JOIN 
                (SELECT S.PlayerSeasonID, count(CASE WHEN S.posplyd > 0 THEN 1 ELSE NULL END) as gms, sum(S.k) AS sk, 
                sum(S.m) AS sm, sum(S.hb) AS shb, sum(S.ff) AS sff, sum(S.fa) AS sfa, sum(S.g) AS sg, sum(S.b) AS sb, 
                sum(S.ho) AS sho, sum(S.t) AS st 
                    FROM players AS PL 
                            INNER JOIN  
                            stats S ON S.PlayerID = PL.PlayerID 
                    WHERE S.Season = %s AND S.Round <= %s  
                    GROUP by PL.PlayerID) AS SS 
                ON RS.PlayerSeasonID = SS.PlayerSeasonID 
                GROUP BY RS.PlayerSeasonID 
                ORDER BY RS.ncalfclub, RS.Position"""
        stats_val = (season, roundno, season, roundno)

        rpoints_sqlstring = """SELECT P.round, P.ncalfclub, 
                CAST((P.pki + P.pm + P.phb + P.pf + P.psc + P.pho + P.pt + P.hoarding) AS DECIMAL(3,1)) AS ptotal, 
                CAST(P.hoarding AS SIGNED) as hoarding, CAST(P.pki AS DECIMAL(3,1)) as pki, 
                CAST(P.pm AS DECIMAL(3,1)) as pm, CAST(P.phb AS DECIMAL(3,1)) as phb, 
                CAST(P.pf AS DECIMAL(3,1)) as pf, CAST(P.psc AS DECIMAL(3,1)) as psc, 
                CAST(P.pho AS DECIMAL(3,1)) as pho, CAST(P.pt AS DECIMAL(3,1)) as pt, 
                CAST(S.k AS UNSIGNED) as k, CAST(S.m AS UNSIGNED) as m, 
                CAST(S.hb AS UNSIGNED) as hb, CAST(S.ff AS UNSIGNED) as ff, CAST(S.fa AS UNSIGNED) as fa, 
                CAST(S.g AS UNSIGNED) as g, CAST(S.b AS UNSIGNED) as b, CAST(S.ho AS UNSIGNED) as ho, 
                CAST(S.t AS UNSIGNED) as t 
            FROM points P
            INNER JOIN
                   (SELECT ncalfclub, sum(k) AS k, sum(m) AS m, sum(hb) AS hb, sum(ff) AS ff, sum(fa) AS fa, 
                   sum(g) AS g, sum(b) AS b, sum(ho) AS ho, sum(t) AS t
                    FROM stats
                    WHERE Season = %s AND Round = %s AND 
                        ncalfclub > 0
                    GROUP BY ncalfclub
                   ) S ON P.ncalfclub = S.ncalfclub
             WHERE P.Season = %s AND P.Round = %s  
             GROUP BY P.ncalfclub
             ORDER BY ptotal DESC"""
        rpoints_val = (season, roundno, season, roundno)

        changes_sqlstring = """SELECT * FROM ncalfdb.changes WHERE season = %s AND round = %s ORDER BY ncalfclub"""
        changes_val = (season, roundno)

        print("Getting the stats and points data for Round " + str(roundno))
        c.execute(stats_sqlstring, stats_val)
        stats = c.fetchall()
        c.execute(rpoints_sqlstring, rpoints_val)
        rpoints = c.fetchall()
        c.execute(changes_sqlstring, changes_val)
        changes = c.fetchall()

        print('Build the lists to output to the spreadsheet')
        for statsrow in stats:
            stats_list.append(statsrow)
        for pointsrow in rpoints:
            rpoints_list.append(pointsrow)
        for change in changes:
            changes_list.append(change)

    stats_dataframe = pd.DataFrame(stats_list)
    rpoints_dataframe = pd.DataFrame(rpoints_list)

    # Write the results to the original results spreadsheet using the template
    output_results_to_old_excel(stats_list, rpoints_list, changes_list, season, roundno)
    # Write the results to the new results spreadsheet
    output_results_to_new_excel(stats_dataframe, rpoints_dataframe, roundno)


def update_changes_sheet(season, roundno):
    # Updates the changes sheet with the updated player list and positions
    print('UPDATING THE CHANGES SPREADSHEET')
    print('Opening Excel')
    excel_app = xw.App(visible=False)
    print('Opening the workbook')
    wb = excel_app.books.open(r'C:\Users\hener\Dropbox\NCALF\Changes\NCALF ' + str(season) + ' Changes.xlsx')

    print('Copy the previous round sheet')
    prev_changes_sheet = wb.sheets('Round ' + str(roundno))
    prev_changes_sheet.copy(after=prev_changes_sheet, name='Round ' + str(roundno + 1))
    new_changes_sheet = wb.sheets('Round ' + str(roundno + 1))
    print('Clear the changes input data for the new round')
    last_changes_row = new_changes_sheet.range('E4').end('down').row
    new_changes_sheet.range((4, 1), (last_changes_row, 3)).clear_contents()
    new_changes_sheet.range((4, 5), (last_changes_row, 6)).clear_contents()

    # Connect to the database
    conn = create_sql_connection()
    c = conn.cursor()

    output_list = []

    # Set the sqlstring to get the stats out for the round
    changes_sqlstring = """SELECT S.PlayerSeasonID, PL.PlayerFirstname, PL.PlayerSurname, S.Position, S.Club
                            FROM players AS PL
                            INNER JOIN
                            stats S ON S.PlayerID = PL.PlayerID
                            WHERE S.Season = %s AND S.Round = %s"""
    changes_val = (season, roundno)

    print("Updating the data sheet for Round " + str(roundno))
    c.execute(changes_sqlstring, changes_val)
    player_list = c.fetchall()

    for player in player_list:
        output_list.append(player)
    output_dataframe = pd.DataFrame(output_list)

    data_sheet = wb.sheets('Lists')
    # data_sheet.tables('Players').update(output_dataframe)
    data_sheet.range('D2').options(index=False).value = output_dataframe
    print('Save and close the spreadsheet')
    wb.save()
    wb.close()
    excel_app.quit()


def calculate_new_positions(season):
    # Calculate the new positions for the new season based on the positions played in the previous season

    # Connect to the database
    sql_conn = create_sql_connection()

    # test for connection being made
    if sql_conn is None:
        print("Error! cannot create the database connection.")
    else:
        # Set the SQL cursor
        sql_cursor = sql_conn.cursor()

        # Get the highest id for the players
        sqlcommand = "SELECT MAX(PlayerSeasonID) FROM stats WHERE Season = %s"
        val = (season - 1)
        sql_cursor.execute(sqlcommand, val)
        intHighID = int(sql_cursor.fetchone()['MAX(PlayerSeasonID)'])

        # Get the posplyd data for the season for each player
        sqlcommand = "SELECT Round, PlayerID, PlayerSeasonID, Position, posplyd FROM stats " \
                     "WHERE Season = %s order by PlayerSeasonID, Round"
        val = (season - 1)
        sql_cursor.execute(sqlcommand, val)
        playerdataset = sql_cursor.fetchall()

        # Set the position counter to zero
        positions = {'D': 0, 'C': 0, 'F': 0, 'RK': 0, 'OB': 0}

        CurrPlayerSeasonID = playerdataset[0]['PlayerSeasonID']
        CurrPlayerPosn = playerdataset[0]['Position']
        CurrPlayerID = playerdataset[0]['PlayerID']
        for playerdata in playerdataset:
            if playerdata['PlayerSeasonID'] == CurrPlayerSeasonID:
                CurrPlayerSeasonID = playerdata['PlayerSeasonID']
                CurrPlayerPosn = playerdata['Position']
                CurrPlayerID = playerdata['PlayerID']
                if playerdata['posplyd'] in range(1, 7):
                    positions['D'] += 1
                if playerdata['posplyd'] in range(7, 10):
                    positions['C'] += 1
                if playerdata['posplyd'] in range(10, 16):
                    positions['F'] += 1
                if playerdata['posplyd'] == 16:
                    positions['RK'] += 1
                if playerdata['posplyd'] in range(17, 19):
                    positions['OB'] += 1

                # If it is the last element in the table, then assign position
                if playerdataset.index(playerdata) == len(playerdataset) - 1:
                    # Get the max
                    maxcount = max(positions.values())
                    # Get all the positions with that max number
                    position_list = [k for k, v in positions.items() if v == maxcount]
                    # Assign the position
                    if CurrPlayerPosn in position_list:
                        newpos = CurrPlayerPosn
                    elif maxcount == 0:
                        newpos = "ROOK"
                    else:
                        newpos = "/".join(position_list)

                    # Save the new position in the offseason table
                    sqlcommand = "UPDATE offseason SET newpos = %s WHERE PlayerID = %s"
                    val = (newpos, CurrPlayerID)
                    sql_cursor.execute(sqlcommand, val)
                    sql_conn.commit()

            else:
                # Change of player
                # Assign the position of the player just completed

                # Get the max
                maxcount = max(positions.values())
                # Get all the positions with that max number
                position_list = [k for k, v in positions.items() if v == maxcount]
                # Assign the position
                if CurrPlayerPosn in position_list:
                    newpos = CurrPlayerPosn
                elif maxcount == 0:
                    newpos = "ROOK"
                else:
                    newpos = "/".join(position_list)

                # Save the new position in the offseason table
                sqlcommand = "UPDATE offseason SET newpos = %s WHERE PlayerID = %s"
                val = (newpos, CurrPlayerID)
                sql_cursor.execute(sqlcommand, val)
                sql_conn.commit()

                # Set up for the new player
                positions = {'D': 0, 'C': 0, 'F': 0, 'RK': 0, 'OB': 0}

                CurrPlayerSeasonID = playerdata['PlayerSeasonID']
                CurrPlayerPosn = playerdata['Position']
                CurrPlayerID = playerdata['PlayerID']
                if playerdata['posplyd'] in range(1, 7):
                    positions['D'] += 1
                if playerdata['posplyd'] in range(7, 10):
                    positions['C'] += 1
                if playerdata['posplyd'] in range(10, 16):
                    positions['F'] += 1
                if playerdata['posplyd'] == 16:
                    positions['RK'] += 1
                if playerdata['posplyd'] in range(17, 19):
                    positions['OB'] += 1

        # Set all the remaining players to 'ROOK'
        sqlcommand = "UPDATE offseason SET newpos = 'ROOK' where newpos = '0'"
        sql_cursor.execute(sqlcommand)
        sql_conn.commit()


def create_draft_stats_spreadsheet(season):
    # Creates a spreadsheet containing the players and stats for the new year

    # Connect to the database
    sql_conn = create_sql_connection()

    # test for connection being made
    if sql_conn is None:
        print("Error! cannot create the database connection.")
    else:
        # Set the SQL cursor
        sql_cursor = sql_conn.cursor()

        # Get the player and stats data for the season for each player
        sqlcommand = """SELECT OS.PlayerSeasonID, OS.PlayerFirstName, OS.PlayerSurname, OS.Club, OS.newpos as Pos, 
                        count(CASE WHEN S.posplyd > 0 THEN 1 ELSE NULL END) AS gms, sum(S.k) As sk, sum(S.m) As sm, 
                        sum(S.hb) As shb, sum(S.ff) As sff, sum(S.fa) As sfa, sum(S.g) As sg, sum(S.b) As sb, 
                        sum(S.ho) As sho, sum(S.t) As st 
                        FROM offseason AS OS 
                        LEFT JOIN 
                            (SELECT * FROM stats WHERE season = %s) S ON OS.PlayerID = S.PlayerID 
                        WHERE OS.Season = %s
                        GROUP BY OS.PlayerID 
                        ORDER BY OS.PlayerSeasonID"""
        val = (season - 1, season)

        # Read the data into a dataframe
        df = pd.read_sql(sqlcommand, sql_conn, params=val)
        df.to_excel(r'C:\ncalf\NCALF ' + str(season) + r'\NCALF ' + str(season) + r' Draft stats.xlsx')


def CreateDraftDatabase(season):
    # Creates the draft database for the new season from the offseason table

    # Connect to the database
    sql_conn = create_sql_connection()

    # test for connection being made
    if sql_conn is None:
        print("Error! cannot create the database connection.")
    else:
        # Set the SQL cursor
        sql_cursor = sql_conn.cursor()

        # Clear any entries in draft_players for the current season if they already exist
        sqlcommand = "DELETE FROM draft_players WHERE Season = %s"
        val = (season,)
        sql_cursor.execute(sqlcommand, val)
        sql_conn.commit()

        # 'Import the players list into the draft database from the offseason table
        # 'and initialise the table values
        sqlcommand = "SELECT * FROM offseason"
        sql_cursor.execute(sqlcommand)
        playerdataset = sql_cursor.fetchall()

        for player in playerdataset:

            # Split the positions up if there are multiple positions
            if '/' in player['newpos']:
                positions = player['newpos'].split('/')
                anotherpos = True
            else:
                positions = [player['newpos']]
                anotherpos = False

            for position in positions:
                sqlcommand = """INSERT INTO draft_players (Season, PlayerSeasonID, PlayerID, FirstName, Surname, 
                                                club, price, sold, ncalfclub, posn, nominated, otherpos, sequence)
                                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
                val = (player['Season'], player['PlayerSeasonID'], player['PlayerID'], player['PlayerFirstName'],
                       player['PlayerSurname'], convert_long_AFL_name_to_short_NCALF_format(player['Club']), 0, 0, 0,
                       position, 0, anotherpos, 0)
                sql_cursor.execute(sqlcommand, val)
            sql_conn.commit()

        # Initialise the positions table to all false
        sqlcommand = "UPDATE draft_positions SET soldyet = 0"
        sql_cursor.execute(sqlcommand)
        sql_conn.commit()

        print("Draft table updated")


def SetupSeasonStats(season):
    # Sets up the stats table for the new season with default records for each player in each round

    # Connect to the database
    sql_conn = create_sql_connection()

    # Set the SQL cursor
    sql_cursor = sql_conn.cursor()

    # Get the player and stats data for the season for each player

    # Process the draft_players table, and update the playerseason table based on the draft results
    sqlcommand = "SELECT * FROM draft_players WHERE Season = %s ORDER BY PlayerSeasonID"
    val = (season,)
    sql_cursor.execute(sqlcommand, val)
    draft_playerlist = sql_cursor.fetchall()

    # Delete all existing records for the current season if they exist
    sqlcommand = "DELETE FROM stats WHERE Season = %s"
    val = (season,)
    sql_cursor.execute(sqlcommand, val)

    # Get the highest playerID for the season
    sqlcommand = 'SELECT MAX(PlayerSeasonID) AS max_id FROM draft_players WHERE Season = %s'
    val = (season,)
    sql_cursor.execute(sqlcommand, val)
    max_PlayerSeasonID = sql_cursor.fetchall()[0]['max_id']

    for i in range(1, max_PlayerSeasonID + 1):
        print("Creating stats table for player id: " + str(i))
        PlayerSeasonID_list = [x for x in draft_playerlist if x['PlayerSeasonID'] == i]
        if len(PlayerSeasonID_list) == 1:
            position = PlayerSeasonID_list[0]['posn']  # Record the position for later processing
            ncalfclub = PlayerSeasonID_list[0]['ncalfclub']
        else:  # Iterate through the list to get the player position
            position = ""  # Reset the position string
            ncalfclub = 0
            for draft_player in PlayerSeasonID_list:
                # If player has been sold to an ncalf club, record the position and ncalfclub and exit the local loop
                if (draft_player['sold'] == 1) and (draft_player['ncalfclub'] > 0):
                    position = draft_player['posn']
                    ncalfclub = draft_player['ncalfclub']
                    break
                # Create a single string of the multiple positions
                if position == "":
                    position = draft_player['posn']
                else:
                    position = position + "/" + draft_player['posn']

        # Save the record to the stats table
        for roundno in range(1, Numrounds + 1):
            sqlcommand = """INSERT INTO stats (Season, Round, PlayerID, PlayerSeasonID, Club, Position, posplyd, 
                            ncalfclub, k, m, hb, ff, fa, g, b, ho, t) 
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
            if roundno == 1:
                draftclub = ncalfclub
            else:
                draftclub = 0

            val = (season, roundno, PlayerSeasonID_list[0]['PlayerID'], PlayerSeasonID_list[0]['PlayerSeasonID'],
                   PlayerSeasonID_list[0]['club'], position, 0, draftclub, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            sql_cursor.execute(sqlcommand, val)

    sql_conn.commit()


def SetupSeasonPoints(season):
    # Sets up the points table for the new season with default records for each round

    # Connect to the database
    sql_conn = create_sql_connection()

    # Set the SQL cursor
    sql_cursor = sql_conn.cursor()

    for roundno in range(1, Numrounds + 1):
        print("Creating points table for Round " + str(roundno))
        for ncalfclub in range(1, NumNCALFClubs + 1):
            sqlcommand = """INSERT INTO points (Season, round, ncalfclub, pki, pm, phb, pf, psc, pho, pt, hoarding)
                            VALUES (%s, %s, %s, 0, 0, 0, 0, 0, 0, 0, 0)"""
            val = (season, roundno, ncalfclub)
            sql_cursor.execute(sqlcommand, val)
    sql_conn.commit()


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


def get_all_current_players():
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
    connection = create_sql_connection()

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
    conn = create_sql_connection()
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


def convert_transition_to_players_table():
    # Moves the data from the transition table to the players table as new data

    # open the sql database

    sql_conn = create_sql_connection()
    # test for connection being made
    if sql_conn is None:
        print("Error! cannot create the database connection.")
    else:
        # Set the SQL cursor
        sql_cursor = sql_conn.cursor()

        # Get all current player records and order by year from the transition table
        sqlite_sqlcommand = "SELECT *  \
                                FROM transition"

        sql_cursor.execute(sqlite_sqlcommand)

        players = sql_cursor.fetchall()

        # for each player in players

        print("Moving player data from transition table to players table")

        for player in players:
            # Get the player details so that can find in update sql later
            # Data structure - 0: FirstName, 1: Surname, 2:Year, 3:Club, 4:Photo, 5:Status, 6:UniqueID
            FirstName = player[0]
            Surname = player[1]
            Player_ID = player[6]
            Player_image = player[4]

            sqlite_sqlcommand = """INSERT INTO 'Players'('PlayerID', 'PlayerFirstname', 'PlayerSurname', 'PlayerImage') 
                                        VALUES (?, ?, ?, ?)"""
            val = (Player_ID, FirstName, Surname, Player_image)
            sql_cursor.execute(sqlite_sqlcommand, val)

        sql_conn.commit()


def AFL_short_to_long(str_teamname):
    # Converts the NCALF short team name long format team name
    if str_teamname == 'Ade':
        return 'Adelaide Crows'
    elif str_teamname == 'Bris':
        return 'Brisbane'
    elif str_teamname == 'Carl':
        return 'Carlton'
    elif str_teamname == 'Coll':
        return 'Collingwood'
    elif str_teamname == 'Ess':
        return 'Essendon'
    elif str_teamname == 'Fitzroy Lions':
        return 'Brisbane'
    elif str_teamname == 'Fre':
        return 'Fremantle'
    elif str_teamname == 'Geel':
        return 'Geelong'
    elif str_teamname == 'GC':
        return 'Gold Coast Suns'
    elif str_teamname == 'GWS':
        return 'GWS Giants'
    elif str_teamname == 'Haw':
        return 'Hawthorn'
    elif str_teamname == 'Melb':
        return 'Melbourne'
    elif str_teamname == 'NM':
        return 'North Melbourne'
    elif str_teamname == 'PA':
        return 'Port Adelaide'
    elif str_teamname == 'Rich':
        return 'Richmond'
    elif str_teamname == 'StK':
        return 'St Kilda'
    elif str_teamname == 'Syd':
        return 'Sydney'
    elif str_teamname == 'WC':
        return 'West Coast'
    elif str_teamname == 'Foot':
        return 'Western Bulldogs'


def get_input(playerlist):
    # Gets the list of players and prints them for the user to select which one
    i = 1
    for player in playerlist:
        print(f"{i:<5}{player[0]:<20}{player[1]:<20}{player[2]:<10}{player[3]:<30}{player[5]:<10}{player[6]:<10}")
        i += 1
    return int(input("Select the index of the correct player:"))


# def create_table(conn, create_table_sql):
#     """ create a table from the create_table_sql statement
#     :param conn: Connection object
#     :param create_table_sql: a CREATE TABLE statement
#     :return:
#     """
#     try:
#         c = conn.cursor()
#         c.execute(create_table_sql)
#         conn.commit()
#     except Error as e:
#         print(e)
#

def get_Playermatch_input(playerlist, firstname, surname, club):
    # Gets the list of players and prints them for the user to select which one
    i = 1
    for players in playerlist:
        print(f"{i:<5}{players['Club']:<10}{players['PlayerID']:<10}{firstname:<20}{surname:<20}{club:<20}")
        i += 1
    return int(input("Select the index of the correct player:"))


def merge_playerseason_into_stats():
    # Transfers all the playerseason information into the stats table to make the playerseason table
    # Connect to the database
    sql_conn = create_sql_connection()

    # test for connection being made
    if sql_conn is None:
        print("Error! cannot create the database connection.")
    else:
        # Set the SQL cursor
        sql_cursor = sql_conn.cursor()

        # For each line in playerseason
        sqlcommand = "select * from playerseason"
        sql_cursor.execute(sqlcommand)

        playerseasonlines = sql_cursor.fetchall()

        for playerseasonline in playerseasonlines:
            sqlcommand = "UPDATE stats SET PlayerID = %s, Club = %s, Position = %s WHERE Season = %s " \
                         "  AND PlayerSeasonID = %s"
            val = (playerseasonline['PlayerID'], playerseasonline['Club'], playerseasonline['Position'],
                   playerseasonline['Season'], playerseasonline['PlayerSeasonID'])
            sql_cursor.execute(sqlcommand, val)

        sql_conn.commit()


def update_player_list_to_offseason():
    # Adds new players to Player List from playerdownload and creates offseason table

    # Connect to the database
    sql_conn = create_sql_connection()

    # test for connection being made
    if sql_conn is None:
        print("Error! cannot create the database connection.")
    else:
        # Set the SQL cursor
        sql_cursor = sql_conn.cursor()
        match_cursor = sql_conn.cursor()

        # Get all player records from the PlayerDownload table
        sqlite_sqlcommand = "SELECT * FROM playerdownload ORDER BY Club, PlayerSurname ASC"

        sql_cursor.execute(sqlite_sqlcommand)

        players = sql_cursor.fetchall()

        # Set the PlayerSeasonID counter
        playerseasonID = 1

        # Set the new season PlayerID counter
        newPlayerID = 23000  # New season PlayerID holder

        # for each player in players
        player_cursor = sql_conn.cursor()
        for player in players:
            # Find the PlayerID for existing players, and create new playerIDs for new players
            # Need to add PlayerID default

            # Find the player in the Players table
            sqlite_sqlcommand = """SELECT Players.PlayerID 
                                    FROM Players 
                                    WHERE (Players.PlayerFirstname = %s OR Players.AltPlayerFirstName = %s) 
                                    AND (Players.PlayerSurname = %s OR Players.AltPlayerSurname = %s)"""
            val = (
                player['PlayerFirstName'], player['PlayerFirstName'], player['PlayerSurname'], player['PlayerSurname'])
            player_cursor.execute(sqlite_sqlcommand, val)
            matchplayers = player_cursor.fetchall()  # List of all players that match firstname and surname
            nummatchplayers = len(matchplayers)  # Number of players matched

            if nummatchplayers == 0:  # If player not found ie. new player
                # create new player id
                sqlite_sqlcommand = "SELECT MAX(PlayerID) FROM players"
                match_cursor.execute(sqlite_sqlcommand)
                currmaxPlayerID = int(match_cursor.fetchone()['MAX(PlayerID)'])
                if currmaxPlayerID < newPlayerID:  # If there haven't been any new season players added to the database
                    PlayerID = newPlayerID
                else:
                    PlayerID = currmaxPlayerID + 1
                # Add player to Player list
                sqlite_sqlcommand = """INSERT INTO Players (PlayerID, PlayerFirstName, PlayerSurname, PlayerImage) 
                                    VALUES (%s, %s, %s, %s)"""
                val = (str(PlayerID), player['PlayerFirstName'], player['PlayerSurname'], player['Photo'])
                match_cursor.execute(sqlite_sqlcommand, val)
            elif nummatchplayers == 1:
                # Record the playerID if single match found
                PlayerID = matchplayers[0]['PlayerID']
                # Update player image in Players table
                sqlite_sqlcommand = """UPDATE Players SET PlayerImage = %s WHERE PlayerID = %s"""
                val = (player['Photo'], str(PlayerID))
                match_cursor.execute(sqlite_sqlcommand, val)
            elif nummatchplayers > 1:
                # List the players with club and ask to choose which one
                sqlite_sqlcommand = """SELECT P.PlayerID, S.Club
                                    FROM Players P
                                    INNER JOIN stats S ON P.PlayerID = S.PlayerID 
                                    WHERE P.PlayerFirstName = %s AND P.PlayerSurname = %s AND S.Season = %s 
                                    AND S.Round = 1"""
                val = (player['PlayerFirstName'], player['PlayerSurname'], season - 1)
                player_cursor.execute(sqlite_sqlcommand, val)
                duplicateplayers = player_cursor.fetchall()
                Playerindex = get_Playermatch_input(duplicateplayers, player['PlayerFirstName'],
                                                    player['PlayerSurname'], player['Club'])
                PlayerID = duplicateplayers[Playerindex - 1]['PlayerID']
                # Update player image in Players table
                sqlite_sqlcommand = """UPDATE Players SET PlayerImage = %s WHERE PlayerID = %s"""
                val = (player['Photo'], str(PlayerID))
                match_cursor.execute(sqlite_sqlcommand, val)

            strPlayerID = str(PlayerID).rjust(5, '0')
            sqlite_sqlcommand = ("INSERT INTO offseason (Season, PlayerID, PlayerSeasonID, PlayerFirstName, \n"
                                 " PlayerSurname, Club, newpos) VALUES (%s, %s, %s, %s, %s, %s, %s)")
            val = (
                season, strPlayerID, playerseasonID, player['PlayerFirstName'], player['PlayerSurname'],
                player['Club'], 0)
            player_cursor.execute(sqlite_sqlcommand, val)

            # Increment the playerseasonID
            playerseasonID += 1

            sql_conn.commit()
