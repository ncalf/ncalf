from sqlite3 import Error
import pyodbc
import sys
import pymysql
import NCALF_bin

season = 2023


def convert_transition_to_players_table():
    # Moves the data from the transition table to the players table as new data

    # open the sql database

    sql_conn = NCALF_bin.create_sql_connection()
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


def create_table(conn, create_table_sql):
    """ create a table from the create_table_sql statement
    :param conn: Connection object
    :param create_table_sql: a CREATE TABLE statement
    :return:
    """
    try:
        c = conn.cursor()
        c.execute(create_table_sql)
        conn.commit()
    except Error as e:
        print(e)


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
    sql_conn = NCALF_bin.create_sql_connection()

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
    sql_conn = NCALF_bin.create_sql_connection()

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


