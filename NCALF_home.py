import NCALF_bin
# import ncalf_gui
# import Get_current_players
# import Update_database


def main():
    # NCALF SEASON OPERATIONS

    season = 2023

    # START OF SEASON
    # Download all new season players and their photos from the AFL website (code needs update for each season)
    # Get_current_players.get_all_current_players()

    # Transfer the player information to the offseason table and assign playerIDs to new players and update photos
    # Update_database.update_player_list_to_offseason()

    # Calculate new positions and update the offseason table
    # NCALF_bin.calculate_new_positions(season)

    # Export the previous season stats to an Excel file to send out to all team owners
    # NCALF_bin.create_draft_stats_spreadsheet(season)

    # Create the draft database for the new season from the offseason table
    # NCALF_bin.CreateDraftDatabase(season)

    # After the draft, set up the stats table and the points table for the new season
    # NCALF_bin.SetupSeasonStats(season)
    # NCALF_bin.SetupSeasonPoints(season)

    # OPERATING THE SEASON
    # chromedriver.exe location: Stored in chromedriverfile variable within NCALF_bin.py

    roundno = 13

    # ncalf_gui.open_window()
    NCALF_bin.import_changes(season, roundno)
    NCALF_bin.Download_player_stats_AFL(season, roundno)
    NCALF_bin.calculate_results(season, roundno)
    NCALF_bin.output_results_reports(season, roundno)
    NCALF_bin.update_changes_sheet(season, roundno)
    print("done")

    # OLD STUFF
    # Get_current_players.check_current_players(database)
    # Get_past_players.main(database)
    # Update_database.assign_unique_ids(database)
    # Update_database.convert_transition_to_players_table(database)
    # Update_database.bring_past_players_over(database)
    # Bring current players over(database)
    # Update_database.merge_stats(database)
    # Update_database.build_stats_season(database)
    # Update_database.merge_points(database)
    # Update_database.update_past_player_ids(database)
    # Update_database.update_player_ids(database)
    # Update_database.update_draft_playerIDs(database)
    # Update_database.transition_to_playerdownload(database)
    # Update_database.update_player_list(database)
    # Update_database.merge_playerseason_into_stats()

    # NCALF_bin.Import_AFL_stats_positions_to_database(database)
    # NCALF_bin.Download_player_stats_footywire(database) - not required


if __name__ == "__main__":
    main()
