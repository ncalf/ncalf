import json

with open('stats.json', 'r') as json_file:
	json_data = json.load(json_file)
	# print(json_data)
	l = len(json_data["lists"])

	print("gamesPlayed = ", json_data["lists"][0]["stats"]["gamesPlayed"])

	print("Items: ", l)
