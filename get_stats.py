# Read AFL Stats
# Install playwright from command line:
# 	pip install pytest-playwright
# Install browsers:
#  	playwright install

from playwright.sync_api import Page, expect, Playwright, sync_playwright

filter = "api.afl.com.au/cfs/afl/statsCentre" # /players?competitionId=CD_S2023014&teamIds=CD_T120%2CCD_T30"

statNum = 1

def onResponse (r):
	if filter in r.url:
		global statNum # We want to use the previously defined value

		print("Response url: " + r.url)
		# print("Response" + r)
		# print("Response" + r.ContentType)

		json = r.body().decode("utf-8")
		# print("JSON: " + json)

		file = "stats_" + str(statNum) + ".json"

		textfile = open(file, "w")
		a = textfile.write(json)
		textfile.close()
		print("File saved: ", file)

		statNum = statNum + 1


def run(playwright: Playwright) -> None:
	browser = playwright.chromium.launch(channel="chrome")
	context = browser.new_context()
	page = context.new_page()

	page.on("response", onResponse)
	# page = Page()
	print("Opening page...")

	url = r"https://www.afl.com.au/afl/matches/4787#player-stats"
	page.goto(url)
	
	context.close()
	browser.close()

with sync_playwright() as playwright:
	print('Starting...')
	run(playwright)
