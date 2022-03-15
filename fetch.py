import urllib.request, json
import pandas as pd

DATA_URL = r"https://picks.cbssports.com/graphql?operationName=PoolSeasonStandingsQuery&variables=%7B%22skipAncestorPools%22%3Afalse%2C%22skipPeriodPoints%22%3Afalse%2C%22gameInstanceUid%22%3A%22cbs-ncaab-tournament-manager%22%2C%22includedEntryIds%22%3A%5B%22ivxhi4tzhi4tqnjtgezdgni%3D%22%5D%2C%22poolId%22%3A%22kbxw63b2gy2dgnzyguza%3D%3D%3D%3D%22%2C%22first%22%3A50%7D&extensions=%7B%22persistedQuery%22%3A%7B%22version%22%3A1%2C%22sha256Hash%22%3A%224084de9ddd4d6369c4b2625dc756a3d6974b774e746960d3cdfcf64686147d7b%22%7D%7D"

content = urllib.request.urlopen(DATA_URL)

data = json.loads(content.read().decode()).get("data").get("gameInstance").get("pool")

total_brackets = data.get("entriesCount")

picked_brackets = data.get("entriesWithPicksCount")

brackets = [x.get("node") for x in data.get("entries").get("edges")]

df = pd.DataFrame(brackets)

df.to_excel("report.xlsx")
