import urllib.request, json
import pandas as pd


def generate_excel(df):
    # Add table formatting, inspired by
    # https://stackoverflow.com/a/57204530

    # initialize ExcelWriter and set df as output
    writer = pd.ExcelWriter("brothers_pullam.xlsx", engine="xlsxwriter")
    df.to_excel(writer, sheet_name="Brackets", index=False)

    # worksheet is an instance of Excel sheet "Cities" - used for inserting the table
    worksheet = writer.sheets["Brackets"]
    # workbook is an instance of the whole book - used i.e. for cell format assignment
    workbook = writer.book

    header_cell_format = workbook.add_format()
    # header_cell_format.set_rotation(90)
    header_cell_format.set_align("center")
    header_cell_format.set_align("vcenter")

    # create list of dicts for header names
    #  (columns property accepts {'header': value} as header name)
    col_names = [{"header": col_name} for col_name in df.columns]

    # add table with coordinates: first row, first col, last row, last col;
    #  header names or formatting can be inserted into dict
    worksheet.add_table(
        0,
        0,
        df.shape[0],
        df.shape[1] - 1,
        {
            "columns": col_names,
            # 'style' = option Format as table value and is case sensitive
            # (look at the exact name into Excel)
            # "style": "Table Style Medium 10",
        },
    )

    # autoscale cols
    # https://stackoverflow.com/a/40535454
    # skip the loop completly if AutoFit for header is not needed
    for i, col in enumerate(col_names):
        # apply header_cell_format to cell on [row:0, column:i] and write text value from col_names in
        # worksheet.write(0, i, col["header"], header_cell_format)

        for idx, col in enumerate(df):  # loop through all columns
            series = df[col]
            max_len = (
                max(
                    (
                        series.astype(str).map(len).max(),  # len of largest item
                        len(str(series.name)),  # len of column name/header
                    )
                )
                + 5
            )  # adding a little extra space
            worksheet.set_column(idx, idx, max_len)  # set column width

    # save writer object and created Excel file with data from DataFrame
    writer.save()


# DATA_URL = r"https://picks.cbssports.com/graphql?operationName=PoolSeasonStandingsQuery&variables=%7B%22skipAncestorPools%22%3Afalse%2C%22skipPeriodPoints%22%3Afalse%2C%22gameInstanceUid%22%3A%22cbs-ncaab-tournament-manager%22%2C%22includedEntryIds%22%3A%5B%22ivxhi4tzhi4tqnjtgezdgni%3D%22%5D%2C%22poolId%22%3A%22kbxw63b2gy2dgnzyguza%3D%3D%3D%3D%22%2C%22first%22%3A50%7D&extensions=%7B%22persistedQuery%22%3A%7B%22version%22%3A1%2C%22sha256Hash%22%3A%224084de9ddd4d6369c4b2625dc756a3d6974b774e746960d3cdfcf64686147d7b%22%7D%7D"
DATA_URL = r'https://picks.cbssports.com/graphql?operationName=PoolSeasonStandingsQuery&variables={"skipAncestorPools":false,"skipPeriodPoints":false,"gameInstanceUid":"cbs-ncaab-tournament-manager","includedEntryIds":["ivxhi4tzhi4tqnjtgezdgni="],"poolId":"kbxw63b2gy2dgnzyguza====","first":1000}&extensions={"persistedQuery":{"version":1,"sha256Hash":"4084de9ddd4d6369c4b2625dc756a3d6974b774e746960d3cdfcf64686147d7b"}}'

content = urllib.request.urlopen(DATA_URL)

data = json.loads(content.read().decode()).get("data").get("gameInstance").get("pool")

total_brackets = data.get("entriesCount")

picked_brackets = data.get("entriesWithPicksCount")

brackets = [x.get("node") for x in data.get("entries").get("edges")]
for bracket in brackets:
    if bracket.get("championTeam"):
        bracket["championTeam"] = bracket["championTeam"].get("abbrev")

df = pd.DataFrame(brackets)[
    [
        "name",
        "poolRank",
        "championTeam",
        "fantasyPoints",
        "maxPoints",
        "correctPicks",
        "url",
    ]
]

generate_excel(df)
# df.to_excel("report.xlsx")
