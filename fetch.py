import json
import urllib.request

import pandas as pd


def generate_excel(df):
    # Add table formatting, inspired by
    # https://stackoverflow.com/a/57204530

    # initialize ExcelWriter and set df as output
    writer = pd.ExcelWriter(
        "/mnt/c/Users/jacob/Documents/Financials/brothers_pullam.xlsx",
        engine="xlsxwriter",
    )
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


def fetch_data():
    DATA_URL = r"https://picks.cbssports.com/graphql?operationName=PoolSeasonStandingsQuery&variables={%22skipAncestorPools%22:false,%22skipPeriodPoints%22:false,%22skipCheckForIncompleteEntries%22:true,%22gameInstanceUid%22:%22cbs-ncaab-tournament-manager%22,%22includedEntryIds%22:[%22ivxhi4tzhiytinrsg4zdemry%22,%22ivxhi4tzhiytmnbugqydonzs%22],%22poolId%22:%22kbxw63b2he2deobzhaza====%22,%22first%22:500,%22orderBy%22:%22CHAMPION%22,%22sortingOrder%22:%22DESC%22}&extensions={%22persistedQuery%22:{%22version%22:1,%22sha256Hash%22:%22797a9386ad10d089d4d493a911bce5a63dae4efb4c02a0a32a40b20de37e002d%22}}&sort=DESC&orderBy=championTeam"
    # content = urllib.request.urlopen(DATA_URL)
    # data = (
    #     json.loads(content.read().decode()).get("data").get("gameInstance").get("pool")
    # )
    # NOTE: CBS added authentication to this URL endpoint, so I'm just copying and pasting
    # the data from the JSON response of DATA_URL in my browser to the "data.json" file
    data = open("./data.json", "r")
    data = (
        json.loads(data.read()).get("data").get("gameInstance").get("pool")
    )

    total_brackets = data.get("entriesCount")

    picked_brackets = data.get("entriesWithPicksCount")

    brackets = [
        x.get("node")
        for x in data.get("entries").get("edges")
        if x.get("node").get("picksCount") == 63
    ]
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


if __name__ == "__main__":
    # running this directly, just run fetch
    fetch_data()
