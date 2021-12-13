import gspread
import xlrd
from oauth2client.service_account import *
from poll_data import *

"""
Handle contacting Google Sheets and getting information from the document

@author: apkick
"""

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('FCFBRollCallBot-2d263a255851.json', scope)
gc = gspread.authorize(credentials)

sh1 = gc.open_by_url('https://docs.google.com/spreadsheets/d/1-8-X9arHYd4r_GlTjmsjVACzxyP9fcHnWqYE1LPrcYA/edit#gid=0')
standingsWorksheet = sh1.worksheet("Standings")
rankingsWorksheet = sh1.worksheet("Rankings")
compositeWorksheet = sh1.worksheet("Composite")

sh5 = gc.open_by_url(
    'https://docs.google.com/spreadsheets/d/1IrBBMKApJVYlU10wCOKp_oW3wvQfFT-xTC_A6EHJlzU/edit?usp=sharing')
fcsStandingsWorksheet = sh5.worksheet("Sheet1")

"""
Get the ACC standings from Google Docs

"""


def parse_acc():
    standings = ("----------------------\n**ACC**\n----------------------\n" +
                 "----------------------\nAtlantic\n----------------------\n")
    team_column = standingsWorksheet.col_values(2)
    team_conference_column = standingsWorksheet.col_values(3)
    team_overall_column = standingsWorksheet.col_values(4)
    for i in range(6, 13):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nCoastal\n----------------------\n"
    team_column = standingsWorksheet.col_values(5)
    team_conference_column = standingsWorksheet.col_values(6)
    team_overall_column = standingsWorksheet.col_values(7)
    for i in range(6, 13):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the AAC standings from Google Docs

"""


def parse_aac():
    standings = "----------------------\n**American**\n----------------------\n----------------------\nEast\n----------------------\n"
    team_column = standingsWorksheet.col_values(9)
    team_conference_column = standingsWorksheet.col_values(10)
    team_overall_column = standingsWorksheet.col_values(11)
    for i in range(6, 12):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nWest\n----------------------\n"
    team_column = standingsWorksheet.col_values(12)
    team_conference_column = standingsWorksheet.col_values(13)
    team_overall_column = standingsWorksheet.col_values(14)
    for i in range(6, 12):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Big 10 standings from Google Docs

"""


def parse_bigten():
    standings = ("----------------------\n**Big Ten**\n----------------------\n" +
                 "----------------------\nEast\n----------------------\n")
    team_column = standingsWorksheet.col_values(16)
    team_conference_column = standingsWorksheet.col_values(17)
    team_overall_column = standingsWorksheet.col_values(18)
    for i in range(6, 13):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nWest\n----------------------\n"
    team_column = standingsWorksheet.col_values(19)
    team_conference_column = standingsWorksheet.col_values(20)
    team_overall_column = standingsWorksheet.col_values(21)
    for i in range(6, 13):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Pac 12 standings from Google Docs

"""


def parse_pac12():
    standings = ("----------------------\n**Pac-12**\n----------------------\n" +
                 "----------------------\nNorth\n----------------------\n")
    team_column = standingsWorksheet.col_values(2)
    team_conference_column = standingsWorksheet.col_values(3)
    team_overall_column = standingsWorksheet.col_values(4)
    for i in range(28, 34):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nSouth\n----------------------\n"
    team_column = standingsWorksheet.col_values(5)
    team_conference_column = standingsWorksheet.col_values(6)
    team_overall_column = standingsWorksheet.col_values(7)
    for i in range(28, 34):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the CUSA standings from Google Docs

"""


def parse_cusa():
    standings = ("----------------------\n**Conference USA**\n----------------------\n" +
                 "----------------------\nEast\n----------------------\n")
    team_column = standingsWorksheet.col_values(2)
    team_conference_column = standingsWorksheet.col_values(3)
    team_overall_column = standingsWorksheet.col_values(4)
    for i in range(17, 24):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nWest\n----------------------\n"
    team_column = standingsWorksheet.col_values(5)
    team_conference_column = standingsWorksheet.col_values(6)
    team_overall_column = standingsWorksheet.col_values(7)
    for i in range(17, 24):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the MAC standings from Google Docs

"""


def parse_mac():
    standings = ("----------------------\n**MAC**\n----------------------\n" +
                 "----------------------\nEast\n----------------------\n")
    team_column = standingsWorksheet.col_values(9)
    team_conference_column = standingsWorksheet.col_values(10)
    team_overall_column = standingsWorksheet.col_values(11)
    for i in range(17, 23):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nWest\n----------------------\n"
    team_column = standingsWorksheet.col_values(12)
    team_conference_column = standingsWorksheet.col_values(13)
    team_overall_column = standingsWorksheet.col_values(14)
    for i in range(17, 23):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Mountain West standings from Google Docs

"""


def parse_mwc():
    standings = ("----------------------\n**Mountain West**\n----------------------\n" +
                 "----------------------\nMountain\n----------------------\n")
    team_column = standingsWorksheet.col_values(16)
    team_conference_column = standingsWorksheet.col_values(17)
    team_overall_column = standingsWorksheet.col_values(18)
    for i in range(17, 23):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nWest\n----------------------\n"
    team_column = standingsWorksheet.col_values(19)
    team_conference_column = standingsWorksheet.col_values(20)
    team_overall_column = standingsWorksheet.col_values(21)
    for i in range(17, 23):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the SEC standings from Google Docs

"""


def parse_sec():
    standings = ("----------------------\n**SEC**\n----------------------\n" +
                 "----------------------\nEast\n----------------------\n")
    team_column = standingsWorksheet.col_values(9)
    team_conference_column = standingsWorksheet.col_values(10)
    team_overall_column = standingsWorksheet.col_values(11)
    for i in range(28, 35):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nWest\n----------------------\n"
    team_column = standingsWorksheet.col_values(12)
    team_conference_column = standingsWorksheet.col_values(13)
    team_overall_column = standingsWorksheet.col_values(14)
    for i in range(28, 35):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Sun Belt standings from Google Docs

"""


def parse_sbc():
    standings = ("----------------------\n**Sun Belt**\n----------------------\n" +
                 "----------------------\nEast\n----------------------\n")
    team_column = standingsWorksheet.col_values(16)
    team_conference_column = standingsWorksheet.col_values(17)
    team_overall_column = standingsWorksheet.col_values(18)
    for i in range(28, 34):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nWest\n----------------------\n"
    team_column = standingsWorksheet.col_values(19)
    team_conference_column = standingsWorksheet.col_values(20)
    team_overall_column = standingsWorksheet.col_values(21)
    for i in range(28, 34):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Big 12 standings from Google Docs

"""


def parse_big12():
    standings = "----------------------\n**Big 12**\n----------------------\n"
    team_column = standingsWorksheet.col_values(2)
    team_conference_column = standingsWorksheet.col_values(3)
    team_overall_column = standingsWorksheet.col_values(4)
    for i in range(38, 43):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    team_column = standingsWorksheet.col_values(5)
    team_conference_column = standingsWorksheet.col_values(6)
    team_overall_column = standingsWorksheet.col_values(7)
    for i in range(38, 43):
        team = team_column[i].strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Independents standings from Google Docs

"""


def parse_independents():
    standings = "----------------------\n**Independents**\n----------------------\n"
    team_column = standingsWorksheet.col_values(9)
    team_overall_column = standingsWorksheet.col_values(11)
    for i in range(38, 42):
        team = team_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + "\n"
        i += 1
    return standings


"""
Get the America East standings from 1212.one

"""


def parse_americaeast():
    standings = ("----------------------\n**America East**\n----------------------\n" +
                 "----------------------\nTri-State\n----------------------\n")
    team_column = fcsStandingsWorksheet.col_values(2)
    team_conference_column = fcsStandingsWorksheet.col_values(3)
    team_overall_column = fcsStandingsWorksheet.col_values(9)
    for i in range(3, 9):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nNew England\n----------------------\n"
    for i in range(3, 9):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Atlantic Sun standings from 1212.one

"""


def parse_atlanticsun():
    standings = ("----------------------\n**Atlantic Sun**\n----------------------\n" +
                 "----------------------\nDusk\n----------------------\n")
    team_column = fcsStandingsWorksheet.col_values(2)
    team_conference_column = fcsStandingsWorksheet.col_values(3)
    team_overall_column = fcsStandingsWorksheet.col_values(9)
    for i in range(20, 27):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nDawn\n----------------------\n"
    for i in range(28, 35):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Big Sky standings from 1212.one

"""


def parse_bigsky():
    standings = ("----------------------\n**Big Sky**\n----------------------\n" +
                 "----------------------\nSouth\n----------------------\n")
    team_column = fcsStandingsWorksheet.col_values(2)
    team_conference_column = fcsStandingsWorksheet.col_values(3)
    team_overall_column = fcsStandingsWorksheet.col_values(9)
    for i in range(39, 46):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nNorth\n----------------------\n"
    for i in range(47, 54):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Carolina Football Conference standings from 1212.one

"""


def parse_cfc():
    standings = (
                "--------------------------------------------\n**Carolina Football Conference**\n--------------------" +
                "------------------------\n" +
                "----------------------\nNorth\n----------------------\n")
    team_column = fcsStandingsWorksheet.col_values(2)
    team_conference_column = fcsStandingsWorksheet.col_values(3)
    team_overall_column = fcsStandingsWorksheet.col_values(9)
    for i in range(58, 64):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nSouth\n----------------------\n"
    for i in range(65, 71):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Colonial standings from 1212.one

"""


def parse_colonial():
    standings = ("----------------------\n**Colonial**\n----------------------\n" +
                 "----------------------\nSouth\n----------------------\n")
    team_column = fcsStandingsWorksheet.col_values(2)
    team_conference_column = fcsStandingsWorksheet.col_values(3)
    team_overall_column = fcsStandingsWorksheet.col_values(9)
    for i in range(75, 81):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nNorth\n----------------------\n"
    for i in range(82, 88):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Delta Intercollegiate standings from 1212.one

"""


def parse_delta():
    standings = ("--------------------------------------------\n**Delta Intercollegiate**\n-------------------" +
                 "-------------------------\n" +
                 "----------------------\nMississippi Valley\n----------------------\n")
    team_column = fcsStandingsWorksheet.col_values(2)
    team_conference_column = fcsStandingsWorksheet.col_values(3)
    team_overall_column = fcsStandingsWorksheet.col_values(9)
    for i in range(92, 100):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nTennessee Valley\n----------------------\n"
    for i in range(82, 88):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Ivy League standings from 1212.one

"""


def parse_ivy():
    standings = "----------------------\n**Ivy League**\n----------------------\n"
    team_column = fcsStandingsWorksheet.col_values(2)
    team_conference_column = fcsStandingsWorksheet.col_values(3)
    team_overall_column = fcsStandingsWorksheet.col_values(6)
    for i in range(113, 121):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Mid-Atlantic standings from 1212.one

"""


def parse_midatlantic():
    standings = ("----------------------\n**Mid Atlantic**\n----------------------\n" +
                 "----------------------\nAtlantic\n----------------------\n")
    team_column = fcsStandingsWorksheet.col_values(2)
    team_conference_column = fcsStandingsWorksheet.col_values(3)
    team_overall_column = fcsStandingsWorksheet.col_values(9)
    for i in range(125, 131):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nAdirondack\n----------------------\n"
    for i in range(132, 138):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Missouri Valley standings from 1212.one

"""


def parse_mvc():
    standings = (
                "--------------------------------------------\n**Missouri Valley**\n-----------------------------------------\n" + \
                "----------------------\nPrairie\n----------------------\n")
    team_column = fcsStandingsWorksheet.col_values(2)
    team_conference_column = fcsStandingsWorksheet.col_values(3)
    team_overall_column = fcsStandingsWorksheet.col_values(9)
    for i in range(142, 149):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    standings = standings + "\n----------------------\nMetro\n----------------------\n"
    for i in range(150, 157):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the Southland standings from 1212.one

"""


def parse_southland():
    standings = "----------------------\n**Southland**\n----------------------\n"
    team_column = fcsStandingsWorksheet.col_values(2)
    team_conference_column = fcsStandingsWorksheet.col_values(3)
    team_overall_column = fcsStandingsWorksheet.col_values(6)
    for i in range(161, 175):
        team = team_column[i].split(" ")[:-1]
        team = ' '.join(team).strip()
        conference_record = team_conference_column[i].strip()
        overall_record = team_overall_column[i].strip()
        standings = standings + " " + team + " " + overall_record + " (" + conference_record + ")\n"
        i += 1
    return standings


"""
Get the stadings data to standings on Discord

"""


def get_standings_data(conference):
    conference = conference.lower()
    try:
        if conference == "acc":
            return parse_acc()
        elif conference == "american" or conference == "aac":
            return parse_aac()
        elif conference == "big ten" or conference == "b1g" or conference == "big 10" or conference == "b10":
            return parse_bigten()
        elif conference == "conference usa" or conference == "cusa" or conference == "c-usa":
            return parse_cusa()
        elif conference == "mac":
            return parse_mac()
        elif conference == "mountain west" or conference == "mwc":
            return parse_mwc()
        elif conference == "pac-12" or conference == "pac 12":
            return parse_pac12()
        elif conference == "sec":
            return parse_sec()
        elif conference == "sun belt" or conference == "sbc":
            return parse_sbc()
        elif conference == "big 12" or conference == "big xii" or conference == "b12":
            return parse_big12()
        elif conference == "independents" or conference == "independent":
            return parse_independents()
        elif conference == "america east":
            return parse_americaeast()
        elif conference == "atlantic sun":
            return parse_atlanticsun()
        elif conference == "big sky":
            return parse_bigsky()
        elif conference == "carolina" or conference == "carolina football conference" or conference == "cfc":
            return parse_cfc()
        elif conference == "colonial":
            return parse_colonial()
        elif conference == "delta" or conference == "delta intercollegiate":
            return parse_delta()
        elif conference == "ivy" or conference == "ivy league":
            return parse_ivy()
        elif conference == "mid atlantic" or conference == "mid-atlantic":
            return parse_midatlantic()
        elif conference == "missouri valley" or conference == "mvc":
            return parse_mvc()
        elif conference == "southland":
            return parse_southland()
        else:
            return "Conference not found"
    except Exception as e:
        print("The following error occurred: " + str(e))
        return None


"""
Parse the rankings worksheet standings

"""


def parse_rankings_worksheet(num_col, team_col, value_col, standings):
    ranks = rankingsWorksheet.col_values(num_col)
    teams = rankingsWorksheet.col_values(team_col)
    values = rankingsWorksheet.col_values(value_col)
    i = 4
    for team in teams[4:-1]:
        value = values[i]
        rank = ranks[i]
        if int(rank) > 25:
            break
        rankings = rankings + "#" + rank + " " + team.strip() + " " + value.strip() + "\n"
        i = i + 1
    return rankings


"""
Get the FBS coaches poll data

"""


def get_fbs_coaches_poll_data(r):
    return get_coaches_poll_data(r, "FBS")


"""
Get the FCS coaches poll data

"""


def get_fcs_coaches_poll_data(r):
    return get_coaches_poll_data(r, "FCS")
