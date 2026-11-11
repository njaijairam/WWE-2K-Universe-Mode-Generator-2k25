import random, os
import FreeSimpleGUI as sg
import openpyxl as xl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import sys
from Brand import *

def getRoster(roster_file):
    with open(roster_file) as f:
        roster = f.readlines()
    return [r.strip() for r in roster if r.strip()]

def draftRoster(roster, brands):
    raw_roster, smackdown_roster, nxt_roster, wcw_roster = [], [], [], []
    b = 1  
    random.shuffle(roster)  

    for r in roster:
        match b:
            case 1:
                raw_roster.append(r)
                if brands > 1: b = 2
                else: b = 1
                continue
            case 2:
                smackdown_roster.append(r)
                if brands > 2: b = 3
                else: b = 1
                continue
            case 3:
                nxt_roster.append(r)
                if brands > 3: b = 4
                else: b = 1
                continue
            case 4:
                wcw_roster.append(r)
                b = 1
                continue

    return [raw_roster, smackdown_roster, nxt_roster, wcw_roster]

def addTeams(roster, teams):
    combined = list(roster)
    for t in teams: combined.append(t[0]); combined.append(t[1])
    return combined

def makeTabLayout(brand_name, brand_key, title_names, key_names):
    brand_layout = [[sg.Text("Brand Name: "), sg.Input(key=brand_key, default_text=brand_name, enable_events=True)],
                    [sg.Text("Men's World Title: "), sg.Input(key=key_names[0], default_text=title_names[0])],
                    [sg.Text("Men's Midcard Title: "), sg.Input(key=key_names[1], default_text=title_names[1])],
                    [sg.Text("Men's Tag Team Title: "), sg.Input(key=key_names[2], default_text=title_names[2])],
                    [sg.Text("Intergender Title: "), sg.Input(key=key_names[3], default_text=title_names[3])],
                    [sg.Text("Women's World Title: "), sg.Input(key=key_names[4], default_text=title_names[4])],
                    [sg.Text("Women's Tag Team Title: "), sg.Input(key=key_names[5], default_text=title_names[5])]
                    ]
    return brand_layout

def makeTxtFile(f, Brand, file_type, team_roster):
    if file_type == "Spreadsheet": return

    title_names = Brand.title_names

    f.write(f"{Brand.name}\n")
    f.write("=" * 30 + "\n\n")

    if not team_roster: full_mens_roster = addTeams(Brand.men, Brand.teams)
    else: full_mens_roster = Brand.men

    f.write("Men's Roster\n")
    f.write("=" * 30 + "\n\n")
    for m in full_mens_roster: f.write(f"{m}\n")
    f.write("\n")

    f.write("Women's Roster\n")
    f.write("=" * 30 + "\n")
    for w in Brand.women: f.write(f"{w}\n")
    f.write("\n")

    f.write("Tag Teams\n")
    f.write("=" * 30 + "\n")
    for t in Brand.teams:
        if team_roster: f.write(f"{t}\n")
        else: f.write(f"{t[0]} & {t[1]}\n")
    f.write("\n")

    f.write("Champions\n")
    f.write("=" * 30 + "\n")
    f.write(f"{title_names[0]} (Men's World): {Brand.champions[0]}\n") 
    f.write(f"{title_names[1]} (Men's Midcard): {Brand.champions[1]}\n")  
    if team_roster: f.write(f"{title_names[2]} (Men's Tag): {Brand.champions[2]}\n") 
    else: f.write(f"{title_names[2]} (Men's Tag): {Brand.champions[2][0]} & {Brand.champions[2][1]}\n")
    f.write(f"{title_names[3]} (Intergender): {Brand.champions[3]}\n")
    f.write(f"{title_names[4]} (Women's World): {Brand.champions[4]}\n") 
    if Brand.champions[5]:
        if team_roster: f.write(f"{title_names[5]} (Women's Tag): {Brand.champions[5]}\n")
        else: f.write(f"{title_names[5]} (Women's Tag): {Brand.champions[5][0]} & {Brand.champions[5][1]}\n")
    f.write("\n\n")

    f.write("Rivalries\n")
    f.write("=" * 30 + "\n")
    for r in Brand.feuds: f.write(f"{r}\n")
    f.write("\n\n")

def writeWomenTag(f, womens_teams, womens_tag_champions, women_team_roster):
    f.write("\n\n")
    f.write("Women's Tag Division\n")
    f.write("=" * 30 + "\n\n")  
    if womens_tag_champions:
        if women_team_roster: f.write(f"Women's Tag Team Champions: {womens_tag_champions}\n\n")
        else: f.write(f"Women's Tag Team Champions: {womens_tag_champions[0]} & {womens_tag_champions[1]}\n\n")
    for w in womens_teams[1:]: 
        if women_team_roster: f.write(f"{w}\n")
        else: f.write(f"{w[0]} & {w[1]}\n")

def writeIntergenderDivision(f, ig_division, ig_champion, title_name):
    f.write("\n\n")
    f.write(f"Shared Intergender Division ({title_name})\n")
    f.write("=" * 30 + "\n\n")
    f.write(f"Intergender Champion: {ig_champion}\n\n")
    for member in ig_division[1:]:
        f.write(f"{member}\n")

def createSpreadsheet(Brand, brand_number, draft_workbook, roster_sheet, tag_teams_sheet, champions_sheet,
        divisions_sheet, feuds_sheet, women_tag_sheet, womens_tag_division, file_type, team_roster, women_team_roster):

    tn = Brand.title_names
    roster_sheet.cell(row=1, column=brand_number).value = Brand.name

    if team_roster:
        full_roster = Brand.men + Brand.women
        full_roster.sort()
        full_roster = full_roster + Brand.teams
    else: full_roster = addTeams(Brand.men + Brand.women, Brand.teams)
    full_roster.sort()

    for row in range(len(full_roster)): roster_sheet.cell(row=row + 2, column=brand_number).value = full_roster[row]

    if not women_team_roster: full_teams = Brand.teams + Brand.women_teams
    else: full_teams = Brand.teams
    tag_teams_sheet.cell(row=1, column=brand_number).value = Brand.name
    for row in range(len(full_teams)):
        if team_roster: tag_teams_sheet.cell(row=row + 2, column=brand_number).value = f"{full_teams[row]}"
        else: tag_teams_sheet.cell(row=row + 2, column=brand_number).value = f"{full_teams[row][0]} & {full_teams[row][1]}"

    if women_team_roster:
        for row in range(len(Brand.women_teams)):
            tag_teams_sheet.cell(row=row + len(Brand.teams) + 2, column=brand_number).value = Brand.women_teams[row]

    champions_sheet.cell(row=2, column=1).value = "Mens World"
    champions_sheet.cell(row=3, column=1).value = "Mens Midcard"
    champions_sheet.cell(row=4, column=1).value = "Mens Tag"
    champions_sheet.cell(row=5, column=1).value = "Intergender"
    champions_sheet.cell(row=6, column=1).value = "Womens World"
    champions_sheet.cell(row=7, column=1).value = "Womens Tag"

    champions_sheet.cell(row=1, column=brand_number + 1).value = Brand.name
    champions_sheet.cell(row=2, column=brand_number + 1).value = Brand.champions[0] 
    champions_sheet.cell(row=3, column=brand_number + 1).value = Brand.champions[1] 
    if team_roster: champions_sheet.cell(row=4, column=brand_number + 1).value = f"{Brand.champions[2]}"
    else: champions_sheet.cell(row=4, column=brand_number + 1).value = f"{Brand.champions[2][0]} & {Brand.champions[2][1]}" 
    champions_sheet.cell(row=5, column=brand_number + 1).value = Brand.champions[3]
    champions_sheet.cell(row=6, column=brand_number + 1).value = Brand.champions[4]
    if Brand.champions[5]:
        if team_roster: champions_sheet.cell(row=7, column=brand_number + 1).value = f"{Brand.champions[5]}"
        else: champions_sheet.cell(row=7, column=brand_number + 1).value = f"{Brand.champions[5][0]} & {Brand.champions[5][1]}"

    feuds_sheet.cell(row=1, column=brand_number).value = Brand.name
    for row in range(len(Brand.feuds)):
        feuds_sheet.cell(row=row + 2, column=brand_number).value = Brand.feuds[row]

    brand_colors = [
        PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid'),
        PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid'),
        PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid'),
        PatternFill(start_color='00B050', end_color='00B050', fill_type='solid')
    ]

    for b in range(1, brand_number + 1):
        roster_sheet.cell(row=1, column=b).fill = brand_colors[b - 1]
        tag_teams_sheet.cell(row=1, column=b).fill = brand_colors[b - 1]
        champions_sheet.cell(row=1, column=b + 1).fill = brand_colors[b - 1]
        feuds_sheet.cell(row=1, column=b).fill = brand_colors[b - 1]

    if file_type == "Spreadsheet": draft_workbook.save("Draft.xlsx")

def formatSpreadsheet(cell_range, style):
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    for cell in cell_range:
        for c in cell:
            if style == "Header":
                c.font = Font(name="Arial", size=8, bold=True)
                c.alignment = Alignment(horizontal='center')
            elif style == "Body":
                c.font = Font(name="Arial", size=9, bold=False)
            c.border = thin_border

def developBrand(Brand, team_number, women_tag_number, draft_file, file_type, teams, women_teams, team_roster_flag):
    if teams: Brand.teams = teams
    else: Brand.createTeams(team_number)

    if women_teams: Brand.women_teams = women_teams
    else: Brand.createWomenTeams(women_tag_number)

    Brand.assignChampions(Brand.teams, team_roster_flag)
    Brand.assignDivisions(Brand.teams, team_roster_flag)
    Brand.assignFeuds(Brand.divisions, team_roster_flag)
    if file_type == "Text":
        makeTxtFile(draft_file, Brand, "Text", team_roster_flag)

def createWomenTagDivision(women_tag_number, brands, draft_file, file_type, women_team_roster):
    if women_tag_number == 0: return []
    womens_tag_division = []
    for b in brands:
        womens_tag_division += b.women_teams
    if not womens_tag_division: return []
    random.shuffle(womens_tag_division)
    womens_tag_champions = womens_tag_division[0]
    if file_type == "Text": writeWomenTag(draft_file, womens_tag_division, womens_tag_champions, women_team_roster)
    return womens_tag_division

def createIntergenderDivision(ig_number, brands, draft_file, file_type):
    if ig_number == 0: return []
    ig_division = []
    for b in brands:
        brand_pool = b.men + b.women
        random.shuffle(brand_pool)
        ig_division += brand_pool[:ig_number]
    if not ig_division: return []
    random.shuffle(ig_division)
    ig_champion = ig_division[0]
    title_name = brands[0].title_names[3] if len(brands[0].title_names) > 3 else "Intergender Championship"
    if file_type == "Text": 
        writeIntergenderDivision(draft_file, ig_division, ig_champion, title_name)
    return ig_division
