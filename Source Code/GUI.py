import os, FreeSimpleGUI as sg
from functions import *
def gui():
    current_directory = os.getcwd()
    sg.theme("DarkBlack")

    # 6 title slots per brand: World, Midcard, Tag, Intergender, Women's World, Women's Tag
    title_names = [
        ["World Heavyweight", "Intercontinental", "Raw Tag Team", "Intergender", "Women's World", "Women's Tag Team"],  # Raw
        ["WWE Universal", "United States", "Smackdown Tag Team", "Intergender", "WWE Women's", "Women's Tag Team"],  # Smackdown
        ["NXT", "North American", "NXT Tag Team", "Intergender", "NXT Women's", "Women's Tag Team"],  # NXT
        ["WCW World", "WCW US", "WCW World Tag Team", "Intergender", "WCW Women's", "Women's Tag Team"]  # WCW
    ]

    key_names = [
        ["raw_world", "raw_mid", "raw_tag", "raw_ig", "raw_women", "raw_w_tag"],  # Raw
        ["sd_world", "sd_mid", "sd_tag", "sd_ig", "sd_women", "sd_w_tag"],  # Smackdown
        ["nxt_world", "nxt_mid", "nxt_tag", "nxt_ig", "nxt_women", "nxt_w_tag"],  # NXT
        ["wcw_world", "wcw_mid", "wcw_tag", "wcw_ig", "wcw_women", "wcw_w_tag"]  # WCW
    ]

    men = sg.FileBrowse("Men's Roster", file_types=[("TXT Files", "*.txt")], initial_folder=current_directory)
    women = sg.FileBrowse("Women's Roster", file_types=[("TXT Files", "*.txt")], initial_folder=current_directory)
    team_roster = sg.FileBrowse("Men's Teams Roster", file_types=[("TXT Files", "*.txt")], initial_folder=current_directory, visible=False)
    women_team_roster = sg.FileBrowse("Women's Teams Roster", file_types=[("TXT Files", "*.txt")], initial_folder=current_directory, visible=False)

    brands = [1, 2, 3, 4]
    teams = [i for i in range(1, 21)]
    women_teams = [i for i in range(0, 11)]
    ig_numbers = [i for i in range(0, 11)]

    main_layout = [
        [sg.Text("Select Your Male and Female Roster Files")],  
        [sg.InputText(key="-FILE_PATH-"), men],  
        [sg.InputText(key="-FILE_PATH2-"), women],  
        [sg.InputText(key='-FILE_PATH3-', visible=False), team_roster], 
        [sg.InputText(key='-FILE_PATH4-', visible=False), women_team_roster],
        [sg.Text("Select the number of brands (Max 4)"),  
         sg.Spin(brands, initial_value=4, key='brands', size=4, enable_events=True)],
        [sg.Text("Select the number of tag teams per brand"),  
         sg.Spin(teams, key='teams', initial_value=4, size=4)],
        [sg.Text("Select the number of Women's Tag Teams you want per brand"),  
         sg.Spin(women_teams, initial_value="0", key='wTag', size=4, enable_events=True)],
        [sg.Text("Intergender participants per brand (Shared Title)"),
         sg.Spin(ig_numbers, initial_value="0", key='ig_num', size=4)],
        [sg.Checkbox("Separate Team Roster", key='team_roster', enable_events=True),
         sg.Checkbox("Separate Women Team Roster", key='women_team_roster', enable_events=True)],
        [sg.Text("Which file type?"), sg.Combo(["Text", "Spreadsheet"], default_value="Text", key='file_type')]
    ]

    raw_layout = makeTabLayout("Raw", "raw_name", title_names[0], key_names[0])
    sd_layout = makeTabLayout("Smackdown", "sd_name", title_names[1], key_names[1])
    nxt_layout = makeTabLayout("NXT", "nxt_name", title_names[2], key_names[2])
    wcw_layout = makeTabLayout("WCW", "wcw_name", title_names[3], key_names[3])

    layout = [[sg.TabGroup([
        [sg.Tab('General', main_layout),
         sg.Tab('Raw', raw_layout, key='raw_tab'),
         sg.Tab('Smackdown', sd_layout, key='sd_tab'),
         sg.Tab('NXT', nxt_layout, key='nxt_tab'),
         sg.Tab('WCW', wcw_layout, key='wcw_tab')
         ]])],
        [sg.Button("Generate"), sg.Exit()]
    ]

    window = sg.Window("2K Universe Generator", layout)

    while True:
        event, values = window.read()

        refresh = ["brands", "raw_name", "sd_name", "nxt_name", "wcw_name"]

        if event in (sg.WIN_CLOSED, 'Exit'):
            sys.exit()

        elif event == "Generate":
            if values["-FILE_PATH4-"] and values["wTag"] == 0: values["wTag"] = 1

            male_roster, female_roster = values["-FILE_PATH-"], values["-FILE_PATH2-"]
            brand_number, tag_teams = values["brands"], values["teams"]
            womens_tag_team_title = values["wTag"]
            intergender_number = values["ig_num"]
            file_type = values["file_type"]

            if file_type == "Spreadsheet":
                try:
                    draft_file = open("Draft.xlsx", 'w')
                    draft_file.close()
                    draft_file = open("Draft.xlsx", 'a')
                except IOError:
                    sg.PopupError("The Draft.xlsx file is already open.\nPlease close the file")
                    continue

            brand_names = [values["raw_name"], values["sd_name"], values["nxt_name"], values["wcw_name"]]

            title_names = [
                [values["raw_world"], values["raw_mid"], values["raw_tag"], values["raw_ig"], values["raw_women"], values["raw_w_tag"]],
                [values["sd_world"], values["sd_mid"], values["sd_tag"], values["sd_ig"], values["sd_women"], values["sd_w_tag"]],
                [values["nxt_world"], values["nxt_mid"], values["nxt_tag"], values["nxt_ig"], values["nxt_women"], values["nxt_w_tag"]],
                [values["wcw_world"], values["wcw_mid"], values["wcw_tag"], values["wcw_ig"], values["wcw_women"], values["wcw_w_tag"]]
            ]

            sep_team, sep_wteam = values["team_roster"], values["women_team_roster"]

            team_roster, women_team_roster = False, False
            if sep_team: team_roster = values["-FILE_PATH3-"]
            if sep_wteam: women_team_roster = values["-FILE_PATH4-"]

            gui_inputs = [male_roster, female_roster, int(brand_number), int(tag_teams),
                          int(womens_tag_team_title), int(intergender_number),
                          brand_names, title_names, file_type, team_roster, women_team_roster]

            if int(brand_number) > 4: gui_inputs[2] = brand_number = 4
            if int(tag_teams) > 20: gui_inputs[3] = tag_teams = 20
            if int(womens_tag_team_title) > 10: gui_inputs[4] = womens_tag_team_title = 10

            if not bool(male_roster) or not bool(female_roster):
                sg.PopupError("Missing Roster File", "Please give both the men and women rosters.")
                continue

            if len(getRoster(male_roster)) == 0 or len(getRoster(female_roster)) == 0:
                sg.PopupError("Empty Roster File", "One or both roster files is empty.")
                continue

            return gui_inputs

        elif event in refresh:
            brand_number = values["brands"]
            window['raw_tab'].update(title=values["raw_name"])
            window['sd_tab'].update(visible=(brand_number >= 2), title=values["sd_name"])
            window['nxt_tab'].update(visible=(brand_number >= 3), title=values["nxt_name"])
            window['wcw_tab'].update(visible=(brand_number >= 4), title=values["wcw_name"])

        if values['team_roster']: window['-FILE_PATH3-'].update(visible=True); team_roster.update(visible=True)
        else: window['-FILE_PATH3-'].update(visible=False); team_roster.update(visible=False)
        if values['women_team_roster']: window['-FILE_PATH4-'].update(visible=True); women_team_roster.update(visible=True)
        else: window['-FILE_PATH4-'].update(visible=False); women_team_roster.update(visible=False)

    window.close()
