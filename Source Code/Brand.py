import random
import os
import PySimpleGUI as sg
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import sys

class Brand:
    def __init__(self, name, men, women, title_names):
        self.name = name
        self.men = men
        self.women = women
        self.roster = men + women
        self.title_names = title_names

    def createTeams(self, team_number):
        men = self.men
        self.teams = []
        random.shuffle(men)
        for i in range(0, team_number * 2, 2):
            self.teams.append((men[i], men[i + 1]))
        return self.teams

    def createWomenTeams(self, womens_team_number):
        women = self.women
        self.women_teams = []
        if womens_team_number == 0: return self.women_teams
        random.shuffle(women)
        for i in range(0, womens_team_number * 2, 2):
            self.women_teams.append((women[i], women[i + 1]))
        return self.women_teams

    def assignChampions(self, teams, second_midcard, womens_midcard, team_roster):
        male_roster = self.men; female_roster = self.women
        random.shuffle(male_roster); random.shuffle(female_roster); random.shuffle(teams)

        self.champions = []
        second_midcard_champion = None
        womens_midcard_champion = None

        singles_division = male_roster
        if not team_roster: singles_division = Brand.removeTeams(male_roster, teams)

        world_champion = singles_division[0]
        midcard_champion = singles_division[1]
        womens_champion = female_roster[0]
        tag_champions = teams[0]
        if second_midcard: second_midcard_champion = singles_division[2]
        if womens_midcard: womens_midcard_champion = female_roster[1]

        self.champions = [world_champion, midcard_champion, womens_champion, tag_champions,
                          second_midcard_champion, womens_midcard_champion]
        return self.champions

    def assignDivisions(self, tag_teams, champions, second_midcard, womens_midcard, team_roster):
        men = self.men
        women = self.women
        self.divisions = []
        world_division, midcard_division, womens_division, tag_division, second_midcard_division, womens_midcard_division = [], [], [], [], [], []
        male_belt_order = 1
        female_belt_order = 1

        if not team_roster: Brand.removeTeams(men, tag_teams)

        random.shuffle(men)
        random.shuffle(women)
        random.shuffle(tag_teams)

        for m in men:
            match male_belt_order:
                case 1: world_division.append(m); male_belt_order = 2
                case 2:
                    midcard_division.append(m)
                    if second_midcard: male_belt_order = 3
                    else: male_belt_order = 1
                case 3: second_midcard_division.append(m); male_belt_order = 1

        for w in women:
            if female_belt_order == 1:
                womens_division.append(w)
                if womens_midcard: female_belt_order = 2
            else: womens_midcard_division.append(w); female_belt_order = 1

        tag_division = tag_teams

        self.divisions = [world_division, midcard_division, womens_division, tag_division,
                          second_midcard_division, womens_midcard_division]
        return self.divisions

    def assignFeuds(self, divisions, team_roster):
        self.feuds = []
        world_div, mid_div, women_div, tag_div = divisions[0], divisions[1], divisions[2], divisions[3]
        random.shuffle(world_div); random.shuffle(mid_div); random.shuffle(tag_div); random.shuffle(women_div)
        self.feuds.append(f"{world_div[0]} Vs. {world_div[1]}")
        self.feuds.append(f"{mid_div[0]} Vs. {mid_div[1]}")
        self.feuds.append(f"{women_div[0]} Vs. {women_div[1]}")

        if team_roster: self.feuds.append(f"{tag_div[0]} Vs. {tag_div[1]}")
        else: self.feuds.append(f"{tag_div[0][0]} & {tag_div[0][1]} Vs. {tag_div[1][0]} & {tag_div[1][1]}")
        return self.feuds

    @staticmethod
    def removeTeams(roster, teams):
        singles_roster = roster
        for r in singles_roster:
            for t in teams:
                if r == t[0] or r == t[1]:
                    singles_roster.remove(r)
        return singles_roster
