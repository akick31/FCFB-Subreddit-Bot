import praw
import discord
import json
from standings_drawer import *


with open('config.json', 'r') as config_file:
    config_data = json.load(config_file)


"""
Login to reddit

"""


def login_reddit():
    r = praw.Reddit(user_agent=config_data['user_agent'],
                    client_id=config_data['client_id'],
                    client_secret=config_data['client_secret'],
                    username=config_data['username'],
                    password=config_data['password'],
                    subreddit=config_data['subreddit'])
    return r