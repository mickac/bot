import discord
import gspread
import os
import re

from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

load_dotenv()
DISCORD_TOKEN = os.getenv('DISCORD_TOKEN')


class SpreadSheetMixtin:

    def __init__(self, anime, anime_link, season_title):
        pattern = re.compile(r"0️⃣|1️⃣|2️⃣|3️⃣|4️⃣|5️⃣|6️⃣|7️⃣|8️⃣|9️⃣|🔟")
        self.anime = pattern.sub("", anime)
        self.anime_link = anime_link
        self.season_title = season_title

    def create_keyfile_dict(self):
        return {
            "type": os.getenv('TYPE'),
            "project_id": os.getenv('PROJECT_ID'),
            "private_key_id": os.getenv('PRIVATE_KEY_ID'),
            "private_key": os.getenv('PRIVATE_KEY'),
            "client_email": os.getenv('CLIENT_EMAIL'),
            "client_id": os.getenv('CLIENT_ID'),
            "auth_uri": os.getenv('AUTH_URI'),
            "token_uri": os.getenv('TOKEN_URI'),
            "auth_provider_x509_cert_url": os.getenv('AUTH_PROVIDER_X509_CERT_URL'),
            "client_x509_cert_url": os.getenv('CLIENT_X509_CERT_URL'),
        }

    def get_rgb_number(self, rgb):
        return rgb/255 if rgb > 0 else rgb

    def add_anime_to_spreadsheet(self):
        ANIME_LIST_ROW = 4
        BORDER_STYLE = {
            'style': 'SOLID',
            'colorStyle': {
                'rgbColor': {
                    'red': self.get_rgb_number(0),
                    'green': self.get_rgb_number(0),
                    'blue': self.get_rgb_number(0)
                }
            }
        }
        FILE_NAME = 'excelek'

        gclient = gspread.authorize(
            ServiceAccountCredentials.from_json_keyfile_dict(
                self.create_keyfile_dict(), [
                    'https://spreadsheets.google.com/feeds',
                    'https://www.googleapis.com/auth/drive'
                ]
            )
        )
        spreadsheet = gclient.open(FILE_NAME)

        try:
            sheet = spreadsheet.worksheet(self.season_title)
        except gspread.exceptions.WorksheetNotFound:
            sheet = spreadsheet.add_worksheet(
                title=self.season_title, rows=50, cols=50)

        anime_list = sheet.row_values(4)
        if not sheet.find(self.anime):
            last_anime = anime_list[-1] if anime_list else None
            sheet.update_cell(
                sheet.find(last_anime).row if last_anime else ANIME_LIST_ROW,
                sheet.find(last_anime).col + 1 if last_anime else 1,
                self.anime
            )
            sheet = spreadsheet.worksheet(self.season_title)

        selected_anime_row = sheet.find(self.anime).row
        selected_anime_col = sheet.find(self.anime).col
        if not sheet.cell(selected_anime_row + 1, selected_anime_col).value:
            incremented_last_episode_number = ANIME_LIST_ROW + 1
            sheet.update_cell(
                incremented_last_episode_number,
                selected_anime_col,
                '1'
            )
        else:
            incremented_last_episode_number = ANIME_LIST_ROW + int(
                sheet.col_values(selected_anime_col)[-1]
            ) + 1
            sheet.update_cell(
                incremented_last_episode_number,
                selected_anime_col,
                incremented_last_episode_number - ANIME_LIST_ROW
            )

        sheet.format(
            gspread.utils.rowcol_to_a1(
                incremented_last_episode_number,
                selected_anime_col
            ), {
                'backgroundColor': {
                    'red': self.get_rgb_number(228),
                    'green': self.get_rgb_number(187),
                    'blue': self.get_rgb_number(217)
                },
                'borders': {
                    'bottom': BORDER_STYLE,
                    'left': BORDER_STYLE,
                    'right': BORDER_STYLE,
                    'top': BORDER_STYLE
                },
                'horizontalAlignment': 'CENTER',
                'textFormat': {
                    'fontSize': 9,
                    'link': {'uri': self.anime_link}
                }
            }
        )


intents = discord.Intents.default()
intents.message_content = True

client = discord.Client(intents=intents)


@client.event
async def on_ready():
    print(f'We have logged in as {client.user}')


@client.event
async def on_message(message):
    if message.channel.name == 'wydawane-anime' and message.content.startswith('Sezon:'):
        content = message.content
        anime = content.splitlines()[1]
        anime_link = re.search(r"Hard CDA:\s+(.*?)\n", content).group(1)
        season_title = re.search(r"Sezon: (.*)", content).group(1)
        if anime and anime_link and season_title:
            spreadsheet = SpreadSheetMixtin(
                anime=anime, anime_link=anime_link, season_title=season_title
            )
            spreadsheet.add_anime_to_spreadsheet()

client.run(DISCORD_TOKEN)
