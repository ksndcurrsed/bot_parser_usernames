from telethon import TelegramClient
from telethon.errors import UsernameNotOccupiedError, UsernameInvalidError
from telethon.tl.types import User, Channel
from telethon.tl.types import UserStatusRecently, UserStatusLastMonth, UserStatusLastWeek, UserStatusOnline, UserStatusOffline
import openpyxl
import os

class Parser:
    def __init__(self, input_file):
        self.api_id = 'АпиАЙДИ'
        self.api_hash = 'Апи'
        self.input_file = input_file
        self.output_file = 'output.xlsx'
        self.client = TelegramClient('session_name', self.api_id, self.api_hash)

    async def check_username(self, username):
        try:
            entity = await self.client.get_entity(username)
            if isinstance(entity, User):
                status = "человек"
                last_online = self.parse_user_status(entity.status)
            elif isinstance(entity, Channel):
                if entity.broadcast:
                    status = "канал"
                else:
                    status = "чат"
                last_online = "N/A"  # статус не отображается для каналов/чатов
            return status, last_online
        except (UsernameNotOccupiedError, UsernameInvalidError, ValueError):
            return "не существует", "N/A"

    def parse_user_status(self, status):
        if isinstance(status, UserStatusRecently):
            return "недавно был в сети"
        elif isinstance(status, UserStatusLastWeek):
            return "был в сети на прошлой неделе"
        elif isinstance(status, UserStatusLastMonth):
            return "был в сети в прошлом месяце"
        elif isinstance(status, UserStatusOnline):
            return f"в сети ({status.expires.strftime('%Y-%m-%d %H:%M:%S')})"
        elif isinstance(status, UserStatusOffline):
            return f"был в сети {status.was_online.strftime('%Y-%m-%d %H:%M:%S')}"
        else:
            return "не видно"

    async def process_usernames(self):
        wb = openpyxl.load_workbook(self.input_file)
        sheet = wb.active

        output_wb = openpyxl.Workbook()
        output_sheet = output_wb.active
        output_sheet.append(['Username', 'Статус', 'Последний раз в сети'])

        async with self.client:
            for row in sheet.iter_rows(min_row=1, values_only=True):
                username = row[0]
                status, last_online = await self.check_username(username)
                output_sheet.append([username, status, last_online])

        output_wb.save(self.output_file)

        return os.path.abspath(self.output_file)

    async def run(self):
        return await self.process_usernames()
