from telethon import TelegramClient
from telethon.errors import UsernameNotOccupiedError, UsernameInvalidError, FloodWaitError
from telethon.tl.types import User, Channel
from telethon.tl.types import UserStatusRecently, UserStatusLastMonth, UserStatusLastWeek, UserStatusOnline, UserStatusOffline
import openpyxl
import os
import logging
import asyncio

class Parser:
    def __init__(self, input_file):
        self.api_id = '22731591'
        self.api_hash = '5fb07e3e7604e11fdfc9c40d7232c944'
        self.input_file = input_file
        self.output_file = 'output.xlsx'
        self.client = TelegramClient('session_name', self.api_id, self.api_hash)

    async def get_entity_safely(self, username):
        try:
            input_entity = await self.client.get_input_entity(username)
            entity = await self.client.get_entity(input_entity)
            return entity
        except UsernameNotOccupiedError:
            return None
        except UsernameInvalidError:
            return None
        except FloodWaitError as e:
            logging.warning(f"FloodWaitError: необходимо ожидание {e.seconds} секунд.")
            await asyncio.sleep(e.seconds)  # Ожидание перед повторной попыткой
            return await self.get_entity_safely(username)
        except Exception as e:
            logging.error(f"Ошибка при получении сущности для {username}: {e}")
            return None

    async def check_username(self, username):
        try:
            if not username:
                return "не существует", "N/A"

            entity = await self.get_entity_safely(username)

            if isinstance(entity, User):
                status = "человек"
                last_online = self.parse_user_status(entity.status)
            elif isinstance(entity, Channel):
                status = "канал" if entity.broadcast else "чат"
                last_online = "N/A"
            else:
                return "не существует", "N/A"

            return status, last_online

        except Exception as e:
            logging.error(f"Ошибка при проверке пользователя {username}: {e}")
            return "ошибка", "N/A"

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
        try:
            wb = openpyxl.load_workbook(self.input_file)
            sheet = wb.active

            # Создание выходного Excel файла
            output_wb = openpyxl.Workbook()
            output_sheet = output_wb.active
            output_sheet.append(['Username', 'Статус', 'Последний раз в сети'])

            usernames = [row[0] for row in sheet.iter_rows(min_row=1, values_only=True) if row[0]]
            total_usernames = len(usernames)

            async with self.client:
                for index, username in enumerate(usernames, start=1):
                    if username:
                        try:
                            status, last_online = await self.check_username(username)
                            output_sheet.append([username, status, last_online])

                            # Вывод в консоль проверенного пользователя и количества оставшихся
                            remaining = total_usernames - index
                            print(f"Проверен: {username}, Статус: {status}, Осталось: {remaining} пользователей")

                            # После каждых 100 проверенных пользователей делать паузу 5 минут (300 секунд)
                            if index % 100 == 0:
                                print(f"Пауза на 5 минут после проверки {index} пользователей.")
                                await asyncio.sleep(300)  # Пауза 5 минут (300 секунд)

                            await asyncio.sleep(5)  # Увеличенная задержка между запросами (5 сек)

                        except FloodWaitError as e:
                            print(f"FloodWaitError: Ожидание {e.seconds} секунд.")
                            await asyncio.sleep(e.seconds)

                    else:
                        logging.info("Пропуск пустой строки или пустого имени пользователя")

            # Сохранение выходного файла
            output_wb.save(self.output_file)

            return os.path.abspath(self.output_file)

        except Exception as e:
            logging.error(f"Ошибка при обработке списка пользователей: {e}")
            return None

    async def run(self):
        return await self.process_usernames()
