import datetime
import asyncio
import logging
import os
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram import F
from aiogram.types.input_file import FSInputFile
from parse import Parser

logging.basicConfig(level=logging.INFO)
bot = Bot(token="Токен")
dp = Dispatcher()

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    await message.answer(
        "Отправь файл Excel, чтобы получить обработанный вариант!")


@dp.message(F.document)
async def get_doc(message: types.Message):
    try:
        file_info = await bot.get_file(message.document.file_id)

        file_name = message.document.file_name
        path = os.path.join(file_name)

        await bot.download_file(file_info.file_path, destination=path)

        parser = Parser(path)
        output_file_path = await parser.run()

        if os.path.exists(output_file_path):
            output_file_to_send = FSInputFile(output_file_path)

            await bot.send_document(chat_id=message.chat.id, document=output_file_to_send)

        else:
            await message.answer("Не удалось создать файл. Пожалуйста, попробуйте снова.")

    except Exception as e:
        logging.error(f"Ошибка при обработке документа: {e}")
        await message.answer("Произошла ошибка при загрузке или обработке файла. Пожалуйста, попробуйте снова.")

    finally:
        if os.path.exists(path):
            os.remove(path)
            logging.info(f"Файл {path} был удалён.")

        if os.path.exists(output_file_path):
            os.remove(output_file_path)
            logging.info(f"Файл {output_file_path} был удалён.")


async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
