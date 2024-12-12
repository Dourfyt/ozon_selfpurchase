from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton
from aiogram.filters import Command
from aiogram.types import FSInputFile
import asyncio
import os
import subprocess

# Токен бота
API_TOKEN = "6501489744:AAEVS3aoQG1JsbkJBWRdTWw7__JuWJvF43w"

# Создаем объект бота и диспетчер
bot = Bot(token=API_TOKEN)
dp = Dispatcher()

# Клавиатура с кнопками
keyboard = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton(text="Загрузить таблицу"), KeyboardButton(text="Скачать таблицу")],
        [KeyboardButton(text="Запустить самовыкуп")]
    ],
    resize_keyboard=True
)

# Обработчик команды /start
@dp.message(Command("start"))
async def cmd_start(message: Message):
    await message.answer("Привет! Выберите действие:", reply_markup=keyboard)

# Обработчик кнопки "Загрузить таблицу"
@dp.message(F.text == "Загрузить таблицу")
async def upload_table_request(message: Message):
    await message.answer("Отправьте файл в формате XLSX. Он будет сохранен как 'file.xlsx'.")

# Обработчик отправки файла
@dp.message(F.document)
async def handle_file(message: Message):
    document = message.document
    if document.file_name.endswith(".xlsx"):
        file_path = "file.xlsx"
        await bot.download(document, file_path)
        await message.answer("Файл успешно загружен и сохранен как 'file.xlsx'.")
    else:
        await message.answer("Пожалуйста, отправьте файл в формате XLSX.")

# Обработчик кнопки "Скачать таблицу"
@dp.message(F.text == "Скачать таблицу")
async def download_table(message: Message):
    file_path = "file.xlsx"
    if os.path.exists(file_path):
        file = FSInputFile(file_path)
        await message.answer_document(file, caption="Вот ваша таблица.")
    else:
        await message.answer("Файл 'file.xlsx' не найден. Сначала загрузите таблицу.")

# Обработчик кнопки "Запустить самовыкуп"
@dp.message(F.text == "Запустить самовыкуп")
async def start_self_buyout(message: Message):
    # Запуск скрипта с помощью subprocess
    try:
        await message.answer("Самовыкуп запущен")
        # Запускаем файл selenium_bot/driver.py
        result = subprocess.run(["python", "selenium_bot/driver.py"], capture_output=True, text=True)
        # Проверим, если скрипт завершился успешно
        if result.returncode == 0:
            await message.answer("Самовыкуп успешно завершен!")
        else:
            await message.answer(f"Произошла ошибка при запуске самовыкупа: {result.stderr}")
    except Exception as e:
        await message.answer(f"Ошибка при запуске самовыкупа: {str(e)}")

# Главная функция запуска бота
async def main():
    print("Бот запущен!")
    try:
        # Запуск диспетчера (поллинг)
        await dp.start_polling(bot)
    finally:
        # Закрываем сессию бота при завершении работы
        await bot.session.close()

if __name__ == "__main__":
    asyncio.run(main())
