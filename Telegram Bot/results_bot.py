import asyncio
from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import FSInputFile

API_TOKEN = "API_TG_BOT"

bot = Bot(token=API_TOKEN)
dp = Dispatcher()

RESULTS_FILE = "results.xlsx"  # Имя файла результатов (должен находиться в рабочей папке с этим ботом)

def is_admin(user_id):
    # Сейчас бот доступен всем. Для ограничения доступа — раскомментировать функцию ниже и прописать ID преподавателей.
    return True

# Сейчас доступен всем. Потом сюда можно вставить проверку ID преподавателя.
#def is_admin(user_id):
#    ADMIN_IDS = [123456789, 987654321]  # <-- сюда ID нужных пользователей
#    return user_id in ADMIN_IDS

@dp.message(Command("start"))
async def cmd_start(message: types.Message):
    """Стартовое сообщение и кнопка для получения файла."""
    await message.answer(
        "Привет! Я бот для получения результатов студентов.\n\n"
        "Нажмите кнопку ниже, чтобы получить файл Excel с результатами.",
        reply_markup=types.ReplyKeyboardMarkup(
            keyboard=[
                [types.KeyboardButton(text="Получить результаты")]
            ],
            resize_keyboard=True
        )
    )

@dp.message(lambda m: m.text and "результат" in m.text.lower())
async def send_results(message: types.Message):
    """
    Проверка доступа и отправка Excel-файла.
    Если пользователь не админ — файл не отправляется.
    """
    if not is_admin(message.from_user.id):
        await message.answer("У вас нет доступа к результатам.")
        return

    try:
        file = FSInputFile(RESULTS_FILE)
        await message.answer_document(file, caption="Результаты всех студентов (Excel).")
    except FileNotFoundError:
        await message.answer("Файл результатов пока не создан. Нет данных или никто не завершил тест.")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
