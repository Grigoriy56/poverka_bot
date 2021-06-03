from aiogram import Bot, types
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from states import *
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.contrib.middlewares.logging import LoggingMiddleware
from data import *
from datetime import timedelta, date
import openpyxl
from openpyxl.styles import PatternFill, Font
from config import TOKEN
from work import create_excel, add_line, preparation

line = 2 #эта переменная хранит строку, куда нужно записывать значение
bot = Bot(token=TOKEN)

dp = Dispatcher(bot, storage=MemoryStorage())
dp.middleware.setup(LoggingMiddleware())

###
counter_list = list(counter.keys())

button_1 = KeyboardButton(counter_list[0])
button_2 = KeyboardButton(counter_list[1])
button_3 = KeyboardButton(counter_list[2])
button_4 = KeyboardButton(counter_list[3])
button_5 = KeyboardButton(counter_list[4])
button_6 = KeyboardButton(counter_list[5])

markup_1 = ReplyKeyboardMarkup(resize_keyboard=True).row(button_1, button_2, button_3)
markup_1.row(button_4, button_5, button_6)

temp = ['Горячая', 'Холодная']
button_hot = KeyboardButton(temp[0])
button_cold = KeyboardButton(temp[1])
markup_temp = ReplyKeyboardMarkup(resize_keyboard=True).add(button_hot, button_cold)

dates = ['Сегодня', 'Ввести дату']
button_today = KeyboardButton(dates[0])
button_date = KeyboardButton(dates[1])
markup_date = ReplyKeyboardMarkup(resize_keyboard=True).add(button_today, button_date)
###


# learn answer on start
@dp.message_handler(commands=['start'], state="*")
async def process_start_command(message: types.Message):
    await message.reply('Привет, выбери счётчик из списка', reply_markup=markup_1, reply=False)
    await Test.start.set()


@dp.message_handler(state=Test.start)
async def echo_message(msg: types.Message, state: FSMContext):
    if msg.text in counter_list:
        global gos_number    #запоминиаем ключ для словаря со счётчиками
        gos_number = msg.text
        await bot.send_message(msg.from_user.id, 'Горячая или холодная?', reply_markup=markup_temp)
        await Test.st_count.set()
        await state.update_data(answer1=gos_number) # в состоянии запоминаю номер счётчика
    else:
        await bot.send_message(msg.from_user.id, 'Счётчик не в списке. Начните заново. /start')
        await Test.start.set()


@dp.message_handler(state=Test.st_count)
async def temp_message(msg: types.Message):
    global temp_now #переменная, чтобы далее выбрать температуру из словаря
    if msg.text == temp[0]:
        temp_now = 2
    elif msg.text == temp[1]:
        temp_now = 1

    await Test.st_temp.set()
    await bot.send_message(msg.from_user.id, 'Когда была произведена поверка?', reply_markup=markup_date)


@dp.message_handler(state=Test.st_temp)
async def temp_message(msg: types.Message):
        global date_now

        if msg.text == dates[0]:
            date_now = date.today()

        if msg.text == dates[1]:
            '''Макар, доделывай ручной ввод данных'''
            pass

        add_line(preparation(gos_number, temp_now, date_now), line, 1)







if __name__ == '__main__':
    executor.start_polling(dp)
