from aiogram import Bot, types
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from states import *
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from data import *
from datetime import date, timedelta
import os.path
from config import TOKEN
from create_excel import create_excel, add_line, preparation

bot = Bot(token=TOKEN)

dp = Dispatcher(bot, storage=MemoryStorage())

###
counter_list = list(counter.keys())

button_1 = KeyboardButton(counter_list[0])
button_2 = KeyboardButton(counter_list[1])
button_3 = KeyboardButton(counter_list[2])
button_4 = KeyboardButton(counter_list[3])
button_5 = KeyboardButton(counter_list[4])
button_6 = KeyboardButton(counter_list[5])
button_old = KeyboardButton('Получить отсчёт за прошедшую неделю')

cancel = 'Сбросить и начать заново'
button_cancel = KeyboardButton(cancel)

markup_1 = ReplyKeyboardMarkup(resize_keyboard=True).row(button_1, button_2, button_3)
markup_1.row(button_4, button_5, button_6)

markup_1.row(button_old)

temp = ['Горячая', 'Холодная']
button_hot = KeyboardButton(temp[0])
button_cold = KeyboardButton(temp[1])
markup_temp = ReplyKeyboardMarkup(resize_keyboard=True).add(button_hot, button_cold).add(button_cancel)

dates = ['Сегодня', 'Ввести дату']
button_today = KeyboardButton(dates[0])
button_date = KeyboardButton(dates[1])
markup_date = ReplyKeyboardMarkup(resize_keyboard=True).add(button_today, button_date).add(button_cancel)
###


# learn answer on start
now = date.today()


if date.weekday(now) == 3:
    if not os.path.exists(f'poverka/poverka{now}.xlsx'):
        create_excel(str(now))
    old = now - timedelta(weeks=1)
else:
    while date.weekday(now) != 3:
        now -= timedelta(days=1)

    old = now - timedelta(weeks=1)
    if not os.path.exists(f'poverka/poverka{now}.xlsx'):
        create_excel(str(now))


@dp.message_handler(commands=['start'])
async def process_start_command(message: types.Message):

    await message.reply('Привет, выбери счётчик из списка. ', reply_markup=markup_1, reply=False)
    await Test.start.set()


@dp.message_handler(state=Test.start)
async def echo_message(msg: types.Message):
    if msg.text in counter_list:
        global gos_number    #запоминиаем ключ для словаря со счётчиками
        gos_number = msg.text
        await bot.send_message(msg.from_user.id, 'Горячая или холодная?', reply_markup=markup_temp)
        await Test.st_count.set()
    elif msg.text == 'Получить отсчёт за прошедшую неделю':
        with open(f'poverka/poverka{now}.xlsx', 'rb') as file:
            await bot.send_document(msg.from_user.id, file, caption='Держи, друг')
    else:
        await bot.send_message(msg.from_user.id, 'Счётчик не в списке. Начните заново. /start', reply_markup=markup_1)
        await Test.start.set()


@dp.message_handler(state=Test.st_count)
async def temp_message(msg: types.Message):
    global temp_now #переменная, чтобы далее выбрать температуру из словаря
    if msg.text == cancel:
        await Test.start.set()
        await bot.send_message(msg.from_user.id, 'Попробуем ещё раз?', reply_markup=markup_1)
    if msg.text == temp[0]:
        temp_now = 1
        await Test.st_temp.set()
        await bot.send_message(msg.from_user.id, 'Когда была произведена поверка? Введи в формате: YYYY-MM-DD', reply_markup=markup_date)
    elif msg.text == temp[1]:
        temp_now = 2
        await Test.st_temp.set()
        await bot.send_message(msg.from_user.id, 'Когда была произведена поверка? Введи в формате: YYYY-MM-DD', reply_markup=markup_date)


@dp.message_handler(state=Test.st_temp)
async def date_message(msg: types.Message):
    if msg.text == cancel:
        await Test.start.set()
        await bot.send_message(msg.from_user.id, 'Попробуем ещё раз?', reply_markup=markup_1)
    global date_now
    if msg.text == dates[0]:
        date_now = date.today()
        add_line(preparation(gos_number, temp_now, date_now), now)
        await Test.start.set()
        await bot.send_message(msg.from_user.id, 'Занесено в таблицу успешно. Можем повторить.', reply_markup=markup_1)
    if msg.text == dates[1]:
        await bot.send_message(msg.from_user.id, 'Когда была произведена поверка? Введи в формате: YYYY-MM-DD', reply_markup=markup_date)
        await Test.wait.set()


@dp.message_handler(state=Test.wait)
async def other_date_message(msg: types.Message):
    if msg.text == cancel:
        await Test.start.set()
        await bot.send_message(msg.from_user.id, 'Попробуем ещё раз?', reply_markup=markup_1)
    lst = msg.text.split('-')
    try:
        date_now = date(year=int(lst[0]), month=int(lst[1]), day=int(lst[2]))
        add_line(preparation(gos_number, temp_now, date_now), now)
        await bot.send_message(msg.from_user.id, 'Занесено в таблицу успешно. Можем повторить.', reply_markup=markup_1)
        await Test.start.set()

    except ValueError:
        await bot.send_message(msg.from_user.id, 'Ты ошибся номером, друг. Введи ещё раз в формате: YYYY-MM-DD')
    except IndexError:
        await bot.send_message(msg.from_user.id, 'Ты ошибся номером, друг. Введи ещё раз в формате: YYYY-MM-DD')


if __name__ == '__main__':
    executor.start_polling(dp)
