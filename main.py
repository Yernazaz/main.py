import asyncio
from contextlib import suppress
import schedule
import pandas as pd
import xlsxwriter
import pyodbc
from config import BOT_TOKEN
import sqlite3
import logging
import datetime
import aiogram.utils.markdown as md
from aiogram import Bot, Dispatcher, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Text
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.utils import executor
from aiogram import types
from aiogram.utils.exceptions import (MessageToEditNotFound, MessageCantBeEdited, MessageCantBeDeleted,
                                      MessageToDeleteNotFound)
import markups as nav

from xlsxwriter.workbook import Workbook

workbook = Workbook('output_users.xlsx')
worksheet = workbook.add_worksheet()

workbook1 = Workbook('output_meetings.xlsx')
worksheet1 = workbook1.add_worksheet()

logging.basicConfig(level=logging.INFO)
bot = Bot(token=BOT_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)


# States
class Form(StatesGroup):
    name = State()
    city = State()
    work = State()
    link = State()
    hobby = State()


class Check(StatesGroup):
    date = State()


users_list = []


@dp.message_handler(commands=["start"])
async def start_command(message: types.Message):
    await message.reply(
        "Приветствую\nКаждую неделю я буду предлагать тебе для встречи интересного человека, случайно выбранного среди других участников сообщества."
        "Для старта ответь на несколько вопросов и прочитай короткую инструкцию. ", reply_markup=nav.create)


# You can use state '*' if you need to handle all states
@dp.message_handler(state='*', commands='cancel')
@dp.message_handler(Text(equals='cancel', ignore_case=True), state='*')
async def cancel_handler(message: types.Message, state: FSMContext):
    """
    Allow user to cancel any action
    """
    current_state = await state.get_state()
    if current_state is None:
        return

    logging.info('Cancelling state %r', current_state)
    # Cancel state and inform user about it
    await state.finish()
    # And remove keyboard (just in case)
    await message.reply('Отменено.', reply_markup=types.ReplyKeyboardRemove())


@dp.message_handler(commands=['create'])
async def process_prof(message: types.Message):
    await Form.name.set()
    await bot.send_message(message.from_user.id, "☕️Напиши Имя и Фамилию ")


@dp.message_handler(state=Form.name)
async def process_name(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['name'] = message.text

    await Form.next()
    await message.reply("🏬 Напиши свой город и нажми \nОтправить")


@dp.message_handler(state=Form.city)
async def process_city(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['city'] = message.text
    # Update state and data
    await Form.next()
    await message.reply("👩‍💻 Напиши свою должность/роль в компании\nПример: менеджер проектов")


@dp.message_handler(state=Form.work)
async def process_work(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['work'] = message.text
    await Form.next()
    await message.reply("А теперь отправь ссылку на свою социальную сеть")


@dp.message_handler(state=Form.link)
async def process_link(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['link'] = message.text
    await Form.next()
    await message.reply("🦄 Расскажи немного о себе и своих увлечениях")


@dp.message_handler(state=Form.hobby)
async def process_gender(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['hobby'] = message.text
        username = message.from_user.username
        try:
            conn = sqlite3.connect("coffee_random.db")
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO `users` (`full_name`, `city`, `work`, `hobby`, `username`, `tg_username`) VALUES (?, ?, ?, ?, ?, ?)",
                (data['name'],
                 data['city'],
                 data['work'],
                 data['hobby'],
                 data['link'],
                 str(username)
                 ))
            await bot.send_message(
                message.chat.id,
                md.text(
                    md.text('Супер! Твой профиль выглядит так:\nИмя:'
                            , md.code(data['name'])),
                    md.text('Город:', md.code(data['city'])),
                    md.text('Занятие:', data['work']),
                    md.text('О себе:', md.code(data['hobby'])),

                    sep='\n',
                ), reply_markup=nav.mainMenu)
            conn.commit()


        except ValueError:
            await message.reply("smth wrong")

        finally:
            if (conn):
                conn.close()

        await state.finish()


@dp.message_handler()
async def bot_message(message: types.Message):
    if message.text == 'Начать общение':
        try:
            conn = sqlite3.connect("coffee_random.db")
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM `users` ORDER BY RANDOM() LIMIT 1")
            result = cursor.fetchall()
            for row in result:
                id1 = row[0]
                name = row[1]
                city = row[2]
                work = row[3]
                hobby = row[4]
                link = row[5]
                tg = row[6]

            await bot.send_message(
                message.chat.id,
                md.text(
                    md.text('Я нашел тебе напарника. Это ', name,
                            '\nГород: ', city,
                            '\nЗанятие:', work,
                            '\nО себе:', hobby,
                            '\nВот его телеграм:', tg,
                            '\n\nНапиши напарнику, и договоритесь о времени встречи или видеозвонка.'
                            '\nВы можете устроить онлайн-коворкинг💻 или запланировать совместный кофе брейк☕'
                            ),
                    sep='\n',
                ), reply_markup=nav.confirm
            )
            cursor.execute("SELECT * FROM `users` WHERE `tg_username`=?", [message.from_user.username], )
            user1 = cursor.fetchall()
            for row1 in user1:
                id0 = row1[0]

            users_list.clear()
            users_list.append(id0)
            users_list.append(id1)
            conn.commit()
            conn.close()

            try:
                timer_seconds = 10
                if timer_seconds <= 0:
                    raise ValueError()
            except (TypeError, ValueError):
                await bot.send_message(chat_id=message.chat.id, text='timer wrong')

            for seconds_left in range(timer_seconds - 1, 0, -1):
                await asyncio.sleep(1)

                if seconds_left == 1:
                    await bot.send_message(chat_id=message.chat.id,
                                           text=f'👋🏻Привет!\nПомнишь, я давал тебе напарника?',
                                           reply_markup=nav.yes_no)
        except:
            await message.reply('smth wrong')
    elif message.text == 'Поехали!':
        await message.reply('Не откладывай, договорись о встрече сразу🙂', reply_markup=types.ReplyKeyboardRemove())
    elif message.text == 'Создать профиль':
        conn = sqlite3.connect("coffee_random.db")
        cursor = conn.cursor()
        cursor.row_factory = lambda cursor, row: row[6]
        cursor.execute("SELECT * FROM `users` WHERE `tg_username`=?", [message.from_user.username], )
        result = cursor.fetchall()
        conn.commit()
        conn.close()

        if not result:
            await Form.name.set()
            await message.reply("☕️Напиши Имя и Фамилию ", reply_markup=types.ReplyKeyboardRemove())
        else:
            await message.reply("У вас уже есть анкета! ", reply_markup=nav.mainMenu)

    elif message.text == 'Нет, не получилось':
        connection = sqlite3.connect("coffee_random.db")
        cursor = connection.cursor()
        cursor.execute("INSERT INTO `meetings` (`user1`,`user2`,`meeted`) VALUES(?,?,?)",
                       (users_list.pop(), users_list.pop(), False))
        connection.commit()
        connection.close()
        await message.reply(f'Спасибо за отзыв!', reply_markup=types.ReplyKeyboardRemove())
    elif message.text == 'Да, мы связались':
        connection = sqlite3.connect("coffee_random.db")
        cursor = connection.cursor()
        cursor.execute("INSERT INTO `meetings` (`user1`,`user2`,`meeted`) VALUES(?,?,?)",
                       (users_list[0], users_list[1], True))
        connection.commit()
        connection.close()
        await message.reply('Спасибо за отзыв!', reply_markup=types.ReplyKeyboardRemove())
    else:
        await message.reply('Неизвестная команда', reply_markup=types.ReplyKeyboardRemove())


if __name__ == '__main__':
    executor.start_polling(dp)
    conn = sqlite3.connect('coffee_random.db')
    c = conn.cursor()
    c.execute("select * from users")
    mysel = c.execute("select * from users ")
    for i, row in enumerate(mysel):
        for j, value in enumerate(row):
            worksheet.write(i, j, value)
    workbook.close()


    c.execute("select * from meetings")
    mysel = c.execute("select * from meetings ")
    for i, row in enumerate(mysel):
        for j, value in enumerate(row):
            worksheet1.write(i, j, value)
    workbook1.close()

