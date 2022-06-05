from aiogram.types import ReplyKeyboardMarkup, KeyboardButton

btnMain = KeyboardButton('Начать общение')
btnConf = KeyboardButton('Поехали!')
CreateProfo = KeyboardButton('Создать профиль')
btnNO = KeyboardButton('Нет, не получилось')
btnYes = KeyboardButton('Да, мы связались')

create = ReplyKeyboardMarkup(resize_keyboard = True).add(CreateProfo)
mainMenu = ReplyKeyboardMarkup(resize_keyboard = True).add(btnMain)
confirm = ReplyKeyboardMarkup(resize_keyboard = True).add(btnConf)
yes_no = ReplyKeyboardMarkup(resize_keyboard = True).add(btnNO,btnYes)


