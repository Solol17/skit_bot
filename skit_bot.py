import logging
import os

import telebot
from configurations import geckodriver_path, firefox_location, source_dir, document_path
from first_project import SkitConnector
from dotenv import load_dotenv

load_dotenv()
# Назначаем токен бота для связи с телеграмом
bot = telebot.TeleBot(os.getenv("bot_key"))
# Создание переменных, в которых хранятся данные для авторизации
current_passwd = os.getenv("current_passwd")
current_login = os.getenv("current_login")

# Создание экземпляра класса SkitConnector
skit = SkitConnector(login=current_login, passwd=current_passwd, source=source_dir)

@bot.message_handler(commands=['start', 'help', 'text', 'report'])
def send_welcome(message):
    if message.text == "/start":
        bot.send_message(message.from_user.id, "Доброго времени суток! Начинаю сбор информации по заявкам.")

  #   Если webdriver закрыт, то происходит его открытие + авторизация на сайте
    if not skit.driver:
        skit.open_driver(geckodriver_path, firefox_location)
        skit.get_authorization()
    # Обращение к методу, создающий словарь из необходимых данных
    skit.read_source()

    if message.text == "/start":
        logging.info(f"Получено сообщение /start")
        yield_answer = skit.get_report()
        for answer_list in yield_answer:
            for answer in answer_list:
                bot.send_message(message.from_user.id, answer, parse_mode="Markdown")
        bot.send_message(message.from_user.id, "Это, на данный момент, вся актуальная информация по заявкам ЦОП.")

    if message.text == "/report":
        logging.info(f"Получено сообщение /report")
        bot.send_message(message.from_user.id, "Доброго времени суток! Начинаю подготовку отчёта.")
        skit.writing_docx(document_path)
        bot.send_document(message.chat.id, open(document_path, 'rb'))

@bot.message_handler(func=lambda message: True)
def echo_all(message):
  bot.reply_to(message, message.text)

# bot.infinity_polling()