import telebot
import sqlite3
import openpyxl
import smtplib
import pandas as pd

token = ('7453810145:AAE8AzJamsCbuwAdVAfIjeoE1SnEmhTJ4YY')
bot = telebot.TeleBot(token)

data = pd.read_excel('homework_data.xlsx')
data1 = pd.read_excel('attendance_data.xlsx')


@bot.message_handler(commands=['start'])
def start_message(message):
    bot.send_message(message.chat.id, "Добро пожаловать! Пожалуйста, представьтесь(ФИО)")


def is_teacher_valid(teacher_name):
    return teacher_name in data, data1['ФИО преподавателя'].values


@bot.message_handler(func=lambda message: True)
def handle_message(message):
    teacher_name = message.text.strip()
    if is_teacher_valid(teacher_name):
        bot.reply_to(message, "Вы успешно авторизованы.")
        chat_id = message.chat.id
        check_homework_completion(chat_id)
        check_homework_given(chat_id)
    else:
        bot.reply_to(message, "Ваше имя не найдено в списке преподавателей.")


def check_homework_completion(chat_id):
    total_homework = len(data)
    checked_homework = data[data['Проверенные ДЗ'] == True].shape[0]

    completion_rate = (checked_homework / total_homework) * 100
    if completion_rate >= 75:
        return
    bot.send_message(chat_id,
                     f"Внимание! Процент проверенных домашних заданий ниже 75% ({completion_rate:.2f}%).")


def check_homework_given(chat_id):
    total_given_homework = len(data)
    given_homework = data[data['Выдано ДЗ'] == True].shape[0]

    given_rate = (given_homework / total_given_homework) * 100
    if given_rate >= 70:
        return
    bot.send_message(chat_id, f"Внимание! Процент выданного домашнего задания ниже 70% ({given_rate:.2f}%).")


bot.polling()
