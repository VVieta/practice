import telebot
import pandas as pd
from sec import sec

token = sec.get('BOT_API_TOKEN')
bot = telebot.TeleBot(token)

homework_data = pd.read_excel('homework_data.xlsx', header=[0, 3])
attendance_data = pd.read_excel('attendance_data.xlsx')


@bot.message_handler(commands=['start'])
def start_message(message):
    bot.send_message(message.chat.id, "Добро пожаловать! Пожалуйста, представьтесь (ФИО)")


def is_teacher_valid(teacher_name):
    return teacher_name in homework_data['ФИО преподавателя'].values or teacher_name in attendance_data[
        'ФИО преподавателя'].values


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
    if ('Кол-во пар', 'Проверено') not in homework_data.columns:
        bot.send_message(chat_id, "столбцы 'Кол-во пар', 'Проверено' в файле не распознаны")
        return

    total_homework = homework_data[('Кол-во пар', 'Выдано')].sum()
    checked_homework = homework_data[('Кол-во пар', 'Проверено')].sum()

    completion_rate = (checked_homework / total_homework) * 100 if total_homework > 0 else 0
    if completion_rate < 75:
        bot.send_message(chat_id,
                         f"Внимание! Процент проверенных домашних заданий ниже 75% ({completion_rate:.2f}%).")


def check_homework_given(chat_id):
    if ('Кол-во пар', 'Выдано') not in homework_data.columns:
        bot.send_message(chat_id, "столбцы 'Кол-во пар', 'Выдано' в файле не распознаны")
        return

    total_given_homework = homework_data[('Кол-во пар', 'Выдано')].sum()
    given_rate = (total_given_homework / len(homework_data)) * 100 if len(homework_data) > 0 else 0
    if given_rate < 70:
        bot.send_message(chat_id,
                         f"Внимание! Процент выданного домашнего задания ниже 70% ({given_rate:.2f}%).")


bot.polling()
