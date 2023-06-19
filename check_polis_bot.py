from polis_request import get_polis_from_maks_exel, get_polis_from_rmis_exel, \
    get_polis_from_foms_exel, used_users, get_last_row_column
from xlsx_to_xml_takes_delete import make_xml_for_takes_delete
from config import TOKEN, PATH
import datetime

# Настройки
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters
from telegram import Update

updater = Updater(token=TOKEN)  # Токен API к Telegram
dispatcher = updater.dispatcher


# Создадим декоратор, для отлова исключений
def error_log(x):
    def error_(*args, **kwargs):
        try:
            return x(*args, **kwargs)
        except Exception as e:
            print(f'Ошибка: {e}')
            raise e

    return error_


@error_log
# Обработка команд
def startCommand(bot, update):
    bot.send_message(chat_id=update.message.chat_id, text='Добрый день')
    user_id = update.message.from_user.id
    name_id = update.message.from_user.id
    username = update.message.from_user.username
    name_i = update.message.from_user.first_name
    name_f = update.message.from_user.last_name
    print(id, username, name_i, name_f, update.message.text)
    used_users(name_id, username, name_i, name_f, update.message.text,\
               'not', datetime.datetime.now(), 'not')
    print(name_id, username, name_i, name_f, update.message.text)


@error_log
def textMessage(bot, update):
    # input_file = update.message.document
    response = 'Я на связи ' #  + update.message.text  # формируем текст ответа
    bot.send_message(chat_id=update.message.chat_id, text=response)
    name_id = update.message.from_user.id
    username = update.message.from_user.username
    name_i = update.message.from_user.first_name
    name_f = update.message.from_user.last_name
    print(name_id, username, name_i, name_f, update.message.text)
    used_users(name_id, username, name_i, name_f, update.message.text,\
               'not', datetime.datetime.now(), 'not')


# отправка файла
@error_log
def send_document_handler(bot, update, file):
    # print('senddd')
    name_id = update.message.from_user.id
    username = update.message.from_user.username
    name_i = update.message.from_user.first_name
    name_f = update.message.from_user.last_name
    bot.sendDocument(chat_id=update.message.chat_id, document=open(file, 'rb'))
    print(name_id, username, name_i, name_f, 'sended')
    used_users(name_id, username, name_i, name_f,\
               'not', 'sended' + update.message.document.file_name)


# сохранение файла
@error_log
def get_document_handler(bot, update):
    name_id = update.message.from_user.id
    username = update.message.from_user.username
    name_i = update.message.from_user.first_name
    name_f = update.message.from_user.last_name

    file = bot.getFile(update.message.document.file_id)
    print("file_name: " + str(update.message.document.file_name))
    file.download(str(update.message.document.file_name))
    rec_last_row, rec_last_column = get_last_row_column(str(update.\
                                        message.document.file_name))
    if str(update.message.document.file_name).split('.')[-1] == 'xlsx':
        if rec_last_column == 8:  # Не актуально  # если так, файл от МАКС
            # Список из полученного файла отправляем на запрос
            # get_polis_from_maks_exel(str(update.message.document.file_name))
            used_users(name_id, username, name_i, name_f, 'not', 'received_' +\
                       update.message.document.file_name, datetime.datetime.now(),\
                       str(rec_last_row) + '*' + str(rec_last_column))
            # Полученный файл отправляем адресату
            bot.sendDocument(chat_id=update.message.chat_id, document=open(PATH +\
                       'checked_' + update.message.document.file_name, 'rb'))
            sent_last_row, sent_last_column = get_last_row_column(PATH + 'checked_' +
                        str(update.message.document.file_name))
            print(name_id, username, name_i, name_f, 'received_' +\
                  update.message.document.file_name, datetime.datetime.now())
            used_users(name_id, username, name_i, name_f, 'not', 'sent_' +\
                        update.message.document.file_name, datetime.datetime.now(),\
                        str(sent_last_row) + '*' + str(sent_last_column))

        elif rec_last_column == 3:  #  Не актуально # если так, файл из РМИС
            # get_polis_from_rmis_exel(str(update.message.document.file_name))
            # print('col_2')
            used_users(name_id, username, name_i, name_f, 'not', 'received_' +\
                        update.message.document.file_name, datetime.datetime.now(),\
                        str(rec_last_row) + '*' + str(rec_last_column))
            bot.sendDocument(chat_id=update.message.chat_id, document=open(PATH +\
                        'checked_' + update.message.document.file_name, 'rb'))
            sent_last_row, sent_last_column = get_last_row_column(PATH + 'checked_' +\
                        str(update.message.document.file_name))
            print(name_id, username, name_i, name_f, 'sent' + update.message.document.file_name,\
                  datetime.datetime.now())
            used_users(name_id, username, name_i, name_f, 'not', 'sent_' +\
                        update.message.document.file_name, datetime.datetime.now(),\
                        str(sent_last_row) + '*' + str(sent_last_column))

        elif rec_last_column == 5:   # если так, файл от ФОМС
            get_polis_from_foms_exel(str(update.message.document.file_name))
            # print('col_2')
            used_users(name_id, username, name_i, name_f, 'not', 'received_' +\
            update.message.document.file_name, datetime.datetime.now(),\
            str(rec_last_row) + '*' + str(rec_last_column))
            bot.sendDocument(chat_id=update.message.chat_id,\
            document=open(PATH + 'checked_' + update.message.document.file_name, 'rb'))
            sent_last_row, sent_last_column = get_last_row_column(PATH + 'checked_' +\
                        str(update.message.document.file_name))
            print(name_id, username, name_i, name_f, 'sent' + update.message.document.file_name,\
                        datetime.datetime.now())
            used_users(name_id, username, name_i, name_f, 'not', 'sent_' + update.message.document.file_name,
                        datetime.datetime.now(), str(sent_last_row) + '*' + str(sent_last_column))

        elif rec_last_column == 1:   # если так, файл MEK от ФОМС надо переделать в XML
            make_xml_for_takes_delete(str(update.message.document.file_name))
            used_users(name_id, username, name_i, name_f, 'not', 'received_' +\
                       update.message.document.file_name, datetime.datetime.now(),\
                       str(rec_last_row) + '*' + str(rec_last_column))
            bot.sendDocument(chat_id=update.message.chat_id, document=open(PATH +\
                       'checked_' + (update.message.document.file_name).split('.')[0] + '.xml', 'rb'))
            sent_last_row, sent_last_column = get_last_row_column(update.message.document.file_name)
            print(name_id, username, name_i, name_f, 'sent' + update.message.document.file_name,\
                        datetime.datetime.now())
            used_users(name_id, username, name_i, name_f, 'not', 'sent_' + update.message.document.file_name,
                       datetime.datetime.now(), str(sent_last_row) + '*' + str(sent_last_column))


        else:
            response = 'Откройте файл МЭК от ТФОМС, удалите все столбцы кроме: \n'\
                       'Описание, Полис, ФИО, Дата_рождения, Подр_МО \n' \
                       'сохраните файл в формате *.xlsx \n' \
                       'и отправьте мне, я все проверю и верну вам \n'\
                       "Вот образец файла"
            bot.send_message(chat_id=update.message.chat_id, text=response)
            bot.sendDocument(chat_id=update.message.chat_id, \
                             document=open(PATH + 'simple.png', 'rb'))
            print(name_id, username, name_i, name_f, update.message.document.file_name,\
                  'non_format_file')
            used_users(name_id, username, name_i, name_f, 'not', 'non_format_file' +\
                       update.message.document.file_name, datetime.datetime.now(), '0*0')

# Хендлеры
start_command_handler = CommandHandler('start', startCommand)
text_message_handler = MessageHandler(Filters.text, textMessage)
get_doc_message_handler = MessageHandler(Filters.document, get_document_handler)  ###
send_doc_message_handler = MessageHandler(Filters.document, send_document_handler)  ###

# Добавляем хендлеры в диспетчер
dispatcher.add_handler(start_command_handler)
dispatcher.add_handler(text_message_handler)
dispatcher.add_handler(get_doc_message_handler)  ###
dispatcher.add_handler(send_doc_message_handler)  ###
# Начинаем поиск обновлений
updater.start_polling(clean=True)
# Останавливаем бота, если были нажаты Ctrl + C
updater.idle()
