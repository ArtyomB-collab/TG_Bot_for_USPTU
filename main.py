import time
import telebot
import os
import pythoncom
import win32com.client
import traceback
import telebot.apihelper
import gc
from dotenv import load_dotenv  
from telebot import types
from telebot.types import BotCommand
load_dotenv()
telebot.apihelper.READ_TIMEOUT = 120
telebot.apihelper.CONNECT_TIMEOUT = 120

bot = telebot.TeleBot(os.getenv('BOT_TOKEN'))

bot.set_my_commands([
    BotCommand("start", "Начать работу"),
    BotCommand("help", "Справка"),
    BotCommand("status", "Проверить статус обработки")
])

TEMP_DIR = "temp_files"

user_files = {}  # {user_id: file_path}
MACROS = {
    'check': {
        'name': 'Проверка',
        'macro': 'USPTU_Check.USPTU'
    }
}

os.makedirs(TEMP_DIR, exist_ok=True)

@bot.message_handler(commands=['start'])
def main(message):
    bot.send_message(
    message.chat.id,
    f'''Привет, {message.from_user.first_name}! 
Присылайте .docx файлы для обработки макросом.

Доступные команды:
/start — Начать работу
/help — Справка
/status — Проверить статус обработки'''
)

@bot.message_handler(commands=['help'])
def help_command(message):
    help_text = """
Доступные команды:
/start - Начать работу
/help - Эта справка
/status - Проверить статус обработки

Просто пришлите .docx файл, и бот запустит на нем макрос
    
Ограничения:
- Только файлы .docx
- Максимальный размер: 20 МБ
- Обработка занимает 1-5 минут
    """
    bot.send_message(message.chat.id, help_text)

def get_unique_filepath(directory, filename):
    base_name, ext = os.path.splitext(filename)
    candidate = os.path.join(directory, filename)
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{base_name} ({counter}){ext}")
        counter += 1
    return candidate

@bot.message_handler(content_types=['document'])
def handler_document(message):
    if not message.document.file_name.endswith('.docx'):
        bot.send_message(message.chat.id, "Только .docx файлы")
        return
    
    if message.document.file_size > 20 * 1024 * 1024:
        bot.send_message(message.chat.id, "Файл больше 20 МБ")
        return

    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)

    input_path = get_unique_filepath(
        TEMP_DIR,
        f"{message.chat.id}_{message.document.file_name}"
    )

    with open(input_path, 'wb') as f:
        f.write(downloaded_file)

    print(f"[DEBUG] Файл сохранён: {input_path}")

    user_files[message.chat.id] = input_path

    markup = types.InlineKeyboardMarkup(row_width=2)

    markup.add(
        types.InlineKeyboardButton(
            "Да",
            callback_data="confirm_yes"
        ),
        types.InlineKeyboardButton(
            "Нет",
            callback_data="confirm_no"
        )
    )

    bot.send_message(
        message.chat.id,
        f"Вы отправили файл:\n{message.document.file_name}\n\nХотите его проверить?",
        reply_markup=markup
    )
@bot.callback_query_handler(
    func=lambda call: call.data in ["confirm_yes", "confirm_no"]
)
def process_confirmation(call):
    user_id = call.message.chat.id
    if user_id not in user_files:
        bot.answer_callback_query(
            call.id,
            "Файл не найден"
        )
        return
    input_path = user_files[user_id]
    if call.data == "confirm_no":
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
            del user_files[user_id]
        except Exception as e:
            print(e)
        bot.edit_message_text(
            "Файл отменён. Можете отправить другой .docx",
            user_id,
            call.message.message_id
        )
        bot.answer_callback_query(call.id)
        return
    output_path = input_path.replace(
        ".docx",
        "_processed.docx"
    )

    bot.edit_message_text(
        "Файл подтверждён. Начинаю обработку...",
        user_id,
        call.message.message_id
    )

    bot.answer_callback_query(call.id)

    result = process_docx_with_macro(
        input_path,
        output_path,
        MACROS["check"]["macro"]
    )

    if result and os.path.exists(result):

        with open(result, 'rb') as f:
            bot.send_document(
                user_id,
                f,
                caption="Проверка завершена"
            )

        bot.edit_message_text(
            "Готово",
            user_id,
            call.message.message_id
        )

        try:
            os.remove(input_path)
            os.remove(result)
            del user_files[user_id]

        except Exception as e:
            print(e)

    else:

        bot.edit_message_text(
            "Ошибка обработки",
            user_id,
            call.message.message_id
        )

def process_docx_with_macro(input_path, output_path, macro_name):
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    
    try:
        doc = word.Documents.Open(os.path.abspath(input_path), ReadOnly=False)
        time.sleep(0.5)
        print(f"[DEBUG] Запускаю макрос {macro_name}...")
        word.Application.Run(macro_name)
        print(f"[DEBUG] Макрос завершён")
        doc.SaveAs(os.path.abspath(output_path))
        print(f"[DEBUG] Сохранено: {output_path}")
        doc.Close()
        return output_path
    except Exception as e:
        print(f"[ERROR] Ошибка: {e}")
        traceback.print_exc()
        return None
    finally:
        word.Quit(SaveChanges=False)
        del word
        gc.collect()
        pythoncom.CoUninitialize()

@bot.message_handler(commands=['status'])
def status_command(message):
    bot.send_message(message.chat.id, "Бот работает нормально. Готов принимать файлы.")

@bot.message_handler(func=lambda message: True)
def handle_all(message):
    if message.text:
        bot.send_message(message.chat.id, "Пожалуйста, пришлите .docx файл для обработки")


if __name__ == '__main__':
    print("Бот запущен...")

    bot.remove_webhook()

    while True:
        try:
            bot.infinity_polling(
                timeout=30,
                long_polling_timeout=30,
                skip_pending=True
            )

        except Exception as e:
            print(f"[ERROR] {e}")
            time.sleep(5)