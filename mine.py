import logging
import pandas as pd
from telegram import (
    Update,
    ReplyKeyboardMarkup,
    KeyboardButton,
    ReplyKeyboardRemove,
)
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)
import os

# Включаем логирование
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO  # Для более подробного логирования можно установить DEBUG
)
logger = logging.getLogger(__name__)

# Определяем состояния для ConversationHandler
REGISTER, MAIN_MENU, ENTER_PERIOD, ENTER_IMPORTANCE = range(4)

# Идентификатор администратора
ADMIN_ID = 461549398  # Замените на ваш Telegram ID

# Путь к Excel файлу
EXCEL_FILE = 'data.xlsx'

# Функция для инициализации Excel файла, если он не существует
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            # Создаём пустой DataFrame для пользователей
            df_users = pd.DataFrame(columns=['telegram_id', 'fio'])
            df_users.to_excel(writer, sheet_name='Users', index=False)

            # Создаём пустой DataFrame для данных
            df_data = pd.DataFrame(columns=['telegram_id', 'fio', 'period', 'importance'])
            df_data.to_excel(writer, sheet_name='Data', index=False)

# Функция для загрузки данных из Excel
def load_data():
    df_users = pd.read_excel(EXCEL_FILE, sheet_name='Users')
    df_data = pd.read_excel(EXCEL_FILE, sheet_name='Data')
    return df_users, df_data

# Функция для сохранения данных в Excel
def save_data(df_users, df_data):
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='w') as writer:
        df_users.to_excel(writer, sheet_name='Users', index=False)
        df_data.to_excel(writer, sheet_name='Data', index=False)

# Функция для создания клавиатуры главного меню
def main_menu_keyboard(user_id):
    buttons = [
        [KeyboardButton("Внести данные")],
        [KeyboardButton("Просмотреть внесённое")]
    ]

    # Добавляем кнопку для администратора
    if user_id == ADMIN_ID:
        buttons.append([KeyboardButton("Скачать таблицу")])

    reply_markup = ReplyKeyboardMarkup(buttons, resize_keyboard=True)
    return reply_markup

# Обработка команды /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    init_excel()

    df_users, _ = load_data()

    if user_id in df_users['telegram_id'].values:
        await update.message.reply_text(
            "Вы уже зарегистрированы.",
            reply_markup=main_menu_keyboard(user_id)
        )
        return MAIN_MENU
    else:
        await update.message.reply_text(
            "Добро пожаловать! Пожалуйста, введите ваше ФИО для регистрации:",
            reply_markup=ReplyKeyboardRemove()
        )
        return REGISTER

# Обработчик регистрации
async def register(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    fio = update.message.text.strip()

    if not fio:
        await update.message.reply_text(
            "ФИО не может быть пустым. Пожалуйста, введите ваше ФИО:"
        )
        return REGISTER

    df_users, df_data = load_data()

    if user_id in df_users['telegram_id'].values:
        await update.message.reply_text(
            "Вы уже зарегистрированы.",
            reply_markup=main_menu_keyboard(user_id)
        )
        return MAIN_MENU

    # Добавляем нового пользователя
    new_user = pd.DataFrame({
        'telegram_id': [user_id],
        'fio': [fio]
    })
    df_users = pd.concat([df_users, new_user], ignore_index=True)

    save_data(df_users, df_data)

    await update.message.reply_text(
        "Регистрация прошла успешно!",
        reply_markup=main_menu_keyboard(user_id)
    )
    return MAIN_MENU

# Обработчик главного меню внутри ConversationHandler

async def main_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text

    if text == "Внести данные":
        await update.message.reply_text(
            "Укажите период отпуска:",
            reply_markup=ReplyKeyboardRemove()
        )
        return ENTER_PERIOD
    elif text == "Просмотреть внесённое":
        await view_data(update, context)
        return MAIN_MENU
    elif text == "Скачать таблицу":
        if user_id == ADMIN_ID:
            await send_excel(update, context)
        else:
            await update.message.reply_text("У вас нет доступа к этой команде.", reply_markup=main_menu_keyboard(user_id))
        return MAIN_MENU
    else:
        await update.message.reply_text("Пожалуйста, выберите опцию из меню.")
        return MAIN_MENU

# Обработчик ввода периода отпуска
async def enter_period(update: Update, context: ContextTypes.DEFAULT_TYPE):
    period = update.message.text.strip()
    if not period:
        await update.message.reply_text("Период отпуска не может быть пустым. Введите период отпуска:")
        return ENTER_PERIOD

    context.user_data['period'] = period
    await update.message.reply_text(
        "Опишите, насколько важен отпуск в этот период:",
        reply_markup=ReplyKeyboardRemove()
    )
    return ENTER_IMPORTANCE

# Обработчик ввода важности отпуска
async def enter_importance(update: Update, context: ContextTypes.DEFAULT_TYPE):
    importance = update.message.text.strip()
    if not importance:
        await update.message.reply_text("Описание важности отпуска не может быть пустым. Опишите важность:")
        return ENTER_IMPORTANCE

    context.user_data['importance'] = importance
    user_id = update.effective_user.id

    df_users, df_data = load_data()

    user_row = df_users[df_users['telegram_id'] == user_id]
    if user_row.empty:
        await update.message.reply_text("Ошибка: Пользователь не зарегистрирован. Используйте /start для регистрации.")
        return ConversationHandler.END

    fio = user_row.iloc[0]['fio']
    period = context.user_data['period']
    importance = context.user_data['importance']

    # Добавляем новую запись данных
    new_data = pd.DataFrame({
        'telegram_id': [user_id],
        'fio': [fio],
        'period': [period],
        'importance': [importance]
    })
    df_data = pd.concat([df_data, new_data], ignore_index=True)

    save_data(df_users, df_data)

    await update.message.reply_text(
        "Ваши данные успешно сохранены!",
        reply_markup=main_menu_keyboard(user_id)
    )
    return MAIN_MENU

# Функция для просмотра внесённых данных пользователя
async def view_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    df_users, df_data = load_data()

    user_row = df_users[df_users['telegram_id'] == user_id]
    if user_row.empty:
        await update.message.reply_text(
            "Вы не зарегистрированы. Пожалуйста, используйте /start для регистрации."
        )
        return

    fio = user_row.iloc[0]['fio']
    user_data = df_data[df_data['telegram_id'] == user_id]

    if user_data.empty:
        await update.message.reply_text(
            "У вас нет внесённых данных.",
            reply_markup=main_menu_keyboard(user_id)
        )
        return

    message = f"Ваши данные, {fio}:\n"
    for idx, row in user_data.iterrows():
        message += f"\nПериод отпуска: {row['period']}\nВажность: {row['importance']}\n"

    await update.message.reply_text(
        message,
        reply_markup=main_menu_keyboard(user_id)
    )

# Функция для отправки Excel-файла администратору
async def send_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != ADMIN_ID:
        await update.message.reply_text("У вас нет доступа к этой команде.")
        return

    if not os.path.exists(EXCEL_FILE):
        await update.message.reply_text("Файл данных не найден.")
        return

    with open(EXCEL_FILE, 'rb') as file:

        await update.message.reply_document(document=file)
        await update.message.reply_text("Файл отправлен.", reply_markup=main_menu_keyboard(user_id))

        # Обработчик неизвестных команд

async def unknown(update: Update, context: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text("Извините, я не понимаю эту команду. Используйте /start для начала.")

def main():
        # Создаём приложение бота
        application = ApplicationBuilder().token('8072085864:AAG564f5XpnTISlbc2O5uMTntOE7Ajmgp74').build()  # Замените на ваш токен

        # Инициализируем Excel файл
        init_excel()

        # Создаём ConversationHandler для регистрации и взаимодействия с меню
        conv_handler = ConversationHandler(
            entry_points=[CommandHandler('start', start)],
            states={
                REGISTER: [MessageHandler(filters.TEXT & ~filters.COMMAND, register)],
                MAIN_MENU: [
                    MessageHandler(filters.Regex('^(Внести данные|Просмотреть внесённое|Скачать таблицу)$'),
                                   main_menu_handler)
                ],
                ENTER_PERIOD: [MessageHandler(filters.TEXT & ~filters.COMMAND, enter_period)],
                ENTER_IMPORTANCE: [MessageHandler(filters.TEXT & ~filters.COMMAND, enter_importance)],
            },
            fallbacks=[CommandHandler('start', start)]
        )

        # Добавляем обработчики
        application.add_handler(conv_handler)
        application.add_handler(MessageHandler(filters.COMMAND, unknown))

        # Запускаем бота
        application.run_polling()

if __name__ == '__main__':
        main()
