import json
import locale
import os
import time
import telebot
from telebot import types
import requests
import pandas as pd
from PIL import Image, ImageDraw, ImageGrab, ImageFont
import win32com.client
import pythoncom

bot = telebot.TeleBot('')  # token to access HTTP API
openweather_API = '' # token to access HTTP API for site: https://openweathermap.org/api
locale.setlocale(locale.LC_TIME, 'ru_RU')
logo = Image.open("img/ava.jpg")
logo = logo.resize((60, 60))
font = ImageFont.truetype('font/Certa Sans Medium.otf', 30)


@bot.message_handler(commands=['start', 'main', 'hello'])
def start(message):
    bot.send_message(message.chat.id,
                     f'Привет, {message.from_user.first_name}! Подсказать погоду? Пиши название города, и я отправлю '
                     f'тебе данные.\n\nДля подробного просмотра команд бота используйте /help!')


@bot.message_handler(commands=['help'])
def commands(message):
    bot.send_message(message.chat.id,
                     'Список команд для бота:\n/start - начать работу бота\n/weather - узнать нынешнюю '
                     'погоду\n/forecast - узнать прогноз погоды')


@bot.message_handler(commands=['forecast'])
def forecast(message):
    markup = types.InlineKeyboardMarkup()
    bt1 = types.InlineKeyboardButton('Город', callback_data='CITY_FORECAST')
    bt2 = types.InlineKeyboardButton('Координаты', callback_data='COR_FORECAST')
    markup.row(bt1, bt2)
    bot.reply_to(message, 'Вы хотите узнать прогноз погоды для определённого города или точных координат?',
                 reply_markup=markup)


@bot.message_handler(commands=['weather'])
def weather(message):
    markup = types.InlineKeyboardMarkup()
    bt1 = types.InlineKeyboardButton('Город', callback_data='CITY_WEATHER')
    bt2 = types.InlineKeyboardButton('Координаты', callback_data='COR_WEATHER')
    markup.row(bt1, bt2)
    bot.reply_to(message, 'Вы хотите узнать погоду в данный момент для определённого города или точных координат?',
                 reply_markup=markup)


@bot.callback_query_handler(func=lambda call: True)
def callback_query(call):
    key = str(call.data)
    if 'CITY' in key:
        bot.send_message(call.message.chat.id, 'Введите название города.')
        if 'FORECAST' in key:
            bot.register_next_step_handler(call.message, get_city_forecast)
        elif 'WEATHER' in key:
            bot.register_next_step_handler(call.message, get_city_weather)

    if 'COR' in key:
        bot.send_message(call.message.chat.id, 'Введите координаты (широту и долготу) через пробел, разделитель '
                                               'разрядов в числах - точка.')
        if 'FORECAST' in key:
            bot.register_next_step_handler(call.message, get_cor_forecast)
        elif 'WEATHER' in key:
            bot.register_next_step_handler(call.message, get_cor_weather)


def get_city_forecast(message):
    city = message.text.strip().lower()
    res = requests.get(
        f'https://api.openweathermap.org/data/2.5/forecast?q={city}&appid={openweather_API}&units=metric&lang=ru')
    if res.status_code == 200:
        print(res.text)
        send_forecast(res, message)
        return

    bot.reply_to(message, 'Город указан некорректно!')


def get_cor_forecast(message):
    cor = parse_cor(message.text)
    if len(cor) == 2:
        res = requests.get(
            f'https://api.openweathermap.org/data/2.5/forecast?lat={cor[0]}&lon={cor[1]}&appid={openweather_API}&units=metric&lang=ru')
        print(res.text)
        if res.status_code == 200:
            send_forecast(res, message)
            return

    bot.reply_to(message, 'Координаты указаны некорректно!')


def send_forecast(res, message):
    data = json.loads(res.text)
    # text = make_info_weather(data)
    # img = requests.get(f'https://openweathermap.org/img/wn/{data["weather"][0]["icon"]}@2x.png',
    #                    stream=True).content
    # bot.send_photo(message.chat.id, img, text, reply_to_message_id=message.id)
    img = make_info_forecast(data)
    bot.send_photo(message.chat.id, img, reply_to_message_id=message.id)


def get_city_weather(message):
    city = message.text.strip().lower()
    res = requests.get(
        f'https://api.openweathermap.org/data/2.5/weather?q={city}&appid={openweather_API}&units=metric&lang=ru')
    if res.status_code == 200:
        send_weather(res, message)
        return

    bot.reply_to(message, 'Город указан некорректно!')


def get_cor_weather(message):
    cor = parse_cor(message.text)
    if len(cor) == 2:
        res = requests.get(
            f'https://api.openweathermap.org/data/2.5/weather?lat={cor[0]}&lon={cor[1]}&appid={openweather_API}'
            f'&units=metric&lang=ru')
        if res.status_code == 200:
            send_weather(res, message)
            return

    bot.reply_to(message, 'Координаты указаны некорректно!')


def send_weather(res, message):
    data = json.loads(res.text)
    text = make_info_weather(data)
    img = requests.get(f'https://openweathermap.org/img/wn/{data["weather"][0]["icon"]}@2x.png',
                       stream=True).content
    bot.send_photo(message.chat.id, img, text, reply_to_message_id=message.id)


@bot.message_handler(content_types=['text'])
def get_weather(message):
    city = message.text.strip().lower()
    res = requests.get(
        f'https://api.openweathermap.org/data/2.5/weather?q={city}&appid={openweather_API}&units=metric&lang=ru')
    if res.status_code == 200:
        print(res.text)
        data = json.loads(res.text)
        text = make_info_weather(data)
        img = requests.get(f'https://openweathermap.org/img/wn/{data["weather"][0]["icon"]}@2x.png',
                           stream=True).content
        bot.send_photo(message.chat.id, img, text, reply_to_message_id=message.id)
    else:
        bot.reply_to(message, 'Если вы просто так пишете мне, то не стоит, я в любом случае буду думать, '
                              'что вы прислали мне название города. И в данном случае город указан некорректно!')


def make_info_weather(data):
    local_time = time.gmtime(time.time() + int(data['timezone']))
    wind = make_info_wind(data["wind"])
    str_wind = f"Ветер: {wind['dir']} {wind['speed']} м/с\n"
    if {wind['gust']} != '—':
        str_wind += f"Порывы ветра: {wind['gust']} м/с\n"

    text = f"В {time.strftime('%H:%M', local_time)} по местному времени:\n" \
           + f"Погода: {data['weather'][0]['description'].capitalize()}\n" \
           + f"Температура: {data['main']['temp']} °C\n" \
           + str_wind \
           + f"Температура: {data['main']['temp']} °C\n" \
           + f"Влажность: {data['main']['humidity']}%\n" \
           + f"Атм. давление: {round(data['main']['pressure'] / 1.333)} мм. рт. ст.\n"
    return text


def make_info_wind(wind):
    wind_deg = wind['deg']
    if wind_deg > 337.5 or wind_deg < 22.5:
        wind_direction = 'С\u2B06'
    elif 22.5 <= wind_deg <= 67.5:
        wind_direction = 'С/В\u2197'
    elif 67.5 < wind_deg < 112.5:
        wind_direction = 'В\u27A1'
    elif 112.5 <= wind_deg <= 157.5:
        wind_direction = 'Ю/В\u2198'
    elif 157.5 < wind_deg < 202.5:
        wind_direction = 'Ю\u2B07'
    elif 202.5 <= wind_deg <= 247.5:
        wind_direction = 'Ю/З\u2199'
    elif 247.5 < wind_deg < 292.5:
        wind_direction = 'З\u2B05'
    else:
        wind_direction = 'С/З\u2196'

    wind_dic = {'dir': str(wind_direction), 'speed': str(wind['speed']), 'gust': '—'}
    if wind.get("gust"):
        wind_dic['gust'] = str(wind['gust'])

    return wind_dic


def parse_cor(str_cor):
    cor = []
    try:
        cor = list(map(float, str_cor.split()))
    finally:
        return cor


def make_info_forecast(data):
    df = pd.DataFrame(
        {'': ['Погода', 'Температура', 'Направление ветра', 'Скорость ветра', 'Порывы ветра', 'Влажность',
              'Атм. давление']})

    for i in range(1, 8, 2):
        times = data['list'][i]
        local_time = time.strftime('%d %b %H:%M', time.gmtime(times['dt']))
        wind = make_info_wind(times['wind'])
        df_forecast = pd.DataFrame(
            {local_time: [times['weather'][0]['description'].capitalize() + ' ', f"{times['main']['temp']} °C",
                          wind['dir'], wind['speed'] + ' м/с', wind['gust'] + ' м/с',
                          f"{times['main']['humidity']}%", f"{round(times['main']['pressure'] / 1.333)} мм. рт. ст."]})
        df = pd.concat([
            df,
            df_forecast
        ], axis=1)

    writer = pd.ExcelWriter('forecast.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='new_forecast', index=False)

    # Получаем объект xlsxwriter workbook и worksheet
    workbook = writer.book
    worksheet = writer.sheets['new_forecast']

    # Устанавливаем ширину столбцов
    for col_num, value in enumerate(df.columns.values):
        column_width = max(df[value].astype(str).apply(len).max(), len(value))
        worksheet.set_column(col_num, col_num, column_width)

    # Устанавливаем выравнивание текста
    cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    for row_num in range(1, len(df) + 1):
        worksheet.set_row(row_num, None, cell_format)

    writer.close()
    xlsx_path = os.path.abspath("forecast.xlsx")
    client = win32com.client.Dispatch("Excel.Application", pythoncom.CoInitialize())
    wb = client.Workbooks.Open(xlsx_path)
    ws = wb.Worksheets("new_forecast")
    ws.Range("A1:E8").CopyPicture(Format=2)
    img = ImageGrab.grabclipboard()

    padding_top = logo.height
    padding_color = (255, 255, 255)
    new_height = img.height + padding_top
    new_img = Image.new("RGB", (img.width, new_height), padding_color)
    new_img.paste(img, (0, padding_top))
    new_img.paste(logo, (0, 0))

    draw = ImageDraw.Draw(new_img)
    text = ''
    if data['city']['name'] == '':
        text = f'Погода по координатам {data["city"]["coord"]["lat"]} {data["city"]["coord"]["lon"]}'
    else:
        text = f'Погода в городе {data["city"]["name"]}'
    text_color = (0, 0, 0)  # Цвет текста (в данном случае - белый)
    text_position = (70, 10)  # Позиция текста (координаты X, Y)
    draw.text(text_position, text, fill=text_color, font=font)

    wb.Close()
    client.Quit()
    os.remove("./forecast.xlsx")
    print(df)

    # local_time = time.gmtime(time.time() + int(data['timezone']))
    #
    # text = f"{time.strftime('%d %b %H:%M ', local_time)} по местному времени:\n" \
    #        + f"Погода: {data['weather'][0]['description'].capitalize()}\n" \
    #        + f"Температура: {data['main']['temp']} °C\n" \
    #        + make_info_wind(data["wind"]) \
    #        + f"Влажность: {data['main']['humidity']}%\n" \
    #        + f"Атм. давление: {round(data['main']['pressure'] / 1.333)} мм. рт. ст.\n"
    return new_img


bot.infinity_polling()
