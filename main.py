from flask import Flask, request, jsonify
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

app = Flask(__name__)

# Налаштування Google Sheets API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('E:\\api-exchange-rates-435414-daa8b542484a.json', scope)
client = gspread.authorize(creds)
spreadsheet = client.open('Exchange Rates')
worksheet = spreadsheet.Rates  # Вкажіть потрібний лист


# Функція для отримання та запису курсу валют за вказаними датами
def update_exchange_rates(update_from=None, update_to=None):
    if update_from is None or update_to is None:
        update_from = update_to = datetime.now().strftime('%Y-%m-%d')

    api_url = f"https://bank.gov.ua/NBUStatService/v1/statdirectory/exchange?date={update_from}&json"
    response = requests.get(api_url)
    data = response.json()

    # Очищуємо Google Sheet перед записом нових даних
    worksheet.clear()

    # Додаємо заголовки
    worksheet.append_row(['Code', 'Name', 'Rate', 'Date'])

    # Записуємо курси валют
    for item in data:
        if update_from <= item['exchangedate'] <= update_to:
            worksheet.append_row([item['cc'], item['txt'], item['rate'], item['exchangedate']])


@app.route('/update_rates', methods=['POST'])
def update_rates():
    # Авторизація (можна використовувати базову або токен)
    auth = request.headers.get('Authorization')
    if auth != 'Bearer your_token':
        return jsonify({"error": "Unauthorized"}), 401

    # Отримуємо параметри з POST-запиту
    update_from = request.form.get('update_from')
    update_to = request.form.get('update_to')

    # Викликаємо функцію для оновлення даних
    update_exchange_rates(update_from, update_to)

    return jsonify({"status": "Update successful"})


if __name__ == '__main__':
    app.run(debug=True)
