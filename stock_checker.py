import gspread
from google.oauth2.service_account import Credentials
import requests
import xml.etree.ElementTree as ET
import json

# Google Sheets'e bağlanmak için kimlik bilgileri
def connect_to_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file('credentials.json', scopes=scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key('1BK8XnyGad3h2OiwlX7Fu2CSKYeI_f_Qm4Wxiknm8gL8').sheet1  # İlk sayfaya erişiyoruz
    return sheet

# Eski stok verilerini Google Sheets'ten çekme
def load_old_stock_data_from_sheets(sheet):
    old_data = {}
    data = sheet.get_all_records()
    for row in data:
        old_data[row['SPR No']] = row['Stock Status']
    return old_data

# Yeni stok verilerini Google Sheets'e yazma
def save_new_stock_data_to_sheets(sheet, new_data):
    for i, (spr_no, stock_status) in enumerate(new_data.items(), start=2):
        sheet.update_cell(i, 1, spr_no)  # SPR No sütunu
        sheet.update_cell(i, 2, stock_status)  # Stock Status sütunu

# Microsoft Teams'e mesaj oluşturma
def send_notification(spr_no, title, stock, emoji, message):
    formatted_title = format_title(title)
    link = f"[{spr_no}](https://www.ikea.com.tr/p{spr_no})"
    return f"{emoji} {message}: {link} - {formatted_title}"

# Başlığı bold yapmak için formatlama
def format_title(title):
    if '/' in title:
        parts = title.split('/')
        parts[0] = f"**{parts[0]}**"
        parts[1] = f"**{parts[1]}**"
        return '/'.join(parts)
    else:
        parts = title.split(' ', 1)
        parts[0] = f"**{parts[0]}**"
        return ' '.join(parts)

# Stok durumunu kontrol et
def check_stock(sheet):
    url = "https://cdn.ikea.com.tr/urunler-xml/products.xml"
    response = requests.get(url)
    xml_content = response.content
    root = ET.fromstring(xml_content)

    old_stock_data = load_old_stock_data_from_sheets(sheet)
    new_stock_data = {}
    messages = []

    for item in root.findall('item'):
        spr_no = item.find('sprNo').text
        stock = item.find('stock').text
        title = item.find('title').text.replace('IKEA ', '')

        new_stock_data[spr_no] = stock

        if spr_no in old_stock_data and old_stock_data[spr_no] != stock:
            if stock == 'Mevcut':
                messages.append(send_notification(spr_no, title, stock, '🟢', 'Stoğa geldi'))
            elif stock == 'Stok Yok':
                messages.append(send_notification(spr_no, title, stock, '🔴', 'Stoğu bitti'))

    save_new_stock_data_to_sheets(sheet, new_stock_data)

    if messages:
        send_combined_message(messages)

# Microsoft Teams'e toplu mesaj gönderme
def send_combined_message(messages):
    combined_message = '  \n'.join(messages)
    message_data = {"text": combined_message}

    webhook_url = 'https://mapaikea.webhook.office.com/...'
    response = requests.post(webhook_url, data=json.dumps(message_data), headers={'Content-Type': 'application/json'})

    if response.status_code == 200:
        print('Mesajlar gönderildi')
    else:
        print(f'Bildirim gönderilirken hata oluştu: {response.status_code}')

if __name__ == "__main__":
    sheet = connect_to_sheets()
    check_stock(sheet)
