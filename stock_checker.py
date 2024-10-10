import gspread
from google.oauth2.service_account import Credentials
import requests
import xml.etree.ElementTree as ET
import json

# Webhook URL'si
webhook_url = 'https://mapaikea.webhook.office.com/webhookb2/2ef0fe0b-c860-49e0-815b-926cf6094a49@d528e232-ef97-4deb-8d8e-c1761bd80396/IncomingWebhook/d6f0c96af01047fdb37e61751b61c692/091d033f-f752-4723-99a6-8a4e4ed8b4a1/V2xjLyMIaDtLEcXqd6HM9CUE2QB3i5faxO7EZCKqeOXzo1'

# Google Sheets'e bağlanmak için kimlik bilgileri
def connect_to_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file('credentials.json', scopes=scope)
    client = gspread.authorize(creds)  # Google Sheets'e bağlanmak için client oluşturuluyor
    sheet = client.open_by_key('1BK8XnyGad3h2OiwlX7Fu2CSKYeI_f_Qm4Wxiknm8gL8').sheet1  # Google Sheets ID'si
    return sheet

# Eski stok verilerini Google Sheets'ten çekme
def load_old_stock_data_from_sheets(sheet):
    old_data = {}
    data = sheet.get_all_records()  # Tüm verileri alıyoruz
    for row in data:
        old_data[row['SPR No']] = row['Stock Status']
    return old_data

# Yeni stok verilerini Google Sheets'e toplu olarak yazma
def save_new_stock_data_to_sheets(sheet, new_data):
    rows = []
    for spr_no, stock_status in new_data.items():
        rows.append([spr_no, stock_status])  # SPR No ve Stock Status her satıra ekleniyor

    if len(rows) > 0:
        # Tüm veriyi toplu olarak yazmak için güncelleniyor
        sheet.update(f'A2:B{len(rows) + 1}', rows)

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

    old_stock_data = load_old_stock_data_from_sheets(sheet)  # Eski stok verilerini alıyoruz
    new_stock_data = {}
    messages = []

    for spr_no in old_stock_data:  # Tüm eski stok verilerini dolaşıyoruz
        item_found = False
        for item in root.findall('item'):
            xml_spr_no = item.find('sprNo').text
            if spr_no == xml_spr_no:
                item_found = True
                stock = item.find('stock').text
                title = item.find('title').text.replace('IKEA ', '')  # IKEA ibaresini kaldırıyoruz

                new_stock_data[spr_no] = stock  # Yeni stok verisini kaydediyoruz

                # Eğer stok durumu değişmişse bildirim hazırlıyoruz
                if old_stock_data[spr_no] != stock:
                    if stock == 'Mevcut':
                        messages.append(send_notification(spr_no, title, stock, '🟢', 'Stoğa geldi'))
                    elif stock == 'Stok Yok':
                        messages.append(send_notification(spr_no, title, stock, '🔴', 'Stoğu bitti'))

        # Eğer ürün XML'de bulunamıyorsa ve daha önce mevcutsa "Sitede kapatıldı" bildirimi gönderiyoruz
        if not item_found and old_stock_data[spr_no] == 'Mevcut':
            messages.append(send_notification(spr_no, "Ürün bulunamadı", '⚪', '⚪', 'Sitede kapatıldı'))

    # Yeni stok verilerini Google Sheets'e yaz
    save_new_stock_data_to_sheets(sheet, new_stock_data)

    if messages:
        send_combined_message(messages)

# Microsoft Teams'e toplu mesaj gönderme
def send_combined_message(messages):
    combined_message = '  \n'.join(messages)  # Mesajları alt alta yazıyoruz
    message_data = {"text": combined_message}

    response = requests.post(webhook_url, data=json.dumps(message_data), headers={'Content-Type': 'application/json'})

    if response.status_code == 200:
        print('Mesajlar gönderildi')
    else:
        print(f'Bildirim gönderilirken hata oluştu: {response.status_code}')

if __name__ == "__main__":
    sheet = connect_to_sheets()  # Google Sheets'e bağlanıyoruz
    check_stock(sheet)  # Stok kontrolünü başlatıyoruz
