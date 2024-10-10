import requests
import xml.etree.ElementTree as ET
import json
import csv

# Webhook URL'si
webhook_url = 'https://mapaikea.webhook.office.com/webhookb2/2ef0fe0b-c860-49e0-815b-926cf6094a49@d528e232-ef97-4deb-8d8e-c1761bd80396/IncomingWebhook/d6f0c96af01047fdb37e61751b61c692/091d033f-f752-4723-99a6-8a4e4ed8b4a1/V2xjLyMIaDtLEcXqd6HM9CUE2QB3i5faxO7EZCKqeOXzo1'

# Eski stok durumu dosyasını yükleme
def load_old_stock_data(file_path):
    old_data = {}
    try:
        with open(file_path, mode='r', newline='', encoding='utf-8') as file:
            csv_reader = csv.reader(file)
            for row in csv_reader:
                old_data[row[0]] = row[1]  # sprNo ve eski stok durumu
    except FileNotFoundError:
        pass  # Eğer dosya yoksa ilk çalıştırma olabilir, bu durumda boş dönecek
    return old_data

# Yeni stok durumunu CSV'ye kaydetme
def save_new_stock_data(file_path, new_data):
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        csv_writer = csv.writer(file)
        for spr_no, stock in new_data.items():
            csv_writer.writerow([spr_no, stock])

# CSV'den kontrol edilecek ürünleri yükleme
def load_products_from_csv(file_path):
    products = []
    with open(file_path, mode='r', newline='', encoding='utf-8') as file:
        csv_reader = csv.reader(file)
        for row in csv_reader:
            products.append(row[0])  # Sadece sprNo'ları alıyoruz
    return products

# Yeni stok durumunu kontrol etme
def check_stock(product_csv, old_stock_csv):
    url = "https://cdn.ikea.com.tr/urunler-xml/products.xml"
    response = requests.get(url)
    xml_content = response.content
    root = ET.fromstring(xml_content)

    product_ids = load_products_from_csv(product_csv)
    old_stock_data = load_old_stock_data(old_stock_csv)
    new_stock_data = {}

    for spr_no in product_ids:
        item_found = False
        for item in root.findall('item'):
            xml_spr_no = item.find('sprNo').text
            if spr_no == xml_spr_no:
                item_found = True
                stock = item.find('stock').text
                title = item.find('title').text.replace('IKEA ', '')  # "IKEA " kelimesini çıkarma

                # Yeni stok durumunu kaydediyoruz
                new_stock_data[spr_no] = stock

                # Eğer stok durumu değişmişse bildirim gönder
                if spr_no in old_stock_data and old_stock_data[spr_no] != stock:
                    if stock == 'Mevcut':
                        send_notification(spr_no, title, stock, '🟢', 'Stoğa Geldi')
                    elif stock == 'Stok Yok':
                        send_notification(spr_no, title, stock, '🔴', 'Stoğu Bitti')
                break

        # Eğer ürün XML'de bulunamıyorsa ve daha önce mevcutsa "Sitede Kapandı" bildirimi gönder
        if not item_found and spr_no in old_stock_data and old_stock_data[spr_no] == 'Mevcut':
            send_notification(spr_no, 'Sitede Kapandı', '⚪', '⚪', 'Sitede Kapandı')

    # Yeni stok durumlarını kaydediyoruz
    save_new_stock_data(old_stock_csv, new_stock_data)

# Microsoft Teams'e mesaj gönderme
def send_notification(spr_no, title, stock, emoji, message):
    formatted_title = format_title(title)  # Başlığı bold yapmak için formatlıyoruz
    message_data = {
        "text": f"{emoji} {message}: {spr_no} - {formatted_title}"
    }

    response = requests.post(webhook_url, data=json.dumps(message_data), headers={'Content-Type': 'application/json'})

    if response.status_code == 200:
        print(f'Mesaj gönderildi: {message} - {spr_no}')
    else:
        print(f'Bildirim gönderilirken hata oluştu: {response.status_code}')

# Başlığı bold yapmak için formatlama (taksim işareti desteği ile)
def format_title(title):
    # Taksim işaretini kontrol ediyoruz
    if '/' in title:
        parts = title.split('/')
        # Taksim işaretinden önce ve sonra gelen kısmı bold yapıyoruz
        parts[0] = f"**{parts[0]}**"
        parts[1] = f"**{parts[1]}**"
        return '/'.join(parts)
    else:
        # Eğer taksim işareti yoksa, sadece ilk kelimeyi bold yap
        parts = title.split(' ')
        parts[0] = f"**{parts[0]}**"
        return ' '.join(parts)

# Stok durumunu kontrol et
check_stock('urun_listesi.csv', 'eski_stok_durumu.csv')  # CSV dosya yollarını buraya ekleyin
