import requests
from bs4 import BeautifulSoup as BS
import openpyxl

def get_html(url):
    response = requests.get(url)
    if response.status_code == 200:
        return response.text
    return None

def get_links(html):
    soup = BS(html, 'html.parser')
    links = []
    content_wrapper = soup.find('div', class_='content-wrapper')
    main_content = content_wrapper.find('div', class_='main-content')
    posts = main_content.find_all('div', class_='listings-wrapper')
    for post in posts:
        title = post.find('div', class_='c-name')
        price = post.find('div', class_='price-dollar').text.strip()
        address = post.find('div', class_='address').text.strip()
        link = title.find('a').get('href')
        full_link = 'https://house.kg' + link
        links.append(full_link)
    return links


def get_posts(html):
    soup = BS(html, 'html.parser')
    content_wrapper = soup.find('div', class_='content-wrapper')
    title = content_wrapper.find('div', class_='c-name').find('h1').text.strip()
    price = content_wrapper.find('div', class_='price-dollar').text.strip()
    address = content_wrapper.find('span', class_='address').text.strip()
    square = content_wrapper.find('div', class_='left').find('h1').text.strip()
    description = content_wrapper.find('div', class_='description').find('p').text.strip()
        
    data = {
        'title': title,
        'price': price,
        'address': address,
        'square': square,
        'description': description[0:100],
    }
    return data

def get_last_page(html):
    soup = BS(html, 'html.parser')
    content_wrapper = soup.find('div', class_='content-wrapper')
    pagination = content_wrapper.find('div', class_='pagination')
    li = pagination.find('li', class_='page-item active')
    last_page = li.find('a').get('data-page')
    return int(last_page)


def save_to_excel(data):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet['A1'] = 'Название'
    sheet['B1'] = 'Цена'
    sheet['C1'] = 'Адрес'
    sheet['D1'] = 'Площадь'
    sheet['E1'] = 'Описание'

    
    for i,item in enumerate(data,2):
        sheet[f'A{i}'] = item['title']
        sheet[f'B{i}'] = item['price']
        sheet[f'C{i}'] = item['address']
        sheet[f'D{i}'] = item['square']
        sheet[f'E{i}'] = item['description']
        
    wb.save('products.xlsx')

def main():
    URL = 'https://www.house.kg/kupit-kvartiru?page=1'
    html = get_html(URL)
    last_page = get_last_page(html)

    for i in range(1, last_page):
        page_url = f'{URL}?page={i}'
        page = get_html(page_url)
        links = get_links(page)
        data = []
        for link in links:
            detail_html = get_html(link)
            data.append(get_posts(detail_html))
        save_to_excel(data)

if __name__ == '__main__':
    main()


