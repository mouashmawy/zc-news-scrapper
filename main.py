import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


def scrapCard(card):
    title = card.find('h3').find('a').text
    link = card.find('h3').find('a')['href']
    full_link = 'https://zewailcity.edu.eg/main/' + link
    time = card.time.em.strong.text
    paragraph = card.find('div', class_='news-text').p.text
    photo = card.img['src'].replace('../', 'https://www.zewailcity.edu.eg/')
    return title, full_link, time, paragraph, photo


def savePageToSheet(pageCards, ws):
    for card in pageCards:
        title, full_link, time, paragraph, photo_link = scrapCard(card)

        row = ws.max_row

        ws.cell(row=row, column=1).value = title
        ws.cell(row=row, column=2).value = time
        ws.cell(row=row, column=3).value = paragraph
        ws.cell(row=row, column=4).value = full_link
        ws.cell(row=row, column=5).value = photo_link




def scrapSite():

    wb = Workbook()
    ws = wb.active

    for page in range(1,2):
        response = requests.get(f'https://www.zewailcity.edu.eg/main/content.php?lang=en&alias=recent_news&page={page}')

        soap = BeautifulSoup(response.text,'lxml')
        pageCards = soap.find_all('div', class_= 'news-content clearfix')

        savePageToSheet(pageCards, ws)




    try:
        wb.save('zc-news')
    except:
        i=0
        while True:
            try:
                i += 1
                wb.save(f'zc-news-{i}')
                return
            except:
                pass




def main():
    scrapSite()


if __name__ == '__main__':
    main()