from fake_useragent import UserAgent
import requests
from bs4 import BeautifulSoup
import time
import random
import json
from xlsxwriter.workbook import Workbook

HEADERS = {
    "User-Agent": UserAgent().random,
    "Accept-Language": "ru-RU,ru;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "DNT": "1",
    "Upgrade-Insecure-Requests": "1",
}

PAGE = 1
URL = f"https://habr.com/ru/flows/develop/articles/page{PAGE}/"


def CollectData(url, page=PAGE):  # Собираем данные с сайта
    data = []
    reps = [",", ":", ";"]
    page = requests.get(url, headers=HEADERS)

    if page.status_code == 200:
        soup = BeautifulSoup(page.content, "lxml")
        post = soup.find_all("article", class_="tm-articles-list__item")

        # Вычисление максимальной страницы + переход на следующую
        # max_page = soup.find_all(
        #     "a", class_="tm-pagination__page")
        # for i in range(1, int(max_page[-1].text.strip()) + 1):
        #     post = soup.find_all(
        #         "article", class_="tm-articles-list__item")

        # Собираем данные
        for item in post:
            # Задержка между запросами
            time.sleep(random.uniform(0.1, 0.5))
            title = item.find(
                "a", class_="tm-title__link").find("span").text.strip()
            link = "https://habr.com" + \
                item.find("a", class_="tm-title__link")["href"]

            # Описание
            description = None
            description_sources = [
                item.find("div", class_="tm-article-body"),
                item.find("div", class_="article-formatted-body")
            ]
            for source in description_sources:
                if source and source.find("p") is not None:
                    description = source.find("p").text.strip()
                    break

            # Изображение
            image = None
            image_sources = [
                item.find("div", "tm-article-snippet__cover"),
                item.find("div", "tm-article-snippet__cover_cover"),
                item.find("div", "article-formatted-body")
            ]
            for source in image_sources:
                if source and source.find("img") is not None:
                    image = source.find("img")["src"]
                    break

            # Удаляем ненужные символы в названиях
            for rep in reps:
                title = title.replace(rep, "")
            # Заполняем список data
            if title is not None and link is not None and image is not None and description is not None:
                data.append(
                    {"Название": title, "Описание": description, "Ссылка": link, "Изображение": image})
        # Переход на следующую страницу
        global PAGE
        PAGE += 1
        return data
    else:
        print(f"Ошибка: {page.status_code} - {page.text}")
        return -1


def SaveData_JSON(data, filename):  # Сохраняем данные в JSON файл
    with open(f"{filename}.json", "w", encoding="utf-8") as file:
        json.dump(data, file, indent=4, ensure_ascii=False)
        print(f"Данные успешно собраны и сохранены в {filename}.json")


def SaveData_CSV(data, filename):  # Сохраняем данные в CSV файл
    with open(f"{filename}.csv", "w", encoding="utf-8") as file:
        file.write("Название,Описание,Ссылка,Изображение\n")
        for item in data:
            file.write(f"{item['Название']},{item['Ссылка']}\n")
        print(f"Данные успешно собраны и сохранены в {filename}.csv")


def SaveData_XLSX(data, filename):  # Сохраняем данные в XLSX файл
    workbook = Workbook(f"{filename}.xlsx")
    worksheet = workbook.add_worksheet()

    # Записываем заголовки
    header_format = workbook.add_format({
        'bold': True,
        'font_color': 'blue',
        'align': 'center',  # Центрирование текста
        'valign': 'vcenter'  # Вертикальное центрирование
    })
    worksheet.write(0, 0, "Название", header_format)
    worksheet.write(0, 1, "Описание", header_format)
    worksheet.write(0, 2, "Ссылка", header_format)
    worksheet.write(0, 3, "Изображение", header_format)

    # Записываем данные
    for row_num, item in enumerate(data, start=1):
        worksheet.write(row_num, 0, item["Название"])
        worksheet.write(row_num, 1, item["Описание"])
        worksheet.write(row_num, 2, item["Ссылка"])
        worksheet.write(row_num, 3, item["Изображение"])

    # Устанавливаем ширину столбцов в зависимости от максимальной длины содержимого
    max_lengths = [len("Название"), len("Ссылка"),
                   len("Описание"), len("Изображение")]
    for item in data:
        max_lengths[0] = max(max_lengths[0], len(item["Название"]))
        max_lengths[1] = max(max_lengths[1], len(item["Описание"])/4)
        max_lengths[2] = max(max_lengths[2], len(item["Ссылка"]))
        max_lengths[3] = max(max_lengths[3], len(item["Изображение"]))
    # Устанавливаем ширину столбцов
    for i, length in enumerate(max_lengths):
        worksheet.set_column(i, i, length + 2)  # Добавляем небольшой запас

    workbook.close()
    print(f"Данные успешно собраны и сохранены в {filename}.xlsx")


def main():
    data = CollectData(URL)
    SaveData_JSON(data, "Parse_HABR")
    SaveData_CSV(data, "Parse_HABR")
    SaveData_XLSX(data, "Parse_HABR")


if __name__ == "__main__":
    main()
