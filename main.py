import collections
import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape

FOUNDATION_YEAR = 1920


def get_ru_word_year(years):
    k = (years // 10 % 10 != 1) * years % 10
    return ['лет', 'года', 'год'][(k == 1) + (1 <= k <= 4)]


def get_wine_maker_age():
    wine_maker_age = datetime.date.today().year - FOUNDATION_YEAR
    return f"{wine_maker_age} {get_ru_word_year(wine_maker_age)}"


def get_products(file_name):
    product_cards = pandas.read_excel(file_name, keep_default_na=False).to_dict(orient='records')
    products = collections.defaultdict(list)
    for product_card in product_cards:
        category = product_card.get('Категория')
        products[category].append(product_card)
    return products


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        wine_maker_age=get_wine_maker_age(),
        products=get_products('wine.xlsx'),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
