import argparse
import collections
import datetime
import os
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

    table_file_name = os.environ.get('TABLE_FILE_NAME', 'wine.xlsx')
    parser = argparse.ArgumentParser(description='Reads the excel table to run the wine shop site')
    parser.add_argument('table_file_name', nargs='?', default=table_file_name,
                        help='path to table with products description')
    args = parser.parse_args()
    table_file_name = args.table_file_name
    print(table_file_name)

    rendered_page = template.render(
        wine_maker_age=get_wine_maker_age(),
        products=get_products(table_file_name),
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
