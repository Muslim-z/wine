from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape
from dotenv import load_dotenv
from datetime import datetime
import pandas
from _collections import defaultdict
from pprint import pprint


def determine_the_age_ending(starting_year):
    age = datetime.now().year - starting_year
    if age % 100 in (11, 12, 13, 14):
        return f'{age} лет'
    elif age % 10 == 1:
        return f'{age} год'
    elif age % 10 in (2, 3, 4):
        return f'{age} года'
    else:
        return f'{age} лет'


def get_excel_wines(excel_file_path):
    excel_data = pandas.read_excel(excel_file_path)
    excel_data.fillna('', inplace=True)
    wines = defaultdict(list)
    for line in excel_data.to_dict(orient='records'):
        category = line.pop('Категория')
        wines[category].append(line)
    return wines


def render_template(data):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html'])
    )
    template = env.get_template('template.html')
    return template.render(data)


def save_index_file(content, path_html):
    with open(path_html, 'w', encoding="utf8") as file:
        file.write(content)


def main():
    load_dotenv()
    starting_year = 1920
    age_data = determine_the_age_ending(starting_year)
    excel_file_path = 'wine3.xlsx'
    wine_data = get_excel_wines(excel_file_path)
    data_to_render = {
        'age': age_data,
        'wines': wine_data
    }
    rendered_page = render_template(data_to_render)
    output_html_file = 'index.html'
    save_index_file(rendered_page, output_html_file)
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
