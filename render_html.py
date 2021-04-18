from jinja2 import Environment, FileSystemLoader
from read_excel_works import read_excel_works
from read_excel_messages_or_accidents import read_excel_messages_or_accidents
from datetime import datetime
# from pprint import pprint


def symbol_filter(arg):
    if arg == 1:
        return '○'
    else:
        return ''


def main():
    env = Environment(loader=FileSystemLoader('./', encoding='utf8'))
    env.filters['symbol_filter'] = symbol_filter
    tmpl = env.get_template('template.html')  

    messages = read_excel_messages_or_accidents(0)
    accidents = read_excel_messages_or_accidents(1)
    excel_works_data = read_excel_works()
    for keys, values in excel_works_data.items():
        works = []
        persons = []
        company_name = keys
        date = values['日付'].strftime('%Y/%m/%d')

        for row in values['作業内容']:
            works.append({'place': row[0], 'content': row[1], 'classifying': row[2]})

        for row in values['出勤名簿']:
            tmp_dict = {}
            tmp_dict['name'] = row[0]
            for i in range(1, 10):
                tmp_dict[i] = int(row[i])
            persons.append(tmp_dict)

        data = {
            'company_name': company_name,
            'date': date,
            'works': works,
            'persons': persons,
            'messages': messages,
            'accidents': accidents
            }
        html = tmpl.render(data)
        out_file_name = '出力ファイル/' + company_name + '.html'
        with open(out_file_name, 'w', encoding='utf-8') as f:
            f.write(str(html))


if __name__ == '__main__':
    main()