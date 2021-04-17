def symbol_filter(arg):
    if arg == 1:
        return '○'
    else:
        return ''


from jinja2 import Environment, FileSystemLoader
from read_excel_works import read_excel_work
from datetime import datetime
from pprint import pprint

env = Environment(loader=FileSystemLoader('./', encoding='utf8'))
env.filters['symbol_filter'] = symbol_filter
tmpl = env.get_template('template.html')

works = []
persons = []

excel_works_data = read_excel_work()
for keys, values in excel_works_data.items():
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

    messages = [
        '横断歩道に歩行者がいたら、一時停止してください。',
        '法定速度を遵守してください。',
        '日が短いため、早めのライトアップを心がけてください。'
    ]

    accidents = [
        '前を走る車両が急ブレーキをかけたため追突しそうになった。十分な車間距離を保つこと。',
        '自転車が急に道路を横断してきた。自転車は急な動きが多いため、見かけたら警戒すること。'
    ]

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

    pprint(persons)

    works = []
    persons = []