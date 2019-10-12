import xlwt
import datetime
import json


def opera_line(txt):
    temp = txt.split('=')[1:]
    line_t = []
    for item in temp:
        start = item.index('"')
        end = item.rindex('"')
        v = item[start + 1:end]
        line_t.append(v)
    return line_t


def save_excel(arg1, save_path):
    workbook = xlwt.Workbook()
    sheet1 = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
    for i, line_t in enumerate(arg1):
        for j, token in enumerate(line_t):
            sheet1.write(i, j, token)

    workbook.save(save_path)
    print('创建excel文件完成!')


if __name__ == '__main__':
    with open('data/config.txt', encoding='utf-8') as config_file:
        config = json.loads(config_file.read())
    txt_path = config['txt_path']
    excel_path = config['excel_path']
    year = config['year']
    month = config['month']
    other_name = config['other_name']

    with open(txt_path) as txt_file:
        lines = txt_file.readlines()
    opera_lines = [['time', 'id', 'name', 'authority', 'card_src']]
    for line in lines[1:-1]:
        line = line.strip()
        temp_line = opera_line(line)
        if temp_line[2] == '':
            if temp_line[1] in other_name:
                temp_line[2] = other_name[temp_line[1]]
        datenow = datetime.datetime.strptime(temp_line[0], '%Y-%m-%d %H:%M:%S')
        y = datenow.year
        m = datenow.month
        if str(year) == str(y) and str(month) == str(m):
            opera_lines.append(temp_line)
    save_excel(opera_lines, excel_path)
