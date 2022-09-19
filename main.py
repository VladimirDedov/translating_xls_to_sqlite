import openpyxl
import sqlite3


def readXLS(data_name: 'name file.xlsx'):
    """read file data.xlsx and return list of lists of requests with authors """
    list_id, list_request, list_go, list_person, list_email = [], [], [], [], []
    count = 1
    try:
        book = openpyxl.open(data_name, read_only=True)
        sheet = book.worksheets[0]  # first list in book
    except:
        print('Не удалось открыть файл data.xlsx')
    for row in range(1, sheet.max_row + 1):
        list_id.append(count)
        list_request.append(sheet[row][0].value)
        list_go.append(sheet[row][1].value)
        list_person.append(sheet[row][2].value)
        list_email.append(sheet[row][3].value)
        count += 1
    book.close()
    return list_id, list_request, list_go, list_person, list_email


def return_data():
    """return list of tuples"""
    list_data = readXLS('data.xlsx')
    lst_tuple = list(zip(list_data[0], list_data[1], list_data[2], list_data[3], list_data[4]))

    return lst_tuple


def main():
    flag = True
    try:
        conn = sqlite3.connect('sd.db')  # connected for sd.db or create BD
        cur = conn.cursor()  # for sql requests in BD
        cur.execute("""CREATE TABLE IF NOT EXISTS sd(
            go_id INT PRIMARY KEY,
            request TEXT,
            go TEXT,
            fio TEXT,
            e_mail TEXT
        )
        """)
    except:
        flag = False
        print("Не удалось подключиться или создать БД")

    conn.commit()

    tmp_lst = return_data()
    try:
        cur.executemany("INSERT INTO sd VALUES(?,?,?,?,?)", tmp_lst)
    except:
        flag = True
        print("Не удалось записать данные")
    conn.commit()
    if flag:
        print("Данные успешно перенесены в БД sd.db")


if __name__ == "__main__":
    main()
