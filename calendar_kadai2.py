"""
calendar_kadai2.py
カレンダーを表示するスクリプト
2019.2.25 T.Katada 作成
"""

import datetime

def year_check(year):
    """
    閏年かどうか？
    """
    plus_day = 0
    if year % 400 == 0:
        plus_day = 1
    elif year % 4 == 0 and year % 100 != 0 :
        plus_day = 1
    else:
        plus_day = 0
    return plus_day

def last_day_check(year,month):
    """
    monthで受け取った月の最終日を返す
    """
    if month != 2:
        if month in [4,6,9,11]:
            return 30
        return 31

    return 28 + year_check(year)


def print_calendar(year,month):
    """
    カレンダーを表示する
    """
    print(str(year)+'年'+str(month)+'月')
    print("日 月 火 水 木 金 土")

    first_day = datetime.date(year,month,1)#入力された年月の初日
    first_day_s = (first_day.isoweekday()) % 7  #入力された曜日を数値で返す,月曜日は1,日曜日は7
    #if first_day_s == 7:
    #    first_day_s = 0
    month_last_day = last_day_check(year,month)#入力された月の最終日を返す

    line = ''
    if first_day_s > 0:
        line += '  '
        line += ('   ' * (first_day_s - 1))

    w = first_day_s
    for d in range(1, month_last_day + 1):
        if w > 0:
            line += ' '
        if d < 10:
            line += ' '

        line += str(d)

        if w >= 6:
            print(line)
            line = ''

        w = (w + 1) % 7

    if line != '':
        print(line)

    print()

#ターミナル上で年月を入力
year = input("年:")
year = int(year)
month = input("月:")
month = int(month)
print_calendar(year,month)