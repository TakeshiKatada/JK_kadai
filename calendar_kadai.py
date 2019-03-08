#ターミナルでカレンダーを表示する

#カレンダーの数を入れるリスト
# year = (2019)
# month = [1,2,3,4,5,6,7,8,9,10,11,12]
# day = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31]

# # dict([('january',31),('february',28),(',mardh',31),('april',30),('may',31),('june',30)
# # ('july',31),('august',30),('september',31),('october',31),('november',30),('December',31)])

# week_name = ("月","火","水","木","金","土","日")

# def one_month(x):
#     print(month(x-1))
#     print(day())

# print(year)
# one_month(1)




Month = input("何月？:")
Month = int(Month)
print(2019,end =' ')
print(Month)
print("月|火|水|木|金|土|日")

if Month == 1:
    print("  |1 |2 |3 |4 |5 |6 |\n7 |8 |9 |10|11|12|13|\n14|15|16|17|18|19|20|\n21|22|23|24|25|26|27|\n28|29|30|31|")
elif Month == 2:
    print("  |  |  |  |  |1 |2 |\n3 |4 |5 |6 |7 |8 |9 |\n10|11|12|13|14|15|16|\n17|18|19|20|21|22|23|\n24|25|26|27|28|")
elif Month == 3:
    print()
elif Month == 4:
    print('4月')
elif Month == 5:
    print('5月')
elif Month == 6:
    print('6月')



else:
    pass