def num(n):
    a = 1
    while a < n:
        if a % 3 == 0:
            print(a,"*")
            a = a + 1
        else:
            print(a)
            a = a + 1

num(41)