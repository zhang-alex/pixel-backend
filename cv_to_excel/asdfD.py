import math
base26 = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
def find_powers(num) :
    i = 1
    counter = 1
    while num >= i:
        i += 26**counter
        counter += 1
    digits = counter - 1
    for i in range(digits) :
        num -= 26**i
    cell_is_A=0
    if num==0 : cell_is_A=1
    base10_26 = []
    for i in range(digits) :
        base10_26.append(0)
        temp_quotient = int(num/(26**(digits-1-i)))
        num -= (26**(digits-1-i)) * temp_quotient
        base10_26[i] += temp_quotient
    cell = ""
    for i in range(len(base10_26)) :
        cell += base26[base10_26[i]]
    print "cell:",cell
find_powers(702)
