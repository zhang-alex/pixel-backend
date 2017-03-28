import math

base26 = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']

def find_powers(num) :

    i = 1
    counter = 1

    while num >= i:
        i += 26**counter
        print num,"is at least",counter,"digits"
        counter += 1

    digits = counter - 1

    print num,"is",digits,"digits"
    print math.log(num,26)

    for i in range(digits) :
        num -= 26**i

    print num

    cell_is_A=0
    if num==0 : cell_is_A=1


    ##num is now a regular integer.

    base10_26 = []

    for i in range(digits) :

        base10_26.append(0)

        print 'power of 26 we are dividing by: 26^',digits - 1 - i

        temp_quotient = int(num/(26**(digits-1-i)))
        print 'temporary quotient',temp_quotient

        num -= (26**(digits-1-i)) * temp_quotient
        print 'number after subtraction:',num

        base10_26[i] += temp_quotient

        ##num -= (26**(((new_digits-1)-i))) * base10_26[i]

    print 'array values:'
    for i in base10_26 :
        print i

    cell = ""

    for i in range(len(base10_26)) :
        cell += base26[base10_26[i]]

    print "cell:",cell


#-----------------------------------------------------------

find_powers(702)
