import re

def char2num(char):
    def toExcelCol(y):
        q,r=divmod(y-1,26)
        return toExcelCol(q)+chr(r+ord('A')) if y!=0 else ''
    def toIdx(y):
        s,p=0,1
        for c in y[::-1]:
            s+=p*(int(c,36)-9)
            p*=26
        return s
    mn_c=toExcelCol(min(toIdx(re.split('(\D+)',c)[1]) for c in char))
    mx_c=toExcelCol(max(toIdx(re.split('(\D+)',c)[1]) for c in char))
    mn_r=str(min(re.split('(\D+)',c)[2] for c in char))
    mx_r=str(max(re.split('(\D+)',c)[2] for c in char))
    return mn_c+mn_r+":"+mx_c+mx_r

print char2num(['A1','B1','A2','B2'])
