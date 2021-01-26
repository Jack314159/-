##引库
import xlrd
import random

##同户名
name=input('请输入你的名字：')
print()
Name=input('请输入题库名称（以文件名为标准）：')
print()

##导入题库
input('请确保题库的xls文件和本程序在同一目录内，\n[按下回车继续]')
while True:
    try:
        book = xlrd.open_workbook('{}.xls'.format(Name))
        table = book.sheet_by_name('Sheet1')
        break
    except:
        input('错误！！！请对照说明找出原因 \n[按下回车键继续]')
print('导入成功')

##调试
q = eval(input('请输入【题干 】所在列（从 0 开始）：'))
a = q+1
b = a+1
c = b+1
d = c+1
e = d+1
s = e+1
m = s+1


##简介
print('******* {} 欢迎进入刷题模式*******'.format(name))
print('*******★☆★{}自测程序★☆★*******'.format(Name))
print('*************★逍遥出品，必属精品★***************')
print()

##模式
print('{:*^20}'.format('模式选择'))
print('【A】模拟考试')
print('【B】顺序遍历')
print('【C】乱序遍历')
print('【D】自定义刷题')
print('【E】错题重做')
print('【F】按难度刷题')
mode = input('{} which mode do you want ?\n'.format(name)).upper()

##模式通式
rows = table.nrows
if mode == 'A':
    num = 100
    ls1 = list(range(1,rows))
    random.shuffle(ls1)
    count = ls1[:num]
elif mode == 'B':
    count = list(range(1,rows))
elif mode == 'C':
    ls1 = list(range(1,rows))
    random.shuffle(ls1)
    count = ls1[:]
elif mode == 'D':
    num = eval(input('{}请输入你要刷的题数： '.format(name)))
    ls1 = list(range(1,rows))
    random.shuffle(ls1)
    count = ls1[:num]
elif mode == 'E':
    count0 = input('{},请按要求输入你的错题编码：'.format(name))
    count = sorted(count0)
elif mode == 'F':
    Fls1=[]
    Fls2=[]
    Fls3=[]
    for i in range(1,rows):
        grade = table.cell(i,m).value
        if grade == '容易':
            Fls1.append(i)
        elif grade == '一般':
            Fls2.append(i)
        else:
            Fls3.append(i)
    g = input('请选择难度：【A】容易 【B】一般 【C】难\n').upper()
    if g == 'A':
        count = Fls1
    elif g == 'B':
        count = Fls2
    elif g == 'C':
        count = Fls3
        
    
####刷题
Count = 0
nwrong = []
right = 0
wrong = 0
for i in count:
    Count+=1
    question = table.cell(i,q).value
    A = table.cell(i,a).value
    B = table.cell(i,b).value
    C = table.cell(i,c).value
    D = table.cell(i,d).value
    E = table.cell(i,e).value
    answer = table.cell(i,s).value
    print('{0}、{1}'.format(Count,question))
    print('【A】、{}'.format(A))
    print('【B】、{}'.format(B))
    print('【C】、{}'.format(C))
    print('【D】、{}'.format(D))
    if E!= '':
        print('【E】、{}'.format(E))
    ianswer = input('给出你的选择：\n').upper()
    if ianswer == answer:
        print('【答案正确】')
        right+=1
    elif ianswer == 'N':
        break
    else:
        wrong+=1
        nwrong.append(i)
        if wrong == 1:
            print('【first kill !!!】')
        else:
            print('【回答错误】')
        print('正确答案为：【{}】'.format(answer))
    
print()
print('用时(暂未开发)：')
print('***************{}恭喜你完成本次刷题之旅***************'.format(name))

##结算
per = right/Count
print('正确个数：{}'.format(right))
print('错误个数：{}'.format(wrong))
print('正确率：{:.2%}'.format(per))

##错题导入
input('系统将为您生成本次错题集，如不需要直接关闭即可\n【按回车键继续】')
file = open('错题.txt','w')
file.write('错题编码为：{}\n\n'.format(nwrong))
file.write('以下为{}的错题\n'.format(name))
for k in sorted(nwrong):
    A = table.cell(k,a).value
    B = table.cell(k,b).value
    C = table.cell(k,c).value
    D = table.cell(k,d).value
    E = table.cell(k,e).value
    answer = table.cell(k,s).value
    file.write('{}、{}\n'.format(k,table.cell(k,q).value))
    file.write('A、{}\n'.format(A))
    file.write('B、{}\n'.format(B))
    file.write('C、{}\n'.format(C))
    file.write('D、{}\n'.format(D))
    if E != '':
        file.write('E、{}\n'.format(E))
    file.write('【正确答案为：{}】\n'.format(answer))
file.close()
input('生成完毕\n【按回车键结束】')
