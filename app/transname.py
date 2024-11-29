from base.myopenpyxl.xls2py import *

coachs = []
students = []
students_gift = []
coachs_pocket = []


def goodboy(coachs, students, students_gift):
    for index_c, coach in enumerate(coachs):

        #获得每一个学生的得分放到students[index][0]里
        for index, student in enumerate(students):
            score = 0
            sujects = student[0].split(",")
            for suject in sujects:
                if suject in coach:
                    score += 1
            students[index][1] = score

        #初始化最大得分和最大得分的学生的序号
        max_score = 0
        max_score_student_index = 0

        #查找最高的分的学生的得分和序号存放在上面的max_score、max_score_student_index里
        for index, student in enumerate(students):
            if students[index][1] > max_score:
                max_score = students[index][1]
                max_score_student_index = index

        #判断最好的学生的得分，如果是0则'-'传入最终结果，否则学生的礼品交给教官
        if max_score == 0:
            coachs_pocket[index_c] = '-'
        else:
            coachs_pocket[index_c] = students_gift[max_score_student_index]

#设置操作文件
coachs = xls2list('G:\\OneDrive - lingdu\\wkzz\\2024.11.26祥峰调度\\更新后的数据.xlsx', 'c5', coachs)
#设置参考模板不规范名位置
students = xls2list("G:\\project\\LHSextract\\企业关键字对应规范名.xlsx", 'A1', students)
for index, student in enumerate(students):
    students[index] = [student, 0]
#设置参考模板规范名位置
students_gift = xls2list("G:\\project\\LHSextract\\企业关键字对应规范名.xlsx", 'B1', students_gift)
coachs_amount = len(coachs)
coachs_pocket = [''] * coachs_amount
goodboy(coachs, students, students_gift)
#设置输出位置
list2xls('G:\\OneDrive - lingdu\\wkzz\\2024.11.26祥峰调度\\更新后的数据.xlsx', 'd5', coachs_pocket)
print(coachs_pocket)




