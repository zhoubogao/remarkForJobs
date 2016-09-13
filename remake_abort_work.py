#-*-coding:utf-8-*-

import xlrd
import xlwt
import time
import datetime
import calendar

data = xlrd.open_workbook('123.xls')
table = data.sheets()[0]
allrows = []
for rownum in range(table.nrows):
    allrows.append(table.row_values(rownum))
allrows = allrows[1:]
nl = []
for ri in allrows:
    nl.append(ri[1])

num_list = list(set(nl))
result_list = []
for num in num_list:
    one_person_list = []
    new_row = []
    time_list = []
    date_list = []
    for ri in allrows:
        if num != ri[1]:
            continue
        time_list.append(ri[2])
        name = ri[0]
    for t in time_list:
        date_list.append(t.split()[0])
    date_list = list(set(date_list))
    date_list = sorted(date_list,
                       key=lambda date_list: int(date_list.replace('/', '')))
    for d in date_list:
        same_date_list = []
        for t in time_list:
            if d not in t:
                continue
            same_date_list.append(t)
        if len(same_date_list) > 2:
            same_date_list = sorted(same_date_list, key=lambda same_date_list: int(same_date_list.split()[1].replace(':',"").strip()))
            same_date_list = [same_date_list[0], same_date_list[-1]]
        new_row.append(name)
        new_row.append(num)
        new_row.append(d)
        if len(same_date_list) == 0:
            new_row.append(u' ')
            new_row.append(u' ')
        elif len(same_date_list) == 1:
            same_date_list = same_date_list[0].split()[1]
            if int(same_date_list.replace(':', '')) >= 120000:
                new_row.append(u' ')
                new_row.append(same_date_list)
            else:
                new_row.append(same_date_list)
                new_row.append(u' ')
        else:
            new_row.append(same_date_list[0].split()[1])
            new_row.append(same_date_list[1].split()[1])

        one_person_list.append(new_row)
        new_row = []

    cur_year = int(one_person_list[0][2].split('/')[0])
    cur_month = int(one_person_list[0][2].split('/')[1])
    cur_cal = range(calendar.monthrange(cur_year, cur_month)[1]+1)[1:]
    cur_cal_temp = cur_cal
    
    if cur_cal[-1] != int(one_person_list[-1][2].split('/')[2]):
        insert_item = one_person_list[-1][:2]
        insert_item.append(str(cur_year) + '/' + str(cur_month) + '/' +str(cur_cal[-1]))
        insert_item.append(u' ')
        insert_item.append(u' ')
        one_person_list.append(insert_item)
    for cc in cur_cal:
        for one in one_person_list:
            cur_date = int(one[2].split('/')[2])
            if cc == cur_date:
            	break
            elif cc < cur_date:
                insert_item = one[:2]
                insert_item.append(str(cur_year) + '/' + str(cur_month) + '/' +str(cc))
                insert_item.append(u' ')
                insert_item.append(u' ')
                one_person_list.insert(one_person_list.index(one), insert_item)
                break
            else:
            	pass
    
    for one in one_person_list:
        cur_date = int(one[2].split('/')[2])
        dayOfWeek = datetime.datetime(cur_year, cur_month, cur_date).isoweekday()

        if dayOfWeek > 5 :
            one.append(u'正常')
            one.append(u' ')
        else:
            if one[3] == u' ' and one[4] == u' ':
                one.append(u'异常')
                one.append(u'无打卡记录')
            elif one[3] != u' ' and one[4] == u' ':
                one.append(u'异常')
                one.append(u'下班未打卡')
            elif one[3] == u' ' and one[4] != u' ':
                one.append(u'异常')
                one.append(u'上班未打卡')
            else:
                endTime = datetime.datetime.strptime(one[4].strip(), "%H:%M:%S")
                startTime = datetime.datetime.strptime(one[3].strip(), "%H:%M:%S")
                print endTime
                print startTime
                dura = (endTime - startTime)
                if str(dura) >= 93000 :
                    one.append(u'正常')
                    one.append(u' ')
                else:
                    one.append(u'异常')
                    one.append(u'未满8小时')

        if dayOfWeek == 1:
            dayOfWeek = u'星期一'
        elif dayOfWeek == 2:
            dayOfWeek = u'星期二'
        elif dayOfWeek == 3:
            dayOfWeek = u'星期三'
        elif dayOfWeek == 4:
            dayOfWeek = u'星期四'
        elif dayOfWeek == 5:
            dayOfWeek = u'星期五'
        elif dayOfWeek == 6:
            dayOfWeek = u'星期六'
        elif dayOfWeek == 7:
            dayOfWeek = u'星期日'
        else:
        	dayOfWeek = u'未知'
        one[2] = one[2] + u' ' + dayOfWeek
    result_list.extend(one_person_list)

result_list = sorted(result_list, key=lambda result_list: int(result_list[1]))
for r in result_list:
    print r

file = xlwt.Workbook(encoding='utf-8')
table = file.add_sheet('result1', cell_overwrite_ok=True)
table.write(0, 0, u'姓名')
table.write(0, 1, u'工号')
table.write(0, 2, u'日期')
table.write(0, 3, u'上班打卡时间')
table.write(0, 4, u'下班打卡时间')
table.write(0, 5, u'状态')
table.write(0, 6, u'异常情况')

row = 1
for ii in result_list:
    col = 0
    for i in ii:

        table.write(row, col, i)
        col += 1
    row += 1

now = time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime(time.time()))
file.save('result-' + now + '.xls')