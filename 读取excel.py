import xlrd

import util.文件
'''
 作者：徐立
 2018-08-15 22:35
'''
class weekcontent(object):
    pass
def get_week(diretory):
    # 上周计划
    lastTaskSet = set()
    # 类型
    lastbugSet = set()
    # 本周计划
    taskSet = set()
    bugSet = set()
    # 上周修复BUG数量
    lastWeekBugCount = set()
    # 上周开发任务数
    lastWeekTaskCount = 0;
    # 本周的BUG数量
    bugSetCount = set()
    for file in util.文件.get_excel(diretory):
        excelFile = xlrd.open_workbook(file)
        # 获取第一个sheet，第一行是日志类型就进行。
        sheet = excelFile.sheet_by_index(0)
        maxrow = sheet.nrows
        # 根据第一行的第一个单元格是类型来判断
        cell0 = sheet.cell(0, 0)
        #  上周工作内容
        if (cell0.value == '日志类型'):
            for row in range(0, maxrow):
                # 第一列的内容
                cellType = sheet.cell(row, 0)
                textValue = sheet.cell(row, 5).value
                if cellType.value == 'BUG修复' and textValue != '':
                    lastWeekBugCount.add(sheet.cell(row, 4).value)
                    lastbugSet.add(textValue)
                if cellType.value == '开发任务' and textValue != '':
                    lastWeekTaskCount += 1
                    lastTaskSet.add(textValue)
        # 本周工作计划,从第一行开始
        if cell0.value == 'BUG编号':
            for row in range(1, maxrow):
                # 需求名称
                bugSet.add(sheet.cell(row, 1).value)
                # BUG数量
                bugSetCount.add(sheet.cell(row, 0).value)
                # 本周工作计划
        if cell0.value == '任务名称':
            for row in range(1, maxrow):
                taskSet.add(sheet.cell(row, 0).value)
    # 每行开头的数字
    num = 1
    result = "上周工作内容：\n"
    if lastTaskSet:
        result += "{0}.开发{1}的任务。\n".format(*[num, '、'.join(lastTaskSet)])
        num = num + 1
    if lastbugSet:
        result += "{0}.修复{1}的BUG。（{2}个BUG）\n".format(*[num, '、'.join(lastbugSet), len(lastWeekBugCount)])
    result += '本周工作计划：\n'
    num = 1
    if taskSet:
        result += "{0}.开发{1}任务。\n".format(*[num, '、'.join(taskSet)])
        num = num + 1
    if bugSet:
        result += "{0}.修复{1}的BUG。（{2}个BUG）\n".format(*[num, '、'.join(bugSet), len(bugSetCount)])
    return result
if __name__ == '__main__':
    # 扫描的文件目录
    downloads = "D:\Downloads"
    print(get_week(downloads))
