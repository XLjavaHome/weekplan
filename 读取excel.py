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
                    lastbugSet.add(textValue)
                if cellType.value == '开发任务' and textValue != '':
                    lastTaskSet.add(textValue)
        # 本周工作计划
        if cell0.value == 'BUG编号':
            for row in range(0, maxrow):
                bugSet.add(sheet.cell(row, 1).value)
                # 本周工作计划
        if cell0.value == '任务名称':
            for row in range(0, maxrow):
                taskSet.add(sheet.cell(row, 0).value)
    num = 1
    result = "上周工作内容：\n"
    if lastTaskSet:
        result += "{0}.开发任务:{1}。\n".format(*[num, '、'.join(lastTaskSet)])
        num = num + 1
    if lastbugSet:
        result += "{0}.BUG修复:{1}。\n".format(*[num, '、'.join(lastbugSet)])
    result += '本周工作计划：\n'
    num = 1
    if taskSet:
        result += "{0}.开发任务:{1}。\n".format(*[num, '、'.join(taskSet)])
        num = num + 1
    if bugSet:
        result += "{0}.BUG修复:{1}。\n".format(*[num, '、'.join(bugSet)])
    return result
if __name__ == '__main__':
    # 扫描的文件目录
    downloads = "D:\Downloads"
    print(get_week(downloads))
