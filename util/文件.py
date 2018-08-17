import os
import winreg
def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                         r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')  # 利用系统的链表
    return str(winreg.QueryValueEx(key, "Desktop")[0])  # 返回的是Unicode类型数据
def get_Temp(*file):
    deskdop = get_desktop()
    # 获取临时文件
    temp = os.path.join(deskdop, 'temp')
    # 判断文件是否存在
    if not os.path.exists(temp):
        os.mkdir(temp)
    if file:
        temp = os.path.join(temp, file)
        if not os.path.exists(temp):
            os.mkdir(temp)
    return temp
def list_dir(path, *ends):
    '''
    这个列表下的所有文件
    :param path:
    :param ends:
    :return:
    '''
    files = os.listdir(path)
    result = set()
    for i in files:
        for end in ends:
            if (i.endswith(end)):
                result.add(os.path.join(path, i))
    return result
def get_excel(path):
    '''
    获取该目录下的excel文件的集合
    :param path:
    :return:
    '''
    return list_dir(path, 'xls', 'xlsx')
if __name__ == '__main__':
    # 就是获取当前目录，并组合成新目录
    getcwd = os.getcwd()
    print(getcwd)
    # 当前文件的名称
    print(os.path.basename(__file__))
    # 获取当前文件的绝对路径
    print(__file__)
    # 新建临时文件
    # print(list_dir(get_Temp(), "xlsx"))
    print(get_excel(get_Temp()))
