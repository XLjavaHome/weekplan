import tkinter.filedialog
from tkinter import *

# 剪切板
import pyperclip

import 读取excel
root = Tk()
def xz():
    filename = tkinter.filedialog.askdirectory()
    if filename != '':
        try:
            result = 读取excel.get_week(filename)
        except:
            print("文件处于打开状态")
        # 删除所有内容
        text.delete('1.0', 'end')
        # 将内容复制进剪切板
        pyperclip.copy(result)
        text.insert(1.0, result)
    else:
        text.config(text="您没有选择任何文件夹");
text = Text(root, height=10)
text.pack()
root.title("读取excel生成周报")
btn = Button(root, text="请选择excel所在目录", command=xz)
btn.pack()
root.mainloop()
