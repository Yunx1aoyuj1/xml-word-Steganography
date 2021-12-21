# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


# def print_hi(name):
#     # Use a breakpoint in the code line below to debug your script.
#     print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.
#
#
# # Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     print_hi('PyCharm')

# 具体思路看论文
import re
import zipfile
import os
# gui套件
import tkinter as tk
import tkinter.messagebox


def str_to_hex(s):
    return ''.join([hex(ord(c)).replace('0x', '') for c in s])


def hex_to_str(s):
    text_list = re.findall(".{2}", s)
    new_text = " ".join(text_list)
    return ''.join([chr(i) for i in [int(b, 16) for b in new_text.split(' ')]])

#def watermark_of_word(resource, watermark):

def watermark_of_word():
    resource = e1.get()
    watermark = e2.get()
    zipFile = zipfile.ZipFile(os.path.join(os.getcwd(), resource + ".docx"), "r")  # 以zip打开word得到xml
    zipFile.extract("word/document.xml")  # 解压
    file = open("word/document.xml", "r+", encoding="utf-8")
    file.seek(0, 0)
    Watermark = str_to_hex(watermark)

    string = file.read()
    cite = 0
    stat = False
    string1 = list(string)
    # Watermark = list(Watermark)
    length = len("w:rsidR=\"")  # 添加水印的版本控制位
    mark_of_Watermark = 0
    for i in range(len(string)):
        cite = string.find("w:rsidR=\"", cite)
        for j in range(0, 7):
            string1[cite + length + j] = Watermark[mark_of_Watermark]
            if (mark_of_Watermark == len(Watermark) - 1):
                stat = True
                break
            else:
                mark_of_Watermark += 1
        if stat == True:
            break
        cite = cite + 1
    string = ''.join(string1)
    file.seek(0, 0)
    file.write(string)
    file.closed

    zout = zipfile.ZipFile(resource + "_watermark.docx", 'w')  # 被写入对象
    zin = zipfile.ZipFile(resource + ".docx", 'r')  # 读取对象
    file = open("word/document.xml", "rb")
    text = file.read()
    for file in zipFile.namelist():
        re_file = zin.read(file)
        if file != "word/document.xml":  # 如果不是目标文件直接复制
            zout.writestr(file, re_file)
        else:
            zout.writestr("word/document.xml", text)  # 是目标文件，将加好水印的目标文件写入
    zout.close()
    zin.close()
    print(text)
    tk.messagebox.showinfo("结果", "成功")

def decode_of_word(resource):  # dome不想写加密 这里的水印主要是告诉程序读几位（对于用了加密的，可将密文补成定长

    resource = e1.get()
    # 本程序默认定长
    zipFile = zipfile.ZipFile(os.path.join(os.getcwd(), resource + ".docx"), "r")  # 以zip打开word得到xml
    zipFile.extract("word/document.xml")  # 解压
    file = open("word/document.xml", "r+", encoding="utf-8")
    file.seek(0, 0)
    string = file.read()
    file.close()

    length_of_watermark = 30
    cite = 0
    stat = False
    string1 = list(string)
    Watermark = ""
    length = len("w:rsidR=\"")  # 添加水印的版本控制位
    for i in range(len(string)):
        cite = string.find("w:rsidR=\"", cite)
        # print(cite)
        for j in range(0, 7):
            Watermark = Watermark + string[cite + length + j]
            length_of_watermark = length_of_watermark - 1
            if length_of_watermark == 0:
                break
        print(length_of_watermark)
        cite = cite + 1
        if length_of_watermark == 0:
            break
    Watermark = hex_to_str(Watermark)
    print(Watermark)  # 输出结果可能会出现乱码
    return  Watermark

def watermark_of_ppt(resource, watermark):#完成度不高 dome不想写加密
    zipFile = zipfile.ZipFile(os.path.join(os.getcwd(), resource + ".pptx"), "r")
    Watermark = "<a: tblStylestyleId =\"{" + "\"" + watermark + "}\"" + "styleName =\"" + watermark + "\"/>"
    zipFile.extract("ppt/tableStyles.xml")
    file = open("ppt/tableStyles.xml", "r+", encoding='utf-8')
    file.seek(0, 0)
    string = file.read()
    start = len(string)  # 默认在文件尾
    strs = string.split("</a:tblStyleLst>");
    text = strs[0] + Watermark + "</a:tblStyleLst>" + strs[1]
    file.seek(0, 0)
    file.write(text)
    file.closed

    zout = zipfile.ZipFile(resource + "_watermark.pptx", 'w')  # 被写入对象
    zin = zipfile.ZipFile(resource + ".pptx", 'r')
    file = open("ppt/tableStyles.xml", "rb")
    text = file.read()
    for file in zipFile.namelist():
        re_file = zin.read(file)
        if file != "ppt/tableStyles.xml":
            zout.writestr(file, re_file)
        else:
            zout.writestr("ppt/tableStyles.xml", text)
    zout.close()
    zin.close()#y


def watermark_of_excel(resource, watermark):#完成度不高 dome不想写加密
    zipFile = zipfile.ZipFile(os.path.join(os.getcwd(), resource + ".xlsx"), "r")
    Watermark = "< x14 : id> {" + watermark + "}" + "</ x14:id>"
    zipFile.extract("xl/styles.xml")
    file = open("xl/styles.xml", "r+", encoding='utf-8')
    file.seek(0, 0)
    string = file.read()
    start = len(string)  # 默认在文件尾
    strs = string.split("<extLst>");
    text = strs[0] + "<extLst>" + Watermark + strs[1]
    file.seek(0, 0)
    file.write(text)
    file.closed

    zout = zipfile.ZipFile(resource + "_watermark.xlsx", 'w')  # 被写入对象
    zin = zipfile.ZipFile(resource + ".xlsx", 'r')
    file = open("xl/styles.xml", "rb")
    text = file.read()
    for file in zipFile.namelist():
        re_file = zin.read(file)
        if file != "xl/styles.xml":
            zout.writestr(file, re_file)
        else:
            zout.writestr("xl/styles.xml", text)
    zout.close()
    zin.close()

def pr_decode_of_word():
    mess= decode_of_word(e1.get)
    tk.messagebox.showinfo("结果", mess)


#watermark_of_word("test", "nuaa")
#decode_of_word("test", "nuaa")
# 第1步，实例化object，建立窗口window
window = tk.Tk()

# 第2步，给窗口的可视化起名字
window.title('水印添加器')

# 第3步，设定窗口的大小(长 * 宽)
window.geometry('500x300')  # 这里的乘是小x

# 第4步，在图形界面上设定输入框控件entry框并放置

w1 = tk.Label(window, text="文件名")
w1.pack()
e1 = tk.Entry(window, show=None)  # 显示成明文形式
e1.pack()


w2 = tk.Label(window, text="水印")
w2.pack()
e2 = tk.Entry(window, show=None)  # 显示成明文形式
e2.pack()



#创建几个button
r1 = tk.Button(window, text='word 加水印', command=watermark_of_word)
r1.pack()

r2 = tk.Button(window, text='word 解水印', command=pr_decode_of_word)
r2.pack()

r3 = tk.Button(window, text='ppt 加水印',)
r3.pack()

r4 = tk.Button(window, text='xlx 加水印', )
r4.pack()



# 第6步，创建并放置按钮触
#b1 = tk.Button(window, text='查询', width=10, height=2, command=search)
#b1.pack()


# 第8步，主窗口循环显示
window.mainloop()