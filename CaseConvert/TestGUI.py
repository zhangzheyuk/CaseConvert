#!encoding:utf-8
import sys
import logging
import xlrd
import traceback
try:
    from lxml import etree
except ImportError:
    import xml.etree.ElementTree as etree
import Tkinter
import tkFileDialog
from Tkinter import *
import tkMessageBox

#设置logger配置
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                    datefmt='%a, %d %b %Y %H:%M:%S',
                    filename='./test.log',
                    filemode='w')

logging.debug('debug message')
logging.info('info message')
logging.warning('warning message')
logging.error('error message')
logging.critical('critical message')


root = Tkinter.Tk()
root.title("Tkinter-选择导入文件")
root.geometry("400x200+480+350")
entry = Entry(root, width=40)
entry.pack()
path = "./"


#选取文件路径
def callback():
    entry.delete(0, END)
    # 清空entry里面的内容
    # 调用filedialog模块的askdirectory()函数去打开文件夹
    # filepath = tkFileDialog.askdirectory()
    filepath = tkFileDialog.askopenfilename()
    if filepath:
        entry.insert(0, filepath) #将选择好的路径加入到entry里面

#处理函数
def tcConvert():
    path = entry.get()
    print(entry.get())
    logging.debug('select path ===' + path)

    f_in = xlrd.open_workbook(path)
    if f_in:
        logging.debug("t")


    sheet = f_in.sheet_by_index(0)
    # create XML
    testcases = etree.Element('testcases')

    print ("row = %d" % sheet.nrows)
    def format_str(rawstr):
        rawstr = "<![CDATA[<p>" + rawstr + "</p>"
        return rawstr


    for seq in range(1, sheet.nrows):
        # print(sheet.row_values(seq))
        try:
            name_ = sheet.row_values(seq)[0]
            summary_ = sheet.row_values(seq)[1]
            pre_ = sheet.row_values(seq)[2]
            importance_ = sheet.row_values(seq)[3]
            step_ = sheet.row_values(seq)[4]
            expect_ = sheet.row_values(seq)[5]
            exe_type_ = sheet.row_values(seq)[6]

            test_case = etree.SubElement(testcases, 'testcase', name=name_)
            summary = etree.SubElement(test_case, 'summary')
            summary.text = format_str(summary_)
            print(summary.text)
            logging.debug('summary.text====' + summary.text)


            preconditions = etree.SubElement(test_case, 'preconditions')
            print(pre_)
            # preconditions.text = u'"{0}"'.format(pre_)
            pre_ = pre_.replace("\n", "</br>")
            print(pre_)
            preconditions.text = format_str(pre_)
            print(preconditions.text)

            # <importance><![CDATA[2]]></importance>
            importance_level = etree.SubElement(test_case, 'importance')
            importance_level.text = str(int(importance_))
            print(importance_level.text)

            steps = etree.SubElement(test_case, 'steps')
            step = etree.SubElement(steps, 'step')
            step_number = etree.SubElement(step, 'step_number')
            step_number.text = str(1)

            # Transform the steps
            actions = etree.SubElement(step, 'actions')
            # actions.text = u'"{0}"'.format(step_)
            step_ = step_.replace("\n", "</br>")
            print(step_)
            actions.text = format_str(step_)
            print(actions.text)
            expectedresults = etree.SubElement(step, 'expectedresults')

            # expectedresults.text = u'"{0}"'.format(expect_)
            expect_ = expect_.replace("\n", "</br>")
            print(expect_)
            expectedresults.text = format_str(expect_)
            print(expectedresults.text)

            execution_type = etree.SubElement(step, 'execution_type')
            execution_type.text = str(int(exe_type_))
            print(execution_type.text)

        except Exception as e:
            print("line:", seq)
            print(str(e))
            for item in sys.exc_info():
                print item

    # s = etree.tostring(testcases, pretty_print=True)
    s = etree.tostring(testcases)
    pathlist = path.split("/")
    pathlist.pop()
    newpath = '/'.join(pathlist)
    f_out = open(newpath + "/my.xml", 'w')
    f_out.write(s)
    # button_success = Button(root, text="Convert Successfully")
    # button_success.pack()
    tkMessageBox.showinfo("showinfo demo", "Convert Successfully")

button = Button(root, text="Open", command=callback)
button.pack()

button_exe = Button(root, text="execute", command=tcConvert)
button_exe.pack()

root.mainloop()


