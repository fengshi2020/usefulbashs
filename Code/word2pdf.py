import os, re
import threading
import pythoncom
from win32com.client import DispatchEx
from gooey import Gooey,GooeyParser

class WordConvertToOther:
    def DocToDocx(docpath, savep):
        '''将doc转存为docx'''
        with semaphore:
            lock.acquire()
            try:
                # CoInitialize初始化,为线程和word对象创建一个套间，令其可以正常关联和执行
                pythoncom.CoInitialize()
                # 用DispatchEx()的方式启动MS Word或与当前已执行的MS Word建立连结
                word = DispatchEx('Word.Application')
                # 打开指定目录下doc文档
                doc = word.Documents.Open(docpath)
                savefile=savep+'\\' + os.path.basename(docpath)
                # 将打开的doc文档存储为docx
                doc.SaveAs(re.sub('.doc$', '.docx', savefile), FileFormat=12)
                # 关闭doc文档
                doc.Close()
            except:
                # 报错则输出报错文件
                print(docpath + '：无法打开')
            else:
                # 无报错输出转换完成
                print(os.path.basename(docpath) + " ： 转换完成")
            finally:
                # 关闭office程序
                word.Quit()
                # 释放资源
                pythoncom.CoUninitialize()
            lock.release()

    def DocToPdf(docpath, savep):
        '''将doc、docx转存为pdf'''
        with semaphore:
            lock.acquire()
            try:
                pythoncom.CoInitialize()
                word = DispatchEx('Word.Application')
                #word.Visible = 0
                #word.DisplayAlerts = 0
                doc = word.Documents.Open(docpath)
                savefile = savep + '\\' + os.path.basename(docpath)
                doc.SaveAs(re.sub('\.doc.*', '.pdf', savefile), FileFormat=17)
                doc.Close()
            except:
                print(docpath + '：无法打开')
            else:
                print(os.path.basename(docpath) + " ： 转换完成")
            finally:
                word.Quit()
                pythoncom.CoUninitialize()
            lock.release()


@Gooey(
    richtext_controls=True,  # 打开终端对颜色支持
    program_name="PDF简单处理",  # 程序名称
    encoding="utf-8",  # 设置编码格式，打包的时候遇到问题
    progress_regex=r"^progress: (\d+)%$"  # 正则，用于模式化运行时进度信息
 )
def main():
    description_msg = '简单的PDF处理，包含Office文档转换为PDF，PDF加密'
    parser = GooeyParser(description = description_msg)
    parser.add_argument('filepath', help='选择Word目录', widget='DirChooser', default=os.getcwd())
    parser.add_argument('savepath', help='选择保存目录', widget='DirChooser', default=os.getcwd())
    args = parser.parse_args()

    # 获取word文件目录绝对路径
    pre = args.filepath
    pdfpath = args.savepath
    for a, b, c in os.walk(pre):
        for file in c:
            if re.search('\.doc', file) != None:
                # 将doc转存为docx
                # threading.Thread(target=WordConvertToOther.DocToDocx, args=(pre + file,)).start()
                # 将doc、docx转存为pdf
                threading.Thread(target=WordConvertToOther.DocToPdf, args=(pre +'\\'+ file, pdfpath)).start()

if __name__ == '__main__':
    # 控制线程最大并发数为12
    semaphore = threading.Semaphore(12)
    # 线程锁
    lock = threading.Lock()
    main()

