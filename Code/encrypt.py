# _*_ coding:utf-8 _*_
import win32com.client as cl
#from win32com.client import DispatchEx
import datetime
import random
import string
import os
import py7zr
from gooey import GooeyParser, Gooey

def outlook(contentlist):
    olook = cl.DispatchEx("outlook.Application")  # 固定写法
    mail = olook.CreateItem(0)  # 固定写法
    #mail.Display = False
    mail.To = contentlist.get('to')  # 收件人
    mail.CC = contentlist.get('cc')  # 抄送人
    mail.Subject = contentlist.get('sub') + '相关文件加密密码，请查收！'  # 邮件主题
    body_html = ''
    body_html = body_html + '<html><body style="font-family:Meiryo UI;font-size:10pt;color:black"><p>您好！<p>'
    body_html = body_html + contentlist.get('sub') + '相关文件加密密码如下请查收，加密文件将通过其他途径发出，请妥善保管文件和密码，谢谢！<br/><p><div style="font-family:Meiryo UI;font-size:12pt;color:red">'
    body_html = body_html + contentlist.get(
        'pass') + '</div><p>建议将加密密码复制粘贴使用，推荐使用7-zip解压缩<a href="http://www.7-zip.org.cn/">7-zip官网</a>，如有任何疑问欢迎通过本邮箱联系<p>'
    body_html = body_html + '<div style="font-family:Meiryo UI;font-size:7pt">恭祝顺祺！<br/>平安国际智慧城市科技股份有限公司<br/></div></body></html>'
    mail.HTMLBody = body_html
    try:
        mail.Send()  # 发送
        return True
    except Exception as e:
        print(e.args)
        return False

def encrypt_file(pwd, file_path, zip_path,filename):
    print('文件加密中，请稍后……………………\n')
    password = pwd  # 密码
    filepath = file_path
    dir_na = zip_path+'\\'+filename+'.zip'  # 压缩后路径
    folder_path = os.path.abspath(filepath)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    z7z_file_path = os.path.abspath(os.path.join(folder_path, f'{dir_na}'))
    if os.path.isdir(filepath):
        with py7zr.SevenZipFile(z7z_file_path, mode='w', password=password) as zf:
            zf.set_encrypted_header(True)
            for dir_path, dir_names, file_names in os.walk(filepath):
                for filename in file_names:
                    fpath = dir_path.replace(filepath, '')
                    file_path = os.path.join(dir_path, filename)
                    filename = os.path.join(fpath, filename)
                    zf.write(file_path, arcname=filename)
    elif os.path.isfile(filepath):
        filename = os.path.basename(filepath)
        with py7zr.SevenZipFile(z7z_file_path, mode='w', password=password) as zf1:
            zf1.set_encrypted_header(True)
            zf1.write(filepath, arcname=filename)
    print('文件加密成功！')

@Gooey(
    richtext_controls=True,  # 打开终端对颜色支持
    language='chinese',
    header_show_title=False,
    required_cols=1,
    program_name="敏感数据加密工具V1.0",  # 程序名称
    encoding="utf-8",  # 设置编码格式，打包的时候遇到问题
    default_size=(620, 500)
)
def mainUI():
    parse = GooeyParser(description='敏感数据加密工具，输入需要加密文件/文件夹路径，客户邮箱，将能自动完成文件加密，密码发送。需要将加密后的文件分享给业务团队。')
    parse.add_argument('项目名称', help='请输入项目名称及文件名称，示例：XX项目后端源代码')
    parse.add_argument('源文件路径', widget='DirChooser', help='请选择需要加密的文件夹，建议需要加密的文件放在一个文件夹中！', default=os.getcwd())
    parse.add_argument('保存文件路径', widget='DirChooser', default=os.getcwd())
    parse.add_argument('发送列表', help='请输入接受加密密码的客户邮箱，多个邮箱地址以英文‘,’分隔')
    parse.add_argument('抄送列表', help='请输入抄送人的邮箱（主要是外发审核人），多个邮箱地址以英文‘;’分隔')
    args = parse.parse_args()
    conlist={
        'sub': args.项目名称,
        'to': args.发送列表,
        'cc': args.抄送列表,
        'pass':createPass()
    }
    encrypt_file(conlist.get('pass'),args.源文件路径,args.保存文件路径,conlist.get('sub'))
    if outlook(conlist):
        print('邮件发送成功！')
    else:
        print('邮件发送失败，请重试！')
    writeLogs(conlist)

def createPass():
    #创建加密密码
    digits = list(string.digits.strip())
    upper = list(string.ascii_uppercase)
    lower = list(string.ascii_lowercase)
    symbol = list("~!@#$%^&*()_+{}[]/?")
    letter_list = digits + upper + lower + symbol
    pwd = ''.join([random.choice(letter_list) for i in range(16)])
    return pwd

def writeLogs(conlist):
    logfile = os.getcwd() + '\\log.txt'
    content = '\n\n\n++++++++++++++++++++++++++++++++++\n'
    content = content + str(datetime.datetime.now()) + '\n'
    content = content + '项目及文件名：'+ conlist.get('sub') +'\n' + '收件人列表：' + conlist.get('to')+'\n'+'抄送人列表：'+conlist.get('cc')+'\n加密密码：'+conlist.get('pass')+'\n++++++++++++++++++++++++++++++++++'
    with open(logfile,'a+',encoding='UTF-8') as log:
        log.write(content)
    print('加密记录成功！')

if __name__ == '__main__':
    mainUI()