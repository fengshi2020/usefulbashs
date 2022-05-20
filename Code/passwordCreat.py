# -*- coding: utf-8 -*-
import random
import string
import time
import os

from gooey import GooeyParser, Gooey


@Gooey(encoding='utf8', program_name='强密码生成器', program_description='用于生成包含大小写字母、数字、特殊符号的强密码', optional_cols=1)
def stronge_password():
    parse = GooeyParser(description='用于生成包含大小写字母、数字、特殊符号的强密码')
    parse.add_argument('数量', widget='TextField', default='8')
    parse.add_argument('长度', widget='TextField', default='12')
    parse.add_argument('数字', widget='Dropdown', choices={"是": 1, "否": 0}, default='是')
    parse.add_argument('大写字母', widget='Dropdown', choices={"是": 1, "否": 0}, default='是')
    parse.add_argument('小写字母', widget='Dropdown', choices={"是": 1, "否": 0}, default='是')
    parse.add_argument('特殊符号', widget='Dropdown', choices={"是": 1, "否": 0}, default='是')
    parse.add_argument('去除词', widget='TextField', default='None')
    args = parse.parse_args()
    if args.数字 == "是":
        digits = list(string.digits.strip())
    else:
        digits = []
    if args.大写字母:
        upper = list(string.ascii_uppercase)
    else:
        upper = []
    if args.小写字母:
        lower = list(string.ascii_lowercase)
    else:
        lower = []
    if args.特殊符号:
        symbol = list("~!@#$%^&*()_+{}[]/?")
    else:
        symbol = []
    if args.去除词 != 'None':
        stop_word = list(args.去除词.strip())
    else:
        stop_word = []
    letter_list = digits + upper + lower + symbol
    for item in stop_word:
        if item in letter_list:
            letter_list.remove(item)
    passlist = '___________________________\n' + time.strftime("%Y-%m-%d %H:%M:%S") + '\n'
    print(passlist)
    for i in range(int(args.数量)):
        result = ''.join([random.choice(letter_list) for i in range(int(args.长度))])
        print(result)
        passlist += result + '\n'
    with open('passList.txt','a+') as f:
        f.write(passlist)

if __name__ == '__main__':
    stronge_password()