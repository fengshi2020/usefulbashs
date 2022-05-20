#_*_ coding:utf-8 _*_
import json
import requests
import os
import pandas as pd
import re
from gooey import Gooey, GooeyParser

baseurl = 'http://itunes.apple.com/search?term='
headres = {
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'accept-encoding': 'gzip, deflate, br',
    'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
    'cache-control': 'max-age=0',
    'cookie': 's_fid=75EB3767DE7C88D8-180C49C529498645; s_vi=[CS]v1|30D3942C502598F7-600000189A9395A0[CE]; dslang=CN-ZH; acn01=g9lTbavfLbNAYZ/4Uin1sGbYSbGkarIs37OUALT2cZ0AEA3VmRiwnw==',
    'dnt': '1',
    'sec-ch-ua': '" Not;A Brand";v="99", "Microsoft Edge";v="97", "Chromium";v="97"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'none',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36 Edg/97.0.1072.55'
}

def getResult(keyword,path):
    url = baseurl+keyword+'&entity=software&Country=CN&limit=200'
    r = requests.post(url)
    if r.status_code == 200:
        alsysResult(keyword, r.text, path)

def data_clean(text):
    # 清洗excel中的非法字符，都是不常见的不可显示字符，例如退格，响铃等
    ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
    text = ILLEGAL_CHARACTERS_RE.sub(r'', text)
    return text
'''
df = df.fillna('').astype(str)
for col in df.columns:
    df[col] = df[col].apply(lambda x: data_clean(x))
df.to_excel(base_path + 'result.xlsx', index=False)
'''

def alsysResult(word, con, path):
    if not con:
        print('未获取到内容！')
    else:
        rows = json.loads(con)
        print(word+'成功搜索到：'+str(rows['resultCount'])+'个结果！')
        results = json.dumps(rows['results'], ensure_ascii=False, indent=10)
        #content = json.loads(results)
        jsonfile = word+'.json'
        with open(jsonfile,'w', encoding="UTF-8") as f:
            f.write(results)
        df = pd.read_json(jsonfile, encoding='utf-8')
        df = df.fillna('').astype(str)
        for col in df.columns:
            df[col] = df[col].apply(lambda x: data_clean(x))
        savefile = path + '\\' + word + '.xlsx'
        df.to_excel(savefile)
        print('结果保存在：' + savefile)
        if os.path.exists(jsonfile):
            os.remove(jsonfile)

@Gooey(
    richtext_controls=True,  # 打开终端对颜色支持
    language='chinese',
    header_show_title=False,
    required_cols=1,
    program_name="IOS应用搜索爬虫V1.0",  # 程序名称
    encoding="utf-8",  # 设置编码格式，打包的时候遇到问题
    default_size=(550, 460)
)
def main():
    parser = GooeyParser(description="一款输入关键字就能查找IOS应用详细情况的小应用！")
    parser.add_argument('output',
                        metavar='保存目录',
                        help='请选择文件保存目录！',
                        widget='DirChooser',
                        default=os.getcwd()
                        )
    parser.add_argument('keywords',
                        metavar='搜索关键字',
                        help='请输入搜索关键字，如有多个以英文逗号’，‘间隔！'
                        )
    args = parser.parse_args()
    savePath = args.output
    keywords = args.keywords
    for keyword in keywords.split(','):
        getResult(keyword, savePath)


if __name__ == '__main__':
    main()


