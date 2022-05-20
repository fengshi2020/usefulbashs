import requests
import time
import os
import arrow
import pandas as pd
import pandas.io.formats.excel
from collections import OrderedDict
import yagmail
from xlsxwriter.utility import xl_rowcol_to_cell
import numpy as np

pandas.io.formats.excel.header_style = None
pd.set_option('display.max_colwidth', -1)  # 能显示的最大宽度, 否则to_html出来的地址就不全

#改这里
data_list = []

t = arrow.now()

endTms = t.format("YYYY-MM-DD")
stTms = t.shift(days=-7).format("YYYY-MM-DD")

headers = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}

# 改这里
userList = [
    {'xxx': [594]},
]


def get_info(Operator, userId, stTms, endTms):
    BasicInfo_url = f'http://xxx.com/api/v1/user/{userId}'
    SevenDay_url = f'http://xxx.com/api/v1/stat/overview?userId={userId}&dataType=USER&groupUnit=BY_DAYS&&&stTms={stTms}&endTms={endTms}'
    SevenDay_res = requests.get(SevenDay_url, headers=headers)
    SevenDay_data = SevenDay_res.json()['all']
    BasicInfo_res = requests.get(BasicInfo_url, headers=headers).json()
    info = OrderedDict()
    for i in SevenDay_data:
        info['运营'] = Operator
        info['名称'] = BasicInfo_res['member']['fullname']
        info['userId'] = i['after']['userId']
        info['公司名称'] = BasicInfo_res['info']['company']
        info['邮箱'] = BasicInfo_res['email']
        info['余额'] = round(float(BasicInfo_res['balance']) / 100, 2)
        info['日预算'] = float(BasicInfo_res['dailyBudget']) / 100
        info[i['after']['timeStamp']] = float('%.2f' % (i['after']['cost'] / 100))
    data_list.append(info)
    print(f"{Operator}数据处理完成！")


#改这里
def get_info_run():
    for item in userList:
        for Operator, list_ in item.items():
            for userId in list_:
                get_info(Operator, userId, stTms, endTms)
            print(f"处理{Operator}数据中。。。")


def gen_report():
    print("生成7日日报中。。。。")
    #改这里
    df = pd.DataFrame(data_list)

    yesterday = t.shift(days=-1).format("YYYY-MM-DD")
    beforday = t.shift(days=-2).format("YYYY-MM-DD")
    colums_days = [t.shift(days=i).format("YYYY-MM-DD") for i in range(-7,0)]
    df['花费同比'] = (df[yesterday]-df[stTms])/ df[stTms].apply(lambda x: x if x != 0 else 1)
    df['花费环比'] = (df[yesterday]-df[beforday])/ df[beforday].apply(lambda x: x if x != 0 else 1)
    df['花费七日均'] = df[colums_days].mean(1).round(2)
    df['预计可消费天数'] = round(df['余额'] / df['花费七日均'].apply(lambda x: x if x != 0 else 1),0)
    writer = pd.ExcelWriter('近7天报告' + time.strftime("%Y%m%d%H%M") + '.xlsx', engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='report')
    number_rows = len(df.index)
    workbook = writer.book

    workbook.formats[0].set_font_name("微软雅黑")

    worksheet = writer.sheets['report']
    worksheet.freeze_panes(1, 2)
    worksheet.autofilter(f'A1:T{number_rows+1}')

    worksheet.set_zoom(90)

    cell_format = workbook.add_format({'font_name':'微软雅黑','font_size':12,'bold': True,'bg_color': 'blue','font_color':'white','align':'center','valign':'vcenter','border':1})
    worksheet.set_row(0,None, cell_format)

    money_fmt = workbook.add_format({'num_format': '￥#,##0.00','font_name':'微软雅黑','border': 1})
    percent_fmt = workbook.add_format({'num_format': '0.00%','font_name':'微软雅黑','border': 1})

    # Total formatting
    total_fmt = workbook.add_format({'align': 'right', 'num_format': '￥#,##0.00',
                                    'bottom': 6,'font_name':'微软雅黑','border': 1,'bold':True})

    # Total percent format
    total_percent_fmt = workbook.add_format({'align': 'right', 'num_format': '0.00%',
                                             'bottom': 6,'font_name':'微软雅黑','border': 1,'bold':True})

    all_border_fmt = workbook.add_format({'border': 1,'font_name':'微软雅黑','align': 'right'})

    worksheet.set_column('A:A', 8,all_border_fmt)
    worksheet.set_column('B:B', 17, all_border_fmt)
    worksheet.set_column('C:C', 8, all_border_fmt)
    worksheet.set_column('D:D', 33,all_border_fmt)
    worksheet.set_column('E:E', 25,all_border_fmt)
    worksheet.set_column('F:F', 15)
    worksheet.set_column('G:G', 18)
    worksheet.set_column('H:O', 15)
    worksheet.set_column('P:R', 15)
    worksheet.set_column('S:S', 17,all_border_fmt)
    worksheet.set_column('T:T', 15,all_border_fmt)

    worksheet.set_column('F:O', 12, money_fmt)
    worksheet.set_column('R:R', 12, money_fmt)

    worksheet.set_column('P:Q', 12, percent_fmt)

    # Add total rows
    for column in range(5, 15):
        # Determine where we will place the formula
        cell_location = xl_rowcol_to_cell(number_rows + 1, column)
        # Get the range to use for the sum formula
        start_range = xl_rowcol_to_cell(1, column)
        end_range = xl_rowcol_to_cell(number_rows, column)
        # Construct and write the formula
        formula = "=SUM({:s}:{:s})".format(start_range, end_range)
        worksheet.write_formula(cell_location, formula, total_fmt)

    # Add a total label
    worksheet.write_string(number_rows + 1, 4, "总计", total_fmt)

    percent_formula_same = "=(N{0}-H{0})/H{0}".format(number_rows + 2)
    worksheet.write_formula(number_rows + 1, 15, percent_formula_same, total_percent_fmt)

    percent_formula_ring = "=(N{0}-M{0})/M{0}".format(number_rows + 2)
    worksheet.write_formula(number_rows + 1, 16, percent_formula_ring, total_percent_fmt)

    formula_mean7 = "=AVERAGE(H{0}:O{0})".format(number_rows + 2)
    worksheet.write_formula(number_rows + 1, 17, formula_mean7,total_fmt)

    formula_predays = "=ROUND(F{0}/R{0},0)".format(number_rows + 2)
    worksheet.write_formula(number_rows + 1, 18, formula_predays)

    color_range_p = "P2:P{}".format(number_rows + 1)
    color_range_q = "Q2:Q{}".format(number_rows + 1)
    color_range_r = "R2:R{}".format(number_rows + 1)
    color_range_s = "S2:S{}".format(number_rows + 2)
    color_range_f = "F2:F{}".format(number_rows + 1)
    color_range_g = "G2:G{}".format(number_rows + 1)

    format1 = workbook.add_format({'bg_color': '#FFC7CE',
                                   'font_color': '#9C0006',
                                   'border': 1})

    format2 = workbook.add_format({'bg_color': '#C6EFCE',
                                   'font_color': '#006100',
                                   'border':1})

    format3 = workbook.add_format({'bg_color': '#FFC7CE',
                                   'font_color': '#9C0006',
                                   'border': 1})

    worksheet.conditional_format(color_range_p, {'type': 'top',
                                               'value': '10',
                                               'format': format1})

    worksheet.conditional_format(color_range_p, {'type': 'bottom',
                                               'value': '10',
                                               'format': format2})

    worksheet.conditional_format(color_range_q, {'type': 'top',
                                               'value': '10',
                                               'format': format1})

    worksheet.conditional_format(color_range_q, {'type': 'bottom',
                                               'value': '10',
                                               'format': format2})

    worksheet.conditional_format(color_range_r, {'type': 'top',
                                               'value': '10',
                                               'format': format1})

    worksheet.conditional_format(color_range_r, {'type': 'bottom',
                                               'value': '10',
                                               'format': format2})

    worksheet.conditional_format(color_range_s, {'type': 'cell',
                                                 'criteria': '<=',
                                                 'value': '7',
                                                 'format': format3})

    worksheet.conditional_format(color_range_f, {'type': 'cell',
                                                 'criteria': '<=',
                                                 'value': '10000',
                                                 'format': format3})

    worksheet.conditional_format(color_range_g, {'type': 'cell',
                                                 'criteria': '<=',
                                                 'value': '1000',
                                                 'format': format3})

    worksheet.conditional_format(f'H2:H{number_rows+1}', {'type': 'data_bar', 'bar_solid': True, 'format': money_fmt})
    worksheet.conditional_format(f'I2:I{number_rows+1}', {'type': 'data_bar', 'bar_solid': True, 'format': money_fmt})
    worksheet.conditional_format(f'J2:J{number_rows+1}', {'type': 'data_bar', 'bar_solid': True, 'format': money_fmt})
    worksheet.conditional_format(f'K2:K{number_rows+1}', {'type': 'data_bar', 'bar_solid': True, 'format': money_fmt})
    worksheet.conditional_format(f'L2:L{number_rows+1}', {'type': 'data_bar', 'bar_solid': True, 'format': money_fmt})
    worksheet.conditional_format(f'M2:M{number_rows+1}', {'type': 'data_bar', 'bar_solid': True, 'format': money_fmt})
    worksheet.conditional_format(f'N2:N{number_rows+1}', {'type': 'data_bar', 'bar_solid': True, 'format': money_fmt})
    worksheet.conditional_format(f'O2:O{number_rows+1}', {'type':'data_bar','bar_solid': True,'format': money_fmt})

    #迷你图-1
    worksheet.write('T1','趋势迷你图')
    for row in range(2,number_rows+2):
        worksheet.add_sparkline('T'+str(row), {'range': 'report!H{0}:N{0}'.format(row),'markers': True})

    # 折线图
    chart_line = workbook.add_chart({'type': 'line'})
    chart_line.add_series({
        'categories': '=report!$H$1:$N$1',
        'values': f'=report!$H${number_rows+2}:$N${number_rows+2}',
    })
    chart_line.set_legend({'none': True})

    column_chart = workbook.add_chart({'type': 'column'})
    column_chart.add_series({
        'categories': '=report!$H$1:$N$1',
        'values': f'=report!$H${number_rows+2}:$N${number_rows+2}',
    })
    chart_line.combine(column_chart)
    chart_line.set_title({ 'name': '总7日走势图'})
    worksheet.insert_chart(f'D{number_rows+3}', chart_line)

    #透视表
    pivot = pd.pivot_table(df,values=colums_days + [endTms],index=['运营'],aggfunc=np.sum)
    pivot['花费同比'] = (pivot[yesterday]-pivot[stTms])/ pivot[stTms].apply(lambda x: x if x != 0 else 1)
    pivot['花费环比'] = (pivot[yesterday]-pivot[beforday])/ pivot[beforday].apply(lambda x: x if x != 0 else 1)
    pivot['花费七日均'] = pivot[colums_days].mean(1).round(2)
    pivot.to_excel(writer,sheet_name='report',startrow=number_rows+3,startcol=6)
    pivot_rows = len(pivot.index)
    worksheet.write(f'T{number_rows+4}', '趋势迷你图')
    for row in range(number_rows+5, number_rows + 4 + pivot_rows +1):
        worksheet.add_sparkline('T' + str(row), {'range': 'report!H{0}:N{0}'.format(row), 'markers': True})

    writer.save()

    print('7日报生成完成！')
    return df,pivot


def get_html_msg(df,pivot):
    #1. 构造html信息
    df.drop('公司名称', axis=1, inplace=True)
    df_html = df.to_html(escape=False,index=False)

    pivot_html = pivot.to_html(escape=False)
    df_html = df_html.replace("\n", "")
    pivot_html = pivot_html.replace("\n", "")

    # html = html.replace("\n", "") 表格部分

    head = \
    '''
    <head>
        <meta charset="utf-8">
        <STYLE TYPE="text/css" MEDIA=screen>

            table.dataframe {
                border-collapse: collapse;
                border: 2px solid #a19da2;
                /*居中显示整个表格*/
                margin: auto;
            }

            table.dataframe thead {
                border: 2px solid #91c6e1;
                background: #f1f1f1;
                padding: 10px 10px 10px 10px;
                color: #333333;
            }

            table.dataframe tbody {
                border: 2px solid #91c6e1;
                padding: 10px 10px 10px 10px;
            }

            table.dataframe tr {

            }

            table.dataframe th {
                vertical-align: top;
                font-size: 14px;
                padding: 10px 10px 10px 10px;
                color: #105de3;
                font-family: 微软雅黑;
                text-align: center;
            }

            table.dataframe td {
                text-align: center;
                padding: 10px 10px 10px 10px;
            }

            body {
                font-family: 微软雅黑;
            }

            h1 {
                color: #5db446
            }

            div.header h2 {
                color: #0002e3;
                font-family: 微软雅黑;
            }

            h3 {
                font-size: 22px;
                background-color: rgba(0, 2, 227, 0.71);
                text-shadow: 2px 2px 1px #de4040;
                color: rgba(239, 241, 234, 0.99);
                line-height: 1.5;
            }

            h4 {
                color: #e10092;
                font-family: 微软雅黑;
                font-size: 20px;
                text-align: center;
            }

        </STYLE>
        </head>
    '''

    # 构造模板的附件（100）
    body = \
    """
        <body>
        <div align="center" class="header">
            <!--标题部分的信息-->
            <h1 align="center">您好，以下为近7日日报，详细内容请看附件！</h1>
        </div>
        <hr>
        <div class="content">
            <!--正文内容-->
            <div>
                <h4>运营汇总报告</h4>
                {1}
                <h4>客户详细报告</h4>
                {0}   
            </div>
            <hr>
            <p style="text-align: center">
                —— 本次报告完 ——
            </p>
        </div>
        </body>
    """.format(df_html,pivot_html)

    foot = \
    '''
    <br/>
    <p>
    xx营<br/>
    Email: xx.com<br/>
    MP: +86 xxx<br/>
    Address: xxxx<br/>
    </p>
    '''
    html_msg = "<html>" + head + body + foot + "</html>"
    html_msg = html_msg.replace("\n", "")
    # 这里是将HTML文件输出，作为测试的时候，查看格式用的，正式脚本中可以注释掉
    fout = open('./t4.html', 'w', encoding='UTF-8', newline='')
    fout.write(html_msg)
    return html_msg


def send_seven_daily_mail(html_msg):
    print('正在准备发送日报邮件...')
    yesterday = t.shift(days=-1).format("YYYY-MM-DD")
    # 链接邮箱服务器
    yag = yagmail.SMTP(user="xxx.com", password="ckxxxx", host='smtp.gmail.com')

    staff = ['xxxx']

    # 发送邮件
    print('正在给部门小伙伴发送7日日报.....')
    yag.send(to=staff,
             subject=yesterday + '七日日报', contents=html_msg,
             attachments='近7天报告' + time.strftime("%Y%m%d%H%M") + '.xlsx')
    print('xx7日日报发送成功!')


def remove_file():
    all_file = os.listdir('./')
    for file in all_file:
        if file.endswith('.xlsx'):
            os.remove(file)


def main():
    #清除历史遗留excel文件
    remove_file()
    get_info_run()
    df, pivot = gen_report()
    html_msg = get_html_msg(df,pivot)
    send_seven_daily_mail(html_msg)


if __name__ == '__main__':
    main()