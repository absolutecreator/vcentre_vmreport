#!/usr/bin/env python3

# Script by spirin-li.
# Telegram: @absolutecreator

# Основные библиотеки
import os
from os.path import basename
from os.path import getmtime
import shutil
import sys
import warnings
import time
from datetime import datetime, timedelta
import csv
import atexit
import base64
from io import BytesIO
import getopt
import glob
import re
import requests

# Библиотеки для построения кадра данных Pandas и графиков
import pandas as pd
import matplotlib.pyplot as plt
from adjustText import adjust_text
import seaborn as sns

# Библиотеки для подключения к vcentre
from pyVim.connect import SmartConnectNoSSL, Disconnect
from pyVmomi import vim, vmodl

# Библиотеки для отправки почты
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formatdate

global Date
global csvHeader

pd.options.display.float_format = '{:.2f}'.format
warnings.simplefilter(action='ignore', category=FutureWarning)

login = "script@vmcod.local"
passwd = "script!"
dr = '/opt/vmware-report/'
bakDir = dr + 'bak/'
CsvDir = dr + 'csv/'
lastCsv = CsvDir + 'last/'
XlsxDir = dr + 'xlsx/'

frmt = '%d.%m.%Y'
timeStr = time.strftime(frmt)

Date = str(timeStr)
if not os.path.exists(dr):
    os.makedirs(dr)
if not os.path.exists(lastCsv):
    os.makedirs(lastCsv)
if not os.path.exists(CsvDir):
    os.makedirs(CsvDir)
if not os.path.exists(bakDir):
    os.makedirs(bakDir)
if not os.path.exists(XlsxDir):
    os.makedirs(XlsxDir)
csv_file = lastCsv + 'vmware_report_' + timeStr + '.csv'
xlsx_file = XlsxDir + 'vmware_report_' + timeStr + '.xlsx'
html_file = dr + 'web_vmware_report_' + timeStr + '.html'

csvHeader = [
    "Date,Name,IP,vCPU,MEM(GB),Guest_HDD(GB),Guest_UsedSpace(GB),Vmdk_Used_Space(GB),Project,Platform,PowerState,GuestOS,Folder,Owner,VMHost,Datastore"]


#      0   1    2   3      4        5         6                   7                     8        9          10      11     12     13    14     15


# Функция создания HTML + Style content
def htmlBuild():
    hd = '''<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> 
        <html xmlns="http://www.w3.org/1999/xhtml">
        <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        </head>
        '''
    pt = "<body>"

    st = [{'selector': 'th', 'props': [
        ('text-align', 'left'),
        ('font-size', '15px'),
        ('padding', '5px')]},
          {'selector': 'td', 'props': [
              ('font-size', '13px'),
              ('padding', '5px')]}]

    cs = '''<style>
        .center img {
            display: block;
            margin-left: 550px;
        }
        .images-container {
            display: flex;
            justify-content: center;
        }
        .images-container img {
            margin-left: 5px;
            margin-right: 5px;
        }
        th {
            font-size: 14px;
            padding: 5px;
            # min-width: 90px;
        }
        td {
            background: #F5D7BF;
            font-size: 15px;
            padding: 5px;
        }
        .table_sort table {
            border-collapse: collapse;
            border-radius: 20px;
        }

        .table_sort th {
            cursor: pointer;
            border-radius: 5px;
            background-color:#3bb3e0;
            font-weight:bold;
            text-decoration:none;
            color:#59646e;
            text-shadow:rgba(255,255,255,0.6) 1px 1px 0;
            position:relative;
            padding:10px 20px;
            padding-right:50px;
            background-image: linear-gradient(bottom, rgb(169, 199, 209) 0%, rgb(176, 217, 232) 100%);
            background-image: -o-linear-gradient(bottom, rgb(169, 199, 209) 0%, rgb(176, 217, 232) 100%);
            background-image: -moz-linear-gradient(bottom, rgb(169, 199, 209) 0%, rgb(176, 217, 232) 100%);
            background-image: -webkit-linear-gradient(bottom, rgb(169, 199, 209) 0%, rgb(176, 217, 232) 100%);
            background-image: -ms-linear-gradient(bottom, rgb(169, 199, 209) 0%, rgb(176, 217, 232) 100%);
            background-image: -webkit-gradient(
            linear,
            left bottom,
            left top,
            color-stop(0, rgb(169, 199, 209)),
            color-stop(1, rgb(176, 217, 232))
            );
            -webkit-border-radius: 5px;
            -moz-border-radius: 5px;
            -o-border-radius: 5px;
            border-radius: 5px;
            -webkit-box-shadow: inset 0px 1px 0px #2ab7ec, 0px 3px 0px 0px #156785, 0px 3px 1px #999;
            -moz-box-shadow: inset 0px 1px 0px #2ab7ec, 0px 3px 0px 0px #156785, 0px 3px 1px #999;
            -o-box-shadow: inset 0px 1px 0px #2ab7ec, 0px 3px 0px 0px #156785, 0px 3px 1px #999;
            box-shadow: inset 0px 1px 0px #2ab7ec, 0px 2px 0px 0px #156785, 0px 3px 1px #999;
        }

        .table_sort th {
            width: 150px;
            height: 40px;
            text-align: center;
        }

        .table_sort th:hover, .table_sort th:focus{
            -webkit-animation: linear 1.2s light infinite;
            -moz-animation: linear 1.2s light infinite;
            -o-animation: linear 1.2s light infinite;
            animation: linear 1.2s light infinite;
        }
            @-webkit-keyframes light{
                0%   { color: #ddd; text-shadow: 0px -1px 0px #000; }
                50%   { color: #fff; text-shadow: 0px -1px 0px #444, 0px 0px 5px #ffd, 0px 0px 8px #fff; }
                100% { color: #ddd; text-shadow: 0px -1px 0px #000; }
            }
            @-moz-keyframes light{
                0%   { color: #ddd; text-shadow: 0px -1px 0px #000; }
                50%   { color: #fff; text-shadow: 0px -1px 0px #444, 0px 0px 5px #ffd, 0px 0px 8px #fff; }
                100% { color: #ddd; text-shadow: 0px -1px 0px #000; }
            }
            @-o-keyframes light{
                0%   { color: #ddd; text-shadow: 0px -1px 0px #000; }
                50%   { color: #fff; text-shadow: 0px -1px 0px #444, 0px 0px 5px #ffd, 0px 0px 8px #fff; }
                100% { color: #ddd; text-shadow: 0px -1px 0px #000; }
            }
            @keyframes light{
                0%   { color: #ddd; text-shadow: 0px -1px 0px #000; }
                50%   { color: #fff; text-shadow: 0px -1px 0px #444, 0px 0px 5px #ffd, 0px 0px 8px #fff; }
                100% { color: #ddd; text-shadow: 0px -1px 0px #000; }
            }
        .table_sort th:active{
            color: #fff;
            text-shadow: 0px -1px 0px #444,0px 0px 5px #ffd,0px 0px 8px #fff;
            box-shadow: 0px 1px 0px #666,0px 2px 0px #444,0px 2px 2px rgba(0, 0, 0, .9);
            -webkit-transform: translateY(3px);
            -moz-transform: translateY(3px);
            -o-transform: translateY(3px);
            transform: translateY(3px);
            -webkit-animation: none;
            -moz-animation: none;
            -o-animation: none;
            animation: none;
        }


        .table_sort tbody tr:nth-child(even) {
            background: #e3e3e3;
        }

        th.sorted[data-order="1"],
        th.sorted[data-order="-1"] {
            position: relative;
        }

        th.sorted[data-order="1"]::after,
        th.sorted[data-order="-1"]::after {
            right: 8px;
            position: absolute;
            bottom:0;
            left: 47%;
            -webkit-animation: linear 1.2s light infinite;
            -moz-animation: linear 1.2s light infinite;
            -o-animation: linear 1.2s light infinite;
            animation: linear 1.2s light infinite;     
        }

        th.sorted[data-order="-1"]::after {
            content: "▼"
        }

        th.sorted[data-order="1"]::after {
            content: "▲"
        }

        table {
            font-family: "Lucida Sans Unicode", "Lucida Grande", Sans-Serif;
            text-align: left;
            font-size: 17px;
            border-collapse: separate;
            border-spacing: 5px;
            background: #ECE9E0;
            color: #656665;
            border: 5px solid #ECE9E0;
            border-radius: 20px;
        }

        section {
            font-family: "Lucida Sans Unicode", "Lucida Grande", Sans-Serif;
            'font-size', '15px'
        }

        .table_sort td {
            font-size: 15px;
        }
        </style>
        '''
    # JavaScript
    s_s = '''
        <script>
        document.addEventListener('DOMContentLoaded', () => {

        const getSort = ({ target }) => {
            const order = (target.dataset.order = -(target.dataset.order || -1));
            const index = [...target.parentNode.cells].indexOf(target);
            const collator = new Intl.Collator(['en', 'ru'], { numeric: true });
            const comparator = (index, order) => (a, b) => order * collator.compare(
                a.children[index].innerHTML,
                b.children[index].innerHTML
            );

            for(const tBody of target.closest('table').tBodies)
                tBody.append(...[...tBody.rows].sort(comparator(index, order)));

            for(const cell of target.parentNode.cells)
                cell.classList.toggle('sorted', cell === target);
        };

        document.querySelectorAll('.table_sort thead').forEach(tableTH => tableTH.addEventListener('click', () => getSort(event)));

        });
        </script>
        '''
    ft = "</body></html>"
    return hd, pt, st, cs, s_s, ft


# Функции обработки информации vCentre
def get_obj(vc, root, vim_type):
    container = vc.content.viewManager.CreateContainerView(root, vim_type, True)
    view = container.view
    container.Destroy()
    return view


def get_filter_spec(containerView, objType, path):
    traverse_spec = vmodl.query.PropertyCollector.TraversalSpec()
    traverse_spec.name = 'traverse'
    traverse_spec.path = 'view'
    traverse_spec.skip = False
    traverse_spec.type = vim.view.ContainerView

    obj_spec = vmodl.query.PropertyCollector.ObjectSpec()
    obj_spec.obj = containerView
    obj_spec.skip = True
    obj_spec.selectSet.append(traverse_spec)

    prop_spec = vmodl.query.PropertyCollector.PropertySpec()
    prop_spec.type = objType
    prop_spec.pathSet = path

    return vmodl.query.PropertyCollector.FilterSpec(propSet=[prop_spec],
                                                    objectSet=[obj_spec])


def process_result(result, objects):
    for o in result.objects:
        if o.obj not in objects:
            objects[o.obj] = {}
        for p in o.propSet:
            objects[o.obj][p.name] = p.val


def collect_properties(vc, root, vim_type, props):
    objects = {}
    view_mgr = vc.content.viewManager
    container = view_mgr.CreateContainerView(root, [vim_type], True)
    try:
        filter_spec = get_filter_spec(container, vim_type, props)
        options = vmodl.query.PropertyCollector.RetrieveOptions()
        pc = vc.content.propertyCollector
        result = pc.RetrievePropertiesEx([filter_spec], options)
        process_result(result, objects)
        while result.token is not None:
            result = pc.ContinueRetrievePropertiesEx(result.token)
            process_result(result, objects)
    finally:
        container.Destroy()
    return objects


# Подсвечивание ячейки в таблице
def highlight(series):
    color = 'background-color:#CFFED4;'
    default = ''
    return [color if e == 'poweredOn' else default for e in series]


# Функция построения графика и вывод его в изображение base64
def plotting(dfi, tl, pt):
    dfi = dfi.round(2)
    dfi.shape
    fig, ax = plt.subplots(figsize=(10, 6))
    if 'Bar' in pt:
        dfi.plot.bar(figsize=(10, 6),
                     title=tl,
                     lw=3,
                     fontsize=10,
                     grid=True, ax=ax, color=colors).legend(loc='best')
    else:
        dfi.plot(figsize=(10, 6),
                 title=tl,
                 lw=3,
                 fontsize=10,
                 grid=True, ax=ax, marker='o', color=colors).legend(loc='best')
    annotations = []
    for y in ax.patches:
        value = y.get_height()
        text = f'{value:}'
        text_x = y.get_x() + y.get_width() / 2
        text_y = y.get_y() + value
        color = y.get_facecolor()
        annotations += [ax.text(text_x, text_y, text, ha='center', va='bottom', color=color,
                                size=12)]

    ax.legend(bbox_to_anchor=(1.0, 1.0))
    plt.xticks(rotation=30)
    adjust_text(annotations, ha='left')
    dfi = BytesIO()
    plt.savefig(dfi, format='png', dpi=60)
    b64 = base64.b64encode(dfi.getvalue()).decode('utf-8')
    return b64


# Получение данных из vCentre
def main(vcr):
    print(vcr)
    vc = SmartConnectNoSSL(host=vcr, user=login, pwd=passwd, port=443)
    atexit.register(Disconnect, vc)
    data = []
    vms = collect_properties(vc, vc.content.rootFolder, vim.VirtualMachine,
                             ['name', 'summary', 'guest', 'customValue', 'runtime', 'datastore'])

    # Получение данных кастомных полей (атрибутов)
    Bdy2 = ''
    for vm in vms.values():
        Project = 'NO_NAME'
        Owner = ''
        Datastore = ''
        Os = ''
        Mem = ''
        Guest_HDD = ''
        Guest_UsedSpace = ''
        Vmdk_Used_Space = ''
        if vm['summary'].runtime.host is None:
            Bdy2 = "Не все ВМ выгружены. Возможно, идет обслуживание. Если скрипт запущен ночью, перезапустите его в дневное время."
            continue

        # Кастомные поля (атрибуты)
        if "hx-dl-esxi-sas" in vm['summary'].runtime.host.name:
            Platform = "HX-DL-SAS"
        elif "hx-dl-esxi-ssd" in vm['summary'].runtime.host.name:
            Platform = "HX-DL-SSD"
        elif "hx-esxi-ssd" in vm['summary'].runtime.host.name:
            Platform = "HX-SSD"
        elif "hx-esxi-sas" in vm['summary'].runtime.host.name:
            Platform = "HX-SAS"
        else:
            Platform = "OLD"
        fields = vc.content.customFieldsManager.field
        for field in fields:
            if field.name == "Owner":
                attr_key = field.key
                for fName in vm['customValue']:
                    if fName.key == attr_key:
                        Owner = fName.value
            elif field.name == "Project":
                attr_key = field.key
                for fName in vm['customValue']:
                    if fName.key == attr_key:
                        Project = fName.value
            match1 = re.search(r"\[(.*?)]", str(vm['summary'].config.vmPathName))
            if match1:
                Datastore = match1.group(1)
        Os = str(vm['guest'].guestFullName)

        # Форматирование дробных значений
        Mem = f"{round(int(vm['summary'].config.memorySizeMB)) / 1024:.0f}"
        Guest_HDD = f"{round(sum([int(d.capacity) for d in vm['guest'].disk])) / 1024 ** 3:.2f}"
        Guest_UsedSpace = f"{round(sum([int(gd.capacity - gd.freeSpace) for gd in vm['guest'].disk])) / 1024 ** 3:.2f}"
        Vmdk_Used_Space = f"{round(int(vm['summary'].storage.committed + vm['summary'].storage.uncommitted)) / 1024 ** 3:.2f}"
        data.append(
            {"Name": vm['name'], "IP": vm['guest'].ipAddress, "vCPU": vm['summary'].config.numCpu, "Mem(GB)": Mem,
             "Guest_HDD(GB)": Guest_HDD, "Guest_UsedSpace(GB)": Guest_UsedSpace,
             "Vmdk_Used_Space(GB)": Vmdk_Used_Space,
             "Project": Project.strip("'"),
             "Platform": Platform, "PowerState": vm['runtime'].powerState, "GuestOS": Os,
             "Folder": vm['summary'].vm.parent.name, "Owner": Owner.strip("'").replace(",", ""),
             "VMHost": vm['summary'].runtime.host.name, "Datastore": Datastore})
    vmList = data

    # Создание csv
    if csv_file:
        with open(csv_file, 'a', newline='', encoding="cp1251", errors="ignore") as csvfile:
            wrtr = csv.writer(csvfile, delimiter='\t', quoting=csv.QUOTE_NONE)
            for vm in vmList:
                list_value = list(vm.values())
                row = f"{Date},{list_value[0]},{list_value[1]},{list_value[2]},{list_value[3]},{list_value[4]},{list_value[5]},{list_value[6]},{list_value[7]},{list_value[8]},{list_value[9]},{list_value[10]},{list_value[11]},{list_value[12]},{list_value[13]},{list_value[14]}"
                wrtr.writerow([row])
        file.close()
    shutil.copy2(csv_file, CsvDir, follow_symlinks=True)
    return vc, Bdy2


# Функция обработки параметров запуска
def params(argv):
    try:
        opts, args = getopt.getopt(argv, "hc:r:", ["diffperiod=", "cleanrepover="])
    except getopt.GetoptError:
        print(
            'Error. Script usage only with params: -r <number of last report files> -c <clean reports greater date by days>')
        sys.exit(2)
    if not opts:
        print(
            'Error. Script usage only with params: -r <number of last report files> -c <clean reports greater date by days>')
        sys.exit(2)
    diffperdays = ''
    cleanrepovdays = ''
    for opt, arg in opts:
        if opt == '-h':
            print(
                'Error. Script usage only with params: -r <number of last report files> -c <clean reports greater date by days>')
            sys.exit()
        if opt in ("-r", "--diffperiod"):
            diffperdays = arg
        elif opt in ("-c", "--cleanrepover"):
            cleanrepovdays = arg
    return diffperdays, cleanrepovdays


# Функция группировки данных pandas по проектам и дате
def dfGroupBy(fle, num):
    locals()["rws " + str(num)] = pd.read_csv(fle, index_col=None, header=0, encoding="cp1251", sep=',')
    locals()["dfr" + str(num)] = locals()["rws " + str(num)].groupby(['Date', 'Project']).agg(
        {'Vmdk_Used_Space(GB)': 'sum', 'MEM(GB)': 'sum', 'vCPU': 'sum'})
    return locals()["dfr" + str(num)]


# Функция отправки почты
def sendMail():
    # Отправка e-mail со вложениями
    report_date = datetime.now()
    recipients = ['absolute.creator@gmail.com']  # 
    msg = MIMEMultipart()
    msg['From'] = "vcreport@sample.ru"
    msg['To'] = ", ".join(recipients)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = f"Выгрузка ВМ из vmware vСentre за {report_date.strftime(frmt)}"

    if prjctCntSame:
        body = "Файлы во вложении.\nКонтакты: absolute.creator@gmail.com\nJob: https://rundeck.absolutecreator.ru/project/VMWare_report/jobs"
        for fi in [html_file, xlsx_file]:
            with open(fi, "rb") as fil:
                part = MIMEApplication(
                    fil.read(),
                    Name=basename(fi)
                )
            # After the file is closed
            part['Content-Disposition'] = 'attachment; filename="%s"' % basename(fi)
            msg.attach(part)

    else:
        body = "В vcentre добавлен или удалён проект. Выполните скрипт вручную или дождитесь следующего запуска по расписанию."
    msg.attach(MIMEText(body))
    if body2:
        msg.attach(MIMEText(body2))

    with smtplib.SMTP('192.168.168.168') as s:
        s.send_message(msg)
    if os.path.exists(lastCsv):
        shutil.rmtree(lastCsv)


# Запуск программы
if __name__ == "__main__":

    with open(csv_file, 'w', newline='', encoding="cp1251", errors="ignore") as file:
        writer = csv.writer(file, delimiter='\t', quoting=csv.QUOTE_NONE)
        writer.writerow(csvHeader)
    file.close()

    vCentre = "vc.vmcod.local"
    vCtr, body2 = main(vCentre)
    vCentre = "vcsa.vmcod.local"
    vCtr, body2 = main(vCentre)
    vCentre = "cod.vmcod.local"
    vCtr, body2 = main(vCentre)
    csvData = pd.read_csv(csv_file, delimiter=',', encoding="cp1251")
    csvData = csvData.drop_duplicates(keep='first')
    header, preTable, styles, css, sort_script, footer = htmlBuild()
    table = csvData.style.set_table_styles(styles).set_table_attributes('class="table_sort"').apply(highlight, subset=[
        'PowerState'], axis=1).hide_index().hide_columns(subset='Date').format({'Guest_HDD(GB)': '{:.2f}',
                                                                                'Vmdk_Used_Space(GB)': '{:.2f}',
                                                                                'Guest_UsedSpace(GB)': '{:.2f}'}).render(
        escape=False)
    # .round(2)
    # HDD Pie (текущие значения по vmdk)
    # Создаем палитру цветов
    grouped = csvData.groupby(['Project']).sum()
    colors = sns.set_palette("hls", n_colors=len(csvData['Project'].unique()))

    # Строим круговую диаграмму
    fig, ax = plt.subplots()
    wedges, texts, autotexts = ax.pie(grouped['Vmdk_Used_Space(GB)'].round(2),
                                      labels=grouped.index,
                                      labeldistance=1.2,
                                      autopct='%1.1f%%',
                                      startangle=90,
                                      radius=1.2,
                                      colors=colors)

    # Устанавливаем шрифт и размер шрифта для лейблов и процентов
    plt.setp(texts, fontweight='normal', fontsize=12)
    plt.setp(autotexts, fontweight='normal', fontsize=10)
    adjust_text(texts, ha='left')
    adjust_text(autotexts, ha='left')

    # Добавляем легенду
    ax.legend(wedges, grouped.index,
              title='Projects',
              loc='center left',
              bbox_to_anchor=(1, 0, 0.5, 1))

    # Устанавливаем заголовок диаграммы
    ax.set_title('Vmdk_Used_Space(GB)', fontsize=16, pad=20)

    # Переводим диаграмму в Base64 и сохраняем в переменную для использования в html
    HDDPie = BytesIO()
    plt.savefig(HDDPie, format='png', dpi=100)
    encoded = base64.b64encode(HDDPie.getvalue()).decode('utf-8')
    HDDPiePng = ('<div class="center img"><img src=\'data:image/png;base64,{}\' alt=\'HDDPie\'/></div>'.format(encoded))
    diffperiod, cleanrepover = params(sys.argv[1:])

    # Чтение отчетов по ресурсам (прошлые даты + текущий) и бекап более ранних отчетов, чем DiffPeriod
    csvs = glob.glob(os.path.join(CsvDir, "*.csv"))
    webreps = glob.glob(os.path.join(dr, "*.html"))
    perpas = diffperiod + 'dcsvs'
    csvsper = CsvDir + perpas
    if os.path.exists(csvsper):
        shutil.rmtree(csvsper)
    os.makedirs(csvsper)
    df = pd.DataFrame()
    cFls = 0
    prevCnt = 0
    curCnt = 0
    for i in webreps:
        wrept = datetime.strptime(str(re.search("(\d{2}.\d{2}.\d{4})", i)[0]), frmt)
        if wrept < datetime.strptime(timeStr, frmt) - timedelta(days=int(cleanrepover)):
            os.remove(i)
    for i in csvs:
        cFls += 1
        rept = datetime.strptime(str(re.search("(\d{2}.\d{2}.\d{4})", i)[0]), frmt)
        if rept < datetime.strptime(timeStr, frmt) - timedelta(days=int(cleanrepover)):
            os.remove(i)
        if rept == (datetime.strptime(timeStr, frmt) - timedelta(days=int(1))):
            prevCsv = i
            df = dfGroupBy(i, cFls)
            df = df.reset_index()
            prevCnt = len(df.index)
            # print(df)
            df.iloc[0:0]
        if rept == (datetime.strptime(timeStr, frmt)):
            curCsv = i
            df = dfGroupBy(i, cFls)
            df = df.reset_index()
            curCnt = len(df.index)
            df.iloc[0:0]
        if rept < datetime.strptime(timeStr, frmt) - timedelta(days=int(diffperiod)):
            continue
        else:
            shutil.copy2(i, csvsper, follow_symlinks=True)
    df = pd.DataFrame()
    list_of_files = []
    for file in glob.glob(os.path.join(csvsper, "*.csv")):
        list_of_files.append((getmtime(file), file))
    csvsprList = [file for _, file in sorted(list_of_files)]
    for filename in csvsprList:
        rows = pd.read_csv(filename, index_col=None, header=0, encoding="cp1251", sep=',')
        # print('Чтение csv')
        df = df.append(rows)
        # print(df)
    prjctCntSame = True
    if cFls > 1:
        if prevCnt != curCnt:
            prjctCntSame = False
            shutil.move(CsvDir, bakDir + timeStr)
            os.makedirs(CsvDir)
            print("Added new project. Start new week. Reports will be backed up. Script being restarted.")
            headers = {'Content-Type': 'text/xml'}
            params = {'authtoken': 'AUTHTOKEN'}
            response = requests.post('http://127.0.0.1:4440/api/40/job/job-id/run', params=params, headers=headers)
            print(response)
            sendMail()
            quit()
    df['Date'] = pd.to_datetime(df['Date'].astype(str), format=frmt)
    df = df.groupby(['Date', 'Project']).agg({'Vmdk_Used_Space(GB)': 'sum', 'MEM(GB)': 'sum', 'vCPU': 'sum'})
    df = df.reset_index()
    df['Date'] = df['Date'].astype(str)
    df = df.set_index(['Date', 'Project'])
    maxPeriod = (sum(df.value_counts(subset=['Date'])))
    datesCount = 0
    for i in df.value_counts(subset=['Project']):
        datesCount = i
        if i:
            break
    minPeriod = maxPeriod / datesCount
    dfMinDiff = df.diff(periods=int(minPeriod))
    dfMaxDiff = df.diff(periods=int(maxPeriod - minPeriod))
    dfMaxDiff.dropna(subset=['Vmdk_Used_Space(GB)', 'MEM(GB)', 'vCPU'], inplace=True)
    shutil.rmtree(csvsper)

    # Создание изображений графиков base64 и добавление в html
    UsedSpaceVmdkHist = UsedSpaceVmdkHistBar = df.unstack()['Vmdk_Used_Space(GB)']
    UsedMemHist = UsedMemHistBar = df.unstack()['MEM(GB)']
    UsedCpuHist = UsedCpuHistBar = df.unstack()['vCPU']
    pltType = ''
    title = 'Total historically Vmdk_Used_Space(GB) (last reports)'
    encoded = plotting(UsedSpaceVmdkHist, title, pltType)

    UsedSpaceVmdkHistChartPng = (
        '<div class="image-container"><img src=\'data:image/png;base64,{}\' alt=\'UsedSpaceVmdkHistChart\'/>'.format(
            encoded))

    pltType = 'Bar'
    title = 'Total historical Vmdk_Used_Space(GB) (last reports)'
    encoded = plotting(UsedSpaceVmdkHistBar, title, pltType)

    UsedSpaceVmdkHistBarPng = (
        '<div class="image-container"><img src=\'data:image/png;base64,{}\' alt=\'UsedSpaceVmdkHistBar\'/>'.format(
            encoded))

    pltType = ''
    title = 'Total historical MEM(GB) usage (last reports)'
    encoded = plotting(UsedMemHist, title, pltType)

    UsedMemHistChartPng = (
        '<img src=\'data:image/png;base64,{}\' alt=\'UsedMemHistChart\'/>'.format(
            encoded))

    pltType = 'Bar'
    title = 'Total historical MEM(GB) usage (last reports)'
    encoded = plotting(UsedMemHistBar, title, pltType)

    UsedMemHistBarPng = (
        '<img src=\'data:image/png;base64,{}\' alt=\'UsedMemHistBar\'/>'.format(
            encoded))

    pltType = ''
    title = 'Total historical vCPU usage (last reports)'
    encoded = plotting(UsedCpuHist, title, pltType)

    UsedCpuHistChartPng = (
        '<img src=\'data:image/png;base64,{}\' alt=\'UsedCpuHistChart\'/></div>'.format(
            encoded))

    pltType = 'Bar'
    title = 'Total historical vCPU usage (last reports)'
    encoded = plotting(UsedCpuHistBar, title, pltType)

    UsedCpuHistBarPng = (
        '<img src=\'data:image/png;base64,{}\' alt=\'UsedCpuHistBar\'/></div>'.format(
            encoded))

    htmlTable = (css + preTable + table + footer + sort_script).encode('utf8')

    with open(html_file, 'w') as f:
        f.write(header)
        f.write(UsedSpaceVmdkHistBarPng)
        f.write(UsedMemHistBarPng)
        f.write(UsedCpuHistBarPng)
        f.write(UsedSpaceVmdkHistChartPng)
        f.write(UsedMemHistChartPng)
        f.write(UsedCpuHistChartPng)
        f.write(HDDPiePng)

    with open(html_file, 'ab') as f:
        f.write(htmlTable)

    # Создание итогового xlsx
    with pd.ExcelWriter(xlsx_file) as writer:
        csvData.to_excel(writer, sheet_name='Выгрузка ВМ')
        df.to_excel(writer, sheet_name='Отчет по ресурсам')
        dfMinDiff.to_excel(writer, sheet_name='1Rep diff')
        dfMaxDiff.to_excel(writer, sheet_name=diffperiod + 'Rep diff')

    sendMail()
