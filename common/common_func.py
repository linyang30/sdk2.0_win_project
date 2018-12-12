import psutil
import logging
from PIL import ImageGrab
import os
from time import sleep
import yaml
import subprocess
import time
import xlsxwriter
import re
import win32gui
import win32con

def get_current_time():
    now = time.strftime('%Y-%m-%d-%H-%M-%S')
    return now

def mkdir(path):
    '''自动创建文件夹'''
    logging.info('mkdir')
    folder = os.path.exists(path)
    if not folder:
        os.mkdir(path)

def get_computer_cpu_info(interval):
    '''获取CPU占用率'''
    logging.info('get_computer_cpu_info')
    cpu_percent = str(psutil.cpu_percent(interval)) + '%'
    return cpu_percent

def get_computer_memory_info():
    '''获取内存占用率'''
    logging.info('get_computer_memory_info')
    phymem = psutil.virtual_memory()
    phymem_percent = str(phymem.percent) + '%'
    return phymem_percent

def close_window(app):
    logging.info('close_window')
    wndtitle = get_config_data()['base_dir'].replace('/', '\\') + app
    logging.info('wndtitle: ' + wndtitle)
    wndclass = "ConsoleWindowClass"
    wnd = win32gui.FindWindow(wndclass, wndtitle)
    win32gui.SendMessage(wnd, win32con.WM_CLOSE)
    sleep(5)


def get_screen_shot(appname):
    '''获取屏幕截图'''
    logging.info('get_screen_shot')
    im = ImageGrab.grab()
    mkdir('./screenshot')
    im.save('./screenshot/' + appname[0: appname.index('.')] + '_' + get_current_time() +'.png', 'png')

def get_config_data():
    '''获取配置文件数据'''
    logging.info('get_config_data')
    with open('./config/config.yaml', 'r') as file:
        data = yaml.load(file)
    return data

def create_bat_for_launch(app_name):
    '''根据应用的名称，生成bat文件'''
    logging.info('create_bat')
    data = get_config_data()
    mkdir('./temp')
    with open('./temp/' + 'start_app.bat' , 'w') as f:
        f.write('@echo off\n@start ' + data['base_dir'] + app_name + '\n@exit')

def gen_app():
    data = get_config_data()
    app_list = os.listdir(data['base_dir'])
    for app in app_list:
        if app.endswith('exe'):
            yield app

def launch_app():
    logging.info('launch_app')
    # base = os.path.dirname(os.path.dirname(__file__))
    base = get_config_data()['tool_base']
    if os.path.exists(base + '/temp/start_app.bat'):
        command = base + '/temp/start_app.bat'
        subprocess.Popen(command, stdout=subprocess.PIPE, shell=True)
    else:
        logging.info('启动文件不存在')
        logging.info(base + '/temp/start_app.bat')

def record_data(test_name, index, worksheet):
    data = get_config_data()
    sleep(5)
    logging.info('start record data: ' + test_name)
    start_cpu_info = get_computer_cpu_info(data['cpu_interval'])
    start_mem_info = get_computer_memory_info()
    logging.info('cpu: ' + start_cpu_info + '  mem: ' + start_mem_info)
    sleep(data['test_time'])
    finish_cpu_info = get_computer_cpu_info(data['cpu_interval'])
    finish_mem_info = get_computer_memory_info()
    logging.info('cpu: ' + finish_cpu_info + '  mem: ' + finish_mem_info)
    worksheet.write_string('A' + str(index), test_name)
    worksheet.write_string('B' + str(index), start_cpu_info)
    worksheet.write_string('C' + str(index), finish_cpu_info)
    worksheet.write_string('D' + str(index), start_mem_info)
    worksheet.write_string('E' + str(index), finish_mem_info)
    logging.info('end record data: ' + test_name)

def init_excel():
    mkdir('./test_result')
    workbook = xlsxwriter.Workbook('./test_result/test_result_' + get_current_time() + '.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.write_string('A1', 'App for test')
    worksheet.write_string('B1', 'Start CPU')
    worksheet.write_string('C1', 'End CPU')
    worksheet.write_string('D1', 'Start Memory')
    worksheet.write_string('E1', 'End Memory')
    worksheet.write_string('F1', 'Test Result')
    return workbook, worksheet

def write_content_verify_result(result, worksheet, index):
    if len(result) > 10:
        worksheet.write_string('F' + str(index), 'Pass')
    else:
        worksheet.write_string('F' + str(index), 'Fail')

def get_app_output(appname):
    sleep(5)
    data = get_config_data()
    # base = os.path.dirname(os.path.dirname(__file__))
    base = get_config_data()['tool_base']
    command = data['base_dir'] + appname + ' > ' + base + '/temp/content.txt'
    p = subprocess.Popen(command, stdout=subprocess.PIPE, shell=True)
    sleep(15)
    os.system('TASKKILL /PID %s /T /F' % p.pid)

def content_verify(app_name, worksheet, index):
    sleep(5)
    file = open('./temp/content.txt', 'r')
    out_put = file.read()
    if app_name == 'astra-tests.exe':
        worksheet.write_string('F' + str(index), '请手动测试')
    elif app_name == 'BodyReaderPoll.exe':
        result = re.findall(r'Body mask: width: 640 height: 480 center value: 0', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'ColorizedBodyViewer-SFML.exe':
        result = re.findall(r'[1-9]\d\.\d fps \(\d+\.\d ms\)', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'ColorReaderEvent.exe':
        result = re.findall(r'color frameIndex: \d+  r: [1-9]\d+    g: [1-9]\d+    b: [1-9]\d+', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'ColorReaderEventCPP.exe':
        result = re.findall(r'color frameIndex: \d+ r: [1-9]\d+ g: [1-9]\d+ b: [1-9]\d+', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'ColorReaderPoll.exe':
        result = re.findall(r'color frameIndex: \d+  r: [1-9]\d+    g: [1-9]\d+    b: [1-9]\d+', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'DebugHandViewer.exe':
        worksheet.write_string('F' + str(index), '请手动测试')

    elif app_name == 'DepthReaderEvent.exe':
        result = re.findall(r'depth frameIndex:  \d+  value:  [1-9]\d+', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'DepthReaderEventCPP.exe':
        result = re.findall(r'depth frameIndex: \d+ value: [1-9]\d+ wX: \d+\.\d+ wY: \d+\.\d+ wZ: [1-9]\d+\.\d+ dX: [1-9]\d+\.\d+ dY: [1-9]\d+\.\d+ dZ: [1-9]\d+\.\d+', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'DepthReaderPoll.exe':
        result = re.findall(r'depth frameIndex \d+ value [1-9]\d+', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'HandReader.exe':
        result = re.findall(r'index \d+ active hand count \d+', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'InfraredColorReaderEvent.exe':
        result1 = re.findall(r'color frameIndex: \d+  r: [1-9]\d+    g: [1-9]\d+    b: [1-9]\d+ ', out_put)
        result2 = re.findall(r'infrared frameIndex:  \d+  value:  [1-9]\d+', out_put)
        result = result1[:6] + result2[:6]
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'InfraredReaderEvent.exe':
        result = re.findall(r'infrared frameIndex:  \d+  value:  [1-9]\d+', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'InfraredReaderPoll.exe':
        result = re.findall(r'infrared frameIndex:  \d+  value:  [1-9]\d+', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'MaskedColorViewer-SFML.exe':
        result = re.findall(r'[1-9]\d+\.\d fps \(\d+\.\d ms\)', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'MultiSensorViewer-SFML.exe':
        result = re.findall(r'[1-9]\d+\.\d fps \(\d+\.\d ms\)', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'SimpleBodyViewer-SFML.exe':
        result = re.findall(r'FPS: [1-9]\d+\.\d \(\d+\.\d+ ms\)', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'SimpleColorViewer-SFML.exe':
        result = re.findall(r'[1-9]\d+\.\d fps \(\d+\.\d ms\)', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'SimpleDepthViewer-SFML.exe':
        result = re.findall(r'[1-9]\d+\.\d fps \(\d+\.\d ms\)', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'SimpleHandViewer-SFML.exe':
        result = re.findall(r'[1-9]\d+\.\d fps \(\d+\.\d ms\)', out_put)
        write_content_verify_result(result, worksheet, index)

    elif app_name == 'SimpleStreamViewer-SFML.exe':
        result = re.findall(r'[1-9]\d+\.\d fps \(\d+\.\d ms\)', out_put)
        write_content_verify_result(result, worksheet, index)
