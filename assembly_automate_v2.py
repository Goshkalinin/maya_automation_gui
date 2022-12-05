import os
import pywinauto as pw
from time import sleep
import sys
import logging
from datetime import datetime
from PIL import ImageGrab


def log(file_name, massage, screenshot = False):

    if not os.path.exists("log"):
        os.mkdir("log")
    
    logging.basicConfig(
        format="%(asctime)s %(message)s",
        encoding='utf-8',
        level=logging.DEBUG,
        handlers=[
        logging.FileHandler("log\\" + file_name + '.log'),
        logging.StreamHandler(sys.stdout)
    ])
    logging.debug(massage)
    
    if screenshot == False:
        pass
    else:
        now = datetime.now()
        current_time = now.strftime("%H_%M_%S_")
        myscreen = ImageGrab.grab()
        myscreen.save("log\\" + current_time + file_name + '.jpg')

def get_files_list():

    path = os.getcwd()  # подхватываем адрес рабочей папки
    files = os.listdir(path + '\\animate')  # собираем список на обработку

    print(files)

    return(files)

def open_shot(file_name):

    path = os.getcwd()
    folder = file_name[-12:-3]

    file_path = path + '\\' + folder + "\\animate\\" + file_name

    pw.Application().start('explorer.exe "C:\\Program Files"')  # стартуем проводник
    app = pw.Application(backend="uia").connect(path="explorer.exe", title="Program Files")  # подключаемся к нему
    dlg = app["Program Files"]  # подключаемся к окошку
    dlg['Address: C:\Program Files'].wait("ready")  # ждём поисковую строку

    dlg.type_keys("%D")    # переключаемся на неё хоткеем
    dlg.type_keys(file_path, with_spaces=True)    # впечатываем туда файл
    dlg.type_keys("{ENTER}")
    sleep(2)

    app.kill()  # закрываем проводник

def check_maya_connected(file):
    print("Trying connect with Maya")
    try:
        app = pw.Application(backend="uia").connect(title_re=".* - Autodesk Maya 2020.4:")
        print("Connected!")
        return app
    except pw.findwindows.ElementNotFoundError:
        for i in ['.  ', '.. ', '...']: 
            sys.stdout.write('\r'"connecting" + i)
            sleep(0.5)
        pass

    try:
        app = pw.Application(backend="uia").connect(title="Maya")
        
        os.system('taskkill /im maya.exe /t /f')
        os.system('taskkill /im cmd.exe /t /f')

        path = os.getcwd()
        folder = file[-12:-3]
        file_path = path + '\\' + folder + "\\animate\\" + file
        os.remove(file_path)
        os.remove("animate\\" + file)

        # print(file_path + " is removed")
        print("animate\\" + file + " is removed")

        log(file, "_NoReferences", screenshot = False)
        
        return process()

    except pw.findwindows.ElementNotFoundError:
        return check_maya_connected(file)

def app_status(app):
    
    """if app was get any cpu for last n seconds
    lets think that app in busy"""
    
    print("Checking Maya load")

    load_points = [1]
    while sum(load_points[-30:]) != 0:
        load = app.cpu_usage()

        if load < 0.5:  # убираем случайные всплески
            load = 0
        
        load_points.append(load)
            
        for i in ['.  ', '.. ', '...']: 
            sys.stdout.write('\r'"connecting" + i + str(load))
            sleep(0.33)

    sys.stdout.write('')
    print("Maya is free now")
    pass

def get_assembly(app, file):
    print("Making Assembly")

    episode = file[0:3]
    dlg = app.top_window()
    dlg["MultyDo"].click_input()
    dlg["Auto Assemble"].click_input(); sleep(1); dlg["Auto Assemble"].click_input()
    dlg["Episodes"].click_input(); dlg["Episodes"].type_keys(episode); dlg["Episodes"].type_keys("{ENTER}")
    dlg[file].click_input(); dlg["Add"].click_input(); dlg["Assemble"].click_input()

    app_status(app)
    print("Assembly is finished")

    os.system('taskkill /im maya.exe /t /f')
    os.system('taskkill /im cmd.exe /t /f')

    path = os.getcwd()
    folder = file[-12:-3]
    file_path = path + '\\' + folder + "\\animate\\" + file
    os.remove(file_path)
    os.remove("animate\\" + file)

    print(file_path + " is removed")
    print("animate\\" + file + " is removed")
    

def process():

    try:
        files = get_files_list()  # загружаем список файлов
        
        for file in files:
        
                open_shot(file)
                app = check_maya_connected(file)
                app_status(app)
                get_assembly(app, file)


    except Exception as e:
        print(e)
        log("OOPS", e, False)

process()
