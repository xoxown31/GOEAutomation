import pandas as pd
import pyautogui as pg
import pyperclip
import time
from tkinter import filedialog, messagebox
import time
import json
import subprocess

def start_setting():
    global menuLoc
    settingData = {}
    try:
        with open('setting.json') as f:
            settingData = json.load(f)
    except:
        with open('setting.json', 'w', encoding='utf-8') as f:
            settingData['GOEMessengerPath'] = filedialog.askopenfilename(
                    initialdir="C:/Program Files (x86)/AtMessenger7", 
                    title="GOE 메신저 선택",
                    filetypes=[("exe file", ".exe")]
                )
            json.dump(settingData, f, indent="\t")
            
    subprocess.run(settingData['GOEMessengerPath'])
    
    time.sleep(1)
    menu_offLoc = pg.locateOnScreen('image/menu_org_off.png')
    if menu_offLoc is not None:
        pg.click(menu_offLoc)
    
    menuLoc = pg.locateOnScreen('image/menu_org.png')
    if menuLoc is None:
        messagebox.showerror(title="세팅 오류", message="'조직도' 메뉴를 선택해주세요.")
        exit()

def validate(l):
    
    tab_start_loc = pg.locateCenterOnScreen('image/tab_start.png')
    if tab_start_loc is None:
        messagebox.showerror(title="사용자 추가 오류", message="GOE 메신저 창 (좌상단)을 찾을 수 없습니다.")
        exit()
    
    tab_end_loc = pg.locateCenterOnScreen('image/tab_end.png')
    if tab_end_loc is None:
        messagebox.showerror(title="사용자 추가 오류", message="GOE 메신저 창 (우하단)을 찾을 수 없습니다.")
        exit()
        
    region = (tab_start_loc.x, tab_start_loc.y, tab_end_loc.x - tab_start_loc.x, tab_end_loc.y - tab_start_loc.y)
    
    school_loc = pg.locateOnScreen('image/school_org.png')
    if school_loc is None:
        messagebox.showerror(title="사용자 추가 오류", message="'소속' 입력 창을 찾을 수 없습니다.")
        exit()

    name_loc = pg.locateOnScreen('image/name_org.png')
    if name_loc is None:
        messagebox.showerror(title="사용자 추가 오류", message="'검색어 입력' 입력 창을 찾을 수 없습니다.")
        exit()

    search_button_loc = pg.locateOnScreen('image/search_button_org.png')
    if search_button_loc is None:
        messagebox.showerror(title="사용자 추가 오류", message="'검색' 버튼을 찾을 수 없습니다.")
        exit()

    pg.moveTo(school_loc)
    pg.click()
    pg.typewrite('.')
    
    pg.moveTo(name_loc)
    pg.click()
    pg.typewrite('.')
    
    pg.moveRel(0, 100)
    pg.click()
    
    delete_school_loc, delete_name_loc = pg.locateAllOnScreen('image/delete.png')
    if delete_school_loc is None:
        messagebox.showerror(title="사용자 추가 오류", message="'삭제' 버튼을 찾을 수 없습니다.")
        exit()
        
    reject_list = []
        
    for school, name in l:
        pg.click(delete_school_loc)
        pyperclip.copy(school)
        pg.hotkey('ctrl', 'v')

        pg.click(delete_name_loc)
        pyperclip.copy(name)
        pg.hotkey('ctrl', 'v')
        
        pg.click(search_button_loc)
        time.sleep(1)
        
        not_user_loc = pg.locateOnScreen('image/not_user.png', region=region)
        if not_user_loc is not None:
            pg.click(not_user_loc)
            reject_list.append((school, name))
            
    print("[ 검색 이상자 목록 ]")
    print('소속\t이름')
    for school, name in reject_list:
        print(school, name, sep='\t')
    print("총 인원 :", len(l))
    print("승인 인원 :", len(l) - len(reject_list))
    print("이상 인원 :", len(reject_list))
    print("이상자 비율 : ", (len(reject_list) / len(l)) * 100, "%", sep='')

if __name__=='__main__':
    
    start_time0 = time.time()
    
    inputPath = filedialog.askopenfilename(
        initialdir=".", 
        title="엑셀 파일 선택",
        filetypes=[("Excel files", ".xlsx .xls")]
        )

    try:
        df = pd.read_excel(inputPath)
    except:
        messagebox.showerror(title="엑셀 파일 오류", message="제대로 된 엑셀 파일을 지정해주세요.")
        exit()
        
    start_setting()
    
    start_time1 = time.time()
        
    l = [x[-2:] for x in df.values.tolist()]
    unit = len(l)
    n = len(l[0])
    
    validate(l)
    print("[ 실행 시간 분석 ]")
    print("Total time (sec) : "+ str(time.time() - start_time0))
    print("Setting time (sec) : "+ str(start_time1 - start_time0))
    print("Automation time (sec) : "+ str(time.time() - start_time1))
    print("Unit : "+ str(unit))
    print("Unit per second (Total) : "+ str(unit /(time.time() - start_time0)))
    print("Unit per second (Automation) : "+ str(unit /(time.time() - start_time1)))
    print("Second per Unit (Total) : "+ str((time.time() - start_time0)/unit))
    print("Second per Unit (Automation) : "+ str((time.time() - start_time1)/unit))