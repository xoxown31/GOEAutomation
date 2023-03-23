import pandas as pd
import pyautogui as pg
import pyperclip
import time
from tkinter import filedialog, messagebox
import json
import time
import subprocess

moveGroup = lambda i : pg.moveTo(pg.center(group0Loc) + pg.Point(0, 60 * (i+3)))
search_menu_loc = school_loc = name_loc = search_button_loc = delete_name_loc = delete_school_loc = complete_loc = tab_start_loc = tab_end_loc = None
reject_list = []

def add(d, l, x):
    if str(l[x+1]) == 'nan' or n-x-1 < 3:
        if l[x] not in d:
            d[l[x]] = [(l[-2], l[-1])]
        else:
            d[l[x]].append((l[-2], l[-1]))
        return d
    
    if l[x] in d:
        d[l[x]] = add(d[l[x]], l, x+1)
    else:
        d[l[x]] = add({}, l, x+1)
        
    return d

def dfs(d, depth):
    if type(d) == list:
        add_user(d, depth)
        return
    for i, key in enumerate(sorted(d.keys(), reverse=True)):
        create_group(key, depth)
        if i == 0:
            moveGroup(depth)
            pg.click()
        dfs(d[key], depth+1)

def start_setting():
    global group0Loc, menuLoc
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
    menu_offLoc = pg.locateOnScreen('image/menu_off.png')
    if menu_offLoc is not None:
        pg.click(menu_offLoc)
    
    menuLoc = pg.locateOnScreen('image/menu.png')
    if menuLoc is None:
        messagebox.showerror(title="세팅 오류", message="'내 목록' 메뉴를 선택해주세요.")
        exit()
        
    pg.moveTo(menuLoc)
    pg.moveRel(60, 0)
    pg.click()
    pg.moveRel(-60, 0)
    pg.click()
    time.sleep(1)
    
    downLocs = list(pg.locateAllOnScreen('image/drop_button_down.png'))
    while downLocs:
        for e in downLocs[::-1]:
            pg.click(e)
        downLocs = list(pg.locateAllOnScreen('image/drop_button_down.png'))
        
    pg.moveTo(menuLoc)
    pg.moveRel(60, 0)
    pg.click()
    pg.moveRel(-60, 0)
    pg.click()
    time.sleep(1)
    
    group0Loc = pg.locateOnScreen('image/group_0.png')
    
    moveGroup(-3)
    pg.click(button='right')
    pg.moveRel(50, 30)
    pg.click()
    
    pyperclip.copy("0 " + str(int(time.time())))
    pg.hotkey('ctrl', 'v')
    pg.press('enter')
    
    moveGroup(-3)
    pg.click(clicks=2, interval=1)
    pg.moveTo(menuLoc)
    pg.moveRel(60, 0)
    pg.click()
    pg.moveRel(-60, 0)
    pg.click()
    
    
def create_group(name, depth):
    moveGroup(depth)
    pg.click(button='right')
    pg.moveRel(60, 195)
    pg.click()
    
    pyperclip.copy(name)
    pg.hotkey('ctrl', 'v')
    pg.press('enter')
    

def add_user(l, depth):
    global search_menu_loc, school_loc, name_loc, search_button_loc, delete_name_loc, delete_school_loc, complete_loc, tab_start_loc, tab_end_loc, reject_list, region0, region1
    
    moveGroup(depth)
    pg.click(button='right')
    pg.moveRel(60, 220)
    pg.click()
    
    time.sleep(1)
    
    if search_menu_loc is None:
        search_menu_loc = pg.locateOnScreen('image/search_menu.png')
        if search_menu_loc is None:
            messagebox.showerror(title="사용자 추가 오류", message="'검색' 메뉴를 찾을 수 없습니다.")
            exit()
    
    if complete_loc is None:
        complete_loc = pg.locateOnScreen('image/complete.png')
        if complete_loc is None:
            messagebox.showerror(title="사용자 추가 오류", message="'검색' 버튼을 찾을 수 없습니다.")
            exit()
    
    pg.moveTo(search_menu_loc)
    pg.click()
    
    if school_loc is None:
        school_loc = pg.locateOnScreen('image/school.png')
        if school_loc is None:
            messagebox.showerror(title="사용자 추가 오류", message="'소속' 입력 창을 찾을 수 없습니다.")
            exit()
    
    if name_loc is None:
        name_loc = pg.locateOnScreen('image/name.png')
        if name_loc is None:
            messagebox.showerror(title="사용자 추가 오류", message="'검색어 입력' 입력 창을 찾을 수 없습니다.")
            exit()
    
    if search_button_loc is None:
        search_button_loc = pg.locateOnScreen('image/search_button.png')
        if search_button_loc is None:
            messagebox.showerror(title="사용자 추가 오류", message="'검색' 버튼을 찾을 수 없습니다.")
            exit()
    
    pg.moveTo(school_loc)
    pg.click()
    pg.typewrite('.')
    
    pg.moveTo(name_loc)
    pg.click()
    pg.typewrite('.')
    
    if tab_start_loc is None:
        tab_start_loc = pg.locateCenterOnScreen('image/add_tab_start.png')
        if tab_start_loc is None:
            messagebox.showerror(title="사용자 추가 오류", message="GOE 메신저 창 (좌상단)을 찾을 수 없습니다.")
            exit()

        tab_end_loc = pg.locateCenterOnScreen('image/add_tab_end.png')
        if tab_end_loc is None:
            messagebox.showerror(title="사용자 추가 오류", message="GOE 메신저 창 (우하단)을 찾을 수 없습니다.")
            exit()
            
        region1 = region0 = (tab_start_loc.x, tab_start_loc.y, tab_end_loc.x - tab_start_loc.x, tab_end_loc.y - tab_start_loc.y)
    
        pg.click(search_button_loc)
        
        not_user_loc = pg.locateOnScreen('image/not_user.png', region=region0)
        while not_user_loc is None:
            not_user_loc = pg.locateOnScreen('image/not_user.png', region=region0)
        
        region0 = tuple(not_user_loc)
        pg.click(not_user_loc)
        
    pg.moveRel(0, 100)
    pg.click()
    
    if delete_school_loc is None:
        delete_school_loc, delete_name_loc = pg.locateAllOnScreen('image/delete.png')
        if delete_school_loc is None:
            messagebox.showerror(title="사용자 추가 오류", message="'삭제' 버튼을 찾을 수 없습니다.")
            exit()
    
    for school, name in l:
        
        pg.click(delete_school_loc)
        pyperclip.copy(school)
        pg.hotkey('ctrl', 'v')
        
        pg.click(delete_name_loc)
        pyperclip.copy(name)
        pg.hotkey('ctrl', 'v')
        
        pg.click(search_button_loc)
        
        not_user_loc = pg.locateOnScreen('image/not_user.png', region=region0)
        user_loc = pg.locateOnScreen('image/user.png', region=region1)
        
        while not_user_loc is None and user_loc is None:
            not_user_loc = pg.locateOnScreen('image/not_user.png', region=region0)
            user_loc = pg.locateOnScreen('image/user.png', region=region1)
            
        if not_user_loc is not None:
            region0 = tuple(not_user_loc)
            pg.click(not_user_loc)
            reject_list.append((school, name))
        else:
            region1 = tuple(user_loc)
            pg.moveRel(-40, 44)
            pg.click(button='right')
            
        
    pg.moveTo(complete_loc)
    pg.click()
    
def validate(l):
    pg.click()
    
    for school, name in l:
        pg.moveTo(delete_school_loc)
        pg.click()
        pyperclip.copy(school)
        pg.hotkey('ctrl', 'v')

        pg.moveTo(delete_name_loc)
        pg.click()
        pyperclip.copy(name)
        pg.hotkey('ctrl', 'v')
        
        pg.moveTo(search_button_loc)
        pg.click()
        time.sleep(2)
        pg.moveRel(-40, 44)
        pg.click(button='right')
        time.sleep(0.1)

if __name__=='__main__':
    
    start_time0 = time.time()
    
    inputPath = filedialog.askopenfilename(
        initialdir=".", 
        title="엑셀 파일 선택",
        filetypes=[("Excel files", ".xlsx .xls")]
        )

    try:
        df = pd.read_excel(inputPath, sheet_name = 'input')
    except:
        messagebox.showerror(title="엑셀 파일 오류", message="제대로 된 엑셀 파일을 지정해주세요.")
        exit()
        
    start_setting()
    
    l = df.values.tolist()
    unit = len(l)
    n = len(l[0])
    
    d = {}

    for line in l:
        add(d, line, 0)
    
    start_time1 = time.time()
    
    dfs(d, 0)
    
    print("[ 검색 이상자 목록 ]")
    print('소속\t이름')
    for school, name in reject_list:
        print(school, name, sep='\t')
    print("총 인원 :", len(l))
    print("승인 인원 :", len(l) - len(reject_list))
    print("이상 인원 :", len(reject_list))
    print("이상자 비율 : ", (len(reject_list) / len(l)) * 100, "%", sep='')
    print()
    print("[ 실행 시간 분석 ]")
    print("Total time (sec) : "+ str(time.time() - start_time0))
    print("Setting time (sec) : "+ str(start_time1 - start_time0))
    print("Automation time (sec) : "+ str(time.time() - start_time1))
    print("Unit : "+ str(unit))
    print("Unit per second (Total) : "+ str(unit /(time.time() - start_time0)))
    print("Unit per second (Automation) : "+ str(unit /(time.time() - start_time1)))
    print("Second per Unit (Total) : "+ str((time.time() - start_time0)/unit))
    print("Second per Unit (Automation) : "+ str((time.time() - start_time1)/unit))