import pandas as pd
import pyautogui as pg
import pyperclip
import time
import os
import win32api
from tkinter import filedialog, messagebox
from collections import defaultdict
from math import isnan
from pprint import pprint
import time

moveGroup = lambda i : pg.moveTo(pg.center(group0Loc) + pg.Point(0, 60 * (i+3)))
search_menu_loc = school_loc = name_loc = search_button_loc = delete_name_loc = delete_school_loc = complete_loc = None

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
    global group0Loc
    menuLoc = pg.locateOnScreen('image/menu.png')
    if menuLoc is None:
        messagebox.showerror(title="세팅 오류", message="'내 목록' 메뉴를 선택해주세요.")
        exit()
        
    pg.moveTo(menuLoc)
    pg.moveRel(60, 0)
    pg.click()
    pg.moveRel(-60, 0)
    pg.click()
    
    if pg.locateOnScreen('image/drop_button_down.png') is not None:
        messagebox.showerror(title="세팅 오류", message="그룹을 펼치지 말고 전부 닫아주세요.")
        exit()
    
    group0Loc = pg.locateOnScreen('image/group_0.png')
    
def create_group(name, depth):
    moveGroup(depth)
    pg.click(button='right')
    pg.moveRel(60, 195)
    pg.click()
    
    pyperclip.copy(name)
    pg.hotkey('ctrl', 'v')
    pg.press('enter')
    

def add_user(l, depth):
    global search_menu_loc, school_loc, name_loc, search_button_loc, delete_name_loc, delete_school_loc, complete_loc
    
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
    
    pg.moveRel(0, 100)
    pg.click()
    
    if delete_school_loc is None:
        delete_school_loc, delete_name_loc = pg.locateAllOnScreen('image/delete.png')
        if delete_school_loc is None:
            messagebox.showerror(title="사용자 추가 오류", message="'삭제' 버튼을 찾을 수 없습니다.")
            exit()
    
    
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
        
    pg.moveTo(complete_loc)
    pg.click()
    

if __name__=='__main__':
    
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
        
    l = df.values.tolist()
    n = len(l[0])
    
    d = {}

    for line in l:
        add(d, line, 0)
    
    dfs(d, 0)
    
    
'''
for station, name in l:
    
    pg.moveTo(delete_station)
    pg.click()
    pyperclip.copy(station)
    pg.hotkey('ctrl', 'v')

    pg.moveTo(delete_name)
    pg.click()
    pyperclip.copy(name)
    pg.hotkey('ctrl', 'v')
    
    pg.moveTo(search_button)
    pg.click()
    time.sleep(2)
    pg.moveRel(-40, 44)
    pg.click(button='right')
    time.sleep(0.1)
'''