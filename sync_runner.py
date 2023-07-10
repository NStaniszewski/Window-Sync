import time
import win32con
import win32gui
import win32ui
import win32api
import pyautogui
import win32com.client
import os
import pynput
#from pynput.keyboard import Key,Listener
import sched

def on_f10(key):
    if key==pynput.keyboard.Key.f10:
        os._exit(0) 
    
def show_open_windows():
    def print_windows(window,holder):
        if win32gui.IsWindowVisible(window):
            print(window,win32gui.GetWindowText(window))
    win32gui.EnumWindows(print_windows,None)
    print()

def get_chosen_windows(names):
    out_list = []
    names=names.split(',')
    def get_specific_windows(window,holder):
        for name in names:
            if win32gui.IsWindowVisible(window) and win32gui.GetWindowText(window).lower().find(name.lower())!=-1:
                out_list.append(window)
    win32gui.EnumWindows(get_specific_windows,None)
    return set(out_list) #dont want same window twice in the list if we ping 2 keywords in its name

def boot_stuff():
    print("Open Windows:")
    show_open_windows()
    window_list=[]
    while window_list==[]:
        name = input("Input Window Name:")
        window_list = get_chosen_windows(name.strip(' '))
    print("Now Running. Press F9 to Start/Stop Syncing. Press F10 to turn off.\nSync: Off")
    return window_list

def auto_input(windows,script):
    vk_map={'W':0x57,'A':0x41,'S':0x53,'D':0x44,'P':0x55,'LSHFT':0xA0,'LCTRL': 0xA2,'SPACE':0x20,"HOME":0x24,'END':0x23} #temp idea until i get string --> hex conversion working
    mouse_map={'LEFT_CLICK':'L','RIGHT_CLICK':'R','MID_CLICK':'M'}
    s=sched.scheduler(time.time,time.sleep)
    for button in script:
        offset_time=1
        for window in windows:
            if button[0] in vk_map:
                press_keyboard(window,s,vk_map[button[0]],float(button[1])+offset_time,float(button[2].strip()))
                offset_time+=.31
            elif button[0] in mouse_map:
                press_mouse(window,s,mouse_map[button[0]],float(button[1])+offset_time,int(button[2]),int(button[3].strip())) #no hold time for mouse button rn, only does clicks atm
                offset_time+=.31
        s.run()

def press_keyboard(window,s:sched.scheduler,key,start_time,hold_time):
    prio=2
    foreground_time=.15
    duration=start_time+hold_time

    s.enter(start_time-foreground_time,prio,win32gui.SetForegroundWindow,argument=(window,))
    s.enter(start_time,prio,win32api.SendMessage,argument=(window,win32con.WM_KEYDOWN,key,0))
    s.enter(duration-foreground_time,prio,win32gui.SetForegroundWindow,argument=(window,))
    s.enter(duration,prio,win32api.SendMessage,argument=(window,win32con.WM_KEYUP,key,0))

def press_mouse(window,s:sched.scheduler,key,start_time,winx,winy):
    prio=2
    foreground_time=.02
    if key=='L':
        s.enter(start_time-foreground_time,prio,win32gui.SetForegroundWindow,argument=(window,))

        s.enter(start_time-.01,prio,win32api.SetCursorPos,argument=((winx,winy),))
        s.enter(start_time,prio,win32api.mouse_event,argument=(win32con.MOUSEEVENTF_LEFTDOWN,winx,winy,0,0))
        s.enter(start_time+.01,prio,win32api.mouse_event,argument=(win32con.MOUSEEVENTF_LEFTUP,winx,winy,0,0))

    elif key=='R':
        s.enter(start_time-foreground_time,prio,win32gui.SetForegroundWindow,argument=(window,))

        s.enter(start_time-.01,prio,win32api.SetCursorPos,argument=((winx,winy),))
        s.enter(start_time,prio,win32api.mouse_event,argument=(win32con.MOUSEEVENTF_RIGHTDOWN,winx,winy,0,0))
        s.enter(start_time+.01,prio,win32api.mouse_event,argument=(win32con.MOUSEEVENTF_RIGHTUP,winx,winy,0,0))
    
    else:
        s.enter(start_time-foreground_time,prio,win32gui.SetForegroundWindow,argument=(window,))

        s.enter(start_time-.01,prio,win32api.SetCursorPos,argument=((winx,winy),))
        s.enter(start_time,prio,win32api.mouse_event,argument=(win32con.MOUSEEVENTF_MIDDLEDOWN,winx,winy,0,0))
        s.enter(start_time+.01,prio,win32api.mouse_event,argument=(win32con.MOUSEEVENTF_MIDDLEUP,winx,winy,0,0))
#sloppy as fuck maybe clean it up later but idc

def check_selected_windows(windows):
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('')
    for window in windows:
        win32gui.SetForegroundWindow(window)
        time.sleep(.2)

def pull_script(filename):
    out_list=[]
    fp=open(filename,'r')
    for line in fp:
        out_list.append(line.split(','))
    fp.close()
    return out_list

def scripty(windows):
    #file_name=input("Script File Name: ")
    file_name='test'
    script=pull_script(file_name+".txt")
    auto_input(windows,script)


def time_key_press(key):
    global t
    t=time.time()
    
def time_key_release(key):
    global t
    global time_held
    global sync_list
    time_held=round(time.time()-t,2)*2 #for some reason the timer would be on average 1/2 of how much real time passed, maybe just my own pc though
    if syncing:
        if key!=pynput.keyboard.Key.f9 and key!=pynput.keyboard.Key.f10:
            sync_list.append([key,time_held])
            print(sync_list)

def mouse_click(x,y,button,placeholder_bool): #the listener puts in a bool so i need the placeholder to avoid an error
    global mouse_list
    if syncing:
        foreground_window=win32gui.GetForegroundWindow()
        rect=win32gui.GetWindowRect(foreground_window)
        winx=rect[0]
        winy=rect[1]
        in_window_x=x-winx
        in_window_y=y-winy
        mouse_list.append([button,in_window_x,in_window_y])
        if len(mouse_list)==2:
            sync_list.append([mouse_list[1][0],mouse_list[1][1],mouse_list[1][2]])
            mouse_list=[]
            print(sync_list)

def input_sync(windows):
    vk_map={"'w'":0x57,"'a'":0x41,"'s'":0x53,"'d'":0x44,"'p'":0x55,'Key.shift':0xA0,'Key.ctrl_l': 0xA2,'Key.space':0x20,'Key.home':0x24,'Key.end':0x23}
    mouse_map={pynput.mouse.Button.left:'L',pynput.mouse.Button.right:'R',pynput.mouse.Button.middle:'M'}
    #yes doing this string thing is goofy, but its the only way i could get the map to work properly with how the different key types are output into the list by win32
    s=sched.scheduler(time.time,time.sleep)
    offset_time=0
    global key
    global time_held
    for window in windows:
        runtime=0
        for input in sync_list:
            if str(input[0]) in vk_map:
                press_keyboard(window,s,vk_map[str(input[0])],offset_time+runtime,input[1])
                runtime+=input[1]
                offset_time+=.31
            elif str(input[0]) in mouse_map:
                rect=win32gui.GetWindowRect(window)
                winx=rect[0]
                winy=rect[1]
                press_mouse(window,s,mouse_map[str(input[0])],offset_time+runtime+.01,input[1]+winx,input[2]+winy)
                runtime+=.1
                offset_time+=.31
    s.run()

def main():
    scripting=False #change to True if doing scripts, False for synching
    global syncing;syncing=False
    global sync_list;sync_list=[]
    global key;key=''
    global mouse_list;mouse_list=[]
    def on_f9(key):
        global syncing
        if key==pynput.keyboard.Key.f9:
            if syncing:
                print("Sync: Off")
                syncing=False
            else:
                print("Sync: On")
                syncing=True

    start_end=pynput.keyboard.Listener(on_release=on_f9)
    prgm_end=pynput.keyboard.Listener(on_release=on_f10)
    prgm_end.start()

    window_list = boot_stuff()
    check_selected_windows(window_list)
    start_end.start()

    if scripting:
        scripty(window_list)
    else:
        timer_start=pynput.keyboard.Listener(on_press=time_key_press)
        timer_end=pynput.keyboard.Listener(on_release=time_key_release)
        timer_start.start()
        timer_end.start()
        mouse_checker=pynput.mouse.Listener(on_click=mouse_click)
        mouse_checker.start()
        while True:
            if not syncing and sync_list!=[]:
                input_sync(window_list)
                sync_list=[]

if __name__ == "__main__":
    main()