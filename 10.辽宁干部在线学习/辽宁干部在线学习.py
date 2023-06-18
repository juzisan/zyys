"""
Spyder Editor

This is a temporary script file.

需要
1. 数据.xlsx
2. add.png
3. 安装  pip install pyautogui opencv-contrib-python pyperclip3 dddd-ocr
4. pip install dddd-ocr
"""
import time
import pyautogui as pyau
import pyperclip3 as pyco

def timer(func):
    """计时器."""

    def warpper():
        """计时器内部."""
        print('\033[1;32;40mstart\033[0m')
        time1 = time.time()
        func()
        seconds = time.time() - time1
        m, s = divmod(seconds, 60)
        print("\033[1;32;40mthe run time is %02d:%.6f\033[0m" % (m, s))
    return warpper



def find_in(xi, wan):
    x_zuobiao = xi.left - wan.left
    y_zoubiao = xi.top - wan.top
    return True if 0 < x_zuobiao <100 and 0< y_zoubiao < 30 else False



def find_xue():
    xue_list = list(pyau.locateAllOnScreen('xuexi.png'))
    xue_wan_list = list(pyau.locateAllOnScreen('wancheng100.png'))
    print('学个：', len(xue_list),len(xue_wan_list))
    fanhui_list = []
    if xue_wan_list:
        for i in xue_list:
            print('判断学习：', i)
            xue_status = any([ find_in(i, w) for w in xue_wan_list])
            print(xue_status)
            if xue_status:
                print(r'已学完', i)
            else:
                fanhui_list.append(i)
                print(r'未学完', i)
    else:
        if xue_list:
            fanhui_list.append(list(xue_list)[0])
        else:
            print(r'全学完了')

    print(fanhui_list)
    if fanhui_list:
        return (fanhui_list[0])
    else:
        return (r'未找到')

@timer
def main():
    kong_position = (30,150)
    pyau.click(*kong_position, clicks=1, interval=0.5, button='left', duration=0.0) # 空白
    for i in range(4):
        zhaodao = find_xue()
        if  zhaodao == r'未找到':
            pyau.press('pagedown', presses=1, interval=0.5)
            print('pagedown')
        else:
            loc =(zhaodao.left+30 , zhaodao.top+10)
            pyau.click(*loc, clicks=1, interval=0.5, button='left', duration=0.0)

            loc  = pyau.locateCenterOnScreen('dianjiguankan.png')
            pyau.click(*loc, clicks=1, interval=0.5, button='left', duration=0.0)

            loc  = pyau.locateCenterOnScreen('kaishixuexi.png')
            pyau.click(*loc, clicks=1, interval=0.5, button='left', duration=0.0)

            pyau.click(*kong_position, clicks=1, interval=0.5, button='left', duration=0.0)
            pyau.press('pagedown', presses=1, interval=0.5)

            loc  = pyau.locateCenterOnScreen('biji.png')
            pyau.move(*loc, duration=0.5)
            pyau.moveRel(0, -200, duration=0.5)
            time.sleep(3)
            pyau.hotkey('ctrl','w')
            time.sleep(1)
            pyau.hotkey('f5')
            time.sleep(1)
            pyau.hotkey('enter')


            print('点击学习')

    print('ok')
    

if __name__ == "__main__":
    main()
    print('all ok')
