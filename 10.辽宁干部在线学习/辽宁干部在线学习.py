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


def find_xue():
    xue_list = pyau.locateAllOnScreen('xuexi.png')
    xue_wan_list = pyau.locateAllOnScreen('wancheng100.png')
    if xue_wan_list:
        for i in xue_list:
            for j in xue_wan_list:
                x_zuobiao = i.left - j.left
                y_zoubiao = i.top - j.top
                print(x_zuobiao,y_zoubiao)
                if 0 < x_zuobiao <100 and 0< y_zoubiao < 30:
                    return (i)
            print(i)
    else:
        if xue_list:
            dianji_click = list(xue_list)[0]
            print(dianji_click)
            return ((dianji_click.left, dianji_click.top))
        else:
            print('all study')

    return (r'未找到')

@timer
def main():
    zhaodao = find_xue()
    if  zhaodao == r'未找到':
        pyau.press('pagedown', presses=1, interval=0.5)
        print('pagedown')
    else:
        x1 = zhaodao.left+30
        y1 = zhaodao.top+10
        pyau.click(x=x1, y=y1, clicks=1, interval=0.5, button=PRIMARY, duration=0.0)
        print('点击学习')

    print('ok')
    

if __name__ == "__main__":
    main()
    print('all ok')
