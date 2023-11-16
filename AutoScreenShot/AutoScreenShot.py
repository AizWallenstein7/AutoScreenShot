import os
import cv2
import time
import ctypes
from ctypes import windll, wintypes
import numpy as np
from PIL import ImageGrab
from datetime import datetime as dt


def screenshot_position():
    try:
        windll.shcore.SetProcessDpiAwareness(True)
    except:
        pass

    def position_get():
        point = wintypes.POINT()
        windll.user32.GetCursorPos(ctypes.byref(point))
        return point.x, point.y

    input('取得したい箇所の"左上"にカーソルを当てEnterキー押してください')
    x1, y1 = position_get()
    print(f'X:{str(x1)}, Y:{str(y1)}')
    input('取得したい箇所の"右下"にカーソルを当てEnterキー押してください')
    x2, y2 = position_get()
    print(f'X:{str(x2)}, Y:{str(y2)}')
    with open('screenshot_position.txt', 'w') as f:
        f.write(f'{str(x1)}, {str(y1)}, {str(x2)}, {str(y2)}')


def screenshot_loop(stop_time, dir_path, x1, y1, x2, y2):
    try:
        while True:
            time.sleep(stop_time)
            img = ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)
            img_ima = np.array(img)
            if 'img_mae' in locals():
                backSub = cv2.createBackgroundSubtractorMOG2()
                fgmask = backSub.apply(img_mae)
                fgmask = backSub.apply(img_ima)
                if np.count_nonzero(fgmask) / fgmask.size * 100 < 0.5:
                    continue
            img.save(f'{dir_path}/{dt.now().strftime("%y%m%d%H%M%S")}.png')
            img_mae = img_ima
    except KeyboardInterrupt:
        print('\n■■キャプチャ終了■■')


def main():
    if not os.path.exists('screenshot_position.txt'):
        print('スクリーンショット位置を設定してください')
        screenshot_position()
    else:
        screenshot_set = input("スクリーンショットする位置を設定しますか？[y/any]")
        if screenshot_set == "y" or screenshot_set == "Y":
            screenshot_position()
    x1, y1, x2, y2 = np.loadtxt('screenshot_position.txt', delimiter=',')
    stop_time = input('撮影間隔(秒)を入力して、ENTERで開始(デフォルト値：2秒、最低値：1秒)：')
    stop_time = 2 if not stop_time else 1 if float(stop_time) < 1 else float(stop_time)
    dir_path = dt.now().strftime('%y%m%d_%H%M')
    if not os.path.exists(dir_path):
        os.mkdir(dir_path)
    print('\nスクリーンショットを開始しました\nCtl+cで終了できます')
    screenshot_loop(stop_time, dir_path, x1, y1, x2, y2)


if __name__ == '__main__':
    main()