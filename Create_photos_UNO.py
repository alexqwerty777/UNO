import time
import psutil

import pyautogui
from PIL import Image
import subprocess

# import os
# print(os.listdir('C:\\Users\\Admin\\Desktop\\UNO\\UNO-photo\\'))
path = 'C:\\Users\\Admin\\Desktop\\UNO\\UNO-photo\\'


def create_photos(img_res: Image):
    photos = ['Леха_2.jpg', 'Диман_2.jpg', 'Рустам_2.jpg', 'ОляТян_2.jpg', 'ОляСан_2.jpg',
              'Рома_2.jpg', 'Наташа_2.jpg', 'Саня_2.jpg', 'Некит_2.jpg', 'Лена_2.jpg']
    # photos = [i for i in range(10)]
    # 'Results.png'
    # img_res = Image.open(path + 'scr.png')
    x, y = img_res.size
    print(x, y)
    split_lines = []
    # find black color, it's split lines
    for i in range(10, x):
        color = img_res.getpixel((i, 3))
        print(color)
        if color == (0, 0, 0, 255) or color == (0, 0, 0):
            split_lines.append(i)
    print(split_lines)
    split_lines = split_lines[:10]
    split_lines.append(x)
    # print(split_lines)
    left_side = img_res.crop((0, 0, split_lines[0], y))
    # left_side.show()
    for i, photo in enumerate(photos):
        img = Image.open(path + '//photos//' + photo)
        right_side = img_res.crop((split_lines[i], 0, split_lines[i + 1], y))
        # right_side.show()
        left_side = left_side.resize((200, 512))
        right_side = right_side.resize((200, 512))
        lift_size_x, _ = left_side.size
        img.paste(left_side, (0, 0))
        img.paste(right_side, (lift_size_x, 0))
        # img.show()
        img.save(path + photo)


def get_screenshot(my_path: str) -> Image:
    my_screen = pyautogui.screenshot(my_path + 'Results.png', region=(22, 262, 1370, 346))
    return my_screen


def open_excel(xl_path: str):
    subprocess.Popen(xl_path, shell=True)


def close_exel():
    for process in (process for process in psutil.process_iter() if process.name() == "EXCEL.EXE"):
        process.kill()


def open_path():
    subprocess.Popen(r'explorer /select,"C:\Users\Admin\Desktop\UNO\UNO-photo\OLD"')


def main():
    path_xl = 'C:\\Users\\Admin\\Desktop\\UNO\\Uno_.xlsx'
    open_excel(path_xl)
    time.sleep(2)
    my_screen = get_screenshot(path)
    # close_exel()
    create_photos(my_screen)
    open_path()
    # screen.show()
    # create_photos()


if __name__ == '__main__':
    main()
