import datetime
import math
import win32api
import win32con
import time
import ctypes
import webbrowser
from pykeyboard import PyKeyboard
from pymouse import PyMouse
import cv2
import numpy as np
import aircv as ac
import time
from PIL import ImageGrab


# 单击
def click(x, y):
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
    time.sleep(1)


# 双击
def double_click(x, y):
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
    # https://blog.csdn.net/zhanglidn013/article/details/35988381
    # https://docs.microsoft.com/zh-cn/windows/desktop/inputdev/virtual-key-codes
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
    time.sleep(0.2)


# 按下一个按钮
def press_one_key(key1):
    ctypes.windll.user32.keybd_event(key1, 0, 0, 0)
    # https://blog.csdn.net/zhanglidn013/article/details/35988381
    # https://docs.microsoft.com/zh-cn/windows/desktop/inputdev/virtual-key-codes
    ctypes.windll.user32.keybd_event(key1, 0, win32con.KEYEVENTF_KEYUP, 0)


# 关闭浏览器
def ctrl_w():
    time.sleep(2)
    win32api.keybd_event(17, 0, 0, 0)
    win32api.keybd_event(87, 0, 0, 0)
    win32api.keybd_event(87, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    # print('ctrl_w')
    time.sleep(1)


# 打开浏览器
# webbrowser.open(url, new=0, autoraise=True)
# 在系统的默认浏览器中访问url地址，
# 如果new = 0, url会在同一个 浏览器窗口中打开；
# 如果new = 1，新的浏览器窗口会被打开;
# 如果new = 2 新的浏览器tab会被打开
def web_open(url):
    webbrowser.open(url, new=0)
    time.sleep(4)


# 模拟键盘输入字符串
def string_input(my_string):
    k.type_string(my_string)
    time.sleep(1)


# 验证码找到x
def get_find_x():
    screen_grab()
    canny()
    find_x = 0
    try:
        if matchImg('yzm_screencap.png', '1.png') is not None:
            # print("1！")
            find_x = str(matchImg('yzm_screencap.png', '1.png')['result'][0])
        elif matchImg('yzm_screencap.png', '2.png') is not None:
            # print("2！")
            find_x = str(matchImg('yzm_screencap.png', '2.png')['result'][0])
        elif matchImg('yzm_screencap.png', '3.png') is not None:
            # print("3！")
            find_x = str(matchImg('yzm_screencap.png', '3.png')['result'][0])
        elif matchImg('yzm_screencap.png', '4.png') is not None:
            # print("4！")
            find_x = str(matchImg('yzm_screencap.png', '4.png')['result'][0])
        elif matchImg('yzm_screencap.png', '5.png') is not None:
            # print("5！")
            find_x = str(matchImg('yzm_screencap.png', '5.png')['result'][0])
        elif matchImg('yzm_screencap.png', '6.png') is not None:
            # print("6！")
            find_x = str(matchImg('yzm_screencap.png', '6.png')['result'][0])
        elif matchImg('yzm_screencap.png', '7.png') is not None:
            # print("7！")
            find_x = str(matchImg('yzm_screencap.png', '7.png')['result'][0])
        else:
            find_x = 0
            # print('没有匹配,find_x is :' + str(find_x))
    except Exception as e:
        print(e)
        print("这里有个异常")
    return find_x


# 对比两张图，找到坐标。
def matchImg(imgsrc, imgobj):
    # imgsrc=原始图像，imgobj=待查找的图片
    imsrc = ac.imread(imgsrc)
    imobj = ac.imread(imgobj)
    match_result = ac.find_template(imsrc, imobj, 0.5)
    # 0.9、confidence是精度，越小对比的精度就越低 {'confidence': 0.5435812473297119,
    # 'rectangle': ((394, 384), (394, 416), (450, 384), (450, 416)), 'result': (422.0, 400.alipay_leave0)}
    if match_result is not None:
        match_result['shape'] = (imsrc.shape[1], imsrc.shape[0])  # 0为高，1为宽
    return match_result


def screen_grab():
    # 参数说明
    # 第一个参数 开始截图的x坐标
    # 第二个参数 开始截图的y坐标
    # 第三个参数 结束截图的x坐标
    # 第四个参数 结束截图的y坐标
    bbox = (900, 420, 1129, 611)
    im = ImageGrab.grab(bbox)
    # 参数 保存截图文件的路径
    im.save('yzm_screencap.png')


def canny():
    # 读入图像
    lenna = cv2.imread("yzm_screencap.png", 0)
    # 图像降噪
    lenna = cv2.GaussianBlur(lenna, (5, 5), 0)
    # Canny边缘检测，50为低阈值low，150为高阈值high
    canny = cv2.Canny(lenna, 150, 190)
    cv2.imwrite('yzm_screencap.png', canny)


#滑动鼠标内部用到的函数
def get_track(distance):
    # 拿到移动轨迹，模仿人的滑动行为，先匀加速后匀减速
    # 匀变速运动基本公式：
    # ①v=v0+at
    # ②s=v0t+(1/2)at²
    # ③v²-v0²=2as
    #
    # :param distance: 需要移动的距离
    # :return: 存放每0.2秒移动的距离
    # 初速度
    v = 0
    # 单位时间为0.2s来统计轨迹，轨迹即0.2内的位移
    t = 0.1
    # 位移/轨迹列表，列表内的一个元素代表0.2s的位移
    tracks = []
    # 当前的位移
    current = 0
    # 到达mid值开始减速
    mid = distance * 4 / 5

    distance += 10  # 先滑过一点，最后再反着滑动回来

    while current < distance:
        if current < mid:
            # 加速度越小，单位时间的位移越小,模拟的轨迹就越多越详细
            a = 2  # 加速运动
        else:
            a = -3  # 减速运动

        # 初速度
        v0 = v
        # 0.2秒时间内的位移
        s = v0 * t + 0.5 * a * (t ** 2)
        # 当前的位置
        current += s
        # 添加到轨迹列表
        tracks.append(round(s))

        # 速度已经达到v,该速度作为下次的初速度
        v = v0 + a * t

    # 反着滑动到大概准确位置
    for i in range(3):
        tracks.append(-2)
    for i in range(4):
        tracks.append(-1)
    return tracks




# 模拟人鼠标滑动操作：缓动滑动（非线性）
def slip_pic(x, y, x_1, y_1):
    my_x_1_half = (x_1 - x)
    # my_x_2 = x + my_x_1_half
    # my_x_3 = x + my_x_1_half + my_x_1_half
    # print(math.ceil(float(my_x_1_half)))
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    time.sleep(1)
    track_list = get_track(my_x_1_half)
    for track in track_list:
        x = x + track
        # print('for: track is ' + str(x))
        m.move(int(x), y_1)
        time.sleep(0.02)
    # m.move(int(my_x_2), y_1)
    # time.sleep(1)
    # m.move(int(my_x_3), y_1)
    # time.sleep(2)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
    time.sleep(1)


#模拟鼠标滑动操作：线性
def mouse_move(x, y, x1, y1):
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)
    time.sleep(1)
    ctypes.windll.user32.SetCursorPos(x1, y1)
    time.sleep(1)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)
    time.sleep(2)


# 验证码滑动示例
def my_22_sms(my_phone_num):
    # 点击空白区域
    click(1300, 462)
    # 点击手机输入框
    click(930, 462)
    # 模拟键盘输入手机号
    string_input(my_phone_num)
    # 点击空白区域
    click(1300, 462)
    # 点击获取验证码
    click(1110, 516)
    time.sleep(4)
    for my_index in range(1, 6, 1):
        my_find_x_f = get_find_x()
        if my_find_x_f == 0:
            # print('my_find_x_f is 0')
            # 点击刷新
            click(1075, 712)
            time.sleep(2)
        else:
            # print(my_find_x_f)
            my_int = 849 + 52 + math.ceil(float(my_find_x_f))
            slip_pic(849, 666, int(my_int), 666)
            # print('滑动了，跳出循环')
            break
    # print('已经跳出循环')
    # 关闭浏览器窗口
    ctrl_w()


# if对比图片并点击
def ifmatchImgClick(myScreencap, mypng):
    if matchImg(myScreencap, mypng) is not None:
        print("-------------！" + mypng + str(
            matchImg(myScreencap, mypng)['result'][0]) + ',' + str(
            matchImg(myScreencap, mypng)['result'][1]))
        myx = str(matchImg(myScreencap, mypng)['result'][0])
        myy = str(matchImg(myScreencap, mypng)['result'][1])
        # click(int(float(myx)), int(float(myy)))
        # print("-------------对比图片return True")
        return True
        # time.sleep(2)
    else:
        # print("-------------对比图片return False")
        return False



# 从上到下找图片：my_img_height为小图片的高度，即为每次增加的像素
def matchImg_up_down(imgsrc, imgobj, my_img_height):
    img = cv2.imread(imgsrc)
    y0, y1, x0, x1 = 0, 0, 0, 1920
    for step_y1 in range(my_img_height, 1080, my_img_height):
        cropped = img[y0:step_y1, x0:x1]  # 裁剪坐标为[y0:y1, x0:x1]
        cv2.imwrite('tmp_' + imgsrc, cropped)
        if ifmatchImgClick('tmp_' + imgsrc, imgobj):
            # print('找到了，' + str(step_y1))
            matchImgClick('tmp_' + imgsrc, imgobj)
            break
        else:
            pass



# 对比图片并点击——偏移量my_x_offset   my_y_offset
def matchImgClick_offset(myScreencap, mypng, my_x_offset, my_y_offset):
    if matchImg(myScreencap, mypng) is not None:
        print("-------------点击按钮！" + mypng + str(
            matchImg(myScreencap, mypng)['result'][0]) + ',' + str(
            matchImg(myScreencap, mypng)['result'][1]))
        myx = str(matchImg(myScreencap, mypng)['result'][0])
        myy = str(matchImg(myScreencap, mypng)['result'][1])
        click(int(float(myx) + my_x_offset), int(float(myy) + my_y_offset))
        time.sleep(2)
        print("-------------结束点击按钮。")


# 返回坐标
def matchImg_return_x_y(myScreencap, mypng):
    if matchImg(myScreencap, mypng) is not None:
        print("-------------！" + mypng + str(
            matchImg(myScreencap, mypng)['result'][0]) + ',' + str(
            matchImg(myScreencap, mypng)['result'][1]))
        myx = str(matchImg(myScreencap, mypng)['result'][0])
        myy = str(matchImg(myScreencap, mypng)['result'][1])
        # click(int(float(myx)), int(float(myy)))
        # print("-------------对比图片return True")
        return matchImg(myScreencap, mypng)['result'][0], matchImg(myScreencap, mypng)['result'][1]
        # time.sleep(2)
    else:
        # print("-------------对比图片return False")
        return 0,0




# 图片添加字
def add_num(im01, mypng, x, y):
    img = Image.open(im01)
    imgmypng = Image.open(mypng)
    # ImageDraw.Draw()函数会创建一个用来对image进行操作的对象，
    # 对所有即将使用ImageDraw中操作的图片都要先进行这个对象的创建。
    draw = ImageDraw.Draw(img)

    # 设置字体和字号
    myfont = ImageFont.truetype('C:/windows/fonts/Arial.TTF', size=20)

    # 设置要添加的数字的颜色为红色
    fillcolor = "#ff0000"

    # 昨天博客中提到过的获取图片的属性
    width, height = imgmypng.size

    # 设置添加数字的位置，具体参数可以自己设置，从左上角开始
    draw.text((x - width / 2, y - height / 2), '999', font=myfont, fill=fillcolor)

    # 将修改后的图片以格式存储
    img.save(im01, 'png')

    return 0


# 涂鸦掉图片
def matchImg_delete2Continue(myScreencap, mypng):
    if matchImg(myScreencap, mypng) is not None:
        print("-------------涂鸦pic！" + mypng + str(
            matchImg(myScreencap, mypng)['result'][0]) + ',' + str(
            matchImg(myScreencap, mypng)['result'][1]))
        myx = str(matchImg(myScreencap, mypng)['result'][0])
        myy = str(matchImg(myScreencap, mypng)['result'][1])
        add_num(myScreencap, mypng, int(float(myx)), int(float(myy)))
        print("-------------结束点击按钮。")




if __name__ == '__main__':
    # 定义鼠标键盘实例
    k = PyKeyboard()
    m = PyMouse()

    while 1 == 1:
        my_time = datetime.datetime.now()
        print(my_time)
        time.sleep(10)
