import io
import os

import cv2
import numpy as np
import pytesseract
import win32com.client
import win32com.client
from PIL import Image

import pybcapclient.bcapclient as bcapclient
from pybcapclient.bcapclient import BCAPClient


def connect(host, port, timeout, provider="CaoProv.DENSO.VRC"):
    client = bcapclient.BCAPClient(host, port, timeout)
    client.service_start("")
    Name = ""
    Provider = provider
    Machine = ("localhost")
    Option = ("")
    hCtrl = client.controller_connect(Name, Provider, Machine, Option)
    hRobot = client.controller_getrobot(hCtrl, "Arm0", "")
    client.robot_execute(hRobot, "TakeArm", [0, 0])
    client.robot_execute(hRobot, "Motor", [1, 0])
    return (client, hCtrl, hRobot)


def disconnect(client, hCtrl, hRobot):
    client.robot_execute(hRobot, "Motor", 0)
    client.robot_execute(hRobot, "GiveArm")
    client.controller_disconnect(hCtrl)
    client.service_stop()


def robot_getvar(client, hRobot, name):
    assert isinstance(client, BCAPClient)
    var_handle = client.robot_getvariable(hRobot, name)
    value = client.variable_getvalue(var_handle)
    client.variable_release(var_handle)
    return value


def take_img(CVconv=True, wb=False, oneshotfocus=False):
    eng = win32com.client.Dispatch("CAO.CaoEngine")
    ctrl = eng.Workspaces(0).AddController("", "CaoProv.Canon.N10-W02", "", "Server=192.168.0.90" + ", Timeout=5000")
    image_handle = ctrl.AddVariable("IMAGE")
    if wb:
        ctrl.Execute("OneShotWhiteBalance")
    if oneshotfocus:
        ctrl.Execute("OneShotFocus")
    image = image_handle.Value
    stream = io.BytesIO(image)
    img = Image.open(stream)
    opencvImage = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    del image_handle
    del ctrl
    del eng
    if CVconv:
        return opencvImage
    else:
        return img


def switch_bcap_to_orin(client, hRobot, caoRobot):
    client.robot_execute(hRobot, "GiveArm")
    caoRobot.Execute("TakeArm", [0, 0])
    caoRobot.Execute("Motor", [1, 0])


def switch_orin_to_bcap(client, hRobot, caoRobot):
    caoRobot.Execute("GiveArm")
    client.robot_execute(hRobot, "TakeArm", [0, 0])
    client.robot_execute(hRobot, "Motor", [1, 0])


def list_to_string_position(pos):
    return "P(" + ", ".join(str(i) for i in pos) + ")"


def list_to_string_joints(pos):
    return "J(" + ", ".join(str(i) for i in pos) + ")"


def move_to_new_pos(client, hRobot, new_x, new_y, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")
    curr_pos[0] = new_x
    curr_pos[1] = new_y
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def move_to_photo_position(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")
    # curr_pos = [x, y, z, Rx, Ry, Rz]
    curr_pos[0] = 150
    curr_pos[1] = -45
    curr_pos[2] = 235
    curr_pos[3] = 180
    curr_pos[4] = 0
    curr_pos[5] = 180
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def move_to_the_highligther(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")
    # curr_pos = [x, y, z, Rx, Ry, Rz]
    # first move on y axis to avoid to hit the highlighter
    curr_pos[1] = -145
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] = 145
    curr_pos[2] = 86
    curr_pos[3] = 180
    curr_pos[4] = 0
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[5] = 180
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def replace_the_highlighter(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")
    curr_pos[0] = 145
    curr_pos[1] = -145
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 89
    curr_pos[3] = 180
    curr_pos[4] = 0
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def go_up(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")
    curr_pos[2] = 200
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def tesseract_ocr(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # write the grayscale image to disk as a temporary file so we can apply OCR to it
    filename = "{}.png".format(os.getpid())
    cv2.imwrite(filename, gray)
    # load the image as a PIL/Pillow image, apply OCR, and then delete the temporary file
    text = pytesseract.image_to_string(Image.open(filename))
    os.remove(filename)
    # print the text that is read by the OCR
    print(text)
    return text


def move_to_initial_writing_position(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")
    # curr_pos = [x, y, z, Rx, Ry, Rz]
    curr_pos[2] = 235
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] = 180
    curr_pos[1] = -45
    curr_pos[3] = 180
    curr_pos[4] = 0
    curr_pos[5] = 180
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


# 20 of old_pos[0]-curr_pos[0] = 2cm more or less
def write_a_character(char, client, hRobot):
    if char == "1":
        write_one(client, hRobot)
    elif char == "2":
        write_two(client, hRobot)
    elif char == "3":
        write_three(client, hRobot)
    elif char == "4":
        write_four(client, hRobot)
    elif char == "5":
        write_five(client, hRobot)
    elif char == "6":
        write_six(client, hRobot)
    elif char == "7":
        write_seven(client, hRobot)
    elif char == "8":
        write_eight(client, hRobot)
    elif char == "9":
        write_nine(client, hRobot)
    elif char == "0":
        write_zero(client, hRobot)
    elif char == "A":
        write_maiusc_a(client, hRobot)
    elif char == "B":
        write_maiusc_b(client, hRobot)
    elif char == "C":
        write_maiusc_c(client, hRobot)
    elif char == "D":
        write_maiusc_d(client, hRobot)
    elif char == "E":
        write_maiusc_e(client, hRobot)
    elif char == "F":
        write_maiusc_f(client, hRobot)
    elif char == "G":
        write_maiusc_g(client, hRobot)
    elif char == "H":
        write_maiusc_h(client, hRobot)
    elif char == "I":
        write_maiusc_i(client, hRobot)


def write_one(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")
    curr_pos[0] -= 5
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_two(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[0] -= 5
    curr_pos[1] += 5
    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[0] += 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[0] -= 15
    curr_pos[1] += 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 20
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_three(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[0] -= 5
    curr_pos[1] += 5
    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[0] += 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[0] -= 5
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # Half of the three
    curr_pos[0] -= 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[0] -= 5
    curr_pos[1] += 5
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[0] += 5
    curr_pos[1] += 5
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 15
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_four(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[0] -= 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 15
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 15
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_five(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] += 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] -= 6
    client.robot_move(hRobot, 2, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] += 6
    client.robot_move(hRobot, 2, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 20
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_six(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 10
    curr_pos[1] += 10
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] -= 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] += 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 6
    curr_pos[1] += 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 16
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_seven(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 20
    curr_pos[1] += 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 20
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_eight(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")
    curr_pos[2] = 85
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] += 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] -= 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] -= 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] += 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[0] += 5
    curr_pos[1] += 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] -= 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] -= 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] += 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] += 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_nine(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[0] -= 20
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 10
    curr_pos[1] -= 10
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] += 5
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] += 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] -= 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 6
    curr_pos[1] -= 6
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 4
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_zero(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[0] -= 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] += 5
    client.robot_move(hRobot, 2, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, 2, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 10
    curr_pos[1] += 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 15
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_maiusc_a(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[0] -= 20
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 20
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 20
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 10
    curr_pos[1] += 2.5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 10
    curr_pos[1] -= 2.5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_maiusc_b(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] += 6
    client.robot_move(hRobot,  1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] -= 5
    client.robot_move(hRobot,  1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 5
    curr_pos[1] += 5
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_maiusc_c(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] += 7
    curr_pos[0] -= 3
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] += 3
    curr_pos[0] -= 7
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 10
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 20
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_maiusc_d(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[0] -= 20
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 7
    curr_pos[0] -= 3
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 3
    curr_pos[0] -= 7
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 10
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] += 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 20
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_maiusc_e(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[0] -= 20
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 10
    curr_pos[1] += 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 10
    curr_pos[1] += 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 20
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_maiusc_f(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[0] -= 20
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 10
    curr_pos[1] += 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 10
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_maiusc_g(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[1] -= 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] += 7
    curr_pos[0] -= 3
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] += 3
    curr_pos[0] -= 7
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 10
    client.robot_move(hRobot, 1, "@P " + list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[0] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 15
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_maiusc_h(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] += 10
    curr_pos[1] += 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 10
    curr_pos[1] += 5
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_maiusc_i(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[2] = 85
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # to allow the subsequent characters to be written on the same line
    curr_pos[0] += 20
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def to_the_next_char(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")

    curr_pos[1] -= 15
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def write_a_word(word, client, hRobot):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")
    curr_pos[1] = 70
    client.robot_move(hRobot, 2, list_to_string_position(curr_pos), "SPEED=100")

    for c in word:
        write_a_character(c, client, hRobot)
        to_the_next_char(client, hRobot)
