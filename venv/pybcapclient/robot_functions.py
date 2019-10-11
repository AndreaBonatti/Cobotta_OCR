import io
import os

import cv2
import math
import numpy as np
import pytesseract
import win32com.client
import win32com.client
from PIL import Image
from numpy import pi

import pybcapclient.bcapclient as bcapclient
from pybcapclient.bcapclient import BCAPClient

PIX_MM_RATIO = 9.222
INVERT = -1

CAMERA_ROBOT_DISTANCE = 52.38570925983608  # mm
OFFSET_CONF = 0.25  # mm
CAMERA_ROBOT_DISTANCE_Y = 0.9874772462474191  # mm


def polar_to_robot_coordinates(angle, robot_x, robot_y, module=CAMERA_ROBOT_DISTANCE):
    offset_x = module * math.cos(math.radians(angle))
    offset_y = -module * math.sin(math.radians(angle))
    return robot_x + offset_x, robot_y + offset_y


def pixels_to_cartesian(img_x, img_y, width=1920, height=1080):  # per coordinate di opencv, non numpy
    cartesian_x = img_x - width / 2.0
    cartesian_y = -img_y + height / 2.0
    return cartesian_x, cartesian_y


def find_polar_coordinates(angle, camera_x, camera_y):
    (x, y) = pixels_to_cartesian(camera_x, camera_y)
    if np.absolute(x) > math.ceil(PIX_MM_RATIO):
        L2 = math.sqrt(x ** 2 + y ** 2) / PIX_MM_RATIO
        alpha2 = math.atan2(y, x) * (180.0 / pi)

        if x > 0 and y < 0:
            alpha3 = 90 - np.absolute(alpha2)
            L3 = math.sqrt(
                CAMERA_ROBOT_DISTANCE ** 2 + L2 ** 2 - 2 * L2 * CAMERA_ROBOT_DISTANCE * math.cos(math.radians(alpha3)))
            alpha4 = math.degrees(math.asin((float(L2) / L3) * (math.sin(math.radians(alpha3)))))
            alpha5 = angle + alpha4

        if x > 0 and y >= 0:
            alpha3 = 90 + alpha2
            L3 = math.sqrt(
                CAMERA_ROBOT_DISTANCE ** 2 + L2 ** 2 - 2 * L2 * CAMERA_ROBOT_DISTANCE * math.cos(math.radians(alpha3)))
            alpha4 = math.degrees(math.asin((float(L2) / L3) * (math.sin(math.radians(alpha3)))))
            alpha5 = angle + alpha4

        if x < 0 and y >= 0:
            alpha3 = 360 - (np.absolute(alpha2) + 90)
            L3 = math.sqrt(
                CAMERA_ROBOT_DISTANCE ** 2 + L2 ** 2 - 2 * L2 * CAMERA_ROBOT_DISTANCE * math.cos(math.radians(alpha3)))
            alpha4 = math.degrees(math.asin((float(L2) / L3) * (math.sin(math.radians(alpha3)))))
            alpha5 = angle - alpha4

        if x < 0 and y < 0:
            alpha3 = np.absolute(alpha2 + 90)
            L3 = math.sqrt(
                CAMERA_ROBOT_DISTANCE ** 2 + L2 ** 2 - 2 * L2 * CAMERA_ROBOT_DISTANCE * math.cos(math.radians(alpha3)))
            alpha4 = math.degrees(math.asin((float(L2) / L3) * (math.sin(math.radians(alpha3)))))
            alpha5 = angle - alpha4

        return (L3, alpha5)

    else:
        L3 = CAMERA_ROBOT_DISTANCE + (y / PIX_MM_RATIO)
        return (L3, angle)


def find_orientation(contour, robot_angle):
    (_, _), (_, _), angle = cv2.fitEllipse(contour)
    new_angle = 0

    if angle <= 90:
        if robot_angle <= 0:
            new_angle = robot_angle + angle
        elif robot_angle > 0 and angle >= 50:
            beta = 90 - angle
            gamma = 90 - robot_angle
            new_angle = -(gamma + beta)
        elif robot_angle > 0 and angle < 50:
            new_angle = robot_angle + angle
    else:
        if robot_angle > 0:
            beta = 180 - angle
            new_angle = robot_angle - beta
        elif robot_angle <= 0 and angle <= 160:
            new_angle = robot_angle + angle
        elif robot_angle <= 0 and angle > 160:
            beta = 180 - angle
            new_angle = robot_angle - beta

    return new_angle


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
    # Orange highligther curr_pos[2] = 108, green one =84
    curr_pos[2] = 84
    curr_pos[3] = 180
    curr_pos[4] = 0
    curr_pos[5] = 90
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[5] = 180
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
    curr_pos[0] = 180
    curr_pos[1] = -45
    curr_pos[2] = 235
    curr_pos[3] = 180
    curr_pos[4] = 0
    curr_pos[5] = 180
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def move_to_the_sheet(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")
    # curr_pos = [x, y, z, Rx, Ry, Rz]
    curr_pos[2] = 86
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")


def test_writing(client, hRobot, mode=2):
    curr_pos = robot_getvar(client, hRobot, "@CURRENT_POSITION")
    # 20 of old_pos[0]-curr_pos[0] = 2cm more or less
    # Test 1: I
    # curr_pos[0] = 160
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[2] = 105
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # Test 2: 1
    # curr_pos[0] = 160
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[2] = 105
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[0] = 180
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[2] = 86
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[0] -= 5
    # curr_pos[1] += 5
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[2] = 105
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # Test 3: 11
    # curr_pos[0] = 160
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[2] = 105
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[0] = 180
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[2] = 86
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[0] -= 5
    # curr_pos[1] += 5
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[2] = 105
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    #
    # curr_pos[0] = 180
    # curr_pos[1] -= 20
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[2] = 86
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    #
    # curr_pos[0] = 160
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[2] = 105
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[0] = 180
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[2] = 86
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[0] -= 5
    # curr_pos[1] += 5
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    # curr_pos[2] = 105
    # client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    # Test 4: 2
    curr_pos[0] -= 5
    curr_pos[1] += 5
    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 86
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[0] += 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, 1, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[0] -= 5
    curr_pos[1] -= 5
    client.robot_move(hRobot, 1, list_to_string_position(curr_pos), "SPEED=100")

    curr_pos[0] -= 15
    curr_pos[1] += 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[1] -= 10
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
    curr_pos[2] = 105
    client.robot_move(hRobot, mode, list_to_string_position(curr_pos), "SPEED=100")
