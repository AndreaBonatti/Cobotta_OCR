import pybcapclient.robot_functions as utility
import cv2 as cv
import win32com.client
import time


def main():
    # Set parameters
    host = "192.168.0.1"
    port = 5007
    timeout = 14400
    eng = win32com.client.Dispatch("CAO.CaoEngine")
    ctrl = eng.Workspaces(0).AddController("", "CaoProv.DENSO.RC8", "", "Server=" + host)
    caoRobot = ctrl.AddRobot("robot0", "")

    (client, hCtrl, hRobot) = utility.connect(host, port, timeout)

    # Open the hand of robot at the begin
    utility.switch_bcap_to_orin(client, hRobot, caoRobot)
    # print("Switch to Orin")
    ctrl.Execute("HandMoveH", [20, 0])
    utility.switch_orin_to_bcap(client, hRobot, caoRobot)
    # print("Switch to Bcap")

    utility.move_to_photo_position(client, hRobot)
    # image = utility.take_img()
    # cv.imshow("Image", image)
    # cv.waitKey()
    # cv.destroyWindow("Image")
    # text = utility.tesseract_ocr(image)
    utility.move_to_the_highligther(client, hRobot)

    # switch to use the gripper
    utility.switch_bcap_to_orin(client, hRobot, caoRobot)
    # print("Switch to Orin")
    ctrl.Execute("HandMoveH", [20, 1])
    utility.switch_orin_to_bcap(client, hRobot, caoRobot)
    # print("Switch to Bcap")

    utility.move_to_initial_writing_position(client, hRobot)
    # Time to put a sheet under the robot
    time.sleep(20)
    utility.move_to_the_sheet(client, hRobot)

    utility.test_writing(client, hRobot)
    utility.go_up(client, hRobot)
    utility.replace_the_highlighter(client, hRobot)
    utility.switch_bcap_to_orin(client, hRobot, caoRobot)
    # print("Switch to Orin")
    ctrl.Execute("HandMoveH", [20, 0])
    utility.switch_orin_to_bcap(client, hRobot, caoRobot)
    # print("Switch to Bcap")

    utility.go_up(client, hRobot)
    utility.disconnect(client, hCtrl, hRobot)


if __name__ == "__main__":
    main()
