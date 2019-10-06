import pybcapclient.robot_functions as utility
import cv2 as cv
import win32com.client


def main():
    # Set parameters
    host = "192.168.0.1"
    port = 5007
    timeout = 14400
    eng = win32com.client.Dispatch("CAO.CaoEngine")
    ctrl = eng.Workspaces(0).AddController("", "CaoProv.DENSO.RC8", "", "Server=" + host)
    caoRobot = ctrl.AddRobot("robot0", "")

    (client, hCtrl, hRobot) = utility.connect(host, port, timeout)

    utility.move_to_photo_position(client, hRobot)
    image = utility.take_img()
    # cv.imshow("Image", image)
    # cv.waitKey()
    # cv.destroyWindow("Image")
    text = utility.tesseract_ocr(image)
    utility.move_to_the_highlighter(client, hRobot)
    # switch to use the gripper
    utility.switch_bcap_to_orin(client, hRobot, caoRobot)
    # print("Switch to Orin")
    # gripper code here, execution of the .pcs file to grab the highlighter

    utility.switch_orin_to_bcap(client, hRobot, caoRobot)
    # print("Switch to Bcap")

    utility.disconnect(client, hCtrl, hRobot)


if __name__ == "__main__":
    main()
