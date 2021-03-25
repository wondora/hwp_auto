import win32com.client as win32
import pandas as pd

hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule") # 보안 모듈 물어보지 않는다.
hwp.Open("D:/경기과학고/상장인쇄/당선증2.hwp")
hwp.XHwpWindows.Item(0).Visible = True

f_list = hwp.GetFieldList().split("\x02")
print(f_list)
hwp.SaveAs(os.path.join(os.getcwd(). "test.hwp")) # 다른이름 저장


for i in range(len(excel)-1):  # ‘-1’은 기존에 한쪽 있기 때문
    hwp.Run('Paste')
    hwp.MovePos(3)

