import pyautogui as pgui

w, h = pgui.size()
print("{\"width\":%d,\"height\":%d}" % (w,h))
pgui.moveTo(w-25, 23, duration=0.0)
pgui.click(w-25, 23)
