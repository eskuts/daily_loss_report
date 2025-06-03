import locale
import os
import subprocess
import time
from datetime import datetime, timedelta

import pyautogui
import pyperclip
import win32com.client
import win32gui

locale.setlocale(locale.LC_TIME, "Russian_Russia.1251")


def find_coordinastes():
    try:
        while True:
            x, y = pyautogui.position()  # Получаем текущие координаты курсора
            print(f"Current Mouse Position: X={x}, Y={y}")
            time.sleep(1)
    except KeyboardInterrupt:
        print("Stop")
    return 0


def clicker(prev_day, path):
    pyautogui.moveTo(883, 595)  # переводит курсор на рабочий экран
    shortcut_path = r"C:\Users\e.kuts\Desktop\Reports.exe — ярлык.lnk"
    subprocess.Popen(["cmd", "/c", shortcut_path])
    time.sleep(2)

    hwnd = win32gui.FindWindow(None, "Введите ваш логин и пароль")
    rect = win32gui.GetClientRect(hwnd)  # (0, 0, 401, 174)

    pyautogui.moveTo(883, 595)  # нажимает ок
    pyautogui.click()
    time.sleep(100)

    pyautogui.moveTo(174, 73)  # нажимает пересчет-1
    pyautogui.click()
    time.sleep(2)

    pyautogui.moveTo(490, 256)  # место для ввода даты
    pyautogui.click()

    pyautogui.write(str(prev_day), interval=0.1, _pause=True)  # день месяца

    pyautogui.moveTo(517, 202)  # нажимает пересчет-2
    pyautogui.click()
    time.sleep(120)

    pyautogui.moveTo(96, 167)  # нажимает поиск
    pyautogui.click()
    pyautogui.write("-3")  # ищет нужный отчёт
    pyautogui.moveTo(347, 459)  # выбираем оп-3м
    pyautogui.doubleClick()
    time.sleep(5)

    pyautogui.moveTo(502, 254)  # место для ввода даты
    pyautogui.click()

    pyautogui.write(str(prev_day), interval=0.1, _pause=True)  # дата

    pyautogui.moveTo(515, 200)  # нажимает отчёт
    pyautogui.click()
    time.sleep(60)
    pyautogui.moveTo(570, 235)  # сохр файл
    pyautogui.click()
    time.sleep(2)
    pyautogui.moveTo(573, 523)  # csv
    pyautogui.click()
    time.sleep(2)
    pyautogui.press("enter")  # Нажимаем Enter

    pyautogui.moveTo(598, 52)  # окно для ввода пути сохранения файла
    pyautogui.click()
    time.sleep(2)

    pyperclip.copy(path)  # копируем путь в буфер чтобы оттуда вставить потом
    pyautogui.hotkey(
        "ctrl", "v"
    )  # вводим путь сохранения в формате Сеть\DC\share\Аналитика\ОП-3М_СОО\Январь\CSV
    time.sleep(2)
    pyautogui.press("enter")  # Нажимаем Enter
    time.sleep(2)

    pyautogui.moveTo(563, 440)  # наводится на имя файла
    pyautogui.click()
    pyautogui.press("enter")  # Нажимаем Enter
    time.sleep(2)
    pyautogui.press("enter")  # Нажимаем Enter
    time.sleep(2)
    pyautogui.hotkey("alt", "f4")
    return 0


def kill_excel():
    """Принудительное завершение всех процессов Excel."""
    os.system("taskkill /F /IM excel.exe")


def safe_refresh(workbook, timeout=60):
    """Безопасное обновление данных с таймаутом."""
    start_time = time.time()
    workbook.RefreshAll()
    while time.time() - start_time < timeout:
        try:
            if workbook.Application.CalculationState == -1:  # xlDone
                break
        except Exception:
            break
        time.sleep(1)
    else:
        print("Excel завис при обновлении. Принудительное завершение.")
        kill_excel()
        raise TimeoutError("Excel завис во время обновления.")


def run_data_extractor():
    # find_coordinastes()

    prev_day = datetime.today() - timedelta(days=1)
    str_prev_day = prev_day.strftime("%d.%m.%Y")
    month = prev_day.strftime("%B").capitalize()

    path = os.environ[r"path_to_op3m_report"] + "\CSV"
    path = path.replace("month", month)
    clicker(str_prev_day, path)

    file_path = os.environ[r"path_to_op3m_report"].replace("month", month)

    excel = win32com.client.Dispatch("Excel.Application")
    workbook = excel.Workbooks.Open(file_path)
    time.sleep(10)
    workbook.RefreshAll()
    excel.CalculateUntilAsyncQueriesDone()
    time.sleep(120)
    workbook.Save()
    workbook.Close(SaveChanges=True)
    excel.Quit()


if __name__ == "__main__":
    raise SystemExit(run_data_extractor())
