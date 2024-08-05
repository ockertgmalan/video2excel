import os
import pyautogui
import time
import subprocess

import pygetwindow as gw

def take_screenshot(file_path, output_folder):
    base_name = os.path.basename(file_path).replace('.xls', '.png')
    outname = os.path.join(output_folder, base_name)
    if not os.path.exists(outname):
        print(f"Processing file: {file_path}")
        # Open the Excel file
        subprocess.Popen(['start', 'excel', file_path], shell=True)
        time.sleep(1)  # Wait for Excel to open

        # Handle potential popup by hitting Tab twice and Enter
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)     
        
        pyautogui.press('enter')
        time.sleep(0.1)

        # Ensure Excel is the active window
        excel_windows = gw.getWindowsWithTitle('Excel')
        if excel_windows:
            excel_window = excel_windows[0]
            excel_window.activate()
            time.sleep(0.1)

        # Adjust the zoom level using the ribbon
        pyautogui.press('alt')
        time.sleep(0.1)
        pyautogui.press('w')
        time.sleep(0.1)
        pyautogui.press('q')
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.hotkey('1', '0')  # Set the zoom level to 10%
        time.sleep(0.1)
        pyautogui.press('enter')
        time.sleep(0.5)

        # Take a screenshot of the specific window
        left, top, right, bottom = excel_window.left, excel_window.top, excel_window.right, excel_window.bottom
        screenshot = pyautogui.screenshot(region=(left, top, right-left, bottom-top))
        
        screenshot.save(outname)

        # Close Excel
        pyautogui.hotkey('alt', 'f4')
        time.sleep(0.5)  # Wait for Excel to close

def main(input_folder, output_folder):
    print(f"Input folder: {input_folder}")
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file_name in os.listdir(input_folder):
        print(f"Processing file: {file_name}")
        if file_name.endswith('.xls'):
            file_path = os.path.join(input_folder, file_name)
            print('taking screenshot')
            take_screenshot(file_path, output_folder)
    print("Screenshot capture complete.")
if __name__ == "__main__":
    input_folder = r'G:\Projects\Excel2video\pixel-spreadsheet\stardog_spreadsheets'
    output_folder = r'G:\Projects\Excel2video\pixel-spreadsheet\stardog_in_excel'
    main(input_folder, output_folder)