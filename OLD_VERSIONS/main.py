import pandas as pd
import pyautogui
from time import sleep
import logging
import keyboard
from pynput import mouse, keyboard
from playsound import playsound


# Configure logging
logger = logging.getLogger('ComaxLogger')
logger.setLevel(logging.DEBUG)

# Create file handler for logging to a file
file_handler = logging.FileHandler('log.txt')
file_handler.setLevel(logging.DEBUG)

# Create console handler for logging to the terminal
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# Create formatter and set it for both handlers
formatter = logging.Formatter('%(asctime)s,%(msecs)d %(name)s %(levelname)s %(message)s', datefmt='%H:%M:%S')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Add handlers to the logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)

def get_cursor_position(message):
    global ctrl_left_click_detected
    ctrl_left_click_detected = False
    print(message + " Press Ctrl + Left Click at the position.")

    def on_click(x, y, button, pressed):
        if pressed and button == mouse.Button.left:
            if ctrl_listener.ctrl_pressed:
                global ctrl_left_click_detected
                ctrl_left_click_detected = True
                position = (x, y)
                print(f"Position confirmed at {position}.")
                logger.debug(f"Position registered: {position}")
                return False

    def on_press(key):
        if key == keyboard.Key.ctrl_l:
            ctrl_listener.ctrl_pressed = True

    def on_release(key):
        if key == keyboard.Key.ctrl_l:
            ctrl_listener.ctrl_pressed = False

    with mouse.Listener(on_click=on_click) as listener:
        with keyboard.Listener(on_press=on_press, on_release=on_release) as ctrl_listener:
            ctrl_listener.ctrl_pressed = False
            listener.join()

    if ctrl_left_click_detected:
        return pyautogui.position()







# Define coordinates for clicks
uuid_field_pos = get_cursor_position("Place your cursor on the UUID input field.")
update_btn_pos = get_cursor_position("Place your cursor on the Update button.")
version_field_pos = get_cursor_position("Place your cursor on the version input field.")
confirm_btn_pos = get_cursor_position("Place your cursor on the Confirm button.")

sleep(5) # wait a while so the user opens the window
def update_udid(udid, version, customer, index, total):
    try:
        # Click on the UDID field and enter the UDID
        pyautogui.click(uuid_field_pos)
        pyautogui.hotkey('ctrl', 'a')  # Select all text
        pyautogui.press('backspace')  # Clear the field
        pyautogui.write(udid)  # Write the new UDID
        pyautogui.press('enter')
        sleep(0.6)

        # Click the update button
        pyautogui.click(update_btn_pos)
        sleep(1)

        # Click the version field and clear it
        pyautogui.click(version_field_pos)
        sleep(1)

        for _ in range(5):  # Clear the version field
            pyautogui.press('backspace')
            sleep(0.1)

        # Enter the new version number
        pyautogui.write(version)  # New version to enter
        sleep(0.5)

        # Click the confirm button
        pyautogui.click(confirm_btn_pos)
        sleep(1)

        # Log success message
        logger.info(f'{index + 1}/{total} - UDID: {udid} of {customer} updated successfully to version {version}.')
    except Exception as e:
        logger.error(f"Failed to update UDID {udid}: {e}")

for file in range(0, 2):
# Load UDIDs from the Excel file with error handling
    try:
        df = pd.read_excel(r'C:\Users\benm\PycharmProjects\AutoCheckActive\Assets\Excels\toUpdate\output_{}.xlsx'.format(file))
        # Process each UDID in the DataFrame
        total_udids = len(df['UDID'])
        if 'UDID' not in df.columns or 'Version' not in df.columns:
            raise ValueError("Missing required columns 'UDID' or 'ver' in the Excel file.")
    except Exception as e:
        logger.error(f"Failed to load or validate Excel file: {e}")
        raise SystemExit("Error: Could not process the Excel file.")
    logger.info(f"Running excel file 'output_{file}.xlsx'")

    for num, row in df.iterrows():
        update_udid(row['UDID'],'808', row['Customer'], num, total_udids)

logger.info("All UDIDs processed.")