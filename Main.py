import time
import comtypes.client as cc
from tkinter import *
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk
import logging
import os
import configparser
from threading import Thread
import sys
import serial.tools.list_ports
import socket
from logging.handlers import TimedRotatingFileHandler


# Setup Logging
scriptDir = os.path.dirname(__file__)  # Absolute dir the script is in
logname = 'LaserHypotCont.log'
absFilePath = os.path.join(scriptDir, logname)
logger = logging.getLogger('Rotating Log')
try:
    # Rotating Logs 1 every day, keep for 30 days
    handler = TimedRotatingFileHandler(absFilePath, when='h', interval=8, backupCount=30)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
    # Set the suffix to include the date format and '.log'
    handler.suffix = "%Y-%m-%d_%H-%M-%S.log"
    logger.addHandler(handler)
except Exception as e:
    print(f'Error creating new Log file, permissions error: {e}')

# Setup Settings File
config = configparser.ConfigParser()
config['Run Cavity'] = {}
config['Laser Enabled'] = {}
config['Admin'] = {}
config['Hypot'] = {}
config['Continuity'] = {}
config['Laser'] = {}
config['Hardware IDs'] = {}

# Driver Variables
cc.GetModule('SC6540.dll')
from comtypes.gen import SC6540Lib

cc.GetModule('ARI38XX_64.dll')
from comtypes.gen import ARI38XXLib

hypotSerial1 = "AQ03JGPEA"
sc6540Serial1 = "B0007EEKA"
sc6540Serial2 = "B0007BEKA"
# hypotSerial2 = "YOUR_SERIAL_NUMBER"
# Setup Laser Connectivity
laser = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # Creates socket
try:
    laser.connect(('127.0.0.1', 50000))  # IP and Port number for laser
except Exception as e:
    print(f'Connection to Laser Marker failed: {e}')
    logger.error(f'Connection to Laser Marker failed: {e}')

# General Variables
adminPassword = '6789'  # Default password if not set in the settings file
cavityContinuitySuccesses = {}
cavityHypotSuccesses = {}
runCavity = {}
laserEnabled = {}
hypotSettings = {
    'voltage': 1240,  # AC Voltage
    'current high limit': 20,  # Current High Limit
    'current low limit': 0,  # Current Low Limit
    'ramp up time': 0.1,  # Ramp up time in seconds
    'ramp down time': 2,  # Dewll time in seconds
    'dwell time': 0,  # RampDownTime in seconds
    'arcsense level': 5,  # ArcSense level
    'arc detection': True,  # Arc detection
    'frequency': ARI38XXLib.ARI38XXFrequency60Hz,  # Frequency
    'continuity test': False,  # Continuity test
    'high limit resistance': 1.5,  # High limit of the continuity resistance
    'low limit resistance': 0,  # Low limit of the continuity resistance
    'resistance offset': .5  # Continuity resistance offset
}
continuitySettings = {
    'voltage': 1240,  # AC Voltage
    'current high limit': 20,  # Current High Limit
    'current low limit': 0,  # Current Low Limit
    'ramp up time': 0.1,  # Ramp up time in seconds
    'ramp down time': 2,  # Dewll time in seconds
    'dwell time': 0,  # RampDownTime in seconds
    'arcsense level': 5,  # ArcSense level
    'arc detection': True,  # Arc detection
    'frequency': ARI38XXLib.ARI38XXFrequency60Hz,  # Frequency
    'continuity test': True,  # Continuity test
    'high limit resistance': 1.5,  # High limit of the continuity resistance
    'low limit resistance': 0,  # Low limit of the continuity resistance
    'resistance offset': .5  # Continuity resistance offset
}
# UI Variables
rectangles = {}
statusText = {}
root = tk.Tk()
root.geometry('1800x900')
root.title('Main')
backgroundColor = '#2A2E32'
canvasColor = '#3D434B'
enabledColor = '#26A671'
halfDisabledColor = '#F9A825'
disabledColor = '#DE0A02'
textBackgroundColor = '#2A2E32'
textColor = 'White'

root.configure(bg=backgroundColor)
# root.attributes('-toolwindow', True)  # Disables bar at min, max, close button in top right
faultState = False


def find_com_port_by_serial_number(targetSerialNumber):
    ports = serial.tools.list_ports.comports()
    for port in ports:
        # Print all device details for debugging purposes
        # print(f"Device: {port.device}, Description: {port.description}, HWID: {port.hwid}")
        if targetSerialNumber in port.hwid:
            logger.info('Serial number: ' + targetSerialNumber + ' Located at: ' + port.device)
            return port.device
    return None


def concat_port(comPort):
    return 'ASRL' + comPort.replace("COM", '') + '::INSTR'


# Avoid using COM# because windows can mix it up

hypotComPort1 = find_com_port_by_serial_number(hypotSerial1)
portNumHy1 = concat_port(hypotComPort1)

# hypotComPort2 = find_com_port_by_serial_number(hypotSerial2)
# portNumHy2 = cachevar(hypotComPort2)

sc6540ComPort1 = find_com_port_by_serial_number(sc6540Serial1)
portNumSC1 = concat_port(sc6540ComPort1)

sc6540ComPort2 = find_com_port_by_serial_number(sc6540Serial2)
portNumSC2 = concat_port(sc6540ComPort2)

print(portNumHy1)
print(portNumSC1)
print(portNumSC2)

# Driver Setup
try:
    hypotDriver1 = cc.CreateObject('ARI38XX.ARI38XX', interface=ARI38XXLib.IARI38XX)
    hypotDriver1.Initialize(portNumHy1, True, False, 'DriverSetup=BaudRate=38400')
except Exception as e:
    print(f'Connection to Hypot1 failed: {e}')
    logger.error(f'Connection to Hypot1 failed: {e}')

try:
    pass
    # hypotDriver2 = cc.CreateObject('ARI38XX.ARI38XX', interface=ARI38XXLib.IARI38XX)
    # hypotDriver2.Initialize('portNumHy2, True, False, 'DriverSetup=BaudRate=38400')
except Exception as e:
    print(f'Connection to Hypot2 failed: {e}')
    logger.error(f'Connection to Hypot2 failed: {e}')

try:
    sc6540Driver1 = cc.CreateObject('SC6540.SC6540', interface=SC6540Lib.ISC6540)
    sc6540OptionString1 = 'Cache=false, InterchangeCheck=false, QueryInstrStatus=true, RangeCheck=false, RecordCoercions=false, Simulate=false'
    sc6540Driver1.Initialize(portNumSC1, True, False, sc6540OptionString1)
except Exception as e:
    print(f'Connection to SC6540 Switch1 failed: {e}')
    logger.error(f'Connection to SC6540 Switch1 failed: {e}')

try:
    sc6540Driver2 = cc.CreateObject('SC6540.SC6540', interface=SC6540Lib.ISC6540)
    sc6540OptionString2 = 'Cache=false, InterchangeCheck=false, QueryInstrStatus=true, RangeCheck=false, RecordCoercions=false, Simulate=false'
    sc6540Driver2.Initialize(portNumSC2, True, False, sc6540OptionString2)
except Exception as e:
    print(f'Connection to SC6540 Switch2 failed: {e}')
    logger.error(f'Connection to SC6540 Switch2 failed: {e}')


def get_settings():
    config.read('settings.ini')
    global adminPassword
    try:
        adminPassword = config['Admin']['Password']
    except Exception as ex:  # Revert to default if password is missing in settings file
        logger.error(f'Error Getting Admin Password: {ex}')
        adminPassword = '6789'
        config['Admin']['Password'] = '6789'
    for x in range(1, 11):
        cav = 'cavity' + str(x)
        try:
            runCavity[cav] = tk.IntVar(value=int(config['Run Cavity'][cav]))
            laserEnabled[cav] = tk.IntVar(value=int(config['Laser Enabled'][cav]))
        except Exception as ex:
            logger.error(f'Error Getting settings file: {ex}')
            runCavity[cav] = tk.IntVar(value=1)
            laserEnabled[cav] = tk.IntVar(value=1)

    for key, value in config['Hypot'].items():  # Have to convert to their proper types so the AddACWTest can parse them properly
        if value == 'True':
            continuitySettings[key] = True
        elif value == 'False':
            continuitySettings[key] = False
        else:
            try:
                continuitySettings[key] = int(value)
            except:
                continuitySettings[key] = float(value)

    for key, value in config['Continuity'].items():  # Have to convert to their proper types so the AddACWTest can parse them properly
        if value == 'True':
            continuitySettings[key] = True
        elif value == 'False':
            continuitySettings[key] = False
        else:
            try:
                continuitySettings[key] = int(value)
            except:
                continuitySettings[key] = float(value)

def save_settings():
    # Write the config object to a file
    with open('settings.ini', 'w') as configfile:
        if config['Admin']['Password']:
            global adminPassword
            adminPassword = config['Admin']['Password']
        for x in range(1, 11):
            cav = 'cavity' + str(x)
            config['Run Cavity'][cav] = str(runCavity[cav].get())
            config['Laser Enabled'][cav] = str(laserEnabled[cav].get())

        for key, value in hypotSettings.items():
            config['Hypot'][key] = str(value)
        for key, value in continuitySettings.items():
            config['Continuity'][key] = str(value)

        config.write(configfile)  # Close and save to settings file

    update_colors(canvas)


def fault():
    faultWindow = tk.Toplevel(root)
    faultWindow.geometry('900x450')
    faultWindow.title('Part Fault')
    faultWindow.attributes('-toolwindow', True)  # Disables bar at the top right: min, max, close button
    faultWindow.attributes('-topmost', True)  # Force it to be above all other program windows
    faultWindow.configure(bg=backgroundColor)
    faultWindow.lift()

    faultLabel = tk.Label(faultWindow, text='Cavity failed a test', font=helv, fg=textColor, bg=backgroundColor)
    faultLabel.grid(row=0, column=2, pady=5)

    faultResetButton = tk.Button(faultWindow, text='Reset', command=lambda: reset(closeWindow=True, window=faultWindow), bg='#000000', fg=textColor, relief='flat', width=7,
                                 height=2, font=helvmedium)
    faultResetButton.grid(row=13, column=2, padx=3, pady=3)

    continuityFaultList = {}
    for cavity, value in cavityContinuitySuccesses.items():
        if not value:  # If failed continuity test
            logger.info('Continuity fail on Cavity: ' + str(cavity))
            continuityFaultList[cavity] = tk.Label(faultWindow, text='Cavity ' + str(cavity), font=helvmedium, fg=textColor, bg=backgroundColor)
            continuityFaultList[cavity].grid(row=cavity + 2, column=0)
    if continuityFaultList:  # If there are any continuity faults, put header
        continuityFaultHeader = tk.Label(faultWindow, text='Continuity Issues', font=helvUnderline, fg=textColor, bg=backgroundColor)
        continuityFaultHeader.grid(row=1, column=0, columnspan=2, pady=5)

    hypotFaultList = {}
    for cavity, value in cavityHypotSuccesses.items():
        if not value:  # If failed continuity test
            logger.info('Hypot fail on Cavity: ' + str(cavity))
            hypotFaultList[cavity] = tk.Label(faultWindow, text='Cavity ' + str(cavity), font=helvmedium, fg=textColor, bg=backgroundColor)
            hypotFaultList[cavity].grid(row=cavity + 2, column=3)
    if hypotFaultList:  # If there are any hypot faults, put header
        continuityFaultHeader = tk.Label(faultWindow, text='Hypot Issues', font=helvUnderline, fg=textColor, bg=backgroundColor)
        continuityFaultHeader.grid(row=1, column=3, columnspan=2, pady=5)


def reset(closeWindow, window):
    global faultState
    faultState = False
    if closeWindow:
        window.destroy()
    sc6540Driver1.Execution.DisableAllChannels()
    sc6540Driver2.Execution.DisableAllChannels()
    # Set all Output Variables to 0
    for cavity in cavityContinuitySuccesses:
        cavityContinuitySuccesses[cavity] = 0
    for cavity in cavityHypotSuccesses:
        cavityHypotSuccesses[cavity] = 0
    time.sleep(1)  # Make sure double clicks dont accidentally start it again
    startButton["state"] = "normal"  # Re-enables start button


def start():
    startButton["state"] = "disabled"  # Disabled start button so its not running twice at the same time due to threading
    disabledCavs = 0
    for cavity, value in runCavity.items():
        if value.get() == 0:
            disabledCavs += 1
    totalProgressBar['value'] = (disabledCavs * 10)
    totalProgressPercentage.configure(text=str(int(totalProgressBar['value'])) + ' %')  # Updates displayed percentage. Conv to int to remove decimals
    for cavity, value in runCavity.items():
        cavitynum = ''.join([char for char in cavity if char.isdigit()])
        cavitynum = int(cavitynum)
        if value.get() == 1:
            print('Running Cavity: ' + str(cavitynum))
            logger.info('Running Cavity: ' + str(cavitynum))

            sc6540Driver1.Execution.DisableAllChannels()
            sc6540Driver2.Execution.DisableAllChannels()

            continuity_setup(cavitynum)
            hypot_execution(continuityTest=True, cavityNum=cavitynum)

            sc6540Driver1.Execution.DisableAllChannels()
            sc6540Driver2.Execution.DisableAllChannels()

            hypot_setup(cavitynum)
            hypot_execution(continuityTest=False, cavityNum=cavitynum)

            sc6540Driver1.Execution.DisableAllChannels()
            sc6540Driver2.Execution.DisableAllChannels()
            totalProgressBar.step(10)
            totalProgressPercentage.configure(text=str(int(totalProgressBar['value'])) + ' %')  # Updates displayed percentage. Conv to int to remove decimals
        if laserEnabled['cavity' + str(cavitynum)].get() == 1:
            print('Lasering Cavity: ' + str(cavitynum))
            logger.info('Lasering Cavity: ' + str(cavitynum))
            # laser(cavitynum)
        print('=================================')  # Separate cavities for testing readability
    if faultState:  # If any part has a problem, have operators acknowledge they took care of it before starting again
        fault()
    else:
        reset(closeWindow=False, window=False)


def startstart():  # This is to put the main loop on a separate thread so it can be emergency stopped
    mainThread = Thread(target=start)
    mainThread.start()


def stop():
    logger.error('Emergency Stop Used!')
    try:
        sc6540Driver1.Execution.DisableAllChannels()
        sc6540Driver2.Execution.DisableAllChannels()
        logger.error('Emergency Stop Done!')
        root.quit()
        sys.exit()
    except Exception as ex:
        logger.error(f"Error during emergency stop!: {ex}")
        sc6540Driver1.Execution.DisableAllChannels()
        sc6540Driver2.Execution.DisableAllChannels()
        logger.error('Emergency Stop Done!')
        root.quit()
        sys.exit()
    finally:
        # Ensure that the Tkinter main loop exits cleanly
        logger.error('Hit finally in emergency stop!')
        logger.error('Emergency Stop Done!')
        root.quit()
        sys.exit()


def on_stop_button_clicked():
    stopThread = Thread(target=stop)
    stopThread.start()


def continuity_setup(cavitynum):
    if (cavitynum <= 5):  # First sc6540 switch and hypot
        sc6540Driver = sc6540Driver1
    else:  # Second sc6540 switch and hypot
        sc6540Driver = sc6540Driver2
    # Enable Continuity (High) channels
    sc6540Driver.Execution.ConfigureContinuityChannels({8})
    # After the multiplexer was configured, the safety or ground bond tester could start output for ground bond test on those connections.
    time.sleep(1)

    # Enable Return (Low) channels
    if cavitynum == 4 or cavitynum == 9:
        rtnChannel = 10  # Module B channel 1
    elif cavitynum == 5 or cavitynum == 10:
        rtnChannel = 12  # Module B channel 3
    else:
        rtnChannel = 2 * cavitynum
    sc6540Driver.Execution.ConfigureReturnChannels({rtnChannel})
    # After the multiplexer was configured, the safety tester could start dual check on those connections.
    time.sleep(1)

    logger.info('Continuity Setup Done')
    print('Continuity Setup Done')


def hypot_setup(cavitynum):
    if (cavitynum <= 5):  # First sc6540 switch and hypot
        sc6540Driver = sc6540Driver1
    else:  # Second sc6540 switch and hypot
        sc6540Driver = sc6540Driver2

    # Withstand test (ACW, DCW)
    # Enable Withstand (High) channels
    if cavitynum == 4 or cavitynum == 9:
        highChannel = 9  # Module B channel 1
    elif cavitynum == 5 or cavitynum == 10:
        highChannel = 11  # Module B channel 3
    else:
        highChannel = 2 * cavitynum - 1
    sc6540Driver.Execution.ConfigureWithstandChannels({highChannel})

    # Enable Return (Low) channels
    if cavitynum == 4 or cavitynum == 9:
        rtnChannel = 10  # Module B channel 2
    elif cavitynum == 5 or cavitynum == 10:
        rtnChannel = 12  # Module B channel 4
    else:
        rtnChannel = 2 * cavitynum
    sc6540Driver.Execution.ConfigureReturnChannels({rtnChannel})
    # After the multiplexer was configured, the safety tester could start output for withstand test on those connections.
    time.sleep(1)

    logger.info('Hypot Setup Done')
    print('Hypot Setup Done')


def hypot_execution(continuityTest, cavityNum):
    if cavityNum <= 5:
        hypotDriver = hypotDriver1
    else:
        hypotDriver = hypotDriver2

    # Create file
    try:
        hypotDriver.Files.Create(1, 'IviTest')
    except Exception as ex:
        logger.error('Exception occured at Hypot execution: ' + str(ex))
        files = hypotDriver.Files.TotalFiles
        for h in range(1, files + 1):
            if 'IviTest' in hypotDriver.Files.QueryFileName(h):
                hypotDriver.Files.Delete(h)
                hypotDriver.Files.Create(h, 'IviTest')
                break
    # Hypot manual results read on page 83
    #   Add ACW test item by AddACWTest()
    if (continuityTest):
        hypotDriver.Steps.AddACWTest(*tuple(continuitySettings.values()))  # * is the unpacking operator to separate tuple into parameters
    else:
        hypotDriver.Steps.AddACWTest(*tuple(hypotSettings.values()))  # * is the unpacking operator to separate tuple into parameters

    hypotDriver.Files.Save()
    # Start test
    hypotDriver.Execution.Execute()
    # Output Results
    read_hypot(continuityTest=continuityTest, hypotDriver=hypotDriver, cavityNum=cavityNum)
    # Reset test and close connection
    hypotDriver.Execution.Abort()

    if (continuityTest):
        print('Continuity Execution Done')
        logger.info('Continuity Execution Done')
    else:
        print('Hypot Execution Done')
        logger.info('Hypot Execution Done')


def read_hypot(continuityTest, hypotDriver, cavityNum):
    global faultState
    lastOpcStatus = False
    while (True):
        output = hypotDriver.Execution.ReadTestDisplayRaw().split(',')  # Split into an array for data parsing
        print(output)
        logger.info('Raw Output: ' + hypotDriver.Execution.ReadTestDisplayRaw())

        hypotDriver.System.WriteString('*OPC?\n')
        opcStatus = hypotDriver.System.ReadString()
        if ('1' in opcStatus and lastOpcStatus):
            # Successes
            if output[2] == 'PASS':
                if continuityTest:
                    cavityContinuitySuccesses[cavityNum] = 1
                    print('Cavity ' + str(cavityNum) + ' passes continuity')
                else:
                    cavityHypotSuccesses[cavityNum] = 1
                    print('Cavity ' + str(cavityNum) + ' passes hypot')
            # Failures
            if 'Cont' in output[2]:  # Failed Continuity (Can show , so check if Cont in output
                faultState = True
                cavityContinuitySuccesses[cavityNum] = 0
                print('Cavity ' + str(cavityNum) + ' fails continuity')
                logger.info('Cavity ' + str(cavityNum) + ' fails continuity')
            if output[2] == 'Breakdown':  # Failed Hypot
                faultState = True
                cavityHypotSuccesses[cavityNum] = 0
                print('Cavity ' + str(cavityNum) + ' fails hypot')
                logger.info('Cavity ' + str(cavityNum) + ' fails hypot')
            break
        lastOpcStatus = '1' in opcStatus
        time.sleep(0.1)


def laser(cavityNum):
    if cavityHypotSuccesses[cavityNum] == 1 and cavityContinuitySuccesses[cavityNum] == 1:  # Only Laser if passes both tests
        msg = 'WX,PRG=' + str(cavityNum) + ',BLK=1, MarkingParameter=80,1500,70'
        laser.send(msg.encode('utf-8'))  # Converts command string to byte format

        read_laser()

        logger.info('Laser Done')
        print('Laser Done')
    else:
        logger.info('Skipping Laser')
        print('Skipping Laser')


def read_laser():
    response = laser.recv(1024)  # Listens for data, max amount of bytes specified in ()
    response = response.decode('utf-8')  # Converts bytes to string

    print(response)
    logger.info('Laser Output: ' + response)


def admin_panel():
    def toggle_cavity():
        for key, value in cavityCheckBoxes.items():
            value.toggle()

    def toggle_laser():
        for key, value in laserCheckBoxes.items():
            value.toggle()

    def quit_admin():
        adminWindow.destroy()
        update_colors(canvas)
        adminTextbox.delete(0, 'end')  # Clears Password

    if (adminTextbox.get() == adminPassword):
        logger.info('Admin Logged in')
        try:  # Prevent duplicate windows being opened
            # noinspection PyUnboundLocalVariable
            if adminWindow.winfo_exists():  # python will raise an exception there if variable doesn't exist
                adminWindow.after(1, lambda: adminWindow.focus_force())  # Refocuses window instead of creating a duplicate
                pass
            else:
                adminWindow = tk.Toplevel(root)
        except NameError:  # exception? we are now here.
            adminWindow = tk.Toplevel(root)
        else:  # no exception and no window? creating window.
            if not adminWindow.winfo_exists():
                adminWindow = tk.Toplevel(root)

        adminWindow.geometry('900x450')
        adminWindow.title('Admin Panel')
        adminWindow.attributes('-toolwindow', True)  # Disables bar at min, max, close button in top right
        adminWindow.attributes('-topmost', True)  # Force it to be above all other program windows
        adminWindow.configure(bg=backgroundColor)
        adminWindow.lift()

        cavityHeaderLabel = tk.Label(adminWindow, text='Enable/Disable Cavities', font=helv, fg=textColor, bg=backgroundColor)
        cavityHeaderLabel.grid(row=0, column=1, columnspan=2)
        laserHeaderLabel = tk.Label(adminWindow, text='Enable/Disable Laser', font=helv, fg=textColor, bg=backgroundColor)
        laserHeaderLabel.grid(row=0, column=5, columnspan=2)

        save_settingsButton = tk.Button(adminWindow, text='Save', command=save_settings, bg='#000000', fg=textColor, relief='flat', width=7, height=2, font=helvmedium)
        save_settingsButton.grid(row=9, column=4, padx=3, pady=3)

        closeAdminButton = tk.Button(adminWindow, text='Close', command=quit_admin, bg='#000000', fg=textColor, relief='flat', width=7, height=2, font=helvmedium)
        closeAdminButton.grid(row=10, column=4, padx=3, pady=3)

        toggleCavityButton = tk.Button(adminWindow, text='Toggle Cavity', command=toggle_cavity, bg='#000000', fg=textColor, relief='flat', width=12, height=2, font=helvmedium)
        toggleCavityButton.grid(row=7, column=2, padx=3, pady=3)
        toggleLaserButton = tk.Button(adminWindow, text='Toggle Laser', command=toggle_laser, bg='#000000', fg=textColor, relief='flat', width=12, height=2, font=helvmedium)
        toggleLaserButton.grid(row=7, column=5, padx=3, pady=3)

        resetButton = tk.Button(adminWindow, text='Reset', command=lambda: reset(closeWindow=False, window=adminWindow), bg='#000000', fg=textColor, relief='flat', width=7, height=2, font=helvmedium)
        resetButton.grid(row=10, column=6, padx=3, pady=3)

        cavityCheckBoxes = {}
        for x in range(1, 11):
            cavityCheckBoxes[x] = Checkbutton(adminWindow, text='Cavity' + str(x),
                                              variable=runCavity['cavity' + str(x)], onvalue=1, offvalue=0, fg='white', selectcolor='Black', bg=backgroundColor,
                                              font=helvmedium)
            if x < 6:
                cavityCheckBoxes[x].grid(row=x, column=1)
            else:
                cavityCheckBoxes[x].grid(row=x - 5, column=2)
        laserCheckBoxes = {}
        for x in range(1, 11):
            laserCheckBoxes[x] = Checkbutton(adminWindow, text='Laser' + str(x),
                                             variable=laserEnabled['cavity' + str(x)], onvalue=1, offvalue=0, fg='white', selectcolor='Black', bg=backgroundColor,
                                             font=helvmedium)
            if x < 6:
                laserCheckBoxes[x].grid(row=x, column=5)
            else:
                laserCheckBoxes[x].grid(row=x - 5, column=6)


# Function to create a grid of rectangles and store references
def create_rectangle_grid(rows, columns, rectWidth, rectHeight, padding, canv):
    cavNum = 1
    for col in range(columns):
        for row in range(rows):
            x1 = col * (rectWidth + padding)
            y1 = row * (rectHeight + padding)
            x2 = x1 + rectWidth
            y2 = y1 + rectHeight
            rectangles[cavNum] = canv.create_rectangle(x1, y1, x2, y2, fill='green')
            canv.create_text((x1 + rectWidth / 2, y1 + 18), text="Cav " + str(cavNum), fill=textColor, font='tkDefaeultFont 20')
            statusText[cavNum] = canv.create_text((x1 + rectWidth / 2, y1 + rectHeight / 2 + 15), text="", fill=textColor, font='tkDefaeultFont 14')
            cavNum += 1
    update_colors(canv)


# Function to change the color of a rectangle based on its cavity number
def change_rectangle_color(cavNum, color, canv):
    if cavNum in rectangles:
        canv.itemconfig(rectangles[cavNum], fill=color)
    else:
        print(f"No rectangle found at position {cavNum}")


def update_colors(canv):
    for x in range(1, 11):
        if laserEnabled['cavity' + str(x)].get() == 0 and runCavity['cavity' + str(x)].get() == 0:
            change_rectangle_color(cavNum=x, color=disabledColor, canv=canv)
            change_rectangle_text(x, text="Fully Disabled")
        elif runCavity['cavity' + str(x)].get() == 0:
            change_rectangle_color(cavNum=x, color=halfDisabledColor, canv=canv)
            change_rectangle_text(x, text="Hypot Disabled")
        elif laserEnabled['cavity' + str(x)].get() == 0:
            change_rectangle_color(cavNum=x, color=halfDisabledColor, canv=canv)
            change_rectangle_text(cavNum=x, text="Laser Disabled")
        else:
            change_rectangle_color(cavNum=x, color=enabledColor, canv=canv)
            change_rectangle_text(cavNum=x, text="Enabled")


# Function to change the status text of a rectangle based on its cavity number
def change_rectangle_text(cavNum, text):
    if cavNum in statusText:
        canvas.itemconfig(statusText[cavNum], text=text)
    else:
        print(f"No rectangle found at position {cavNum}")


# Setting values to make sure theyre populated when referenced, or if no settings file found initially
for y in range(1, 11):
    c = 'cavity' + str(y)  # Using cavity(y) instead of int so settings file is a bit more readable
    runCavity[c] = tk.IntVar(value=0)
    laserEnabled[c] = tk.IntVar(value=0)

# Get settings on program start
get_settings()

#       Starting UI Setup
# Fonts
helv = tkfont.Font(family='Helvetica', size=20, weight='bold')
helvUnderline = tkfont.Font(family='Helvetica', size=20, weight='bold', underline=True)
helvmedium = tkfont.Font(family='Helvetica', size=15, weight='bold')
helvsmall = tkfont.Font(family='Helvetica', size=10, weight='bold')

# UI Setup
startButton = tk.Button(root, text='START', command=startstart, bg='#000000', fg=textColor, relief='flat', width=35,
                        height=20, font=helv)
startButton.pack(side=LEFT, padx=20)
stopButton = tk.Button(root, text='Emergency STOP', command=on_stop_button_clicked, bg='#000000', fg=textColor, relief='flat', width=18,
                       height=3, font=helvmedium)
stopButton.place(x=850, y=875)
root.protocol("WM_DELETE_WINDOW", on_stop_button_clicked)  # Gracefully shuts down program if window closed

totalProgressText = tk.Label(root, text='Total Progress', fg=textColor, bg=textBackgroundColor, font=helv)
totalProgressText.pack(side=tk.TOP, pady=10)
totalProgressBar = ttk.Progressbar(root, length=360, maximum=100)
totalProgressBar.pack(side=tk.TOP, ipady=6, pady=7)
totalProgressPercentage = tk.Label(root, text='0%', fg='Black', bg='#E6E6E6', font=helvmedium)
totalProgressPercentage.place(in_=totalProgressBar, relx=0.5, rely=0.5, anchor=tk.CENTER)

mainFrame = ttk.Frame(root, padding=(5, 5, 5, 5))
mainFrame.pack(side=tk.TOP)
canvas = Canvas(mainFrame, width=650, height=700, bg=canvasColor, highlightthickness=5, highlightbackground=canvasColor)
canvas.pack(fill=tk.BOTH, expand=True)

# Create the grid of rectangles
create_rectangle_grid(rows=5, columns=2, rectWidth=300, rectHeight=100, padding=50, canv=canvas)

# Admin UI
adminLabel = tk.Label(root, text='Admin Settings', fg=textColor, bg=textBackgroundColor, font=helv)
adminLabel.place(x=1525, y=850)
adminText = tk.StringVar()
adminTextbox = ttk.Entry(root, show='*', width=25)
adminTextbox.place(x=1550, y=900)
adminSubmitButton = tk.Button(root, text='Submit', command=admin_panel, bg='#000000', fg=textColor, relief='flat',
                              width=9, height=2, font=helvsmall)
adminSubmitButton.place(x=1550, y=925)

try:
    root.mainloop()
except KeyboardInterrupt:
    on_stop_button_clicked()
    sys.exit()
