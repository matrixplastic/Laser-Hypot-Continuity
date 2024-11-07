import time
import tkinter

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
errors = []
print('Setting up Logging')
scriptDir = os.path.dirname(__file__)  # Absolute dir the script is in
logname = 'logs/LaserHypotCont.log'
absFilePath = os.path.join(scriptDir, logname)
logger = logging.getLogger('Rotating Log')
try:
    if not os.path.exists('logs/'):
        os.makedirs('logs/')
    # Rotating Logs 1 every day, keep for 30 days
    handler = TimedRotatingFileHandler(absFilePath, when='h', interval=8, backupCount=30)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
    # Set the suffix to include the date format and '.log'
    handler.suffix = "%Y-%m-%d_%H-%M-%S.log"
    logger.addHandler(handler)

except Exception as e:
    print(f'Error creating new Log file: {e}')
    errors.append(f'Error creating new Log file: {e}')

print('Setting up Settings File')
# Setup Settings File
config = configparser.ConfigParser()
config['Run Cavity'] = {}
config['Laser Enabled'] = {}
config['Admin'] = {}
config['Hypot'] = {}
config['Continuity'] = {}
config['Laser'] = {}
config['Hardware IDs'] = {}

def connect_hardware(serialToConnect):
    pass
print('Setting up Drivers')
# Driver Variables
cc.GetModule('SC6540.dll')
from comtypes.gen import SC6540Lib

cc.GetModule('ARI38XX_64.dll')
from comtypes.gen import ARI38XXLib

hypotSerial1 = "AQ03JGPEA"
hypotSerial2 = "A107A3OCA"
switchSerial1 = "B0007EEKA"
switchSerial2 = "B0007BEKA"


# Setup Laser Connectivity
laser = socket.socket(socket.AF_INET, socket.SOCK_STREAM)  # Creates socket
laserIP = '10.10.0.167'
try:
    laser.connect((laserIP, 50000))  # IP and Port number for laser
except Exception as e:
    print(f'Connection to Laser Marker failed: {e}')
    logger.error(f'Connection to Laser Marker failed: {e}')
    errors.append(f'Connection to Laser Marker failed')

# General Variables
adminPassword = '6789'  # Default password if not set in the settings file
faultState = False
cavityContinuitySuccesses = {}
cavityHypotSuccesses = {}
runCavity = {}
laserEnabled = {}
hypotSettings = {}
continuitySettings = {}
defaultHypotSettings = {
    'voltage': 1000,  # AC Voltage
    'currenthighlimit': 10,  # Current High Limit
    'currentlowlimit': 0.001,  # Current Low Limit
    'rampuptime': 1,  # Ramp up time in seconds
    'dwelltime': 5,  # RampDownTime in seconds
    'rampdowntime': 1,  # Dewll time in seconds
    'arcsenselevel': 1,  # ArcSense level
    'arcdetection': False,  # Arc detection
    'frequency': ARI38XXLib.ARI38XXFrequency60Hz,  # Frequency
    'continuitytest': False,  # Continuity test
    'highlimitresistance': 1.5,  # High limit of the continuity resistance
    'lowlimitresistance': 0.01,  # Low limit of the continuity resistance
    'resistanceoffset': 0.5  # Continuity resistance offset
}
defaultContinuitySettings = {
    'voltage': 1000,  # AC Voltage
    'currenthighlimit': 10,  # Current High Limit
    'currentlowlimit': 0.001,  # Current Low Limit
    'rampuptime': 0.1,  # Ramp up time in seconds
    'dwelltime': 0.1,  # RampDownTime in seconds
    'rampdowntime': 0.1,  # Dewll time in seconds
    'arcsenselevel': 1,  # ArcSense level
    'arcdetection': False,  # Arc detection
    'frequency': ARI38XXLib.ARI38XXFrequency60Hz,  # Frequency
    'continuitytest': True,  # Continuity test
    'highlimitresistance': 1.5,  # High limit of the continuity resistance
    'lowlimitresistance': 0.01,  # Low limit of the continuity resistance
    'resistanceoffset': 0.5  # Continuity resistance offset
}

# Admin Panel Settings Variables
continuityTkinterObjs = {}
continuityTkinterObjsLabel = {}
continuityArcDetectionBool = None
hypotTkinterObjs = {}
hypotTkinterObjsLabel = {}
hypotArcDetectionBool = None

# UI Variables
rectangles = {}
statusText = {}
root = tk.Tk()
root.geometry('1800x900')
root.title('Main')
root.state('zoomed')
backgroundColor = '#2A2E32'
canvasColor = '#3D434B'
enabledColor = '#26A671'
halfDisabledColor = '#F9A825'
disabledColor = '#DE0A02'
textBackgroundColor = '#2A2E32'
textColor = 'White'

root.configure(bg=backgroundColor)
# root.attributes('-toolwindow', True)  # Disables bar at min, max, close button in top right


def find_com_port_by_serial_number(targetSerialNumber):
    ports = serial.tools.list_ports.comports()
    for port in ports:
        # Print all device details for debugging purposes
        print(f"Device: {port.device}, Description: {port.description}, HWID: {port.hwid}")
        logger.info(f"Device: {port.device}, Description: {port.description}, HWID: {port.hwid}")
        if targetSerialNumber in port.hwid:
            logger.info('Serial number: ' + targetSerialNumber + ' Located at: ' + port.device)
            return port.device
    return None


def concat_port(comPort):
    try:
        return 'ASRL' + comPort.replace("COM", '') + '::INSTR'
    except Exception as ex:
        logger.error(f'Concat port error with {comPort}: {ex}')
        print(f'Concat port error with {comPort}: {ex}')

# Avoid using COM# because windows can mix it up

hypotComPort1 = find_com_port_by_serial_number(hypotSerial1)
portNumHy1 = concat_port(hypotComPort1)

hypotComPort2 = find_com_port_by_serial_number(hypotSerial2)
portNumHy2 = concat_port(hypotComPort2)

switchComPort1 = find_com_port_by_serial_number(switchSerial1)
portNumSC1 = concat_port(switchComPort1)

switchComPort2 = find_com_port_by_serial_number(switchSerial2)
portNumSC2 = concat_port(switchComPort2)

print(portNumHy1)
logger.info(f"Hypot1 Port: {portNumHy1}")
print(portNumHy2)
logger.info(f"Hypot2 Port: {portNumHy1}")
print(portNumSC1)
logger.info(f"Switch1 Port: {portNumHy1}")
print(portNumSC2)
logger.info(f"Switch2 Port: {portNumHy1}")

# Driver Setup
try:
    hypotDriver1 = cc.CreateObject('ARI38XX.ARI38XX', interface=ARI38XXLib.IARI38XX)
    hypotDriver1.Initialize(portNumHy1, True, False, 'DriverSetup=BaudRate=38400')
except Exception as e:
    print(f'Connection to Hypot1 failed: {e}')
    logger.error(f'Connection to Hypot1 failed: {e}')
    errors.append('Connection to Hypot1 failed')

try:
    pass
    hypotDriver2 = cc.CreateObject('ARI38XX.ARI38XX', interface=ARI38XXLib.IARI38XX)
    hypotDriver2.Initialize(portNumHy2, True, False, 'DriverSetup=BaudRate=38400')
except Exception as e:
    print(f'Connection to Hypot2 failed: {e}')
    logger.error(f'Connection to Hypot2 failed: {e}')
    errors.append('Connection to Hypot2 failed')

try:
    switchDriver1 = cc.CreateObject('SC6540.SC6540', interface=SC6540Lib.ISC6540)
    switchOptionString1 = 'Cache=false, InterchangeCheck=false, QueryInstrStatus=true, RangeCheck=false, RecordCoercions=false, Simulate=false'
    switchDriver1.Initialize(portNumSC1, True, False, switchOptionString1)
except Exception as e:
    print(f'Connection to SC6540 Switch1 failed: {e}')
    logger.error(f'Connection to SC6540 Switch1 failed: {e}')
    errors.append('Connection to SC6540 Switch1 failed')
try:
    switchDriver2 = cc.CreateObject('SC6540.SC6540', interface=SC6540Lib.ISC6540)
    switchOptionString2 = 'Cache=false, InterchangeCheck=false, QueryInstrStatus=true, RangeCheck=false, RecordCoercions=false, Simulate=false'
    switchDriver2.Initialize(portNumSC2, True, False, switchOptionString2)
except Exception as e:
    print(f'Connection to SC6540 Switch2 failed: {e}')
    logger.error(f'Connection to SC6540 Switch2 failed: {e}')
    errors.append('Connection to SC6540 Switch2 failed')


def get_settings():
    config.read('settings.ini')
    global errors
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
            logger.error(f"Error reading Settings.ini file, creating new enabled variables. {ex}")
            runCavity[cav] = tk.IntVar(value=1)
            laserEnabled[cav] = tk.IntVar(value=1)
    # Hypot
    for key, defaultValue in defaultHypotSettings.items():
        # Make sure values exist
        if key not in config['Hypot']:
            value = defaultValue
        else:
            value = config['Hypot'][key]

        # Convert to their proper types so the AddACWTest can parse them properly
        try:
            if key in ('arcdetection', 'continuitytest'):
                hypotSettings[key] = bool(value)
            elif key in ('voltage', 'currenthighlimit', 'currentlowlimit', 'arcsenselevel', 'frequency'):
                hypotSettings[key] = int(value)
            elif key in ('rampuptime', 'rampdowntime', 'dwelltime', 'highlimitresistance', 'lowlimitresistance', 'resistanceoffset'):
                hypotSettings[key] = float(value)
        except Exception as ex:
            logger.error(f"Error reading Hypot values from Settings.ini file! Delete it!: {ex}")
            print(f"Error reading Hypot values from Settings.ini file! Delete it!: {ex}")
            errors.append("Error reading Settings.ini! Delete it, restart Program!")
    global hypotSerial1
    hypotSerial1 = config['Hardware IDs']['hypot1']
    global hypotSerial2
    hypotSerial2 = config['Hardware IDs']['hypot2']
    global switchSerial1
    switchSerial1 = config['Hardware IDs']['switch1']
    global switchSerial2
    switchSerial2 = config['Hardware IDs']['switch2']

    # Continuity
    for key, defaultValue in defaultContinuitySettings.items():
        # Make sure values exist
        if key not in config['Continuity']:
            value = defaultValue
        else:
            value = config['Continuity'][key]

        # Convert to their proper types so the AddACWTest can parse them properly
        try:
            if key in ('arcdetection', 'continuitytest'):
                continuitySettings[key] = bool(value)
            elif key in ('voltage', 'currenthighlimit', 'currentlowlimit', 'arcsenselevel', 'frequency'):
                continuitySettings[key] = int(value)
            elif key in ('rampuptime', 'rampdowntime', 'dwelltime', 'highlimitresistance', 'lowlimitresistance', 'resistanceoffset'):
                continuitySettings[key] = float(value)
        except Exception as ex:
            logger.error(f"Error reading Hypot values from Settings.ini file! Delete it!: {ex}")
            print(f"Error reading Hypot values from Settings.ini file! Delete it!: {ex}")
            errors.append("Error reading Settings.ini file! Delete it, and restart program.")

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

        for key, value in hypotTkinterObjs.items():
            config['Hypot'][key] = str(value.get())
        config['Hypot']['arcdetection'] = str(hypotArcDetectionBool.get())

        for key, value in continuityTkinterObjs.items():
            config['Continuity'][key] = str(value.get())
        config['Continuity']['arcdetection'] = str(continuityArcDetectionBool.get())


        config.write(configfile)  # Close and save to settings file
    update_colors(canvas)


def save_serials():
    if hypotSerial1 != config['Hardware IDs']['hypot1']:
        try:
            config['Hardware IDs']['hypot1'] = hypotSerial1
        except Exception as ex:
            logger.error(f"Error Connecting to or Saving Hypot 1 to settings file! {ex}")
    if hypotSerial2 != config['Hardware IDs']['hypot2']:
        try:
            config['Hardware IDs']['hypot2'] = hypotSerial2
        except Exception as ex:
            logger.error(f"Error Connecting to or Saving Hypot 2 to settings file! {ex}")
    if switchSerial1 != config['Hardware IDs']['switch1']:
        try:
            config['Hardware IDs']['switch1'] = switchSerial1
        except Exception as ex:
            logger.error(f"Error Connecting to or Saving Switch 1 to settings file! {ex}")
    if switchSerial2 != config['Hardware IDs']['switch2']:
        try:
            config['Hardware IDs']['switch2'] = switchSerial2
        except Exception as ex:
            logger.error(f"Error Connecting to or Saving Switch 2 to settings file! {ex}")


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
    switchDriver1.Execution.DisableAllChannels()
    switchDriver2.Execution.DisableAllChannels()
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

            switchDriver1.Execution.DisableAllChannels()
            switchDriver2.Execution.DisableAllChannels()

            continuity_setup(cavitynum)
            hypot_execution(continuityTest=True, cavityNum=cavitynum)

            switchDriver1.Execution.DisableAllChannels()
            switchDriver2.Execution.DisableAllChannels()

            hypot_setup(cavitynum)
            hypot_execution(continuityTest=False, cavityNum=cavitynum)

            switchDriver1.Execution.DisableAllChannels()
            switchDriver2.Execution.DisableAllChannels()
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
        switchDriver1.Execution.DisableAllChannels()
        switchDriver2.Execution.DisableAllChannels()
        logger.error('Emergency Stop Done!')
        root.quit()
        root.destroy()
        sys.exit()
    except Exception as ex:
        logger.error(f"Error during emergency stop!: {ex}")
        switchDriver1.Execution.DisableAllChannels()
        switchDriver2.Execution.DisableAllChannels()
        logger.error('Emergency Stop Done!')
        root.quit()
        root.destroy()
        sys.exit()
    finally:
        # Ensure that the Tkinter main loop exits cleanly
        logger.error('Hit finally in emergency stop!')
        logger.error('Emergency Stop Done!')
        root.quit()
        root.destroy()
        sys.exit()


def on_stop_button_clicked():
    stopThread = Thread(target=stop)
    stopThread.start()


def continuity_setup(cavitynum):
    if (cavitynum <= 5):  # First sc6540 switch and hypot
        switchDriver = switchDriver1
    else:  # Second sc6540 switch and hypot
        switchDriver = switchDriver2
        cavitynum -= 5 # Reduce value for proper switch port assignments

    # After the multiplexer was configured, the safety or ground bond tester could start output for ground bond test on those connections.
    time.sleep(1)

    # Enable Return (Low) channels
    rtnChannel = 2 * cavitynum - 1

    # Enable Continuity (High) channels
    switchDriver.Execution.ConfigureContinuityChannels({16})
    switchDriver.Execution.ConfigureReturnChannels({rtnChannel})
    # After the multiplexer was configured, the safety tester could start dual check on those connections.
    time.sleep(1)

    logger.info('Continuity Setup Done')
    print('Continuity Setup Done')


def hypot_setup(cavitynum):
    if (cavitynum <= 5):  # First sc6540 switch and hypot
        switchDriver = switchDriver1
    else:  # Second sc6540 switch and hypot
        switchDriver = switchDriver2
        cavitynum -= 5 # Reduce value for proper switch port assignments

    # Enable Return (Low) channels
    rtnChannel = 2 * cavitynum - 1
    highChannel = 2 * cavitynum

    switchDriver.Execution.ConfigureWithstandChannels({highChannel})
    switchDriver.Execution.ConfigureReturnChannels({rtnChannel})

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
        logger.error('Missing Test File creating a new one, or Issue: ' + str(ex))
        files = hypotDriver.Files.TotalFiles
        for h in range(1, files + 1):
            if 'IviTest' in hypotDriver.Files.QueryFileName(h):
                hypotDriver.Files.Delete(h)
                hypotDriver.Files.Create(h, 'IviTest')
                break
    # Hypot manual results read on page 83
    #   Add ACW test item by AddACWTest()
    try:
        if (continuityTest):
            # Add an ACW step and then edit parameters
            hypotDriver.Steps.AddACWTestWithDefaults()
            hypotDriver.Parameters.Voltage = continuitySettings['voltage']
            hypotDriver.Parameters.HighLimit = continuitySettings['currenthighlimit']
            hypotDriver.Parameters.LowLimit = continuitySettings['currentlowlimit']
            hypotDriver.Parameters.RampUp = continuitySettings['rampuptime']
            hypotDriver.Parameters.Dwell = continuitySettings['dwelltime']
            hypotDriver.Parameters.RampDown = continuitySettings['rampdowntime']
            hypotDriver.Parameters.ArcSense = continuitySettings['arcsenselevel']
            hypotDriver.Parameters.ArcDetectEnabled = continuitySettings['arcdetection']
            hypotDriver.Parameters.Frequency = continuitySettings['frequency']
            hypotDriver.Parameters.ContinuityEnabled = continuitySettings['continuitytest']
            hypotDriver.Parameters.ContHiLimit = continuitySettings['highlimitresistance']
            hypotDriver.Parameters.ContLoLimit = continuitySettings['lowlimitresistance']
            hypotDriver.Parameters.ContOffset = continuitySettings['resistanceoffset']
        else:
            hypotDriver.Steps.AddACWTestWithDefaults()
            hypotDriver.Parameters.Voltage = hypotSettings['voltage']
            hypotDriver.Parameters.HighLimit = hypotSettings['currenthighlimit']
            hypotDriver.Parameters.LowLimit = hypotSettings['currentlowlimit']
            hypotDriver.Parameters.RampUp = hypotSettings['rampuptime']
            hypotDriver.Parameters.Dwell = hypotSettings['dwelltime']
            hypotDriver.Parameters.RampDown = hypotSettings['rampdowntime']
            hypotDriver.Parameters.ArcSense = hypotSettings['arcsenselevel']
            hypotDriver.Parameters.ArcDetectEnabled = hypotSettings['arcdetection']
            hypotDriver.Parameters.Frequency = hypotSettings['frequency']
            hypotDriver.Parameters.ContinuityEnabled = hypotSettings['continuitytest']
            hypotDriver.Parameters.ContHiLimit = hypotSettings['highlimitresistance']
            hypotDriver.Parameters.ContLoLimit = hypotSettings['lowlimitresistance']
            hypotDriver.Parameters.ContOffset = hypotSettings['resistanceoffset']

        hypotDriver.Files.Save()
        # Start test
        hypotDriver.Execution.Execute()
        # Output Results
        read_hypot(continuityTest=continuityTest, hypotDriver=hypotDriver, cavityNum=cavityNum)
        # Reset test and close connection
    except Exception as ex:
        logger.error('Exception occured at Hypot execution: ' + str(ex))
        errors.append('Exception occured at Hypot execution: ' + str(ex))
        print('Exception occured at Hypot execution: ' + str(ex))
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
    get_settings()
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
        except (NameError, tk.TclError):  # exception? we are now here.
            adminWindow = tk.Toplevel(root)
        else:  # no exception and no window? creating window.
            if not adminWindow.winfo_exists():
                adminWindow = tk.Toplevel(root)

        adminWindow.geometry('1600x800')
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
            cavityCheckBoxes[x] = Checkbutton(adminWindow, text='Cavity' + str(x), variable=runCavity['cavity' + str(x)], onvalue=1, offvalue=0, fg='white', selectcolor='Black', bg=backgroundColor, font=helvmedium)
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

        # Breaks up view a bit, improves legibility
        verticalSeparator = Frame(adminWindow, bg="red", height=800, width=2)
        verticalSeparator.place(x=710, y=0)

        # Continuity and Hypot Settings

        continuityTkinterObjsLabel['header'] = tk.Label(adminWindow, text='Continuity Settings', font=helv, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['header'].grid(row=0, column=9, columnspan=2, padx=30)

        continuityTkinterObjsLabel['voltage'] = tk.Label(adminWindow, text='Voltage', font=helvsmall, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['voltage'].grid(row=1, column=9, padx=20)
        continuityTkinterObjs['voltage'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=5000, increment=10)
        continuityTkinterObjs['voltage'].set(continuitySettings['voltage'])
        continuityTkinterObjs['voltage'].grid(row=1, column=10, padx=20)

        continuityTkinterObjsLabel['currenthighlimit'] = tk.Label(adminWindow, text='Current High Limit', font=helvsmall, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['currenthighlimit'].grid(row=2, column=9, padx=20)
        continuityTkinterObjs['currenthighlimit'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=20, increment=1)
        continuityTkinterObjs['currenthighlimit'].set(continuitySettings['currenthighlimit'])
        continuityTkinterObjs['currenthighlimit'].grid(row=2, column=10, padx=20)

        continuityTkinterObjsLabel['currentlowlimit'] = tk.Label(adminWindow, text='Current Low Limit', font=helvsmall, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['currentlowlimit'].grid(row=3, column=9, padx=20)
        continuityTkinterObjs['currentlowlimit'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=9, increment=1)
        continuityTkinterObjs['currentlowlimit'].set(continuitySettings['currentlowlimit'])
        continuityTkinterObjs['currentlowlimit'].grid(row=3, column=10, padx=20)

        continuityTkinterObjsLabel['rampuptime'] = tk.Label(adminWindow, text='Ramp Up Time', font=helvsmall, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['rampuptime'].grid(row=4, column=9, padx=20)
        continuityTkinterObjs['rampuptime'] = ttk.Spinbox(adminWindow, width=10, from_=0.1, to=999, increment=0.1)
        continuityTkinterObjs['rampuptime'].set(continuitySettings['rampuptime'])
        continuityTkinterObjs['rampuptime'].grid(row=4, column=10, padx=20)

        continuityTkinterObjsLabel['rampdowntime'] = tk.Label(adminWindow, text='Ramp Down Time', font=helvsmall, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['rampdowntime'].grid(row=5, column=9, padx=20)
        continuityTkinterObjs['rampdowntime'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=999, increment=0.1)
        continuityTkinterObjs['rampdowntime'].set(continuitySettings['rampdowntime'])
        continuityTkinterObjs['rampdowntime'].grid(row=5, column=10, padx=20)

        continuityTkinterObjsLabel['dwelltime'] = tk.Label(adminWindow, text='Dwell Time', font=helvsmall, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['dwelltime'].grid(row=6, column=9, padx=20)
        continuityTkinterObjs['dwelltime'] = ttk.Spinbox(adminWindow, width=10, from_=0.3, to=999, increment=0.1)
        continuityTkinterObjs['dwelltime'].set(continuitySettings['dwelltime'])
        continuityTkinterObjs['dwelltime'].grid(row=6, column=10, padx=20)

        continuityTkinterObjsLabel['arcsenselevel'] = tk.Label(adminWindow, text='Arc Sense', font=helvsmall, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['arcsenselevel'].grid(row=7, column=9, padx=20)
        continuityTkinterObjs['arcsenselevel'] = ttk.Spinbox(adminWindow, width=10, from_=1, to=9, increment=1)
        continuityTkinterObjs['arcsenselevel'].set(continuitySettings['arcsenselevel'])
        continuityTkinterObjs['arcsenselevel'].grid(row=7, column=10, padx=20)

        def print_value():  # Have to have command parameter in radio buttons or default value doesn't select properly. Possible tkinter bug
            print("Arc Detection:", continuityArcDetectionBool.get())

        global continuityArcDetectionBool
        continuityArcDetectionBool = tk.BooleanVar(value=continuitySettings['arcdetection'])

        continuityTkinterObjsLabel['arcdetection'] = tk.Label(adminWindow, text='Arc Detection', font=helvsmall, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['arcdetection'].grid(row=8, column=9, padx=20)
        continuityArcDetectionRadioTrue = ttk.Radiobutton(adminWindow, text='Yes', value=True, variable=continuityArcDetectionBool, command=print_value)
        continuityArcDetectionRadioTrue.grid(row=8, column=10, padx=20)
        continuityArcDetectionRadioFalse = ttk.Radiobutton(adminWindow, text='No', value=False, variable=continuityArcDetectionBool, command=print_value)
        continuityArcDetectionRadioFalse.grid(row=8, column=11, padx=20)

        continuityTkinterObjsLabel['highlimitresistance'] = tk.Label(adminWindow, text='High Limit Resistance', font=helvsmall, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['highlimitresistance'].grid(row=9, column=9, padx=20)
        continuityTkinterObjs['highlimitresistance'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=1.5, increment=0.01)
        continuityTkinterObjs['highlimitresistance'].set(continuitySettings['highlimitresistance'])
        continuityTkinterObjs['highlimitresistance'].grid(row=9, column=10, padx=20)

        continuityTkinterObjsLabel['lowlimitresistance'] = tk.Label(adminWindow, text='Low Limit Resistance', font=helvsmall, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['lowlimitresistance'].grid(row=10, column=9, padx=20)
        continuityTkinterObjs['lowlimitresistance'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=1.5, increment=0.01)
        continuityTkinterObjs['lowlimitresistance'].set(continuitySettings['lowlimitresistance'])
        continuityTkinterObjs['lowlimitresistance'].grid(row=10, column=10, padx=20)

        continuityTkinterObjsLabel['resistanceoffset'] = tk.Label(adminWindow, text='Resistance Offset', font=helvsmall, fg=textColor, bg=backgroundColor)
        continuityTkinterObjsLabel['resistanceoffset'].grid(row=11, column=9, padx=20)
        continuityTkinterObjs['resistanceoffset'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=0.5, increment=0.01)
        continuityTkinterObjs['resistanceoffset'].set(continuitySettings['resistanceoffset'])
        continuityTkinterObjs['resistanceoffset'].grid(row=11, column=10, padx=20)

        # Hypot Settings

        hypotTkinterObjsLabel['header'] = tk.Label(adminWindow, text='Hypot Settings', font=helv, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['header'].grid(row=0, column=13, columnspan=2, padx=20)

        hypotTkinterObjsLabel['voltage'] = tk.Label(adminWindow, text='Voltage', font=helvsmall, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['voltage'].grid(row=1, column=12, padx=20)
        hypotTkinterObjs['voltage'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=5000, increment=10)
        hypotTkinterObjs['voltage'].set(hypotSettings['voltage'])
        hypotTkinterObjs['voltage'].grid(row=1, column=13, padx=20)

        hypotTkinterObjsLabel['currenthighlimit'] = tk.Label(adminWindow, text='Current High Limit', font=helvsmall, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['currenthighlimit'].grid(row=2, column=12, padx=20)
        hypotTkinterObjs['currenthighlimit'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=20, increment=1)
        hypotTkinterObjs['currenthighlimit'].set(hypotSettings['currenthighlimit'])
        hypotTkinterObjs['currenthighlimit'].grid(row=2, column=13, padx=20)

        hypotTkinterObjsLabel['currentlowlimit'] = tk.Label(adminWindow, text='Current Low Limit', font=helvsmall, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['currentlowlimit'].grid(row=3, column=12, padx=20)
        hypotTkinterObjs['currentlowlimit'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=9.9, increment=0.1)
        hypotTkinterObjs['currentlowlimit'].set(hypotSettings['currentlowlimit'])
        hypotTkinterObjs['currentlowlimit'].grid(row=3, column=13, padx=20)

        hypotTkinterObjsLabel['rampuptime'] = tk.Label(adminWindow, text='Ramp Up Time', font=helvsmall, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['rampuptime'].grid(row=4, column=12, padx=20)
        hypotTkinterObjs['rampuptime'] = ttk.Spinbox(adminWindow, width=10, from_=0.1, to=999, increment=0.1)
        hypotTkinterObjs['rampuptime'].set(hypotSettings['rampuptime'])
        hypotTkinterObjs['rampuptime'].grid(row=4, column=13, padx=20)

        hypotTkinterObjsLabel['rampdowntime'] = tk.Label(adminWindow, text='Ramp Down Time', font=helvsmall, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['rampdowntime'].grid(row=5, column=12, padx=20)
        hypotTkinterObjs['rampdowntime'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=999, increment=0.1)
        hypotTkinterObjs['rampdowntime'].set(hypotSettings['rampdowntime'])
        hypotTkinterObjs['rampdowntime'].grid(row=5, column=13, padx=20)

        hypotTkinterObjsLabel['dwelltime'] = tk.Label(adminWindow, text='Dwell Time', font=helvsmall, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['dwelltime'].grid(row=6, column=12, padx=20)
        hypotTkinterObjs['dwelltime'] = ttk.Spinbox(adminWindow, width=10, from_=0.3, to=999, increment=0.1)
        hypotTkinterObjs['dwelltime'].set(hypotSettings['dwelltime'])
        hypotTkinterObjs['dwelltime'].grid(row=6, column=13, padx=20)

        hypotTkinterObjsLabel['arcsenselevel'] = tk.Label(adminWindow, text='Arc Sense', font=helvsmall, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['arcsenselevel'].grid(row=7, column=12, padx=20)
        hypotTkinterObjs['arcsenselevel'] = ttk.Spinbox(adminWindow, width=10, from_=1, to=9, increment=1)
        hypotTkinterObjs['arcsenselevel'].set(hypotSettings['arcsenselevel'])
        hypotTkinterObjs['arcsenselevel'].grid(row=7, column=13, padx=20)

        def print_value():  # Have to have command parameter in radio buttons or default value doesn't select properly. Possible tkinter bug
            print("Arc Detection:", hypotArcDetectionBool.get())

        global hypotArcDetectionBool
        hypotArcDetectionBool = tk.BooleanVar(value=hypotSettings['arcdetection'])

        hypotTkinterObjsLabel['arcdetection'] = tk.Label(adminWindow, text='Arc Detection', font=helvsmall, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['arcdetection'].grid(row=8, column=12, padx=20)
        hypotArcDetectionRadioTrue = ttk.Radiobutton(adminWindow, text='Yes', value=True, variable=hypotArcDetectionBool, command=print_value)
        hypotArcDetectionRadioTrue.grid(row=8, column=13, padx=20)
        hypotArcDetectionRadioFalse = ttk.Radiobutton(adminWindow, text='No', value=False, variable=hypotArcDetectionBool, command=print_value)
        hypotArcDetectionRadioFalse.grid(row=8, column=14, padx=20)

        hypotTkinterObjsLabel['highlimitresistance'] = tk.Label(adminWindow, text='High Limit Resistance', font=helvsmall, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['highlimitresistance'].grid(row=9, column=12, padx=20)
        hypotTkinterObjs['highlimitresistance'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=1.5, increment=0.01)
        hypotTkinterObjs['highlimitresistance'].set(hypotSettings['highlimitresistance'])
        hypotTkinterObjs['highlimitresistance'].grid(row=9, column=13, padx=20)

        hypotTkinterObjsLabel['lowlimitresistance'] = tk.Label(adminWindow, text='Low Limit Resistance', font=helvsmall, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['lowlimitresistance'].grid(row=10, column=12, padx=20)
        hypotTkinterObjs['lowlimitresistance'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=1.5, increment=0.01)
        hypotTkinterObjs['lowlimitresistance'].set(hypotSettings['lowlimitresistance'])
        hypotTkinterObjs['lowlimitresistance'].grid(row=10, column=13, padx=20)

        hypotTkinterObjsLabel['resistanceoffset'] = tk.Label(adminWindow, text='Resistance Offset', font=helvsmall, fg=textColor, bg=backgroundColor)
        hypotTkinterObjsLabel['resistanceoffset'].grid(row=11, column=12, padx=20)
        hypotTkinterObjs['resistanceoffset'] = ttk.Spinbox(adminWindow, width=10, from_=0, to=0.5, increment=0.01)
        hypotTkinterObjs['resistanceoffset'].set(hypotSettings['resistanceoffset'])
        hypotTkinterObjs['resistanceoffset'].grid(row=11, column=13, padx=20)

    else: # Wrong password
        wrongPassLabel = tk.Label(root, text="Wrong Password!")
        wrongPassLabel.place(x=1550, y=925)
        root.after(3000, lambda: wrongPassLabel.destroy())  # time in ms


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
            update_rectangle_text(x, text="Fully Disabled")
        elif runCavity['cavity' + str(x)].get() == 0:
            change_rectangle_color(cavNum=x, color=halfDisabledColor, canv=canv)
            update_rectangle_text(x, text="Hypot Disabled")
        elif laserEnabled['cavity' + str(x)].get() == 0:
            change_rectangle_color(cavNum=x, color=halfDisabledColor, canv=canv)
            update_rectangle_text(cavNum=x, text="Laser Disabled")
        else:
            change_rectangle_color(cavNum=x, color=enabledColor, canv=canv)
            update_rectangle_text(cavNum=x, text="Enabled")


# Function to change the status text of a rectangle based on its cavity number
def update_rectangle_text(cavNum, text):
    if cavNum in statusText:
        canvas.itemconfig(statusText[cavNum], text=text)
    else:
        print(f"No rectangle found at position {cavNum}")

def update_error_text():
    if not errors:
        errorString = 'All Hardware Connected'
        errorText.config(text=errorString, fg='green')
    else:
        errorString = '\n'.join(errors)
        errorText.config(text=errorString, fg='red')

# Setting values to make sure theyre populated when referenced, or if no settings file found initially
for y in range(1, 11):
    c = 'cavity' + str(y)  # Using cavity(y) instead of int so settings file is a bit more readable
    runCavity[c] = tk.IntVar(value=0)
    laserEnabled[c] = tk.IntVar(value=0)

# Get settings on program start
get_settings()

#       Starting UI Setup
# Fonts and Styles
helv = tkfont.Font(family='Helvetica', size=20, weight='bold')
helvUnderline = tkfont.Font(family='Helvetica', size=20, weight='bold', underline=True)
helvmedium = tkfont.Font(family='Helvetica', size=15, weight='bold')
helvsmall = tkfont.Font(family='Helvetica', size=10, weight='bold')

# UI Setup
startButton = tk.Button(root, text='START', command=startstart, bg='#000000', fg=textColor, relief='flat', width=20, height=12, font=helv)
startButton.place(x=50, y=400)
stopButton = tk.Button(root, text='Emergency STOP', command=on_stop_button_clicked, bg='#000000', fg=textColor, relief='flat', width=18, height=3, font=helvmedium)
stopButton.place(x=850, y=875)
root.protocol("WM_DELETE_WINDOW", on_stop_button_clicked)  # Gracefully shuts down program if window closed

progressCanvas = Canvas(root, width=400, height=300, bg=canvasColor, highlightthickness=5, highlightbackground=canvasColor)
progressCanvas.place(x=1000, y=10)
progressCanvas.create_rectangle(0, 0, 0, 0, fill='white')

totalProgressText = tk.Label(progressCanvas, text='Total Progress', fg=textColor, bg=canvasColor, font=helv)
totalProgressText.pack(side=TOP)
totalProgressBar = ttk.Progressbar(progressCanvas, length=360, maximum=100)
totalProgressBar.pack(side=TOP)
totalProgressPercentage = tk.Label(progressCanvas, text='0%', fg='Black', bg='#E6E6E6', font=helvmedium)
totalProgressPercentage.place(in_=totalProgressBar, relx=0.5, rely=0.5, anchor=tk.CENTER)

rectangleFrame = ttk.Frame(root, padding=(5, 5, 5, 5), width=670, height=720)
rectangleFrame.place(x=1000, y=110)
canvas = Canvas(rectangleFrame, width=650, height=700, bg=canvasColor, highlightthickness=5, highlightbackground=canvasColor)
canvas.place(x=0, y=0)

errorFrame = ttk.Frame(root, padding=(5, 5, 5, 5), width=520, height=270)
errorFrame.place(x=100, y=50)
errorCanvas = Canvas(errorFrame, width=500, height=250, bg=canvasColor, highlightthickness=5, highlightbackground=canvasColor)
errorCanvas.place(x=0, y=0)
errorText = tk.Label(errorCanvas, text='', fg='red', bg=canvasColor, font=helvmedium)
errorText.place(x=0, y=0)
update_error_text()


# Create the grid of rectangles
create_rectangle_grid(rows=5, columns=2, rectWidth=300, rectHeight=100, padding=50, canv=canvas)

# Admin UI
adminLabel = tk.Label(root, text='Admin Settings', fg=textColor, bg=textBackgroundColor, font=helv)
adminLabel.place(x=1525, y=850)
adminText = tk.StringVar()
adminTextbox = ttk.Entry(root, show='*', width=25)
adminTextbox.place(x=1550, y=900)
adminSubmitButton = tk.Button(root, text='Submit', command=admin_panel, bg='#000000', fg=textColor, relief='flat', width=9, height=2, font=helvsmall)
adminSubmitButton.place(x=1550, y=925)

try:
    root.mainloop()
except KeyboardInterrupt:
    on_stop_button_clicked()
    sys.exit()
