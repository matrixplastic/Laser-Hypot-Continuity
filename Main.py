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

from machinedata import MachineData

VERSION = '2026.09.30'


# Setup Logging
errors = []
print('Setting up Logging')
def resource_path(relative_path: str) -> str:
    """ Get absolute path to resource, works for dev and PyInstaller """
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

# Prepare log directory and full path to log file
log_dir: str = resource_path('logs')
os.makedirs(log_dir, exist_ok=True)

log_path: str = os.path.join(log_dir, 'LaserHypotCont.log')

# Setup logger
logger = logging.getLogger('Rotating Log')
handler = TimedRotatingFileHandler(filename=log_path,when='h',interval=8,backupCount=30)
handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
handler.suffix = "%Y-%m-%d_%H-%M-%S.log"
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)


print('Setting up Settings File')
# Setup Settings File
SETTINGS_PATH = resource_path('settings.ini')  # next to the exe; a shortcut's Start-in folder must not pick the file
config = configparser.ConfigParser()
config['Run Cavity'] = {}
config['Laser Enabled'] = {}
config['Admin'] = {}
config['Hypot'] = {}
config['Laser'] = {}
config['Hardware IDs'] = {}
config['MachineData'] = {}


print('Setting up Drivers')
# Driver Variables
cc.GetModule('SC6540.dll')
from comtypes.gen import SC6540Lib

cc.GetModule('ARI38XX_64.dll')
from comtypes.gen import ARI38XXLib


hypotHwid1 = "BG01PD8AA"
hypotHwid2 = "BG01PFCQA"
switchHwid1 = "B0007EEKA"
switchHwid2 = "B0007BEKA"


# Setup Laser Connectivity
laserIP = '192.180.0.11'  # default; settings.ini [Laser] ip overrides
laserSocket = None
laserConnected = False


def connect_laser():
    global laserSocket, laserConnected
    try:
        laserSocket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        laserSocket.settimeout(10)
        laserSocket.connect((laserIP, 50000))  # IP and Port number for laser
        laserConnected = True
        if 'Connection to Laser Marker failed' in errors:
            errors.remove('Connection to Laser Marker failed')
    except Exception as e:
        laserConnected = False
        print(f'Connection to Laser Marker failed: {e}')
        logger.error(f'Connection to Laser Marker failed: {e}')
        if 'Connection to Laser Marker failed' not in errors:
            errors.append('Connection to Laser Marker failed')
    return laserConnected

# General Variables
adminPassword = '6789'  # Default password if not set in the settings file
faultState = False
cavityContinuitySuccesses = {} # 0=Failure, 1=Success, 2=SkippedIfContFail
cavityHypotSuccesses = {}
cavityLaserSuccesses = {}  # 0=Failure, 1=Success, 2=SkippedTestFailed, 3=Disabled
runCavity = {}
laserEnabled = {}
hypotSettings = {}
settingsValid = True
usbHwids = set()
# Test parameters for the single ACW step (continuity runs first in the same step). These are
# the controlled process values: they must match the LHC work instruction, and the WI revision
# they came from belongs in the comment below whenever they change. settings.ini overrides
# them on the machine; every run logs the values it used and reads them back from the tester.
# WI reference: to be filled in against the controlled LHC work instruction.
defaultHypotSettings = {
    'voltage': 1000,  # AC voltage
    'currenthighlimit': 10,  # Current high limit (mA)
    'currentlowlimit': 0,  # Current low limit (mA)
    'rampuptime': 0.1,  # Ramp up time in seconds
    'dwelltime': 0.3,  # Dwell time in seconds
    'rampdowntime': 0.0,  # Ramp down time in seconds
    'arcsenselevel': 1,  # ArcSense level
    'arcdetection': True,  # Arc detection
    'frequency': ARI38XXLib.ARI38XXFrequency60Hz,  # Frequency
    'continuitytest': True,  # Continuity test runs in the same step, before the withstand test
    'highlimitresistance': 1.5,  # High limit of the continuity resistance (ohm)
    'lowlimitresistance': 0.01,  # Low limit of the continuity resistance (ohm)
    'resistanceoffset': 0.5  # Continuity resistance offset (ohm)
}
# settings key, driver property. Used to load the tester and to read it back.
HYPOT_PARAMS = (
    ('voltage', 'Voltage'), ('currenthighlimit', 'HighLimit'), ('currentlowlimit', 'LowLimit'),
    ('rampuptime', 'RampUp'), ('dwelltime', 'Dwell'), ('rampdowntime', 'RampDown'),
    ('arcsenselevel', 'ArcSense'), ('arcdetection', 'ArcDetectEnabled'), ('frequency', 'Frequency'),
    ('continuitytest', 'ContinuityEnabled'), ('highlimitresistance', 'ContHiLimit'),
    ('lowlimitresistance', 'ContLoLimit'), ('resistanceoffset', 'ContOffset'),
)

SETTING_LABELS = {
    'voltage': 'Voltage (V)', 'currenthighlimit': 'Current high limit (mA)', 'currentlowlimit': 'Current low limit (mA)',
    'rampuptime': 'Ramp up (s)', 'dwelltime': 'Dwell (s)', 'rampdowntime': 'Ramp down (s)', 'arcsenselevel': 'Arc sense',
    'arcdetection': 'Arc detection', 'frequency': 'Frequency', 'continuitytest': 'Continuity test',
    'highlimitresistance': 'Continuity high limit (\u03a9)', 'lowlimitresistance': 'Continuity low limit (\u03a9)', 'resistanceoffset': 'Resistance offset (\u03a9)',
}
settingsPanelRows = {}   # setting key -> value label on the main window
devicePanelRows = {}     # role -> (dot canvas, oval id, detail label) on the main window

# Admin Panel Settings Variables
hypotTkinterObjs = {}
hypotTkinterObjsLabel = {}
hypotArcDetectionBool = None

# UI Variables
rectangles = {}
statusText = {}
root = tk.Tk()
root.geometry('1800x1000')
root.title('Laser Hipot Continuity')


def toolwindow(window):
    """Hide the minimise, maximise and close buttons on a popup (Windows only; ignored elsewhere)."""
    try:
        window.attributes('-toolwindow', True)
    except tk.TclError:
        pass


def maximize():
    try:
        root.state('zoomed')  # Windows
    except tk.TclError:
        root.geometry('1800x1000')


maximize()
backgroundColor = '#2A2E32'
canvasColor = '#3D434B'
enabledColor = '#26A671'
halfDisabledColor = '#F9A825'
disabledColor = '#DE0A02'
textBackgroundColor = '#2A2E32'
textColor = 'White'

root.configure(bg=backgroundColor)
# toolwindow(root)


def get_usb_hwids():
    ports = serial.tools.list_ports.comports()
    for port in ports:
        hwid = port.hwid.split('SER=')[1]
        usbHwids.add(hwid)
    print(f"HWIDs: {usbHwids}")
    logger.info(f"HWIDs: {usbHwids}")


def find_com_port_by_hwid_number(targetHwidNumber):
    ports = serial.tools.list_ports.comports()
    for port in ports:
        # Print all device details for debugging purposes
        print(f"Device: {port.device}, Description: {port.description}, HWID: {port.hwid}")
        logger.info(f"Device: {port.device}, Description: {port.description}, HWID: {port.hwid}")
        if targetHwidNumber in port.hwid:
            logger.info('Hwid number: ' + targetHwidNumber + ' Located at: ' + port.device)
            return port.device
    return None


def concat_port(comPort):
    try:
        print(f'Com Port: {comPort}')
        logger.info(f'Com Port: {comPort}')
        print('Serial Port Alias: ASRL' + comPort.replace("COM", '') + '::INSTR')
        logger.info('Serial Port Alias: ASRL' + comPort.replace("COM", '') + '::INSTR')
        return 'ASRL' + comPort.replace("COM", '') + '::INSTR'
    except Exception as ex:
        logger.error(f'Concat port error with {comPort}: {ex}')
        print(f'Concat port error with {comPort}: {ex}')
        return None


# Devices are found by USB serial number, never by COM number, because Windows renumbers ports.
DEVICE_ROLES = {'hypot1': 'Hypot 1', 'hypot2': 'Hypot 2', 'switch1': 'Switch 1', 'switch2': 'Switch 2'}
SWITCH_OPTIONS = 'Cache=false, InterchangeCheck=false, QueryInstrStatus=true, RangeCheck=false, RecordCoercions=false, Simulate=false'
drivers = {role: None for role in DEVICE_ROLES}
deviceStatus = {role: {'ok': False, 'detail': 'not connected', 'hwid': ''} for role in DEVICE_ROLES}
hypotDriver1 = hypotDriver2 = switchDriver1 = switchDriver2 = None


def role_hwid(role):
    return {'hypot1': hypotHwid1, 'hypot2': hypotHwid2, 'switch1': switchHwid1, 'switch2': switchHwid2}[role]


def set_role_hwid(role, hwid):
    global hypotHwid1, hypotHwid2, switchHwid1, switchHwid2
    if role == 'hypot1':
        hypotHwid1 = hwid
    elif role == 'hypot2':
        hypotHwid2 = hwid
    elif role == 'switch1':
        switchHwid1 = hwid
    else:
        switchHwid2 = hwid


def connect_device(role, hwid=None):
    """(Re)connects one tester or switch by its USB serial, closing the old driver first. Safe to
    call while idle, from the device panel, without restarting the program. Returns True on success."""
    global hypotDriver1, hypotDriver2, switchDriver1, switchDriver2
    name = DEVICE_ROLES[role]
    if hwid:
        set_role_hwid(role, hwid)
    hwid = role_hwid(role)
    old = drivers.get(role)
    if old is not None:
        try:
            old.close()
        except Exception as ex:
            logger.warning(f'{name}: closing the old driver failed: {ex}')
    drivers[role] = None
    port = concat_port(find_com_port_by_hwid_number(hwid)) if hwid else None
    try:
        if port is None:
            raise RuntimeError(f'no USB device with serial {hwid or "(none)"} attached')
        if role.startswith('hypot'):
            driver = cc.CreateObject('ARI38XX.ARI38XX', interface=ARI38XXLib.IARI38XX)
            driver.Initialize(port, True, False, 'DriverSetup=BaudRate=38400, QueryInstrStatus=true')
        else:
            driver = cc.CreateObject('SC6540.SC6540', interface=SC6540Lib.ISC6540)
            driver.Initialize(port, True, False, SWITCH_OPTIONS)
        drivers[role] = driver
        deviceStatus[role] = {'ok': True, 'detail': f'{port}, serial {hwid}', 'hwid': hwid}
        logger.info(f'{name} connected on {port} (serial {hwid})')
        print(f'{name} connected on {port} (serial {hwid})')
        if f'Connection to {name} failed' in errors:
            errors.remove(f'Connection to {name} failed')
    except Exception as ex:
        deviceStatus[role] = {'ok': False, 'detail': str(ex)[:90], 'hwid': hwid or ''}
        logger.error(f'Connection to {name} failed: {ex}')
        print(f'Connection to {name} failed: {ex}')
        if f'Connection to {name} failed' not in errors:
            errors.append(f'Connection to {name} failed')
    hypotDriver1, hypotDriver2, switchDriver1, switchDriver2 = drivers['hypot1'], drivers['hypot2'], drivers['switch1'], drivers['switch2']
    refresh_device_panel()
    return deviceStatus[role]['ok']


def disable_all_channels():
    for role in ('switch1', 'switch2'):
        if drivers[role] is not None:
            drivers[role].Execution.DisableAllChannels()


def write_settings():
    with open(SETTINGS_PATH, 'w') as configfile:
        config.write(configfile)


def get_settings():
    config.read(SETTINGS_PATH)
    global errors
    global adminPassword
    global hypotHwid1
    global hypotHwid2
    global switchHwid1
    global switchHwid2
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
    # Hypot. A value that will not parse is a stop, not a warning: the test must not run on a
    # half-loaded settings dictionary.
    global settingsValid
    settingsValid = True
    if 'Hypot' not in config:
        config['Hypot'] = {}
    for key, defaultValue in defaultHypotSettings.items():
        if key not in config['Hypot']:
            hypotSettings[key] = defaultValue
            continue
        try:
            if key in ('arcdetection', 'continuitytest'):
                hypotSettings[key] = config['Hypot'].getboolean(key)  # bool('False') is True; getboolean is not
            elif key in ('voltage', 'currenthighlimit', 'arcsenselevel', 'frequency'):
                hypotSettings[key] = int(float(config['Hypot'][key]))
            else:  # currentlowlimit and the times and resistances are decimals
                hypotSettings[key] = float(config['Hypot'][key])
        except Exception as ex:
            settingsValid = False
            logger.error(f"Bad value for {key} in settings.ini: {ex}")
            print(f"Bad value for {key} in settings.ini: {ex}")
            errors.append(f"Bad value for {key} in settings.ini. Fix it and restart.")

    updateHWIDS = {}
    try:
        hypotHwid1 = config['Hardware IDs']['hypot1']
    except Exception as ex:
        logger.error(f"No hypot1 hwid var in settings.ini: {ex}")
        print(f"No hypot1 hwid var in settings.ini: {ex}")
        updateHWIDS['hypot1'] = hypotHwid1
    try:
        hypotHwid2 = config['Hardware IDs']['hypot2']
    except Exception as ex:
        logger.error(f"No hypot2 hwid var in settings.ini: {ex}")
        print(f"No hypot2 hwid var in settings.ini: {ex}")
        updateHWIDS['hypot2'] = hypotHwid2
    try:
        switchHwid1 = config['Hardware IDs']['switch1']
    except Exception as ex:
        logger.error(f"No switch1 hwid var in settings.ini: {ex}")
        print(f"No switch1 hwid var in settings.ini: {ex}")
        updateHWIDS['switch1'] = switchHwid1
    try:
        switchHwid2 = config['Hardware IDs']['switch2']
    except Exception as ex:
        logger.error(f"No switch2 hwid var in settings.ini: {ex}")
        print(f"No switch2 hwid var in settings.ini: {ex}")
        updateHWIDS['switch2'] = switchHwid2

    for device, hwid in updateHWIDS.items():    # Call to write hwids to settings.ini if they're missing
        default_hwid_conf(device, hwid)

    global laserIP
    if 'Laser' not in config:
        config['Laser'] = {}
    if not config['Laser'].get('ip'):
        config['Laser']['ip'] = laserIP
    laserIP = config['Laser']['ip'].strip()
    refresh_settings_panel()


def default_hwid_conf(device, hwid):
    config['Hardware IDs'][device] = hwid
    write_settings()


def save_settings():
    # Write the config object to a file
    print("Attempting to Save Settings")
    logger.info("Attempting to Save Settings")
    with open(SETTINGS_PATH, 'w') as configfile:
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
        config.write(configfile)  # Close and save to settings file
    get_settings()  # reload so the next run and the settings view use what was just saved
    update_colors(canvas)


def attached_ports():
    """Serial devices attached right now as (label, serial); a fresh list on every call."""
    out = []
    for port in serial.tools.list_ports.comports():
        ser = port.hwid.split('SER=')[1].split(' ')[0] if 'SER=' in port.hwid else ''
        out.append((f'{port.device}   {port.description}   serial {ser or "?"}', ser))
    return out


def device_settings():
    """Change which tester or switch fills each role while the program runs. Connect swaps the
    driver in place and saves the serial to settings.ini; no restart and no file editing."""
    global deviceWindow
    try:
        if deviceWindow.winfo_exists():
            deviceWindow.focus_force()
            return
    except (NameError, tk.TclError):
        pass
    deviceWindow = tk.Toplevel(root)
    deviceWindow.title('Devices')
    deviceWindow.attributes('-topmost', True)
    deviceWindow.configure(bg=backgroundColor, padx=14, pady=10)
    tk.Label(deviceWindow, text='Devices', font=helv, fg=textColor, bg=backgroundColor).grid(row=0, column=0, columnspan=4, sticky='w', pady=(0, 2))
    tk.Label(deviceWindow, text='Pick the attached device for each role and press Connect. The serial number is saved to settings.ini so it survives a restart. Not available while a test is running.',
             font=helvsmall, fg='#B0B8C0', bg=backgroundColor, wraplength=900, justify='left').grid(row=1, column=0, columnspan=4, sticky='w', pady=(0, 8))
    combos, status = {}, {}

    def paint(role):
        st = deviceStatus[role]
        status[role].config(text=('Connected: ' if st['ok'] else 'Not connected: ') + st['detail'], fg=enabledColor if st['ok'] else halfDisabledColor)

    def refresh_lists():
        ports = attached_ports()
        for role, combo in combos.items():
            combo['values'] = [p[0] for p in ports]
            current = next((p[0] for p in ports if p[1] and p[1] == role_hwid(role)), '')
            combo.set(current or (f'(serial {role_hwid(role)} is not attached)' if role_hwid(role) else '(none)'))
            paint(role)

    def connect(role):
        if startButton['state'] == 'disabled':
            messagebox.showwarning('Devices', 'Wait for the test to finish before changing devices.', parent=deviceWindow)
            return
        chosen = combos[role].get()
        serialNo = next((p[1] for p in attached_ports() if p[0] == chosen), None)
        if not serialNo:
            messagebox.showwarning('Devices', 'Pick an attached device from the list first.', parent=deviceWindow)
            return
        if connect_device(role, serialNo):
            config['Hardware IDs'][role] = serialNo
            write_settings()
        paint(role)
        update_error_text()

    row = 2
    for role, name in DEVICE_ROLES.items():
        tk.Label(deviceWindow, text=name, font=helvmedium, fg=textColor, bg=backgroundColor).grid(row=row, column=0, sticky='w', pady=5, padx=(0, 10))
        combos[role] = ttk.Combobox(deviceWindow, width=54, state='readonly')
        combos[role].grid(row=row, column=1, padx=(0, 8))
        tk.Button(deviceWindow, text='Connect', command=lambda r=role: connect(r), bg='#000000', fg=textColor, relief='flat', width=9, font=helvsmall).grid(row=row, column=2, padx=(0, 10))
        status[role] = tk.Label(deviceWindow, text='', font=helvsmall, fg=textColor, bg=backgroundColor, anchor='w', width=52)
        status[role].grid(row=row, column=3, sticky='w')
        row += 1

    tk.Label(deviceWindow, text='Laser marker', font=helvmedium, fg=textColor, bg=backgroundColor).grid(row=row, column=0, sticky='w', pady=5, padx=(0, 10))
    ipVar = tk.StringVar(value=laserIP)
    ttk.Entry(deviceWindow, textvariable=ipVar, width=24).grid(row=row, column=1, sticky='w')
    laserStatus = tk.Label(deviceWindow, text='', font=helvsmall, fg=textColor, bg=backgroundColor, anchor='w', width=52)
    laserStatus.grid(row=row, column=3, sticky='w')

    def paint_laser():
        laserStatus.config(text=f'Connected: {laserIP}:50000' if laserConnected else f'Not connected: {laserIP} did not answer', fg=enabledColor if laserConnected else halfDisabledColor)

    def reconnect_laser():
        global laserIP
        if startButton['state'] == 'disabled':
            messagebox.showwarning('Devices', 'Wait for the test to finish before changing devices.', parent=deviceWindow)
            return
        laserIP = ipVar.get().strip() or laserIP
        config['Laser']['ip'] = laserIP
        write_settings()
        connect_laser()
        paint_laser()
        refresh_device_panel()
        update_error_text()

    tk.Button(deviceWindow, text='Reconnect', command=reconnect_laser, bg='#000000', fg=textColor, relief='flat', width=9, font=helvsmall).grid(row=row, column=2, padx=(0, 10))
    row += 1
    tk.Button(deviceWindow, text='Rescan ports', command=refresh_lists, bg='#000000', fg=textColor, relief='flat', width=12, font=helvsmall).grid(row=row, column=1, sticky='w', pady=(12, 4))
    tk.Button(deviceWindow, text='Close', command=deviceWindow.destroy, bg='#000000', fg=textColor, relief='flat', width=9, font=helvsmall).grid(row=row, column=2, pady=(12, 4))
    refresh_lists()
    paint_laser()


def refresh_device_panel():
    """Status dots on the main window: testers, switches, laser and the machine-data link."""
    if not devicePanelRows:
        return
    for role, (dot, oval, detail) in devicePanelRows.items():
        if role == 'laser':
            ok, text = laserConnected, (f'{laserIP}:50000' if laserConnected else f'{laserIP} did not answer')
        elif role == 'machinedata':
            md = globals().get('machineData')
            ok = md is not None and md.client is not None
            text = f'{md.host} as machine {md.mid}' if ok else ('off in settings.ini' if md is None or not md.enabled else 'not connected')
        else:
            ok, text = deviceStatus[role]['ok'], deviceStatus[role]['detail']
        dot.itemconfig(oval, fill=enabledColor if ok else ('#777' if role == 'machinedata' else disabledColor))
        detail.config(text=text)


def refresh_settings_panel():
    """Read-only view of the values the next run will use, amber where they differ from the
    program defaults (a settings.ini that has drifted from the code)."""
    if not settingsPanelRows:
        return
    drift = 0
    for key, _ in HYPOT_PARAMS:
        value = hypotSettings.get(key, '')
        text = ('Yes' if value else 'No') if isinstance(defaultHypotSettings[key], bool) else str(value)
        same = values_match(value, defaultHypotSettings[key])
        drift += 0 if same else 1
        settingsPanelRows[key].config(text=text, fg=textColor if same else halfDisabledColor)
    settingsNote.config(text='These must match the LHC work instruction. Change them in Admin Settings; every run logs them and reads them back from both testers before the first cavity.'
                        + (f' {drift} value{"s" if drift > 1 else ""} differ from the program defaults.' if drift else ''))


def fault():
    faultWindow = tk.Toplevel(root)
    faultWindow.geometry('1300x520')
    faultWindow.title('Part Fault')
    for col in (0, 3, 6):  # three result columns share the width
        faultWindow.columnconfigure(col, weight=1, minsize=280)
    toolwindow(faultWindow)
    faultWindow.attributes('-topmost', True)  # Force it to be above all other program windows
    faultBackgroundColor = '#DE3C4B'
    faultWindow.configure(bg=faultBackgroundColor)
    faultWindow.lift()

    faultLabel = tk.Label(faultWindow, text='Cavity failed a test', font=helv, fg=textColor, bg=faultBackgroundColor)
    faultLabel.grid(row=0, column=0, columnspan=8, pady=8)

    faultResetButton = tk.Button(faultWindow, text='Reset', command=lambda: reset(closeWindow=True, window=faultWindow), bg='#000000', fg=textColor, relief='flat', width=7,
                                 height=2, font=helvmedium)
    faultResetButton.grid(row=13, column=0, columnspan=8, padx=3, pady=12)

    continuityFaultList = {}
    continuityFaultHeader = tk.Label(faultWindow, text='Continuity Failures', font=helvUnderline, fg=textColor, bg=faultBackgroundColor)
    continuityFaultHeader.grid(row=1, column=0, columnspan=2, pady=5)
    for cavity, value in cavityContinuitySuccesses.items():
        if not value:  # If failed continuity test
            logger.info('Continuity fail on Cavity: ' + str(cavity))
            continuityFaultList[cavity] = tk.Label(faultWindow, text='Cavity ' + str(cavity), font=helvmedium, fg=textColor, bg=faultBackgroundColor)
            continuityFaultList[cavity].grid(row=cavity + 2, column=0)

    hypotFaultList = {}
    hypotFaultHeader = tk.Label(faultWindow, text='Hypot Failures', font=helvUnderline, fg=textColor, bg=faultBackgroundColor)
    hypotFaultHeader.grid(row=1, column=3, columnspan=2, pady=5)
    for cavity, value in cavityHypotSuccesses.items():
        if not value:  # If failed hypot test
            logger.info('Hypot fail on Cavity: ' + str(cavity))
            hypotFaultList[cavity] = tk.Label(faultWindow, text='Cavity ' + str(cavity), font=helvmedium, fg=textColor, bg=faultBackgroundColor)
            hypotFaultList[cavity].grid(row=cavity + 2, column=3)

    laserFaultList = {}
    laserFaultHeader = tk.Label(faultWindow, text='Not Marked', font=helvUnderline, fg=textColor, bg=faultBackgroundColor)
    laserFaultHeader.grid(row=1, column=6, columnspan=2, pady=5)
    for cavity, value in cavityLaserSuccesses.items():
        if not value:  # Passed both tests but the laser did not mark it
            logger.info('Laser fail on Cavity: ' + str(cavity))
            laserFaultList[cavity] = tk.Label(faultWindow, text='Cavity ' + str(cavity), font=helvmedium, fg=textColor, bg=faultBackgroundColor)
            laserFaultList[cavity].grid(row=cavity + 2, column=6)


def non_fault():
    nonFaultWindow = tk.Toplevel(root)
    nonFaultWindow.geometry('700x350')
    nonFaultWindow.title('Test Complete')
    toolwindow(nonFaultWindow)
    nonFaultWindow.attributes('-topmost', True)  # Force it to be above all other program windows
    nonFaultBackgroundColor = '#769C1C'
    nonFaultWindow.configure(bg=nonFaultBackgroundColor)
    nonFaultWindow.lift()

    nonFaultLabel = tk.Label(nonFaultWindow, text='All Parts Good', font=helv, fg=textColor, bg=nonFaultBackgroundColor)
    nonFaultLabel.place(x=100, y=50)

    nonFaultResetButton = tk.Button(nonFaultWindow, text='Reset', command=lambda: reset(closeWindow=True, window=nonFaultWindow), bg='#000000', fg=textColor, relief='flat', width=7, height=2, font=helvmedium)
    nonFaultResetButton.place(x=100, y=100)



def reset(closeWindow, window):
    global faultState
    faultState = False
    if closeWindow:
        window.destroy()
    disable_all_channels()
    # Set all Output Variables to 0
    for cavity in cavityContinuitySuccesses:
        cavityContinuitySuccesses[cavity] = 0
    for cavity in cavityHypotSuccesses:
        cavityHypotSuccesses[cavity] = 0
    for cavity in cavityLaserSuccesses:
        cavityLaserSuccesses[cavity] = 0
    machineData.state('IDLE', 'reset')
    time.sleep(1)  # Make sure double clicks dont accidentally start it again
    startButton["state"] = "normal"  # Re-enables start button

def settings_summary():
    return ', '.join(f'{key}={hypotSettings[key]}' for key, _ in HYPOT_PARAMS)


def values_match(actual, expected):
    if isinstance(expected, bool):
        return bool(actual) == expected
    try:
        return abs(float(actual) - float(expected)) < 0.005
    except (TypeError, ValueError):
        return actual == expected


def create_hypot_tests():
    """Load the test onto both testers and read every parameter back. Returns False, and the run
    must not start, if a step throws or a parameter reads back different from what was sent."""
    logger.info(f'Test settings for this run: {settings_summary()}')
    print(f'Test settings for this run: {settings_summary()}')
    for name, hypotDriver in (('hypot1', hypotDriver1), ('hypot2', hypotDriver2)):
        # Hypot manual results read on page 83
        #   Add ACW test item by AddACWTest()
        try:
            hypotDriver.Files.Delete(1)
        except Exception as ex:
            logger.error(f"Error deleting hypot test on {name}, run aborted: {ex}")
            print(f"Error deleting hypot test on {name}, run aborted: {ex}")
            errors.append(f'{name}: could not clear the old test. Reset and try again.')
            return False
        try:
            hypotDriver.Steps.AddACWTestWithDefaults()
            for key, prop in HYPOT_PARAMS:
                setattr(hypotDriver.Parameters, prop, hypotSettings[key])
            hypotDriver.Files.Save()
        except Exception as ex:
            # Before this returned False the tester was left holding AddACWTestWithDefaults()
            # and the run went ahead on the instrument's own defaults.
            logger.error(f"Error creating hypot test on {name}, run aborted: {ex}")
            print(f"Error creating hypot test on {name}, run aborted: {ex}")
            errors.append(f'{name}: could not load the test settings. Reset and try again.')
            return False
        for key, prop in HYPOT_PARAMS:
            try:
                actual = getattr(hypotDriver.Parameters, prop)
            except Exception as ex:
                logger.warning(f'{name} {prop} could not be read back: {ex}')
                continue
            if not values_match(actual, hypotSettings[key]):
                logger.error(f'{name} {prop} read back {actual}, expected {hypotSettings[key]}; run aborted')
                print(f'{name} {prop} read back {actual}, expected {hypotSettings[key]}; run aborted')
                errors.append(f'{name}: {key} on the tester is {actual}, not {hypotSettings[key]}. Do not run.')
                machineData.alarm('SETTINGS', f'{name} {key} read back {actual}, expected {hypotSettings[key]}')
                return False
        logger.info(f'{name} test loaded and verified')
    return True


def preflight():
    """Everything that has to be true before a fixture run may start."""
    if not settingsValid:
        errors.append('settings.ini has a bad value. Fix it and restart the program.')
        return False
    missing = [DEVICE_ROLES[r] for r in DEVICE_ROLES if drivers[r] is None]
    if missing:
        errors.append('Not connected: ' + ', '.join(missing) + '. Use Change devices.')
        machineData.alarm('DEVICES', 'not connected: ' + ', '.join(missing))
        return False
    laserWanted = any(runCavity[c].get() == 1 and laserEnabled[c].get() == 1 for c in runCavity)
    if laserWanted and not laserConnected and not connect_laser():
        errors.append('Laser marker not connected. Reconnect it, or disable the laser for this run.')
        machineData.alarm('LASER', 'laser marker not connected at start of run')
        return False
    return create_hypot_tests()


def run_samples():
    """Per-run values for the machine database: every test parameter the run used, with the
    program's own default as the limits so a settings.ini that drifted from the code shows as out
    of spec, plus each cavity's result."""
    samples = {}
    for key, _ in HYPOT_PARAMS:
        expected = defaultHypotSettings[key]
        value = hypotSettings[key]
        samples[key] = {'v': float(value), 'lo': float(expected), 'hi': float(expected)}
    for cav in range(1, 11):
        if runCavity['cavity' + str(cav)].get() == 1:
            samples[f'cavity{cav}_continuity'] = cavityContinuitySuccesses[cav]
            samples[f'cavity{cav}_hypot'] = cavityHypotSuccesses[cav]
            samples[f'cavity{cav}_laser'] = cavityLaserSuccesses[cav]
    return samples


def start():
    startButton["state"] = "disabled"  # Disabled start button so its not running twice at the same time due to threading
    disabledCavs = 0
    global faultState
    if not preflight():
        update_error_text()
        startButton["state"] = "normal"
        return
    runStarted = time.time()
    machineData.state('RUNNING', 'fixture run')
    for cavity, value in runCavity.items():
        if value.get() == 0:
            disabledCavs += 1
    totalProgressBar['value'] = (disabledCavs * 10)
    totalProgressPercentage.configure(text=str(int(totalProgressBar['value'])) + ' %')  # Updates displayed percentage. Conv to int to remove decimals
    for i in range(1, 11):
        cavityContinuitySuccesses[i] = 0
        cavityHypotSuccesses[i] = 0
        cavityLaserSuccesses[i] = 0

    for cavity, value in runCavity.items():
        cavitynum = ''.join([char for char in cavity if char.isdigit()])
        cavitynum = int(cavitynum)
        if value.get() == 1:    # If cavity Enabled
            print('Running Cavity: ' + str(cavitynum))
            logger.info('Running Cavity: ' + str(cavitynum))

            disable_all_channels()

            hypot_setup(cavitynum)
            hypot_execution(cavityNum=cavitynum)

            disable_all_channels()

            totalProgressBar.step(10)
            totalProgressPercentage.configure(text=str(int(totalProgressBar['value'])) + ' %')  # Updates displayed percentage. Conv to int to remove decimals
        else: # If cavity Disabled
            cavityContinuitySuccesses[cavitynum] = 3
            cavityHypotSuccesses[cavitynum] = 3  # Dont show on fault window, but don't do other functions either
            cavityLaserSuccesses[cavitynum] = 3
        if laserEnabled['cavity' + str(cavitynum)].get() == 1:
            print('Lasering Cavity: ' + str(cavitynum))
            logger.info('Lasering Cavity: ' + str(cavitynum))
            laser(cavitynum)
        else:
            cavityLaserSuccesses[cavitynum] = 3
            print('Laser Disabled. Skipping Cavity: ' + str(cavitynum))
            logger.info('Laser Disabled. Skipping Cavity: ' + str(cavitynum))
        print(f"Continuity results: {cavityContinuitySuccesses}")
        logger.info(f"Continuity results: {cavityContinuitySuccesses}")
        print(f"Hypot results:      {cavityHypotSuccesses}")
        logger.info(f"Hypot results:      {cavityHypotSuccesses}")

        if cavityContinuitySuccesses[cavitynum] == 0 or cavityHypotSuccesses[cavitynum] == 0 or cavityLaserSuccesses[cavitynum] == 0:
            print(f"Fault State True")
            logger.info(f"Fault State True")
            faultState = True

        print('=================================')  # Separate cavities for testing readability
    tested = [c for c in range(1, 11) if runCavity['cavity' + str(c)].get() == 1]
    good = sum(1 for c in tested if cavityContinuitySuccesses[c] == 1 and cavityHypotSuccesses[c] == 1 and cavityLaserSuccesses[c] != 0)
    machineData.cycle(round(time.time() - runStarted, 1), good, len(tested) - good, run_samples())
    if faultState:  # If any part has a problem, have operators acknowledge they took care of it before starting again
        machineData.state('DOWN', 'cavity failed, waiting for reset')
        fault()
    else:
        machineData.state('IDLE', 'run complete')
        non_fault()


def start_start():  # This is to put the main loop on a separate thread so it can be emergency stopped
    mainThread = Thread(target=start)
    mainThread.start()

def close_drivers():
    for role, driver in drivers.items():
        if driver is None:
            continue
        try:
            driver.close()
        except Exception as ex:
            logger.warning(f'{DEVICE_ROLES[role]}: close failed: {ex}')

def stop():
    logger.error('Emergency Stop Used!')
    print('Emergency Stop Used!')
    machineData.close()
    try:
        disable_all_channels()
        close_drivers()
        print('Program exited cleanly')
        # noinspection PyProtectedMember
        os._exit(os.X_OK)  # Force exits program with status OK
    except Exception as ex:
        logger.error(f"Error during emergency stop!: {ex}")
        print(f"Error during emergency stop!: {ex}")
        close_drivers()
        # noinspection PyProtectedMember
        os._exit(os.X_OK)  # Force exits program with status OK
    finally:
        # Ensure that the Tkinter main loop exits cleanly
        logger.error('Hit finally in emergency stop!')
        print('Hit finally in emergency stop!')
        # noinspection PyProtectedMember
        os._exit(os.X_OK)  # Force exits program with status OK


def on_stop_button_clicked():
    stopThread = Thread(target=stop)
    stopThread.start()


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
    time.sleep(0.1)

    logger.info('Hypot Setup Done')
    print('Hypot Setup Done')


def hypot_execution(cavityNum):
    if cavityNum <= 5:
        hypotDriver = hypotDriver1
    else:
        hypotDriver = hypotDriver2

    try:
        # Start test
        hypotDriver.Execution.Execute()
        # Output Results
        read_hypot(hypotDriver=hypotDriver, cavityNum=cavityNum)
    except Exception as ex:
        logger.error('Exception occured at Hypot execution: ' + str(ex))
        errors.append('Exception occured at Hypot execution: ' + str(ex))
        print('Exception occured at Hypot execution: ' + str(ex))

    hypotDriver.Execution.Abort()
    print('Hypot Execution Done')
    logger.info('Hypot Execution Done')


def read_hypot(hypotDriver, cavityNum):
    global faultState
    lastOpcStatus = False
    deadline = time.time() + hypotSettings['rampuptime'] + hypotSettings['dwelltime'] + hypotSettings['rampdowntime'] + 30
    while (True):
        if time.time() > deadline:
            logger.error(f'Cavity {cavityNum}: no result from the tester in time; counted as a failure')
            print(f'Cavity {cavityNum}: no result from the tester in time; counted as a failure')
            errors.append(f'Cavity {cavityNum}: tester did not answer. Check the tester.')
            machineData.alarm('HYPOT_TIMEOUT', f'cavity {cavityNum}: no result from tester')
            faultState = True
            break
        output = hypotDriver.Execution.ReadTestDisplayRaw().split(',')  # Split into an array for data parsing
        print(output)
        logger.info('Raw Output: ' + hypotDriver.Execution.ReadTestDisplayRaw())

        hypotDriver.System.WriteString('*OPC?\n')
        opcStatus = hypotDriver.System.ReadString()

        continuityFailureTypes = ['Cont. Hi-Lmt']
        hypotFailureTypes = ['HI-LIMIT', 'Short', 'Breakdown']
        if ('1' in opcStatus and lastOpcStatus):
            # Successes
            if output[2] == 'PASS':
                cavityHypotSuccesses[cavityNum] = 1
                cavityContinuitySuccesses[cavityNum] = 1
                print('Cavity ' + str(cavityNum) + ' passes hypot & continuity')
                logger.info('Cavity ' + str(cavityNum) + ' passes hypot & continuity')
            elif output[2] in continuityFailureTypes:
                cavityContinuitySuccesses[cavityNum] = 0
                cavityHypotSuccesses[cavityNum] = 2
                print('Cavity ' + str(cavityNum) + ' fails Continuity')
                logger.info('Cavity ' + str(cavityNum) + ' fails Continuity')
                faultState = True
            elif output[2] in hypotFailureTypes:
                cavityContinuitySuccesses[cavityNum] = 1
                cavityHypotSuccesses[cavityNum] = 0
                print('Cavity ' + str(cavityNum) + ' fails Hypot')
                logger.info('Cavity ' + str(cavityNum) + ' fails Hypot')
                faultState = True
            break
        lastOpcStatus = '1' in opcStatus
        time.sleep(0.1)

def laser_ok(response):
    parts = response.split(',')
    return len(parts) > 1 and parts[1].strip() == 'OK'


def laser(cavityNum):
    global faultState
    if cavityHypotSuccesses[cavityNum] == 1 and cavityContinuitySuccesses[cavityNum] == 1:  # Only Laser if passes both tests
        programNo = cavityNum - 1  # Laser programs array starts at 0
        if (laser_ok(send_laser('RX,Ready\r'))
                and laser_ok(send_laser('WX,ProgramNo=' + str(programNo) + '\r'))
                and laser_ok(send_laser('WX,StartMarking\r'))):
            send_laser('RX,ProgramNo\r')
            cavityLaserSuccesses[cavityNum] = 1
        else:
            # A passed part that did not get marked is a fault, not a log line
            cavityLaserSuccesses[cavityNum] = 0
            faultState = True
            print(f'Cavity {cavityNum}: laser did not mark')
            logger.error(f'Cavity {cavityNum}: laser did not mark')
            machineData.alarm('LASER', f'cavity {cavityNum}: laser did not answer OK')
    else:
        cavityLaserSuccesses[cavityNum] = 2
        print('Skipping Laser due to failed cont or hypot')
        logger.info('Skipping Laser due to failed cont or hypot')

    print('Laser Done')
    logger.info('Laser Done')

def send_laser(msg):
    try:
        print(f'Sending to laser: {msg}')
        logger.info(f'Sending to laser: {msg}')
        laserSocket.send(msg.encode('utf-8'))
        return read_laser()
    except Exception as ex:
        logger.error(f'Issue sending commands to Laser Marker: {ex}')
        print(f'Issue sending commands to Laser Marker: {ex}')
        return 'ng,ng,ng'  # Return no good string


def read_laser():
    try:
        response = laserSocket.recv(1024)  # Listens for data, max amount of bytes specified in ()
        response = response.decode('utf-8')  # Converts bytes to string
        print(f'Laser Output: {response}')
        logger.info(f'Laser Output: {response}')
        return response
    except Exception as ex:
        logger.error(f'Issue receiving commands from Laser Marker: {ex}')
        print(f'Issue receiving commands from Laser Marker: {ex}')
        return 'ng,ng,ng'   # Return no good string


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
        toolwindow(adminWindow)
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

        hardwareSettingsButton = tk.Button(adminWindow, text='Change\ndevices', command=device_settings, bg='#000000', fg=textColor, relief='flat', width=11, height=3, font=helvsmall)
        hardwareSettingsButton.grid(row=12, column=4)

    else: # Wrong password
        adminMessage.config(text='Wrong password')
        root.after(3000, lambda: adminMessage.config(text=''))  # time in ms


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
        errorText.config(text='Ready. All devices connected.', fg=enabledColor)
    else:
        errorText.config(text='\n'.join(errors[-6:]), fg='#FF6B6B')

# Setting values to make sure theyre populated when referenced, or if no settings file found initially
for y in range(1, 11):
    c = 'cavity' + str(y)  # Using cavity(y) instead of int so settings file is a bit more readable
    runCavity[c] = tk.IntVar(value=0)
    laserEnabled[c] = tk.IntVar(value=0)

# Get settings on program start
get_settings()
machineData = MachineData(config['MachineData'], logger, VERSION)
machineData.start()
for _role in DEVICE_ROLES:
    connect_device(_role)
connect_laser()

#       Starting UI Setup
# Fonts and Styles
helv = tkfont.Font(family='Helvetica', size=20, weight='bold')
helvUnderline = tkfont.Font(family='Helvetica', size=20, weight='bold', underline=True)
helvmedium = tkfont.Font(family='Helvetica', size=15, weight='bold')
helvsmall = tkfont.Font(family='Helvetica', size=10, weight='bold')

# UI Setup. Three columns: devices and controls, progress and cavities, settings and admin.
root.columnconfigure(1, weight=1)
root.rowconfigure(1, weight=1)
header = tk.Frame(root, bg=canvasColor, padx=16, pady=8)
header.grid(row=0, column=0, columnspan=3, sticky='ew')
tk.Label(header, text='Laser Hipot Continuity', font=helv, fg=textColor, bg=canvasColor).pack(side=LEFT)
tk.Label(header, text=f'v{VERSION}    settings.ini: {SETTINGS_PATH}', font=helvsmall, fg='#B0B8C0', bg=canvasColor).pack(side=RIGHT)

left = tk.Frame(root, bg=backgroundColor, padx=16, pady=12)
left.grid(row=1, column=0, sticky='nsw')
center = tk.Frame(root, bg=backgroundColor, padx=8, pady=12)
center.grid(row=1, column=1, sticky='n')
right = tk.Frame(root, bg=backgroundColor, padx=16, pady=12)
right.grid(row=1, column=2, sticky='nse')

# Devices
deviceFrame = tk.LabelFrame(left, text=' Devices ', font=helvmedium, fg=textColor, bg=canvasColor, padx=12, pady=8, bd=0)
deviceFrame.pack(fill=X)
for _i, (_role, _name) in enumerate(list(DEVICE_ROLES.items()) + [('laser', 'Laser marker'), ('machinedata', 'Machine data')]):
    _dot = tk.Canvas(deviceFrame, width=14, height=14, bg=canvasColor, highlightthickness=0)
    _dot.grid(row=_i, column=0, padx=(0, 8), pady=3)
    _oval = _dot.create_oval(2, 2, 12, 12, fill='#777', outline='')
    tk.Label(deviceFrame, text=_name, font=helvsmall, fg=textColor, bg=canvasColor, anchor='w', width=13).grid(row=_i, column=1, sticky='w')
    _detail = tk.Label(deviceFrame, text='', font=helvsmall, fg='#B0B8C0', bg=canvasColor, anchor='w', width=44)
    _detail.grid(row=_i, column=2, sticky='w')
    devicePanelRows[_role] = (_dot, _oval, _detail)
tk.Button(deviceFrame, text='Change devices\u2026', command=device_settings, bg='#000000', fg=textColor, relief='flat', width=16, font=helvsmall).grid(row=len(devicePanelRows), column=1, columnspan=2, sticky='w', pady=(8, 2))

# Messages
messageFrame = tk.LabelFrame(left, text=' Messages ', font=helvmedium, fg=textColor, bg=canvasColor, padx=12, pady=8, bd=0)
messageFrame.pack(fill=X, pady=(12, 0))
errorText = tk.Label(messageFrame, text='', fg='#FF6B6B', bg=canvasColor, font=helvsmall, justify='left', anchor='nw', wraplength=470, height=6)
errorText.pack(fill=X)
update_error_text()
if errors:
    machineData.alarm('STARTUP', '; '.join(errors))
    machineData.state('DOWN', errors[0])
else:
    machineData.state('IDLE', 'program started')

startButton = tk.Button(left, text='START', command=start_start, bg=enabledColor, activebackground='#1f8a5d', fg=textColor, relief='flat', width=20, height=6, font=helv)
startButton.pack(pady=(28, 10))
stopButton = tk.Button(left, text='Emergency STOP', command=on_stop_button_clicked, bg=disabledColor, activebackground='#b00', fg=textColor, relief='flat', width=20, height=2, font=helvmedium)
stopButton.pack()
root.protocol("WM_DELETE_WINDOW", on_stop_button_clicked)  # Gracefully shuts down program if window closed

# Progress and the cavity grid
progressCanvas = Canvas(center, width=660, height=96, bg=canvasColor, highlightthickness=0)
progressCanvas.pack(pady=(0, 10))
totalProgressText = tk.Label(progressCanvas, text='Total Progress', fg=textColor, bg=canvasColor, font=helvmedium)
totalProgressText.pack(side=TOP, pady=(8, 4))
totalProgressBar = ttk.Progressbar(progressCanvas, length=600, maximum=100)
totalProgressBar.pack(side=TOP, pady=(0, 8))
totalProgressPercentage = tk.Label(totalProgressBar, text='0 %', fg='Black', bg='#E6E6E6', font=helvsmall)
totalProgressPercentage.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
canvas = Canvas(center, width=650, height=700, bg=canvasColor, highlightthickness=0)
canvas.pack()
create_rectangle_grid(rows=5, columns=2, rectWidth=300, rectHeight=100, padding=50, canv=canvas)

# Test settings, read only
settingsFrame = tk.LabelFrame(right, text=' Test settings (read only) ', font=helvmedium, fg=textColor, bg=canvasColor, padx=12, pady=8, bd=0)
settingsFrame.pack(fill=X)
for _i, (_key, _) in enumerate(HYPOT_PARAMS):
    tk.Label(settingsFrame, text=SETTING_LABELS.get(_key, _key), font=helvsmall, fg='#B0B8C0', bg=canvasColor, anchor='w', width=26).grid(row=_i, column=0, sticky='w', pady=1)
    _value = tk.Label(settingsFrame, text='', font=helvsmall, fg=textColor, bg=canvasColor, anchor='e', width=12)
    _value.grid(row=_i, column=1, sticky='e', pady=1)
    settingsPanelRows[_key] = _value
settingsNote = tk.Label(settingsFrame, text='', font=helvsmall, fg='#B0B8C0', bg=canvasColor, wraplength=330, justify='left')
settingsNote.grid(row=len(HYPOT_PARAMS), column=0, columnspan=2, sticky='w', pady=(10, 0))

# Admin UI
adminFrame = tk.LabelFrame(right, text=' Admin settings ', font=helvmedium, fg=textColor, bg=canvasColor, padx=12, pady=8, bd=0)
adminFrame.pack(fill=X, pady=(14, 0))
adminText = tk.StringVar()
adminTextbox = ttk.Entry(adminFrame, show='*', width=24)
adminTextbox.grid(row=0, column=0, padx=(0, 8), pady=4)
adminSubmitButton = tk.Button(adminFrame, text='Open', command=admin_panel, bg='#000000', fg=textColor, relief='flat', width=9, font=helvsmall)
adminSubmitButton.grid(row=0, column=1)
adminMessage = tk.Label(adminFrame, text='', font=helvsmall, fg=halfDisabledColor, bg=canvasColor, anchor='w')
adminMessage.grid(row=1, column=0, columnspan=2, sticky='w')

# Populate settings.ini file. Need to start admin_panel to populate fields to save
adminTextbox.delete(0, 'end') # Clears Password
adminTextbox.insert(0, adminPassword) # Set Password
admin_panel()
save_settings()
update_colors(canvas)
for widget in root.winfo_children(): # Close admin window
    if isinstance(widget, tk.Toplevel):
        widget.destroy()
adminTextbox.delete(0, 'end')  # Clears Password
refresh_settings_panel()
refresh_device_panel()
root.lift()
#Test each cavity
#switchDriver1.Execution.DisableAllChannels()
#switchDriver2.Execution.DisableAllChannels()
#hypot_setup(10)


try:
    root.mainloop()

except KeyboardInterrupt:
    on_stop_button_clicked()
    sys.exit()