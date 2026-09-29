"""Runs Main.py without hardware: the Ikonix COM drivers, the serial port list, the laser socket
and the machine-data broker are all stubbed, so the window can be opened, driven and screenshotted
on any machine (in CI: xvfb-run python tests/ui_smoke.py). Screenshots land in SHOTS (default /out).

Scenarios, chosen with SCENARIO=<name>:
  connected   every device answers, a run passes all cavities (default)
  missing     hypot 2 and the laser are absent
  fail        cavity 7 fails hipot, cavity 3 is not marked
"""
import os
import subprocess
import sys
import time
import types

SCENARIO = os.environ.get('SCENARIO', 'connected')
SHOTS = os.environ.get('SHOTS', '/out')
os.makedirs(SHOTS, exist_ok=True)
os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.getcwd())

# ---- comtypes and the two IVI driver type libraries ---------------------------------
comtypes = types.ModuleType('comtypes')
client = types.ModuleType('comtypes.client')
gen = types.ModuleType('comtypes.gen')
SC = types.ModuleType('comtypes.gen.SC6540Lib'); SC.ISC6540 = object
ARI = types.ModuleType('comtypes.gen.ARI38XXLib'); ARI.IARI38XX = object; ARI.ARI38XXFrequency60Hz = 60
client.GetModule = lambda name: None


class Bag:
    pass


class FakeHypot:
    created = 0

    def __init__(self):
        FakeHypot.created += 1
        self.n = FakeHypot.created
        self.port = None
        self.Parameters = Bag()
        self.Files = Bag(); self.Files.Delete = lambda i: None; self.Files.Save = lambda: None
        self.Steps = Bag(); self.Steps.AddACWTestWithDefaults = lambda: None
        self.Execution = Bag(); self.Execution.Execute = self.execute; self.Execution.Abort = lambda: None
        self.Execution.ReadTestDisplayRaw = self.display
        self.System = Bag(); self.System.WriteString = lambda s: None; self.System.ReadString = lambda: '1'
        self.runs = 0

    def execute(self):
        self.runs += 1   # one Execute per cavity; hypot 2 runs cavities 6 to 10

    def Initialize(self, port, a, b, opts):
        if port is None or (SCENARIO == 'missing' and 'ASRL4' in str(port)):
            raise RuntimeError('IVI: resource not found ' + str(port))
        self.port = port

    def display(self):
        if SCENARIO == 'fail' and self.port == 'ASRL4::INSTR' and self.runs == 2:   # hypot 2, second cavity = 7
            return '1000,0.5,HI-LIMIT,0'
        return '1000,0.1,PASS,0'

    def close(self):
        pass


class FakeSwitch:
    def __init__(self):
        self.Execution = Bag()
        for name in ('DisableAllChannels',):
            setattr(self.Execution, name, lambda: None)
        self.Execution.ConfigureWithstandChannels = lambda ch: None
        self.Execution.ConfigureReturnChannels = lambda ch: None

    def Initialize(self, port, a, b, opts):
        if port is None:
            raise RuntimeError('IVI: resource not found')

    def close(self):
        pass


client.CreateObject = lambda progid, interface=None: FakeHypot() if 'ARI' in progid else FakeSwitch()
comtypes.client = client
sys.modules.update({'comtypes': comtypes, 'comtypes.client': client, 'comtypes.gen': gen,
                    'comtypes.gen.SC6540Lib': SC, 'comtypes.gen.ARI38XXLib': ARI})

# ---- serial ports -------------------------------------------------------------------
import serial.tools.list_ports as lp


class Port:
    def __init__(self, dev, desc, ser):
        self.device, self.description, self.hwid = dev, desc, f'USB VID:PID=0403:6001 SER={ser} LOCATION=1-1'


PORTS = [Port('COM3', 'USB Serial Port (Hypot 3805)', 'BG01PD8AA'), Port('COM4', 'USB Serial Port (Hypot 3805)', 'BG01PFCQA'),
         Port('COM6', 'USB Serial Port (SC6540)', 'B0007EEKA'), Port('COM7', 'USB Serial Port (SC6540)', 'B0007BEKA'),
         Port('COM9', 'USB Serial Port (spare tester)', 'BG01ZZZZA')]
if SCENARIO == 'missing':
    PORTS = [p for p in PORTS if 'PFCQA' not in p.hwid]
lp.comports = lambda: PORTS

# ---- laser socket -------------------------------------------------------------------
import socket


class FakeSocket:
    def __init__(self, *a, **k):
        self.sent = []

    def settimeout(self, t): pass

    def connect(self, addr):
        if SCENARIO == 'missing':
            raise OSError('No route to host')

    def send(self, data):
        self.sent.append(data)

    def recv(self, n):
        last = self.sent[-1] if self.sent else b''
        if SCENARIO == 'fail' and b'ProgramNo=2' in last:   # cavity 3
            return b'ER,ProgramNo,22\r'
        return (last[:2] + b',OK\r') if last else b'RX,OK\r'

    def close(self): pass


socket.socket = FakeSocket

# ---- drive the window once it is up ---------------------------------------------------
import tkinter as tk

step = 0


def shot(name):
    subprocess.run(['import', '-window', 'root', os.path.join(SHOTS, f'{SCENARIO}-{name}.png')], check=False)


def drive(root):
    try:
        _drive(root)
    except Exception as ex:
        import traceback; traceback.print_exc()
        print('FAILED:', ex)
    finally:
        root.after(200, root.destroy)


def _drive(root):
    M = sys.modules['Main']
    shot('main')
    print('devices:', {k: (v.get('ok'), v.get('detail')) for k, v in M.deviceStatus.items()})
    print('settings panel rows:', len(M.settingsPanelRows))
    print('start enabled:', M.startButton['state'])
    if SCENARIO == 'missing':
        M.start()   # preflight must refuse: hypot 2 is absent
        print('refused:', M.startButton['state'] == 'normal', '|', M.errorText['text'].replace('\n', ' / '))
        assert 'Not connected: Hypot 2' in M.errorText['text']
    else:
        M.start()   # synchronous here; the real button runs it on a thread
        root.update()
        shot('after-run')
        print('results:', M.cavityContinuitySuccesses, M.cavityHypotSuccesses, M.cavityLaserSuccesses)
        for w in root.winfo_children():
            if isinstance(w, tk.Toplevel):
                w.destroy()
        M.reset(False, None)
    M.device_settings()
    root.update()
    shot('devices')
    M.adminTextbox.delete(0, 'end'); M.adminTextbox.insert(0, M.adminPassword); M.admin_panel()
    root.update()
    shot('admin')
    print('OK')


orig = tk.Tk.mainloop


def mainloop(self):
    self.after(1500, lambda: drive(self))
    orig(self)


tk.Tk.mainloop = mainloop
import Main  # noqa: E402,F401  runs the program
