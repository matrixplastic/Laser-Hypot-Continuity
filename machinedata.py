"""Publishes the LHC's state, cycles and test settings to the plant machine-data broker.

Topics and payloads follow machinedata/README.md in the ERP repo:
  matrix/machine/<MachineID>/state      {"state":"RUNNING","source":"lhc","detail":"..."}
  matrix/machine/<MachineID>/cycle      {"count":1,"cycle_s":92.4,"good":9,"reject":1,"samples":{...}}
  matrix/machine/<MachineID>/alarm      {"code":"...","text":"...","active":true}
  matrix/machine/<MachineID>/heartbeat  {"fw":"2026.09.29","uptime":1234}

One cycle is one fixture run (up to ten cavities). Its samples carry every test parameter the run
used, with the expected value as the low and high limit, so a run on the wrong settings shows up
as out of spec in the database instead of going unnoticed for a year.

Everything here is fail-open: if the broker is unreachable the tests still run and the outage is
logged. Data collection must never stop production.
"""
import json
import threading
import time

try:
    import paho.mqtt.client as mqtt
except ImportError:  # the exe was built without paho; publishing is simply off
    mqtt = None

DEFAULTS = {
    'enabled': '0',
    'host': '10.15.80.26',
    'port': '1883',
    'user': 'device',
    'password': '',
    'machine_id': '887',   # production.tblmachines.ID for LHC 1, never the friendly number
}


class MachineData:
    def __init__(self, section, logger, version):
        # Fill the section with defaults so the first save_settings() writes the keys out
        for key, value in DEFAULTS.items():
            if key not in section:
                section[key] = value
        self.log = logger
        self.version = version
        self.enabled = section.getboolean('enabled', fallback=False) and mqtt is not None
        self.host = section.get('host')
        self.port = section.getint('port', fallback=1883)
        self.user = section.get('user')
        self.password = section.get('password')
        self.mid = section.getint('machine_id', fallback=887)
        self.topic = f'matrix/machine/{self.mid}/'
        self.started = time.time()
        self.client = None
        if section.getboolean('enabled', fallback=False) and mqtt is None:
            self.log.error('Machine data enabled in settings.ini but paho-mqtt is not installed')

    def start(self):
        if not self.enabled:
            self.log.info('Machine data publishing is off (settings.ini [MachineData] enabled=0)')
            return
        try:
            self.client = mqtt.Client(mqtt.CallbackAPIVersion.VERSION2, client_id=f'lhc{self.mid}')
            self.client.username_pw_set(self.user, self.password)
            self.client.will_set(self.topic + 'state', json.dumps({'state': 'OFF', 'source': 'lhc', 'detail': 'program closed'}), qos=1)
            self.client.on_connect = lambda c, u, f, rc, p=None: self.log.info(f'Machine data broker connected: {rc}')
            self.client.on_disconnect = lambda c, u, f, rc, p=None: self.log.warning(f'Machine data broker disconnected: {rc}')
            self.client.connect_async(self.host, self.port, keepalive=60)
            self.client.loop_start()   # reconnects on its own; publishes queue while offline
            threading.Thread(target=self._heartbeat, daemon=True).start()
        except Exception as ex:
            self.log.error(f'Machine data publishing failed to start, continuing without it: {ex}')
            self.client = None

    def _publish(self, kind, payload):
        if self.client is None:
            return
        try:
            self.client.publish(self.topic + kind, json.dumps(payload), qos=1)
        except Exception as ex:
            self.log.warning(f'Machine data publish {kind} failed: {ex}')

    def _heartbeat(self):
        while True:
            self._publish('heartbeat', {'fw': self.version, 'uptime': int(time.time() - self.started)})
            time.sleep(60)

    def state(self, state, detail=''):
        self._publish('state', {'state': state, 'source': 'lhc', 'detail': detail[:120]})

    def cycle(self, cycle_s, good, reject, samples):
        self._publish('cycle', {'count': 1, 'cycle_s': cycle_s, 'good': good, 'reject': reject, 'samples': samples})

    def alarm(self, code, text, active=True):
        self._publish('alarm', {'code': code[:40], 'text': text[:255], 'active': active})

    def close(self):
        if self.client is None:
            return
        try:
            self.state('OFF', 'program closed')
            self.client.loop_stop()
            self.client.disconnect()
        except Exception:
            pass
