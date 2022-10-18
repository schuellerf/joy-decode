#!/usr/bin/env python3
""" Library and helper program to interface
DPM-8605, DPM-8608, DPM-8616, DPM-8624
https://joy-it.net/de/products
Document "JT-8600_communication-protocol.pdf" recieved on 18th October 2021

Currently only "Simple communication Protocol" is implemented
(i.e. Modbus is not implemented)

Main usecase is just to have logging for https://oekotrainer.de/oekotrainer/01000
when there is a https://oekotrainer.de/spannungsregler/03000 connected
via an RS485 to USB interface e.g. https://joy-it.net/de/products/SBC-TTL-RS485
to your PC

The generated CSV can be imported by e.g. libreoffice with the language "English (USA)" to make sure that the numbers are correct

Usage:
    monitor.py [--interface=<dev>] [--baud=<baud>] [--address_code=<addr>] [--delay=<delay>] [--output=<output>]
               [--comment=<comment>] [--max-watt=<max-watt>] [--voltage-limit=<volt>] [--verbose] [--mqtt-server=<mqttserver>]

Options:
    -i --interface=<dev>           Serial Device
    -b --baud=<baud>               Baudrate [default: 9600]
    -a --address_code=<addr>       Address code of device [default: 1]
    -d --delay=<delay>             Delay for polling of values in milliseconds. Set to 0 to disable polling. [default: 1000]
    -o --output=<output>           Filename to output CSV data, will append if existing [default: power_log.csv]
    -c --comment=<comment>         Optional comment to be added to the data. (e.g. person doing the workout)
                                   Will also be used to override the MQTT topic (if an MQTT server is given) [default: ]
    -w --max-watt=<watt>           Set output current to limit to this power (in watts) [default: 50]
    -l --voltage-limit=<volt>      Set output voltage limit (in volts) [default: 14.5]
    -v --verbose                   Debug output on stderr [default: false]
    -m --mqtt-server=<mqttserver>  Send the data to and MQTT Server
"""

import os
import serial
import re
import time
import datetime
import csv
import sys
import math
import paho.mqtt.client as mqtt

try:
    import pynput
    use_key = True
except:
    print("pynput not found - no keypress support")
    use_key = False

from docopt import docopt
from enum import Enum

if os.name == 'nt':
    print('Windows was not yet tested')
    DEFAULT_INTERFACE = None # might be COM1 - not yet tested
elif os.name == 'posix':
    DEFAULT_INTERFACE = '/dev/ttyUSB0'
else:
    print(f'{os.name} is not supported')
    DEFAULT_INTERFACE = None # should fail anyway

WATT_LOWER_LIMIT = 5
WATT_UPPER_LIMIT = 300

WATT_STEP = 5

TOLERANCE_WATT = 1

class DPM8600:

    START = ':'
    READ = 'r'
    WRITE = 'w'
    END = '\r\n'

    class WriteFunction(Enum):

        WRITE_VOLTAGE_LIMIT = 10 # V/100
        WRITE_CURRENT_LIMIT = 11 # mA
        WRITE_OUTPUT_STATUS = 12 # output off (0), output on (1)
        WRITE_VOLTAGE_AND_CURRENT_LIMIT = 20

        def convert(self, val):
            if self == self.WRITE_VOLTAGE_LIMIT:
                return math.trunc(val * 100)
            elif self == self.WRITE_CURRENT_LIMIT:
                return math.trunc(val * 1000)
            elif self == self.WRITE_OUTPUT_STATUS:
                return "1" if val else "0"
            elif self == self.READ_OUTPUT_TYPE:
                v = math.trunc(val[0] * 100)
                a = math.trunc(val[1] * 1000)
                return f"{v},{a}"

    class ReadFunction(Enum):

        READ_MAX_OUTPUT_VOLTAGE = 0 # V/100
        READ_MAX_OUTPUT_CURRENT = 1 # mA - 5A -> DPM-8605, ...
        READ_VOLTAGE_SETTING = 10 # send 0 to get response
        READ_CURRENT_SETTING = 11 # send 0 to get response
        READ_OUTPUT_STATUS = 12 # output off (0), output on (1)
        READ_OUTPUT_VOLTAGE = 30 # V/100, send 0 to get response
        READ_OUTPUT_CURRENT = 31 # mA,  send 0 to get response
        READ_OUTPUT_TYPE = 32 # ConstantVoltage (CV) = 0, ConstantCurrent (CC) = 1
        READ_TEMPERATURE = 33 # °C

        def convert(self, val):
            if self == self.READ_OUTPUT_STATUS:
                return bool(val)
            elif self in [self.READ_MAX_OUTPUT_VOLTAGE, self.READ_VOLTAGE_SETTING, self.READ_OUTPUT_VOLTAGE]:
                return val / 100
            elif self in [self.READ_MAX_OUTPUT_CURRENT, self.READ_CURRENT_SETTING, self.READ_OUTPUT_CURRENT]:
                return val / 1000
            elif self == self.READ_TEMPERATURE:
                return val
            elif self == self.READ_OUTPUT_TYPE:
                return "CV" if val == 0 else "CC"


    def __init__(self, address_code = 1, baud_rate = 9600, interface = DEFAULT_INTERFACE):
        self.address_code = address_code
        self.baud_rate = baud_rate
        self.interface = interface
        self.serial = None

        try:
            self.serial = serial.Serial(interface, baud_rate, timeout=5)
        except serial.serialutil.SerialException as se:
            print(f"SerialException: {se}")
            sys.exit(1)

        try:
            self.serial.open()
        except serial.serialutil.SerialException:
            pass

        print(f"Initially clearing input buffer")
        self._clearInput()

        self.cmd_re = re.compile(f"{self.START}(?P<addr>\d+)(?P<func>[wr])(?P<func_num>\d+)=((?P<operand>\d+),?)+.?{self.END}")

    def __del__(self):
        if self.serial:
            self.serial.close()

    def _send(self, cmd, operands = [0]):
        global debug
        if isinstance(cmd, self.ReadFunction):
            cmd_name = self.ReadFunction(cmd).name
        else:
            cmd_name = self.WriteFunction(cmd).name
            operands = [cmd.convert(o) for o in operands]

        if cmd_name.startswith("WRITE_"):
            func = self.WRITE
        else:
            func = self.READ

        operands_str = [str(o) for o in operands]
        operand = ",".join(operands_str)

        raw_cmd = f"{self.START}{self.address_code:02}{func}{cmd.value:02}={operand},{self.END}"

        if debug: print(f"out >{raw_cmd}<",file=sys.stderr)
        self.serial.write(raw_cmd.encode())
        self.serial.flush()

    def _clearInput(self):
        dropped = b""
        while self.serial.in_waiting > 0:
            dropped += self.serial.read()
        print(f"dropped '{dropped.decode()}'")

    def _read(self, cmd):
        global debug

        try:
            ret = self.serial.read_until()
        except serial.serialutil.SerialException as se:
            print(f"SerialException {se}", file=sys.stderr)
            return None
        try:
            ret = ret.decode()
        except UnicodeDecodeError as e:
            print(f"UnicodeDecodeError {e}", file=sys.stderr)
            return None
        if debug: print(f"in >{ret}<",file=sys.stderr)

        if ret is None or len(ret) == 0:
            return None

        if not ret.endswith(self.END):
            print(f"TIMEOUT, in get_voltage() only got '{ret}', clearing input")
            self._clearInput()
            return None

        if ret.strip() == ":01ok":
            return True

        m = self.cmd_re.match(ret)

        if m is None:
            print(f"Could not decode '{ret}', clearing input")
            self._clearInput()
            return None

        if int(m.group("func_num")) != cmd.value:
            print(f"Wrong answer! got {m.group('func_num')} expected {cmd.value}, clearing input")
            self._clearInput()
            return None

        ret = cmd.convert(int(m.group("operand")))

        return ret

    def set_voltage_limit(self, v):
        self._send(self.WriteFunction.WRITE_VOLTAGE_LIMIT, [v])
        # need to read confirmation "ok" - although not documented in spec
        return self._read(self.WriteFunction.WRITE_VOLTAGE_LIMIT)

    def set_current_limit(self, a):
        self._send(self.WriteFunction.WRITE_CURRENT_LIMIT, [a])
        # need to read confirmation "ok" - although not documented in spec
        return self._read(self.WriteFunction.WRITE_CURRENT_LIMIT)

    def set_output(self, output):
        self._send(self.WriteFunction.WRITE_OUTPUT_STATUS, [output])
        # need to read confirmation "ok" - although not documented in spec
        return self._read(self.WriteFunction.WRITE_OUTPUT_STATUS)

    def set_max_voltage_and_current(self, v, a):
        self._send(self.WriteFunction.WRITE_VOLTAGE_AND_CURRENT_LIMIT, [v,a])

    def get_voltage(self):
        self._send(self.ReadFunction.READ_OUTPUT_VOLTAGE)
        return self._read(self.ReadFunction.READ_OUTPUT_VOLTAGE)

    def get_current(self):
        self._send(self.ReadFunction.READ_OUTPUT_CURRENT)
        return self._read(self.ReadFunction.READ_OUTPUT_CURRENT)

    def get_output_status(self):
        self._send(self.ReadFunction.READ_OUTPUT_STATUS)
        return self._read(self.ReadFunction.READ_OUTPUT_STATUS)

    def get_output_type(self):
        self._send(self.ReadFunction.READ_OUTPUT_TYPE)
        return self._read(self.ReadFunction.READ_OUTPUT_TYPE)

    def get_temperature(self):
        self._send(self.ReadFunction.READ_TEMPERATURE)
        return self._read(self.ReadFunction.READ_TEMPERATURE)

    def get_voltage_limit(self):
        self._send(self.ReadFunction.READ_VOLTAGE_SETTING)
        return self._read(self.ReadFunction.READ_VOLTAGE_SETTING)

    def get_current_limit(self):
        self._send(self.ReadFunction.READ_CURRENT_SETTING)
        return self._read(self.ReadFunction.READ_CURRENT_SETTING)


def key_press(key):
    global max_watt

    if key == pynput.keyboard.Key.up:
        if max_watt + WATT_STEP <= WATT_UPPER_LIMIT:
            max_watt += WATT_STEP
    elif key == pynput.keyboard.Key.down:
        if max_watt - WATT_STEP >= WATT_LOWER_LIMIT:
            max_watt -= WATT_STEP

# The callback for when the MQTT client receives a CONNACK response from the server.
def on_connect(client, userdata, flags, rc):
    print(f"Connected with result code {rc}")

    # Subscribing in on_connect() means that if we lose the connection and
    # reconnect then subscriptions will be renewed.
    client.subscribe("$SYS/#")

# The callback for when a MQTT  PUBLISH message is received from the server.
def on_message(client, userdata, msg):
    print(f"{msg.topic} {msg.payload}")

if __name__ == "__main__":
    global debug
    global max_watt
    args = docopt(__doc__, version="0.1")

    if args["--interface"] is None:
        args["--interface"] = DEFAULT_INTERFACE
    print(args)
    debug = bool(args["--verbose"])

    dev = DPM8600(address_code = int(args["--address_code"]),
                  baud_rate = int(args["--baud"]),
                  interface = args["--interface"])

    comment = args["--comment"]
    mqtt_server = args.get("--mqtt-server")
    if mqtt_server:
        print(f"Connecting to {mqtt_server}...")
        mqttc = mqtt.Client()
        mqttc.on_connect = on_connect
        mqttc.on_message = on_message
        mqttc.connect(mqtt_server)
        mqttc.loop_start()
        if comment:
            mqtt_topic = comment
        else:
            mqtt_topic = "joy_charger"

        print(f"MQTT Topic is '{mqtt_topic}'")
    else:
        mqttc = None

    max_watt = int(args["--max-watt"])
    max_volt = float(args["--voltage-limit"])
    last_v = 0.0
    last_a = 0.0
    wh_sum = 0.0
    wh_gross = 0.0
    v_sum = 0.0
    a_sum = 0.0
    stat_count = 0
    a_limit = None
    delay_seconds = int(args["--delay"]) / 1000.0

    dev.set_voltage_limit(max_volt)

    if use_key:
        print(f"pynput found - use UP or DOWN key to change watt")
        key_handler = pynput.keyboard.Listener(on_press=key_press)
        key_handler.start()

    with open(args["--output"],'a') as csvfile:
        fieldnames = ['timestamp', 'session_time', 'realtime', 'V', 'A', 'W', 'Wh', 'Wh Sum', 'Vmax', 'Amax', 'Wmax', 'OutputState', 'ConstVolt_ConstCurr','comment']

        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        # only write header once
        if csvfile.tell() == 0:
            writer.writeheader()

        start_time = datetime.datetime.now()
        start_time_monotonic = time.monotonic()
        now = start_time
        now_monotonic = start_time_monotonic
        time.sleep(delay_seconds)

        while True:
            last_time = now
            last_time_monotonic = now_monotonic

            now = datetime.datetime.now()
            now_monotonic = time.monotonic()

            v = dev.get_voltage()
            a = dev.get_current()
            state = dev.get_output_status()
            typ = dev.get_output_type()
            # temp = dev.get_temperature()
            v_lim = dev.get_voltage_limit()
            a_lim = dev.get_current_limit()

            if v is None or a is None:
                continue

            w = v*a

            timespan_s = now_monotonic - last_time_monotonic

            wh = w * timespan_s / 3600.0
            wh_sum += wh

            full_timespan_s = now_monotonic - start_time_monotonic

            stat_count += 1 # also count idle for gross performance
            if v > 0 and a > 0:
                v_sum += v
                a_sum += a
                wh_gross = ( (v_sum/stat_count) * (a_sum/stat_count) ) * full_timespan_s / 3600.0

            print(f"[{now.strftime('%Y-%m-%d %H:%M:%S.%f')} / {now - start_time}] {v:.2f}/{v_lim:.2f} V, {a:.3f}/{a_lim:.3f} A, {w:.3f}/{max_watt:.3f} W, {wh_sum:.3f} Wh, {wh_gross:.3f} Wh_gross, {'ON' if state else 'OFF'}, {typ}")

            d = {'realtime': now.strftime('%Y-%m-%d %H:%M:%S.%f'),
                 'timestamp': now.strftime('%s%f'),
                 'session_time': f"{now - start_time}",
                 'V': v,
                 'Vmax': v_lim,
                 'A': a,
                 'Amax': a_lim,
                 'W': w,
                 'Wmax': max_watt,
                 'Wh': wh,
                 'Wh Sum': wh_sum,
                 'OutputState': 'ON' if state else 'OFF',
                 'ConstVolt_ConstCurr': typ,
                 'comment': comment
                 }
            writer.writerow(d)
            if mqttc:
                for k in d:
                    try:
                        mqttc.publish(f"{mqtt_topic}/{k}", d[k])
                    except Exception as e:
                        print(f"Tried to send value for {k} but failed - {type(d[k])}")
                        print(e)
                        raise

            # Adapt power (only when active - i.e. state==True)
            if w > max_watt + TOLERANCE_WATT and state:
                a_limit = max_watt / v
                print(f"Got {w:.3f} watt - I only want {max_watt:.3f}, lower current limit to {a_limit:.3f}A")
            
            if w < max_watt - TOLERANCE_WATT and typ == "CC" and state:
                a_limit = max_watt / v
                print(f"Got {w:.3f} watt - but I want {max_watt:.3f}, raise current limit to {a_limit:.3f}A")

            if a_limit:
                dev.set_current_limit(a_limit)
                a_limit = None

            all_time_s = now_monotonic - start_time_monotonic
            offset_s = all_time_s % delay_seconds
            time.sleep(delay_seconds - offset_s)

            last_v = v
            last_a = a

