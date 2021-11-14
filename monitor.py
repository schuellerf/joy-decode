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
    monitor.py [--interface=<dev>] [--baud=<baud>] [--address_code=<addr>] [--delay=<delay>] [--output=<output>] [--comment=<comment>] [--max-watt=<max-watt>] [--verbose]

Options:
    -i --interface=<dev>      Serial Device
    -b --baud=<baud>          Baudrate [default: 9600]
    -a --address_code=<addr>  Address code of device [default: 1]
    -d --delay=<delay>        Delay for polling of values in milliseconds. Set to 0 to disable polling. [default: 1000]
    -o --output=<output>      Filename to output CSV data, will append if existing [default: power_log.csv]
    -c --comment=<comment>    Optional comment to be added to the data. (e.g. person doing the workout) [default: ]
    -w --max-watt=<watt>      Set output current to limit to this power (in watts) [default: 50]
    -v --verbose              Debug output on stderr [default: false]
"""

import os
import serial
import re
import time
import datetime
import csv
import sys
import math

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
        except SerialException as se:
            print(f"SerialException {se}", file=sys.stderr)
            return None
        ret = ret.decode()
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

if __name__ == "__main__":
    global debug
    args = docopt(__doc__, version="0.1")

    if args["--interface"] is None:
        args["--interface"] = DEFAULT_INTERFACE
    print(args)
    debug = bool(args["--verbose"])

    dev =  DPM8600(address_code = int(args["--address_code"]),
                   baud_rate = int(args["--baud"]),
                   interface = args["--interface"])

    max_watt = int(args["--max-watt"])
    last_v = 0.0
    last_a = 0.0
    wh_sum = 0.0
    wh_exact = 0.0
    v_sum = 0.0
    a_sum = 0.0
    stat_count = 0
    a_limit = None
    delay_seconds = int(args["--delay"]) / 1000.0

    comment = args["--comment"]

    with open(args["--output"],'a') as csvfile:
        fieldnames = ['timestamp', 'session_time', 'realtime', 'V', 'A', 'W', 'Wh', 'Wh Sum', 'Vmax', 'Amax', 'Wmax', 'OutputState', 'ConstVolt_ConstCurr','comment']

        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        # only write header once
        if csvfile.tell() == 0:
            writer.writeheader()

        start_time = datetime.datetime.now()
        now = start_time
        time.sleep(delay_seconds)

        while True:
            last_time = now
            now = datetime.datetime.now()
            v = dev.get_voltage()
            a = dev.get_current()
            state = dev.get_output_status()
            typ = dev.get_output_type()
            temp = dev.get_temperature()
            v_lim = dev.get_voltage_limit()
            a_lim = dev.get_current_limit()

            if v is None or a is None:
                continue

            w = v*a

            timespan = now - last_time
            timespan_s = timespan.seconds + (timespan.microseconds/1000000.0)

            wh = w * timespan_s / 3600.0
            wh_sum += wh

            full_timespan = now - start_time
            full_timespan_s = full_timespan.seconds + ( full_timespan.microseconds / 1000000.0 )

            if v > 0 and a > 0:
                v_sum += v
                a_sum += a
                stat_count += 1
                wh_exact = ( (v_sum/stat_count) * (a_sum/stat_count) ) * full_timespan_s / 3600.0

            print(f"[{now.strftime('%Y-%m-%d %H:%M:%S.%f')} / {now - start_time}] {v:.2f}/{v_lim:.2f} V, {a:.3f}/{a_lim:.2f} A, {w:.3f}/{max_watt:.3f} W, {wh_sum:.3f} Wh, {wh_exact:.3f} Wh_exact, {'ON' if state else 'OFF'}, {typ}")

            d = {'realtime': now.strftime('%Y-%m-%d %H:%M:%S.%f'),
                 'timestamp': now.strftime('%s%f'),
                 'session_time': now - start_time,
                 'V': v,
                 'Vmax': v_lim,
                 'A': a,
                 'Amax': a_lim,
                 'W': w,
                 'Wmax': max_watt,
                 'Wh': wh,
                 'Wh Sum': wh_exact,
                 'OutputState': 'ON' if state else 'OFF',
                 'ConstVolt_ConstCurr': typ,
                 'comment': comment
                 }
            writer.writerow(d)

            # Adapt power (only when active - i.e. state==True)
            if w > max_watt + TOLERANCE_WATT and state:
                a_limit = max_watt / v
                print(f"Got {w} watt - want only {max_watt}, lower current limit to {a_limit}A")
            
            if w < max_watt - TOLERANCE_WATT and typ == "CC" and state:
                a_limit = max_watt / v
                print(f"Got {w} watt - but I want {max_watt}, raise current limit to {a_limit}A")

            if a_limit:
                dev.set_current_limit(a_limit)
                a_limit = None

            # wait for next round
            if (2*delay_seconds - timespan_s) > 0:
                time.sleep(2*delay_seconds - timespan_s)

            last_v = v
            last_a = a

