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
    monitor.py [--interface=<dev>] [--baud=<baud>] [--address_code=<addr>] [--delay=<delay>] [--output=<output>] [--comment=<comment>]

Options:
    -i --interface=<dev>      Serial Device
    -b --baud=<baud>          Baudrate [default: 9600]
    -a --address_code=<addr>  Address code of device [default: 1]
    -d --delay=<delay>        Delay for polling of values in milliseconds. Set to 0 to disable polling. [default: 1000]
    -o --output=<output>      Filename to output CSV data, will append if existing [default: power_log.csv]
    -c --comment=<comment>    Optional comment to be added to the data. (e.g. person doing the workout) [default: ]
"""

import os
import serial
import re
import time
import datetime
import csv

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



class DPM8600:

    START = ':'
    READ = 'r'
    WRITE = 'w'
    END = '\r\n'

    #WRITE_VOLTAGE = 10 # V/100
    #WRITE_CURRENT = 11 # mA
    #WRITE_OUTPUT_STATUS = 12 # output off (0), output on (1)
    #WRITE_VOLTAGE_AND_CURRENT = 20

    class Function(Enum):

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

        self.serial = serial.Serial(interface, baud_rate, timeout=5)
        try:
            self.serial.open()
        except serial.serialutil.SerialException:
            pass
        self.cmd_re = re.compile(f"{self.START}(?P<addr>\d+)(?P<func>[wr])(?P<func_num>\d+)=((?P<operand>\d+),?)+.?{self.END}")

    def __del__(self):
        self.serial.close()

    def _send(self, cmd, operands = [0]):
        cmd_name = self.Function(cmd).name

        if cmd_name.startswith("WRITE_"):
            func = self.WRITE
        else:
            func = self.READ

        operands_str = [str(o) for o in operands]
        operand = ",".join(operands_str)

        raw_cmd = f"{self.START}{self.address_code:02}{func}{cmd.value:02}={operand},{self.END}"

        self.serial.write(raw_cmd.encode())
        self.serial.flush()

    def _read(self, cmd):

        ret = self.serial.read_until()
        ret = ret.decode()

        if ret is None or len(ret) == 0:
            return None

        if not ret.endswith(self.END):
            print(f"TIMEOUT, in get_voltage() only got '{ret}'")
            return None
        m = self.cmd_re.match(ret)

        if m is None:
            print(f"Could not decode '{ret}'")
            return None

        if int(m.group("func_num")) != cmd.value:
            print(f"Wrong answer! got {m.group('func_num')} expected {cmd.value}")
            return None

        ret = cmd.convert(int(m.group("operand")))

        return ret

    def get_voltage(self):

        self._send(self.Function.READ_OUTPUT_VOLTAGE)
        return self._read(self.Function.READ_OUTPUT_VOLTAGE)

    def get_current(self):

        self._send(self.Function.READ_OUTPUT_CURRENT)
        return self._read(self.Function.READ_OUTPUT_CURRENT)

    def get_output_status(self):

        self._send(self.Function.READ_OUTPUT_STATUS)
        return self._read(self.Function.READ_OUTPUT_STATUS)

    def get_output_type(self):

        self._send(self.Function.READ_OUTPUT_TYPE)
        return self._read(self.Function.READ_OUTPUT_TYPE)

    def get_temperature(self):

        self._send(self.Function.READ_TEMPERATURE)
        return self._read(self.Function.READ_TEMPERATURE)

if __name__ == "__main__":
    args = docopt(__doc__, version="0.1")

    if args["--interface"] is None:
        args["--interface"] = DEFAULT_INTERFACE
    print(args)

    dev =  DPM8600(address_code = int(args["--address_code"]),
                   baud_rate = int(args["--baud"]),
                   interface = args["--interface"])

    last_v = None
    last_a = None
    wh_sum = 0.0
    delay_seconds = int(args["--delay"]) / 1000.0

    comment = args["--comment"]

    with open(args["--output"],'a') as csvfile:
        fieldnames = ['timestamp', 'realtime', 'V', 'A', 'W', 'Wh', 'Wh Sum','comment']

        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        # only write header once
        if csvfile.tell() == 0:
            writer.writeheader()
        while True:
            v = dev.get_voltage()
            a = dev.get_current()
            state = dev.get_output_status()
            typ = dev.get_output_type()
            temp = dev.get_temperature()

            w = v*a

            wh = w * (delay_seconds / 3600.0)

            now = datetime.datetime.now()
            
            print(f"[{now.strftime('%Y-%m-%d %H:%M:%S.%f')}] {v:.2f} V, {a:.3f} A, {w:.3f} W, {wh_sum:.3f} Wh")
            d = {'realtime': now.strftime('%Y-%m-%d %H:%M:%S.%f'),
                 'timestamp': now.strftime('%s%f'),
                 'V': v,
                 'A': a,
                 'W': w,
                 'Wh': wh,
                 'Wh Sum': wh_sum,
                 'comment': comment
                 }
            writer.writerow(d)

            time.sleep(delay_seconds)

            last_v = v
            last_a = a
            wh_sum += wh

