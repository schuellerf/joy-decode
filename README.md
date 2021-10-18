# joy-decode

Library and helper program to interface
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
