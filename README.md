# EEMacros
This is a VBA enabled spreadsheet that contains functions for electrical engineering. The functions can be used to do things like
size overcurrent protection devices, size current carrying wires, grounds and conduits, output ampacities of wires, calculate voltage drop,
produce the full load amps for certain AC motors, output residential unit demand factors etc.

The functions are based off the 2008 NEC but do not always apply to every situation and do not encompass every table. 
If you are a professional electrical engineer using these functions to design electrical power systems, always compare the
output values to the NEC and other applicable codes in your jurisdiction. I do not assume any liability for the accuracy, relevancy, or applicability of
these functions.

Examples of each function are given in the upper left-hand corner of the spreadsheet.

Example of how to use:

If you want to return the full load amps of a 10 horsepower motor.

=FLA(10,480,3)
