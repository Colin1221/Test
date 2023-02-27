
# please first pip install modbus_tk, using 'pip install modbus_tk'
# -*- coding: utf-8 -*-

"""
Created 01/30/2023
@author: ColinChu
"""

import modbus_tk
import modbus_tk.defines as cst
import modbus_tk.modbus_tcp as modbus_tcp
from modbus_tk import utils
import time
import csv
import datetime
from datetime import datetime
import pytz
import threading
import pyvisa as visa
# Tesla Python #
from uds_client_application import Uds_Client_Application
import canapi2.dll_utils
logger = modbus_tk.utils.create_logger("console")


class SmartChamber:
    def __init__(self, ip_host, port=5000):
        self.master = modbus_tcp.TcpMaster(ip_host, port)
        self.master.set_timeout(5.0)
        logger.info("connected")

    def get_temperature(self):
        temperature = self.master.execute(1, cst.READ_HOLDING_REGISTERS, cst.READ_CURR_TEMP, 1)[0]
        if temperature > 3000:
            temperature = -65535 + temperature - 1
        return temperature / 10

    def get_humidity(self):
        humidity = self.master.execute(1, cst.READ_HOLDING_REGISTERS, cst.READ_CURR_HUMI, 1)
        return humidity[0] / 10

    def read_set_temperature(self):
        set_temperature = self.master.execute(1, cst.READ_HOLDING_REGISTERS, cst.SET_TEMP, 1)[0]
        if set_temperature > 3000:
            set_temperature = -65535 + set_temperature - 1

        return set_temperature / 10

    def read_set_humidity(self):
        set_humid = self.master.execute(1, cst.READ_HOLDING_REGISTERS, cst.SET_HUMI, 1)
        return set_humid[0] / 10

    def set_temperature(self, temp):
        if chamber.read_set_temperature() != temp:
            if temp > 155:
                self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.SET_TEMP, output_value=155 * 10)
            elif temp < -75:
                self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.SET_TEMP, output_value=-75 * 10)
            else:
                self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.SET_TEMP, output_value=temp * 10)
        # may stuck at this point
        # time.sleep(0.1)
        # temp1=temp
        # if temp<0:
        #     temp1=(65535+temp*10+1)/10
        # while self.master.execute(1, cst.READ_HOLDING_REGISTERS, cst.SET_TEMP, 1)[0]/10!=temp1:
        #     time.sleep(0.1)
        #     self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.SET_TEMP, output_value=temp*10)

        return self.read_set_temperature()

    def set_humidity(self, humidity):
        if chamber.read_set_humidity() != humidity:
            if humidity > 100:
                self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.SET_HUMI, output_value=100 * 10)
            elif humidity < 0:
                self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.SET_HUMI, output_value=0)
            else:
                self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.SET_HUMI, output_value=humidity * 10)
        # may stuck at this point
        # time.sleep(0.1)
        # while self.master.execute(1, cst.READ_HOLDING_REGISTERS, cst.SET_HUMI, 1)[0] / 10 != humi:
        #     time.sleep(0.1)
        #     self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.SET_HUMI, output_value=humi * 10)
        return self.read_set_humidity()

    def chamber_on_off_state(self, state_request: bool):
        if state_request:
            self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.TURN_ON_OFF, output_value=1)
        else:
            self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.TURN_ON_OFF, output_value=0)

    def read_chamber_state(self):
        state_request = self.master.execute(1, cst.READ_HOLDING_REGISTERS, cst.TURN_ON_OFF, 1)
        return state_request[0]

    def chamber_light(self, state: bool):
        if state:
            self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.ChamberLightSwitch, output_value=1)
        elif state:
            self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.ChamberLightSwitch, output_value=0)
        return state

    def read_chamber_light(self):
        light_request = self.master.execute(1, cst.READ_HOLDING_REGISTERS, cst.ChamberLightSwitch, 1)
        return light_request[0]

    def set_humidity_slope(self, delta_humid: float):
        """
            RH% change per minute
        """
        if delta_humid > 5:
            self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.HUMI_SLOPE, output_value=5 * 10)
        else:
            self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.HUMI_SLOPE, output_value=delta_humid * 10)

    def read_humidity_slope(self):
        humi_slope = self.master.execute(1, cst.READ_HOLDING_REGISTERS, cst.HUMI_SLOPE, 1)
        return humi_slope[0] / 10

    def set_temp_slope(self, delta_temp: float):
        """
              temp change per minute
        """
        if delta_temp > 10:
            self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.TEMP_SLOPE, output_value=10 * 10)
        else:
            self.master.execute(1, cst.WRITE_SINGLE_REGISTER, cst.TEMP_SLOPE, output_value=delta_temp * 10)

    def read_temp_slope(self):
        temp_slope = self.master.execute(1, cst.READ_HOLDING_REGISTERS, cst.TEMP_SLOPE, 1)
        return temp_slope[0] / 10


def connect_hvc():
    """
    Set up comms with HVC through PCAN-USB Pro
    """
    global hvp, hvbms
    link = canapi2.canapi2_link.Canapi2LinkHwHandle('16', 500000)
    link.reset_hardware()
    link.open()
    hvp = Uds_Client_Application(link, 'HVP')
    hvp.get_security_access()
    hvp.session_control(3)

    link1 = canapi2.canapi2_link.Canapi2LinkHwHandle('15', 500000)
    link1.reset_hardware()
    link1.open()
    hvbms = Uds_Client_Application(link1, 'HVBMS')
    hvbms.get_security_access()
    hvbms.session_control(3)
    # USAGEID = hvbms.read_data(DID.USAGE)
    # print("usageid:", USAGEID)


def init_hv_source(key):
    """Connect and initialize HV_source.
       Directory path to visa32.dll so it can be used by pyVISA
    """
    rm = visa.ResourceManager('C:\\Windows\\System32\\visa32.dll')  # use PyVisa
    resource_list = rm.list_resources()
    # print(resource_list)
    for item in resource_list:
        try:
            access = rm.open_resource(item, open_timeout=visa.constants.VI_TMO_INFINITE)
        except visa.errors.VisaIOError:
            continue
        return_text = access.query('*IDN?')
        if key in return_text:
            break
    return access


def set_hv_source():
    hv_source = init_hv_source('N5772A')
    Current_set = hv_source.write("SOUR:CURR 1.5")
    voltage_set = hv_source.write("SOUR:VOLT 180")
    power_on = hv_source.write("OUTP ON")
    time.sleep(15)
    voltage_set = hv_source.write("SOUR:VOLT 220")
    time.sleep(15)
    voltage_set = hv_source.write("SOUR:VOLT 260")
    time.sleep(15)
    voltage_set = hv_source.write("SOUR:VOLT 310")
    time.sleep(15)
    power_off = hv_source.write("OUTP OFF")


def read_pack_voltage():
    result = hvbms.read_data(0x412).hex()  # hvbms did 0x412
    pack_pos = complement_conversion(int(result[:8], 16))
    pack_neg = complement_conversion(int(result[8:16], 16))
    pack_voltage = (pack_pos + pack_neg) / 1000
    print('pack voltage: {} V'.format(pack_voltage))
    return pack_voltage


def complement_conversion(num):
    if num > 2 ** 31:
        num = 2 ** 32 - num
    return num


def start_log_csv():
    global csv_handle, csv_writer
    date = datetime.now(tz=pytz.utc)
    now = date.astimezone(pytz.timezone('Asia/Shanghai'))
    header = ['time', 'packvoltage/V', 'chambertemp/℃', 'chamberhumi/%', 'settemp/℃', 'sethumi/%', 'chamberstate']
    csv_handle = open(f'Chamber_{now:%Y-%m-%d-%H-%M-%S}.csv', 'w', newline='')
    csv_writer = csv.writer(csv_handle)
    csv_writer.writerow(header)


def periodic_log_csv():
    date = datetime.now(tz=pytz.utc)
    now = date.astimezone(pytz.timezone('Asia/Shanghai'))
    csv_writer.writerow([now, pack_voltage, chamber_temp, chamber_humidity, set_temp, set_humidity, chamber_state])
    csv_handle.flush()


def chamber_on_off_with_hvc():
    """
    Control Chamber ON/OFF with variables from HVC
    """

    if pack_voltage < 200:
        chamber.chamber_on_off_state(False)

    elif 200 < pack_voltage < 290:
        chamber.chamber_on_off_state(True)

    elif pack_voltage > 300:
        chamber.chamber_on_off_state(False)


def set_chamber_with_hvc():
    """
    To control Chamber with with variables from HVC
    """
    if pack_voltage < 200:
        chamber.set_temperature(75)
        chamber.set_humidity(80)
        chamber.set_temp_slope(5)
        chamber.set_humidity_slope(8)

    elif 240 < pack_voltage < 290:
        chamber.set_temperature(60)
        chamber.set_humidity(60)

    elif pack_voltage > 300:
        chamber.set_temperature(80)
        chamber.set_humidity(80)


def periodic_task():
    global pack_voltage, chamber_temp, chamber_humidity, set_temp, set_humidity, chamber_state
    start_time = time.time()
    periodic_interval = 1
    pack_voltage = read_pack_voltage()
    chamber_on_off_with_hvc()  # switch on/off chamber with HVC
    set_chamber_with_hvc()  # Set chamber parameters with HVC
    chamber_state = chamber.read_chamber_state()
    set_humidity = chamber.read_set_humidity()
    chamber_temp = chamber.get_temperature()
    chamber_humidity = chamber.get_humidity()
    set_temp = chamber.read_set_temperature()
    periodic_log_csv()  # LOG CSV
    end_time = time.time()
    delta_time = end_time - start_time
    if periodic_interval > delta_time:
        time.sleep(periodic_interval - delta_time)
        print(delta_time)


if __name__ == "__main__":

    connect_hvc()  # connect HVC
    time.sleep(0.3)
    # start a new thread to set hv_source N5772A
    set_hv = threading.Thread(target=set_hv_source)
    set_hv.start()
    chamber = SmartChamber("192.168.1.101", 5000)  # connect RD chamber
    logger.info(chamber.chamber_light(True))
    start_log_csv()
    logger.info(threading.current_thread())
    while 1:
        periodic_task()
        if chamber.read_chamber_light() == 0:
            chamber.chamber_on_off_state(False)
            break
