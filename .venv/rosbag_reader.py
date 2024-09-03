# make sure you have : 
# rosbags module , if not : pip install rosbags
# scipy.io module, if not : python -m pip install scipy
# numpy module, if not : pip install numpy
from rosbags.rosbag2 import Reader
from rosbags.serde import deserialize_cdr
from rosbags.typesys import get_types_from_msg, register_types
from os import listdir
from pathlib import Path
from scipy.io import savemat
import numpy as np
import os
# Creating custom messages 
MSG_TYPE_LIST = listdir(str(Path.cwd()) + '\lar_interfaces\msg')
MY_PATH = str(Path.cwd())
print(MSG_TYPE_LIST)

# Creating custom msg
for msg_type in MSG_TYPE_LIST: 
    custom_msg_path = Path(str(Path.cwd()) + "\\lar_interfaces\\msg\\" + msg_type)
    x = custom_msg_path.read_text(encoding='utf-8')
    register_types(get_types_from_msg(
        x ,'lar_interfaces/msg/' + msg_type[:-4] ))
from rosbags.typesys.types import lar_interfaces__msg__BaData as BaDataMsg
BaData = BaDataMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__AdcData as AdcDataMsg
AdcData = AdcDataMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__BmeData as BmeDataMsg
BmeData = BmeDataMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__Vn200Data as Vn200DataMsg
Vn200Data = Vn200DataMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__InsStatus as InsStatusMsg
InsStatus = InsStatusMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__GnssVn200 as GnssVn200Msg
GnssVn200 = GnssVn200Msg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__JetMotorData as JetMotorDataMsg
JetMotorData = JetMotorDataMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__MotorGuidance as MotorGuidanceMsg
MotorGuidance = MotorGuidanceMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__Joystick as JoystickMsg
Joystick = JoystickMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__ServoData as ServoDataMsg
ServoData = ServoDataMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__TigerMode as TigerModeMsg
TigerMode =TigerModeMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__Parameter as ParameterMsg
Parameter = ParameterMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__Waypoint as WaypointMsg
Waypoint =WaypointMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__PidLog as PidLogMsg
PidLog = PidLogMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__Status as StatusMsg
Status = StatusMsg.__msgtype__
from rosbags.typesys.types import lar_interfaces__msg__QgcMode as QgcModeMsg
QgcMode = QgcModeMsg.__msgtype__
# until here


def read_bag(file_name):
    ''' This Function convert rosbag file to .mat file inside the folder of the rosbag file.'''
    bag_file = str(Path.cwd())+ '\\rosbag_files\\'+ file_name
    with Reader(bag_file) as reader:

        adc = {'timestamp': [],'ain0': [], 'ain1': [], 'ain2': [], 'ain3': [], 'ain4': [], 'ain5': [], 'ain6': [], 'ain7': [], 'ain8': [], 'ain9': [], 'ain10': [], 'ain11': []}

        ba30 = {'timestamp': [],'temperature': [], 'pressure': [], 'depth': []}

        bme_back = {'timestamp': [], 'temperature': [], 'pressure': [], 'humidity': []}

        bme_front = {'timestamp': [], 'temperature': [], 'pressure': [], 'humidity': []}

        vn = {'timestamp': [],"time_startup": [], "yaw": [], "pitch": [], "roll": [], "poslla_x": [], "poslla_y": [], "poslla_z": [], "velned_x": [], "velned_y": [], "velned_z": [], "posu": [], "velu": []}

        joystick = {'timestamp': [],"ch1": [], "ch2": [], "ch3": [], "ch4": [], "ch5": [], "ch6": [], "ch7": [], "ch8": []}

        jet = {'timestamp': [],"avg_motor_current": [], "rpm": [], "temp_fet": [], "temp_motor": [], "v_in": []}

        mogu = {'timestamp': [],"srv_1_angle": [], "srv_2_angle": [], "srv_3_angle": [], "srv_4_angle": [], "jet_rpm": []}

        param = {"id": [], "value": [], "type": [], "count": [], "index": []}

        servo = {'timestamp': [],"data_valid": [], "srv1_angle": [], "srv1_voltage": [], "srv1_temperature": [], "srv2_angle": [], "srv2_voltage": [], "srv2_temperature": [], "srv3_angle": [], "srv3_voltage": [], "srv3_temperature": [], "srv4_angle": [], "srv4_voltage": [], "srv4_temperature": []}

        status = {'timestamp': [],"code": []}

        tiger = {'timestamp': [],"tiger_mode": []}

        qgc = {"mode": []}

        waypoint = {"seq": [], "command": [], "alt_depth": [], "stay_at_pos_time": [], "uncertainty_radius": [], "tool_speed": [], "latitude": [], "longitude": [], "altitude": []}

        insstatus = {'timestamp': [],"ins_mode": [], "ins_error": [], "ins_fix": []}

        gnssVn200 = {'timestamp': [],'numsats': [], 'fix': [], 'poslla_x': [], 'poslla_y': [], 'poslla_z': []}

        pid_heading = {'timestamp': [], 'accumulator' : [], 'preverror' : [], 'identifier' : []}
        
        pid_depth = {'timestamp': [], 'accumulator' : [], 'preverror' : [], 'identifier' : []}

        pid_pitch = {'timestamp': [], 'accumulator' : [], 'preverror' : [], 'identifier' : []}

        print('Topics + Message Types : ' )
        for connection in reader.connections:
            print('Topic : ', connection.topic,' | Message Type : ' ,connection.msgtype)
        # iterate over messages
        for connection, Time_Stamp, rawdata in reader.messages():
            # ADC - CHECKED ! 
            if connection.topic == '/ADC_measurements':
                msg = deserialize_cdr(rawdata, AdcData)
                adc["timestamp"].append(msg.timestamp)
                adc['ain0'].append(msg.v_24)
                adc['ain1'].append(msg.v_12)
                adc['ain2'].append(msg.v_5)
                adc['ain3'].append(msg.v_3)
                adc['ain4'].append(msg.i_24)
                adc['ain5'].append(msg.i_12)
                adc['ain6'].append(msg.i_5)
                adc['ain7'].append(msg.i_3)
                adc['ain8'].append(msg.i_in)
                adc['ain9'].append(msg.v_in)
                adc['ain10'].append(msg.ain_10)
                adc['ain11'].append(msg.ain_11)
            # BA30 - CHECKED ! 
            elif connection.topic == '/BA30_measurements':
                msg = deserialize_cdr(rawdata, BaData)
                ba30["timestamp"].append(msg.timestamp)
                ba30['temperature'].append(msg.temperature)
                ba30['pressure'].append(msg.pressure)
                ba30['depth'].append(msg.depth)
            # BME Front - CHECKED !
            elif connection.topic == '/BME_Front_measurements':
                msg = deserialize_cdr(rawdata, BmeData)
                bme_front["timestamp"].append(msg.timestamp)
                bme_front["humidity"].append(msg.humidity) 
                bme_front["temperature"].append(msg.temperature)
                bme_front["pressure"].append(msg.pressure)
            # BME Back - CHECKED ! 
            elif connection.topic == '/BME_Back_measurements':
                msg = deserialize_cdr(rawdata, BmeData)
                bme_back["timestamp"].append(msg.timestamp)
                bme_back["humidity"].append(msg.humidity) 
                bme_back["temperature"].append(msg.temperature)
                bme_back["pressure"].append(msg.pressure)
            # Waypoint - CHECKED ! 
            elif connection.topic == '/waypoint':
                msg = deserialize_cdr(rawdata, Waypoint)
                waypoint["seq"].append(msg.seq)
                waypoint["command"].append(msg.command)
                waypoint["alt_depth"].append(msg.alt_depth)
                waypoint["stay_at_pos_time"].append(msg.stay_at_pos_time)
                waypoint["uncertainty_radius"].append(msg.uncertainty_radius)
                waypoint["tool_speed"].append(msg.tool_speed)
                waypoint["latitude"].append(msg.latitude)
                waypoint["longitude"].append(msg.longitude)
                waypoint["altitude"].append(msg.altitude)
            # Joystick - CHECKED ! 
            elif connection.topic == '/joystick_data':
                msg = deserialize_cdr(rawdata, Joystick)
                joystick["timestamp"].append(msg.timestamp)
                joystick["ch1"].append(msg.chan1)
                joystick["ch2"].append(msg.chan2)
                joystick["ch3"].append(msg.chan3)
                joystick["ch4"].append(msg.chan4)
                joystick["ch5"].append(msg.chan5)
                joystick["ch6"].append(msg.chan6)
                joystick["ch7"].append(msg.chan7)
                joystick["ch8"].append(msg.chan8)
            # QGC MODE - CHECKED ! 
            elif connection.topic == '/qgc_mode':
                msg = deserialize_cdr(rawdata, QgcMode)
                qgc["mode"].append(msg.qgc_mode)
            # Servo - CHECKED ! 
            elif connection.topic == '/servo_motor_data':
                msg = deserialize_cdr(rawdata, ServoData)
                servo["timestamp"].append(msg.timestamp)
                servo["data_valid"].append(msg.servo_data_valid)
                servo["srv1_angle"].append(msg.srv1_angle)
                servo["srv1_voltage"].append(msg.srv1_voltage)
                servo["srv1_temperature"].append(msg.srv1_temperature)
                servo["srv2_angle"].append(msg.srv2_angle)
                servo["srv2_voltage"].append(msg.srv2_voltage)
                servo["srv2_temperature"].append(msg.srv2_temperature)
                servo["srv3_angle"].append(msg.srv3_angle)
                servo["srv3_voltage"].append(msg.srv3_voltage)
                servo["srv3_temperature"].append(msg.srv3_temperature)
                servo["srv4_angle"].append(msg.srv4_angle)
                servo["srv4_voltage"].append(msg.srv4_voltage)
                servo["srv4_temperature"].append(msg.srv4_temperature)
            # MotorGuidanceData - CHECKED ! 
            elif connection.topic == '/motor_guidance_data':
                msg = deserialize_cdr(rawdata, MotorGuidance)
                mogu["timestamp"].append(msg.timestamp)
                mogu["srv_1_angle"].append(msg.srv_1_angle)
                mogu["srv_2_angle"].append(msg.srv_2_angle)
                mogu["srv_3_angle"].append(msg.srv_3_angle)
                mogu["srv_4_angle"].append(msg.srv_4_angle)
                mogu["jet_rpm"].append(msg.jet_rpm)
            # JetMotorData - CHECKED ! 
            elif connection.topic == '/jet_motor_data':
                msg = deserialize_cdr(rawdata, JetMotorData)
                jet["timestamp"].append(msg.timestamp)
                jet["avg_motor_current"].append(msg.avg_motor_current)
                jet["rpm"].append(msg.rpm)
                jet["temp_fet"].append(msg.temp_fet)
                jet["temp_motor"].append(msg.temp_motor)
                jet["v_in"].append(msg.v_in)
            # tiger mode - CHECKED ! 
            elif connection.topic == '/tiger_mode':
                msg = deserialize_cdr(rawdata, TigerMode)
                tiger["timestamp"].append(msg.timestamp)
                tiger["tiger_mode"].append(msg.tiger_mode)
            # parameter - CHECKED ! 
            elif connection.topic == '/parameter':
                msg = deserialize_cdr(rawdata, Parameter)
                param["id"].append(msg.param_id)
                param["value"].append(msg.param_value)
                param["type"].append(msg.param_type)
                param["count"].append(msg.param_count)
                param["index"].append(msg.param_index)
            #GnssVn200 - CHECKED ! 
            elif connection.topic == '/GnssVn200':
                msg = deserialize_cdr(rawdata, GnssVn200)
                gnssVn200["timestamp"].append(msg.timestamp)
                gnssVn200["numsats"].append(msg.numsats)
                gnssVn200["fix"].append(msg.fix)
                gnssVn200["poslla_x"].append(msg.poslla.x)
                gnssVn200["poslla_y"].append(msg.poslla.y)
                gnssVn200["poslla_z"].append(msg.poslla.z)
            # InsStatus - CHECKED ! 
            elif connection.topic == '/InsStatus':
                msg = deserialize_cdr(rawdata, InsStatus)
                insstatus["timestamp"].append(msg.timestamp)
                insstatus["ins_mode"].append(msg.ins_mode)
                insstatus["ins_error"].append(msg.ins_error)
                insstatus["ins_fix"].append(msg.ins_fix)
            # VN200 - CHECKED ! 
            elif connection.topic == '/Vn200Data':
                msg = deserialize_cdr(rawdata, Vn200Data)
                vn["timestamp"].append(msg.timestamp)
                vn["time_startup"].append(msg.timestartup)
                vn["yaw"].append(msg.yawpitchroll.x)
                vn["pitch"].append(msg.yawpitchroll.y)
                vn["roll"].append(msg.yawpitchroll.z)
                vn["poslla_x"].append(msg.poslla.x)
                vn["poslla_y"].append(msg.poslla.y)
                vn["poslla_z"].append(msg.poslla.z)
                vn["velned_x"].append(msg.velned.x)
                vn["velned_y"].append(msg.velned.y)
                vn["velned_z"].append(msg.velned.z)
                vn["posu"].append(msg.posu)
                vn["velu"].append(msg.velu)
            # PID - CHECKED ! 
            elif connection.topic == 'pid_data_log':
                msg = deserialize_cdr(rawdata,PidLog)
                if msg.identifier == 'heading':
                    pid_heading["timestamp"].append(msg.timestamp)
                    pid_heading["accumulator"].append(msg.accumulator)
                    pid_heading["preverror"].append(msg.preverror)
                    pid_heading["identifier"].append(msg.identifier)
                elif msg.identifier == 'depth':
                    pid_depth["timestamp"].append(msg.timestamp)
                    pid_depth["accumulator"].append(msg.accumulator)
                    pid_depth["preverror"].append(msg.preverror)
                    pid_depth["identifier"].append(msg.identifier)
                elif msg.identifier == 'pitch':
                    pid_pitch["timestamp"].append(msg.timestamp)
                    pid_pitch["accumulator"].append(msg.accumulator)
                    pid_pitch["preverror"].append(msg.preverror)
                    pid_pitch["identifier"].append(msg.identifier)

        data = {'adc': adc, 'ba30': ba30, 'bme_front': bme_front, 'bme_back': bme_back, 'vn': vn, 'joystick': joystick, 'jet': jet, 'mogu': mogu, 'param': param, 'servo': servo, 'status': status, 'tiger': tiger, 'qgc': qgc, 'waypoint': waypoint, 'insstatus': insstatus, 'gnssVn200': gnssVn200}
        savemat(bag_file+"\\"+file_name+".mat", data)

ROSBAG_NAMES = listdir(str(Path.cwd()) + '\\rosbag_files')
# print(ROSBAG_NAMES)
# print(str(Path.cwd()))
# ROSBAG_FILE = str(input('ROSBAG FILE:'))
# # read_bag('rosbag2_2022_12_20-11_10_57')
# read_bag(ROSBAG_FILE)
for rosbag_file in ROSBAG_NAMES:
    print(rosbag_file)
    read_bag(rosbag_file)

