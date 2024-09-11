# timestamp        : Epoch time [ms]
# servo_data_valid : True or False indication to understand if the servo motors are still connected and running
# srv"x"_angle     : Angle for all 4 servo motors with x equals [1:4] [Degrees]
# srv"x"_voltage   : Voltage for all 4 servo motors with x equals [1:4] [V] 

int64              timestamp
bool               servo_data_valid
float32            srv1_angle
float32            srv1_voltage
float32            srv1_temperature
float32            srv2_angle
float32            srv2_voltage
float32            srv2_temperature
float32            srv3_angle
float32            srv3_voltage
float32            srv3_temperature
float32            srv4_angle
float32            srv4_voltage
float32            srv4_temperature