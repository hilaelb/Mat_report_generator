# timestamp                 : Epoch time [ms]
# timestartup               : Time since sensor is running [ns]
# VpeStatus                 : 
# yawpitchroll              : Three variable vector giving Yaw Pitch Roll
# quaternion                : 
# dcm                       : The estimated attitude directional cosine matrix given in column major order.
# magned                    : The current estimated magnetic field (Gauss), given in the North East Down (NED) frame.
# accelned                  : The estimated acceleration (with gravity) reported, given in the North East Down (NED) frame. [m/s^2]
# linearaccelbody           : The estimated linear acceleration (without gravity) reported, given in the body frame. [m/s^2]
# linearaccelned            : The estimated linear acceleration (without gravity) reported, given in the North East Down frame. [m/s^2]




int64                       timestamp
uint64                      timestartup
# VpeStatus                   vpestatus         
geometry_msgs/Vector3       yawpitchroll      
# geometry_msgs/Quaternion    quaternion       
# float32[9]                  dcm            
geometry_msgs/Vector3       magned         
geometry_msgs/Vector3       accelned         
# geometry_msgs/Vector3       linearaccelbody   
# geometry_msgs/Vector3       linearaccelned   

