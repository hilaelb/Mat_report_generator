# timestamp                 : Epoch time [ms]
# timestartup               : Time since sensor is running [ns]
# insstatus                 :
# poslla                    : A three variable float vector (latitude: degrees, longtitude: degrees, altitude: meters)
# velbody                   : A three variable float vector (velocity body x, velocity body y, velocity body z) [m/s]
# velned                    : A three variable float vector (velocity north, velocity east, velocity down) [m/s]
# posu                      : Position uncertainty [m]
# velu                      : Forward velocity uncertainty [m/s]




int64 timestamp
uint64 timestartup
uint16 insstatus
geometry_msgs/Point poslla
geometry_msgs/Vector3 velbody
geometry_msgs/Vector3 velned
float32 posu
float32 velu
