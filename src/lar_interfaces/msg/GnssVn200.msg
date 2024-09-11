# timestamp                 : Epoch time [ms]
# numsats                   : Number of gnss satellites
# fix                       : Fix operating number [0: no fix, 1: Time only, 2: 2D, 3: 3D, 4: SBAS]
# poslla                    : A three variable float vector (latitude: degrees, longtitude: degrees, altitude: meters)





int64 timestamp
uint8 numsats
uint8 fix
geometry_msgs/Point poslla
