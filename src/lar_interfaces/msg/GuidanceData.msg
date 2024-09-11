# timestamp                       : Epoch time [ms]
# destbearing                     : The destination to the next waypoint from current position [deg]
# deisttodest                     : The distance between current location and next waypoint [m]
# headingerror                    : The difference between current heading and desired heading [deg]
# deptherror                      : the difference between current depth as read by BA30 and waypoint depth [m]
# headingcmd                      : Heading command sent to servo motors [deg]
# depthcmd                        : Depth command sent to pitch PID 
# pitcherror                      : The difference between current pitch and desired pitch [deg]
# pitchcmd                        : Pitch command sent to the servo motors [deg]
# pitch                           : Current pitch [deg]
# yaw                             : Current yaw [deg]
# roll                            : Current roll [deg]




int64         timestamp
float32       destbearing
float32       disttodest
float32       headingerror
float32       deptherror
float32       headingcmd
float32       depthcmd
float32       pitcherror
float32       pitchcmd
float32       pitch
float32       yaw
float32       roll