# This is a message to publish Tiger mode
#
# Tiger_IDLE = 1                  # idle (pause button)
# Tiger_JOYSTICK = 2              # manual control by joystick (pause)
# Tiger_FIXED_MANUVER = 4         # doing one fixed manuver based on param (pause)
# Tiger_HOLD_POSITION = 8         # stopes go to home (stop button)
# Tiger_ABORTED_ONLY_MANUAL = 16  # stop and dont use the hold position (stop)
# Tiger_ABORTED_LPM = 32          # stoped (no battery) only enable comm, NO MOTORS (stop)
# Tiger_CALIBRATION = 64          # in magnetometer calibration mode (no used)
# Tiger_GUIDED_MISSION = 128      # in wp mission mode mode (play)

int64 timestamp
int64 tiger_mode