# seq - waypoints order
# command - 16: WP_POINT, 19: WP_HOVERING (time loiter in qgc), 21: WP_GPS, 22: WP_SAEFTY, 25: WP_HOME (check this number)
# alt_depth - 1: altitude, 0: depth
# stay_at_pos_time - param1 - time to stay at this position before advancing
# uncertainty_radius - param2 - uncertainty radius in meters
# tool_speed - param4 - tool speed 
# latitude - x - latitude in degrees
# longitude - y - longitude in degrees
# altitude - z - altitude in meters

int64 seq
int64 command
int64 alt_depth
float64 stay_at_pos_time
float64 uncertainty_radius
float64 tool_speed
float64 latitude
float64 longitude
float64 altitude