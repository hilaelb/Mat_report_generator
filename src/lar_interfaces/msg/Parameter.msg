# param_id                  : Onboard parameter id, terminated by NULL if the length is less than 16 human-readable chars and WITHOUT null termination (NULL) byte if the length is exactly 16 chars - applications have to provide 16+1 bytes storage if the ID is stored as string (type:char)
# param_value               : Onboard parameter value (type:float)
# param_type                : Onboard parameter type: see the MAV_PARAM_TYPE enum for supported data types. (type:uint8_t, values:MAV_PARAM_TYPE)
# param_count               : Total number of onboard parameters (type:uint16_t)
# param_index               : Index of this onboard parameter (type:uint16_t)

string  param_id
float32 param_value
int8    param_type
int16   param_count
int16   param_index