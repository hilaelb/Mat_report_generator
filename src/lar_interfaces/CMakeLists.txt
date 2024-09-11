cmake_minimum_required(VERSION 3.8)
project(lar_interfaces)

# Default to C++14
if(NOT CMAKE_CXX_STANDARD)
  set(CMAKE_CXX_STANDARD 14)
endif()

if(CMAKE_COMPILER_IS_GNUCXX OR CMAKE_CXX_COMPILER_ID MATCHES "Clang")
  add_compile_options(-Wall -Wextra -Wpedantic)
endif()

# find dependencies - all of this dependencies are very importent!
find_package(ament_cmake REQUIRED)
find_package(ament_cmake_python REQUIRED)
find_package(rclcpp REQUIRED)
find_package(rclpy REQUIRED)

find_package(rosidl_default_generators REQUIRED)
find_package(builtin_interfaces REQUIRED)
find_package(std_msgs REQUIRED)
find_package(geometry_msgs REQUIRED)


rosidl_generate_interfaces(${PROJECT_NAME}
  "msg/TimeSample.msg"
  "msg/TimeUtc.msg"
  "msg/VpeStatus.msg"
  "msg/BmeData.msg"
  "msg/BaData.msg"
  "msg/AdcData.msg"
  "msg/Vector3.msg"
  "msg/Imu.msg"
  "msg/GnssVn200.msg"
  "msg/InsVn200.msg"
  "msg/AttVn200.msg"
  "msg/TigerMode.msg"
  "msg/QgcMode.msg"
  "msg/Waypoint.msg"
  "msg/Parameter.msg"
  "msg/Joystick.msg"
  "msg/JetMotorData.msg"
  "msg/ServoData.msg"
  "msg/MotorGuidance.msg"
  "msg/GuidanceData.msg"
  "msg/PidLog.msg"
  "msg/YprData.msg"
  "msg/Status.msg"
  "msg/Vn200Data.msg"
  "msg/Vn200DataNoFix.msg"
  "msg/InsStatus.msg"
  "msg/WaypointReached.msg"
  "srv/SendBmeData.srv"
  "srv/SendBaData.srv"
  "srv/SendAdcData.srv"
  "srv/SendBmiData.srv"
  "srv/SendGPSData.srv"
  "srv/SendLoraData.srv"
  "srv/GetYprData.srv"
  "srv/GetParameters.srv"
  "srv/SetParameter.srv"
  "srv/Parameters.srv"
  DEPENDENCIES builtin_interfaces std_msgs geometry_msgs             # This line is very importent for costume messages 
  ADD_LINTER_TESTS
)


ament_package()
