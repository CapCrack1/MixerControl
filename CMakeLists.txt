cmake_minimum_required(VERSION 3.10)
project(MixerControl)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_executable(MixerControl mixer_control.cpp)
target_link_libraries(MixerControl PRIVATE ws2_32 ole32)