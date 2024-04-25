cmake_minimum_required(VERSION 3.0.0)
project(cpp-xlsx-demo VERSION 0.1.0 LANGUAGES C CXX)

include(CTest)
enable_testing()

add_executable(libxl_read_xlsx_demo libxl_read_xlsx_demo.cpp)
add_executable(libxl_write_xlsx_demo libxl_write_xlsx_demo.cpp utils.cpp utils.h)
add_executable(opxl_read_xlsx_demo opxl_read_xlsx_demo.cpp)
add_executable(opxl_write_xlsx_demo opxl_write_xlsx_demo.cpp utils.cpp utils.h)

include_directories(D:/libxl-4.3.0.14/include_cpp)
target_link_libraries(libxl_read_xlsx_demo D:/libxl-4.3.0.14/lib64/libxl.lib)
target_link_libraries(libxl_write_xlsx_demo D:/libxl-4.3.0.14/lib64/libxl.lib)
# target_link_libraries(libxl_write_xlsx_demo stdc++fs) # g++ 13不需要配置这个，8 配置了也没用

add_subdirectory(OpenXLSX)

target_link_libraries(opxl_read_xlsx_demo OpenXLSX::OpenXLSX)
target_link_libraries(opxl_write_xlsx_demo OpenXLSX::OpenXLSX)

set_target_properties(libxl_read_xlsx_demo PROPERTIES CXX_STANDARD 17)
set_target_properties(libxl_write_xlsx_demo PROPERTIES CXX_STANDARD 17)
set_target_properties(opxl_read_xlsx_demo PROPERTIES CXX_STANDARD 17)
set_target_properties(opxl_write_xlsx_demo PROPERTIES CXX_STANDARD 17)

set(CPACK_PROJECT_NAME ${PROJECT_NAME})
set(CPACK_PROJECT_VERSION ${PROJECT_VERSION})
include(CPack)
