cmake_minimum_required(VERSION 3.12)

add_executable(test_docx WIN32 main.cpp)

target_link_libraries(test_docx PRIVATE docxlib)