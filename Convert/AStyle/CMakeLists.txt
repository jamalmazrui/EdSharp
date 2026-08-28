cmake_minimum_required(VERSION 3.8)

if(CMAKE_VERSION VERSION_GREATER "3.19.0")
  cmake_policy(SET CMP0112 OLD)
endif()

project(astyle CXX)

include( build/cmake/Options.cmake REQUIRED )
include( build/cmake/CompilerOptions.cmake REQUIRED )
include( build/cmake/LinkerOptions.cmake REQUIRED )
include( build/cmake/InstallOptions.cmake REQUIRED )
include( build/cmake/OutputFiles.cmake REQUIRED )
include( build/cmake/PrintInfo.cmake REQUIRED )
