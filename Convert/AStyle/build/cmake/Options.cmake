# Build Options - executable by default, libraries on request
option(BUILD_JAVA_LIBS   "Build java library"   OFF)
option(BUILD_SHARED_LIBS "Build shared library" OFF)
option(BUILD_STATIC_LIBS "Build static library" OFF)

# Release Build by default (except for Borland)
if(NOT CMAKE_BUILD_TYPE)
    set(CMAKE_BUILD_TYPE "Release")
endif()

# Linux Soname Version
set(MAJORVER 3)
set(MINORVER 5)
set(PATCHVER 0)
set(SOLIBVER ${MAJORVER}.${MINORVER}.${PATCHVER})

# AStyle Source
file(GLOB SRCS src/*.cpp)

# AStyle Documentation, selected files
list(APPEND DOCS
    doc/astyle.html
    doc/install.html
    doc/news.html
    doc/notes.html
    doc/styles.css)

list(APPEND MAN
    man/astyle.1)

# Define java as a shared library
 if(BUILD_JAVA_LIBS)
    set(BUILD_SHARED_LIBS ON)
 endif()

 # Define the output type
 if(BUILD_SHARED_LIBS OR BUILD_STATIC_LIBS)
    add_library(astyle ${SRCS})
else()
    add_executable(astyle ${SRCS})
endif()
