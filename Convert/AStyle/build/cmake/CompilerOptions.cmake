# Set default compile options for supported compilers
if(APPLE)
    target_compile_options( astyle PRIVATE -W -Wall -fno-rtti -fno-exceptions -std=c++17 -stdlib=libc++)
elseif(NOT WIN32)   # Linux
    target_compile_options(astyle PRIVATE -Wall -Wextra -fno-rtti -fno-exceptions -std=c++17)
    if ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "gnu")
        execute_process(COMMAND ${CMAKE_CXX_COMPILER} -dumpversion OUTPUT_VARIABLE GCC_VERSION)
        if (NOT (GCC_VERSION VERSION_GREATER 4.6 OR GCC_VERSION VERSION_EQUAL 4.6))
            message(FATAL_ERROR "${PROJECT_NAME} c++11 support requires g++ 4.6 or greater, but it is ${GCC_VERSION}")
        endif ()
	elseif(${CMAKE_CXX_COMPILER_ID} STREQUAL "Intel" AND CMAKE_BUILD_TYPE STREQUAL "Release")
        # remove intel remarks for release build
        target_compile_options(astyle PRIVATE -wd11074,11076)
    endif()
elseif(MINGW)
    target_compile_options(astyle PRIVATE -Wall -Wextra -fno-rtti -fno-exceptions -std=c++17)
elseif(BORLAND)     # Release must be explicitly requested for Borland
    target_compile_options(astyle PRIVATE -q -w -x-)   # Cannot use no-rtti (-RT-)
endif()

# Set build-specific compile options
if(BUILD_SHARED_LIBS OR BUILD_STATIC_LIBS)
    if(BUILD_JAVA_LIBS)
        find_package(JNI)
        if (NOT JNI_FOUND)
            set(err_jni "Use '-DJAVA_HOME=<JNI Directory>' to define the Java path")
            message(FATAL_ERROR ${err_jni})
        endif()
        target_compile_options(astyle PRIVATE -DASTYLE_JNI)
        target_include_directories(astyle PRIVATE ${JNI_INCLUDE_DIRS})
    else()
        target_compile_options(astyle PRIVATE -DASTYLE_LIB)
    endif()
    # Windows DLL exports removed
    set_property(TARGET astyle PROPERTY DEFINE_SYMBOL "")
    # Linux solib version added
    set_property(TARGET astyle PROPERTY VERSION ${SOLIBVER})
    set_property(TARGET astyle PROPERTY SOVERSION ${MAJORVER})
endif()
