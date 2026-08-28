# Set linker options

# Strip release builds
if(CMAKE_BUILD_TYPE STREQUAL "Release" OR CMAKE_BUILD_TYPE STREQUAL "MinSizeRel")
	if(NOT BUILD_STATIC_LIBS AND NOT BORLAND)
		add_custom_command(TARGET astyle POST_BUILD
						   COMMAND ${CMAKE_STRIP} $<TARGET_FILE_NAME:astyle>)
	endif()
endif()
# Shared library options
if(BUILD_SHARED_LIBS)
	set(CMAKE_SHARED_LIBRARY_LINK_CXX_FLAGS "")     # remove -rdynamic
	if(${CMAKE_CXX_COMPILER_ID} STREQUAL "Intel")
		# need static linking because cannot find the intel libraries at runtime
		set(CMAKE_SHARED_LINKER_FLAGS "${CMAKE_SHARED_LINKER_FLAGS} -static-intel")
	elseif(MINGW)
		# minGW dlls don't work
		# tdm-gcc dlls work with everything except python
		set(CMAKE_SHARED_LINKER_FLAGS
			"${CMAKE_SHARED_LINKER_FLAGS} -Wl,--add-stdcall-alias -Wl,--dll")
	elseif(BORLAND)
		# use a development environment to compile Borland dlls
	endif()
endif()
