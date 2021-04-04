set ( EXE test_fortran-excel )

set ( SRCS 
  ./basic.f90
  ../../src/excel.c
  ../../src/excel_mo.f90
  )

set ( LIBS
  -lxlsxwriter
  -L/usr/local/lib
  )

set ( CMAKE_Fortran_FLAGS "-ffree-line-length-none" )
set ( CMAKE_Fortran_FLAGS_DEBUG "-Ddebug -O0 -fbounds-check -Wall" )
set ( CMAKE_Fortran_FLAGS_RELEASE "-Drelease -O3" )
set ( CMAKE_C_FLAGS "-O")

add_executable ( ${EXE} ${SRCS} )

target_link_libraries ( ${EXE} ${LIBS} )

include_directories ( /usr/local/include )

add_test ( NAME ${EXE} COMMAND ${EXE} ${ARGS} WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR} )

install ( TARGETS ${EXE} DESTINATION bin )
