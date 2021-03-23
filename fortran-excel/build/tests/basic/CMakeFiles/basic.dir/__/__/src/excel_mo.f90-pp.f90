# 1 "../src/excel_mo.f90"
# 1 "<built-in>"
# 1 "<command-line>"
# 1 "../src/excel_mo.f90"
module excel_mo

  use, intrinsic :: iso_c_binding

  implicit none

  type, bind(c) :: datetime_ty

    integer(c_int) :: year
    integer(c_int) :: month
    integer(c_int) :: day 
    integer(c_int) :: hour 
    integer(c_int) :: min 
    real(c_double) :: sec 

  end type

  interface

    subroutine workbook_new ( workbook, file ) &
        bind ( c, name = 'workbook_new_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: workbook
      character(c_char), intent(in) :: file
    end subroutine

    subroutine workbook_add_worksheet ( workbook, worksheet, name ) &
        bind ( c, name = 'workbook_add_worksheet_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: workbook, worksheet
      character(c_char), intent(in) :: name
    end subroutine

    subroutine workbook_add_format ( workbook, format ) &
        bind ( c, name = 'workbook_add_format_c' )
      import c_ptr
      type(c_ptr), intent(in) :: workbook, format
    end subroutine

    subroutine worksheet_write_datetime( workbook, row, col, datetime, format ) &
        bind ( c, name = 'worksheet_write_datetime_c' )
      import c_ptr, c_int, c_char, datetime_ty
      type(c_ptr),       intent(in)        :: workbook, format
      integer(c_int),    intent(in), value :: row, col
      type(datetime_ty), intent(in)        :: datetime
    end subroutine

    subroutine format_set_shrink ( format ) &
        bind ( c, name = 'format_set_shrink_c' )
      import c_ptr
      type(c_ptr), intent(in) :: format
    end subroutine

    subroutine format_set_bold ( format ) &
        bind ( c, name = 'format_set_bold_c' )
      import c_ptr
      type(c_ptr), intent(in) :: format
    end subroutine

    subroutine format_set_italic ( format ) &
        bind ( c, name = 'format_set_italic_c' )
      import c_ptr
      type(c_ptr), intent(in) :: format
    end subroutine

    subroutine worksheet_autofilter ( worksheet, first_row, first_col, last_row, last_col ) &
        bind ( c, name = 'worksheet_autofilter_c' )
      import c_ptr, c_int
      type(c_ptr), intent(in) :: worksheet
      integer(c_int), intent(in), value :: first_row, first_col, last_row, last_col
    end subroutine

    subroutine worksheet_set_row ( worksheet, format, row, height ) &
        bind ( c, name = 'worksheet_set_row_c' )
      import c_ptr, c_int
      type(c_ptr), intent(in) :: worksheet, format
      integer(c_int), intent(in), value :: row 
      integer(c_int), intent(in), value :: height
    end subroutine

    subroutine worksheet_set_column ( worksheet, format, col, width ) &
        bind ( c, name = 'worksheet_set_column_c' )
      import c_ptr, c_int
      type(c_ptr), intent(in) :: worksheet, format
      integer(c_int), intent(in), value :: col 
      integer(c_int), intent(in), value :: width
    end subroutine

    subroutine worksheet_write_formula ( worksheet, row, col, formula, format ) &
        bind ( c, name = 'worksheet_write_formula_c' )
      import c_ptr, c_int, c_char
      type(c_ptr), intent(in) :: worksheet, format
      integer(c_int),    intent(in), value :: row, col
      character(c_char), intent(in)        :: formula
    end subroutine

    subroutine worksheet_write_number ( worksheet, format, row, col, value ) &
        bind ( c, name = 'worksheet_write_number_c' )
      import c_ptr, c_int, c_double
      type(c_ptr), intent(in) :: worksheet, format
      integer(c_int), intent(in), value :: row, col
      real(c_double), intent(in), value :: value
    end subroutine

    subroutine worksheet_write_string ( worksheet, format, row, col, text ) &
        bind ( c, name = 'worksheet_write_string_c' )
      import c_ptr, c_int, c_char
      type(c_ptr),       intent(in)        :: worksheet, format
      integer(c_int),    intent(in), value :: row, col
      character(c_char), intent(in)        :: text
    end subroutine

    subroutine worksheet_insert_image ( worksheet, row, col, file ) &
        bind ( c, name = 'worksheet_insert_image_c' )
      import c_ptr, c_int, c_char
      type(c_ptr),       intent(in)        :: worksheet
      integer(c_int),    intent(in), value :: row, col
      character(c_char), intent(in)        :: file
    end subroutine

    subroutine workbook_close ( workbook ) &
        bind ( c, name = 'workbook_close_c' )
      import c_ptr
      type(c_ptr), intent(in) :: workbook
    end subroutine

    ! Font size
    subroutine format_set_font_size ( format, size ) &
        bind ( c, name = 'format_set_font_size_c' )
      import c_ptr, c_int
      type(c_ptr),    intent(in)        :: format
      integer(c_int), intent(in), value :: size
    end subroutine

    ! Font name
    subroutine format_set_font_name ( format, name ) &
        bind ( c, name = 'format_set_font_name_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: name
    end subroutine

    ! Number format
    subroutine format_set_num_format ( format, style ) &
        bind ( c, name = 'format_set_num_format_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: style
    end subroutine

    ! Font color
    subroutine format_set_font_color ( format, color ) &
        bind ( c, name = 'format_set_font_color_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: color
    end subroutine

    ! Pattern
    subroutine format_set_pattern ( format, pattern ) &
        bind ( c, name = 'format_set_pattern_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: pattern
    end subroutine

    ! Background color
    subroutine format_set_bg_color ( format, color ) &
        bind ( c, name = 'format_set_bg_color_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: color
    end subroutine

    ! Alignment
    subroutine format_set_align ( format, align ) &
        bind ( c, name = 'format_set_align_c' )
      import c_ptr, c_char
      type(c_ptr), intent(in) :: format
      character(c_char), intent(in) :: align
    end subroutine

    ! Border
    subroutine format_set_border ( format, style ) &
        bind ( c, name = 'format_set_border_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: style
    end subroutine

    subroutine format_set_top ( format, style ) &
        bind ( c, name = 'format_set_top_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: style
    end subroutine

    subroutine format_set_bottom ( format, style ) &
        bind ( c, name = 'format_set_bottom_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: style
    end subroutine

    subroutine format_set_right ( format, style ) &
        bind ( c, name = 'format_set_right_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: style
    end subroutine

    subroutine format_set_left ( format, style ) &
        bind ( c, name = 'format_set_left_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: style
    end subroutine

  end interface

  contains

  function cs ( fs ) result ( fsnull )

    character(*), intent(in)  :: fs
    character(:), allocatable :: fsnull

    fsnull = trim(fs)//c_null_char

  end function

end module
