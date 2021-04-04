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

  type, bind(c) :: image_options_ty
    
    integer(c_int)    :: x_offset        = 0
    integer(c_int)    :: y_offset        = 0
    real(c_double)    :: x_scale         = 1.d0
    real(c_double)    :: y_scale         = 1.d0
    integer(c_int)    :: object_position = 0
    integer(c_int)    :: decorative      = 0
    character(c_char) :: description     = c_null_char
    character(c_char) :: url             = c_null_char
    character(c_char) :: tip             = c_null_char

  end type

  interface

    integer(c_int) function name2row ( name ) &
        bind ( c, name = 'name2row_c' )
      import c_char, c_int
      character(c_char), intent(in) :: name
    end function

    integer(c_int) function name2col ( name ) &
        bind ( c, name = 'name2col_c' )
      import c_char, c_int
      character(c_char), intent(in) :: name
    end function

    integer(c_int) function name2row2 ( name ) &
        bind ( c, name = 'name2row2_c' )
      import c_char, c_int
      character(c_char), intent(in) :: name
    end function

    integer(c_int) function name2col2 ( name ) &
        bind ( c, name = 'name2col2_c' )
      import c_char, c_int
      character(c_char), intent(in) :: name
    end function

    subroutine workbook_new ( workbook, file ) &
        bind ( c, name = 'workbook_new_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: workbook
      character(c_char), intent(in) :: file
    end subroutine

    type(c_ptr) function workbook_new2 ( file ) &
        bind ( c, name = 'workbook_new_c2' )
      import c_ptr, c_char
      character(c_char), intent(in) :: file
    end function

    subroutine workbook_add_worksheet ( workbook, worksheet, name ) &
        bind ( c, name = 'workbook_add_worksheet_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: workbook, worksheet
      character(c_char), intent(in) :: name
    end subroutine

    ! test
    type(c_ptr) function workbook_add_worksheet2 ( workbook, name ) &
        bind ( c, name = 'workbook_add_worksheet_c2' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: workbook
      character(c_char), intent(in) :: name
    end function

    subroutine worksheet_activate ( worksheet ) &
        bind ( c, name = 'worksheet_activate_c' )
      import c_ptr
      type(c_ptr), intent(in) :: worksheet
    end subroutine

    subroutine worksheet_hide ( worksheet ) &
        bind ( c, name = 'worksheet_hide_c' )
      import c_ptr
      type(c_ptr), intent(in) :: worksheet
    end subroutine

    subroutine worksheet_set_first_sheet ( worksheet ) &
        bind ( c, name = 'worksheet_set_first_sheet_c' )
      import c_ptr
      type(c_ptr), intent(in) :: worksheet
    end subroutine

    subroutine workbook_add_format ( workbook, format ) &
        bind ( c, name = 'workbook_add_format_c' )
      import c_ptr
      type(c_ptr), intent(in) :: workbook, format
    end subroutine

    type(c_ptr) function workbook_add_format2 ( workbook ) &
        bind ( c, name = 'workbook_add_format_c2' )
      import c_ptr
      type(c_ptr), intent(in) :: workbook
    end function

    subroutine workbook_define_name ( workbook, name, formula ) &
        bind ( c, name = 'workbook_define_name_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: workbook
      character(c_char), intent(in) :: name, formula
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

    subroutine worksheet_merge_range ( worksheet, first_row, first_col, last_row, last_col, string, format ) &
        bind ( c, name = 'worksheet_merge_range_c' )
      import c_ptr, c_int, c_char
      type(c_ptr),       intent(in)        :: worksheet, format
      integer(c_int),    intent(in), value :: first_row, first_col, last_row, last_col
      character(c_char), intent(in)        :: string
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

    subroutine worksheet_write_string ( worksheet, format, row, col, string ) &
        bind ( c, name = 'worksheet_write_string_c' )
      import c_ptr, c_int, c_char
      type(c_ptr),       intent(in)        :: worksheet, format
      integer(c_int),    intent(in), value :: row, col
      character(c_char), intent(in)        :: string
    end subroutine

    type(image_options_ty) function image_set_options &
        ( x_offset, y_offset, x_scale, y_scale, position, description, url, tip ) &
        bind ( c, name = 'image_set_options_c' )
      import c_int, c_double, c_char, image_options_ty
      integer(c_int),    intent(in), value    :: x_offset, y_offset
      real(c_double),    intent(in), value    :: x_scale, y_scale
      character(c_char), intent(in), optional :: position
      character(c_char), intent(in), optional :: description, url, tip
    end function

    type(c_ptr) function image_set_options2 &
        ( x_offset, y_offset, x_scale, y_scale, position, description, url, tip ) &
        bind ( c, name = 'image_set_options_c2' )
      import c_int, c_double, c_char, c_ptr
      integer(c_int),    intent(in), value    :: x_offset, y_offset
      real(c_double),    intent(in), value    :: x_scale, y_scale
      character(c_char), intent(in), optional :: position
      character(c_char), intent(in), optional :: description, url, tip
    end function

    subroutine worksheet_insert_image_opt ( worksheet, row, col, file, options ) &
        bind ( c, name = 'worksheet_insert_image_opt_c' )
      import c_ptr, c_int, c_char, image_options_ty
      type(c_ptr),            intent(in)        :: worksheet
      type(image_options_ty), intent(in)        :: options
      integer(c_int),         intent(in), value :: row, col
      character(c_char),      intent(in)        :: file
    end subroutine

    subroutine worksheet_insert_image_opt2 ( worksheet, row, col, file, options ) &
        bind ( c, name = 'worksheet_insert_image_opt_c' ) ! same as the above
      import c_ptr, c_int, c_char
      type(c_ptr),       intent(in)        :: worksheet
      type(c_ptr),       intent(in)        :: options
      integer(c_int),    intent(in), value :: row, col
      character(c_char), intent(in)        :: file
    end subroutine

    subroutine worksheet_insert_image ( worksheet, row, col, file ) &
        bind ( c, name = 'worksheet_insert_image_c' )
      import c_ptr, c_int, c_char
      type(c_ptr),       intent(in)        :: worksheet
      integer(c_int),    intent(in), value :: row, col
      character(c_char), intent(in)        :: file
    end subroutine

    subroutine worksheet_write_comment ( worksheet, row, col, string ) &
        bind ( c, name = 'worksheet_write_comment_c' )
      import c_ptr, c_int, c_char
      type(c_ptr),       intent(in)        :: worksheet
      integer(c_int),    intent(in), value :: row, col
      character(c_char), intent(in)        :: string
    end subroutine

    subroutine worksheet_set_header ( worksheet, string ) &
        bind ( c, name = 'worksheet_set_header_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: worksheet
      character(c_char), intent(in) :: string
    end subroutine

    subroutine worksheet_set_footer ( worksheet, string ) &
        bind ( c, name = 'worksheet_set_footer_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: worksheet
      character(c_char), intent(in) :: string
    end subroutine

    subroutine workbook_close ( workbook ) &
        bind ( c, name = 'workbook_close_c' )
      import c_ptr
      type(c_ptr), intent(in) :: workbook
    end subroutine

    ! URL
    subroutine worksheet_write_url ( worksheet, row, col, url, format ) &
        bind ( c, name = 'worksheet_write_url_c' )
      import c_ptr, c_int, c_char
      type(c_ptr), intent(in) :: worksheet, format
      integer(c_int),    intent(in), value :: row, col
      character(c_char), intent(in)        :: url
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

    ! Tab color
    subroutine worksheet_set_tab_color ( worksheet, color ) &
        bind ( c, name = 'worksheet_set_tab_color_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: worksheet
      character(c_char), intent(in) :: color
    end subroutine

    ! Font color
    subroutine format_set_font_color ( format, color ) &
        bind ( c, name = 'format_set_font_color_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: color
    end subroutine

    ! Underline
    subroutine format_set_underline ( format, style ) &
        bind ( c, name = 'format_set_underline_c' )
      import c_ptr, c_char
      type(c_ptr),       intent(in) :: format
      character(c_char), intent(in) :: style
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
