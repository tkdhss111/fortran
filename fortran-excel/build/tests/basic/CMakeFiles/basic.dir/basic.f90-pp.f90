# 1 "../tests/basic/basic.f90"
# 1 "<built-in>"
# 1 "<command-line>"
# 1 "../tests/basic/basic.f90"
program main

  use excel_mo

  implicit none

  integer, parameter :: NFMS = 9 ! Number of formats

  type(c_ptr)            :: wb(1)      ! Workbook pointer
  type(c_ptr)            :: ws(1:3)    ! Worksheet pointer
  type(c_ptr)            :: fm(0:NFMS) ! Format pointer ! 0 is reserved for no format
  type(datetime_ty)      :: datetime
  type(image_options_ty) :: image_options

  integer i





  ! Macro
  print *, 'CELL:', name2row(cs('A2')), name2col(cs('A2'))
  print *, 'COLS:', name2col(cs('A:C')), name2col2(cs('A:C'))
  print *, 'RANGE:', name2row(cs('A1:C3')), name2col(cs('A1:C3')), name2row2(cs('A1:C3')), name2col2(cs('A1:C3'))

  ! Workbook and worksheets
  call workbook_new ( wb(1), file = cs('test.xlsx') )
  call workbook_add_worksheet ( wb(1), ws(1), cs('first_sheet') )
  call workbook_add_worksheet ( wb(1), ws(2), cs('second_sheet') )
  call workbook_add_worksheet ( wb(1), ws(3), cs('third_sheet') )
  call worksheet_activate ( ws(2) )
  call worksheet_hide     ( ws(3) )
  call worksheet_set_tab_color ( ws(1), cs('lime') )

  ! Header and footer
  call worksheet_set_header ( ws(1), cs('header') )
  call worksheet_set_footer ( ws(1), cs('footer') )

  ! Format settings
  do i = 0, NFMS 

    call workbook_add_format ( wb(1), fm(i) )

    if ( mod ( i, 2 ) == 0 .and. i > 0 .and. i < 5 ) then
      call format_set_bg_color ( fm(i), cs('cyan') )
    end if

  end do

  ! Header of table
  call worksheet_set_row     ( ws(1), fm(0), row = 1, height = 40 )
  call worksheet_set_column  ( ws(1), fm(0), col = 2, width  = 40 )
  call format_set_bold       ( fm(1) )
  call format_set_bg_color   ( fm(1), cs('navy')  )
  call format_set_font_color ( fm(1), cs('white') )
  call format_set_font_name  ( fm(1), cs('YuKyokasho') )
  call format_set_align      ( fm(1), cs('center') )
  call format_set_align      ( fm(1), cs('vertical_center') )
  call format_set_font_size  ( fm(1), 20 )
  call worksheet_write_string ( ws(1), fm(1), row = 1, col = 1, string = cs('Ho-gens')     )
  call worksheet_write_string ( ws(1), fm(1), row = 1, col = 2, string = cs('Panic Level') )

  ! Table
  call worksheet_write_string ( ws(1), fm(2), row = 2, col = 1, string = cs('Yoka')        )
  call worksheet_write_string ( ws(1), fm(3), row = 3, col = 1, string = cs('Batten')      )
  call worksheet_write_string ( ws(1), fm(4), row = 4, col = 1, string = cs('Batten-gara') )
  call worksheet_write_number ( ws(1), fm(2), row = 2, col = 2, value = 1.02d0 )
  call worksheet_write_number ( ws(1), fm(3), row = 3, col = 2, value = 3.14d0 )
  call worksheet_write_number ( ws(1), fm(4), row = 4, col = 2, value = 99.0d0 )
  call worksheet_autofilter  ( ws(1), 1, 1, 4, 2 ) 

  ! Black line
  call worksheet_set_row ( ws(1), fm(0), row = 5, height = 5 )

  ! Totals
  call format_set_align      ( fm(8), cs('right') )
  call format_set_align      ( fm(8), cs('vertical_right') )
  call format_set_bold       ( fm(8) )
  call worksheet_write_string  ( ws(1), fm(8), row = 6, col = 1, string = cs('Totals:') )
  call worksheet_write_comment ( ws(1), row = 6, col = 1, string = cs('This is comment') )

  call format_set_num_format ( fm(9), cs('0.0') )
  call format_set_bg_color   ( fm(9), cs('yellow')  )
  call format_set_border     ( fm(9), cs('thick')  )
  call worksheet_write_formula ( ws(1), row = 6, col = 2, formula = cs('=SUM(B2:B4)'), format = fm(9) )

  ! Insert image
  image_options = image_options_ty ( x_offset = 10, y_offset = 10, x_scale = 0.5d0, y_scale = 0.5d0 )
  call worksheet_insert_image_opt ( ws(1), row = 15, col = 1, file = cs('fig.png'), options = image_options )
  call worksheet_insert_image ( ws(1), row = 8, col = 1, file = cs('fig.png') )

  ! Datetime
  call format_set_num_format ( fm(7), cs('yyyy-mm-dd hh:mm:ss') )
  datetime = datetime_ty( year = 2021, month = 3, day = 1, hour = 2, min = 30, sec = 10.d0 ) 
  call worksheet_write_datetime( ws(1), name2row(cs('B7')), name2col(cs('B7')), datetime, fm(7) )

  ! URL
  call format_set_font_color ( fm(5), cs('blue') )
  call format_set_underline ( fm(5), cs('single') )
  call worksheet_write_url ( ws(1), row = 8, col = 2, url = cs('http://libxlsxwriter.github.io'), format = fm(5) )

  ! Defined name
  call workbook_define_name ( wb(1), name = cs('Exchange_rate'), formula = cs('=110.0') )
  call worksheet_write_formula ( ws(1), row = 9, col = 2, formula = cs('=Exchange_rate'), format = fm(0) )

  ! Merge range
  call worksheet_merge_range ( ws(1), name2row(cs('B8:C8')), name2col(cs('B8:C8')), name2row2(cs('B8:C8')), name2col2(cs('B8:C8')), cs('Merged range'), fm(1) )

  ! Close the workbook
  call workbook_close ( wb(1) )

end program
