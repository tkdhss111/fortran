program main

  use excel_mo

  implicit none

  integer, parameter :: NFMS = 9 ! Number of formats

  type(c_ptr) :: wb         ! Workbook pointer
  type(c_ptr) :: ws         ! Worksheet pointer
  type(c_ptr) :: fm(0:NFMS) ! Format pointer ! 0 is reserved for no format
  type(datetime_ty) :: datetime

  integer i

  ! Workbook and worksheets
  call workbook_new ( wb, file = cs('test.xlsx') )
  call workbook_add_worksheet ( wb, ws, cs('new_sheet') )
  call worksheet_set_tab_color ( ws, cs('lime') )

  ! Format settings
  do i = 0, NFMS 

    call workbook_add_format ( wb, fm(i) )

    if ( mod ( i, 2 ) == 0 .and. i > 0 .and. i < 5 ) then
      call format_set_bg_color ( fm(i), cs('cyan') )
    end if

  end do

  ! Header
  call worksheet_set_row     ( ws, fm(0), row = 1, height = 40 )
  call worksheet_set_column  ( ws, fm(0), col = 2, width  = 40 )
  call format_set_bold       ( fm(1) )
  call format_set_bg_color   ( fm(1), cs('navy')  )
  call format_set_font_color ( fm(1), cs('white') )
  call format_set_font_name  ( fm(1), cs('YuKyokasho') )
  call format_set_align      ( fm(1), cs('center') )
  call format_set_align      ( fm(1), cs('vertical_center') )
  call format_set_font_size  ( fm(1), 20 )
  call worksheet_write_string ( ws, fm(1), row = 1, col = 1, text = cs('Ho-gens')     )
  call worksheet_write_string ( ws, fm(1), row = 1, col = 2, text = cs('Panic Level') )

  ! Table
  call worksheet_write_string ( ws, fm(2), row = 2, col = 1, text = cs('Yoka')        )
  call worksheet_write_string ( ws, fm(3), row = 3, col = 1, text = cs('Batten')      )
  call worksheet_write_string ( ws, fm(4), row = 4, col = 1, text = cs('Batten-gara') )
  call worksheet_write_number ( ws, fm(2), row = 2, col = 2, value = 1.02d0 )
  call worksheet_write_number ( ws, fm(3), row = 3, col = 2, value = 3.14d0 )
  call worksheet_write_number ( ws, fm(4), row = 4, col = 2, value = 99.0d0 )
  call worksheet_autofilter  ( ws, 1, 1, 4, 2 ) 

  ! Black line
  call worksheet_set_row ( ws, fm(0), row = 5, height = 5 )

  ! Totals
  call format_set_align      ( fm(8), cs('right') )
  call format_set_align      ( fm(8), cs('vertical_right') )
  call format_set_bold       ( fm(8) )
  call worksheet_write_string  ( ws, fm(8), row = 6, col = 1, text = cs('Totals:')     )

  call format_set_num_format ( fm(9), cs('0.0') )
  call format_set_bg_color   ( fm(9), cs('yellow')  )
  call format_set_border     ( fm(9), cs('thick')  )
  call worksheet_write_formula ( ws, row = 6, col = 2, formula = cs('=SUM(B2:B4)'), format = fm(9) )

  ! Insert image
  call worksheet_insert_image ( ws, row = 8, col = 1, file = cs('fig.png') )

  ! Datetime
  call format_set_num_format ( fm(7), cs('yyyy-mm-dd hh:mm:ss') )
  datetime = datetime_ty( year = 2021, month = 3, day = 1, hour = 2, min = 30, sec = 10.d0 ) 
  call worksheet_write_datetime( ws, 7, 2, datetime, fm(7) )

  ! URL
  call format_set_font_color ( fm(5), cs('blue') )
  call format_set_underline ( fm(5), cs('single') )
  call worksheet_write_url ( ws, row = 8, col = 2, url = cs('http://libxlsxwriter.github.io'), format = fm(5) )

  ! Defined name
  call workbook_define_name ( wb, name = cs('Exchange_rate'), formula = cs('=110.0') )
  call worksheet_write_formula ( ws, row = 9, col = 2, formula = cs('=Exchange_rate'), format = fm(0) )

  ! Close the workbook
  call workbook_close ( wb )

end program
