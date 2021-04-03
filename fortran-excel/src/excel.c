#include "xlsxwriter.h"

// See the following URL to add more functions.
// USL: https://libxlsxwriter.github.io/working_with_formats.html

void workbook_new_c( lxw_workbook **workbook, const char* file )
{
  *workbook = workbook_new( file );
}

void workbook_add_worksheet_c( lxw_workbook **workbook, lxw_worksheet **worksheet, const char* name )
{
  *worksheet = workbook_add_worksheet( *workbook, name );
}

void workbook_add_format_c( lxw_workbook **workbook, lxw_format **format )
{
  *format = workbook_add_format( *workbook );
}

void workbook_define_name_c( lxw_workbook **workbook, const char* name, const char* formula )
{
  workbook_define_name( *workbook, name, formula );
}

void worksheet_write_datetime_c( lxw_worksheet **worksheet, int row, int col, lxw_datetime *datetime, lxw_format **format )
{
  worksheet_write_datetime( *worksheet, row - 1, col - 1, datetime, *format );
}

void format_set_shrink_c( lxw_format **format )
{
  format_set_shrink( *format );
}

void format_set_bold_c( lxw_format **format )
{
  format_set_bold( *format );
}

void format_set_italic_c( lxw_format **format )
{
  format_set_italic( *format );
}

void worksheet_autofilter_c( lxw_worksheet **worksheet, int first_row, int first_col, int last_row, int last_col )
{
  worksheet_autofilter( *worksheet, first_row - 1, first_col - 1, last_row - 1, last_col - 1 );
}

void worksheet_set_row_c( lxw_worksheet **worksheet, lxw_format **format, int row, int height )
{
  worksheet_set_row( *worksheet, row - 1, height, *format );
}

void worksheet_set_column_c( lxw_worksheet **worksheet, lxw_format **format, int col, int width )
{
  worksheet_set_column( *worksheet, 0, col - 1, width, *format );
}

void worksheet_write_formula_c( lxw_worksheet **worksheet, int row, int col, const char* formula, lxw_format **format )
{
  worksheet_write_formula( *worksheet, row - 1, col - 1, formula, *format );
}

void worksheet_write_number_c( lxw_worksheet **worksheet, lxw_format **format, int row, int col, double value )
{
  worksheet_write_number( *worksheet, row - 1, col - 1, value, *format );
}

void worksheet_write_string_c( lxw_worksheet **worksheet, lxw_format **format, int row, int col, const char* text )
{
  worksheet_write_string( *worksheet, row - 1, col - 1, text, *format );
}

void worksheet_write_comment_c( lxw_worksheet **worksheet, int row, int col, const char* text )
{
  worksheet_write_comment( *worksheet, row - 1, col - 1, text );
}

void worksheet_insert_image_c( lxw_worksheet **worksheet, int row, int col, const char* file )
{
  worksheet_insert_image( *worksheet, row - 1, col - 1, file );
}

void workbook_close_c( lxw_workbook **workbook )
{
  workbook_close( *workbook );
}

// URL
void worksheet_write_url_c( lxw_worksheet **worksheet, int row, int col, const char* url, lxw_format **format )
{
  worksheet_write_url( *worksheet, row - 1, col - 1, url, *format );
}

// Font size
void format_set_font_size_c( lxw_format **format, int size ) { format_set_font_size( *format, (double)size ); }

// Font name
void format_set_font_name_c( lxw_format **format, const char* name ) { format_set_font_name( *format, name ); }

// Number format 
void format_set_num_format_c( lxw_format **format, const char* style ) { format_set_num_format( *format, style ); }

// Tab color
void worksheet_set_tab_color_c( lxw_worksheet **worksheet, const char* color )
{
  if ( strcmp(color, "black"  ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_BLACK  ); 
  if ( strcmp(color, "blue"   ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_BLUE   ); 
  if ( strcmp(color, "brown"  ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_BROWN  ); 
  if ( strcmp(color, "cyan"   ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_CYAN   ); 
  if ( strcmp(color, "gray"   ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_GRAY   ); 
  if ( strcmp(color, "green"  ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_GREEN  ); 
  if ( strcmp(color, "lime"   ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_LIME   ); 
  if ( strcmp(color, "magenta") == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_MAGENTA); 
  if ( strcmp(color, "navy"   ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_NAVY   ); 
  if ( strcmp(color, "orange" ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_ORANGE ); 
  if ( strcmp(color, "pink"   ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_PINK   ); 
  if ( strcmp(color, "purple" ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_PURPLE ); 
  if ( strcmp(color, "red"    ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_RED    ); 
  if ( strcmp(color, "silver" ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_SILVER ); 
  if ( strcmp(color, "white"  ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_WHITE  ); 
  if ( strcmp(color, "yellow" ) == 0 ) worksheet_set_tab_color( *worksheet, LXW_COLOR_YELLOW ); 
}

// Font color
void format_set_font_color_c( lxw_format **format, const char* color )
{ 
  if ( strcmp(color, "black"  ) == 0 ) format_set_font_color( *format, LXW_COLOR_BLACK  ); 
  if ( strcmp(color, "blue"   ) == 0 ) format_set_font_color( *format, LXW_COLOR_BLUE   ); 
  if ( strcmp(color, "brown"  ) == 0 ) format_set_font_color( *format, LXW_COLOR_BROWN  ); 
  if ( strcmp(color, "cyan"   ) == 0 ) format_set_font_color( *format, LXW_COLOR_CYAN   ); 
  if ( strcmp(color, "gray"   ) == 0 ) format_set_font_color( *format, LXW_COLOR_GRAY   ); 
  if ( strcmp(color, "green"  ) == 0 ) format_set_font_color( *format, LXW_COLOR_GREEN  ); 
  if ( strcmp(color, "lime"   ) == 0 ) format_set_font_color( *format, LXW_COLOR_LIME   ); 
  if ( strcmp(color, "magenta") == 0 ) format_set_font_color( *format, LXW_COLOR_MAGENTA); 
  if ( strcmp(color, "navy"   ) == 0 ) format_set_font_color( *format, LXW_COLOR_NAVY   ); 
  if ( strcmp(color, "orange" ) == 0 ) format_set_font_color( *format, LXW_COLOR_ORANGE ); 
  if ( strcmp(color, "pink"   ) == 0 ) format_set_font_color( *format, LXW_COLOR_PINK   ); 
  if ( strcmp(color, "purple" ) == 0 ) format_set_font_color( *format, LXW_COLOR_PURPLE ); 
  if ( strcmp(color, "red"    ) == 0 ) format_set_font_color( *format, LXW_COLOR_RED    ); 
  if ( strcmp(color, "silver" ) == 0 ) format_set_font_color( *format, LXW_COLOR_SILVER ); 
  if ( strcmp(color, "white"  ) == 0 ) format_set_font_color( *format, LXW_COLOR_WHITE  ); 
  if ( strcmp(color, "yellow" ) == 0 ) format_set_font_color( *format, LXW_COLOR_YELLOW ); 
}

// Background color
void format_set_bg_color_c( lxw_format **format, const char* color )
{ 
  if ( strcmp(color, "black"  ) == 0 ) format_set_bg_color( *format, LXW_COLOR_BLACK  ); 
  if ( strcmp(color, "blue"   ) == 0 ) format_set_bg_color( *format, LXW_COLOR_BLUE   ); 
  if ( strcmp(color, "brown"  ) == 0 ) format_set_bg_color( *format, LXW_COLOR_BROWN  ); 
  if ( strcmp(color, "cyan"   ) == 0 ) format_set_bg_color( *format, LXW_COLOR_CYAN   ); 
  if ( strcmp(color, "gray"   ) == 0 ) format_set_bg_color( *format, LXW_COLOR_GRAY   ); 
  if ( strcmp(color, "green"  ) == 0 ) format_set_bg_color( *format, LXW_COLOR_GREEN  ); 
  if ( strcmp(color, "lime"   ) == 0 ) format_set_bg_color( *format, LXW_COLOR_LIME   ); 
  if ( strcmp(color, "magenta") == 0 ) format_set_bg_color( *format, LXW_COLOR_MAGENTA); 
  if ( strcmp(color, "navy"   ) == 0 ) format_set_bg_color( *format, LXW_COLOR_NAVY   ); 
  if ( strcmp(color, "orange" ) == 0 ) format_set_bg_color( *format, LXW_COLOR_ORANGE ); 
  if ( strcmp(color, "pink"   ) == 0 ) format_set_bg_color( *format, LXW_COLOR_PINK   ); 
  if ( strcmp(color, "purple" ) == 0 ) format_set_bg_color( *format, LXW_COLOR_PURPLE ); 
  if ( strcmp(color, "red"    ) == 0 ) format_set_bg_color( *format, LXW_COLOR_RED    ); 
  if ( strcmp(color, "silver" ) == 0 ) format_set_bg_color( *format, LXW_COLOR_SILVER ); 
  if ( strcmp(color, "white"  ) == 0 ) format_set_bg_color( *format, LXW_COLOR_WHITE  ); 
  if ( strcmp(color, "yellow" ) == 0 ) format_set_bg_color( *format, LXW_COLOR_YELLOW ); 
}

void format_set_underline_c( lxw_format **format, const char* style )
{ 
  if ( strcmp(style, "single"            ) == 0 ) format_set_underline( *format, LXW_UNDERLINE_SINGLE            );
  if ( strcmp(style, "double"            ) == 0 ) format_set_underline( *format, LXW_UNDERLINE_DOUBLE            );
  if ( strcmp(style, "single_accounting" ) == 0 ) format_set_underline( *format, LXW_UNDERLINE_SINGLE_ACCOUNTING );
  if ( strcmp(style, "double_accounting" ) == 0 ) format_set_underline( *format, LXW_UNDERLINE_DOUBLE_ACCOUNTING );
}

// Pattern
void format_set_pattern_c( lxw_format **format, const char* pattern )
{ 
    int LXW_PATTERN_NONE = 0;

    if ( strcmp(pattern, "none"            ) == 0 ) format_set_pattern( *format, LXW_PATTERN_NONE            );
    if ( strcmp(pattern, "solid"           ) == 0 ) format_set_pattern( *format, LXW_PATTERN_SOLID           );
    if ( strcmp(pattern, "medium_gray"     ) == 0 ) format_set_pattern( *format, LXW_PATTERN_MEDIUM_GRAY     );
    if ( strcmp(pattern, "dark_gray"       ) == 0 ) format_set_pattern( *format, LXW_PATTERN_DARK_GRAY       );
    if ( strcmp(pattern, "light_gray"      ) == 0 ) format_set_pattern( *format, LXW_PATTERN_LIGHT_GRAY      );
    if ( strcmp(pattern, "dark_horizontal" ) == 0 ) format_set_pattern( *format, LXW_PATTERN_DARK_HORIZONTAL );
    if ( strcmp(pattern, "dark_vertical"   ) == 0 ) format_set_pattern( *format, LXW_PATTERN_DARK_VERTICAL   );
    if ( strcmp(pattern, "dark_down"       ) == 0 ) format_set_pattern( *format, LXW_PATTERN_DARK_DOWN       );
    if ( strcmp(pattern, "dark_up"         ) == 0 ) format_set_pattern( *format, LXW_PATTERN_DARK_UP         );
    if ( strcmp(pattern, "dark_grid"       ) == 0 ) format_set_pattern( *format, LXW_PATTERN_DARK_GRID       );
    if ( strcmp(pattern, "dark_trellis"    ) == 0 ) format_set_pattern( *format, LXW_PATTERN_DARK_TRELLIS    );
    if ( strcmp(pattern, "light_horizontal") == 0 ) format_set_pattern( *format, LXW_PATTERN_LIGHT_HORIZONTAL);
    if ( strcmp(pattern, "light_vertical"  ) == 0 ) format_set_pattern( *format, LXW_PATTERN_LIGHT_VERTICAL  );
    if ( strcmp(pattern, "light_down"      ) == 0 ) format_set_pattern( *format, LXW_PATTERN_LIGHT_DOWN      );
    if ( strcmp(pattern, "light_up"        ) == 0 ) format_set_pattern( *format, LXW_PATTERN_LIGHT_UP        );
    if ( strcmp(pattern, "light_grid"      ) == 0 ) format_set_pattern( *format, LXW_PATTERN_LIGHT_GRID      );
    if ( strcmp(pattern, "light_trellis"   ) == 0 ) format_set_pattern( *format, LXW_PATTERN_LIGHT_TRELLIS   );
    if ( strcmp(pattern, "gray_125"        ) == 0 ) format_set_pattern( *format, LXW_PATTERN_GRAY_125        );
    if ( strcmp(pattern, "gray_0625"       ) == 0 ) format_set_pattern( *format, LXW_PATTERN_GRAY_0625       );
}

// Alignment
void format_set_align_c ( lxw_format **format, const char* align )
{
  int LXW_ALIGN_NONE = 0;

  if ( strcmp(align, "none"                 ) == 0 ) format_set_align( *format, LXW_ALIGN_NONE                 );
  if ( strcmp(align, "left"                 ) == 0 ) format_set_align( *format, LXW_ALIGN_LEFT                 );
  if ( strcmp(align, "center"               ) == 0 ) format_set_align( *format, LXW_ALIGN_CENTER               );
  if ( strcmp(align, "right"                ) == 0 ) format_set_align( *format, LXW_ALIGN_RIGHT                );
  if ( strcmp(align, "fill"                 ) == 0 ) format_set_align( *format, LXW_ALIGN_FILL                 );
  if ( strcmp(align, "justify"              ) == 0 ) format_set_align( *format, LXW_ALIGN_JUSTIFY              );
  if ( strcmp(align, "ceter_across"         ) == 0 ) format_set_align( *format, LXW_ALIGN_CENTER_ACROSS        );
  if ( strcmp(align, "distributed"          ) == 0 ) format_set_align( *format, LXW_ALIGN_DISTRIBUTED          );
  if ( strcmp(align, "vertical_top"         ) == 0 ) format_set_align( *format, LXW_ALIGN_VERTICAL_TOP         );
  if ( strcmp(align, "vertical_bottom"      ) == 0 ) format_set_align( *format, LXW_ALIGN_VERTICAL_BOTTOM      );
  if ( strcmp(align, "vertical_center"      ) == 0 ) format_set_align( *format, LXW_ALIGN_VERTICAL_CENTER      );
  if ( strcmp(align, "vertical_justify"     ) == 0 ) format_set_align( *format, LXW_ALIGN_VERTICAL_JUSTIFY     );
  if ( strcmp(align, "vertical_distributed" ) == 0 ) format_set_align( *format, LXW_ALIGN_VERTICAL_DISTRIBUTED );
}

//
// Boarders
//

// Top, bottom, right and left
void format_set_border_c( lxw_format **format, const char* style )
{ 
  int LXW_BORDER_NONE = 0;

  if ( strcmp(style, "none"                ) == 0 ) format_set_border( *format, LXW_BORDER_NONE                );
  if ( strcmp(style, "thin"                ) == 0 ) format_set_border( *format, LXW_BORDER_THIN                );
  if ( strcmp(style, "medium"              ) == 0 ) format_set_border( *format, LXW_BORDER_MEDIUM              );
  if ( strcmp(style, "dashed"              ) == 0 ) format_set_border( *format, LXW_BORDER_DASHED              );
  if ( strcmp(style, "dotted"              ) == 0 ) format_set_border( *format, LXW_BORDER_DOTTED              );
  if ( strcmp(style, "thick"               ) == 0 ) format_set_border( *format, LXW_BORDER_THICK               );
  if ( strcmp(style, "double"              ) == 0 ) format_set_border( *format, LXW_BORDER_DOUBLE              );
  if ( strcmp(style, "hair"                ) == 0 ) format_set_border( *format, LXW_BORDER_HAIR                );
  if ( strcmp(style, "medium_dashed"       ) == 0 ) format_set_border( *format, LXW_BORDER_MEDIUM_DASHED       );
  if ( strcmp(style, "dash_dot"            ) == 0 ) format_set_border( *format, LXW_BORDER_DASH_DOT            );
  if ( strcmp(style, "medium_dash_dot"     ) == 0 ) format_set_border( *format, LXW_BORDER_MEDIUM_DASH_DOT     );
  if ( strcmp(style, "dash_dot_dot"        ) == 0 ) format_set_border( *format, LXW_BORDER_DASH_DOT_DOT        );
  if ( strcmp(style, "medium_dash_dot_dot" ) == 0 ) format_set_border( *format, LXW_BORDER_MEDIUM_DASH_DOT_DOT );
  if ( strcmp(style, "slant_dash_dot"      ) == 0 ) format_set_border( *format, LXW_BORDER_SLANT_DASH_DOT      );
}

// Top
void format_set_top_c( lxw_format **format, const char* style )
{ 
  int LXW_BORDER_NONE = 0;

  if ( strcmp(style, "none"                ) == 0 ) format_set_top( *format, LXW_BORDER_NONE                );
  if ( strcmp(style, "thin"                ) == 0 ) format_set_top( *format, LXW_BORDER_THIN                );
  if ( strcmp(style, "medium"              ) == 0 ) format_set_top( *format, LXW_BORDER_MEDIUM              );
  if ( strcmp(style, "dashed"              ) == 0 ) format_set_top( *format, LXW_BORDER_DASHED              );
  if ( strcmp(style, "dotted"              ) == 0 ) format_set_top( *format, LXW_BORDER_DOTTED              );
  if ( strcmp(style, "thick"               ) == 0 ) format_set_top( *format, LXW_BORDER_THICK               );
  if ( strcmp(style, "double"              ) == 0 ) format_set_top( *format, LXW_BORDER_DOUBLE              );
  if ( strcmp(style, "hair"                ) == 0 ) format_set_top( *format, LXW_BORDER_HAIR                );
  if ( strcmp(style, "medium_dashed"       ) == 0 ) format_set_top( *format, LXW_BORDER_MEDIUM_DASHED       );
  if ( strcmp(style, "dash_dot"            ) == 0 ) format_set_top( *format, LXW_BORDER_DASH_DOT            );
  if ( strcmp(style, "medium_dash_dot"     ) == 0 ) format_set_top( *format, LXW_BORDER_MEDIUM_DASH_DOT     );
  if ( strcmp(style, "dash_dot_dot"        ) == 0 ) format_set_top( *format, LXW_BORDER_DASH_DOT_DOT        );
  if ( strcmp(style, "medium_dash_dot_dot" ) == 0 ) format_set_top( *format, LXW_BORDER_MEDIUM_DASH_DOT_DOT );
  if ( strcmp(style, "slant_dash_dot"      ) == 0 ) format_set_top( *format, LXW_BORDER_SLANT_DASH_DOT      );
}

// Bottom
void format_set_bottom_c( lxw_format **format, const char* style )
{ 
  int LXW_BORDER_NONE = 0;

  if ( strcmp(style, "none"                ) == 0 ) format_set_bottom( *format, LXW_BORDER_NONE                );
  if ( strcmp(style, "thin"                ) == 0 ) format_set_bottom( *format, LXW_BORDER_THIN                );
  if ( strcmp(style, "medium"              ) == 0 ) format_set_bottom( *format, LXW_BORDER_MEDIUM              );
  if ( strcmp(style, "dashed"              ) == 0 ) format_set_bottom( *format, LXW_BORDER_DASHED              );
  if ( strcmp(style, "dotted"              ) == 0 ) format_set_bottom( *format, LXW_BORDER_DOTTED              );
  if ( strcmp(style, "thick"               ) == 0 ) format_set_bottom( *format, LXW_BORDER_THICK               );
  if ( strcmp(style, "double"              ) == 0 ) format_set_bottom( *format, LXW_BORDER_DOUBLE              );
  if ( strcmp(style, "hair"                ) == 0 ) format_set_bottom( *format, LXW_BORDER_HAIR                );
  if ( strcmp(style, "medium_dashed"       ) == 0 ) format_set_bottom( *format, LXW_BORDER_MEDIUM_DASHED       );
  if ( strcmp(style, "dash_dot"            ) == 0 ) format_set_bottom( *format, LXW_BORDER_DASH_DOT            );
  if ( strcmp(style, "medium_dash_dot"     ) == 0 ) format_set_bottom( *format, LXW_BORDER_MEDIUM_DASH_DOT     );
  if ( strcmp(style, "dash_dot_dot"        ) == 0 ) format_set_bottom( *format, LXW_BORDER_DASH_DOT_DOT        );
  if ( strcmp(style, "medium_dash_dot_dot" ) == 0 ) format_set_bottom( *format, LXW_BORDER_MEDIUM_DASH_DOT_DOT );
  if ( strcmp(style, "slant_dash_dot"      ) == 0 ) format_set_bottom( *format, LXW_BORDER_SLANT_DASH_DOT      );
}

// Right
void format_set_right_c( lxw_format **format, const char* style )
{ 
  int LXW_BORDER_NONE = 0;

  if ( strcmp(style, "none"                ) == 0 ) format_set_right( *format, LXW_BORDER_NONE                );
  if ( strcmp(style, "thin"                ) == 0 ) format_set_right( *format, LXW_BORDER_THIN                );
  if ( strcmp(style, "medium"              ) == 0 ) format_set_right( *format, LXW_BORDER_MEDIUM              );
  if ( strcmp(style, "dashed"              ) == 0 ) format_set_right( *format, LXW_BORDER_DASHED              );
  if ( strcmp(style, "dotted"              ) == 0 ) format_set_right( *format, LXW_BORDER_DOTTED              );
  if ( strcmp(style, "thick"               ) == 0 ) format_set_right( *format, LXW_BORDER_THICK               );
  if ( strcmp(style, "double"              ) == 0 ) format_set_right( *format, LXW_BORDER_DOUBLE              );
  if ( strcmp(style, "hair"                ) == 0 ) format_set_right( *format, LXW_BORDER_HAIR                );
  if ( strcmp(style, "medium_dashed"       ) == 0 ) format_set_right( *format, LXW_BORDER_MEDIUM_DASHED       );
  if ( strcmp(style, "dash_dot"            ) == 0 ) format_set_right( *format, LXW_BORDER_DASH_DOT            );
  if ( strcmp(style, "medium_dash_dot"     ) == 0 ) format_set_right( *format, LXW_BORDER_MEDIUM_DASH_DOT     );
  if ( strcmp(style, "dash_dot_dot"        ) == 0 ) format_set_right( *format, LXW_BORDER_DASH_DOT_DOT        );
  if ( strcmp(style, "medium_dash_dot_dot" ) == 0 ) format_set_right( *format, LXW_BORDER_MEDIUM_DASH_DOT_DOT );
  if ( strcmp(style, "slant_dash_dot"      ) == 0 ) format_set_right( *format, LXW_BORDER_SLANT_DASH_DOT      );
}

// Left
void format_set_left_c( lxw_format **format, const char* style )
{ 
  int LXW_BORDER_NONE = 0;

  if ( strcmp(style, "none"                ) == 0 ) format_set_left( *format, LXW_BORDER_NONE                );
  if ( strcmp(style, "thin"                ) == 0 ) format_set_left( *format, LXW_BORDER_THIN                );
  if ( strcmp(style, "medium"              ) == 0 ) format_set_left( *format, LXW_BORDER_MEDIUM              );
  if ( strcmp(style, "dashed"              ) == 0 ) format_set_left( *format, LXW_BORDER_DASHED              );
  if ( strcmp(style, "dotted"              ) == 0 ) format_set_left( *format, LXW_BORDER_DOTTED              );
  if ( strcmp(style, "thick"               ) == 0 ) format_set_left( *format, LXW_BORDER_THICK               );
  if ( strcmp(style, "double"              ) == 0 ) format_set_left( *format, LXW_BORDER_DOUBLE              );
  if ( strcmp(style, "hair"                ) == 0 ) format_set_left( *format, LXW_BORDER_HAIR                );
  if ( strcmp(style, "medium_dashed"       ) == 0 ) format_set_left( *format, LXW_BORDER_MEDIUM_DASHED       );
  if ( strcmp(style, "dash_dot"            ) == 0 ) format_set_left( *format, LXW_BORDER_DASH_DOT            );
  if ( strcmp(style, "medium_dash_dot"     ) == 0 ) format_set_left( *format, LXW_BORDER_MEDIUM_DASH_DOT     );
  if ( strcmp(style, "dash_dot_dot"        ) == 0 ) format_set_left( *format, LXW_BORDER_DASH_DOT_DOT        );
  if ( strcmp(style, "medium_dash_dot_dot" ) == 0 ) format_set_left( *format, LXW_BORDER_MEDIUM_DASH_DOT_DOT );
  if ( strcmp(style, "slant_dash_dot"      ) == 0 ) format_set_left( *format, LXW_BORDER_SLANT_DASH_DOT      );
}
