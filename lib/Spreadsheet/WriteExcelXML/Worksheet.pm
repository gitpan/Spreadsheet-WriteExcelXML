package Spreadsheet::WriteExcelXML::Worksheet;

###############################################################################
#
# Worksheet - A writer class for Excel Worksheets.
#
#
# Used in conjunction with Spreadsheet::WriteExcelXML
#
# Copyright 2000-2004, John McNamara, jmcnamara@cpan.org
#
# Documentation after __END__
#

use Exporter;
use strict;
use Carp;
use Spreadsheet::WriteExcelXML::XMLwriter;
use Spreadsheet::WriteExcelXML::Format;





use vars qw($VERSION @ISA);
@ISA = qw(Spreadsheet::WriteExcelXML::XMLwriter);

$VERSION = '0.01';

###############################################################################
#
# new()
#
# Constructor. Creates a new Worksheet object from a XMLwriter object
#
sub new {

    my $class                   = shift;
    my $self                    = Spreadsheet::WriteExcelXML::XMLwriter->new();
    my $rowmax                  = 65536; # 16384 in Excel 5
    my $colmax                  = 256;
    my $strmax                  = 32767;

    $self->{_name}              = $_[0];
    $self->{_index}             = $_[1];
    $self->{_filehandle}        = $_[2];
    $self->{_indentation}       = $_[3];
    $self->{_activesheet}       = $_[4];
    $self->{_firstsheet}        = $_[5];

    $self->{_ext_sheets}        = [];
    $self->{_fileclosed}        = 0;
    $self->{_offset}            = 0;
    $self->{_xls_rowmax}        = $rowmax;
    $self->{_xls_colmax}        = $colmax;
    $self->{_xls_strmax}        = $strmax;
    $self->{_dim_rowmin}        = $rowmax +1;
    $self->{_dim_rowmax}        = 0;
    $self->{_dim_colmin}        = $colmax +1;
    $self->{_dim_colmax}        = 0;
    $self->{_dim_changed}       = 0;
    $self->{_colinfo}           = [];
    $self->{_selection}         = [0, 0];
    $self->{_panes}             = [];
    $self->{_active_pane}       = 3;
    $self->{_frozen}            = 0;
    $self->{_selected}          = 0;

    $self->{_paper_size}        = 0x0;
    $self->{_orientation}       = 0x1;
    $self->{_header}            = '';
    $self->{_footer}            = '';
    $self->{_hcenter}           = 0;
    $self->{_vcenter}           = 0;
    $self->{_margin_head}       = 0.50;
    $self->{_margin_foot}       = 0.50;
    $self->{_margin_left}       = 0.75;
    $self->{_margin_right}      = 0.75;
    $self->{_margin_top}        = 1.00;
    $self->{_margin_bottom}     = 1.00;

    $self->{_title_rowmin}      = undef;
    $self->{_title_rowmax}      = undef;
    $self->{_title_colmin}      = undef;
    $self->{_title_colmax}      = undef;
    $self->{_print_rowmin}      = undef;
    $self->{_print_rowmax}      = undef;
    $self->{_print_colmin}      = undef;
    $self->{_print_colmax}      = undef;

    $self->{_print_gridlines}   = 1;
    $self->{_screen_gridlines}  = 1;
    $self->{_print_headers}     = 0;

    $self->{_fit_page}          = 0;
    $self->{_fit_width}         = 0;
    $self->{_fit_height}        = 0;

    $self->{_hbreaks}           = [];
    $self->{_vbreaks}           = [];

    $self->{_protect}           = 0;
    $self->{_password}          = undef;

    $self->{_col_sizes}         = {};
    $self->{_row_sizes}         = {};

    $self->{_col_formats}       = {};
    $self->{_row_formats}       = {};

    $self->{_zoom}              = 100;
    $self->{_print_scale}       = 100;

    $self->{_leading_zeros}     = 0;

    $self->{_outline_row_level} = 0;
    $self->{_outline_style}     = 0;
    $self->{_outline_below}     = 1;
    $self->{_outline_right}     = 1;
    $self->{_outline_on}        = 1;



    $self->{prev_row}           = -1;
    $self->{prev_col}           = -1;

    $self->{_table}             = [];
    $self->{_datatypes}         = {String   => 1,
                                   Number   => 2,
                                   DateTime => 3,
                                   Formula  => 4,
                                   Blank    => 5,
                                  };


    bless $self, $class;
    $self->_initialize();
    return $self;
}


###############################################################################
#
# _initialize()
#
# Open a tmp file to store the majority of the Worksheet data. If this fails,
# for example due to write permissions, store the data in memory. This can be
# slow for large files.
#
sub _initialize {

    my $self = shift;

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _close()
#
# Add data to the beginning of the workbook (note the reverse order)
# and to the end of the workbook.
#
sub _close {

    my $self = shift;
    my $sheetnames = shift;
    my $num_sheets = scalar @$sheetnames;

    $self->_write_xml_start_tag(1, 1, 0, 'Worksheet', 'ss:Name', $self->{_name});


    # Update the sheet dimensions
    $self->_store_dimensions();


    # Write the Table element and the child Row, Cell and Data elements.
    $self->_write_xml_table();


    ################################################
    # Prepend in reverse order!!
    #

    # Prepend the sheet password
    $self->_store_password();

    # Prepend the sheet protection
    $self->_store_protect();

    # Prepend the page setup
    $self->_store_setup();

    # Prepend the bottom margin
    $self->_store_margin_bottom();

    # Prepend the top margin
    $self->_store_margin_top();

    # Prepend the right margin
    $self->_store_margin_right();

    # Prepend the left margin
    $self->_store_margin_left();

    # Prepend the page vertical centering
    $self->_store_vcenter();

    # Prepend the page horizontal centering
    $self->_store_hcenter();

    # Prepend the page footer
    $self->_store_footer();

    # Prepend the page header
    $self->_store_header();

    # Prepend the vertical page breaks
    $self->_store_vbreak();

    # Prepend the horizontal page breaks
    $self->_store_hbreak();

    # Prepend WSBOOL
    $self->_store_wsbool();

    # Prepend GRIDSET
    $self->_store_gridset();

    # Prepend GUTS
    $self->_store_guts();

    # Prepend PRINTGRIDLINES
    $self->_store_print_gridlines();

    # Prepend PRINTHEADERS
    $self->_store_print_headers();

    # Prepend EXTERNSHEET references
    for (my $i = $num_sheets; $i > 0; $i--) {
        my $sheetname = @{$sheetnames}[$i-1];
        $self->_store_externsheet($sheetname);
    }

    # Prepend the EXTERNCOUNT of external references.
    $self->_store_externcount($num_sheets);

    # Prepend the COLINFO records if they exist
    if (@{$self->{_colinfo}}){
        while (@{$self->{_colinfo}}) {
            my $arrayref = pop @{$self->{_colinfo}};
            $self->_store_colinfo(@$arrayref);
        }
        $self->_store_defcol();
    }

    #
    # End of prepend. Read upwards from here.
    ################################################

    # Append
    $self->_store_window2();
    $self->_store_zoom();
    $self->_store_panes(@{$self->{_panes}}) if @{$self->{_panes}};
    $self->_store_selection(@{$self->{_selection}});


    # Close Workbook tag. WriteExcel _store_eof().
    $self->_write_xml_end_tag(1, 1, 1, 'Worksheet');



}


###############################################################################
#
# get_name().
#
# Retrieve the worksheet name.
#
sub get_name {

    my $self    = shift;

    return $self->{_name};
}


###############################################################################
#
# get_data().
#
# Retrieves data from memory in one chunk, or from disk in $buffer
# sized chunks.
#
sub get_data {

    my $self   = shift;

    # TODO Update for ExcelXML format
}


###############################################################################
#
# select()
#
# Set this worksheet as a selected worksheet, i.e. the worksheet has its tab
# highlighted.
#
sub select {

    my $self = shift;

    $self->{_selected} = 1;
}


###############################################################################
#
# activate()
#
# Set this worksheet as the active worksheet, i.e. the worksheet that is
# displayed when the workbook is opened. Also set it as selected.
#
sub activate {

    my $self = shift;

    $self->{_selected} = 1;
    ${$self->{_activesheet}} = $self->{_index};
}


###############################################################################
#
# set_first_sheet()
#
# Set this worksheet as the first visible sheet. This is necessary
# when there are a large number of worksheets and the activated
# worksheet is not visible on the screen.
#
sub set_first_sheet {

    my $self = shift;

    ${$self->{_firstsheet}} = $self->{_index};
}


###############################################################################
#
# protect($password)
#
# Set the worksheet protection flag to prevent accidental modification and to
# hide formulas if the locked and hidden format properties have been set.
#
sub protect {

    my $self = shift;

    $self->{_protect}   = 1;
    $self->{_password}  = $self->_encode_password($_[0]) if defined $_[0];

}


###############################################################################
#
# set_column($firstcol, $lastcol, $width, $format, $hidden, $level)
#
# Set the width of a single column or a range of columns.
# See also: _store_colinfo
#
sub set_column {

    my $self = shift;
    my $cell = $_[0];

    # Check for a cell reference in A1 notation and substitute row and column
    if ($cell =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);

        # Returned values $row1 and $row2 aren't required here. Remove them.
        shift  @_;       # $row1
        splice @_, 1, 1; # $row2
    }

    push @{$self->{_colinfo}}, [ @_ ];


    # Store the col sizes for use when calculating image vertices taking
    # hidden columns into account. Also store the column formats.
    #
    return if @_ < 3; # Ensure at least $firstcol, $lastcol and $width

    my $width  = $_[4] ? 0 : $_[2]; # Set width to zero if column is hidden
    my $format = $_[3];

    my ($firstcol, $lastcol) = @_;

    foreach my $col ($firstcol .. $lastcol) {
        $self->{_col_sizes}->{$col}   = $width;
        $self->{_col_formats}->{$col} = $format if defined $format;
    }
}


###############################################################################
#
# set_selection()
#
# Set which cell or cells are selected in a worksheet: see also the
# sub _store_selection
#
sub set_selection {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    $self->{_selection} = [ @_ ];
}


###############################################################################
#
# freeze_panes()
#
# Set panes and mark them as frozen. See also _store_panes().
#
sub freeze_panes {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    $self->{_frozen} = 1;
    $self->{_panes}  = [ @_ ];
}


###############################################################################
#
# thaw_panes()
#
# Set panes and mark them as unfrozen. See also _store_panes().
#
sub thaw_panes {

    my $self = shift;

    $self->{_frozen} = 0;
    $self->{_panes}  = [ @_ ];
}


###############################################################################
#
# set_portrait()
#
# Set the page orientation as portrait.
#
sub set_portrait {

    my $self = shift;

    $self->{_orientation} = 1;
}


###############################################################################
#
# set_landscape()
#
# Set the page orientation as landscape.
#
sub set_landscape {

    my $self = shift;

    $self->{_orientation} = 0;
}


###############################################################################
#
# set_paper()
#
# Set the paper type. Ex. 1 = US Letter, 9 = A4
#
sub set_paper {

    my $self = shift;

    $self->{_paper_size} = $_[0] || 0;
}


###############################################################################
#
# set_header()
#
# Set the page header caption and optional margin.
#
sub set_header {

    my $self   = shift;
    my $string = $_[0] || '';

    if (length $string >= 255) {
        carp 'Header string must be less than 255 characters';
        return;
    }

    $self->{_header}      = $string;
    $self->{_margin_head} = $_[1] || 0.50;
}


###############################################################################
#
# set_footer()
#
# Set the page footer caption and optional margin.
#
sub set_footer {

    my $self   = shift;
    my $string = $_[0] || '';

    if (length $string >= 255) {
        carp 'Footer string must be less than 255 characters';
        return;
    }


    $self->{_footer}      = $string;
    $self->{_margin_foot} = $_[1] || 0.50;
}


###############################################################################
#
# center_horizontally()
#
# Center the page horizontally.
#
sub center_horizontally {

    my $self = shift;

    if (defined $_[0]) {
        $self->{_hcenter} = $_[0];
    }
    else {
        $self->{_hcenter} = 1;
    }
}


###############################################################################
#
# center_vertically()
#
# Center the page horinzontally.
#
sub center_vertically {

    my $self = shift;

    if (defined $_[0]) {
        $self->{_vcenter} = $_[0];
    }
    else {
        $self->{_vcenter} = 1;
    }
}


###############################################################################
#
# set_margins()
#
# Set all the page margins to the same value in inches.
#
sub set_margins {

    my $self = shift;

    $self->set_margin_left($_[0]);
    $self->set_margin_right($_[0]);
    $self->set_margin_top($_[0]);
    $self->set_margin_bottom($_[0]);
}


###############################################################################
#
# set_margins_LR()
#
# Set the left and right margins to the same value in inches.
#
sub set_margins_LR {

    my $self = shift;

    $self->set_margin_left($_[0]);
    $self->set_margin_right($_[0]);
}


###############################################################################
#
# set_margins_TB()
#
# Set the top and bottom margins to the same value in inches.
#
sub set_margins_TB {

    my $self = shift;

    $self->set_margin_top($_[0]);
    $self->set_margin_bottom($_[0]);
}


###############################################################################
#
# set_margin_left()
#
# Set the left margin in inches.
#
sub set_margin_left {

    my $self = shift;

    $self->{_margin_left} = defined $_[0] ? $_[0] : 0.75;
}


###############################################################################
#
# set_margin_right()
#
# Set the right margin in inches.
#
sub set_margin_right {

    my $self = shift;

    $self->{_margin_right} = defined $_[0] ? $_[0] : 0.75;
}


###############################################################################
#
# set_margin_top()
#
# Set the top margin in inches.
#
sub set_margin_top {

    my $self = shift;

    $self->{_margin_top} = defined $_[0] ? $_[0] : 1.00;
}


###############################################################################
#
# set_margin_bottom()
#
# Set the bottom margin in inches.
#
sub set_margin_bottom {

    my $self = shift;

    $self->{_margin_bottom} = defined $_[0] ? $_[0] : 1.00;
}


###############################################################################
#
# repeat_rows($first_row, $last_row)
#
# Set the rows to repeat at the top of each printed page. See also the
# _store_name_xxxx() methods in Workbook.pm.
#
sub repeat_rows {

    my $self = shift;

    $self->{_title_rowmin}  = $_[0];
    $self->{_title_rowmax}  = $_[1] || $_[0]; # Second row is optional
}


###############################################################################
#
# repeat_columns($first_col, $last_col)
#
# Set the columns to repeat at the left hand side of each printed page.
# See also the _store_names() methods in Workbook.pm.
#
sub repeat_columns {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);

        # Returned values $row1 and $row2 aren't required here. Remove them.
        shift  @_;       # $row1
        splice @_, 1, 1; # $row2
    }

    $self->{_title_colmin}  = $_[0];
    $self->{_title_colmax}  = $_[1] || $_[0]; # Second col is optional
}


###############################################################################
#
# print_area($first_row, $first_col, $last_row, $last_col)
#
# Set the area of each worksheet that will be printed. See also the
# _store_names() methods in Workbook.pm.
#
sub print_area {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    return if @_ != 4; # Require 4 parameters

    $self->{_print_rowmin} = $_[0];
    $self->{_print_colmin} = $_[1];
    $self->{_print_rowmax} = $_[2];
    $self->{_print_colmax} = $_[3];
}


###############################################################################
#
# hide_gridlines()
#
# Set the option to hide gridlines on the screen and the printed page.
# There are two ways of doing this in the Excel BIFF format: The first is by
# setting the DspGrid field of the WINDOW2 record, this turns off the screen
# and subsequently the print gridline. The second method is to via the
# PRINTGRIDLINES and GRIDSET records, this turns off the printed gridlines
# only. The first method is probably sufficient for most cases. The second
# method is supported for backwards compatibility. Porters take note.
#
sub hide_gridlines {

    my $self   = shift;
    my $option = $_[0];

    $option = 1 unless defined $option; # Default to hiding printed gridlines

    if ($option == 0) {
        $self->{_print_gridlines}  = 1; # 1 = display, 0 = hide
        $self->{_screen_gridlines} = 1;
    }
    elsif ($option == 1) {
        $self->{_print_gridlines}  = 0;
        $self->{_screen_gridlines} = 1;
    }
    else {
        $self->{_print_gridlines}  = 0;
        $self->{_screen_gridlines} = 0;
    }
}


###############################################################################
#
# print_row_col_headers()
#
# Set the option to print the row and column headers on the printed page.
# See also the _store_print_headers() method below.
#
sub print_row_col_headers {

    my $self = shift;

    if (defined $_[0]) {
        $self->{_print_headers} = $_[0];
    }
    else {
        $self->{_print_headers} = 1;
    }
}


###############################################################################
#
# fit_to_pages($width, $height)
#
# Store the vertical and horizontal number of pages that will define the
# maximum area printed. See also _store_setup() and _store_wsbool() below.
#
sub fit_to_pages {

    my $self = shift;

    $self->{_fit_page}      = 1;
    $self->{_fit_width}     = $_[0] || 0;
    $self->{_fit_height}    = $_[1] || 0;
}


###############################################################################
#
# set_h_pagebreaks(@breaks)
#
# Store the horizontal page breaks on a worksheet.
#
sub set_h_pagebreaks {

    my $self = shift;

    push @{$self->{_hbreaks}}, @_;
}


###############################################################################
#
# set_v_pagebreaks(@breaks)
#
# Store the vertical page breaks on a worksheet.
#
sub set_v_pagebreaks {

    my $self = shift;

    push @{$self->{_vbreaks}}, @_;
}


###############################################################################
#
# set_zoom($scale)
#
# Set the worksheet zoom factor.
#
sub set_zoom {

    my $self  = shift;
    my $scale = $_[0] || 100;

    # Confine the scale to Excel's range
    if ($scale < 10 or $scale > 400) {
        carp "Zoom factor $scale outside range: 10 <= zoom <= 400";
        $scale = 100;
    }

    $self->{_zoom} = int $scale;
}


###############################################################################
#
# set_print_scale($scale)
#
# Set the scale factor for the printed page.
#
sub set_print_scale {

    my $self  = shift;
    my $scale = $_[0] || 100;

    # Confine the scale to Excel's range
    if ($scale < 10 or $scale > 400) {
        carp "Print scale $scale outside range: 10 <= zoom <= 400";
        $scale = 100;
    }

    # Turn off "fit to page" option
    $self->{_fit_page}    = 0;

    $self->{_print_scale} = int $scale;
}


###############################################################################
#
# keep_leading_zeros()
#
# Causes the write() method to treat integers with a leading zero as a string.
# This ensures that any leading zeros such, as in zip codes, are maintained.
#
sub keep_leading_zeros {

    my $self = shift;

    if (defined $_[0]) {
        $self->{_leading_zeros} = $_[0];
    }
    else {
        $self->{_leading_zeros} = 1;
    }
}


###############################################################################
#
# write($row, $col, $token, $format)
#
# Parse $token and call appropriate write method. $row and $column are zero
# indexed. $format is optional.
#
# Returns: return value of called subroutine
#
sub write {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    my $token = $_[2];

    # Handle undefs as blanks
    $token = '' unless defined $token;

    # Match an array ref.
    if (ref $token eq "ARRAY") {
        return $self->write_row(@_);
    }
    # Match integer with leading zero(s)
    elsif ($self->{_leading_zeros} and $token =~ /^0\d+$/) {
        return $self->write_string(@_);
    }
    # Match number
    elsif ($token =~ /^([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?$/) {
        return $self->write_number(@_);
    }
    # Match http, https or ftp URL
    elsif ($token =~ m|^[fh]tt?ps?://|) {
        return $self->write_url(@_);
    }
    # Match mailto:
    elsif ($token =~ m/^mailto:/) {
        return $self->write_url(@_);
    }
    # Match internal or external sheet link
    elsif ($token =~ m[^(?:in|ex)ternal:]) {
        return $self->write_url(@_);
    }
    # Match formula
    elsif ($token =~ /^=/) {
        return $self->write_formula(@_);
    }
    # Match blank
    elsif ($token eq '') {
        splice @_, 2, 1; # remove the empty string from the parameter list
        return $self->write_blank(@_);
    }
    # Default: match string
    else {
        return $self->write_string(@_);
    }
}


###############################################################################
#
# write_row($row, $col, $array_ref, $format)
#
# Write a row of data starting from ($row, $col). Call write_col() if any of
# the elements of the array ref are in turn array refs. This allows the writing
# of 1D or 2D arrays of data in one go.
#
# Returns: the first encountered error value or zero for no errors
#
sub write_row {

    my $self = shift;


    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    # Catch non array refs passed by user.
    if (ref $_[2] ne 'ARRAY') {
        croak "Not an array ref in call to write_row()$!";
    }

    my $row     = shift;
    my $col     = shift;
    my $tokens  = shift;
    my @options = @_;
    my $error   = 0;
    my $ret;

    foreach my $token (@$tokens) {

        # Check for nested arrays
        if (ref $token eq "ARRAY") {
            $ret = $self->write_col($row, $col, $token, @options);
        } else {
            $ret = $self->write    ($row, $col, $token, @options);
        }

        # Return only the first error encountered, if any.
        $error ||= $ret;
        $col++;
    }

    return $error;
}


###############################################################################
#
# write_col($row, $col, $array_ref, $format)
#
# Write a column of data starting from ($row, $col). Call write_row() if any of
# the elements of the array ref are in turn array refs. This allows the writing
# of 1D or 2D arrays of data in one go.
#
# Returns: the first encountered error value or zero for no errors
#
sub write_col {

    my $self = shift;


    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    # Catch non array refs passed by user.
    if (ref $_[2] ne 'ARRAY') {
        croak "Not an array ref in call to write_row()$!";
    }

    my $row     = shift;
    my $col     = shift;
    my $tokens  = shift;
    my @options = @_;
    my $error   = 0;
    my $ret;

    foreach my $token (@$tokens) {

        # write() will deal with any nested arrays
        $ret = $self->write($row, $col, $token, @options);

        # Return only the first error encountered, if any.
        $error ||= $ret;
        $row++;
    }

    return $error;
}


###############################################################################
#
# write_comment($row, $col, $comment)
#
# Write a comment to the specified row and column (zero indexed). The maximum
# comment size is 30831 chars. Excel5 probably accepts 32k-1 chars. However, it
# can only display 30831 chars. Excel 7 and 2000 will crash above 32k-1.
#
# In Excel 5 a comment is referred to as a NOTE.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long comment truncated to 30831 chars
#
sub write_comment {

    my $self      = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }


    if (@_ < 3) { return -1 } # Check the number of args

    my $row       = $_[0];
    my $col       = $_[1];
    my $str       = $_[2];
    my $strlen    = length($_[2]);
    my $str_error = 0;
    my $str_max   = 30831;
    my $note_max  = 2048;

    if ($row >= $self->{_xls_rowmax}) { return -2 }
    if ($col >= $self->{_xls_colmax}) { return -2 }
    if ($row <  $self->{_dim_rowmin}) { $self->{_dim_rowmin} = $row }
    if ($row >  $self->{_dim_rowmax}) { $self->{_dim_rowmax} = $row }
    if ($col <  $self->{_dim_colmin}) { $self->{_dim_colmin} = $col }
    if ($col >  $self->{_dim_colmax}) { $self->{_dim_colmax} = $col }

    # String must be <= 30831 chars
    if ($strlen > $str_max) {
        $str       = substr($str, 0, $str_max);
        $strlen    = $str_max;
        $str_error = -3;
    }

    # A comment can be up to 30831 chars broken into segments of 2048 chars.
    # The first NOTE record contains the total string length. Each subsequent
    # NOTE record contains the length of that segment.
    #
    my $comment = substr($str, 0, $note_max, '');
    $self->_store_comment($row, $col, $comment, $strlen); # First NOTE

    # Subsequent NOTE records
    while ($str) {
        $comment = substr($str, 0, $note_max, '');
        $strlen  = length($comment);
        # Row is -1 to indicate a continuation NOTE
        $self->_store_comment(-1, 0, $comment, $strlen);
    }

    return $str_error;
}


###############################################################################
#
# _XF()
#
# Returns an index to the XF record in the workbook.
#
# Note: this is a function, not a method.
#
sub _XF {

    my $self   = $_[0];
    my $row    = $_[1];
    my $col    = $_[2];
    my $format = $_[3];

    if (ref($format)) {
        return $format->get_xf_index();
    }
    elsif (exists $self->{_row_formats}->{$row}) {
        return $self->{_row_formats}->{$row}->get_xf_index();
    }
    elsif (exists $self->{_col_formats}->{$col}) {
        return $self->{_col_formats}->{$col}->get_xf_index();
    }
    else {
        return 0; # 0x0F for Spreadsheet::WriteExcel
    }
}


###############################################################################
###############################################################################
#
# Internal methods
#



###############################################################################
#
# _substitute_cellref()
#
# Substitute an Excel cell reference in A1 notation for  zero based row and
# column values in an argument list.
#
# Ex: ("A4", "Hello") is converted to (3, 0, "Hello").
#
sub _substitute_cellref {

    my $self = shift;
    my $cell = uc(shift);

    # Convert a column range: 'A:A' or 'B:G'.
    # A range such as A:A is equivalent to A1:A16384, so add rows as required
    if ($cell =~ /\$?([A-I]?[A-Z]):\$?([A-I]?[A-Z])/) {
        my ($row1, $col1) =  $self->_cell_to_rowcol($1 .'1');
        my ($row2, $col2) =  $self->_cell_to_rowcol($2 .'16384');
        return $row1, $col1, $row2, $col2, @_;
    }

    # Convert a cell range: 'A1:B7'
    if ($cell =~ /\$?([A-I]?[A-Z]\$?\d+):\$?([A-I]?[A-Z]\$?\d+)/) {
        my ($row1, $col1) =  $self->_cell_to_rowcol($1);
        my ($row2, $col2) =  $self->_cell_to_rowcol($2);
        return $row1, $col1, $row2, $col2, @_;
    }

    # Convert a cell reference: 'A1' or 'AD2000'
    if ($cell =~ /\$?([A-I]?[A-Z]\$?\d+)/) {
        my ($row1, $col1) =  $self->_cell_to_rowcol($1);
        return $row1, $col1, @_;

    }

    croak("Unknown cell reference $cell");
}


###############################################################################
#
# _cell_to_rowcol($cell_ref)
#
# Convert an Excel cell reference in A1 notation to a zero based row and column
# reference; converts C1 to (0, 2).
#
# See also: http://www.perlmonks.org/index.pl?node_id=270352
#
# Returns: ($row, $col, $row_absolute, $col_absolute)
#
#
sub _cell_to_rowcol {

    my $self =  shift;

    my $cell =  $_[0];
       $cell =~ /(\$?)([A-I]?[A-Z])(\$?)(\d+)/;

    my $col_abs = $1 eq "" ? 0 : 1;
    my $col     = $2;
    my $row_abs = $3 eq "" ? 0 : 1;
    my $row     = $4;

    # Convert base26 column string to number
    # All your Base are belong to us.
    my @chars  = split //, $col;
    my $expn   = 0;
    $col       = 0;

    while (@chars) {
        my $char = pop(@chars); # LS char first
        $col += (ord($char) -ord('A') +1) * (26**$expn);
        $expn++;
    }

    # Convert 1-index to zero-index
    $row--;
    $col--;

    # TODO Check row and column range
    return $row, $col, $row_abs, $col_abs;
}


###############################################################################
#
# _sort_pagebreaks()
#
#
# This is an internal method that is used to filter elements of the array of
# pagebreaks used in the _store_hbreak() and _store_vbreak() methods. It:
#   1. Removes duplicate entries from the list.
#   2. Sorts the list.
#   3. Removes 0 from the list if present.
#
sub _sort_pagebreaks {

    my $self= shift;

    my %hash;
    my @array;

    @hash{@_} = undef;                       # Hash slice to remove duplicates
    @array    = sort {$a <=> $b} keys %hash; # Numerical sort
    shift @array if $array[0] == 0;          # Remove zero

    # 1000 vertical pagebreaks appears to be an internal Excel 5 limit.
    # It is slightly higher in Excel 97/200, approx. 1026
    splice(@array, 1000) if (@array > 1000);

    return @array
}


###############################################################################
#
# _encode_password($password)
#
# Based on the algorithm provided by Daniel Rentz of OpenOffice.
#
#
sub _encode_password {

    use integer;

    my $self      = shift;
    my $plaintext = $_[0];
    my $password;
    my $count;
    my @chars;
    my $i = 0;

    $count = @chars = split //, $plaintext;

    foreach my $char (@chars) {
        my $low_15;
        my $high_15;
        $char     = ord($char) << ++$i;
        $low_15   = $char & 0x7fff;
        $high_15  = $char & 0x7fff << 15;
        $high_15  = $high_15 >> 15;
        $char     = $low_15 | $high_15;
    }

    $password  = 0x0000;
    $password ^= $_ for @chars;
    $password ^= $count;
    $password ^= 0xCE4B;

    return $password;
}


###############################################################################
#
# outline_settings($visible, $symbols_below, $symbols_right, $auto_style)
#
# This method sets the properties for outlining and grouping. The defaults
# correspond to Excel's defaults.
#
sub outline_settings {

    my $self                = shift;

    $self->{_outline_on}    = defined $_[0] ? $_[0] : 1;
    $self->{_outline_below} = defined $_[1] ? $_[1] : 1;
    $self->{_outline_right} = defined $_[2] ? $_[2] : 1;
    $self->{_outline_style} =         $_[3] || 0;

    # Ensure this is a boolean vale for Window2
    $self->{_outline_on}    = 1 if $self->{_outline_on};
}




###############################################################################
###############################################################################
#
# BIFF RECORDS
#


###############################################################################
#
# write_number($row, $col, $num, $format)
#
# Write a double to the specified row and column (zero indexed).
# An integer can be written as a double. Excel will display an
# integer. $format is optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_number {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 3) { return -1 }                    # Check the number of args


    my $row     = $_[0];                         # Zero indexed row
    my $col     = $_[1];                         # Zero indexed column
    my $num     = $_[2];
    my $xf      = _XF($self, $row, $col, $_[3]); # The cell format
    my $type    = $self->{_datatypes}->{Number}; # The data type

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

    $self->{_table}->[$row]->[$col] = [$type, $num, $xf];

    return 0;
}


###############################################################################
#
# write_string ($row, $col, $string, $format)
#
# Write a string to the specified row and column (zero indexed).
# NOTE: there is an Excel 5 defined limit of 255 characters.
# $format is optional.
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 255 chars
#
sub write_string {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 3) { return -1 }                    # Check the number of args

    my $row     = $_[0];                         # Zero indexed row
    my $col     = $_[1];                         # Zero indexed column
    my $str     = $_[2];
    my $xf      = _XF($self, $row, $col, $_[3]); # The cell format
    my $type    = $self->{_datatypes}->{String}; # The data type

    my $str_error = 0;

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

    if (length $str > $self->{_xls_strmax}) { # LABEL must be < 32767 chars
        $str       = substr($str, 0, $self->{_xls_strmax});
        $str_error = -3;
    }


    $self->{_table}->[$row]->[$col] = [$type, $str, $xf];

    return $str_error;
}


###############################################################################
#
# write_blank($row, $col, $format)
#
# Write a blank cell to the specified row and column (zero indexed).
# A blank cell is used to specify formatting without adding a string
# or a number.
#
# A blank cell without a format serves no purpose. Therefore, we don't write
# a BLANK record unless a format is specified. This is mainly an optimisation
# for the write_row() and write_col() methods.
#
# Returns  0 : normal termination (including no format)
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_blank {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    # Check the number of args
    return -1 if @_ < 2;

    # Don't write a blank cell unless it has a format
    return 0 if not defined $_[2];


    my $record  = 0x0201;                        # Record identifier
    my $length  = 0x0006;                        # Number of bytes to follow

    my $row     = $_[0];                         # Zero indexed row
    my $col     = $_[1];                         # Zero indexed column
    my $xf      = _XF($self, $row, $col, $_[2]); # The cell format
    my $type    = $self->{_datatypes}->{Blank};  # The data type

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

    $self->{_table}->[$row]->[$col] = [$type, undef, $xf];

    return 0;
}


###############################################################################
#
# write_formula($row, $col, $formula, $format)
#
# Write a formula to the specified row and column (zero indexed).
# The textual representation of the formula is passed to the parser in
# Formula.pm which returns a packed binary string.
#
# $format is optional.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub write_formula{

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 3) { return -1 }   # Check the number of args

    my $record  = 0x0006;     # Record identifier
    my $length;                 # Bytes to follow

    my $row     = $_[0];      # Zero indexed row
    my $col     = $_[1];      # Zero indexed column
    my $formula = $_[2];      # The formula text string


    my $xf      = _XF($self, $row, $col, $_[3]);  # The cell format
    my $type    = $self->{_datatypes}->{Formula}; # The data type



    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);

    # Add the = sign if it doesn't exist
    $formula    =~ s{^[^=]}{=};

    # Convert A1 style references in the formula to R1C1 references
    $formula    = $self->_convert_formula($row, $col, $formula);

    $self->{_table}->[$row]->[$col] = [$type, $formula, $xf];

    return 0;
}


###############################################################################
#
# store_formula($formula)
#
# Pre-parse a formula. This is used in conjunction with repeat_formula()
# to repetitively rewrite a formula without re-parsing it.
#
sub store_formula{

    my $self    = shift;
    my $formula = $_[0];      # The formula text string

    # Strip the = sign at the beginning of the formula string
    $formula    =~ s(^=)();

    # Parse the formula using the parser in Formula.pm
    my $parser  = $self->{_parser};

    # In order to raise formula errors from the point of view of the calling
    # program we use an eval block and re-raise the error from here.
    #
    my @tokens;
    eval { @tokens = $parser->parse_formula($formula) };

    if ($@) {
        $@ =~ s/\n$//;  # Strip the \n used in the Formula.pm die()
        croak $@;       # Re-raise the error
    }


    # Return the parsed tokens in an anonymous array
    return [@tokens];
}


###############################################################################
#
# repeat_formula($row, $col, $formula, $format, ($pattern => $replacement,...))
#
# Write a formula to the specified row and column (zero indexed) by
# substituting $pattern $replacement pairs in the $formula created via
# store_formula(). This allows the user to repetitively rewrite a formula
# without the significant overhead of parsing.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#
sub repeat_formula {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    if (@_ < 2) { return -1 }   # Check the number of args

    my $record      = 0x0006;   # Record identifier
    my $length;                 # Bytes to follow

    my $row         = shift;    # Zero indexed row
    my $col         = shift;    # Zero indexed column
    my $formula_ref = shift;    # Array ref with formula tokens
    my $format       = shift;   # XF format
    my @pairs       = @_;       # Pattern/replacement pairs


    # Enforce an even number of arguments in the pattern/replacement list
    croak "Odd number of elements in pattern/replacement list" if @pairs %2;

    # Check that $formula is an array ref
    croak "Not a valid formula" if ref $formula_ref ne 'ARRAY';

    my @tokens  = @$formula_ref;

    # Ensure that there are tokens to substitute
    croak "No tokens in formula" unless @tokens;

    while (@pairs) {
        my $pattern = shift @pairs;
        my $replace = shift @pairs;

        foreach my $token (@tokens) {
            last if $token =~ s/$pattern/$replace/;
        }
    }


    # Change the parameters in the formula cached by the Formula.pm object
    my $parser    = $self->{_parser};
    my $formula   = $parser->parse_tokens(@tokens);

    croak "Unrecognised token in formula" unless defined $formula;


    # Excel normally stores the last calculated value of the formula in $num.
    # Clearly we are not in a position to calculate this a priori. Instead
    # we set $num to zero and set the option flags in $grbit to ensure
    # automatic calculation of the formula when the file is opened.
    #
    my $xf        = _XF($self, $row, $col, $format); # The cell format
    my $num       = 0x00;                            # Current value of formula
    my $grbit     = 0x03;                            # Option flags
    my $chn       = 0x0000;                          # Must be zero

    # Check that row and col are valid and store max and min values
    return -2 if $self->_check_dimensions($row, $col);


    # TODO Update for ExcelXML format

    return 0;
}


###############################################################################
#
# write_url($row, $col, $url, $string, $format)
#
# Write a hyperlink. This is comprised of two elements: the visible label and
# the invisible link. The visible label is the same as the link unless an
# alternative string is specified. The label is written using the
# write_string() method. Therefore the 255 characters string limit applies.
# $string and $format are optional and their order is interchangeable.
#
# The hyperlink can be to a http, ftp, mail, internal sheet, or external
# directory url.
#
# Returns  0 : normal termination
#         -1 : insufficient number of arguments
#         -2 : row or column out of range
#         -3 : long string truncated to 255 chars
#
sub write_url {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    # Check the number of args
    return -1 if @_ < 3;

    # Add start row and col to arg list
    return $self->write_url_range($_[0], $_[1], @_);
}


###############################################################################
#
# write_url_range($row1, $col1, $row2, $col2, $url, $string, $format)
#
# This is the more general form of write_url(). It allows a hyperlink to be
# written to a range of cells. This function also decides the type of hyperlink
# to be written. These are either, Web (http, ftp, mailto), Internal
# (Sheet1!A1) or external ('c:\temp\foo.xls#Sheet1!A1').
#
# See also write_url() above for a general description and return values.
#
sub write_url_range {

    my $self = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    # Check the number of args
    return -1 if @_ < 5;


    # Reverse the order of $string and $format if necessary. We work on a copy
    # in order to protect the callers args. We don't use "local @_" in case of
    # perl50005 threads.
    #
    my @args = @_;

    ($args[5], $args[6]) = ($args[6], $args[5]) if ref $args[5];

    my $url = $args[4];


    # Check for internal/external sheet links or default to web link
    return $self->_write_url_internal(@args) if $url =~ m[^internal:];
    return $self->_write_url_external(@args) if $url =~ m[^external:];
    return $self->_write_url_web(@args);
}



###############################################################################
#
# _write_url_web($row1, $col1, $row2, $col2, $url, $string, $format)
#
# Used to write http, ftp and mailto hyperlinks.
# The link type ($options) is 0x03 is the same as absolute dir ref without
# sheet. However it is differentiated by the $unknown2 data stream.
#
# See also write_url() above for a general description and return values.
#
sub _write_url_web {

    my $self    = shift;

    my $record      = 0x01B8;                       # Record identifier
    my $length      = 0x00000;                      # Bytes to follow

    my $row1        = $_[0];                        # Start row
    my $col1        = $_[1];                        # Start column
    my $row2        = $_[2];                        # End row
    my $col2        = $_[3];                        # End column
    my $url         = $_[4];                        # URL string
    my $str         = $_[5];                        # Alternative label
    my $xf          = $_[6] || $self->{_url_format};# The cell format


    # Write the visible label using the write_string() method.
    $str            = $url unless defined $str;
    my $str_error   = $self->write_string($row1, $col1, $str, $xf);
    return $str_error if $str_error == -2;


    # TODO Update for ExcelXML format

    return $str_error;
}


###############################################################################
#
# _write_url_internal($row1, $col1, $row2, $col2, $url, $string, $format)
#
# Used to write internal reference hyperlinks such as "Sheet1!A1".
#
# See also write_url() above for a general description and return values.
#
sub _write_url_internal {

    my $self    = shift;

    my $record      = 0x01B8;                       # Record identifier
    my $length      = 0x00000;                      # Bytes to follow

    my $row1        = $_[0];                        # Start row
    my $col1        = $_[1];                        # Start column
    my $row2        = $_[2];                        # End row
    my $col2        = $_[3];                        # End column
    my $url         = $_[4];                        # URL string
    my $str         = $_[5];                        # Alternative label
    my $xf          = $_[6] || $self->{_url_format};# The cell format

    # Strip URL type
    $url            =~ s[^internal:][];


    # Write the visible label
    $str            = $url unless defined $str;
    my $str_error   = $self->write_string($row1, $col1, $str, $xf);
    return $str_error if $str_error == -2;


    # TODO Update for ExcelXML format

    return $str_error;
}


###############################################################################
#
# _write_url_external($row1, $col1, $row2, $col2, $url, $string, $format)
#
# Write links to external directory names such as 'c:\foo.xls',
# c:\foo.xls#Sheet1!A1', '../../foo.xls'. and '../../foo.xls#Sheet1!A1'.
#
# Note: Excel writes some relative links with the $dir_long string. We ignore
# these cases for the sake of simpler code.
#
# See also write_url() above for a general description and return values.
#
sub _write_url_external {

    my $self    = shift;

    # Network drives are different. We will handle them separately
    # MS/Novell network drives and shares start with \\
    return $self->_write_url_external_net(@_) if $_[4] =~ m[^external:\\\\];


    my $record      = 0x01B8;                       # Record identifier
    my $length      = 0x00000;                      # Bytes to follow

    my $row1        = $_[0];                        # Start row
    my $col1        = $_[1];                        # Start column
    my $row2        = $_[2];                        # End row
    my $col2        = $_[3];                        # End column
    my $url         = $_[4];                        # URL string
    my $str         = $_[5];                        # Alternative label
    my $xf          = $_[6] || $self->{_url_format};# The cell format


    # Strip URL type and change Unix dir separator to Dos style (if needed)
    #
    $url            =~ s[^external:][];
    $url            =~ s[/][\\]g;


    # Write the visible label
    ($str = $url)   =~ s[\#][ - ] unless defined $str;
    my $str_error   = $self->write_string($row1, $col1, $str, $xf);
    return $str_error if $str_error == -2;


    # Determine if the link is relative or absolute:
    # Absolute if link starts with DOS drive specifier like C:
    # Otherwise default to 0x00 for relative link.
    #
    my $absolute    = 0x00;
       $absolute    = 0x02  if $url =~ m/^[A-Za-z]:/;


    # Determine if the link contains a sheet reference and change some of the
    # parameters accordingly.
    # Split the dir name and sheet name (if it exists)
    #
    my ($dir_long , $sheet) = split /\#/, $url;
    my $link_type           = 0x01 | $absolute;
    my $sheet_len;

    if (defined $sheet) {
        $link_type |= 0x08;
        $sheet_len  = pack("V", length($sheet) + 0x01);
        $sheet      = join("\0", split('', $sheet));
        $sheet     .= "\0\0\0";
    }
    else {
        $sheet_len  = '';
        $sheet      = '';
    }

    # TODO Update for ExcelXML format

    return $str_error;
}




###############################################################################
#
# _write_url_external_net($row1, $col1, $row2, $col2, $url, $string, $format)
#
# Write links to external MS/Novell network drives and shares such as
# '//NETWORK/share/foo.xls' and '//NETWORK/share/foo.xls#Sheet1!A1'.
#
# See also write_url() above for a general description and return values.
#
sub _write_url_external_net {

    my $self    = shift;

    my $record      = 0x01B8;                       # Record identifier
    my $length      = 0x00000;                      # Bytes to follow

    my $row1        = $_[0];                        # Start row
    my $col1        = $_[1];                        # Start column
    my $row2        = $_[2];                        # End row
    my $col2        = $_[3];                        # End column
    my $url         = $_[4];                        # URL string
    my $str         = $_[5];                        # Alternative label
    my $xf          = $_[6] || $self->{_url_format};# The cell format


    # Strip URL type and change Unix dir separator to Dos style (if needed)
    #
    $url            =~ s[^external:][];
    $url            =~ s[/][\\]g;


    # Write the visible label
    ($str = $url)   =~ s[\#][ - ] unless defined $str;
    my $str_error   = $self->write_string($row1, $col1, $str, $xf);
    return $str_error if $str_error == -2;


    # Determine if the link contains a sheet reference and change some of the
    # parameters accordingly.
    # Split the dir name and sheet name (if it exists)
    #
    my ($dir_long , $sheet) = split /\#/, $url;
    my $link_type           = 0x0103; # Always absolute
    my $sheet_len;

    if (defined $sheet) {
        $link_type |= 0x08;
        $sheet_len  = pack("V", length($sheet) + 0x01);
        $sheet      = join("\0", split('', $sheet));
        $sheet     .= "\0\0\0";
    }
    else {
        $sheet_len   = '';
        $sheet       = '';
    }

    # TODO Update for ExcelXML format

    return $str_error;
}


###############################################################################
#
# set_row($row, $height, $XF, $hidden, $level)
#
# This method is used to set the height and XF format for a row.
# Writes the  BIFF record ROW.
#
sub set_row {

    my $self        = shift;
    my $record      = 0x0208;               # Record identifier
    my $length      = 0x0010;               # Number of bytes to follow

    my $rw          = $_[0];                # Row Number
    my $colMic      = 0x0000;               # First defined column
    my $colMac      = 0x0000;               # Last defined column
    my $miyRw;                              # Row height
    my $irwMac      = 0x0000;               # Used by Excel to optimise loading
    my $reserved    = 0x0000;               # Reserved
    my $grbit       = 0x0000;               # Option flags
    my $ixfe;                               # XF index
    my $height      = $_[1];                # Format object
    my $format      = $_[2];                # Format object
    my $hidden      = $_[3] || 0;           # Hidden flag
    my $level       = $_[4] || 0;           # Outline level


    # Check for a format object
    if (ref $format) {
        $ixfe = $format->get_xf_index();
    }
    else {
        $ixfe = 0x0F;
    }


    # Set the row height in units of 1/20 of a point. Note, some heights may
    # not be obtained exactly due to rounding in Excel.
    #
    if (defined $height) {
        $miyRw = $height *20;
    }
    else {
        $miyRw = 0xff; # The default row height
    }


    # Set the limits for the outline levels (0 <= x <= 7).
    $level = 0 if $level < 0;
    $level = 7 if $level > 7;

    $self->{_outline_row_level} = $level if $level >$self->{_outline_row_level};


    # Set the options flags. fUnsynced is used to show that the font and row
    # heights are not compatible. This is usually the case for WriteExcelXML.
    # The collapsed flag 0x10 doesn't seem to be used to indicate that a row
    # is collapsed. Instead it is used to indicate that the previous row is
    # collapsed. The zero height flag, 0x20, is used to collapse a row.
    #
    $grbit |= $level;
    $grbit |= 0x0020 if $hidden;
    $grbit |= 0x0040; # fUnsynced
    $grbit |= 0x0080 if $format;
    $grbit |= 0x0100;


    # TODO Update for ExcelXML format

    # Store the row sizes for use when calculating image vertices.
    # Also store the column formats.
    #
    return if @_ < 2;# Ensure at least $row and $height

    $self->{_row_sizes}->{$_[0]}   = $height;
    $self->{_row_formats}->{$_[0]} = $format if defined $format;
}


###############################################################################
#
# _check_dimensions($row, $col)
#
# Check that $row and $col are valid and store max and min values for use in
# DIMENSIONS record. See, _store_dimensions().
#
sub _check_dimensions {

    my $self    = shift;
    my $row     = $_[0];
    my $col     = $_[1];

    if ($row >= $self->{_xls_rowmax}) { return -2 }
    if ($col >= $self->{_xls_colmax}) { return -2 }

    $self->{_dim_changed} = 1;

    if ($row <  $self->{_dim_rowmin}) { $self->{_dim_rowmin} = $row }
    if ($row >  $self->{_dim_rowmax}) { $self->{_dim_rowmax} = $row }
    if ($col <  $self->{_dim_colmin}) { $self->{_dim_colmin} = $col }
    if ($col >  $self->{_dim_colmax}) { $self->{_dim_colmax} = $col }

    return 0;
}


###############################################################################
#
# _store_dimensions()
#
# Writes Excel DIMENSIONS to define the area in which there is data.
#
sub _store_dimensions {

    my $self = shift;

    # Set the data range if data has been written to the worksheet
    if ($self->{_dim_changed}) {
        $self->{_dim_rowmax} = $self->{_dim_rowmax} +1;
        $self->{_dim_colmax} = $self->{_dim_colmax} +1;
    }
    else {
        # Special case, not data was written
        $self->{_dim_rowmax} = 0;
        $self->{_dim_colmax} = 0;
    }
}


###############################################################################
#
# _store_window2()
#
# Write BIFF record Window2.
#
sub _store_window2 {

    use integer;    # Avoid << shift bug in Perl 5.6.0 on HP-UX

    my $self           = shift;
    my $record         = 0x023E;     # Record identifier
    my $length         = 0x000A;     # Number of bytes to follow

    my $grbit          = 0x00B6;     # Option flags
    my $rwTop          = 0x0000;     # Top row visible in window
    my $colLeft        = 0x0000;     # Leftmost column visible in window
    my $rgbHdr         = 0x00000000; # Row/column heading and gridline color

    # The options flags that comprise $grbit
    my $fDspFmla       = 0;                          # 0 - bit
    my $fDspGrid       = $self->{_screen_gridlines}; # 1
    my $fDspRwCol      = 1;                          # 2
    my $fFrozen        = $self->{_frozen};           # 3
    my $fDspZeros      = 1;                          # 4
    my $fDefaultHdr    = 1;                          # 5
    my $fArabic        = 0;                          # 6
    my $fDspGuts       = $self->{_outline_on};       # 7
    my $fFrozenNoSplit = 0;                          # 0 - bit
    my $fSelected      = $self->{_selected};         # 1
    my $fPaged         = 1;                          # 2

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_defcol()
#
# Write BIFF record DEFCOLWIDTH if COLINFO records are in use.
#
sub _store_defcol {

    my $self     = shift;
    my $record   = 0x0055;      # Record identifier
    my $length   = 0x0002;      # Number of bytes to follow

    my $colwidth = 0x0008;      # Default column width

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_colinfo($firstcol, $lastcol, $width, $format, $hidden)
#
# Write BIFF record COLINFO to define column widths
#
# Note: The SDK says the record length is 0x0B but Excel writes a 0x0C
# length record.
#
sub _store_colinfo {

    my $self     = shift;
    my $record   = 0x007D;          # Record identifier
    my $length   = 0x000B;          # Number of bytes to follow

    my $colFirst = $_[0] || 0;      # First formatted column
    my $colLast  = $_[1] || 0;      # Last formatted column
    my $width    = $_[2] || 8.43;   # Col width in user units, 8.43 is default
    my $coldx;                      # Col width in internal units
    my $pixels;                     # Col width in pixels

    # Excel rounds the column width to the nearest pixel. Therefore we first
    # convert to pixels and then to the internal units. The pixel to users-units
    # relationship is different for values less than 1.
    #
    if ($width < 1) {
        $pixels = int($width *12);
    }
    else {
        $pixels = int($width *7 ) +5;
    }

    $coldx = int($pixels *256/7);


    my $ixfe;                       # XF index
    my $grbit    = 0x0000;          # Option flags
    my $reserved = 0x00;            # Reserved
    my $format   = $_[3];           # Format object
    my $hidden   = $_[4] || 0;      # Hidden flag
    my $level    = $_[5] || 0;      # Outline level


    # TODO Update for ExcelXML format


    # Set the limits for the outline levels (0 <= x <= 7).
    $level = 0 if $level < 0;
    $level = 7 if $level > 7;


    # Set the options flags.
    # The collapsed flag 0x10 doesn't seem to be used to indicate that a col
    # is collapsed. Instead it is used to indicate that the previous col is
    # collapsed. The zero height flag, 0x20, is used to collapse a col.
    #
    $grbit |= 0x0001 if $hidden;
    $grbit |= $level << 8;


    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_selection($first_row, $first_col, $last_row, $last_col)
#
# Write BIFF record SELECTION.
#
sub _store_selection {

    my $self     = shift;
    my $record   = 0x001D;                  # Record identifier
    my $length   = 0x000F;                  # Number of bytes to follow

    my $pnn      = $self->{_active_pane};   # Pane position
    my $rwAct    = $_[0];                   # Active row
    my $colAct   = $_[1];                   # Active column
    my $irefAct  = 0;                       # Active cell ref
    my $cref     = 1;                       # Number of refs

    my $rwFirst  = $_[0];                   # First row in reference
    my $colFirst = $_[1];                   # First col in reference
    my $rwLast   = $_[2] || $rwFirst;       # Last  row in reference
    my $colLast  = $_[3] || $colFirst;      # Last  col in reference

    # Swap last row/col for first row/col as necessary
    if ($rwFirst > $rwLast) {
        ($rwFirst, $rwLast) = ($rwLast, $rwFirst);
    }

    if ($colFirst > $colLast) {
        ($colFirst, $colLast) = ($colLast, $colFirst);
    }


    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_externcount($count)
#
# Write BIFF record EXTERNCOUNT to indicate the number of external sheet
# references in a worksheet.
#
# Excel only stores references to external sheets that are used in formulas.
# For simplicity we store references to all the sheets in the workbook
# regardless of whether they are used or not. This reduces the overall
# complexity and eliminates the need for a two way dialogue between the formula
# parser the worksheet objects.
#
sub _store_externcount {

    my $self     = shift;
    my $record   = 0x0016;          # Record identifier
    my $length   = 0x0002;          # Number of bytes to follow

    my $cxals    = $_[0];           # Number of external references

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_externsheet($sheetname)
#
#
# Writes the Excel BIFF EXTERNSHEET record. These references are used by
# formulas. A formula references a sheet name via an index. Since we store a
# reference to all of the external worksheets the EXTERNSHEET index is the same
# as the worksheet index.
#
sub _store_externsheet {

    my $self      = shift;

    my $record    = 0x0017;         # Record identifier
    my $length;                     # Number of bytes to follow

    my $sheetname = $_[0];          # Worksheet name
    my $cch;                        # Length of sheet name
    my $rgch;                       # Filename encoding

    # References to the current sheet are encoded differently to references to
    # external sheets.
    #
    if ($self->{_name} eq $sheetname) {
        $sheetname = '';
        $length    = 0x02;  # The following 2 bytes
        $cch       = 1;     # The following byte
        $rgch      = 0x02;  # Self reference
    }
    else {
        $length    = 0x02 + length($_[0]);
        $cch       = length($sheetname);
        $rgch      = 0x03;  # Reference to a sheet in the current workbook
    }

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_panes()
#
#
# Writes the Excel BIFF PANE record.
# The panes can either be frozen or thawed (unfrozen).
# Frozen panes are specified in terms of a integer number of rows and columns.
# Thawed panes are specified in terms of Excel's units for rows and columns.
#
sub _store_panes {

    my $self    = shift;
    my $record  = 0x0041;       # Record identifier
    my $length  = 0x000A;       # Number of bytes to follow

    my $y       = $_[0] || 0;   # Vertical split position
    my $x       = $_[1] || 0;   # Horizontal split position
    my $rwTop   = $_[2];        # Top row visible
    my $colLeft = $_[3];        # Leftmost column visible
    my $pnnAct  = $_[4];        # Active pane


    # Code specific to frozen or thawed panes.
    if ($self->{_frozen}) {
        # Set default values for $rwTop and $colLeft
        $rwTop   = $y unless defined $rwTop;
        $colLeft = $x unless defined $colLeft;
    }
    else {
        # Set default values for $rwTop and $colLeft
        $rwTop   = 0  unless defined $rwTop;
        $colLeft = 0  unless defined $colLeft;

        # Convert Excel's row and column units to the internal units.
        # The default row height is 12.75
        # The default column width is 8.43
        # The following slope and intersection values were interpolated.
        #
        $y = 20*$y      + 255;
        $x = 113.879*$x + 390;
    }


    # Determine which pane should be active. There is also the undocumented
    # option to override this should it be necessary: may be removed later.
    #
    if (not defined $pnnAct) {
        $pnnAct = 0 if ($x != 0 && $y != 0); # Bottom right
        $pnnAct = 1 if ($x != 0 && $y == 0); # Top right
        $pnnAct = 2 if ($x == 0 && $y != 0); # Bottom left
        $pnnAct = 3 if ($x == 0 && $y == 0); # Top left
    }

    $self->{_active_pane} = $pnnAct; # Used in _store_selection

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_setup()
#
# Store the page setup SETUP BIFF record.
#
sub _store_setup {

    use integer;    # Avoid << shift bug in Perl 5.6.0 on HP-UX

    my $self         = shift;
    my $record       = 0x00A1;                  # Record identifier
    my $length       = 0x0022;                  # Number of bytes to follow

    my $iPaperSize   = $self->{_paper_size};    # Paper size
    my $iScale       = $self->{_print_scale};   # Print scaling factor
    my $iPageStart   = 0x01;                    # Starting page number
    my $iFitWidth    = $self->{_fit_width};     # Fit to number of pages wide
    my $iFitHeight   = $self->{_fit_height};    # Fit to number of pages high
    my $grbit        = 0x00;                    # Option flags
    my $iRes         = 0x0258;                  # Print resolution
    my $iVRes        = 0x0258;                  # Vertical print resolution
    my $numHdr       = $self->{_margin_head};   # Header Margin
    my $numFtr       = $self->{_margin_foot};   # Footer Margin
    my $iCopies      = 0x01;                    # Number of copies


    my $fLeftToRight = 0x0;                     # Print over then down
    my $fLandscape   = $self->{_orientation};   # Page orientation
    my $fNoPls       = 0x0;                     # Setup not read from printer
    my $fNoColor     = 0x0;                     # Print black and white
    my $fDraft       = 0x0;                     # Print draft quality
    my $fNotes       = 0x0;                     # Print notes
    my $fNoOrient    = 0x0;                     # Orientation not set
    my $fUsePage     = 0x0;                     # Use custom starting page


    $grbit           = $fLeftToRight;
    $grbit          |= $fLandscape    << 1;
    $grbit          |= $fNoPls        << 2;
    $grbit          |= $fNoColor      << 3;
    $grbit          |= $fDraft        << 4;
    $grbit          |= $fNotes        << 5;
    $grbit          |= $fNoOrient     << 6;
    $grbit          |= $fUsePage      << 7;


    # TODO Update for ExcelXML format

}

###############################################################################
#
# _store_header()
#
# Store the header caption BIFF record.
#
sub _store_header {

    my $self    = shift;

    my $record  = 0x0014;               # Record identifier
    my $length;                         # Bytes to follow

    my $str     = $self->{_header};     # header string
    my $cch     = length($str);         # Length of header string
    $length     = 1 + $cch;

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_footer()
#
# Store the footer caption BIFF record.
#
sub _store_footer {

    my $self    = shift;

    my $record  = 0x0015;               # Record identifier
    my $length;                         # Bytes to follow

    my $str     = $self->{_footer};     # Footer string
    my $cch     = length($str);         # Length of footer string
    $length     = 1 + $cch;

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_hcenter()
#
# Store the horizontal centering HCENTER BIFF record.
#
sub _store_hcenter {

    my $self     = shift;

    my $record   = 0x0083;              # Record identifier
    my $length   = 0x0002;              # Bytes to follow

    my $fHCenter = $self->{_hcenter};   # Horizontal centering

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_vcenter()
#
# Store the vertical centering VCENTER BIFF record.
#
sub _store_vcenter {

    my $self     = shift;

    my $record   = 0x0084;              # Record identifier
    my $length   = 0x0002;              # Bytes to follow

    my $fVCenter = $self->{_vcenter};   # Horizontal centering

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_margin_left()
#
# Store the LEFTMARGIN BIFF record.
#
sub _store_margin_left {

    my $self    = shift;

    my $record  = 0x0026;                   # Record identifier
    my $length  = 0x0008;                   # Bytes to follow

    my $margin  = $self->{_margin_left};    # Margin in inches

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_margin_right()
#
# Store the RIGHTMARGIN BIFF record.
#
sub _store_margin_right {

    my $self    = shift;

    my $record  = 0x0027;                   # Record identifier
    my $length  = 0x0008;                   # Bytes to follow

    my $margin  = $self->{_margin_right};   # Margin in inches

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_margin_top()
#
# Store the TOPMARGIN BIFF record.
#
sub _store_margin_top {

    my $self    = shift;

    my $record  = 0x0028;                   # Record identifier
    my $length  = 0x0008;                   # Bytes to follow

    my $margin  = $self->{_margin_top};     # Margin in inches

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_margin_bottom()
#
# Store the BOTTOMMARGIN BIFF record.
#
sub _store_margin_bottom {

    my $self    = shift;

    my $record  = 0x0029;                   # Record identifier
    my $length  = 0x0008;                   # Bytes to follow

    my $margin  = $self->{_margin_bottom};  # Margin in inches

    # TODO Update for ExcelXML format
}


###############################################################################
#
# merge_cells($first_row, $first_col, $last_row, $last_col)
#
# This is an Excel97/2000 method. It is required to perform more complicated
# merging than the normal align merge in Format.pm
#
sub merge_cells {

    my $self    = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }

    my $record  = 0x00E5;                   # Record identifier
    my $length  = 0x000A;                   # Bytes to follow

    my $cref     = 1;                       # Number of refs
    my $rwFirst  = $_[0];                   # First row in reference
    my $colFirst = $_[1];                   # First col in reference
    my $rwLast   = $_[2] || $rwFirst;       # Last  row in reference
    my $colLast  = $_[3] || $colFirst;      # Last  col in reference


    # Excel doesn't allow a single cell to be merged
    return if $rwFirst == $rwLast and $colFirst == $colLast;

    # Swap last row/col with first row/col as necessary
    ($rwFirst,  $rwLast ) = ($rwLast,  $rwFirst ) if $rwFirst  > $rwLast;
    ($colFirst, $colLast) = ($colLast, $colFirst) if $colFirst > $colLast;

    # TODO Update for ExcelXML format
}


###############################################################################
#
# merge_range($first_row, $first_col, $last_row, $last_col, $string, $format)
#
# This is a wrapper to ensure correct use of the merge_cells method, i.e., write
# the first cell of the range, write the formatted blank cells in the range and
# then call the merge_cells record. Failing to do the steps in this order will
# cause Excel 97 to crash.
#
sub merge_range {

    my $self    = shift;

    # Check for a cell reference in A1 notation and substitute row and column
    if ($_[0] =~ /^\D/) {
        @_ = $self->_substitute_cellref(@_);
    }
    croak "Incorrect number of arguments" if @_ != 6;
    croak "Final argument must be a format object" unless ref $_[5];

    my $rwFirst  = $_[0];
    my $colFirst = $_[1];
    my $rwLast   = $_[2];
    my $colLast  = $_[3];
    my $string   = $_[4];
    my $format   = $_[5];


    # Set the merge_range property of the format object. For BIFF8+.
    $format->set_merge_range();

    # Excel doesn't allow a single cell to be merged
    croak "Can't merge single cell" if $rwFirst  == $rwLast and
                                       $colFirst == $colLast;

    # Swap last row/col with first row/col as necessary
    ($rwFirst,  $rwLast ) = ($rwLast,  $rwFirst ) if $rwFirst  > $rwLast;
    ($colFirst, $colLast) = ($colLast, $colFirst) if $colFirst > $colLast;

    # Write the first cell
    $self->write($rwFirst, $colFirst, $string, $format);

    # Pad out the rest of the area with formatted blank cells.
    for my $row ($rwFirst .. $rwLast) {
        for my $col ($colFirst .. $colLast) {
            next if $row == $rwFirst and $col == $colFirst;
            $self->write_blank($row, $col, $format);
        }
    }

    $self->merge_cells($rwFirst, $colFirst, $rwLast, $colLast);
}


###############################################################################
#
# _store_print_headers()
#
# Write the PRINTHEADERS BIFF record.
#
sub _store_print_headers {

    my $self        = shift;

    my $record      = 0x002a;                   # Record identifier
    my $length      = 0x0002;                   # Bytes to follow

    my $fPrintRwCol = $self->{_print_headers};  # Boolean flag

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_print_gridlines()
#
# Write the PRINTGRIDLINES BIFF record. Must be used in conjunction with the
# GRIDSET record.
#
sub _store_print_gridlines {

    my $self        = shift;

    my $record      = 0x002b;                    # Record identifier
    my $length      = 0x0002;                    # Bytes to follow

    my $fPrintGrid  = $self->{_print_gridlines}; # Boolean flag

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_gridset()
#
# Write the GRIDSET BIFF record. Must be used in conjunction with the
# PRINTGRIDLINES record.
#
sub _store_gridset {

    my $self        = shift;

    my $record      = 0x0082;                        # Record identifier
    my $length      = 0x0002;                        # Bytes to follow

    my $fGridSet    = not $self->{_print_gridlines}; # Boolean flag

    # TODO Update for ExcelXML format

}


###############################################################################
#
# _store_guts()
#
# Write the GUTS BIFF record. This is used to configure the gutter margins
# where Excel outline symbols are displayed. The visibility of the gutters is
# controlled by a flag in WSBOOL. See also _store_wsbool().
#
# We are all in the gutter but some of us are looking at the stars.
#
sub _store_guts {

    my $self        = shift;

    my $record      = 0x0080;   # Record identifier
    my $length      = 0x0008;   # Bytes to follow

    my $dxRwGut     = 0x0000;   # Size of row gutter
    my $dxColGut    = 0x0000;   # Size of col gutter

    my $row_level   = $self->{_outline_row_level};
    my $col_level   = 0;


    # Calculate the maximum column outline level. The equivalent calculation
    # for the row outline level is carried out in set_row().
    #
    foreach my $colinfo (@{$self->{_colinfo}}) {
        # Skip cols without outline level info.
        next if @{$colinfo} < 6;
        $col_level = @{$colinfo}[5] if @{$colinfo}[5] > $col_level;
    }


    # Set the limits for the outline levels (0 <= x <= 7).
    $col_level = 0 if $col_level < 0;
    $col_level = 7 if $col_level > 7;


    # The displayed level is one greater than the max outline levels
    $row_level++ if $row_level > 0;
    $col_level++ if $col_level > 0;

    # TODO Update for ExcelXML format

}


###############################################################################
#
# _store_wsbool()
#
# Write the WSBOOL BIFF record, mainly for fit-to-page. Used in conjunction
# with the SETUP record.
#
sub _store_wsbool {

    my $self        = shift;

    my $record      = 0x0081;   # Record identifier
    my $length      = 0x0002;   # Bytes to follow

    my $grbit       = 0x0000;   # Option flags

    # Set the option flags
    $grbit |= 0x0001;                            # Auto page breaks visible
    $grbit |= 0x0020 if $self->{_outline_style}; # Auto outline styles
    $grbit |= 0x0040 if $self->{_outline_below}; # Outline summary below
    $grbit |= 0x0080 if $self->{_outline_right}; # Outline summary right
    $grbit |= 0x0100 if $self->{_fit_page};      # Page setup fit to page
    $grbit |= 0x0400 if $self->{_outline_on};    # Outline symbols displayed


    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_hbreak()
#
# Write the HORIZONTALPAGEBREAKS BIFF record.
#
sub _store_hbreak {

    my $self    = shift;

    # Return if the user hasn't specified pagebreaks
    return unless @{$self->{_hbreaks}};

    # Sort and filter array of page breaks
    my @breaks  = $self->_sort_pagebreaks(@{$self->{_hbreaks}});

    my $record  = 0x001b;               # Record identifier
    my $cbrk    = scalar @breaks;       # Number of page breaks
    my $length  = ($cbrk + 1) * 2;      # Bytes to follow


    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_vbreak()
#
# Write the VERTICALPAGEBREAKS BIFF record.
#
sub _store_vbreak {

    my $self    = shift;

    # Return if the user hasn't specified pagebreaks
    return unless @{$self->{_vbreaks}};

    # Sort and filter array of page breaks
    my @breaks  = $self->_sort_pagebreaks(@{$self->{_vbreaks}});

    my $record  = 0x001a;               # Record identifier
    my $cbrk    = scalar @breaks;       # Number of page breaks
    my $length  = ($cbrk + 1) * 2;      # Bytes to follow


    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_protect()
#
# Set the Biff PROTECT record to indicate that the worksheet is protected.
#
sub _store_protect {

    my $self        = shift;

    # Exit unless sheet protection has been specified
    return unless $self->{_protect};

    my $record      = 0x0012;               # Record identifier
    my $length      = 0x0002;               # Bytes to follow

    my $fLock       = $self->{_protect};    # Worksheet is protected

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_password()
#
# Write the worksheet PASSWORD record.
#
sub _store_password {

    my $self        = shift;

    # Can't store passwords in ExcelXML

    # TODO Update for ExcelXML format
}


###############################################################################
#
# insert_bitmap($row, $col, $filename, $x, $y, $scale_x, $scale_y)
#
# Insert a 24bit bitmap image in a worksheet. The main record required is
# IMDATA but it must be proceeded by a OBJ record to define its position.
#
sub insert_bitmap {

    my $self        = shift;

    # Can't store images in ExcelXML

    # TODO Update for ExcelXML format

}


###############################################################################
#
#  _position_image()
#
# Calculate the vertices that define the position of the image as required by
# the OBJ record.
#
#         +------------+------------+
#         |     A      |      B     |
#   +-----+------------+------------+
#   |     |(x1,y1)     |            |
#   |  1  |(A1)._______|______      |
#   |     |    |              |     |
#   |     |    |              |     |
#   +-----+----|    BITMAP    |-----+
#   |     |    |              |     |
#   |  2  |    |______________.     |
#   |     |            |        (B2)|
#   |     |            |     (x2,y2)|
#   +---- +------------+------------+
#
# Example of a bitmap that covers some of the area from cell A1 to cell B2.
#
# Based on the width and height of the bitmap we need to calculate 8 vars:
#     $col_start, $row_start, $col_end, $row_end, $x1, $y1, $x2, $y2.
# The width and height of the cells are also variable and have to be taken into
# account.
# The values of $col_start and $row_start are passed in from the calling
# function. The values of $col_end and $row_end are calculated by subtracting
# the width and height of the bitmap from the width and height of the
# underlying cells.
# The vertices are expressed as a percentage of the underlying cell width as
# follows (rhs values are in pixels):
#
#       x1 = X / W *1024
#       y1 = Y / H *256
#       x2 = (X-1) / W *1024
#       y2 = (Y-1) / H *256
#
#       Where:  X is distance from the left side of the underlying cell
#               Y is distance from the top of the underlying cell
#               W is the width of the cell
#               H is the height of the cell
#
# Note: the SDK incorrectly states that the height should be expressed as a
# percentage of 1024.
#
sub _position_image {

    my $self = shift;

    # Can't store images in ExcelXML

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _size_col($col)
#
# Convert the width of a cell from user's units to pixels. Excel rounds the
# column width to the nearest pixel. If the width hasn't been set by the user
# we use the default value. If the column is hidden we use a value of zero.
#
sub _size_col {

    my $self = shift;
    my $col  = $_[0];

    # Look up the cell value to see if it has been changed
    if (exists $self->{_col_sizes}->{$col}) {
        my $width = $self->{_col_sizes}->{$col};

        # The relationship is different for user units less than 1.
        if ($width < 1) {
            return int($width *12);
        }
        else {
            return int($width *7 ) +5;
        }
    }
    else {
        return 64;
    }
}


###############################################################################
#
# _size_row($row)
#
# Convert the height of a cell from user's units to pixels. By interpolation
# the relationship is: y = 4/3x. If the height hasn't been set by the user we
# use the default value. If the row is hidden we use a value of zero. (Not
# possible to hide row yet).
#
sub _size_row {

    my $self = shift;
    my $row  = $_[0];

    # Look up the cell value to see if it has been changed
    if (exists $self->{_row_sizes}->{$row}) {
        if ($self->{_row_sizes}->{$row} == 0) {
            return 0;
        }
        else {
            return int (4/3 * $self->{_row_sizes}->{$row});
        }
    }
    else {
        return 17;
    }
}


###############################################################################
#
# _store_obj_picture(   $col_start, $x1,
#                       $row_start, $y1,
#                       $col_end,   $x2,
#                       $row_end,   $y2 )
#
# Store the OBJ record that precedes an IMDATA record. This could be generalise
# to support other Excel objects.
#
sub _store_obj_picture {

    my $self        = shift;

    # Can't store images in ExcelXML

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _process_bitmap()
#
# Convert a 24 bit bitmap into the modified internal format used by Windows.
# This is described in BITMAPCOREHEADER and BITMAPCOREINFO structures in the
# MSDN library.
#
sub _process_bitmap {

    my $self   = shift;

    # Can't store images in ExcelXML

    # TODO Update for ExcelXML format

}


###############################################################################
#
# _store_zoom($zoom)
#
#
# Store the window zoom factor. This should be a reduced fraction but for
# simplicity we will store all fractions with a numerator of 100.
#
sub _store_zoom {

    my $self        = shift;

    # If scale is 100 we don't need to write a record
    return if $self->{_zoom} == 100;

    my $record      = 0x00A0;               # Record identifier
    my $length      = 0x0004;               # Bytes to follow

    # TODO Update for ExcelXML format
}


###############################################################################
#
# _store_comment
#
# Store the Excel 5 NOTE record. This format is not compatible with the Excel 7
# record.
#
sub _store_comment {

    my $self      = shift;
    if (@_ < 3) { return -1 }

    my $record    = 0x001C;                 # Record identifier
    my $length ;                            # Bytes to follow

    my $row       = $_[0];                  # Zero indexed row
    my $col       = $_[1];                  # Zero indexed column
    my $str       = $_[2];
    my $strlen    = $_[3];

    # The length of the first record is the total length of the NOTE.
    # Therefore, it can be greater than 2048.
    #
    if ($strlen > 2048) {
        $length = 0x06 + 2048;
    }
    else{
        $length = 0x06 + $strlen;
    }


    # TODO Update for ExcelXML format

}








###############################################################################
#
# New XML code
#
###############################################################################






###############################################################################
#
# _write_xml_table()
#
# Write the stored data into the Table element.
#
# TODO note about data structure
#
sub _write_xml_table {

    my $self      = shift;

    $self->_write_xml_start_tag(2, 1, 1, 'Table',
                                         'ss:ExpandedColumnCount',
                                          $self->{_dim_colmax},
                                          'ss:ExpandedRowCount',
                                          $self->{_dim_rowmax},
                                       );


    # Iterate over the stored data

    my $row_num = 0;

    for my $row (@{$self->{_table}}) {
        if ($row) {
            $self->_write_xml_row($row_num);

            my $col_num = 0;

            for my $col (@{$row}) {
                if ($col) {
                    $self->_write_xml_cell($row_num, $col_num);
                }

                $col_num++;
            }
            $self->_write_xml_end_tag(3, 1, 0, 'Row');
        }
        $row_num++;
    }

    $self->_write_xml_end_tag(2, 1, 0, 'Table');
}


###############################################################################
#
# _write_xml_row()
#
# Write a Row element start tag.
#
sub _write_xml_row {

    my $self  = shift;

    my $row     = $_[0];
    my @attribs;

    push @attribs, "ss:Index", $row +1 if $row != $self->{prev_row} +1;

    $self->_write_xml_start_tag(3, 1, 0, 'Row', @attribs);

    $self->{prev_row} = $row;
}


###############################################################################
#
# _write_xml_cell()
#
# Write a Cell element start tag.
#
sub _write_xml_cell {

    my $self      = shift;

    my $row         = $_[0];
    my $col         = $_[1];

    my $datatype    = $self->{_table}->[$row]->[$col]->[0];
    my $data        = $self->{_table}->[$row]->[$col]->[1];
    my $format      = $self->{_table}->[$row]->[$col]->[2];

    my @attribs;


    push @attribs, "ss:Index",   $col +1       if $col != $self->{prev_col} +1;
    push @attribs, "ss:StyleID", "s" . $format if $format;



    if ($datatype == $self->{_datatypes}->{Formula}) {
        push @attribs, "ss:Formula", $data;
    }


    # Write the Number data element
    if ($datatype == $self->{_datatypes}->{Number}) {
        $self->_write_xml_start_tag(4, 1, 0, 'Cell', @attribs);
        $self->_write_xml_cell_data('Number', $data);
        $self->_write_xml_end_tag(4, 1, 0, 'Cell');
        $self->{prev_col} = $col;
        return;
    }


    # Write the String data element
    if ($datatype == $self->{_datatypes}->{String}) {
        $self->_write_xml_start_tag(4, 1, 0, 'Cell', @attribs);
        $self->_write_xml_cell_data('String', $data);
        $self->_write_xml_end_tag(4, 1, 0, 'Cell');
        $self->{prev_col} = $col;
        return;
    }


    # Write an empty Data element for a formula data
    if ($datatype == $self->{_datatypes}->{Formula}) {
        $self->_write_xml_element(4, 1, 0, 'Cell', @attribs);
        $self->{prev_col} = $col;
    }


    # Write an empty Data element for a blank cell
    if ($datatype == $self->{_datatypes}->{Blank}) {
        $self->_write_xml_element(4, 1, 0, 'Cell', @attribs);
    $self->{prev_col} = $col;
        return;
    }

}


###############################################################################
#
# _write_xml_cell_data()
#
# Write a Data element start tag.
#
sub _write_xml_cell_data {

    my $self  = shift;

    my $datatype    = $_[0];
    my $data        = $_[1];

    $self->_write_xml_start_tag(5, 0, 0, 'Data', 'ss:Type', $datatype);

    if ($datatype eq 'Number') {$self->_write_xml_unencoded_content($data)}
    else                       {$self->_write_xml_content($data)          }

    $self->_write_xml_end_tag(0, 1, 0, 'Data');
}


###############################################################################
#
# _convert_formula($A1_formula)
#
# Converts a string containing an Excel formula in A1 notation into a string
# containing a formula in R1C1 notation.
#
# Instead of parsing the formula into its component parts, as Spreadsheet::
# WriteExcel::Formula does, we convert the A1 style references to R1C1
# references using regexes. This avoid the significant overhead of the
# Parse::RecDescent parser in S::WE::Formula. The main problem with this
# simplified approach is that there is potential for false matches. Such as
# B5 in the following formula (only the last is a valid match).
#
# "= "B5" & SheetB5!B5"
#
# The method used here is to replace potential false matches before converting
# the real A1 cell references and then substitute back the replaced data.
#
# Returns: a string. A representation of a formula in R1C1 notation.
#
sub _convert_formula {

    my $self = shift;

    my $row     = $_[0];
    my $col     = $_[1];
    my $formula = $_[2];

    my @strings;
    my @sheets;

    # Replace double quoted strings in formula. Strings may contain escaped
    # double quotes. Regex by merlyn.
    # See http://www.perlmonks.org/index.pl?node_id=330280
    #
    push @strings, $1 while $formula =~ s/("([^"]|"")*")/__swe__str__/; # "


    # Replace worksheet references in formula, such as Sheet1! or 'Sheet 1'!
    #
    push @sheets,  $1 while $formula =~ s/(('[^']+'|[\w\.]+)!)/__swe__sht__/;


    # Replace valid A1 cell references with R1C1 references. Cell ranges such
    # as B5::G10 are replaced in two goes.
    #
    $formula =~ s{(?<![A-Z])(\$?[A-I]?[A-Z]\$?\d+)}
                 {$self->_A1_to_R1C1($row, $col, $1)}eg;


    # Replace row ranges such as 2:9 with R1C1 references.
    #
    $formula =~ s{(\$?\d+:\$?\d+)}
                 {$self->_row_range_to_R1C1($row, $1)}eg;


    # Replace column ranges such as A:Z with R1C1 references.
    #
    $formula =~ s{(?<!])(\$?[A-I]?[A-Z]:\$?[A-I]?[A-Z])(?!\[)}
                 {$self->_col_range_to_R1C1($col, $1)}eg;


    # Quoted A1 style alphanumeric sheetnames don't need quoting when
    # converted to R1C1 style. For example "='A1'!A1" becomes "=A1!RC" (without
    # the single quotes since A1 isn't a reserved name in R1C1 notation).
    s/^'([a-zA-Z0-9]+)'!$/$1!/ for @sheets;


    # However, sheet names that looks like R1C1 notation do have to be single
    # quoted. For example "='R4C'!A1"  becomes "='R4C'!RC".
    #
    s/^((R\d*|R\[\d+\])?(C\d*|C\[\d+\])?)!$/\'$1\'!/ for @sheets;


    # Replace temporarily escaped strings. Note that the s///s are performed in
    # reverse order to the substitutions above in case of nested strings.
    $formula =~ s/__swe__sht__/shift @sheets /e while @sheets;
    $formula =~ s/__swe__str__/shift @strings/e while @strings;

    return $formula;
}


###############################################################################
#
# _A1_to_R1C1($A1_string)
#
# Converts a string containing an Excel cell reference in A1 notation into a
# string containing a formula in R1C1 notation. For example:
#
#   '=G1' in cell (0, 0) becomes '=RC[6]'.
#
# The R1C1 value is relative to the row and column from which it is referred.
# With reference to the above example:
#
#   '=G1' in cell (1, 0) becomes '=R[-1]C[6]'.
#
# Returns: a string. A representation of a cell reference in R1C1 notation.
#
#
sub _A1_to_R1C1 {

    my $self =  shift;

    my $current_row = $_[0];
    my $current_col = $_[1];

    my ($row, $col, $row_abs, $col_abs) = $self->_cell_to_rowcol($_[2]);

    # Row part
    my $r1c1 = 'R';

    if ($row_abs) {
        $r1c1 .= $row +1; # 1 based
    }
    else {
        $r1c1 .= '[' . ($row -$current_row) . ']' unless $row == $current_row;
    }

    # Column part
    $r1c1 .= 'C';

    if ($col_abs) {
        $r1c1 .= $col +1; # 1 based
    }
    else {
        $r1c1 .= '[' . ($col -$current_col) . ']' unless $col == $current_col;
    }

    return $r1c1;
}


###############################################################################
#
# _row_range_to_R1C1($string)
#
# Replace row ranges with R1C1 references. For example:
#
#   '=20:120' in cell (7, 0) becomes '=R[12]:R[112]'
#
# Returns: a string. A representation of a row cell reference in R1C1 notation.

#
sub _row_range_to_R1C1 {

    my $self =  shift;

    my $current_row = $_[0] +1; # One based
    my $range       = $_[1];


    # Split the range into 2 rows
    my ($row1, $row2) = split ':', $range;

    for my $row ($row1, $row2) {

        my $row_abs = $row =~ s/\$//;

        # TODO Check row range

        my $r1c1 = 'R';

        if ($row_abs) {
            $r1c1 .= $row;
        }
        else {
            $r1c1 .= '['.($row -$current_row) .']' unless $row == $current_row;
        }

        $row = $r1c1;
    }

    # A single row range such as 'R2:R2' is represented as 'R2'
    if ($row1 eq $row2) {return $row1        }
    else                {return "$row1:$row2"}
}


###############################################################################
#
# _col_range_to_R1C1($string)
#
# Replace column ranges with R1C1 references. For example:
#
#   '=D:Z' in cell (6, 0) becomes '=C[3]:C[25]'
#
# Returns: a string. A representation of a col range reference in R1C1 notation.
#
sub _col_range_to_R1C1 {

    my $self =  shift;

    my $current_col = $_[0] +1; # One based
    my $range       = $_[1];

    my %columns;

    @columns{'A' .. 'IV'} = (1 ..256); # Cheap and cheerful or quick and dirty.


    # Split the range into 2 cols
    my ($col1, $col2) = split ':', $range;

    for my $col ($col1, $col2) {

        my $col_abs = $col =~ s/\$//;

        if (not exists $columns{$col}) {
            warn "$col is not an Excel column label.\n"; # TODO Carp
            return $range;
        }
        else {
            $col = $columns{$col};
        }

        my $r1c1 = 'C';

        if ($col_abs) {
            $r1c1 .= $col;
        }
        else {
            $r1c1 .= '['.($col -$current_col).']' unless $col == $current_col;
        }
        $col = $r1c1;
    }

    # A single column range such as 'C3:C3' is represented as 'C3'
    if ($col1 eq $col2) {return $col1        }
    else                {return "$col1:$col2"}
}


1;


__END__


=head1 NAME

Worksheet - A writer class for Excel Worksheets.

=head1 SYNOPSIS

See the documentation for Spreadsheet::WriteExcelXML

=head1 DESCRIPTION

This module is used in conjunction with Spreadsheet::WriteExcelXML.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 PATENT LICENSE

Software programs that read or write files that comply with the Microsoft specifications for the Office Schemas must include the following notice:

"This product may incorporate intellectual property owned by Microsoft Corporation. The terms and conditions upon which Microsoft is licensing such intellectual property may be found at http://msdn.microsoft.com/library/en-us/odcXMLRef/html/odcXMLRefLegalNotice.asp."

=head1 COPYRIGHT

� MM-MMIV, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
