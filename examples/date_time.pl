#!/usr/bin/perl -w

###############################################################################
#
# An example program showing how to use the write_date_time() worksheet method
# in Spreadsheet::WriteExcelXML.
#
# reverse('©'), March 2004, John McNamara, jmcnamara@cpan.org
#



use strict;
use Spreadsheet::WriteExcelXML;

my $workbook   = Spreadsheet::WriteExcelXML->new("datetime.xml");

# Always check that the file was created.
die "Couldn't create new Excel file: $!.\n" unless defined $workbook;

# Add a worksheet and a simple format.
my $worksheet = $workbook->add_worksheet();
my $bold      = $workbook->add_format(bold =>  1);

$worksheet->set_column('A:B', 30);


# Some example date and time formats
my $format1   = $workbook->add_format(num_format => 'General Date');
my $format2   = $workbook->add_format(num_format => 'Short Date'  );
my $format3   = $workbook->add_format(num_format => 'Medium Date' );
my $format4   = $workbook->add_format(num_format => 'Long Date'   );
my $format5   = $workbook->add_format(num_format => 'Short Time'  );
my $format6   = $workbook->add_format(num_format => 'Medium Time' );
my $format7   = $workbook->add_format(num_format => 'Long Time'   );
my $format8   = $workbook->add_format(num_format => 'mm/dd/yy'    );
my $format9   = $workbook->add_format(num_format => 'dd/mm/yy'    );


# Write some explanatory labels
$worksheet->write_date_time('A1', '"General Date" format', $bold);
$worksheet->write_date_time('A2', '"Short Date" format',   $bold);
$worksheet->write_date_time('A3', '"Medium Date" format',  $bold);
$worksheet->write_date_time('A4', '"Long Date" format',    $bold);
$worksheet->write_date_time('A5', '"Short Time" format',   $bold);
$worksheet->write_date_time('A6', '"Medium Time" format',  $bold);
$worksheet->write_date_time('A7', '"Long Time" format',    $bold);
$worksheet->write_date_time('A8', '"mm/dd/yy" format',     $bold);
$worksheet->write_date_time('A9', '"dd/mm/yy" format',     $bold);


# Write the same date with different formatting
$worksheet->write_date_time('B1', '2004-05-11T23:20', $format1);
$worksheet->write_date_time('B2', '2004-05-11T23:20', $format2);
$worksheet->write_date_time('B3', '2004-05-11T23:20', $format3);
$worksheet->write_date_time('B4', '2004-05-11T23:20', $format4);
$worksheet->write_date_time('B5', '2004-05-11T23:20', $format5);
$worksheet->write_date_time('B6', '2004-05-11T23:20', $format6);
$worksheet->write_date_time('B7', '2004-05-11T23:20', $format7);
$worksheet->write_date_time('B8', '2004-05-11T23:20', $format8);
$worksheet->write_date_time('B9', '2004-05-11T23:20', $format9);


__END__

