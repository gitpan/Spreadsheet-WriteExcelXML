#!/usr/bin/perl -w

##############################################################################
#
# A simple formatting example using Spreadsheet::WriteExcelXML.
#
# This program demonstrates the diagonal border cell format.
#
# reverse('�'), May 2004, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcelXML;


my $workbook  = Spreadsheet::WriteExcelXML->new('diag_border.xls');
my $worksheet = $workbook->add_worksheet();


my $format1   = $workbook->add_format(diag_type       => '1');

my $format2   = $workbook->add_format(diag_type       => '2');

my $format3   = $workbook->add_format(diag_type       => '3');

my $format4   = $workbook->add_format(
                                      diag_type       => '3',
                                      diag_border     => '7',
                                      diag_color      => 'red',
                                     );


$worksheet->write('B3',  'Text', $format1);
$worksheet->write('B6',  'Text', $format2);
$worksheet->write('B9',  'Text', $format3);
$worksheet->write('B12', 'Text', $format4);



__END__

