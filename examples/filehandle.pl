#!/usr/bin/perl -w

###############################################################################
#
# Example of using Spreadsheet::WriteExcelXML to write to alternative filehandles.
#
# reverse('©'), April 2003, John McNamara, jmcnamara@cpan.org
#

use strict;
use Spreadsheet::WriteExcelXML;
use IO::Scalar;




###############################################################################
#
# Example 1. This demonstrates the standard way of creating an Excel file by
# specifying a file name.
#

my $workbook1  = Spreadsheet::WriteExcelXML->new('fh_01.xml');

die "Couldn't create new Excel file: $!.\n" unless defined $workbook1;

my $worksheet1 = $workbook1->add_worksheet();

$worksheet1->write(0, 0,  "Hi Excel!");




###############################################################################
#
# Example 2. Write an Excel file to an existing filehandle.
#

open TEST, "> fh_02.xml" or die "Couldn't create new Excel file: $!.\n";

binmode TEST; # Always do this regardless of whether the platform requires it.

my $workbook2  = Spreadsheet::WriteExcelXML->new(\*TEST);
my $worksheet2 = $workbook2->add_worksheet();

$worksheet2->write(0, 0,  "Hi Excel!");




###############################################################################
#
# Example 3. Write an Excel file to an existing OO style filehandle.
#

my $fh = FileHandle->new("> fh_03.xml")
         or die "Couldn't create new Excel file: $!.\n";

binmode($fh);

my $workbook3  = Spreadsheet::WriteExcelXML->new($fh);
my $worksheet3 = $workbook3->add_worksheet();

$worksheet3->write(0, 0,  "Hi Excel!");




###############################################################################
#
# Example 4. Write an Excel file to a string via IO::Scalar. Please refer to
# the IO::Scalar documentation for further details.
#

my $xml_str;

tie *xml, 'IO::Scalar', \$xml_str;

my $workbook4  = Spreadsheet::WriteExcelXML->new(\*xml);
my $worksheet4 = $workbook4->add_worksheet();

$worksheet4->write(0, 0, "Hi Excel 4");
$workbook4->close(); # This is required before we use the scalar


# The Excel file is now in $xml_str. As a demonstration, print it to a file.
open    TMP, "> fh_04.xml";
binmode TMP;
print   TMP  $xml_str;
close   TMP;




###############################################################################
#
# Example 5. Write an Excel file to a string via IO::Scalar newer interface.
# Please refer to the IO::Scalar documentation for further details.
#
my $xml_str2;

my $fh5 = IO::Scalar->new(\$xml_str2);


my $workbook5  = Spreadsheet::WriteExcelXML->new($fh5);
my $worksheet5 = $workbook5->add_worksheet();

$worksheet5->write(0, 0, "Hi Excel 5");
$workbook5->close();

# The Excel file is now in $xml_str. As a demonstration, print it to a file.
open    TMP, "> fh_05.xml" or die "Couldn't create new Excel file: $!.\n";


binmode TMP;
print   TMP  $xml_str2; # This is required before we use the scalar
close   TMP;


