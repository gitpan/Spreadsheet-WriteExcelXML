#!/usr/bin/perl -w

##############################################################################
#
# A simple example of converting some Unicode text to an Excel file using
# Spreadsheet::WriteExcelXML and perl 5.8.
#
# This example generates some Chinese from a file with BIG5 encoded text.
#
#
# reverse('�'), September 2004, John McNamara, jmcnamara@cpan.org
#



# Perl 5.8 or later is required for proper utf8 handling. For older perl
# versions you should use UTF16 and the write_unicode() method.
# See the write_unicode section of the Spreadsheet::WriteExcelXML docs.
#
require 5.008;

use strict;
use Spreadsheet::WriteExcelXML;


my $workbook  = Spreadsheet::WriteExcelXML->new("unicode_big5.xls");
my $worksheet = $workbook->add_worksheet();
   $worksheet->set_column('A:A', 80);


my $file = 'unicode_big5.txt';

open FH, '<:encoding(big5)', $file  or die "Couldn't open $file: $!\n";

my $row = 0;

while (<FH>) {
    next if /^#/; # Ignore the comments in the sample file.
    chomp;
    $worksheet->write($row++, 0,  $_);
}


__END__

