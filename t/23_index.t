#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcelXML.
#
# Tests the implicit/explicit ss:Index attribute of <Row> and <Cell> elements.
#
# reverse('�'), July 2004, John McNamara, jmcnamara@cpan.org
#


use strict;
use Spreadsheet::WriteExcelXML;
use Test::More tests => 19;



##############################################################################
#
# Create a new Excel XML file with row data set.
#
my $test_file = "temp_test_file.xml";
my $workbook  = Spreadsheet::WriteExcelXML->new($test_file);
my $worksheet = $workbook->add_worksheet();

# Test for false "$col == $self->{prev_col} +1" from _write_xml_cell().
# Bug in versions < 0.05.
#
$worksheet->write('A1',  'A1');
$worksheet->write('B2',  'B2');
$worksheet->write('C3',  'C3');

# Test for consecutive columns.
$worksheet->write('A5',  'A5');
$worksheet->write('B5',  'B5');
$worksheet->write('C5',  'C5');

# Test for non-consecutive columns.
$worksheet->write('A7',  'A7');
$worksheet->write('C7',  'C7');
$worksheet->write('D7',  'D7');

# Test for consecutive rows.
$worksheet->write('A9',  'A9' );
$worksheet->write('A10', 'A10');
$worksheet->write('A11', 'A11');

# Test for non-consecutive rows.
$worksheet->write('A13', 'A13');
$worksheet->write('A15', 'A15');
$worksheet->write('A17', 'A17');

# Test for invalid cells. Module should ignore these.
$worksheet->write('IW1',      'IW1'     ); # > Col limit.
$worksheet->write('A65537',   'A65537'  ); # > Row limit.
$worksheet->write('IW165537', 'IW165537'); # > Row and col limit.

# Test for valid/invalid cells.
my $err1 = $worksheet->write('IV1',      'IV1'     ); # Col limit.
my $err2 = $worksheet->write('A65536',   'A65536'  ); # Row limit.
my $err3 = $worksheet->write('IV65536',  'IV65536' ); # Row and col limit.
my $err4 = $worksheet->write('IW1',      'IW1'     ); # > Col limit.
my $err5 = $worksheet->write('A65537',   'A65537'  ); # > Row limit.
my $err6 = $worksheet->write('IW165537', 'IW165537'); # > Row and col limit.

$workbook->close();


##############################################################################
#
# Re-open and reread the Excel file.
#
open XML, $test_file or die "Couldn't open $test_file: $!\n";
my @swex_data = extract_rows(*XML);
close XML;
unlink $test_file;


##############################################################################
#
# Read the data from the Excel file in the __DATA__ section
#
my @test_data = extract_rows(*DATA);


##############################################################################
#
# Check for the same number of elements.
#

is(@swex_data, @test_data, " \tCheck for data size");


##############################################################################
#
# Test that the SWEX elements and Excel are the same.
#

# Pad the SWEX data if necessary.
push @swex_data, ('') x (@test_data -@swex_data);

for my $i (0 .. @test_data -1) {
    is($swex_data[$i],$test_data[$i], " \tTesting ss:Index attribute");

}


##############################################################################
#
# Test the "row or column out of range" return values.
#
is($err1,  0, " \tChecking   valid row/column");
is($err2,  0, " \tChecking   valid row/column");
is($err3,  0, " \tChecking   valid row/column");
is($err4, -2, " \tChecking invalid row/column");
is($err5, -2, " \tChecking invalid row/column");
is($err6, -2, " \tChecking invalid row/column");


##############################################################################
#
# Extract <Row> elements from a given filehandle.
#
sub extract_rows {

    my $fh     = $_[0];
    my $in_row = 0;
    my $row    = '';
    my @rows;

    while (<$fh>) {
        s/^\s+//;
        s/\s+$//;

        $in_row = 1 if m/<Row/;

        $row .= $_ if $in_row;

        if (m[</Row>]) {
            $in_row  = 0;
            push @rows, $row;
            $row     = '';
        }
    }

    return @rows;
}



# The following file was created by Excel. Some redundant data is removed.
__DATA__
<?xml version="1.0"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Sheet1">
  <Table ss:ExpandedColumnCount="256" ss:ExpandedRowCount="65536"
   x:FullColumns="1" x:FullRows="1">
   <Row>
    <Cell><Data ss:Type="String">A1</Data></Cell>
    <Cell ss:Index="256"><Data ss:Type="String">IV1</Data></Cell>
   </Row>
   <Row>
    <Cell ss:Index="2"><Data ss:Type="String">B2</Data></Cell>
   </Row>
   <Row>
    <Cell ss:Index="3"><Data ss:Type="String">C3</Data></Cell>
   </Row>
   <Row ss:Index="5">
    <Cell><Data ss:Type="String">A5</Data></Cell>
    <Cell><Data ss:Type="String">B5</Data></Cell>
    <Cell><Data ss:Type="String">C5</Data></Cell>
   </Row>
   <Row ss:Index="7">
    <Cell><Data ss:Type="String">A7</Data></Cell>
    <Cell ss:Index="3"><Data ss:Type="String">C7</Data></Cell>
    <Cell><Data ss:Type="String">D7</Data></Cell>
   </Row>
   <Row ss:Index="9">
    <Cell><Data ss:Type="String">A9</Data></Cell>
   </Row>
   <Row>
    <Cell><Data ss:Type="String">A10</Data></Cell>
   </Row>
   <Row>
    <Cell><Data ss:Type="String">A11</Data></Cell>
   </Row>
   <Row ss:Index="13">
    <Cell><Data ss:Type="String">A13</Data></Cell>
   </Row>
   <Row ss:Index="15">
    <Cell><Data ss:Type="String">A15</Data></Cell>
   </Row>
   <Row ss:Index="17">
    <Cell><Data ss:Type="String">A17</Data></Cell>
   </Row>
   <Row ss:Index="65536">
    <Cell><Data ss:Type="String">A65536</Data></Cell>
    <Cell ss:Index="256"><Data ss:Type="String">IV65536</Data></Cell>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Selected/>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
