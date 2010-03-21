#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcelXML.
#
# Test autofilters.
#
# reverse('�'), April 2005, John McNamara, jmcnamara@cpan.org
#


use strict;
use Spreadsheet::WriteExcelXML;
use Test::More tests => 13;


##############################################################################
#
# Create a new Excel XML file with different formats on each page.
#
my $test_file  = "temp_test_file.xml";
my $workbook   = Spreadsheet::WriteExcelXML->new($test_file);
my $worksheet  = $workbook->add_worksheet();
my $bold       = $workbook->add_format(bold => 1);

$worksheet->set_column('A:D', 12);
$worksheet->set_row(0, 20, $bold);

my @data =  [
                ['Region',   'Item',     'Volume',   'Month',    ],
                ['East',     'Apple',    '9000',     'July',     ],
                ['East',     'Apple',    '5000',     'July',     ],
                ['South',    'Orange',   '9000',     'September',],
                ['North',    'Apple',    '2000',     'November', ],
                ['West',     'Apple',    '9000',     'November', ],
                ['East',     'Pear',     '7000',     'October',  ],
                ['North',    'Pear',     '9000',     'August',   ],
                ['West',     'Orange',   '1000',     'December', ],
                ['West',     'Grape',    '1000',     'November', ],
                ['South',    'Pear',     '10000',    'April',    ],
            ];


$worksheet->write('A1', \@data);


$worksheet->autofilter('A1:D11');


$worksheet->filter_column('A', 'x eq East');
$worksheet->filter_column('C', 'x > 1000 and x < 9000');


$workbook->close();


##############################################################################
#
# Re-open and reread the Excel file.
#
open 'XML', $test_file or die "Couldn't open $test_file: $!\n";
my @swex_data = extract_data(*XML);
close XML;
unlink $test_file;


##############################################################################
#
# Read the data from the Excel file in the __DATA__ section
#
my @test_data = extract_data(*DATA);


##############################################################################
#
# Pad the SWEX and test data if necessary.
#

push @swex_data, ('') x (@test_data -@swex_data);
push @test_data, ('') x (@swex_data -@test_data);


##############################################################################
#
# Run the tests
#
for my $i (0 .. @test_data -1) {
    is($swex_data[$i], $test_data[$i], " \t" . $test_data[$i]);

}


##############################################################################
#
# Extract autofilter elements from a given filehandle.
#
sub extract_data {

    my $fh     = $_[0];
    my $in_opt = 0;
    my $setup    = '';
    my @options;

    while (<$fh>) {
        s/^\s+([<| ])/$1/;
        s/\s+$//;

        push @options, $_ if /^<NamedRange/;
        push @options, $_ if /AutoFilter/;
        push @options, $_ if /FilterOn/;
    }

    return @options;
}


# The following data was generated by Excel.
__DATA__
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
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
  <Style ss:ID="s21">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:Bold="1"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Sheet1">
  <Names>
   <NamedRange ss:Name="_FilterDatabase" ss:RefersTo="=Sheet1!R1C1:R11C4" ss:Hidden="1"/>
  </Names>
  <Table ss:ExpandedColumnCount="4" ss:ExpandedRowCount="11" x:FullColumns="1"
   x:FullRows="1">
   <Column ss:AutoFitWidth="0" ss:Width="66.75" ss:Span="3"/>
   <Row ss:AutoFitHeight="0" ss:Height="19.5" ss:StyleID="s21">
    <Cell><Data ss:Type="String">Region</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Item</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Volume</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Month</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
   </Row>
   <Row ss:Hidden="1">
    <Cell><Data ss:Type="String">East</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Apple</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="Number">9000</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">July</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
   </Row>
   <Row>
    <Cell><Data ss:Type="String">East</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Apple</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="Number">5000</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">July</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
   </Row>
   <Row ss:Hidden="1">
    <Cell><Data ss:Type="String">South</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Orange</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="Number">9000</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">September</Data><NamedCell
      ss:Name="_FilterDatabase"/></Cell>
   </Row>
   <Row ss:Hidden="1">
    <Cell><Data ss:Type="String">North</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Apple</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="Number">2000</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">November</Data><NamedCell
      ss:Name="_FilterDatabase"/></Cell>
   </Row>
   <Row ss:Hidden="1">
    <Cell><Data ss:Type="String">West</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Apple</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="Number">9000</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">November</Data><NamedCell
      ss:Name="_FilterDatabase"/></Cell>
   </Row>
   <Row>
    <Cell><Data ss:Type="String">East</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Pear</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="Number">7000</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">October</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
   </Row>
   <Row ss:Hidden="1">
    <Cell><Data ss:Type="String">North</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Pear</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="Number">9000</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">August</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
   </Row>
   <Row ss:Hidden="1">
    <Cell><Data ss:Type="String">West</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Orange</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="Number">1000</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">December</Data><NamedCell
      ss:Name="_FilterDatabase"/></Cell>
   </Row>
   <Row ss:Hidden="1">
    <Cell><Data ss:Type="String">West</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Grape</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="Number">1000</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">November</Data><NamedCell
      ss:Name="_FilterDatabase"/></Cell>
   </Row>
   <Row ss:Hidden="1">
    <Cell><Data ss:Type="String">South</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">Pear</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="Number">10000</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
    <Cell><Data ss:Type="String">April</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>
   </Row>
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Selected/>
   <FilterOn/>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
  <AutoFilter x:Range="R1C1:R11C4" xmlns="urn:schemas-microsoft-com:office:excel">
   <AutoFilterColumn x:Type="Custom">
    <AutoFilterCondition x:Operator="Equals" x:Value="East"/>
   </AutoFilterColumn>
   <AutoFilterColumn x:Index="3" x:Type="Custom">
    <AutoFilterAnd>
     <AutoFilterCondition x:Operator="GreaterThan" x:Value="1000"/>
     <AutoFilterCondition x:Operator="LessThan" x:Value="9000"/>
    </AutoFilterAnd>
   </AutoFilterColumn>
  </AutoFilter>
 </Worksheet>
</Workbook>