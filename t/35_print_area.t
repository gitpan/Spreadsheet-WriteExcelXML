#!/usr/bin/perl -w

###############################################################################
#
# A test for Spreadsheet::WriteExcelXML.
#
# Tests the print_area method.
#
# reverse('�'), November 2004, John McNamara, jmcnamara@cpan.org
#


use strict;
use Spreadsheet::WriteExcelXML;
use Test::More tests => 6;


##############################################################################
#
# Create a new Excel XML file with different formats on each page.
#
my $test_file  = "temp_test_file.xml";
my $workbook   = Spreadsheet::WriteExcelXML->new($test_file);

# Test with older cell limits.
$workbook->use_lower_cell_limits();

# We use 'Sheet n' worksheet names so that they are single quoted in the
# Excel XML output.

my $worksheet1 = $workbook->add_worksheet('Sheet 1');
   $worksheet1->print_area("A1:A1");

my $worksheet2 = $workbook->add_worksheet('Sheet 2');
   $worksheet2->print_area("A1:A2");

my $worksheet3 = $workbook->add_worksheet('Sheet 3');
   $worksheet3->print_area("A1:B1");

my $worksheet4 = $workbook->add_worksheet('Sheet 4');
   $worksheet4->print_area("A1:B5");

my $worksheet5 = $workbook->add_worksheet('Sheet 5');
   $worksheet5->print_area("A:A");

my $worksheet6 = $workbook->add_worksheet('Sheet 6');
   $worksheet6->print_area(0,0,0,255);

# Should be ignored
my $worksheet7 = $workbook->add_worksheet('Sheet 7');
   $worksheet7->print_area("A1:IV65536");

# Should be ignored
my $worksheet8 = $workbook->add_worksheet('Sheet 8');
   $worksheet8->print_area("A1");


$workbook->close();


##############################################################################
#
# Re-open and reread the Excel file.
#
open XML, $test_file or die "Couldn't open $test_file: $!\n";
my @swex_data = extract_names(*XML);
close XML;
unlink $test_file;


##############################################################################
#
# Read the data from the Excel file in the __DATA__ section
#
my @test_data = extract_names(*DATA);



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
    is($swex_data[$i], $test_data[$i], "Testing print_area()");

}



##############################################################################
#
# Extract <Name> sub-elements from a given filehandle.
#
sub extract_names {

    my $fh     = $_[0];
    my $in_elem = 0;
    my $element    = '';
    my @elements;

    while (<$fh>) {
        s/^\s+([<| ])/$1/;
        s/\s+$//;

        $in_elem  = 1 if (m[^<Names] .. m[^</Names]);

        $element .= $_ if $in_elem and not m[</?Names];

        if (m[^</Names>]) {
            push @elements, $element;
            $in_elem = 0;
            $element = '';
        }
    }
    return @elements;
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
 </Styles>
 <Worksheet ss:Name="Sheet 1">
  <Names>
   <NamedRange ss:Name="Print_Area" ss:RefersTo="='Sheet 1'!R1C1"/>
  </Names>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>-3</HorizontalResolution>
    <VerticalResolution>0</VerticalResolution>
   </Print>
   <Selected/>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="Sheet 2">
  <Names>
   <NamedRange ss:Name="Print_Area" ss:RefersTo="='Sheet 2'!R1C1:R2C1"/>
  </Names>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>-3</HorizontalResolution>
    <VerticalResolution>0</VerticalResolution>
   </Print>
   <Panes>
    <Pane>
     <Number>3</Number>
     <RangeSelection>R1C1:R2C1</RangeSelection>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="Sheet 3">
  <Names>
   <NamedRange ss:Name="Print_Area" ss:RefersTo="='Sheet 3'!R1C1:R1C2"/>
  </Names>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>-3</HorizontalResolution>
    <VerticalResolution>0</VerticalResolution>
   </Print>
   <Panes>
    <Pane>
     <Number>3</Number>
     <RangeSelection>R1C1:R1C2</RangeSelection>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="Sheet 4">
  <Names>
   <NamedRange ss:Name="Print_Area" ss:RefersTo="='Sheet 4'!R1C1:R5C2"/>
  </Names>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>-3</HorizontalResolution>
    <VerticalResolution>0</VerticalResolution>
   </Print>
   <Panes>
    <Pane>
     <Number>3</Number>
     <RangeSelection>R1C1:R5C2</RangeSelection>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="Sheet 5">
  <Names>
   <NamedRange ss:Name="Print_Area" ss:RefersTo="='Sheet 5'!C1"/>
  </Names>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>-3</HorizontalResolution>
    <VerticalResolution>0</VerticalResolution>
   </Print>
   <Panes>
    <Pane>
     <Number>3</Number>
     <ActiveRow>28</ActiveRow>
     <ActiveCol>5</ActiveCol>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="Sheet 6">
  <Names>
   <NamedRange ss:Name="Print_Area" ss:RefersTo="='Sheet 6'!R1"/>
  </Names>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>-3</HorizontalResolution>
    <VerticalResolution>0</VerticalResolution>
   </Print>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="Sheet 7">
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <Print>
    <ValidPrinterInfo/>
    <PaperSizeIndex>9</PaperSizeIndex>
    <HorizontalResolution>-3</HorizontalResolution>
    <VerticalResolution>0</VerticalResolution>
   </Print>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
 <Worksheet ss:Name="Sheet 8">
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>

