package Spreadsheet::WriteExcelXML::XMLwriter;

###############################################################################
#
# XMLwriter - An abstract base class for Excel workbooks and worksheets.
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








use vars qw($VERSION @ISA);
@ISA = qw(Exporter);

$VERSION = '0.02';

###############################################################################
#
# new()
#
# Constructor
#
sub new {

    my $class  = $_[0];

    my $self   = {
                    _indentation => "    ",
                    _filehandle  => undef,
                    _no_encoding => 0,
                    _printed     => 1,
                 };

    bless  $self, $class;
    return $self;
}


###############################################################################
#
# _format_tag($level, $nl, $list, @attributes)
#
# This function formats an XML element tag for printing. Adds indentation and
# newlines as specified. Keeps attributes, if any, on one line or formats
# them one per line.
#
# Args:
#       $level      = The indentation level (int)
#       $nl         = Number of newlines after tag (int)
#       $list       = List attributes on separate lines (0, 1, 2)
#                       0 = No list
#                       1 = Automatic list
#                       2 = Explicit list
#       @attributes = Attribute/Value pairs
#
# The list option puts the attributes on separate lines if there if there is
# more than one attribute. List option 2 generates this effect even when there
# is only one attribute.
#
sub _format_tag {

    my $self    = shift;

    my $level   = shift;
    my $nl      = shift;
    my $list    = shift;

    my $element = $self->{_indentation} x $level. '<' . shift;

    # Autolist option. Only use list format if there is more than 1 attribute.
    $list = 0 if $list == 1 and @_ <= 2;


    # Special case. If _indentation is "" avoid all unnecessary whitespace
    $list = 0 if $self->{_indentation} eq "";
    $nl   = 0 if $self->{_indentation} eq "";


    while (@_) {
        my $attrib = shift;
        my $value  = $self->_encode_xml_escapes(shift);

        if ($list) {$element .= "\n" . $self->{_indentation} x ($level +1);}
        else       {$element .= ' ';                                       }

        $element .= $attrib;
        $element .= '="' . $value . '"';
    }

    $nl = $nl ? "\n" x $nl : "";

    return $element . '>'. $nl;
}


###############################################################################
#
# _encode_xml_escapes()
#
# Encode standard XML escapes, namely " & < >. The apostrophe character isn't
# escaped since it will only occur in double quoted strings.
#
sub _encode_xml_escapes {

    my $self  = shift;
    my $value = $_[0];

    # Print un-encoded entities for debugging
    return $value if $self->{_no_encoding};

    for ($value) {
        s/&/&amp;/g;
        s/</&lt;/g;
        s/>/&gt;/g;
        s/"/&quot;/g; # "
        #s/'/&pos;/g; # Not used
        s/\n/&#10;/g;
    }

    return $value;
}


###############################################################################
#
# _write_xml_start_tag()
#
# Creates a formatted XML opening tag. Prints to the current filehandle by
# default.
#
# Ex: <Worksheet ss:Name="Sheet1">
#
sub _write_xml_start_tag {

    my $self = shift;

    my $tag  = $self->_format_tag(@_);

    local $\; # Make print() ignore -l on the command line.
    print {$self->{_filehandle}} $tag if $self->{_printed};

    return $tag;
}


###############################################################################
#
# _write_xml_directive()
#
# Creates a formatted XML <? ?> directive. Prints to the current filehandle by
# default.
#
# Ex: <?xml version="1.0"?>
#
sub _write_xml_directive {

    my $self = shift;

    my $tag  =  $self->_format_tag(@_);
       $tag  =~ s[<][<?];
       $tag  =~ s[>][?>];

    local $\; # Make print() ignore -l on the command line.
    print {$self->{_filehandle}} $tag if $self->{_printed};

    return $tag;
}


###############################################################################
#
# _write_xml_end_tag()
#
# Creates the closing tag of an XML element. Prints to the current filehandle
# by default.
#
# Ex: </Worksheet>
#
sub _write_xml_end_tag {

    my $self = shift;

    my $tag  =  $self->_format_tag(@_);
       $tag  =~ s[<][</];

    local $\; # Make print() ignore -l on the command line.
    print {$self->{_filehandle}} $tag if $self->{_printed};

    return $tag;

}


###############################################################################
#
# _write_xml_element()
#
# Creates a single open and closed XML element. Prints to the current
# filehandle by default.
#
# Ex: <Alignment ss:Vertical="Bottom"/> or XML <Alignment/>
#
sub _write_xml_element {

    my $self = shift;

    my $tag  =  $self->_format_tag(@_);
       $tag  =~ s[>][/>];

    local $\; # Make print() ignore -l on the command line.
    print {$self->{_filehandle}} $tag if $self->{_printed};

    return $tag;
}


###############################################################################
#
# _write_xml_content()
#
# Creates an encoded XML element content. Prints to the current filehandle
# by default.
#
# Ex: Hello in <Data ss:Type="String">Hello</Data>
#
sub _write_xml_content {

    my $self = shift;

    my $tag  = $self->_encode_xml_escapes($_[0]);

    local $\; # Make print() ignore -l on the command line.
    print {$self->{_filehandle}} $tag if $self->{_printed};

    return $tag;

}


###############################################################################
#
# _write_xml_unencoded_content()
#
# Creates an un-encoded XML element content. Prints to the current filehandle
# by default. Used for numerical or other data that doesn't need to be
# encoded.
#
# Ex: 1.2345 in <Data ss:Type="Number">1.2345</Data>
#
sub _write_xml_unencoded_content {

    my $self = shift;

    my $tag  = $_[0];

    local $\; # Make print() ignore -l on the command line.
    print {$self->{_filehandle}} $tag if $self->{_printed};

    return $tag;
}


###############################################################################
#
# _set_printed()
#
# Turn the option to print on or off. By default this option is 1 = on.
# It is mainly only turned off for testing pupropes.
#
sub _set_printed {

    my $self = shift;
       $self->{_printed} = $_[0];
}


###############################################################################
#
# set_indentation()
#
# Set indentation string used to indent the output. The default is 4 spaces.
#
sub set_indentation {

    my $self = shift;
       $self->{_indentation} = defined $_[0] ? $_[0] : '    ';
}


1;


__END__


=head1 NAME

XMLwriter - An abstract base class for Excel workbooks and worksheets.

=head1 SYNOPSIS

See the documentation for Spreadsheet::WriteExcelXML

=head1 DESCRIPTION

This module is used in conjunction with Spreadsheet::WriteExcelXML.

=head1 AUTHOR

John McNamara jmcnamara@cpan.org

=head1 COPYRIGHT

© MM-MMIV, John McNamara.

All Rights Reserved. This module is free software. It may be used, redistributed and/or modified under the same terms as Perl itself.
