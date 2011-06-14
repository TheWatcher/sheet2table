## @file
#
# @author  Chris Page &lt;chris@starforge.co.uk&gt;
# @version 1.0
# @date    12 June 2011
# @copy    2011, Chris Page &lt;chris@starforge.co.uk&gt;
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 2 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

## @class Spreadsheet::ODSBook
#
package Spreadsheet::ODSBook;

use Archive::Zip;
use XML::Simple;
use Clone qw(clone);
use Spreadsheet::ODSSheet;
use Spreadsheet::ODSCell;
use strict;

## @cmethod $ new($filename)
# Create a new ODSBook object from the content.xml inside the specified .ods file.
# This will attempt to load and parse the file as an .ods file, stripping out lots
# of unnecessary data along the way. This *DOES NOT* retain any kind of formatting
# data - it's intended purely to allow reading of cell values, nothing more.
#
# @param filename The name of the .ods file to load.
# @return A new Spreadsheet::ODSBook object on success. This will die with an error
#         message if any problems are encountered.
sub new {
    my $invocant = shift;
    my $class    = ref($invocant) || $invocant;
    my $filename = shift;

    # Object constructors don't get much more minimal than this...
    my $self = {
        "filename" => $filename,
    };

    my $obj = bless $self, $class;

    # First we need to pull the content.xml out of the .ods file
    my $ods = eval { Archive::Zip -> new($filename) };
    die "Unable to open '$filename' for reading: $@\n" if(!$ods);

    my $content = $ods -> contents('content.xml');
    die "Unable to obtain content.xml from '$filename'.\n" if(!$content);

    # Nuke everything that isn't inside the body, we don't care about anything else.
    $content =~ s|^.*?<office\:body>||gm;
    $content =~ s|</office\:body>.*$||gm;

    # Remove column descriptions, styling, and other junk as they're pointless to us too.
    $content =~ s|<table\:table-column.*?/>||g;
    $content =~ s|table\:style-name=".*?"||g;
    $content =~ s|table\:print=".*?"||g;
    $content =~ s|office\:value-type=".*?"||g;                  # We don't care about types
    $content =~ s{office\:((string|boolean)\-)?value=".*?"}{}g; # This should be duplicated inside the cell body
    $content =~ s|table\:formula=".*?"||g;                      # We can't do anything with formulae anyway

    # Convert covered table cells to empty cells, so they parse correctly
    $content =~ s|table:covered-table-cell|table:table-cell merged="1"|g;

    # We can also just pull the text out into the content, 'cos this is silly.
    $content =~ s|<text:p>(.*?)</text:p>|$1|g;

    # Now squirt the content.xml into XML::Simple.
    $self -> {"doc"} = eval { XMLin($content, ForceContent => 1, KeepRoot => 1, ForceArray => ['table:table', 'table:table-row', 'table:table-cell'], KeyAttr => {'table:table' => 'table:name'}) };
    die "Unable to parse contents of '$filename': $@\n" if($@);

    # make a more useful reference if we can
    if($self -> {"doc"} -> {"office:spreadsheet"} && $self -> {"doc"} -> {"office:spreadsheet"} -> {"table:table"}) {
        $self -> {"doc"} = $self -> {"doc"} -> {"office:spreadsheet"} -> {"table:table"};
    } else {
        die "Document does not appear to be a spreadsheet!\n";
    }

    # Normalise the tables
    $obj -> normalise();

    return $obj;
}


## @method void normalise()
# Process the tables loaded from the .ods, expanding repeat cells and doing general cleanup.
# This will also precalculate things like row and column counts, and identify empty sheets.
sub normalise {
    my $self = shift;

    foreach my $sheet (keys(%{$self -> {"doc"}})) {
        # convert sheet to an ODSSheet...
        $self -> {"doc"} -> {$sheet} = Spreadsheet::ODSSheet -> new($self -> {"doc"} -> {$sheet}, $sheet);

        # While we're here, mark and skip empty sheets
        if(scalar(@{$self -> {"doc"} -> {$sheet} -> {"table:table-row"}}) == 1 &&
           scalar(@{$self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [0] -> {'table:table-cell'}}) == 1 &&
           scalar(keys(%{$self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [0] -> {'table:table-cell'} -> [0]})) == 0) {
            $self -> {"doc"} -> {$sheet} -> {"empty"} = 1;
            next;
        }

        # Table is not empty, look for repeats...
        for(my $row = 0; $row < scalar(@{$self -> {"doc"} -> {$sheet} -> {"table:table-row"}}); ++$row) {
            for(my $col = 0; $col < scalar(@{$self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [$row] -> {"table:table-cell"}}); ++$col) {
                next if(!defined($self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [$row] -> {"table:table-cell"} -> [$col]));

                # Convert the cell to a ODSCell object
                my $cell = $self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [$row] -> {"table:table-cell"} -> [$col] =
                    Spreadsheet::ODSCell -> new($self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [$row] -> {"table:table-cell"} -> [$col]);

                # how many copies do we need?
                my $copies = $cell -> {"table:number-columns-repeated"};

                # If the current cell has a repeat, process it as long as this isn't an empty last column (if it is, we
                # risk having massive amounts of empty columns hanging off the end for no reason)
                if($copies && $copies > 1) {
                    # remove the repeat
                    delete $cell -> {"table:number-columns-repeated"};

                    # Precalculate these to make life easier
                    my $islast  = ($col == (scalar(@{$self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [$row] -> {"table:table-cell"}}) - 1));
                    my $isempty = !defined($self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [$row] -> {"table:table-cell"} -> [$col] -> {"content"});

                    # If this isn't an empty last column, make the copies
                    if(!($islast && $isempty)) {
                        # create an array of copies of the cell, note that we need to use clone
                        # or we would just end up with a bunch of references to the same cell...
                        # (in theory, this is not a problem, but it might
                        my @copylist;
                        for(my $i = 0; $i < $copies; ++$i) {
                            push(@copylist, clone($cell));
                        }

                        # And splice the data in
                        splice(@{$self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [$row] -> {"table:table-cell"}}, $col, 1, @copylist);

                    # If this is an empty last column, just remove it
                    } elsif($islast && $isempty) {
                        delete $self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [$row] -> {"table:table-cell"} -> [$col];
                    }
                }

                # mark merged cells, and add their row spans to the merge areas list
                if($cell -> {"table:number-columns-spanned"} > 1 || $cell -> {"table:number-rows-spanned"} > 1) {
                    $cell -> {"merged"} = 1;

                    my $ecol = $col + (($cell -> {"table:number-columns-spanned"} ? $cell -> {"table:number-columns-spanned"} : 1) - 1);
                    my $erow = $row + (($cell -> {"table:number-rows-spanned"} ? $cell -> {"table:number-rows-spanned"} : 1) - 1);

                    push(@{$self -> {"doc"} -> {$sheet} -> {"merges"}}, [$row, $col, $erow, $ecol]);
                }
            }

            $self -> {"doc"} -> {$sheet} -> {"cols"} = scalar(@{$self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [$row] -> {"table:table-cell"}})
                if(scalar(@{$self -> {"doc"} -> {$sheet} -> {"table:table-row"} -> [$row] -> {"table:table-cell"}}) > $self -> {"doc"} -> {$sheet} -> {"cols"});
        }

        # NOTE: May need to add code to handle number-rows-repeated here.

        $self -> {"doc"} -> {$sheet} -> {"rows"} = scalar(@{$self -> {"doc"} -> {$sheet} -> {"table:table-row"}})
            if(scalar(@{$self -> {"doc"} -> {$sheet} -> {"table:table-row"}}) > $self -> {"doc"} -> {$sheet} -> {"rows"});
    }
}


## @method $ worksheet($sheet)
# Get the Spreadsheet::ODSSheet object for the requested sheet. The argument may either
# be a sheet name, or sheet number. Sheets are ordered alphanumerically by name.
#
# @param sheet The name of the sheet to return, or the position in the sorted sheet list.
# @return The Spreadsheet::ODSSheet object for the selected sheet, or undef if the sheet
#         does not exist.
sub worksheet {
    my $self  = shift;
    my $sheet = shift;

    # 'sheet' may be a string name, or an integer. Working out which may be fun, as a user
    # can call a sheet '0' too, so try it as a name first, fall back on an index if that fails.
    if($self -> {"doc"} -> {$sheet}) {
        return $self -> {"doc"} -> {$sheet};
    }

    # If sheet only contains digits, try using it as an array index.
    if($sheet =~ /^\d+$/) {
        my @sheets = sort keys(%{$self -> {"doc"}});
        return $self -> {"doc"} -> {$sheets[$sheet]};
    }

    return undef;
}


## @method @ worksheets()
# Obtain an array of worksheets in the workbook. This returns an array of
# Spreasheet::ODSSheet objects, sorted by sheet name.
#
# @return An array of spreadsheet objects.
sub worksheets {
    my $self  = shift;

    my @sheets;
    foreach my $name (sort(keys(%{$self -> {"doc"}}))) {
        my $sheet = $self -> worksheet($name);
        push(@sheets, $sheet);
    }

    return @sheets;
}


## @method $ worksheet_count()
# Obtain the number of worksheets in the workbook.
#
# @return The number of worksheets in the workbook.
sub worksheet_count {
    my $self = shift;

    return scalar(keys(%{$self -> {"doc"}}));
}


## @method $ get_filename()
# Returns the name of the file this Spreadsheet::ODSBook was created from.
#
# @return The filename the ODSBook was created from.
sub get_filename {
    my $self = shift;

    return $self -> {"filename"};
}

1;
