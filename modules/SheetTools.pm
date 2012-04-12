## @file
# This file contains the implementation of the SheetTools class. This
# class provides a number of useful utility fonctions for dealing with
# workbooks and worksheets.
#
# @author  Chris Page &lt;chris@starforge.co.uk&gt;
# @version 1.2
# @date    15 Jun 2011
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

## @class SheetTools
# This class provides a number of functions to simplify the loading
# and processing of excel (xls and xlsx) files and opendocument
# spreadsheet (ods) files. Note that, despite supporting the loading of
# ods files, this class relies heavily on the terminology used for
# excel, in part because it was originally written for Excel files, and
# in part because the terminology is usefully distinct.
package SheetTools;

use strict;
use Spreadsheet::ParseExcel;
use Spreadsheet::XLSX;
use Spreadsheet::ODSBook;

use Utils qw(path_join);

## @cmethod $ new(%args)
# Create a new SheetTools object.
#
sub new {
    my $invocant = shift;
    my $class    = ref($invocant) || $invocant;

    # Object constructors don't get much more minimal than this...
    my $self = {  @_,
    };

    die "No settings object provided to ExcelTools\n" if(!$self -> {"settings"});
    die "No database object provided to ExcelTools\n" if(!$self -> {"dbh"});
    die "No template object provided to ExcelTools\n" if(!$self -> {"template"});
    die "No CGI object provided to ExcelTools\n"      if(!$self -> {"cgi"});

    return bless $self, $class;
}


## @fn $ load_workbook($local_name, $type)
# Load the specified workbook from the filesystem into memory, and return a reference
# to it.
#
# @param local_name The name of the file, without any path.
# @param type       The file type, must be 'xls' or 'xlsx'
# @return A reference to the workbook, or an error message string.
sub load_workbook {
    my $self       = shift;
    my $local_name = shift;
    my $type       = shift;
    my $filename   = path_join($self -> {"settings"} -> {"config"} -> {"file_dir"}, $local_name);

    # If the file is an xls file, we can use the nice ParseExcel module...
    if($type eq "xls") {
        my $parser = Spreadsheet::ParseExcel -> new();
        my $workbook = $parser -> parse($filename);

        return $workbook or $parser -> error();

    # If it's an xlsx file, we need to use the somewhat nastier XLSX module. This will
    # die if left to its own devices, so we need to help it out a bit...
    } elsif($type eq "xlsx") {
        my $workbook = eval { Spreadsheet::XLSX -> new($filename) };

        # if eval error is set, the parse failed, so return it as a sane string
        return $@ if($@);

        # otherwise, return the workbook
        return $workbook;

    # If it's a .ods file, we need to fall back on our ODSBook hack module
    } elsif($type eq "ods") {
        my $workbook = eval { Spreadsheet::ODSBook -> new($filename) };

        # if eval error is set, the parse failed, so return it as a sane string
        return $@ if($@);

        # otherwise, return the workbook
        return $workbook;

    # Any other filetype gets an errpr
    } else {
        return $self -> {"template"} -> replace_langvar("GLOBAL_BADTYPE");
    }
}


## @fn @ get_worksheet_size($worksheet)
# Obtain the correct row and column minimum and maximums. This combines the worksheet
# row_range() and col_range() functions in one
sub get_worksheet_size {
    my $self      = shift;
    my $worksheet = shift;

    my ($rowmin, $rowmax) = $worksheet -> row_range();
    my ($colmin, $colmax) = $worksheet -> col_range();

    # trim empty columns off the right side.
    my $isempty = 1;
    for(; $colmax >= $colmin; --$colmax) {
        for(my $row = $rowmin; $row <= $rowmax; ++$row) {
            my $cell = $worksheet -> get_cell($row, $colmax);

            if(defined($cell)) {
                # Nuke cells that have nothing but spaces in them, as this can fuck up sheets...
                $cell -> {"Val"} =~ s/^\s+$// if($cell -> {"Val"});
                $isempty = !defined($cell -> {"Val"}) || !$cell -> {"Val"};
                last if(!$isempty);
            }
        }
        last if(!$isempty);
    }

    # And off the bottom too.
    $isempty = 1;
    for(; $rowmax >= $rowmin; --$rowmax) {
        for(my $col = $colmin; $col <= $colmax; ++$col) {
            my $cell = $worksheet -> get_cell($rowmax, $col);

            if(defined($cell)) {
                # Nuke cells that have nothing but spaces in them, as this can fuck up sheets...
                $cell -> {"Val"} =~ s/^\s+$// if($cell -> {"Val"});
                $isempty = !defined($cell -> {"Val"}) || !$cell -> {"Val"};
                last if(!$isempty);
            }
        }
        last if(!$isempty);
    }

    return ($rowmin, $rowmax, $colmin, $colmax);
}


## @fn $ get_worksheets($workbook)
# Obtain a list of worksheet names within the specified workbook. This will create
# an array of worksheet names, including empty worksheets (as there is no safe and
# reliable method around it)
#
# @param workbook The wookbooj to list worksheets in.
# @return A reference to an array of worksheet names.
sub get_worksheets {
    my $self       = shift;
    my $workbook = shift;
    my @result;

    foreach my $sheet ($workbook -> worksheets()) {
        my ($rmin, $rmax, $cmin, $cmax) = $self -> get_worksheet_size($sheet);

        # Make a nice helpful row/colum message, including an empty identifier as needed
        my $size = sprintf(" (%d columns, %d rows)", ($cmax - $cmin) + 1, ($rmax - $rmin) + 1);
        $size = " (empty worksheet)" if($cmax < $cmin || $rmax < $rmin);

        # Store away....
        push(@result, $sheet -> get_name().$size);
    }

    return \@result;
}


## @fn void mark_worksheet_merges($worksheet)
# This will go through each worksheet merge area and mark each of the cells in it
# with a merge id. This can then be used by optimise_worksheet() to recalculate
# merge areas after removing dead columns
#
# @param worksheet The worksheet to mark areas in.
sub mark_worksheet_merges {
    my $self      = shift;
    my $worksheet = shift;
    my $merged_areas = $worksheet -> get_merged_areas();

    my $areaid = 0;

    # This will do nothing if there are no areas to merge
    foreach my $area (@$merged_areas) {
        # Process all cells in the area, marking them with the areaid
        for(my $row = $area -> [0]; $row <= $area -> [2]; ++$row) {
            for(my $col = $area -> [1]; $col <= $area -> [3]; ++$col) {
                my $cell = $worksheet -> get_cell($row, $col);
                $cell -> {"mergearea"} = $areaid if($cell);
            }
        }

        # Ensure that the value for the merged area is in the top left cell
        my $tlcell = $worksheet -> get_cell($area -> [0], $area -> [1]);
        if($tlcell -> value() eq '') {
            for(my $row = $area -> [0]; $tlcell -> {"Val"} eq '' && $row <= $area -> [2]; ++$row) {
                for(my $col = $area -> [1]; $tlcell -> {"Val"} eq '' && $col <= $area -> [3]; ++$col) {
                    my $cell = $worksheet -> get_cell($row, $col);
                    $tlcell -> {"Val"} = $cell -> {"Val"} if($cell -> {"Val"} ne '');
                }
            }
        }

        ++$areaid;
    }
}


## @fn void mark_headers($id, $worksheet)
# Mark all header cells set in the database in the provided worksheet. This will go
# through all of the headers set for the upload and mark the cells in the worksheet
# as header cells accordingly.
#
# @param id        The id of this upload.
# @param worksheet The worksheet to mark header cells in.
sub mark_headers {
    my $self      = shift;
    my $id        = shift;
    my $worksheet = shift;

    # Ask the database for a list of headers for this sheet
    my $headh = $self -> {"dbh"} -> prepare("SELECT rownum, colnum
                                             FROM ".$self -> {"settings"} -> {"database"} -> {"headers"}."
                                             WHERE sheetid = ?");
    $headh -> execute($id)
         or $self -> {"logger"} -> die_log($self -> {"cgi"} -> remote_host(), "Unable to perform popup lookup: ".$self -> {"dbh"} -> errstr);

    # For each header listed in the database, mark the cell in the sheet appropriately
    while(my $header = $headh -> fetchrow_arrayref()) {
        my $cell = $worksheet -> get_cell($header -> [0], $header -> [1]);

        if($cell) {
            $cell -> {"isheader"} = 1;
        } else {
            $cell = Spreadsheet::ParseExcel::Cell -> new(isheader => 1);
            $worksheet -> {"Cells"} -> [$header -> [0]] -> [$header -> [1]] = $cell;
        }
    }
}


## @fn void optimise_worksheet($worksheet, $remove_nuked)
# Process the specified worksheet, marking colspan and rowspan information in the
# top-left, and optionally removing any cells marked as nuked from the sheet.
# Note that the overall integrity of the sheet may be compromised if nuked cells
# do not occur in columns.
#
# @param worksheet The worksheet to optimise.
# @param remove_nuked If true, remove any nuked cells from the worksheet
sub optimise_worksheet {
    my $self         = shift;
    my $worksheet    = shift;
    my $remove_nuked = shift;

    # Only bother doing the complicated stuff if nuking is eabled
    if($remove_nuked) {
        my ($rowmin, $rowmax, $colmin, $colmax) = $self -> get_worksheet_size($worksheet);

        my $outcells; # store for the sheet after removing columns

        # We can increment the out row here, as we only ever remove columns, but
        # we need separate counters as we're forcibly moving to 0,0
        my ($outc, $outr) = (0, 0);
        for(my $inr = $rowmin; $inr <= $rowmax; ++$inr, ++$outr) {
            $outc = 0;
            for(my $inc = $colmin; $inc <= $colmax; ++$inc) {
                my $incell = $worksheet -> get_cell($inr, $inc);

                if($incell) {
                    # handling of nuked cells is complicated by merges. I hate them.
                    if($incell -> {"nuke"}) {
                        # is the cell merged? If it is, does the cell have any content we need to save?
                        if($incell -> is_merged() && $incell -> {"Val"} ne '') {
                            # hell, yes it does. There's no point looking down or up, as nuked cells are in columns
                            # of nuked cells. As this is the top left, if the cell to the right is part of the same
                            # merge group, we can copy the content there and be safe.
                            my $sidecell = $worksheet -> get_cell($inr, $inc + 1);

                            $sidecell -> {"Val"} = $incell -> {"Val"}
                                if($sidecell && $sidecell -> {"mergearea"} && $sidecell -> {"mergearea"} == $incell -> {"mergearea"});
                        }
                        next;
                    }

                    # Not nuked, so just copy over
                    $outcells -> [$outr] -> [$outc] = $incell;
                }
                ++$outc;
            }
        }

        my @merges;
        my $areaid = 0;
        # Now recalculate merge areas
        for(my $row = 0; $row < $outr; ++$row) {
            for(my $col = 0; $col < $outc; ++$col) {
                my $cell = $worksheet -> get_cell($row, $col);

                # If we have a cell, and it's got a merge group, and it hasn't been processed, do it
                if($cell && $cell -> {"mergearea"} && !$cell -> {"updated"}) {
                    my ($top, $left, $bottom, $right) = ($row, $col, $row, $col);

                    my $area = $cell -> {"mergearea"};
                    # Look down vertically from where we are
                    while($cell -> {"mergearea"} == $area) {
                        # And horizontally...
                        $right = $left;
                        while($cell -> {"mergearea"} == $area) {
                            $cell -> {"mergearea"} = $areaid;
                            $cell -> {"updated"}   = 1;

                            ++$right;
                            $cell = $worksheet -> get_cell($row, $col);
                        }
                        ++$bottom;
                        $cell = $worksheet -> get_cell($row, $col);
                    }

                    # We'll always overshoot by one row and column
                    --$bottom; --$right;

                    push(@merges, [$top, $left, $bottom, $right]);
                }
            }
        }

        # Update the values in the worksheet
        $worksheet -> {"MinRow"}     = 0;
        $worksheet -> {"MaxRow"}     = $outr - 1;
        $worksheet -> {"MinCol"}     = 0;
        $worksheet -> {"MaxCol"}     = $outc - 1;
        $worksheet -> {"Cells"}      = $outcells;
        $worksheet -> {"MergedArea"} = \@merges;
    }

    # Go through each of the merge areas and set up rowspan and colspan
    foreach my $area (@{$worksheet -> {"MergedArea"}}) {
        my $tlcell = $worksheet -> get_cell($area -> [0], $area -> [1]);

        if($tlcell) {
            $tlcell -> {"isdatacell"} = 1;
            $tlcell -> {"rowspan"} = ($area -> [2] - $area -> [0]) + 1 if(($area -> [2] - $area -> [0]) > 0);
            $tlcell -> {"colspan"} = ($area -> [3] - $area -> [1]) + 1 if(($area -> [3] - $area -> [1]) > 0);
        }
    }
}

1;
