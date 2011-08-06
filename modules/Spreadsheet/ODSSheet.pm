## @file
#
# @author  Chris Page &lt;chris@starforge.co.uk&gt;
# @version 1.1
# @date    16 June 2011
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

## @class Spreadsheet::ODSSheet
#
package Spreadsheet::ODSSheet;
use strict;

## @cmethod $ new($sheet, $name)
# Create a new ODSSheet object using the specified sheet as the basis
# for the object. This will convert the sheet to a Spreadsheet::ODSSheet
# and return it.
#
# @param sheet The spreadsheet to convert to an ODSSheet.
# @param name  The name of the spreadsheet.
# @return A new ODSSheet object.
sub new {
    my $invocant = shift;
    my $class    = ref($invocant) || $invocant;
    my $sheet    = shift;
    my $name     = shift;

    $sheet -> {"sheetname"} = $name;

    return bless $sheet, $class;
}


## @method $ get_cell($row, $col)
# Obtain a reference to the ODSCell at the specified row and column. If
# no cell exists at that location, this will return undef.
#
# @param row The row the cell is on, the top row is row 0.
# @param col The column the cell is on, the left column is col 0.
# @return A reference to the cell at the specified row and column, or
#         undef if there is no cell there.
sub get_cell {
    my $self = shift;
    my $row  = shift;
    my $col  = shift;

    return $self -> {"Cells"} -> [$row] -> [$col];
}


## @method @ row_range()
# Returns the minimum and maximum row numbers. If the sheet is empty,
# the maximum row will be less than the minimum.
#
# @return An array of two values: the minimum row number, and the maximum
#         row number.
sub row_range {
    my $self = shift;

    if($self -> {"empty"}) {
        return (0, -1);
    }

    return ($self -> {"MinRow"}, $self -> {"MaxRow"});
}


## @method @ col_range()
# Returns the minimum and maximum column numbers. If the sheet is empty,
# the maximum column will be less than the minimum.
#
# @return An array of two values: the minimum col number, and the maximum
#         col number.
sub col_range {
    my $self = shift;

    if($self -> {"empty"}) {
        return (0, -1);
    }

    return ($self -> {"MinCol"}, $self -> {"MaxCol"});
}


## @method $ get_name()
# Returns the name of the current sheet.
#
# @return The current sheet name.
sub get_name {
    my $self = shift;

    return $self -> {"sheetname"};
}


# @method $ get_merged_areas()
# Obtains the list of merged areas in the spreadsheet. This will return a
# reference to an array of merged areas, each element in the array will
# be a reference to an array containing [start row, start col, end row, end col]
#
# @return A reference to an array of merged areas.
sub get_merged_areas {
    my $self = shift;

    return $self -> {"MergedArea"};
}

1;
