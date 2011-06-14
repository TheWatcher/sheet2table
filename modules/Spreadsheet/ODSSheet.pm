package Spreadsheet::ODSSheet;

use strict;

sub new {
    my $invocant = shift;
    my $class    = ref($invocant) || $invocant;
    my $sheet    = shift;
    my $name     = shift;

    $sheet -> {"sheetname"} = $name;

    return bless $sheet, $class;
}

sub get_cell {
    my $self = shift;
    my $row  = shift;
    my $col  = shift;

    return $self -> {"table:table-row"} -> [$row] -> {"table:table-cell"} -> [$col];
}


sub row_range {
    my $self = shift;

    if($self -> {"empty"}) {
        return (0, -1);
    }

    return (0, $self -> {"rows"} - 1);
}


sub col_range {
    my $self = shift;

    if($self -> {"empty"}) {
        return (0, -1);
    }

    return (0, $self -> {"cols"} - 1);
}


sub get_name {
    my $self = shift;

    return $self -> {"sheetname"};
}


sub get_merged_areas {
    my $self = shift;

    return $self -> {"merges"};
}

1;
