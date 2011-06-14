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

    return $self -> {"Cells"} -> [$row] -> [$col];
}


sub row_range {
    my $self = shift;

    if($self -> {"empty"}) {
        return (0, -1);
    }

    return ($self -> {"MinRow"}, $self -> {"MaxRow"});
}


sub col_range {
    my $self = shift;

    if($self -> {"empty"}) {
        return (0, -1);
    }

    return ($self -> {"MinCol"}, $self -> {"MaxCol"});
}


sub get_name {
    my $self = shift;

    return $self -> {"sheetname"};
}


sub get_merged_areas {
    my $self = shift;

    return $self -> {"MergeArea"};
}

1;
