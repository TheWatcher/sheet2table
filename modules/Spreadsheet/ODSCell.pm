package Spreadsheet::ODSCell;

sub new {
    my $invocant = shift;
    my $class    = ref($invocant) || $invocant;
    my $sheet    = shift;

    # ParseExcel compatibility
    $sheet -> {"Val"} = $sheet -> {"content"};

    return bless $sheet, $class;
}


sub value {
    my $self = shift;

    return $self -> {"content"};
}


sub unformatted {
    my $self = shift;

    return $self -> {"content"};
}


sub is_merged {
    my $self = shift;

    return $self -> {"merged"};
}

1;
