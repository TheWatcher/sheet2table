package Spreadsheet::ODSCell;

sub new {
    my $invocant = shift;
    my $class    = ref($invocant) || $invocant;
    my $cell     = shift;

    # ParseExcel compatibility
    $cell -> {"Val"} = $cell -> {"content"};

    return bless $cell, $class;
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
