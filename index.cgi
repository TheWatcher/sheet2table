#!/usr/bin/perl -wT

use strict;
use lib qw(/home/webperl);
use lib qw(modules);
use utf8;

# System modules
use CGI::Compress::Gzip qw/:standard -utf8/;   # Enabling utf8 here is kinda risky, with the file uploads, but eeegh
use CGI::Carp qw(fatalsToBrowser set_message); # Catch as many fatals as possible and send them to the user as well as stderr
use DBI;
use Digest;
use Encode;
use File::Copy;
use HTML::Entities;
use MIME::Base64;
use Time::HiRes qw(time);

# Webperl modules
use ConfigMicro;
use Logging qw(start_log end_log die_log);
use HTMLValidator;
use Template;
use Utils qw(path_join is_defined_numeric get_proc_size);

# local modules
use SheetTools;

my $dbh;                                   # global database handle, required here so that the END block can close the database connection
my $contact = 'webmaster@starforge.co.uk'; # global contact address, for error messages

# install more useful error handling
BEGIN {
    $ENV{"PATH"} = ""; # Force no path.

    delete @ENV{qw(IFS CDPATH ENV BASH_ENV)}; # Clean up ENV
    sub handle_errors {
        my $msg = shift;
        print "<h1>Software error</h1>\n";
        print '<p>Server time: ',scalar(localtime()),'<br/>Error was:</p><pre>',$msg,'</pre>';
        print '<p>Please report this error to ',$contact,' giving the text of this error and the time and date at which it occured</p>';
    }
    set_message(\&handle_errors);
}
END {
    # Nicely close the database connection. Possibly not vital, but good to be sure..
    $dbh -> disconnect() if($dbh);

    # Stop logging if it has been enabled.
    end_log();
}

# IDs of the stages
use constant STAGE_UPLOAD  => 0;
use constant STAGE_OPTIONS => 1;
use constant STAGE_HEADERS => 2;
use constant STAGE_POPUPS  => 3;
use constant STAGE_OUTPUT  => 4;

use constant GC_TIME       => 3600;
use constant RETAIN_TIME   => 86400;

# Stages in the process
my $stages = [ { "active"   => "templates/default/images/stage/upload_active.png",
                 "inactive" => "templates/default/images/stage/upload_inactive.png",
                 "passed"   => "templates/default/images/stage/upload_passed.png",
                 "width"    => 80,
                 "height"   => 40,
                 "alt"      => "Upload File",
                 "icon"     => "upload",
                 "func"     => \&build_stage0_upload },
               { "active"   => "templates/default/images/stage/options_active.png",
                 "inactive" => "templates/default/images/stage/options_inactive.png",
                 "passed"   => "templates/default/images/stage/options_passed.png",
                 "width"    => 80,
                 "height"   => 40,
                 "alt"      => "Select options",
                 "icon"     => "picksheet",
                 "hasback"  => 1,
                 "func"     => \&build_stage1_options },
               { "active"   => "templates/default/images/stage/headers_active.png",
                 "inactive" => "templates/default/images/stage/headers_inactive.png",
                 "passed"   => "templates/default/images/stage/headers_passed.png",
                 "width"    => 80,
                 "height"   => 40,
                 "alt"      => "Choose headers",
                 "icon"     => "headers",
                 "hasback"  => 1,
                 "func"     => \&build_stage2_headers },
               { "active"   => "templates/default/images/stage/popups_active.png",
                 "inactive" => "templates/default/images/stage/popups_inactive.png",
                 "passed"   => "templates/default/images/stage/popups_passed.png",
                 "width"    => 80,
                 "height"   => 40,
                 "alt"      => "Choose popups",
                 "icon"     => "popups",
                 "hasback"  => 1,
                 "func"     => \&build_stage3_popups },
               { "active"   => "templates/default/images/stage/output_active.png",
                 "inactive" => "templates/default/images/stage/output_inactive.png",
                 "passed"   => "templates/default/images/stage/output_passed.png",
                 "width"    => 80,
                 "height"   => 40,
                 "alt"      => "Output ready",
                 "icon"     => "tables",
                 "hasback"  => 1,
                 "func"     => \&build_stage4_output } ];

# Approved html tags
my @approved_html = ("a", "pre", "code", "br", "object", "embed",
                     "table", "tr", "td", "th", "tbody", "thead",
                     "ul", "ol", "li",
                     "dl", "dt", "dd",
                     "h1", "h2", "h3", "h4", "h5", "h6", "h7",
                     "hr",
                     "sub","sup",
                     "tt", "b", "i", "u", "div", "span", "strong", "blockquote",

    );

# output function hash
my %converters = (
    "format_html"      => \&worksheet_to_html,
    "format_mediawiki" => \&worksheet_to_mediawiki,
    );

# =============================================================================
#  Generation code

## @fn $ revert_property_quotes($content)
# Given a string representing the contents of a html tag, convert any &quot; entities
# back to " to ensure proper parsing by the html cleaning code.
#
# @param content The contents of the html tag to process.
# @return The contents with &quot;s replaced by "s
sub revert_property_quotes {
    my $content = shift;

    $content =~ s/&quot;/"/g;
    return "<".$content.">";
}


## @fn $ cleanup_cell_html($content)
# Remove any potentially dangerous content from the specified cell. This will sanitise
# any content that could break the table when generated as html. Note that this is
# still potentially a security risk in practice - as any situation that allows user-side
# html is
#
# @param content The cell contents to fix up.
# @return The 'safe' cell contents.
sub cleanup_cell_html {
    my $content = shift;

    return '' if(!$content);

    $content = encode_entities($content);

    # Now fix up approved tags
    foreach my $tag (@approved_html) {
        $content =~ s{&lt;(/?\s*$tag.*?)&gt;}{<$1>}gis;
    }

    # Do we have any literal <s at this point now? IF so, we have one or more
    # html tags we need to check up on
    if($content =~ /</) {
        # Restore the quotes around element properties
        $content =~ s/<([^>]+)>/revert_property_quotes($1)/igse;

        # Scrub the html to remove any particularly nasty crap (like dynamic stylesheets),
        # and get a first pass guess at html validity.
        $content = scrub_html($content)
            or return "Unable to process html content in cell.";

        # Get htmltidy to clean it up some more - scrubbing helps remove dangerous stuff,
        # but may not have removed all possible crap, this should.
        $content = tidy_html($content, { tidy_mark      => 0,
                                         output_xhtml   => 1,
                                         merge_spans    => 0,
                                         merge_divs     => 0,
                                         show_body_only => 1,
                             })
            or return "Unable to run tidy over content in cell.";
    }

    $content =~ s/^\s*(.*?)\s*$/$1/;

    return $content;
}


## @fn $ format_popup_html($sysargs, $anchor, $body)
# Generate a html format popup with the specified anchor and body.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @param anchor  The text to show on the page where the popup should be, may contain html or wiki markup.
# @param body    The body of the popup, may contain html or wiki markup.
# @return The formatted popup tag.
sub format_popup_html {
    my $sysargs = shift;
    my $anchor  = shift;
    my $body    = shift;

    return "<span class=\"twpopup\">$anchor<span class=\"twpopup-inner\">".encode_base64(encode("UTF-8", $body), '')."</span></span>";
}


## @fn $ format_popup_mediawiki($sysargs, $anchor, $body)
# Generate a mediawiki-format popup with the specified anchor and body.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @param anchor  The text to show on the page where the popup should be, may contain html or wiki markup.
# @param body    The body of the popup, may contain html or wiki markup.
# @return The formatted popup tag.
sub format_popup_mediawiki {
    my $sysargs = shift;
    my $anchor  = shift;
    my $body    = shift;

    return "<popup title=\"$anchor\">$body</popup>";
}


## @fn void process_popups($sysvars, $id, $worksheet, $formatter)
# Process the specified worksheet, merging cells marked as belonging to popups into
# the title cell. This goes through each popup column pair set for this sheet and
# shoves the body cell data into the title cell via the formatter function.
#
# @param sysvars   A reference to a hash containing the system template, database, cgi, and settings objects.
# @param id        The id of the upload being processed.
# @param worksheet The worksheet containing the data to process.
# @param formatter A reference to a function that takes three arguments - the sysargs,
#                  the popup anchor text, and the popup body - and returns a string
#                  containing the popup.
sub process_popups {
    my $sysvars   = shift;
    my $id        = shift;
    my $worksheet = shift;
    my $formatter = shift;

    # Ask the database for a list of popup columns for this worksheet
    my $poph = $sysvars -> {"dbh"} -> prepare("SELECT title_col, body_col
                                               FROM ".$sysvars -> {"settings"} -> {"database"} -> {"popups"}."
                                               WHERE sheetid = ?");
    $poph -> execute($id)
        or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to perform popup lookup: ".$sysvars -> {"dbh"} -> errstr);

    # For each popup, go down the title column, merging in the content from the
    # body column and marking the body cell as removable. If the cell in the
    # body col is empty, no popup is generated.
    my ($minrow, $maxrow) = $sysvars -> {"sheet"} -> get_worksheet_size($worksheet);
    while(my $popup = $poph -> fetchrow_hashref()) {
        # For each row, process the popup...
        for(my $row = $minrow; $row <= $maxrow; ++$row) {
            # Get the two cells, if possible
            my $title_cell = $worksheet -> get_cell($row, $popup -> {"title_col"});

            if($title_cell) {
                print STDERR "Got cell at $row, ".$popup -> {"title_col"};

                # If the title cell is in a merge, what we really want is the top left cell of the merge
                if($title_cell -> is_merged()) {
                    my $areas = $worksheet -> get_merged_areas();
                    my $area = $areas -> [$title_cell -> {"mergearea"}];

                    # Get the top left cell
                    $title_cell = $worksheet -> get_cell($area -> [0], $area -> [1]);
                    print STDERR "Moved to merge base at $area->[0], $area->[1]";
                }

                my $body_cell  = $worksheet -> get_cell($row, $popup -> {"body_col"});

                # Don't try doing anything if we have no body cell, there's no point.
                if($body_cell) {
                    print STDERR "Got body at $row, ".$popup -> {"body_col"};
                    # Mark the body cell as junk for later killing
                    $body_cell -> {"nuke"} = 1;

                    # trim whitespace, so that popup bodies do not end up parsed as pre blocks in mediawiki
                    $body_cell -> {"Val"} =~ s/^\s*(.*?)\s*$/$1/;

                    # Set the popup up, provided we have a title cell with a value, and the title cell is not
                    # a header and hasn't previously been processed as a popup.
                    if($title_cell && defined($title_cell -> {"Val"}) && !$title_cell -> {"isheader"} && !$title_cell -> {"popup"} && $body_cell -> {"Val"}) {
                        $title_cell -> {"Val"} = $formatter -> ($sysvars, $title_cell -> {"Val"}, $body_cell -> {"Val"});
                        $body_cell -> {"Val"} = ''; # Remove the content.
                        $title_cell -> {"popup"} = 1;
                    }
                }
            }
        }
    }
}


## @fn $ get_popup_colmap($sysvars, $id)
# Obtain a hash mapping column numbers to popup anchors and bodies.
#
# @param sysvars   A reference to a hash containing the system template, database, cgi, and settings objects.
# @param id        The id of the upload being processed.
# @return A reference to a hash mapping columns to pop ids and types
sub get_popup_colmap {
    my $sysvars = shift;
    my $id      = shift;
    my $pophash;

    # Ask the database for a list of popup columns for this worksheet
    my $poph = $sysvars -> {"dbh"} -> prepare("SELECT popupid, title_col, body_col
                                               FROM ".$sysvars -> {"settings"} -> {"database"} -> {"popups"}."
                                               WHERE sheetid = ?");
    $poph -> execute($id)
        or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to perform popup lookup: ".$sysvars -> {"dbh"} -> errstr);

    while(my $popup = $poph -> fetchrow_hashref()) {
        $pophash -> {$popup -> {"title_col"}} = {"id" => $popup -> {"popupid"},
                                                 "type" => "anchor"};
        $pophash -> {$popup -> {"body_col"}}  = {"id" => $popup -> {"popupid"},
                                                 "type" => "body"};
    }

    return $pophash;
}


## @fn void preprocess_worksheet($sysvars, $id, $worksheet, $formatter, $do_headers, $do_popups)
# Run preprocessing and optimisation functions on the worksheet to prepare it for
# conversion to the output format.
#
# @param sysvars    A reference to a hash containing the system template, database, cgi, and settings objects.
# @param id         The id of the upload being processed.
# @param worksheet  The worksheet containing the data to process.
# @param formatter  A reference to a formatter function for process_popups, may be undef if do_popups is false.
# @param do_headers Enable processing of headers.
# @param do_popups  Enable processing of popups.
sub preprocess_worksheet {
    my $sysvars    = shift;
    my $id         = shift;
    my $worksheet  = shift;
    my $do_headers = shift;
    my $do_popups  = shift;
    my $formatter  = shift;

    # Mark all the merge areas so that we can do spanning and optimisation more easily
    $sysvars -> {"sheet"} -> mark_worksheet_merges($worksheet);

    # If we are processing headers, go and mark them
    $sysvars -> {"sheet"} -> mark_headers($id, $worksheet) if($do_headers);

    # If we're doing popups, process them now
    process_popups($sysvars, $id, $worksheet, $formatter) if($do_popups);

    # Remove columns emptied by popups, and speed up spanning.
    $sysvars -> {"sheet"} -> optimise_worksheet($worksheet, $do_popups);
}


## @fn $ worksheet_to_html($sysvars, $worksheet, $options)
# Convert a worksheet to html, potentially including headers and popups. The valid
# settings that can be made in the options hash are:
#
# show_headers Enable the controls to set headers.
# do_headers   Enable processing of headers.
# show_popups  Enable the controls to set popups.
# do_popups    Enable processing of popups.
# do_zebra     Enable zebra tables.
# preview      Enable preview mode.
#
# @param sysvars   A reference to a hash containing the system template, database, cgi, and settings objects.
# @param id        The id of the upload being processed.
# @param worksheet The worksheet containing the data to process.
# @param options   A reference to a hash of options to control output processing.
# @return A string containing the html table
sub worksheet_to_html {
    my $sysvars    = shift;
    my $id         = shift;
    my $worksheet  = shift;
    my $options    = shift;
    my $table;

    # Forcibly turn off other options if show_popups is on
    $options -> {"show_headers"} = $options -> {"do_popups"} = $options -> {"do_headers"} = 0
        if($options -> {"show_popups"});

    # And turn off other options if show headers is on.
    $options -> {"show_popups"} = $options -> {"do_popups"} = $options -> {"do_headers"} = 0
        if($options -> {"show_headers"});

    # Set everything up ready for generation
    preprocess_worksheet($sysvars, $id, $worksheet, $options -> {"do_headers"} || $options -> {"show_headers"}, $options -> {"do_popups"}, \&format_popup_html);

    my ($rowmin, $rowmax, $colmin, $colmax) = $sysvars -> {"sheet"} -> get_worksheet_size($worksheet);

    my $colmap = get_popup_colmap($sysvars, $id);
    my ($modes, $anchors, $bodies, $nextid) = ("", "", "", 0);

    # Okay, we're as close to sorted as we're going to get
    $table  = "<table";
    $table .= " class=\"zebra\"" if($options -> {"do_zebra"});
    $table .= " id=\"preview\"" if($options -> {"preview"});
    $table .= ">\n";

    # If we have popups controls enabled, we need to preface everything by a row of controls
    if($options -> {"show_popups"}) {
        $table .= "    <tr>";
        for(my $col = $colmin; $col <= $colmax; ++$col) {
            my ($mode, $extra) = ("", "");

            if($colmap -> {$col}) {
                $mode = ($colmap -> {$col} -> {"type"} eq "anchor") ? "anchor" : "body";
                $extra = $sysvars -> {"template"} -> replace_langvar("POPUP_ID", {"***id***" => $colmap -> {$col} -> {"id"}});

                $modes   .= "modes[$col] = '$mode';\n";
                $anchors .= "anchors[$col] = ".($mode eq "anchor" ? $colmap -> {$col} -> {"id"} : "-1").";\n";
                $bodies  .= "bodies[$col] = ".($mode eq "body" ? $colmap -> {$col} -> {"id"} : "-1").";\n";
                $nextid  = $colmap -> {$col} -> {"id"} + 1 if(($mode eq "anchor" || $mode eq "body") && $colmap -> {$col} -> {"id"} >= $nextid);
            } else {
                $mode = "anchor_add";
                $modes   .= "modes[$col] = 'anchor_add';\n";
                $anchors .= "anchors[$col] = -1;\n";
                $bodies  .= "bodies[$col] = -1;\n";
            }
            $table .= $sysvars -> {"template"} -> load_template("blocks/popuppick.tem", {"***col***"   => $col,
                                                                                         "***mode***"  => $mode,
                                                                                         "***extra***" => $extra});
        }
        $table .= "</tr>\n";
    }

    # Force preview if show headers or popups is enabled
    $options -> {"preview"} = $options -> {"show_headers"} || $options -> {"show_popups"};

    # Now go through the cells themselves...
    for(my $row = $rowmin; $row <= $rowmax; ++$row) {
        $table .= "    <tr>";
        $table .= $sysvars -> {"template"} -> load_template("blocks/rowpick.tem", {"***row***" => $row})
            if($options -> {"show_headers"});

        for(my $col = $colmin; $col <= $colmax; ++$col) {
            my $cell = $worksheet -> get_cell($row, $col) || {};

            # Skip non-data merged cells
            next if(defined($cell -> {"mergearea"}) && !$cell -> {"isdatacell"});

            $table .= ($cell -> {"isheader"} && $options -> {"do_headers"}) ? '<th' : '<td';
            if($options -> {"show_headers"} || $options -> {"show_popups"}) {
                my $class = $options -> {"show_headers"} ? "sethead" : "";
                $class .= " isanchor" if($options -> {"show_popups"} && $colmap -> {$col} -> {"type"} eq "anchor");
                $class .= " isbody"   if($options -> {"show_popups"} && $colmap -> {$col} -> {"type"} eq "body");
                $class .= " ishead"   if($options -> {"show_headers"} && $cell -> {"isheader"});
                $table .= " id=\"r${row}c${col}\" ";
                $table .= " class=\"$class\"" if($class);
            }
            $table .= ' colspan="'.$cell -> {"colspan"}.'"' if($cell -> {"colspan"});
            $table .= ' rowspan="'.$cell -> {"rowspan"}.'"' if($cell -> {"rowspan"});
            $table .= ">".cleanup_cell_html($cell -> {"Val"});
            $table .= ($cell -> {"isheader"} && $options -> {"do_headers"}) ? '</th>' : '</td>';
        }
        $table .= "</tr>\n";
    }

    $table .= "</table>";

    # Need to tack on the stuff for the popups if it has been set
    $table .= $sysvars -> {"template"} -> load_template("blocks/popupjs.tem", {"***modes***"   => $modes,
                                                                               "***anchors***" => $anchors,
                                                                               "***bodies***"  => $bodies,
                                                                               "***nextid***"  => $nextid})
        if($modes && $anchors && $bodies);

    return $table;
}


## @fn $ worksheet_to_html($sysvars, $worksheet, $options)
# Convert a worksheet to mediawiki markup, potentially including headers and popups. The valid
# settings that can be made in the options hash are:
#
# do_headers   Enable processing of headers.
# do_popups    Enable processing of popups.
# do_zebra     Enable zebra tables.
#
# @param sysvars   A reference to a hash containing the system template, database, cgi, and settings objects.
# @param id        The id of the upload being processed.
# @param worksheet The worksheet containing the data to process.
# @param options   A reference to a hash of options to control output processing.
# @return A string containing the mediawiki table
sub worksheet_to_mediawiki {
    my $sysvars    = shift;
    my $id         = shift;
    my $worksheet  = shift;
    my $options    = shift;
    my $table;

    die "medaiwiki converter can not provide previews or interactive settings"
        if($options -> {"preview"} || $options -> {"show_headers"} || $options -> {"show_popups"});

    # Set everything up ready for generation
    preprocess_worksheet($sysvars, $id, $worksheet, $options -> {"do_headers"}, $options -> {"do_popups"}, \&format_popup_mediawiki);

    my ($rowmin, $rowmax, $colmin, $colmax) = $sysvars -> {"sheet"} -> get_worksheet_size($worksheet);

    $table  = "{|";
    $table .= " class=\"zebra\"" if($options -> {"do_zebra"});

    for(my $row = $rowmin; $row <= $rowmax; ++$row) {
        $table .= "\n|-\n";

        for(my $col = $colmin; $col <= $colmax; ++$col) {
            my $cell = $worksheet -> get_cell($row, $col) || {};

            # Skip non-data merged cells
            next if(defined($cell -> {"mergearea"}) && !$cell -> {"isdatacell"});

            $table .= ($cell -> {"isheader"} && $options -> {"do_headers"}) ? '!' : '|';
            $table .= ($cell -> {"isheader"} && $options -> {"do_headers"}) ? '!' : '|' if($col > $colmin); # Repeat for cells other than the first in line

            $table .= ' colspan="'.$cell -> {"colspan"}.'"' if($cell -> {"colspan"});
            $table .= ' rowspan="'.$cell -> {"rowspan"}.'"' if($cell -> {"rowspan"});
            $table .= '|' if($cell -> {"colspan"} || $cell -> {"rowspan"});
            $table .= cleanup_cell_html($cell -> {"Val"});
        }
    }

    $table .= "\n|}\n";

    return $table;
}



# =============================================================================
#  Database interaction

## @fn $ get_formats($sysvars)
# Obtain a list of supported output formats.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @return A reference to an array of output formats.
sub get_formats {
    my $sysvars = shift;
    my @result;

    my $fmth = $sysvars -> {"dbh"} -> prepare("SELECT name
                                               FROM ".$sysvars -> {"settings"} -> {"database"} -> {"formats"}."
                                               ORDER BY id");
    $fmth -> execute()
        or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to perform format lookup: ".$sysvars -> {"dbh"} -> errstr);

    while(my $format = $fmth -> fetchrow_arrayref()) {
        push(@result, $format -> [0]);
    }

    return \@result;
}


## @fn $ get_upload_data($sysvars, $id)
# Obtain the data for the upload identified by the specified id, provided that the
# address stored for the upload matches the address the request is coming from.
#
# @note This function can be set to only return an upload record when some or all of
#       the ip address of the current user matches the ip address recorded for the
#       upload. This is controlled via the ip_security setting: if the security level
#       is 0 then any IP will match and the record can potentially be obtained by
#       anyone. Levels 1 to 3 match increasing levels of ip range: at level 1, the
#       upper octet must match, level two the upper two octets must match, etc. Level
#       4 requires an exact IP match, and represents the most secure option. Level 4
#       (and possibly level 3) may produce false denials if the user is behind a
#       load-balanced proxy.
# @todo IPv6 compatibility needed here.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @param id      The ID of the upload to obtain the data for
# @return A reference to a hash containing the upload data, or undef if the ID is not valid.
sub get_upload_data {
    my $sysvars = shift;
    my $id      = shift;

    # Work out the address we must match against
    my @addrbits = split(/\./, $sysvars -> {"cgi"} -> remote_addr());
    my $address = "";
    my $bit = 0;

    # build up the string to match against the remote_addr field
    for(; $bit < $sysvars -> {"settings"} -> {"config"} -> {"ip_security"}; ++$bit) {
        $address .= $addrbits[$bit];
        $address .= "." if($bit < 3);
    }
    $address .= "%" if($bit < 4);

    my $sheeth = $sysvars -> {"dbh"} -> prepare("SELECT s.*,f.function
                                                 FROM ".$sysvars -> {"settings"} -> {"database"} -> {"sheets"}." AS s
                                                 LEFT JOIN ".$sysvars -> {"settings"} -> {"database"} -> {"formats"}." AS f
                                                 ON f.id = s.output_type
                                                 WHERE s.id = ?
                                                 AND s.remote_addr LIKE ?");
    $sheeth -> execute($id, $address)
        or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to perform sheet lookup: ".$sysvars -> {"dbh"} -> errstr);

    return $sheeth -> fetchrow_hashref();
}


## @fn $ start_upload($sysvars)
# Take the file uploaded from the user, perform some basic checks on it, and if
# it seems to check out, copy it into the storage directory and create a new
# record for it.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @return A reference to a hash containing the new upload's database data, or
#         a string containing an error message if there was a problem.
sub start_upload {
    my $sysvars = shift;

    # First, do we have a file to grab?
    my $file = $sysvars -> {'cgi'} -> upload("excelfile");

    # bomb if the file is not found
    return $sysvars -> {"template"} -> replace_langvar("UPLOAD_ERR_NOFILE") if(!$file);

    my $namearg = $sysvars -> {'cgi'} -> param("excelfile");

    # Strip any path information from the name
    my ($name) = $namearg =~ m{^(?:.*[\\/])?([^\\/]+)$};

    # Fix non-ascii
    $name =~ s{[^\w_\-\.]}{_}g;

    # Get the extension, if possible
    my ($ext) = $name =~ /\.([^\.]+)$/;
    return $sysvars -> {"template"} -> replace_langvar("UPLOAD_ERR_NOEXT") if(!$ext);

    # The extension must be .xls, .xlsx, or .ods
    return $sysvars -> {"template"} -> replace_langvar("UPLOAD_ERR_BADEXT")
        if($ext !~ /^xlsx?$/i && $ext !~ /^ods$/i);

    # Okay, if we get here, the file is probably a excel workbook, so we can start doing things
    # First, we need to create a target name
    my $sha256 = Digest -> new("SHA-256");
    $sha256 -> add($file, $sysvars -> {"cgi"} -> remote_host(), time(), $$);
    my $destname = $sha256 -> hexdigest();

    # copy over...
    my $src_tainted = $sysvars -> {'cgi'} -> tmpFileName($file);
    my ($srcfile) = $src_tainted =~ /^(.*)$/; # horribly messy, but we need to trust cgi, so...

    my $destfile =  path_join($sysvars -> {"settings"} -> {"config"} -> {"file_dir"}, $destname);
    copy($srcfile, $destfile)
        or return $sysvars -> {"template"} -> replace_langvar("UPLOAD_ERR_BADCOPY", {"***error***" => $!});

    # create a new entry in the database for this
    my $sheeth = $sysvars -> {"dbh"} -> prepare("INSERT INTO ".$sysvars -> {"settings"} -> {"database"} -> {"sheets"}."
                                                 (source_name, local_name, file_type, set_headers, remote_addr, last_updated)
                                                 VALUES(?, ?, ?, 1, ?, UNIX_TIMESTAMP())");
    $sheeth -> execute($name, $destname, lc($ext), $sysvars -> {"cgi"} -> remote_addr())
        or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to perform new sheet insert: ".$sysvars -> {"dbh"} -> errstr);

    my $newid = $sysvars -> {"dbh"} -> {"mysql_insertid"};
    my $entry = get_upload_data($sysvars, $newid);

    return $entry || $sysvars -> {"template"} -> replace_langvar("UPLOAD_ERR_BADID");
}


## @fn void touch_sheet($sysvars, $id)
# Update the last updated time on the specified sheet to the current time. This is
# used to mark the sheet as updated each time an operation is performed on resources
# associated with it.
sub touch_sheet {
    my $sysvars = shift;
    my $id      = shift;

    my $touch = $sysvars -> {"dbh"} -> prepare("UPDATE ".$sysvars -> {"settings"} -> {"database"} -> {"sheets"}."
                                                SET last_updated = UNIX_TIMESTAMP()
                                                WHERE id = ?");
    $touch -> execute($id)
        or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to touch sheet record. Error was: ".$sysvars -> {"dbh"} -> errstr);
}


## @fn $ set_options($sysvars, $id, $workbook)
# Set the options data (active worksheet, output format, zebra flag, etc) for the
# specified upload based on the contents of posted data. This will validate the
# values provided by the user, update the worksheet data accordingly, and return
# any errors that may have occurred.
#
# @param sysvars  A reference to a hash containing the system template, database, cgi, and settings objects.
# @param id       The id of the upload being processed.
# @param workbook The workbook the upload corresponds to.
# @return undef on success, otherwise this returns a string containing error messages.
sub set_options {
    my $sysvars  = shift;
    my $id       = shift;
    my $workbook = shift;
    my $args     = {};
    my ($errors, $error);

    # First the easy check - do we have a valid worksheet?
    ($args -> {"sheet_num"}, $error) = valid_numeric_option($sysvars, "worksheet", {"nicename" => $sysvars -> {"template"} -> replace_langvar("WORKSHEET_NAME"),
                                                                                    "required" => 1,
                                                                                    "minimum"  => 0,
                                                                                    "maximum"  => $workbook -> worksheet_count() - 1 });
    $errors .= $error."<br/>" if($error);

    # Output formats need to check against the database
    ($args -> {"output_type"}, $error) = valid_numeric_option($sysvars, "output", {"nicename" => $sysvars -> {"template"} -> replace_langvar("WORKSHEER_OUTT"),
                                                                                   "required" => 1,
                                                                                   "table"    => $sysvars -> {"settings"} -> {"database"} -> {"formats"},
                                                                                   "column"   => "id"});
    $errors .= $error."<br/>" if($error);

    # The zebra switch is either on or off...
    $args -> {"zebra"} = defined($sysvars -> {"cgi"} -> param("zebra"));

    # As are the headers and popups flags
    $args -> {"set_headers"} = defined($sysvars -> {"cgi"} -> param("headers"));
    $args -> {"set_popups"}  = defined($sysvars -> {"cgi"} -> param("popups"));

    # Okay, even if we have errors, at this point we can update the table - the values
    # will either be NULL, which is safe, or something valid, which is also safe...
    my $seth = $sysvars -> {"dbh"} -> prepare("UPDATE ".$sysvars -> {"settings"} -> {"database"} -> {"sheets"}."
                                               SET sheet_num = ?, output_type = ?, zebra = ?, set_headers = ?, set_popups = ?, last_updated = UNIX_TIMESTAMP()
                                               WHERE id = ?");
    $seth -> execute($args -> {"sheet_num"},
                     $args -> {"output_type"},
                     $args -> {"zebra"},
                     $args -> {"set_headers"},
                     $args -> {"set_popups"},
                     $id)
        or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to update sheet data: ".$sysvars -> {"dbh"} -> errstr);

    return $errors;
}


## @fn $ set_header_cells($sysvars, $id)
# Set the head cells for the specified upload based on the contents of posted data.
# This will go through the values specified by the user and update the data in the
# database accordingly.
#
# @param sysvars  A reference to a hash containing the system template, database, cgi, and settings objects.
# @param id       The id of the upload being processed.
sub set_header_cells {
    my $sysvars = shift;
    my $id      = shift;

    # Remove all the old header entries first, as we need them not at all
    my $nukeheads = $sysvars -> {"dbh"} -> prepare("DELETE FROM ".$sysvars -> {"settings"} -> {"database"} -> {"headers"}."
                                                    WHERE sheetid = ?");
    $nukeheads -> execute($id)
        or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to remove old header data for ID $id: ".$sysvars -> {"dbh"} -> errstr);

    # Prepare a query to shove in new entries
    my $newheads = $sysvars -> {"dbh"} -> prepare("INSERT INTO ".$sysvars -> {"settings"} -> {"database"} -> {"headers"}."
                                                   (sheetid, colnum, rownum)
                                                   VALUES(?, ?, ?)");

    my $hlist = $sysvars -> {"cgi"} -> param("hlist");
    if($hlist) {
        # headers should be encoded as r<num>c<num>; so we can just do a progressive match
        # over the whole string, and shove the results into the database.
        while($hlist =~ /r(\d+)c(\d+);/g) {
            my ($row, $col) = ($1, $2);

            $newheads -> execute($id, $col, $row)
                or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to create new header data for ID $id: ".$sysvars -> {"dbh"} -> errstr);
        }
    }

    touch_sheet($sysvars, $id);
}


## @fn $ set_popup_cols($sysvars, $id)
# Set the popup columns for the specified upload based on the contents of posted data.
# This will go through the values specified by the user and update the data in the
# database accordingly.
#
# @param sysvars  A reference to a hash containing the system template, database, cgi, and settings objects.
# @param id       The id of the upload being processed.
sub set_popup_cols {
    my $sysvars = shift;
    my $id      = shift;

    # Remove all the old popup entries first, as we need them not at all
    my $nukepops = $sysvars -> {"dbh"} -> prepare("DELETE FROM ".$sysvars -> {"settings"} -> {"database"} -> {"popups"}."
                                                   WHERE sheetid = ?");
    $nukepops -> execute($id)
        or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to remove old popup data for ID $id: ".$sysvars -> {"dbh"} -> errstr);

    # Prepare a query to shove in new entries
    my $newpops = $sysvars -> {"dbh"} -> prepare("INSERT INTO ".$sysvars -> {"settings"} -> {"database"} -> {"popups"}."
                                                  (sheetid, popupid, title_col, body_col)
                                                  VALUES(?, ?, ?, ?)");

    my $plist = $sysvars -> {"cgi"} -> param("plist");
    if($plist) {
        # popups should be encoded as a<num>b<num>; so we can just do a progressive match
        # over the whole string, and shove the results into the database.
        my $popid = 0;
        while($plist =~ /a(\d+)b(\d+);/g) {
            my ($anchor, $body) = ($1, $2);

            $newpops -> execute($id, $popid++, $anchor, $body)
                or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to create new popup data for ID $id: ".$sysvars -> {"dbh"} -> errstr);
        }
    }

    touch_sheet($sysvars, $id);
}


## @fn $ garbage_collect($sysvars)
# Delete old content from the database and filesystem to prevent cruft from piling up.
# This will go through the list of sheets stored in the database and, if any are older
# than the retention period, any popups or headers associated with the sheet are deleted
# and the sheet file is removed from the filesystem before the entry itself is removed.
sub garbage_collect {
    my $sysvars = shift;
    my $now = time();

    # Do nothing if we're still in the last gc period
    return if($sysvars -> {"settings"} -> {"config"} -> {"last_gc"} + GC_TIME > $now);

    # We need to garbage collect, set the time so no other scripts try it for a while
    my $setgc = $sysvars -> {"dbh"} -> prepare("UPDATE ".$sysvars -> {"settings"} -> {"database"} -> {"settings"}."
                                                SET value = UNIX_TIMESTAMP()
                                                WHERE name LIKE 'last_gc'");
    $setgc -> execute()
        or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to set gargage collect time: ".$sysvars -> {"dbh"} -> errstr);

    my $retainto = time() - RETAIN_TIME;

    # Prepare some queries for later...
    my $deadheads   = $sysvars -> {"dbh"} -> prepare("DELETE FROM ".$sysvars -> {"settings"} -> {"database"} -> {"headers"}."
                                                      WHERE sheetid = ?");
    my $deadpops    = $sysvars -> {"dbh"} -> prepare("DELETE FROM ".$sysvars -> {"settings"} -> {"database"} -> {"popups"}."
                                                      WHERE sheetid = ?");
    my $deadsheets  = $sysvars -> {"dbh"} -> prepare("DELETE FROM ".$sysvars -> {"settings"} -> {"database"} -> {"sheets"}."
                                                      WHERE id = ?");

    my $oldsheets = $sysvars -> {"dbh"} -> prepare("SELECT id, local_name FROM ".$sysvars -> {"settings"} -> {"database"} -> {"sheets"}."
                                                     WHERE last_updated < ?");
    $oldsheets -> execute($retainto)
        or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to begin garbage collect sequence: ".$sysvars -> {"dbh"} -> errstr);

    while(my $sheet = $oldsheets -> fetchrow_hashref()) {
        $deadheads -> execute($sheet -> {"id"})
            or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to garbage collect headers for sheet ".$sheet -> {"id"}.": ".$sysvars -> {"dbh"} -> errstr);

        $deadpops -> execute($sheet -> {"id"})
            or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to garbage collect popups for sheet ".$sheet -> {"id"}.": ".$sysvars -> {"dbh"} -> errstr);

        unlink(path_join($sysvars -> {"settings"} -> {"config"} -> {"file_dir"}, $sheet -> {"local_name"}))
            or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to remove file for sheet ".$sheet -> {"id"}.": $!");

        $deadsheets -> execute($sheet -> {"id"})
            or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to garbage collect sheet ".$sheet -> {"id"}.": ".$sysvars -> {"dbh"} -> errstr);
    }
}


# =============================================================================
#  Options and validation support

## @fn $ build_options($sysvars, $optlist, $select)
# Build a string containing html option elements, one for each entry in the
# optlist array provided. If select is specified, the option it corresponds
# to will be selected.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @param optlist A reference to an array of options.
# @param select  An optional default selection value.
# @return A string containing the list of options.
sub build_options {
    my $sysvars = shift;
    my $optlist = shift;
    my $select  = shift;
    my $result  = "";

    for(my $id = 0; $id < scalar(@$optlist); ++$id) {
        $result .= "<option value=\"$id\"";
        $result .= ' selected="selected"' if(defined($select) && $select == $id);
        $result .= ">".$sysvars -> {"template"} -> html_clean($optlist -> [$id])."</option>\n";
    }

    return $result;
}


## @fn @ valid_numeric_option($sysvars, $param, $args)
# Determine whether the value specified for the named option is valid. This will
# attempt to ensure that the value is numeric, and lies within the specified
# range or exists in a specified table. Valid entries in the args hash are:
#
# nicename   - The human-readable name of the field.
# required   - If true, the argument must be provided.
# minimum    - Optional minimum value to check the option against.
# maximum    - Optional maximum value to check the option against.
# table      - Optional Table to check the option against.
# column     - the field in the column to check the option against, must be given
#              if the table is specified.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @param param   The name of the parameter to validate.
# @param args    A reference to a hash of arguments for the validation
# @return An array of two values: the first is the value set for the option, or
#         undef on error, the second is the error message if an error was encountered.
sub valid_numeric_option {
    my $sysvars = shift;
    my $param   = shift;
    my $args    = shift;

    # First, do we even have a value?
    my $value = is_defined_numeric($sysvars -> {"cgi"}, $param);

    # If the value is missing, and one is required, fall over
    return (undef, $sysvars -> {"template"} -> replace_langvar("GLOBAL_VALOPT_NOVALUE", {"***field***" => $args -> {"nicename"}}))
        if(!defined($value) && $args -> {"required"});

    # Check against the minimum if needed
    return (undef, $sysvars -> {"template"} -> replace_langvar("GLOBAL_VALOPT_LOWVALUE", {"***field***"   => $args -> {"nicename"},
                                                                                          "***minimum***" => $args -> {"minimum"}}))
        if(defined($args -> {"minimum"}) && defined($value) && $value < $args -> {"minimum"});

    # And the same for the maximum
    return (undef, $sysvars -> {"template"} -> replace_langvar("GLOBAL_VALOPT_HIVALUE", {"***field***"   => $args -> {"nicename"},
                                                                                         "***maximum***" => $args -> {"maximum"}}))
        if(defined($args -> {"maximum"}) && defined($value) && $value > $args -> {"maximum"});

    # If we have a table and column, check the value there
    if(defined($value) && $args -> {"table"} && $args -> {"column"}) {
        my $checkh = $sysvars -> {"dbh"} -> prepare("SELECT `".$args -> {"column"}."`
                                                     FROM `".$args -> {"table"}."`
                                                     WHERE `".$args -> {"column"}."` = ?");
        $checkh -> execute($value)
            or die_log($sysvars -> {"cgi"} -> remote_host(), "Unable to perform validation lookup: ".$sysvars -> {"dbh"} -> errstr);

        my $checkr = $checkh -> fetchrow_arrayref();

        return (undef, $sysvars -> {"template"} -> replace_langvar("GLOBAL_VALOPT_BADVALUE", {"***field***" => $args -> {"nicename"}}))
            if(!$checkr);
    }

    # Get here and it has passed all tests...
    return ($value, undef);
}


# =============================================================================
#  Wizard step functions

## @fn @ build_stage0_upload($sysvars, $error)
# Generate a box containing the message to show on the upload step. This step allows
# the user to select a file to upload to begin the conversion process.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @param error   An optional error message to show in the form.
# @return An array of two values: the title, and the message box.
sub build_stage0_upload {
    my $sysvars = shift;
    my $error   = shift;

    # If we have an error, encapsulate it
    $error = $sysvars -> {"template"} -> load_template("blocks/stage_error.tem", {"***error***" => $error})
        if($error);

    # Make a nice session digest for the upload
    my $sha256 = Digest -> new("SHA-256");
    $sha256 -> add($sysvars -> {"cgi"} -> remote_host(), time(), $$);

    # Now generate the title, message.
    my $title    = $sysvars -> {"template"} -> replace_langvar("UPLOAD_TITLE");
    my $message  = $sysvars -> {"template"} -> wizard_box($sysvars -> {"template"} -> replace_langvar("UPLOAD_TITLE"),
                                                          $error ? "warn" : $stages -> [0] -> {"icon"},
                                                          $stages, 0,
                                                          $sysvars -> {"template"} -> replace_langvar("UPLOAD_LONGDESC"),
                                                          $sysvars -> {"template"} -> load_template("blocks/stage0form.tem", {"***error***"  => $error,
                                                                                                                              "***sessid***" => $sha256 -> hexdigest()}));
    return ($title, $message);
}


## @fn @ build_stage1_options($sysvars)
# Generate a box containing the message to show on the options selection step.
# This step allows the user to select which of the worksheets in the workbook should
# be converted, and the format the output should be in, and other options.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @return An array of two values: the title, and the message box.
sub build_stage1_options {
    my $sysvars = shift;
    my $error;
    my $entry;

    my $id = is_defined_numeric($sysvars -> {"cgi"}, "uid");

    # Do we have an upload id? If we have an id, try to fetch the data for it
    if($id) {
        $entry = get_upload_data($sysvars, $id) or
            $sysvars -> {"template"} -> replace_langvar("GLOBAL_BADID");

    # If we don't have an id, we're being called from stage0, so start an upload
    } else {
        $entry = start_upload($sysvars);
    }

    # If entry is not a hashref here, it's an error message
    return build_stage0_upload($sysvars, $entry) if(!ref($entry));

    # We have an ID at this point, now we need to form a list of worksheets, so we
    # need to actually load the workbook...
    my $workbook = $sysvars -> {"sheet"} -> load_workbook($entry -> {"local_name"}, $entry -> {"file_type"});

    # If workbook is not a reference, it is an error message
    return build_stage0_upload($sysvars, "Error:".$workbook) if(!ref($workbook));

    my $sheetlist = $sysvars -> {"sheet"} -> get_worksheets($workbook);
    my $formatlist = get_formats($sysvars);

    # If we are being called after the user has set the options, set them...
    $error = set_options($sysvars, $id, $workbook)
        if(defined($sysvars -> {"cgi"} -> param("setopts")));

    my($title, $message);

    # If the set has been done, and we have no errors, generate the review page
    if(defined($sysvars -> {"cgi"} -> param("setopts")) && !defined($error)) {
        # Need to get an updated copy of the data
        $entry = get_upload_data($sysvars, $id);

        # The stage we go to next depends on whether we have header or popups enabled
        my $stage = $entry -> {"set_headers"} ? STAGE_HEADERS : $entry -> {"set_popups"} ? STAGE_POPUPS : STAGE_OUTPUT;

        # Now we need human readable stuff for the options
        my $zebra   = $sysvars -> {"template"} -> replace_langvar("OPTIONS_ZEBRA_"  .($entry -> {"zebra"}       ? "ON" : "OFF"));
        my $headers = $sysvars -> {"template"} -> replace_langvar("OPTIONS_HEADERS_".($entry -> {"set_headers"} ? "ON" : "OFF"));
        my $popups  = $sysvars -> {"template"} -> replace_langvar("OPTIONS_POPUPS_" .($entry -> {"set_popups"}  ? "ON" : "OFF"));

        # Now generate the title, message.
        $title    = $sysvars -> {"template"} -> replace_langvar("OPTIONS_TITLE");
        $message  = $sysvars -> {"template"} -> wizard_box($sysvars -> {"template"} -> replace_langvar("OPTIONS_TITLE"),
                                                           $error ? "warn" : $stages -> [1] -> {"icon"},
                                                           $stages, 1,
                                                           $sysvars -> {"template"} -> replace_langvar("OPTIONS_LONGDESC"),
                                                           $sysvars -> {"template"} -> load_template("blocks/stage1conf.tem", {"***sheet***"   => $sheetlist -> [$entry -> {"sheet_num"}],
                                                                                                                               "***format***"  => $formatlist -> [$entry -> {"output_type"}],
                                                                                                                               "***zebra***"   => $zebra,
                                                                                                                               "***headers***" => $headers,
                                                                                                                               "***popups***"  => $popups,
                                                                                                                               "***stage***"   => $stage,
                                                                                                                               "***uid***"     => $entry -> {"id"}}));
    # If we are being called from stage 0, or there are errors, show the options page...
    } else { # if(defined($sysvars -> {"cgi"} -> param("setops")) && !defined($error)) {

        # Okay, we have the workbook, make me a list of worksheet names
        my $worksheets = build_options($sysvars, $sheetlist, $entry -> {"sheet_num"});

        # And a list of possible output formats
        my $formats = build_options($sysvars, $formatlist, $entry -> {"output_type"});

        # right, now we can make the page... If we have an error, encapsulate it
        $error = $sysvars -> {"template"} -> load_template("blocks/stage_error.tem", {"***error***" => $error})
            if($error);

        # Now generate the title, message.
        $title    = $sysvars -> {"template"} -> replace_langvar("OPTIONS_TITLE");
        $message  = $sysvars -> {"template"} -> wizard_box($sysvars -> {"template"} -> replace_langvar("OPTIONS_TITLE"),
                                                           $error ? "warn" : $stages -> [1] -> {"icon"},
                                                           $stages, 1,
                                                           $sysvars -> {"template"} -> replace_langvar("OPTIONS_LONGDESC"),
                                                           $sysvars -> {"template"} -> load_template("blocks/stage1form.tem", {"***error***"   => $error,
                                                                                                                               "***sheets***"  => $worksheets,
                                                                                                                               "***formats***" => $formats,
                                                                                                                               "***zebra***"   => $entry -> {"zebra"}       ? 'checked="checked"' : '',
                                                                                                                               "***headers***" => $entry -> {"set_headers"} ? 'checked="checked"' : '',
                                                                                                                               "***popups***"  => $entry -> {"set_popups"}  ? 'checked="checked"' : '',
                                                                                                                               "***uid***"     => $entry -> {"id"}}));
    }
    return ($title, $message);
}


## @fn @ build_stage2_headers($sysvars)
# Generate a box containing the message to show on the headers selection step.
# This step allows the user to select which of the cells in the worksheet should
# be marked as headers, and the format the output should be in, and other options.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @return An array of two values: the title, and the message box.
sub build_stage2_headers {
    my $sysvars = shift;
    my $entry;

    my $id = is_defined_numeric($sysvars -> {"cgi"}, "uid");

    # Do we have an upload id? If we have an id, try to fetch the data for it
    if($id) {
        $entry = get_upload_data($sysvars, $id) or
            $sysvars -> {"template"} -> replace_langvar("GLOBAL_BADID");
    } else {
        $entry = $sysvars -> {"template"} -> replace_langvar("GLOBAL_NOID");
    }

    # If entry is not a hashref here, it's an error message. Drop back to stage 0 here,
    # as bad/missing IDs are not recoverable at stage 1.
    return build_stage0_upload($sysvars, $entry) if(!ref($entry));

    # We have an ID at this point, now we need to form a list of worksheets, so we
    # need to actually load the workbook...
    my $workbook = $sysvars -> {"sheet"} -> load_workbook($entry -> {"local_name"}, $entry -> {"file_type"});

    # If workbook is not a reference, it is an error message. Again, drop back to 0
    # as a broken upload is not recoverable at stage 1.
    return build_stage0_upload($sysvars, $workbook) if(!ref($workbook));

    # We can now safely get a worksheet!
    my $worksheet = $workbook -> worksheet($entry -> {"sheet_num"});

    my($title, $message);

    # If we're not setting, we've just been called and need to show the set page...
    if(!defined($sysvars -> {"cgi"} -> param("sethead"))) {
        my $table = worksheet_to_html($sysvars, $id, $worksheet, {"show_headers" => 1, "preview" => 1});

        # We need the row and column range from the worksheet
        my ($minrow, $maxrow, $mincol, $maxcol) = $sysvars -> {"sheet"} -> get_worksheet_size($worksheet);

        $title    = $sysvars -> {"template"} -> replace_langvar("HEADERS_TITLE");
        $message  = $sysvars -> {"template"} -> wizard_box($sysvars -> {"template"} -> replace_langvar("HEADERS_TITLE"),
                                                           $stages -> [2] -> {"icon"},
                                                           $stages, 2,
                                                           $sysvars -> {"template"} -> replace_langvar("HEADERS_LONGDESC"),
                                                           $sysvars -> {"template"} -> load_template("blocks/stage2form.tem", {"***uid***"    => $id,
                                                                                                                               "***table***"  => $table,
                                                                                                                               "***mincol***" => $mincol,
                                                                                                                               "***maxcol***" => $maxcol,
                                                                                                                               "***minrow***" => $minrow,
                                                                                                                               "***maxrow***" => $maxrow}));
    # We can has submission
    } else {
        # Record the header selection
        set_header_cells($sysvars, $id);

        # Generate the table with the new headers set
        my $table = worksheet_to_html($sysvars, $id, $worksheet, {"do_headers" => 1, "do_zebra" => 1});

        # The stage we go to next depends on whether we have popups enabled
        my $stage = $entry -> {"set_popups"} ? STAGE_POPUPS : STAGE_OUTPUT;

        $title    = $sysvars -> {"template"} -> replace_langvar("HEADERS_TITLE");
        $message  = $sysvars -> {"template"} -> wizard_box($sysvars -> {"template"} -> replace_langvar("HEADERS_TITLE"),
                                                           $stages -> [2] -> {"icon"},
                                                           $stages, 2,
                                                           $sysvars -> {"template"} -> replace_langvar("HEADERS_LONGDESC"),
                                                           $sysvars -> {"template"} -> load_template("blocks/stage2conf.tem", {"***uid***"    => $id,
                                                                                                                               "***stage***"  => $stage,
                                                                                                                               "***table***"  => $table}));
    }

    return ($title, $message);
}


## @fn @ build_stage3_popups($sysvars)
# Generate a box containing the message to show on the popups selection step.
# This step allows the user to select which of the columns in the selected
# worksheet should be merged into popups.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @return An array of two values: the title, and the message box.
sub build_stage3_popups {
    my $sysvars = shift;
    my $entry;

    my $id = is_defined_numeric($sysvars -> {"cgi"}, "uid");

    # Do we have an upload id? If we have an id, try to fetch the data for it
    if($id) {
        $entry = get_upload_data($sysvars, $id) or
            $sysvars -> {"template"} -> replace_langvar("GLOBAL_BADID");
    } else {
        $entry = $sysvars -> {"template"} -> replace_langvar("GLOBAL_NOID");
    }

    # If entry is not a hashref here, it's an error message. Drop back to stage 0 here,
    # as bad/missing IDs are not recoverable at stage 1.
    return build_stage0_upload($sysvars, $entry) if(!ref($entry));

    # We have an ID at this point, now we need to form a list of worksheets, so we
    # need to actually load the workbook...
    my $workbook = $sysvars -> {"sheet"} -> load_workbook($entry -> {"local_name"}, $entry -> {"file_type"});

    # If workbook is not a reference, it is an error message. Again, drop back to 0
    # as a broken upload is not recoverable at stage 1.
    return build_stage0_upload($sysvars, $workbook) if(!ref($workbook));

    # We can now safely get a worksheet!
    my $worksheet = $workbook -> worksheet($entry -> {"sheet_num"});

    my($title, $message);

    # If we're not setting, we've just been called and need to show the set page...
    if(!defined($sysvars -> {"cgi"} -> param("setpops"))) {
        my $table = worksheet_to_html($sysvars, $id, $worksheet, {"show_popups" => 1, "preview" => 1});

        # We need the row and column range from the worksheet
        my ($minrow, $maxrow, $mincol, $maxcol) = $sysvars -> {"sheet"} -> get_worksheet_size($worksheet);

        $title    = $sysvars -> {"template"} -> replace_langvar("POPUPS_TITLE");
        $message  = $sysvars -> {"template"} -> wizard_box($sysvars -> {"template"} -> replace_langvar("POPUPS_TITLE"),
                                                           $stages -> [3] -> {"icon"},
                                                           $stages, 3,
                                                           $sysvars -> {"template"} -> replace_langvar("POPUPS_LONGDESC"),
                                                           $sysvars -> {"template"} -> load_template("blocks/stage3form.tem", {"***uid***"    => $id,
                                                                                                                               "***pid***"    => $entry -> {"set_headers"} ? "2" : "1",
                                                                                                                               "***table***"  => $table,
                                                                                                                               "***mincol***" => $mincol,
                                                                                                                               "***maxcol***" => $maxcol,
                                                                                                                               "***minrow***" => $minrow,
                                                                                                                               "***maxrow***" => $maxrow}));
    } else {
        # Record the popup selection
        set_popup_cols($sysvars, $id);

        # Generate the table with the new popups set
        my $table = worksheet_to_html($sysvars, $id, $worksheet, {"do_headers" => 1, "do_popups" => 1, "do_zebra" => 1});

        $title    = $sysvars -> {"template"} -> replace_langvar("POPUPS_TITLE");
        $message  = $sysvars -> {"template"} -> wizard_box($sysvars -> {"template"} -> replace_langvar("POPUPS_TITLE"),
                                                           $stages -> [3] -> {"icon"},
                                                           $stages, 3,
                                                           $sysvars -> {"template"} -> replace_langvar("POPUPS_LONGDESC"),
                                                           $sysvars -> {"template"} -> load_template("blocks/stage3conf.tem", {"***uid***"    => $id,
                                                                                                                               "***table***"  => $table}));

    }

    return ($title, $message);
}


## @fn @ build_stage4_output($sysvars)
# Generate a box containing the message to show on the output step.
# This step shows the user the table preview, and allows them to copy the appropriate
# code out of the output box.
#
# @param sysvars A reference to a hash containing the system template, database, cgi, and settings objects.
# @return An array of two values: the title, and the message box.
sub build_stage4_output {
    my $sysvars = shift;
    my $entry;

    my $id = is_defined_numeric($sysvars -> {"cgi"}, "uid");

    # Do we have an upload id? If we have an id, try to fetch the data for it
    if($id) {
        $entry = get_upload_data($sysvars, $id) or
            $sysvars -> {"template"} -> replace_langvar("GLOBAL_BADID");
    } else {
        $entry = $sysvars -> {"template"} -> replace_langvar("GLOBAL_NOID");
    }

    # If entry is not a hashref here, it's an error message. Drop back to stage 0 here,
    # as bad/missing IDs are not recoverable at stage 1.
    return build_stage0_upload($sysvars, $entry) if(!ref($entry));

    # We have an ID at this point, now we need to form a list of worksheets, so we
    # need to actually load the workbook...
    my $workbook = $sysvars -> {"sheet"} -> load_workbook($entry -> {"local_name"}, $entry -> {"file_type"});

    # If workbook is not a reference, it is an error message. Again, drop back to 0
    # as a broken upload is not recoverable at stage 1.
    return build_stage0_upload($sysvars, $workbook) if(!ref($workbook));

    # We can now safely get a worksheet!
    my $worksheet = $workbook -> worksheet($entry -> {"sheet_num"});

    my($title, $message);

    # Make the preview table...
    my $showtable = worksheet_to_html($sysvars, $id, $worksheet, {"do_popups"  => $entry -> {"set_popups"},
                                                                  "do_headers" => $entry -> {"set_headers"},
                                                                  "do_zebra"   => 1});

    $workbook = $sysvars -> {"sheet"} -> load_workbook($entry -> {"local_name"}, $entry -> {"file_type"});

    # If workbook is not a reference, it is an error message. Again, drop back to 0
    # as a broken upload is not recoverable at stage 1.
    return build_stage0_upload($sysvars, $workbook) if(!ref($workbook));

    # We can now safely get a worksheet!
    $worksheet = $workbook -> worksheet($entry -> {"sheet_num"});

    # And generate the actual markup
    my $realtable = $converters{$entry -> {"function"}} -> ($sysvars, $id, $worksheet, {"do_popups"  => $entry -> {"set_popups"},
                                                                                        "do_headers" => $entry -> {"set_headers"},
                                                                                        "do_zebra"   => $entry -> {"zebra"}});

    # The stage we go to when hitting Back depends on whether we have header or popups enabled
    my $stage = $entry -> {"set_popups"} ? STAGE_POPUPS : $entry -> {"set_headers"} ? STAGE_HEADERS : STAGE_OPTIONS;

    $title    = $sysvars -> {"template"} -> replace_langvar("OUTPUT_TITLE");
    $message  = $sysvars -> {"template"} -> wizard_box($sysvars -> {"template"} -> replace_langvar("OUTPUT_TITLE"),
                                                       $stages -> [4] -> {"icon"},
                                                       $stages, 4,
                                                       $sysvars -> {"template"} -> replace_langvar("OUTPUT_LONGDESC"),
                                                       $sysvars -> {"template"} -> load_template("blocks/stage4form.tem", {"***uid***"        => $id,
                                                                                                                           "***pid***"        => $stage,
                                                                                                                           "***showtable***"  => $showtable,
                                                                                                                           "***realtable***"  => $sysvars -> {"template"} -> html_clean($realtable)}));
    return ($title, $message);
}


# =============================================================================
#  Core page code and dispatcher

## @fn $ page_display($sysvars)
# Generate the contents of the page based on the current step in the wizard.
#
# @param sysvars A reference to a hash containing references to the template,
#                database, settings, and cgi objects.
# @return A string containing the page to display.
sub page_display {
    my $sysvars = shift;
    my ($title, $body, $extrahead) = ("", "", "");

    # Clean up before we do anything
    garbage_collect($sysvars);

    # Get the current stage, and make it zero if there's no stage defined
    my $stage = is_defined_numeric($sysvars -> {"cgi"}, "stage");
    $stage = 0 if(!defined($stage));

    # Check that the stage is in range, fix it if not
    $stage = scalar(@$stages) - 1 if($stage >= scalar(@$stages));

    # some stages may provide a back button, in which case we may need to go back a stage if the back is pressed...
    if(defined($sysvars -> {"cgi"} -> param('back')) && $stage > 0 && $stages -> [$stage - 1] -> {"hasback"}) {
        my $bstage = is_defined_numeric($sysvars -> {"cgi"}, "bstage");
        $stage = $bstage if(defined($bstage));
    }

    # Do we have a function?
    my $func = $stages -> [$stage] -> {"func"}; # these two lines could be done in one, but it would look horrible...
    ($title, $body, $extrahead) = $func -> ($sysvars) if($func);

    return $sysvars -> {"template"} -> load_template("page.tem",
                                                     { "***title***"     => $title,
                                                       "***extrahead***" => $extrahead,
                                                       "***core***"      => $body || '<p class="error">No page content available, this should not happen.</p>'});
}

## @fn $ cgi_upload_hook(($filename, $buffer, $bytes_read, $file)
# Upload hook required when uploading files
sub cgi_upload_hook {
    my ($filename, $buffer, $bytes_read, $file) = @_;

    # Get our sessid from the form submission. Assumes the form was submitted as
    # index.cgi?<sessid>
    my ($querysessid) = $ENV{QUERY_STRING};
    my ($sessid) = $querysessid =~ /^([a-fA-F0-9]+)/;

    # Calculate the (rough estimation) of the file size. This isn't
    # accurate because the CONTENT_LENGTH includes not only the file's
    # contents, but also the length of all the other form fields as well,
    # so it's bound to be at least a few bytes larger than the file size.
    my $length = $ENV{'CONTENT_LENGTH'};
    my $percent = 0;

    # Work out the percentage if there is a length
    $percent = sprintf("%.1f", (($bytes_read / $length) * 100))
        if($length);

    # Write this data to the session file.
    open(SES, ">./uploadsess/$sessid.session");
    print SES "$bytes_read:$length:$percent";
    close(SES);
}

my $starttime = time();

# Create a new CGI object to generate page content through
my $out = CGI::Compress::Gzip -> new(\&cgi_upload_hook);

# Load the system config
my $settings = ConfigMicro -> new("config/site.cfg")
    or die_log($out -> remote_host(), "index.cgi: Unable to obtain configuration file: ".$ConfigMicro::errstr);

# Database initialisation. Errors in this will kill program.
$dbh = DBI->connect($settings -> {"database"} -> {"database"},
                    $settings -> {"database"} -> {"username"},
                    $settings -> {"database"} -> {"password"},
                    { RaiseError => 0, AutoCommit => 1, mysql_enable_utf8 => 1 })
    or die_log($out -> remote_host(), "index.cgi: Unable to connect to database: ".$DBI::errstr);

# Pull configuration data out of the database into the settings hash
$settings -> load_db_config($dbh, $settings -> {"database"} -> {"settings"});

# Start doing logging if needed
start_log($settings -> {"config"} -> {"logfile"}) if($settings -> {"config"} -> {"logfile"});

# Create the template handler object
my $template = Template -> new(basedir => path_join($settings -> {"config"} -> {"base"}, "templates"))
    or die_log($out -> remote_host(), "Unable to create template handling object: ".$Template::errstr);

# And the excel tools object
my $sheet = SheetTools -> new("template" => $template,
                              "dbh"      => $dbh,
                              "settings" => $settings,
                              "cgi"      => $out);

my $content = page_display({"template" => $template,
                            "dbh"      => $dbh,
                            "settings" => $settings,
                            "cgi"      => $out,
                            "sheet"    => $sheet});
print $out -> header(-charset => 'utf-8');

my $endtime = time();
my ($user, $system, $cuser, $csystem) = times();
my $debug = $template -> load_template("debug.tem", {"***secs***"   => sprintf("%.2f", $endtime - $starttime),
                                                     "***user***"   => $user,
                                                     "***system***" => $system,
                                                     "***memory***" => $template -> bytes_to_human(get_proc_size())});

print Encode::encode_utf8($template -> process_template($content, {"***debug***" => $debug}));
$template -> set_module_obj(undef);

