#!/usr/bin/perl -wT

# Support script required to obtain the status information for uploads.

use strict;
use CGI;

# make a CGI object to make life easier
my $cgi = CGI -> new();

# Now go to work...
print $cgi -> header();

# obtain the id of the session to get progress for
my $qsessid = $cgi -> param("sessid");
my ($sessid) = $qsessid =~ /^([a-fA-F0-9]+)/;

# Do we have a session?
if($sessid) {
    # Does the progress file for the session exist?
    if (-f "./uploadsess/$sessid.session") {

        # Read the session file and print its contents to the caller
        open (READ, "./uploadsess/$sessid.session");
        my $data = <READ>;
        close (READ);
        print $data;

    # No status file, send back errors
    } else {
        print "0:0:0:Session $sessid doesn't exist.";
    }

# No session specified, send back errors..
} else {
    print "0:0:0:No session specified.";
}
