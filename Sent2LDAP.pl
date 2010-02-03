#!/usr/bin/perl 
#
# Sent2LDAP.pl (v 0.1)
#
# Copyright (c) 2010 Reinier Haasjes <reinier@haasjes.com>. All rights reserved.
# This program is free software; you can redistribute it and/or
# modify it under the same terms as Perl itself.
#
# dn: OU=Sent, CN=John Smith, OU=<user> ,DC=example,DC=com
#

use strict;
use warnings;
use diagnostics;
use diagnostics -verbose;

use Mail::Header;
use Mail::Address;
use Net::LDAP;
use Unicode::String qw(utf8 latin1 utf16);
use Encode qw/decode/;

# Variables
my $ldap_server		= "ldap.example.com";
my $homedir 		= "/home/";
my $sent_folder		= ".maildir/.Sent/cur/";

my $server      	= "localhost";
my $binddn      	= "cn=admin,dc=example,dc=com";
my $bindpasswd  	= "password";
my $base        	= "dc=example,dc=com";

# Global variables
my %recp;

# Sub declarations
sub prepare_ldap($);
sub add_recipients($); 
sub remove_sent($);
sub add_ou($);
sub mime_decode($);

# Get users
opendir(my $dh, $homedir) || die "can't opendir $homedir: $!";
my @users = grep { !/^\./ && -d "$homedir/$_" } readdir($dh);
closedir $dh;

#@users = qw/robert/; #DEBUG: only process reinier

# Connect to LDAP
my $ldap        = Net::LDAP->new( $server ) or die "$@";
#my $ldap 	= Net::LDAP->new( $server, async => 0 ) or die "$@";
$ldap->bind( $binddn, password => $bindpasswd, version => 3 );

# Process one user at a time
my ($user, $result);
foreach $user (@users)
{
#	print "## $user\n";

	# Clear array
	%recp = ();	

	prepare_ldap($user);
	
	# Get mail files
	my $sentdir = $homedir . $user . "/" . $sent_folder;
	opendir(my $dm, $sentdir) || die "can't opendir $sentdir: $!";
	my @files = grep { !/^\./ && -f "$sentdir/$_" } readdir($dm);
	closedir $dm;

	my ($file, $to, $cc, $bcc);
	my $item;
	foreach $file (@files)
	{
		open(MAILFILE, $sentdir . $file) || die ("Could not open file: " . $file);
		my $header = Mail::Header->new( \*MAILFILE );
		close (MAILFILE);

		my @tags = $header->tags;

		add_recipients($header->get("to")) if grep (/^to$/i, @tags);
		add_recipients($header->get("cc")) if grep (/^cc$/i, @tags);
		add_recipients($header->get("bcc")) if grep (/^bcc$/i, @tags);
	}

	while ( my ($k,$v) = each %recp ) {
		unless (length($v)) { $v = $k; }

		# Skip keys (email addresses)
		next if ($k =~ /\"/);	# "
		next if ($k =~ /\+/);	# +

		# Strip quotes from begin and end
		$v =~ s/^[\'\"]+//;
		$v =~ s/[\'\"]+$//;

		# decode value
		my $dv = mime_decode($v);

		$result = $ldap->add( "mail=$k, ou=Sent, ou=$user, $base", 
			attr => [
				'cn'   => "$dv",
				'sn'   => "$dv",
				'mail' => "$k",
				'objectclass' => ['top', 'person',
					'organizationalPerson',
					'inetOrgPerson' ],
				]
		);
		$result->code && warn "failed to add entry ($k -> $dv - $v): ", $result->error ;
	}
}

# Disconnect LDAP
$ldap->unbind;

### Functions
sub mime_decode($) {
	my $val = shift;

	# convert from ISO8559-1 to UTF-8
	if ($val =~ /[\x80-\xFF]/) {
		$val = Unicode::String::latin1($val);
	}

	# Decode MIME encoded fields
	$val = decode('MIME-Header', $val);

	# Return UTF-8 data
	return $val;
}

sub prepare_ldap($) {
	my $user = shift(@_);

	remove_sent($user);
	add_ou($user);
}

sub add_ou($) {
	my $user = shift(@_);

	# Add user OU
	$ldap->add( "ou=$user,$base", 
		attrs => [
			ou => "$user",
			objectClass => [ 'top', 'organizationalUnit' ]
		]
	);

	# Add Sent OU
	$ldap->add( "ou=Sent,ou=$user,$base", 
		attrs => [
			ou => 'Sent',
			objectClass => [ 'top', 'organizationalUnit' ]
		]
	);
}

sub remove_sent($) {
	my $user = shift(@_);
	my $delbranch   = "ou=Sent,ou=$user,$base";

	my $search = $ldap->search( base   => $delbranch,
				 scope  => 'one',
                                 filter => "(objectclass=*)" );
	foreach my $e (sort { $b->dn =~ tr/,// <=> $a->dn =~ tr/,// } $search->entries()) {
		$result = $ldap->delete($e);
		$result->code && warn "failed to delete entry (" . $e->dn . "): ", $result->error ;
	}
}

sub add_recipients($) {
	my $line = shift(@_);

	my @line_array = Mail::Address->parse($line);
	foreach my $item (@line_array) {
		my $addr = lc($item->address);
		if (exists $recp{$addr}) {
			my $cur_len = length($recp{$addr});
			my $new_len = length($item->phrase);
			next if ($new_len <= $cur_len);
		}
		$recp{$addr} = $item->phrase;
	}
}
