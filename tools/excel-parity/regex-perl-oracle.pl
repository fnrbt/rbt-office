#!/usr/bin/perl
# perl tools/excel-parity/regex-perl-oracle.pl < regex-cases.jsonl > regex-perl.jsonl
# For each case of make-regex-cases.py: whether perl compiles the pattern (with /a, so that \w \d \s and
# the POSIX classes are ASCII, as PCRE2's are without UCP; /i for a caseless case), the groups of the
# first match (start and end in code points, -1 unset) and every match of a global search.
use strict;
use warnings;
no warnings;
use JSON::PP;
binmode STDIN, ':encoding(UTF-8)';
binmode STDOUT, ':encoding(UTF-8)';
my $json = JSON::PP->new->canonical(1);
while (my $line = <STDIN>) {
    chomp $line;
    next unless length $line;
    my $c = $json->decode($line);
    my ($p, $s) = ($c->{pattern}, $c->{subject});
    my $re = eval { $c->{i} ? qr/(?a)(?i)$p/ : qr/(?a)$p/ };
    my %out = (pattern => $p, subject => $s, i => $c->{i});
    if (!defined $re) {
        $out{error} = JSON::PP::true;
    } else {
        $out{error} = JSON::PP::false;
        if ($s =~ $re) {
            my @g;
            for my $k (0 .. $#+) {
                push @g, (defined $-[$k] ? ($-[$k], $+[$k]) : (-1, -1));
            }
            $out{first} = \@g;
        } else {
            $out{first} = undef;
        }
        my @all;
        my $t = $s;
        while ($t =~ /$re/g) {
            push @all, [$-[0], $+[0]];
            last if @all > 1000;
        }
        $out{all} = \@all;
    }
    print $json->encode(\%out), "\n";
}
