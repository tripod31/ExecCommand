use strict;

use Win32::OLE qw(in with);
my $obj = Win32::OLE->new('ExecCommand.ExecCommand') || die "Win32::OLE->new";
$obj->{ExeFile}="ping.exe";
$obj->{Arg}="localhost -n 1";

my $i = $obj->Exec();
print $obj->{StdOut};
