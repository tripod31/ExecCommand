use strict;

use Win32::OLE qw(in with);
my $obj = Win32::OLE->new('ClassLibraryForVBA.ExecCommand') ;
$obj->InitRemoteClient();

my $msg=$obj->{"RemoteData"};
print "[client]msg from server:[$msg]\n";
