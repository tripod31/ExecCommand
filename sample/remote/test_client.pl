use strict;

use Win32::OLE qw(in with);
my $obj = Win32::OLE->new('ExecCommand.ExecCommand') ;
$obj->InitRemoteClient();

my $msg=$obj->{"RemoteData"};
print "Remote data read by test_client.pl:$msg";
#set remote data
$obj->{"RemoteData"} = "This string was set by test_client.pl";
