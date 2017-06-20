use strict;

use Win32::OLE qw(in with);
my $obj = Win32::OLE->new('ExecCommand.ExecCommand') ;
$obj->InitRemoteClient();

my $msg=$obj->{"RemoteData"};
print $msg;
#set remote data
$obj->{"RemoteData"} = "RETURN";
