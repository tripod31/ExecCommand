use strict;

use Win32::OLE qw(in with);
my $obj = Win32::OLE->new('ClassLibraryForVBA.ExecCommand') ;
$obj->InitRemoteServer();
$obj->{"ServerData"}="[server]hello!";
system("perl test_client.pl");	#クライアント起動
