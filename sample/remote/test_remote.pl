use strict;

use Win32::OLE qw(in with);
my $obj = Win32::OLE->new('ClassLibraryForVBA.ExecCommand') ;
$obj->InitRemoteServer();
$obj->{"ServerData"}="[server]hello!";

#クライアント起動
$obj->{ExeFile}="perl";
$obj->{Arg}="test_client.pl";
$obj->Exec();
my $msg;
$msg= $obj->{StdOut};
print "[server]client output:$msg:\n";

$msg = $obj->{"ServerData"};
print "[server]msg from client:$msg\n";
