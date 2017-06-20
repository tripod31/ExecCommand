use strict;

use Win32::OLE qw(in with);
my $obj = Win32::OLE->new('ExecCommand.ExecCommand') ;
$obj->InitRemoteServer();
$obj->{"ServerData"}="This string was set by test_remote.pl";

#クライアント実行
$obj->{ExeFile}="perl";
$obj->{Arg}="test_client.pl";
my $i=$obj->Exec();
if ($i!=0) {
    print("failed to run test_client.pl");
    exit(-1);
}

my $msg;
$msg= $obj->{StdOut};
print "Following string is output by test_client.pl:$msg\n";

$msg = $obj->{"ServerData"};
print "This string was read by test_remote.pl:$msg\n";
