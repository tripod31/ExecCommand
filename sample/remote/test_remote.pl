use strict;

use Win32::OLE qw(in with);
my $obj = Win32::OLE->new('ExecCommand.ExecCommand') ;
$obj->InitRemoteServer();
$obj->{"ServerData"}="ARGUMENT";

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
if ($msg eq "ARGUMENT"){
    print "Argument was passed succeccfully.\n";
}else{
    print "Argument was'nt passed.";
}

$msg = $obj->{"ServerData"};
if ($msg eq "RETURN"){
    print "Return value was passed succeccfully.\n";
}else{
    print "Return value was'nt passed.";
}
