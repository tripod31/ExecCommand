use strict;

use Win32::OLE qw(in with);
my $oServer = Win32::OLE->new('ExecCommand.RemoteServer') ;
my $i=$oServer->Init();
if ($i !=0){
    print("err at init server:". $oServer->{Err});
    exit(-1);
}
$oServer->{"Data"}="ARGUMENT";

#クライアント実行
if (! -f "test_client.pl"){
    print("test_client.pl not found.\n");
    exit(-1);
}

my $oExec = Win32::OLE->new('ExecCommand.ExecCommand') ;
$oExec->{ExeFile}="perl";
$oExec->{Arg}="test_client.pl";
my $i=$oExec->Exec();
if ($i!=0) {
    print("err at exec test_client.pl".$oExec->{Err});
    exit(-1);
}

print $oExec->{StdErr};
my $msg;
$msg= $oExec->{StdOut};

if ($msg eq "ARGUMENT"){
    print sprintf("Argument[%s] was passed successfully.\n",$msg);
}else{
    print sprintf("Argument was'nt passed.[%s]\n",$msg);
}

$msg = $oServer->{"Data"};
if ($msg eq "RETURN"){
    print sprintf("Return value[%s] was passed successfully.\n",$msg);
}else{
    print sprintf("Return value was'nt passed.[%s]\n",$msg);
}

$i=$oServer->Close();
if ($i !=0){
    print("err at close server:". $oServer->{Err});
    exit(-1);
}

