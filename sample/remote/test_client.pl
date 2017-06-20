use strict;
use Win32::OLE qw(in with);

my $obj = Win32::OLE->new('ExecCommand.RemoteClient') ;
my $i=$obj->InitRemoteClient();
if ($i!=0) {
    print(sprintf("err at init client:%s\n",$obj->{Err}));
    exit(-1);
}

my $msg=$obj->{"RemoteData"};
print sprintf("%s",$msg);
#set remote data
$obj->{"RemoteData"} = "RETURN";
