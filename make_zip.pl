use strict;
use warnings;
use Archive::ZIP;
use File::Find;
use Cwd;
use File::Copy 'copy';

my $APP_NAME="ClassLibraryForVBA";
my @FILES=("readme.txt","src.zip","regist_dll.bat","unregist_dll.bat","bin/release/ClassLibraryForVBA.dll","bin/release/ClassLibraryForVBA.tlb");
my @DIRS=("sample");
my $DIST_DIR="D:/data_priv/HP/myHP/files";

my $ZIP_FILE="$APP_NAME.zip";
my $zip;

#ソースZIP
$zip = Archive::Zip->new();
print "creating src.zip-->\n";
find({wanted=>\&process_file,no_chdir=>1}, ("./"));
$zip->writeToFileNamed("src.zip");
print "<--src.zip\n";

#ZIP
$zip = Archive::Zip->new();
for my $file (@FILES) {
    $zip->addFile($file);
}
for my $dir (@DIRS) {
    $zip->addTree($dir,$dir);
}
$zip->writeToFileNamed($ZIP_FILE);
copy $ZIP_FILE,$DIST_DIR || die $!;

#ソースファイル判定
sub hantei{
    my $path=shift;
    my $bRet=1;

    #除外パス
    if ( $path =~ /\/(bin|obj|sample)/) {
        $bRet= 0;
    }
    #除外ファイル
    if ( $path =~ /(make_zip|\.zip|readme\.txt|\.bat)/) {
        $bRet= 0;
    }

    return $bRet;
}

#ソースファイル処理
sub process_file{
    #$File::Find::dir   #カレントディレクトリ
    #$_                 #ファイル名

    #フルパスのファイル名
    my $path=  $File::Find::name;

    if (&hantei($path)==0) {
        print "skip:$path\n";
        return;
    }

    print "add :$path\n";
    if (-d $path){
        $zip->addDirectory($path);
    }else{
        $zip->addFile($path);
    }
}

