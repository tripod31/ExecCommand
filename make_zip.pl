use strict;
use warnings;
use Archive::ZIP;
use File::Find;
use Cwd;
use File::Copy 'copy';

my $APP_NAME="ClassLibraryForVBA";
my @FILES=("readme.txt","regist_dll.bat","unregist_dll.bat","bin/release/ClassLibraryForVBA.dll","bin/release/ClassLibraryForVBA.tlb");
my @DIRS=("sample");

my $ZIP_FILE="$APP_NAME.zip";
my $zip;

#ZIP
$zip = Archive::Zip->new();
for my $file (@FILES) {
    $zip->addFile($file);
}
for my $dir (@DIRS) {
    $zip->addTree($dir,$dir);
}
$zip->writeToFileNamed($ZIP_FILE);

