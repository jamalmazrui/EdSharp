#chm2txt 1.0
#August 15, 2007
#Copyright 2007 by Jamal Mazrui
#Modified GPL license

use strict;
use warnings;

use Win32::OLE qw(in valof with OVERLOAD);
use Text::CHM;
        use File::OldSlurp;
use HTML::Stripper;

my (@s, @sChms, @oFiles, @sFiles);
my ($bDir);
my ($i, $iParams, $iLength, $iChms, $iChm,$iFiles, $iFile, $iConverted);
my ($s, $sParam1, $sParam2, $sMessage, $sChm, $sSource, $sTarget, $sTemp, $sContents, $sBody, $sHtml, $sHeading, $sText, $sFiles, $sFile);
my ($o, $oChms, $oChm, $oFiles, $oFile, $oApp, $oDoc, $oBody, $oRange);
$iParams = @ARGV;
$sParam1 = $ARGV[0] if $iParams > 0;
$sParam2 = $ARGV[1] if $iParams > 1;
$s = ' -? -h -help /? /h /help ';
if (!$iParams || index($s, ' ' . lc($sParam1) . ' ') >= 0) {
$sMessage = "Syntax:\n";
$sMessage .= 'chm2txt "SourceFile.chm" "TargetFile.txt"' . "\n";
$sMessage .= "Quotes around a file may be omitted if it does not include a space\n";
$sMessage .= "Target may be omitted to produce one named like source except for extension.\n";
#$sMessage .= "The source may instead be a directory containing .chm files, with similar target paths assumed.\n";
die($sMessage);
}

$sSource = $sParam1;
$sSource =~ s/^\".*\"$//;
die("Cannot find source $sSource\n") if !-e $sSource;

if (-d $sSource) {
$bDir = 1;
@sChms = `dir /b $sSource\\*.chm`;
foreach $sChm (@sChms) {
chomp($sChm);
$sChm = $sSource . "\\" . $sChm;
}

$iChms = @sChms;
$s = 'file';
$s .= 's' unless $iChms == 1;
print "$iChms CHM $s\n";
}
else {
$bDir = 0;
@sChms = ($sSource);
$iChms = 1;

}

$sTemp = "c:\\temp\\temp.htm";
#$sTemp = $PerlApp::RUNLIB . "temp.htm";

for ($iChm = 0; $iChm < $iChms; $iChm++) {

$sSource = $sChms[$iChm];
$s = $sSource;
$s =~ s/.*\\//;
print "Converting $s\n";

if ($bDir || $iParams == 1) {
$iLength = length($sSource);
$sTarget = substr($sSource, 0, $iLength - 3) . "txt";
}
else {
$sTarget = $sParam2;
$sTarget =~ s/^\".*\"$//;
}

`del $sTarget` if -e $sTarget;

my $oChm = Text::CHM->new($sSource);
my @oFiles = $oChm->get_filelist();
$iFiles = @oFiles;
$s = 'topic';
$s .= 's' unless $iFiles == 1;
print "$iFiles $s\n";
for ($iFile = 0; $iFile < $iFiles; $iFile++) {
#print int(100 * $iFile / $iFiles) . "%\n" if $iFile && $iFile % 20 == 0;
my $oFile = $oFiles[$iFile];
$sHtml = $oChm->get_object($oFile->{path}); 

if ($sHtml && index(lc($sHtml), "<html>") >= 0) {
my $oStripper = HTML::Stripper->new(skip_cdata => 1, strip_ws => 0);
$sText = $oStripper->strip_html($sHtml);
$sText =~ s/\A\s*//;
$sText =~ s/\s*\Z//;

if ($sText) {
$sHeading = $sText;
$sHeading =~ s/(\r|\n)(.|\n)*\Z//;
$sContents .= $sHeading . "\r\n";
eval {$sText =~ s/\A$sHeading\s*$sHeading/$sHeading/i};
$sBody .= "\r\n----------\r\n\f\r\n" . $sText;
} # $sText
} # $sHtml

undef $oFile;
} # $iFile

$oChm->close();
undef @oFiles;
undef $oChm;

$s = $sSource;
$s =~ s/.*\\//;
$s =~ s/\..*\Z//;
$sContents = $s . "\r\n\r\nContents\r\n" . $sContents;
$sBody .= "\r\n----------\r\nEnd of Document\r\n";
$sBody = $sContents . $sBody;
$sBody =~ s/ +\r\n/\r\n/g;
$sBody =~ s/(\r\n){3,}/\r\n\r\n/g;
$sBody =~ s/\r\n/\n/;
$sBody =~ s/\r\r/\r/g;
$sBody =~ s/\r\n/\n/g;
write_file($sTarget, $sBody);
`del $sTemp` if -e $sTemp;

if (-e $sTarget) {
print "Done!\n";
$iConverted++;
}
else {
print "Error!\n";
}
} # $iChm 

print "converted $iConverted out of $iChms\n" if $bDir;
