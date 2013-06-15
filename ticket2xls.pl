#<!-- -*- encoding: utf-8n -*- -->
use strict;
use warnings;
use utf8;

use Encode;
use Encode::JP;

use LWP::UserAgent;
use HTTP::Request::Common;

use XML::Simple;
use Data::Dumper;
#use XML::XPath;

use Spreadsheet::WriteExcel;

use open IN   => ":encoding(cp932)"; # 入力ファイルはcp932
use open OUT  => ":encoding(cp932)"; # 出力ファイルはcp932

binmode STDIN, ':encoding(cp932)'; # 標準入力はcp932
binmode STDOUT, ':encoding(cp932)'; # 標準出力はcp932
binmode STDERR, ':encoding(cp932)'; #エラー出力はcp932
#binmode STDOUT => ":utf8";

print "ticket2xls ver. 0.13.06.15.\n";
my $args = @ARGV;
if( $args < 4 ){
    warn "Usage: ticket2xls <site address> <rest key> <ticket no> <output xls file>\n";
    exit;
}

# ARGVは文字コードを変更してやらないとダメっぽい。
my $address   = decode('cp932', $ARGV[0] ); # site address
my $restApikey = decode('cp932', $ARGV[1] ); # rest api key
my $ticketNo = decode('cp932', $ARGV[2] ); # ticket no
my $outputFile   = decode('cp932', $ARGV[3] ); # output file

my @outArray = ();
#my $url = "$address\/issues\/$ticketNo\.xml?key=$restApikey&include=journals";
my $url = "$address\/issues\/$ticketNo\.xml?key=$restApikey";
my $xml = requestRestApi($url);
#print Dumper( $xml );
print "$xml->{description}\n";
exportExcel($xml,$outputFile);
exit;

sub requestRestApi
{
	my ($url) = @_;
	print "request[$url]\n";
	my $ua = LWP::UserAgent->new;
	$ua->timeout(1000);
	$ua->agent('TestPerl');
	my $req1 = HTTP::Request->new(GET => $url);
	my $res1 = $ua->request($req1);
	my $response = decode('utf8', $res1->as_string);
	if( $response !~ /HTTP\/1\.1 200 OK/g ){
		die "Request ERROR[$url]\n";
	}

	my $xml_string =$response;
	$xml_string =~ s/\n//g;
	if( $xml_string =~ /^.+?(<\?xml.+)$/ ){
		$xml_string = $1;
	}
	my $parser = XML::Simple->new;
	#my $xml = $parser->XMLin( encode('utf8',$xml_string) );
	my $xml = $parser->XMLin( $xml_string );
	return $xml;
}

#
# 'priority' => { 'name' => "\x{901a}\x{5e38}, 'id' => '2' },
# 'tracker' => {'name' => "\x{30bf}\x{30b9}\, 'id' => '4'},
# 'subject' => 'test',
# 'status' => { 'name' => "\x{65b0}\x{898f}", 'id' => '1' },
# 'project' => {'name' => "\x{798f}\x{4e95}\,'id' => '1' },
# 'author' => { 'name' => "02 \x{5927}\x{5bb6,'id' => '3'},
# 'description' => "\@[test]\@
# 'updated_on' => '2013-06-15T14:19:48Z',
# 'due_date' => {},
# 'created_on' => '2013-06-15T13:54:39Z',
# 'estimated_hours' => {},
# 'done_ratio' => '0',
# 'start_date' => '2013-06-15',
# 'id' => '20',
# 'spent_hours' => '0.0'


sub getItems
{
	my ($xml, $items) = @_;
	my @itemNameList = ();
	my @array = split /\n/, $xml->{'description'};
	my $itemName = undef;
	foreach my $line (@array){
		if( $line =~ /^\s*\@\[(.+?)\]\@$/ ){
			$itemName = $1;
			$items->{$itemName} = '';
			push @itemNameList, $itemName;
		}
		elsif( defined $itemName){
			$items->{$itemName} = $items->{$itemName}.$line;
		}
	}
	return @itemNameList;
}

sub exportExcel
{
	my ($xml,$file) = @_;
	print "exportExcel[$file]\n";
	
	
	my $file_sjis = encode('cp932', $file);
	# Create a new Excel workbook
	my $workbook = Spreadsheet::WriteExcel->new($file_sjis);
	# Add a worksheet
	my $worksheet = $workbook->add_worksheet();
	#  Add and define a format
	my $format = $workbook->add_format(); # Add a format
	my $format2 = $workbook->add_format(); # Add a format
	$format2->set_bg_color('silver');
 	$format2->set_bold();
 	$format2->set_align('left');
 	$format2->set_align('top');
	$format2->set_text_wrap();

	my %items = ();
	my @itemNameList = getItems($xml, \%items);

	# titles
	for(my $x=0;$x<@itemNameList;$x++){
		$worksheet->write( 0, $x, $itemNameList[$x], $format2);
	}
	$worksheet->set_column( 0, eval(@itemNameList), 20); # 幅設定
	
	# data
	for(my $x=0;$x<@itemNameList;$x++){
		$worksheet->write( 1, $x, $items{ $itemNameList[$x] }, $format);
	}
	$workbook->close;
}
