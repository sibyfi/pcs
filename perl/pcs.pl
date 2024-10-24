#!/usr/bin/env perl 
#===============================================================================
#
#         FILE: pcs.pl
#
#        USAGE: ./pcs.pl <filename> 
#
#  DESCRIPTION: Script to re-format the "Upcoming Class Schedule.csv" report.
#
#      OPTIONS: ---
# REQUIREMENTS: ---
#         BUGS: Probably
#        NOTES: Changed the deep link format for "ENROLL"
#       AUTHOR: George Rucker, NetApp
# ORGANIZATION: NetAppU
#      VERSION: 8
#      CREATED: 09/15/2016 08:37:06 AM
#     REVISION: 2
#===============================================================================

use strict;
use warnings;
use utf8;

use Text::ParseWords;
use Time::Piece;
use Excel::Writer::XLSX;
use Getopt::Std;

our($opt_i,$opt_o);
getopts ("o:i:");

# if ($#ARGV != 3) {
# 	Usage();
# }

########
# vars #
########
# my $outputdir = "output";
# my $input = $ARGV[0];
my $time = Tdate();
# https://netapp.sabacloud.com/Saba/Web_spf/NA1PRD0047/common/leclassview/virtc-00361210
my $link_base = "https://netapp.sabacloud.com/Saba/Web_spf/NA1PRD0047/common/leclassview/";
#my $link = "http://learningcenter.netapp.com/LC?ObjectType=ILT&ObjectID=";
my $more_offers_pre = "https://netapp.sabacloud.com/Saba/Web_spf/NA1PRD0047/common/ledetail/";
my @fields;

if (!defined($opt_o) || !defined($opt_i) ) {
    Usage();
}

########
# help #
########
sub Usage {
	print "Usage: $0 -i <input file> -o <output file>\n";
	exit 0;
}

###################
# parse each line #
###################
sub ParseInput {
	@fields = parse_line(',', 0, $_);
	return;
}

###########
# fix url #
###########
sub FixUrl {
	# <a href=http://www.flane.de/netapp>Enroll</a>	
	my $newurl = $_ =~ m/\=(.*)\>/;
	return;
}

########################
# date for output file #
########################
sub Tdate {
	my @time      = localtime;
	my $year      = ($time[5]) + 1900;
	my $mon       = $time[4];
	my $day       = $time[3];
	my $hour      = $time[2];
	my $min       = $time[1];
	my $sec       = $time[0];
	my @abbr = qw(Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec);
	my $timestamp = join('',$year,$abbr[$mon],$day,"_",$hour,$min,$sec);
	return $timestamp;
}

##############################
# open file and create array #
##############################
open("FH1","<$opt_i") or die "Could not open file $!\n";
my @elements = <FH1>;
shift(@elements);
chomp(@elements);
my $elements_len = @elements;

#################
# init workbook #
#################
my $outputfile = "Class_Schedule" ."_" . $time . ".xlsx";
my $workbook = Excel::Writer::XLSX->new("$opt_o/$outputfile");
my $worksheet = $workbook->add_worksheet();
$worksheet->freeze_panes(1, 0);
$worksheet->autofilter('A1:S1');
$worksheet->set_row(0, 30);

####################
# excel formatting #
####################
my $format_header = $workbook->add_format (
	border       => 1,
	border_color => "#0067C5",
	valign       => 'vcenter',
	align        => 'center',
	bold         => 1,
	bg_color     => "#0067C5",
	color        => "#FFFFFF",
	text_wrap    => 1,
);

my $format1 = $workbook->add_format (
	border       => '1',
	border_color => "#0067C5",
	valign       => 'vcenter',
	align        => 'center',
);

my $format2 = $workbook->add_format (
	border       => '1',
	border_color => "#0067C5",
	valign       => 'vcenter',
	#align       => 'center',
	text_wrap    => '1',
);

my $format_link = $workbook->add_format (
	border       => '1',
	border_color => "#0067C5",
	valign       => 'vcenter',
	align        => 'center',
	underline    => '1',
	color        => 'blue',
);

my $format_red = $workbook->add_format (
	border       => '1',
	border_color => "#0067C5",
	valign       => 'vcenter',
	align        => 'center',
	bg_color     => 'red',
	color        => 'white',
	bold         => '1',
);

my $format_yellow = $workbook->add_format (
	border       => '1',
	border_color => "#0067C5",
	valign       => 'vcenter',
	align        => 'center',
	bg_color     => 'yellow',
	color        => 'black',
	bold         => '1',
);

my $format_green = $workbook->add_format (
	border       => '1',
	border_color => "#0067C5",
	valign       => 'vcenter',
	align        => 'center',
	bg_color     => 'green',
	color        => 'white',
	bold         => '1',
);

##########################
# set custom format type #
##########################
my $format_num = $workbook->add_format (
	border       => '1',
	border_color => "#0067C5",
	valign       => 'vcenter',
	align        => 'center',
	num_format   => '0000000000',
);

################
# write header #
################
#$worksheet->set_row(0,undef,$format_header);
my @header = ("Course Name",
	"Course Number",
	"Start Date",
	"End Date",
	"Duration (Days)",
	"Days Remaining",
	"Location",
	"GEO",
	"Max Count",
	"Student Count",
	"Open Seats",
	"Service Rep",
	"Offering ID",
	"More Offerings",
	"Link",
	"Catalog Domain Name",
	"Offering Domain",
	"Instructor",
	"Display for Learner");

my $col = 0;
foreach (@header) {
	chomp;
	$worksheet->write(0,$col,$_,$format_header);
	$col++;
}

#######################################
# parse file and populate spreadsheet #
#######################################
my $row = 1;
my $line_num = 1;
foreach (@elements) {
	chomp;
	s/What\'s/What\\'s/g;
	&ParseInput;
	(my $course_num,
	my $course_name,
	my $cat_dom_name,
	my $offering_dom,
	my $offering_num,
	my $offering_start_date,
	my $offering_end_date,
	my $offering_loc,
	my $offering_reg,
	my $offering_instr,
	my $disp_for_learner,
	my $enroll_link,
	my $curr_enrolled,
	my $max_stud_count,
	my $cust_serv_rep,
	my $class_type,
	my $domain,
	my $offering_status) = @fields;

	my $perc_full;
	my $min_size = 6;
	my $lms = "1" if ($enroll_link =~ /learningcenter|sabacloud/i);

	my $open_seats = $max_stud_count - $curr_enrolled;
	my $after = Time::Piece->strptime("$offering_end_date", "%Y-%m-%d");
	my $now = localtime;
	my $days_remaining = int(($after - $now) / 86400);
	if ($curr_enrolled == $max_stud_count ) {
		$perc_full = 100;
	} else {
		$perc_full = int(($curr_enrolled / $max_stud_count) * 100);
	}
	my $more_offers = $more_offers_pre . $course_num;

	my $edate = Time::Piece->strptime("$offering_end_date", "%Y-%m-%d");
	my $sdate = Time::Piece->strptime("$offering_start_date", "%Y-%m-%d");
	my $duration = (($edate - $sdate)/86400) + 1;
	my $link;

	print "(" . $row . ") " . $offering_num . "\n";

	$col = 0;
	while($col < '15') {
		$worksheet->write($row,0,$course_name,$format2);
		$worksheet->write($row,1,$course_num,$format2);
		$worksheet->write($row,2,$offering_start_date,$format1);
		$worksheet->write($row,3,$offering_end_date,$format1);
		$worksheet->write($row,4,$duration,$format1);
		$worksheet->write($row,5,$days_remaining,$format1);
		if ($offering_loc =~ /(Virtual\sClass)/) {
			my ($loc,$gmt) = ($offering_loc =~ m/(.*)(\(.*)/);
			$worksheet->write($row,6,$loc,$format1);
		} else {
			$worksheet->write($row,6,$offering_loc,$format1);
		}
		$offering_reg = "AMER" if ($offering_reg =~ /AMERICAS|Americas/);
		$worksheet->write($row,7,$offering_reg,$format1);
		$worksheet->write($row,8,$max_stud_count,$format1);

		# ############
		# Student Count coloring
		# ############
		# if (($curr_enrolled >= $min_size) || ($curr_enrolled < $min_size && $days_remaining >= '30')) {
		$worksheet->write($row,9,$curr_enrolled,$format1);
		# } else {
		# 	$worksheet->write($row,9,$curr_enrolled,$format_yellow);
		# }

		##############
		# Open Seats coloring
		# ############
		# class full
		if ($curr_enrolled == $max_stud_count) {
			$worksheet->write($row,10,$open_seats,$format_green);
		# Less than 6 and less than 14 days
		} elsif ($curr_enrolled < $min_size && $days_remaining <= '14') {
			$worksheet->write($row,10,$open_seats,$format_red);
		# Less than 6 and less than 30 days
		} elsif ($curr_enrolled < $min_size && $days_remaining <= '30') {
			$worksheet->write($row,10,$open_seats,$format_yellow);
		} else {
			$worksheet->write($row,10,$open_seats,$format1);
		}

		$worksheet->write($row,11,$cust_serv_rep,$format2);
		$worksheet->write($row,12,$offering_num,$format_num);
		if (defined($lms)) {
			if ($offering_loc =~ /Virtual/i) {
				$link = "${link_base}virtc-${offering_num}";
			} else {
				$link = "${link_base}class-${offering_num}";
			}
			$worksheet->write_url($row,13,"${more_offers}",$format_link,'More Offerings');
			$worksheet->write_url($row,14,"${link}",$format_link,'ENROLL');
		} else {
			$worksheet->write_url($row,13,"${more_offers}",$format_link,'More Offerings');
			$worksheet->write($row,14,"FixUrl($enroll_link)",$format_link,'ENROLL');
		}
		$worksheet->write($row,15,$offering_dom,$format1);
		$worksheet->write($row,16,$offering_dom,$format1);
		$worksheet->write($row,17,$offering_instr,$format1);
		$worksheet->write($row,18,$disp_for_learner,$format1);
		$col++;
	}
	$row++;
	$line_num++;
}
$worksheet->write(1,20,"Class is full",$format_green);
$worksheet->write(2,20,"Less than 6 Students and Less than 30 days",$format_yellow);
$worksheet->write(3,20,"Less than 6 Students and Less than 2 weeks",$format_red);

#####################
# set column widths #
#####################
$worksheet->set_column('A:A',30);
$worksheet->set_column('B:B',28);
$worksheet->set_column('C:D',10);
$worksheet->set_column('E:E',10);
$worksheet->set_column('F:F',12);
$worksheet->set_column('G:G',14);
$worksheet->set_column('H:H',12);
$worksheet->set_column('I:I',8);
$worksheet->set_column('J:J',10);
$worksheet->set_column('K:K',8);
$worksheet->set_column('L:L',20);
$worksheet->set_column('M:M',12);
$worksheet->set_column('N:N',15);
$worksheet->set_column('O:O',8);
$worksheet->set_column('P:P',12);
$worksheet->set_column('Q:Q',12);
$worksheet->set_column('R:R',12);
$worksheet->set_column('S:S',12);
$worksheet->set_column('U:U',45);

##################
# summary output #
##################
print "\nFilename: $outputfile\n";
print "Complete!\n";



