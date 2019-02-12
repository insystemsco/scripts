#!/usr/bin/perl
# Script Name:			
#         auto_rip
# 	
# Version History:
#         version history moved to auto_rip-changelog.txt
#
# Script Description:
#         auto_rip automates the execution of RegRipper according to an examination process. Select plug-ins are ran to
#		  create reports showing types of information. The possible reports are:
#			- General Information about the Operating System and Its Configuration
#			- User Account Information
#			- Installed Software Information
#			- Networking Configuration Information
#			- Storage Information
#           - Device Information
#			- Program Execution Information
#			- Autostart Locations Information
#			- Logging Information
#			- Malware Indicators
#			- Web Browsing Information
#			- User Account Configuration Information
#			- User Account General Activity
#			- User Account Network Activity
#			- User Account File/Folder Access Activity
#			- User Account Virtualization Access Activity
#			- Communication Software Information	
#
#  Copyright (C) 2013 Corey Harrell (Journey Into Incident Response - http://journeyintoir.blogspot.com/)
#
# This software is released via the GNU General Public License:
# http://www.gnu.org/licenses/gpl.html
#
use strict;
use Getopt::Long;
use FileHandle;

# ------------------------------------------------------------------------------
# Initialization Section
# ------------------------------------------------------------------------------

# Define variables, arrays, and hashes here
my $version = "auto_rip v2013.08.06";
my $optionsresult;  # traps if GetOptions is successful or not
my $configdir;		# folder containing the SAM, Security, Software, and System hives
my $amcachedir;     # folder containing the Amcache hive
my $ntuserdir;		# folder containing the NTUSER.DAT hive
my $usrclassdir;	# folder containing the UsrClass.dat hive
my @categories;		# array to store the selected categories
my @all = qw(os users software network storage device execution autoruns log malware web user_config user_act user_network user_file user_virtual comm); 
my $reportdir = "auto_rip-reports";  #variable to store the directory name
my $logfile;        # variable to store the logfile path
my $logfile_han;    # variable for logfile filehandle
my $sam_var;            #variable to store registry hive name
my $software_var;       #variable to store registry hive name
my $system_var;         #variable to store registry hive name
my $security_var;       #variable to store registry hive name
my $amcache_var;        #variable to store registry hive name
my $ntuser_var;         #variable to store registry hive name
my $usrclass_var;       #variable to store registry hive name
my $file;				#variable to store file name (SIFT wouldn't work when I only used $_)

# Show help if no command-line options used
if (!@ARGV) {&HelpMessage();}  # needs to be placed before GetOptions since that call sets @ARGV to 1

# Define and process command-line options
$optionsresult = GetOptions ('h|help|?' => sub { &HelpMessage() },
					's|system:s' => \$configdir,
					'a|amcache:s' => \$amcachedir,
					'n|ntuser:s' => \$ntuserdir,
					'u|usrclass:s' => \$usrclassdir,
					'c|cat:s' => \@categories,
					'r|reportdir:s'=> \$reportdir);
 
# Show help if an unknown option is used 
if (!$optionsresult) {
	print ("\n");
	&HelpMessage;}

# Configuring the category array
if (@categories) { @categories = split(/,/,join(',',@categories))};  # removing the commas from the array
if (!@categories) {@categories = @all};  # default for categories is all if nothing is specified
foreach (@categories) {if ($_ eq "all") {@categories = @all};} #populates array with all categories when "all" is used

# ------------------------------------------------------------------------------
# MAIN Processing Section
# ------------------------------------------------------------------------------

# Formatting the directory paths for processing to avoid any potential errors

# Removing quotes from directory paths if present
if ($configdir) {$configdir =~ s!"!!g};
if ($amcachedir) {$amcachedir =~ s!"!!g};
if ($ntuserdir) {$ntuserdir =~ s!"!!g};
if ($usrclassdir) {$usrclassdir =~ s!"!!};
if ($reportdir) {$reportdir =~ s!"!!};

# Adding the trailing slash if not present based on the operating system
if ($^O eq "MSWin32") { #detecting Windows systems
	if ($configdir) {$configdir=~ s!\\*$!\\!};
	if ($amcachedir) {$amcachedir=~ s!\\*$!\\!};
	if ($ntuserdir) {$ntuserdir=~ s!\\*$!\\!};
	if ($usrclassdir) {$usrclassdir=~ s!\\*$!\\!};
	}
if ($^O ne "MSWin32") { #detecting non-Windows systems
	if ($configdir) {$configdir=~ s!/*$!/!};
	if ($amcachedir) {$amcachedir=~ s!/*$!/!};
	if ($ntuserdir) {$ntuserdir=~ s!/*$!/!};
	if ($usrclassdir) {$usrclassdir=~ s!/*$!/!};
	}

# Storing the registry hive names into variables (avoids issue with case sensitivity in Linux)
if ($configdir) { 
	opendir (DIR, $configdir) or die $!;
	while ($file = readdir(DIR)) {
		if ($file =~ m/^sam$/i) {$sam_var = $file}
		elsif ($file =~ m/^security$/i) {$security_var = $file}
		elsif ($file =~ m/^software$/i) {$software_var = $file}
		elsif ($file =~ m/^system$/i) {$system_var = $file}
	}
	closedir(DIR);
}
if ($amcachedir) { 
	opendir (DIR, $amcachedir) or die $!;
	while ($file = readdir(DIR)) {
		if ($file =~ m/^amcache\.hve$/i) {$amcache_var = $file}
	}
	closedir(DIR);
}
if ($ntuserdir) { 
	opendir (DIR, $ntuserdir) or die $!;
	while ($file = readdir(DIR)) {
		if ($file =~ m/^ntuser\.dat$/i) {$ntuser_var = $file}
	}
	closedir(DIR);
}
if ($usrclassdir) { 
	opendir (DIR, $usrclassdir) or die $!;
	while ($file = readdir(DIR)) {
		if ($file =~ m/^usrclass\.dat$/i) {$usrclass_var = $file}
	}
	closedir(DIR);
}
	
#Creating the report directory
if (mkdir $reportdir) {};	
if ($^O eq "MSWin32") { #detecting Windows systems to add slash
	if ($reportdir) {$reportdir=~ s!\\*$!\\!};
	}

if ($^O ne "MSWin32") { #detecting non-Windows systems to add slash
	if ($reportdir) {$reportdir=~ s!/*$!/!};
	}

# Setting up timestamp for the logfile
my ($d,$m,$y) = (localtime) [3,4,5]; #storing the time a variable for logging
my ($day,$month,$year) = ($d,$m+1,$y+1900);
my $timestamp = "$year-$month-$day:";

#Creating the logfile
$logfile = $reportdir . "auto_rip-logfile.txt"; #storing the log filepath in a variable
$logfile_han = FileHandle->new(">$logfile");
if ($configdir) {print $logfile_han ("$timestamp The folder containing the SAM, Security, Software, and System hives is $configdir\n")};
if ($amcachedir) {print $logfile_han ("$timestamp The folder containing the AMcache.hve hive is $amcachedir\n")};
if ($ntuserdir) {print $logfile_han ("$timestamp The folder containing the NTUSER.DAT hive is $ntuserdir\n")};
if ($usrclassdir) {print $logfile_han ("$timestamp The folder containing the UsrClass.dat hive is $usrclassdir\n")};
if (@categories) {print $logfile_han ("$timestamp The categories selected are: @categories\n")};
close $logfile_han;

# Main loop to process through the categories in the array
foreach (@categories) {
	if ($_ eq "os") {&os()};
	if ($_ eq "users") {&users()};
	if ($_ eq "software") {&software()};
	if ($_ eq "network") {&network()};
	if ($_ eq "storage") {&storage()};
    if ($_ eq "device") {&device()};
	if ($_ eq "execution") {&execution()};
	if ($_ eq "autoruns") {&autoruns()};
	if ($_ eq "log") {&log()};
	if ($_ eq "malware") {&malware()};
	if ($_ eq "web") {&web()};
	if ($_ eq "user_config") {&user_config()};
	if ($_ eq "user_act") {&user_act()};
	if ($_ eq "user_network") {&user_network()};
	if ($_ eq "user_file") {&user_file()};
	if ($_ eq "user_virtual") {&user_virtual()};
	if ($_ eq "comm") {&comm()};
	}



# ------------------------------------------------------------------------------
# Subroutine Section
# ------------------------------------------------------------------------------

sub HelpMessage 	{
	print<< "EOT";

$version

auto_rip [-s path] [-n path] [-u path] [-r report-directory] [-c categories]

-h, --help      lists all of the available options
-s, --system    path to the folder containing the SAM, Security, Software, and System hives
-a, --amcache   path to the folder containing the AMcache.hve hive (only in Windows 8)
-n, --ntuser    path to the folder containing the NTUSER.DAT hive
-u, --usrclass  path to the folder containing the UsrClass.dat hive
-r, --reportdir path to the folder to store the output reports
-c, --cat       specifies the plug-in categories to run. Seperate multiple categories with a comma

	Supported Categories:

		all             gets information from all categories
		os              gets General Operating System Information
		users           gets User Account Information
		software        gets Installed Software Information
		network         gets Networking Configuration Information
		storage         gets Storage Information
		device          gets Device Information
		execution       gets Program Execution Information
		autoruns        gets Autostart Locations Information
		log             gets Logging Information
		malware         gets Malware Indicators
		web             gets Web Browsing Information
		user_config     gets User Account Configuration Information
		user_act        gets User Account General Activity
		user_network    gets User Account Network Activity
		user_file       gets User Account File/Folder Access Activity
		user_virtual    gets User Account Virtualization Access Activity
		comm            Communication Software Information

Usage:

Extract all information from the SAM, Security, Software, and System hives. 
C:\\>auto_rip -s H:\\Windows\\System32\\config -c all

Extract file and network access information from NTUSER.DAT hive (Windows XP user profile)
C:\\>auto_rip -n "H:\\Documents and Settings\\Corey" -c user_network,user_file

Extract file access information from NTUSER.DAT and UsrClass.dat hive (Windows 7 profile)
C:\\>auto_rip -n H:\\Users\\Corey -u H:\\Users\\Corey\\AppData\\Local\\Microsoft\\Windows -c user_file

Extract all information from all Windows 7 registry hives without using -c switch.
C:\\>auto_rip -s H:\\Windows\\System32\\config -n H:\\Users\\Corey -u H:\\Users\\Corey\\AppData\\Local\\Microsoft\\Windows

Extract all information from all Windows 8 registry hives without using -c switch.
C:\\>auto_rip -s H:\\Windows\\System32\\config -a H:\\Windows\\AppCompat\\Programs\ -n H:\\Users\\Corey -u H:\\Users\\Corey\\AppData\\Local\\Microsoft\\Windows

Extract all information from the SAM, Security, Software, and System hives then store output reports in a specified directory. 
C:\\>auto_rip -s H:\\Windows\\System32\\config -r C:\\reports

Copyright 2015  Corey Harrell (jIIr)
EOT
exit;
}

sub os	{
	print ("---- Processing the os category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the os category\n");
	# build filename for report 
	my $osreport = $reportdir . "01_operating_system_information.txt"; #filename for the os report
	my $osreport_han = FileHandle->new(">$osreport");
	#formatting the report
	print $osreport_han ("=========================================================================================================\n");
	print $osreport_han ("General Information about the Operating System and Its Configuration\n");
	print $osreport_han ("=========================================================================================================\n");
	print $osreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $osreport_han;
	close $logfile_han;
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$software" -p winnt_cv >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p producttype >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p win_cv >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p timezone >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p shutdown >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p shutdowncount >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$security" -p polacdms >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p winlogon >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p uac >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p disablesr >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p diag_sr >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p spp_clients >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p backuprestore >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p winbackup >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p bitbucket >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p disablelastaccess >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p dfrg >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p secctr >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p pagefile >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p hibernate >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p processor_architecture >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p crashcontrol >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p regback >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p ctrlpnl >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p banner >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$system" -p nolmhash >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		system (qq{perl rip.pl -r "$software" -p susclient >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
        system (qq{perl rip.pl -r "$software" -p gpohist >>"$osreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$osreport"});
		system (qq{echo .>>"$osreport"});
		}
}
sub users	{
	print ("---- Processing the users category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the users category\n");
	# build filename for report 
	my $usersreport = $reportdir . "02_user_account_information.txt"; #filename for the users report
	my $usersreport_han = FileHandle->new(">$usersreport");
	#formatting the report
	print $usersreport_han ("=========================================================================================================\n");
	print $usersreport_han ("User Account Information\n");
	print $usersreport_han ("=========================================================================================================\n");
	print $usersreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $usersreport_han;
	close $logfile_han;
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$sam" -p samparse >>"$usersreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$usersreport"});
		system (qq{echo .>>"$usersreport"});
		system (qq{perl rip.pl -r "$software" -p profilelist >>"$usersreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$usersreport"});
		system (qq{echo .>>"$usersreport"});
	}
}
sub software	{
	print ("---- Processing the software category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the software category\n");
	# build filename for report 
	my $softwarereport = $reportdir . "03_installed_software_information.txt"; #filename for the software report
	my $softwarereport_han = FileHandle->new(">$softwarereport");
	#formatting the report
	print $softwarereport_han ("=========================================================================================================\n");
	print $softwarereport_han ("Installed Software Information\n");
	print $softwarereport_han ("=========================================================================================================\n");
	print $softwarereport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $softwarereport_han;
	close $logfile_han;
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$software" -p uninstall >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
		system (qq{perl rip.pl -r "$software" -p apppaths >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
		system (qq{perl rip.pl -r "$software" -p assoc >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
		system (qq{perl rip.pl -r "$software" -p installedcomp >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
		system (qq{perl rip.pl -r "$software" -p msis >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
		system (qq{perl rip.pl -r "$software" -p product >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
		system (qq{perl rip.pl -r "$software" -p installer >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
		system (qq{perl rip.pl -r "$software" -p clsid >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
	}
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p listsoft >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
		system (qq{perl rip.pl -r "$ntuser" -p fileexts >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
		system (qq{perl rip.pl -r "$ntuser" -p arpcache >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
		system (qq{perl rip.pl -r "$ntuser" -p startpage >>"$softwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$softwarereport"});
		system (qq{echo .>>"$softwarereport"});
	}
}
sub network	{
	print ("---- Processing the network category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the network category\n");
	# build filename for report 
	my $networkreport = $reportdir . "04_network_configuration_information.txt"; #filename for the network report
	my $networkreport_han = FileHandle->new(">$networkreport");
	#formatting the report
	print $networkreport_han ("=========================================================================================================\n");
	print $networkreport_han ("Networking Configuration Information\n");
	print $networkreport_han ("=========================================================================================================\n");
	print $networkreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $networkreport_han;
	close $logfile_han;
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$system" -p compname >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$software" -p networkcards >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$system" -p nic >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$system" -p nic2 >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$software" -p macaddr >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$system" -p shares >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$system" -p fw_config >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$system" -p routes >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$software" -p networklist >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$software" -p ssid >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$software" -p networkuid >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$system" -p network >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$system" -p termserv >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$system" -p termcert >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$system" -p rdpport >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
		system (qq{perl rip.pl -r "$software" -p sql_lastconnect >>"$networkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$networkreport"});
		system (qq{echo .>>"$networkreport"});
	}
}
sub storage	{
	print ("---- Processing the storage category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the storage category\n");
	# build filename for report 
	my $storagereport = $reportdir . "05_storage_information.txt"; #filename for the storage report
	my $storagereport_han = FileHandle->new(">$storagereport");
	#formatting the report
	print $storagereport_han ("=========================================================================================================\n");
	print $storagereport_han ("Storage Information\n");
	print $storagereport_han ("=========================================================================================================\n");
	print $storagereport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $storagereport_han;
	close $logfile_han;
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$system" -p mountdev2 >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$system" -p ide >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$system" -p usbdevices >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$system" -p usbstor >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$system" -p devclass >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$software" -p emdmgmt >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$system" -p wpdbusenum >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$system" -p bthport >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$system" -p btconfig >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$system" -p imagedev >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$system" -p stillimage >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
	}
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p mp2 >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$ntuser" -p mndmru >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$ntuser" -p knowndev >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
		system (qq{perl rip.pl -r "$ntuser" -p ddo >>"$storagereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$storagereport"});
		system (qq{echo .>>"$storagereport"});
	}
}
sub device	{
	print ("---- Processing the device category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the device category\n");
	# build filename for report 
	my $devicereport = $reportdir . "06_device_information.txt"; #filename for the device report
	my $devicereport_han = FileHandle->new(">$devicereport");
	#formatting the report
	print $devicereport_han ("=========================================================================================================\n");
	print $devicereport_han ("Device Information\n");
	print $devicereport_han ("=========================================================================================================\n");
	print $devicereport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $devicereport_han;
	close $logfile_han;
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$software" -p audiodev >>"$devicereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$devicereport"});
		system (qq{echo .>>"$devicereport"});
	}
}
sub execution	{
	print ("---- Processing the execution category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the execution category\n");
	# build filename for report 
	my $executionreport = $reportdir . "07_program_execution_information.txt"; #filename for the execution report
	my $executionreport_han = FileHandle->new(">$executionreport");
	#formatting the report
	print $executionreport_han ("=========================================================================================================\n");
	print $executionreport_han ("Program Execution Information\n");
	print $executionreport_han ("=========================================================================================================\n");
	print $executionreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $executionreport_han;
	close $logfile_han;
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$system" -p prefetch >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
		system (qq{perl rip.pl -r "$system" -p appcompatcache >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
		system (qq{perl rip.pl -r "$system" -p legacy >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
		system (qq{perl rip.pl -r "$software" -p tracing >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
        system (qq{perl rip.pl -r "$software" -p at >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
		system (qq{perl rip.pl -r "$software" -p direct >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
	}
	if ($amcachedir) { 
		#setup registry variables
		my $amcache = $amcachedir . $amcache_var;
		#running select plugins
		system (qq{perl rip.pl -r "$amcache" -p amcache >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
	}
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p muicache >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
	}
	if ($usrclassdir) { 
		#setup registry variables
		my $usrclass = $usrclassdir . $usrclass_var;
		#running select plugins
		system (qq{perl rip.pl -r "$usrclass" -p muicache >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
	}
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p userassist >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
		system (qq{perl rip.pl -r "$ntuser" -p appcompatflags >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
		system (qq{perl rip.pl -r "$ntuser" -p winscp >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
        system (qq{perl rip.pl -r "$ntuser" -p mixer >>"$executionreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$executionreport"});
		system (qq{echo .>>"$executionreport"});
	}
}
sub autoruns	{
	print ("---- Processing the autoruns category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the autoruns category\n");
	# build filename for report 
	my $autorunsreport = $reportdir . "08_autoruns_information.txt"; #filename for the autoruns report
	my $autorunsreport_han = FileHandle->new(">$autorunsreport");
	#formatting the report
	print $autorunsreport_han ("=========================================================================================================\n");
	print $autorunsreport_han ("Autostart Locations Information\n");
	print $autorunsreport_han ("=========================================================================================================\n");
	print $autorunsreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $autorunsreport_han;
	close $logfile_han;
	# the first autoruns listed are the Run keys since one of the most common locations for malware
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$software" -p soft_run >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
	}
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p user_run >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
	}
	# the second autoruns listed are the services since one of the most common locations for malware
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$system" -p services >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$system" -p svc >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$system" -p svcdll >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
	}
	# Remaining autoruns locations
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select Software plugins
		system (qq{perl rip.pl -r "$software" -p appinitdlls >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$software" -p init_dlls >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$software" -p bho >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$software" -p installedcomp >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$software" -p imagefile >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$software" -p winlogon >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$software" -p svchost >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$software" -p drivers32 >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$software" -p cmd_shell >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$software" -p shellexec >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$software" -p shellext >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$software" -p schedagent >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		#running select System plugins
		system (qq{perl rip.pl -r "$system" -p appcertdlls >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$system" -p lsa_packages >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$system" -p safeboot >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$system" -p dllsearch >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$system" -p securityproviders >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
        system (qq{perl rip.pl -r "$system" -p profiler >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
	}
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running NTUSER select plugins
		system (qq{perl rip.pl -r "$ntuser" -p load >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$ntuser" -p winlogon_u >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$ntuser" -p cmdproc >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
		system (qq{perl rip.pl -r "$ntuser" -p startup >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
        system (qq{perl rip.pl -r "$ntuser" -p cached >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
        system (qq{perl rip.pl -r "$ntuser" -p profiler >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
	}
	if ($usrclassdir) { 
		#setup registry variables
		my $usrclass = $usrclassdir . $usrclass_var;
		#running NTUSER select plugins
		system (qq{perl rip.pl -r "$usrclass" -p cmd_shell_u >>"$autorunsreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$autorunsreport"});
		system (qq{echo .>>"$autorunsreport"});
	}
}
sub log	{
	print ("---- Processing the log category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the log category\n");
	# build filename for report 
	my $logreport = $reportdir . "09_log_information.txt"; #filename for the log report
	my $logreport_han = FileHandle->new(">$logreport");
	#formatting the report
	print $logreport_han ("=========================================================================================================\n");
	print $logreport_han ("Logging Information\n");
	print $logreport_han ("=========================================================================================================\n");
	print $logreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $logreport_han;
	close $logfile_han;
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$software" -p mrt >>"$logreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$logreport"});
		system (qq{echo .>>"$logreport"});
		system (qq{perl rip.pl -r "$security" -p auditpol >>"$logreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$logreport"});
		system (qq{echo .>>"$logreport"});
		system (qq{perl rip.pl -r "$system" -p eventlog >>"$logreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$logreport"});
		system (qq{echo .>>"$logreport"});
		system (qq{perl rip.pl -r "$system" -p eventlogs >>"$logreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$logreport"});
		system (qq{echo .>>"$logreport"});
		system (qq{perl rip.pl -r "$software" -p winevt >>"$logreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$logreport"});
		system (qq{echo .>>"$logreport"});
		system (qq{perl rip.pl -r "$system" -p auditfail >>"$logreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$logreport"});
		system (qq{echo .>>"$logreport"});
		system (qq{perl rip.pl -r "$software" -p drwatson >>"$logreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$logreport"});
		system (qq{echo .>>"$logreport"});
	}
}
sub malware	{
	print ("---- Processing the malware category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the malware category\n");
	# build filename for report 
	my $malwarereport = $reportdir . "10_malware_indicators.txt"; #filename for the malware report
	my $malwarereport_han = FileHandle->new(">$malwarereport");
	#formatting the report
	print $malwarereport_han ("=========================================================================================================\n");
	print $malwarereport_han ("Malware Indicators\n");
	print $malwarereport_han ("=========================================================================================================\n");
	print $malwarereport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $malwarereport_han;
	close $logfile_han;
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$system" -p pending >>"$malwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$malwarereport"});
		system (qq{echo .>>"$malwarereport"});
		system (qq{perl rip.pl -r "$system" -p netsvcs >>"$malwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$malwarereport"});
		system (qq{echo .>>"$malwarereport"});
		system (qq{perl rip.pl -r "$software" -p inprocserver >>"$malwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$malwarereport"});
		system (qq{echo .>>"$malwarereport"});
        system (qq{perl rip.pl -r "$software" -p fileless >>"$malwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$malwarereport"});
		system (qq{echo .>>"$malwarereport"});
	}
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p cpldontload >>"$malwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$malwarereport"});
		system (qq{echo .>>"$malwarereport"});
        system (qq{perl rip.pl -r "$ntuser" -p fileless >>"$malwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$malwarereport"});
		system (qq{echo .>>"$malwarereport"});
        system (qq{perl rip.pl -r "$ntuser" -p inprocserver >>"$malwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$malwarereport"});
		system (qq{echo .>>"$malwarereport"});
	}
	if ($usrclassdir) { 
		#setup registry variables
		my $usrclass = $usrclassdir . $usrclass_var;
		#running select plugins
		system (qq{perl rip.pl -r "$usrclass" -p inprocserver >>"$malwarereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$malwarereport"});
		system (qq{echo .>>"$malwarereport"});
	}
}
sub web	{
	print ("---- Processing the web category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the web category\n");
	# build filename for report 
	my $webreport = $reportdir . "11_web-browsing_information.txt"; #filename for the web report
	my $webreport_han = FileHandle->new(">$webreport");
	#formatting the report
	print $webreport_han ("=========================================================================================================\n");
	print $webreport_han ("Web Browsing Information\n");
	print $webreport_han ("=========================================================================================================\n");
	print $webreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $webreport_han;
	close $logfile_han;
	if ($configdir) { 
		#setup registry variables
		my $software = $configdir . $software_var;
		my $system = $configdir . $system_var;
		my $security = $configdir . $security_var;
		my $sam = $configdir . $sam_var;
		#running select plugins
		system (qq{perl rip.pl -r "$software" -p defbrowser >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		system (qq{perl rip.pl -r "$software" -p startmenuinternetapps_lm >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		# IE related information starts here
		system (qq{echo ++++++++++  Internet Explorer Related  ++++++++++>>"$webreport"});
		system (qq{echo +>>"$webreport"});
		system (qq{perl rip.pl -r "$software" -p ie_version >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		system (qq{perl rip.pl -r "$software" -p ie_zones >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
        system (qq{perl rip.pl -r "$software" -p javasoft >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		# IE related information ends here
	}
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p startmenuinternetapps_cu >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		system (qq{perl rip.pl -r "$ntuser" -p proxysettings >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		system (qq{perl rip.pl -r "$ntuser" -p menuorder >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		# IE related information starts here
		system (qq{echo ++++++++++  Internet Explorer Related  ++++++++++>>"$webreport"});
		system (qq{echo +>>"$webreport"});
		system (qq{perl rip.pl -r "$ntuser" -p ie_settings >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		system (qq{perl rip.pl -r "$ntuser" -p internet_settings_cu >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		system (qq{perl rip.pl -r "$ntuser" -p internet_explorer_cu >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		system (qq{perl rip.pl -r "$ntuser" -p domains >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		system (qq{perl rip.pl -r "$ntuser" -p typedurls >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
        system (qq{perl rip.pl -r "$ntuser" -p ie_zones >>"$webreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$webreport"});
		system (qq{echo .>>"$webreport"});
		# IE related information ends here
	}
}	
sub user_config	{
	print ("---- Processing the user_config category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the user_config category\n");
	# build filename for report 
	my $userconfigreport = $reportdir . "12_user-configuration_information.txt"; #filename for the user_config report
	my $userconfigreport_han = FileHandle->new(">$userconfigreport");
	#formatting the report
	print $userconfigreport_han ("=========================================================================================================\n");
	print $userconfigreport_han ("User Account Configuration Information\n");
	print $userconfigreport_han ("=========================================================================================================\n");
	print $userconfigreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $userconfigreport_han;
	close $logfile_han;	
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p shellfolders >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p policies_u >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p environment >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p userinfo >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p vista_bitbucket >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p autorun >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p attachmgr >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p autoendtasks >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p winlogon_u >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p user_win >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
        system (qq{perl rip.pl -r "$ntuser" -p gpohist >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		# User Software Related Information starts here
		system (qq{echo ++++++++++  User Software Related Information  ++++++++++>>"$userconfigreport"});
		system (qq{echo +>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p ccleaner >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p sysinternals >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		# User Software Related Information ends here
		# User Network Settings Information starts here
		system (qq{echo ++++++++++  User Network Settings Information  ++++++++++>>"$userconfigreport"});
		system (qq{echo +>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p logonusername >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p ntusernetwork >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p printers >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p winvnc >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		system (qq{perl rip.pl -r "$ntuser" -p userlocsvc >>"$userconfigreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userconfigreport"});
		system (qq{echo .>>"$userconfigreport"});
		# User Network Settings Information ends here
	}
}
sub user_act	{
	print ("---- Processing the user_act category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the user_act category\n");
	# build filename for report 
	my $useractreport = $reportdir . "13_user-account-general-activity.txt"; #filename for the user_act report
	my $useractreport_han = FileHandle->new(">$useractreport");
	#formatting the report
	print $useractreport_han ("=========================================================================================================\n");
	print $useractreport_han ("User Account General Activity\n");
	print $useractreport_han ("=========================================================================================================\n");
	print $useractreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $useractreport_han;
	close $logfile_han;	
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p typedpaths >>"$useractreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$useractreport"});
		system (qq{echo .>>"$useractreport"});
		system (qq{perl rip.pl -r "$ntuser" -p mmc >>"$useractreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$useractreport"});
		system (qq{echo .>>"$useractreport"});
		system (qq{perl rip.pl -r "$ntuser" -p runmru >>"$useractreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$useractreport"});
		system (qq{echo .>>"$useractreport"});
		system (qq{perl rip.pl -r "$ntuser" -p applets >>"$useractreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$useractreport"});
		system (qq{echo .>>"$useractreport"});
		system (qq{perl rip.pl -r "$ntuser" -p acmru >>"$useractreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$useractreport"});
		system (qq{echo .>>"$useractreport"});
		system (qq{perl rip.pl -r "$ntuser" -p wordwheelquery >>"$useractreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$useractreport"});
		system (qq{echo .>>"$useractreport"});
		system (qq{perl rip.pl -r "$ntuser" -p cdstaginginfo >>"$useractreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$useractreport"});
		system (qq{echo .>>"$useractreport"});
		system (qq{perl rip.pl -r "$ntuser" -p gthist >>"$useractreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$useractreport"});
		system (qq{echo .>>"$useractreport"});
	}
}
sub user_network	{
	print ("---- Processing the user_network category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the user_network category\n");
	# build filename for report 
	my $usernetworkreport = $reportdir . "14_user-account-network-activity.txt"; #filename for the user_network report
	my $usernetworkreport_han = FileHandle->new(">$usernetworkreport");
	#formatting the report
	print $usernetworkreport_han ("=========================================================================================================\n");
	print $usernetworkreport_han ("User Account Network Activity\n");
	print $usernetworkreport_han ("=========================================================================================================\n");
	print $usernetworkreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $usernetworkreport_han;
	close $logfile_han;	
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p mndmru >>"$usernetworkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$usernetworkreport"});
		system (qq{echo .>>"$usernetworkreport"});
		system (qq{perl rip.pl -r "$ntuser" -p compdesc >>"$usernetworkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$usernetworkreport"});
		system (qq{echo .>>"$usernetworkreport"});
		system (qq{perl rip.pl -r "$ntuser" -p tsclient >>"$usernetworkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$usernetworkreport"});
		system (qq{echo .>>"$usernetworkreport"});
		system (qq{perl rip.pl -r "$ntuser" -p rdphint >>"$usernetworkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$usernetworkreport"});
		system (qq{echo .>>"$usernetworkreport"});
		system (qq{perl rip.pl -r "$ntuser" -p ssh_host_keys >>"$usernetworkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$usernetworkreport"});
		system (qq{echo .>>"$usernetworkreport"});
		system (qq{perl rip.pl -r "$ntuser" -p winscp_sessions >>"$usernetworkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$usernetworkreport"});
		system (qq{echo .>>"$usernetworkreport"});
		system (qq{perl rip.pl -r "$ntuser" -p vncviewer >>"$usernetworkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$usernetworkreport"});
		system (qq{echo .>>"$usernetworkreport"});
		system (qq{perl rip.pl -r "$ntuser" -p vnchooksapplicationprefs >>"$usernetworkreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$usernetworkreport"});
		system (qq{echo .>>"$usernetworkreport"});
	}
}	
sub user_file	{
	print ("---- Processing the user_file category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the user_file category\n");
	# build filename for report 
	my $userfilereport = $reportdir . "15_user-account-file-access-activity.txt"; #filename for the user_file report
	my $userfilereport_han = FileHandle->new(">$userfilereport");
	#formatting the report
	print $userfilereport_han ("=========================================================================================================\n");
	print $userfilereport_han ("User Account File/Folder Access Activity\n");
	print $userfilereport_han ("=========================================================================================================\n");
	print $userfilereport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $userfilereport_han;
	close $logfile_han;	
	if ($usrclassdir) { 
		#setup registry variables
		my $usrclass = $usrclassdir . $usrclass_var;
		#running select plugins
		system (qq{perl rip.pl -r "$usrclass" -p shellbags >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
	}
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p shellbags >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p itempos >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p comdlg32 >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p recentdocs >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p winzip >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p winrar >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p sevenzip >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p mspaper >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p nero >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		# Microsoft Office Related starts
		system (qq{echo ++++++++++  Microsoft Office Files Accessed  ++++++++++>>"$userfilereport"});
		system (qq{echo +>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p officedocs >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p officedocs2010 >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p reading_locations >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p oisc >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p trustrecords >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p snapshot_viewer >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		# Microsoft Office Related ends
		# Adobe Related starts
		system (qq{echo ++++++++++  Adobe Files Accessed  ++++++++++>>"$userfilereport"});
		system (qq{echo +>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p adoberdr >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		# Adobe Related ends
		# Multimedia Related starts
		system (qq{echo ++++++++++  Multimedia Files Accessed  ++++++++++>>"$userfilereport"});
		system (qq{echo +>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p wallpaper >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p mpmru >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		system (qq{perl rip.pl -r "$ntuser" -p realplayer6 >>"$userfilereport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$userfilereport"});
		system (qq{echo .>>"$userfilereport"});
		# Multimedia Related ends
	}
}
sub user_virtual	{
	print ("---- Processing the user_virtual category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the user_virtual category\n");
	# build filename for report 
	my $uservirtualreport = $reportdir . "16_user-account-virtual-access.txt"; #filename for the user_virtual report
	my $uservirtualreport_han = FileHandle->new(">$uservirtualreport");
	#formatting the report
	print $uservirtualreport_han ("=========================================================================================================\n");
	print $uservirtualreport_han ("User Account Virtualization Access Activity\n");
	print $uservirtualreport_han ("=========================================================================================================\n");
	print $uservirtualreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $uservirtualreport_han;
	close $logfile_han;	
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		system (qq{perl rip.pl -r "$ntuser" -p vmplayer >>"$uservirtualreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$uservirtualreport"});
		system (qq{echo .>>"$uservirtualreport"});
		system (qq{perl rip.pl -r "$ntuser" -p vmware_vsphere_client >>"$uservirtualreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$uservirtualreport"});
		system (qq{echo .>>"$uservirtualreport"});
	}
}
sub comm	{
	print ("---- Processing the comm category\n");
	$logfile_han = FileHandle->new(">>$logfile");
	print $logfile_han ("---- Processing the comm category\n");
	# build filename for report 
	my $commreport = $reportdir . "17_communications_information.txt"; #filename for the comm report
	my $commreport_han = FileHandle->new(">$commreport");
	#formatting the report
	print $commreport_han ("=========================================================================================================\n");
	print $commreport_han ("Communication Software Information\n");
	print $commreport_han ("=========================================================================================================\n");
	print $commreport_han ("\n");
	# closing handles since redirection with system will error out if they are open
	close $commreport_han;
	close $logfile_han;	
	if ($ntuserdir) { 
		#setup registry variables
		my $ntuser = $ntuserdir . $ntuser_var;
		#running select plugins
		# Email related information starts
		system (qq{echo ++++++++++  Email Communication Information  ++++++++++>>"$commreport"});
		system (qq{echo +>>"$commreport"});
		system (qq{perl rip.pl -r "$ntuser" -p outlook >>"$commreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$commreport"});
		system (qq{echo .>>"$commreport"});
		system (qq{perl rip.pl -r "$ntuser" -p olsearch >>"$commreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$commreport"});
		system (qq{echo .>>"$commreport"});
		system (qq{perl rip.pl -r "$ntuser" -p unreadmail >>"$commreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$commreport"});
		system (qq{echo .>>"$commreport"});
		# Email related information ends
		# Telecommunications related information starts
		system (qq{echo ++++++++++  Telecommunications Information  ++++++++++>>"$commreport"});
		system (qq{echo +>>"$commreport"});
		system (qq{perl rip.pl -r "$ntuser" -p skype >>"$commreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$commreport"});
		system (qq{echo .>>"$commreport"});
		# Telecommunications related information ends
		# Messaging related information starts
		system (qq{echo ++++++++++  Messaging Communication Information  ++++++++++>>"$commreport"});
		system (qq{echo +>>"$commreport"});
		system (qq{perl rip.pl -r "$ntuser" -p aim >>"$commreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$commreport"});
		system (qq{echo .>>"$commreport"});
		system (qq{perl rip.pl -r "$ntuser" -p liveContactsGUID >>"$commreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$commreport"});
		system (qq{echo .>>"$commreport"});
		system (qq{perl rip.pl -r "$ntuser" -p yahoo_cu >>"$commreport" 2>>"$logfile"});
		system (qq{echo ...........................................................................................................>>"$commreport"});
		system (qq{echo .>>"$commreport"});
		# Messaging related information ends
	}
}	
