# This software was developed at the National Institute of Standards and Technology by employees of 
# the Federal Government in the course of their official duties.  Pursuant to Title 17 Section 105 
# of the United States Code this software is not subject to copyright protection and is in the 
# public domain. This software is an experimental system.  NIST assumes no responsibility whatsoever 
# for its use by other parties, and makes no guarantees, expressed or implied, about its quality, 
# reliability, or any other characteristic.  We would appreciate acknowledgement if the software is 
# used.
# 
# This software can be redistributed and/or modified freely provided that any derivative works bear 
# some notice that they are derived from it, and any modified versions bear some notice that they 
# have been modified.

# The latest version of the source code is available at: https://github.com/usnistgov/SFA

# ----------------------------------------------------------------------------------------------
# The STEP File Analyzer can only be built with Tcl 8.5.15 or earlier
# More recent versions are incompatibile with the IFCsvr toolkit that is used to read STEP files
# ----------------------------------------------------------------------------------------------
# This is the main routine for the STEP File Analyzer command-line version

global env

set scriptName [info script]
set wdir [file dirname $scriptName]
set auto_path [linsert $auto_path 0 $wdir]

#-------------------------------------------------------------------------------
# start 
set progtime 0
foreach fname [glob -nocomplain -directory $wdir *.tcl] {
  set mtime [file mtime $fname]
  if {$mtime > $progtime} {set progtime $mtime}
}

puts "\n--------------------------------------------------------------------------------"
puts "NIST STEP File Analyzer (v[getVersion] - Updated: [string trim [clock format $progtime -format "%e %b %Y"]])"

# for freeWrap the following lappend commands add package locations to auto_path, must be before package commands
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/tcom3.9
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/twapi3.0.32
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Tclx8.4

# Tcl packages, check if they will load
if {[catch {
  package require tcom
  package require twapi
  package require Tclx
} emsg]} {
  set dir $wdir
  set c1 [string first [file tail [info nameofexecutable]] $dir]
  if {$c1 != -1} {set dir [string range $dir 0 $c1-1]}
  puts "\nERROR: $emsg\n\nThere might be a problem running this program from a directory with accented, non-English, or symbol characters in the pathname.\n\n     [file nativename $dir]\n\nRun the software from a directory without any special characters in the pathname.\n\nPlease contact Robert Lipman (robert.lipman@nist.gov) for other problems."
  exit
}

catch {
  lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/vfs1.4.2
  package require vfs::zip
}

# no arguments, no file, print help, and exit
if {$argc == 1} {set arg [string tolower [lindex $argv 0]]}
if {$argc == 0 || ($argc == 1 && ($arg == "help" || $arg == "-help" || $arg == "-h" || $arg == "-v"))} {
  puts "\nUsage: sfa-cl.exe myfile.stp \[csv\] \[viz\] \[noopen\] \[file\]"
  puts "\n  If myfile.stp has spaces, put quotes around the file name, e.g., \"C:/mydir/my file.stp\"."
  puts "\nOptional command line settings:"
  puts "  csv     Generate CSV files"                                                                                        
  puts "  viz     Generate only Visualizations and no spreadsheet or CSV files"                                                                                        
  puts "  noopen  Do not open the spreadsheet after it has been generated"
  puts "  file    Name of custom options file, e.g., C:/mydir/myoptions.dat"
  puts "          This file should be similar to STEP-File-Analyzer-options.dat"
  puts "          in your home directory."
  puts "
If not already installed, the IFCsvr toolkit (used to read STEP files) and STEP
schema files will be installed the first time this program is run.

Disclaimers:

This software was developed at the National Institute of Standards and Technology
by employees of the Federal Government in the course of their official duties.  
Pursuant to Title 17 Section 105 of the United States Code this software is not
subject to copyright protection and is in the public domain.  This software is an
experimental system.  NIST assumes no responsibility whatsoever for its use by
other parties, and makes no guarantees, expressed or implied, about its quality,
reliability, or any other characteristic.

This software uses Microsoft Excel and IFCsvr which are covered by their own
EULAs (End-User License Agreements).

See the NIST Disclaimer at: https://www.nist.gov/disclaimer"
  exit
}

# -----------------------------------------------------------------------------------------------------
# set drive, myhome, mydocs, mydesk
setHomeDir

# set program files
set pf32 "C:\Program Files (x86)"
if {[info exists env(ProgramFiles)]}  {set pf32 $env(ProgramFiles)}
if {[string first "x86" $pf32] == -1} {append pf32 " (x86)"}
set pf64 "C:\Program Files"
if {[info exists env(ProgramW6432)]} {set pf64 $env(ProgramW6432)}

# detect if NIST version
set nistVersion 1
#foreach item $auto_path {if {[string first "sfa-cl" $item] != -1} {set nistVersion 1}}

# get STEP file name
set localName [lindex $argv 0]
if {[string first ":" $localName] == -1} {set localName [file join [pwd] $localName]}
set localName [file nativename $localName]
if {![file exists $localName]} {
  outputMsg "\n*** STEP file not found: [truncFileName $localName]"
  exit
}
set remoteName $localName

# check for IFCsvr toolkit
set sfaType "CL"
set ifcsvrDir [file join $pf32 IFCsvrR300 dll]
if {![file exists [file join $ifcsvrDir IFCsvrR300.dll]]} {installIFCsvr} 

# -----------------------------------------------------------------------------------------------------
# initialize variables
foreach id {XL_OPEN XL_KEEPOPEN XL_LINK1 XL_FPREC XL_SORT LOGFILE \
            VALPROP PMIGRF PMISEM VIZPMI VIZFEA VIZTES VIZPMIVP INVERSE DEBUG1 \
            PR_STEP_AP242 PR_USER PR_STEP_KINE PR_STEP_COMP PR_STEP_COMM PR_STEP_GEOM PR_STEP_QUAN \
            PR_STEP_FEAT PR_STEP_PRES PR_STEP_TOLR PR_STEP_REPR PR_STEP_CPNT PR_STEP_SHAP} {set opt($id) 1}

set opt(DEBUG1) 0
set opt(DEBUGINV) 0
set opt(FIRSTTIME) 1
set opt(gpmiColor) 2
set opt(INVERSE) 0
set opt(PR_STEP_CPNT) 0
set opt(PR_STEP_GEOM)  0
set opt(PR_USER) 0
set opt(VIZFEA) 0
set opt(VIZPMI) 0
set opt(VIZPMIVP) 0
set opt(writeDirType) 0
set opt(XL_KEEPOPEN) 0
set opt(XL_ROWLIM) 1048576
set opt(XL_SORT) 0
set opt(XLSBUG1) 30
set opt(XLSCSV) Excel

set coverageSTEP 0
set dispCmd ""
set dispCmds {}
set excelYear ""
set firsttime 1
set lastX3DOM ""
set lastXLS  ""
set lastXLS1 ""
set openFileList {}
set pointLimit 2
set sfaVersion 0
set upgrade 0
set userXLSFile ""
set x3dFileName ""
set x3dStartFile 1

set fileDir  $mydocs
set fileDir1 $mydocs
set userWriteDir $mydocs
set writeDir $userWriteDir

set developer 0
if {$env(USERNAME) == "lipman"} {set developer 1}

# initialize other data
initData
initDataInverses

# -----------------------------------------------------------------------------------------------------
# check for custom options file
set optionsFile [file nativename [file join $fileDir STEP-File-Analyzer-options.dat]]
set customFile ""
for {set i 1} {$i <= 10} {incr i} {
  set arg [lindex $argv $i]
  if {[file exists $arg]} {
    set customFile [file nativename $arg]
    puts "\n*** Using custom options file: [truncFileName $customFile]"
    append endMsg "\nA custom options file was used: [truncFileName $customFile]"
    set optionsFile $customFile
  }
}

# check for options file and read (source)
if {[file exists $optionsFile]} {
  if {[catch {
    source $optionsFile
  } emsg]} {
    set msg "\nError reading options file: [truncFileName $optionsFile]\n $emsg\nFix or delete the file."
    append endMsg $msg
    puts $msg
  }
} else {
  puts "\n*** No options file was found.  Default options will be used."
  append endMsg "\nNo options file was found.  Default options were used."
}

# adjust some variables
if {[info exists userEntityFile]} {
  if {![file exists $userEntityFile]} {
    set userEntityFile ""
    set opt(PR_USER) 0
  }
}
set opt(XL_KEEPOPEN) 0

# get command line options
for {set i 1} {$i <= 10} {incr i} {
  set arg [string tolower [lindex $argv $i]]
  if {$arg != ""} {
    if {[string first "noo" $arg] == 0} {set opt(XL_OPEN) 0}                              
    if {[string first "csv" $arg] == 0} {set opt(XLSCSV) "CSV"}                              
    if {[string first "viz" $arg] == 0} {set opt(XLSCSV) "None"}                              
  }
}

# copy schema rose files that are in the Tcl Virtual File System (VFS) or STEP Tools runtime to the IFCsvr dll directory
set copyrose 0
if {$opt(FIRSTTIME) || $sfaVersion < [getVersion]} {
  set opt(FIRSTTIME) 0
  set sfaVersion [getVersion]
  set copyrose 1
  saveState
}
if {$copyrose} {copyRoseFiles}

# -----------------------------------------------------------------------------------------------------
# generate spreadsheet or CSV files
genExcel

# repeat options file messages
if {[info exists endMsg]} {
  puts "\n*** Options file messages"
  puts $endMsg
}
