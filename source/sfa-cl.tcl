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

# This is the main routine for the STEP File Analyzer and Viewer command-line version

global env

set scriptName [info script]
set wdir [file dirname $scriptName]
set auto_path [linsert $auto_path 0 $wdir]
set contact [getContact]

#-------------------------------------------------------------------------------
# start 
set progtime 0
foreach fname [glob -nocomplain -directory $wdir *.tcl] {
  set mtime [file mtime $fname]
  if {$mtime > $progtime} {set progtime $mtime}
}

puts "\n--------------------------------------------------------------------------------"
puts "NIST STEP File Analyzer and Viewer [getVersion] (Updated [string trim [clock format $progtime -format "%e %b %Y"]])"

# for building your own version with freewrap, uncomment and modify C:/Tcl/lib/teapot directory if necessary
# the lappend commands add package locations to auto_path, must be before package commands below
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
  puts "\nERROR: $emsg\n\nThere might be a problem running this program from a directory with accented, non-English, or symbol characters in the pathname."
  puts "     [file nativename $dir]\nTry running the software from a directory without any of the special characters in the pathname."
  puts "\nContact [lindex $contact 0] ([lindex $contact 1]) if you cannot run the STEP File Analyzer and Viewer."
  exit
}

catch {
  lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/vfs1.4.2
  package require vfs::zip
}

# no arguments, no file, print help, and exit
set helpText "\nUsage: sfa-cl.exe myfile.stp \[csv\] \[viz\] \[stats\] \[noopen\] \[file\]

Optional command line settings:
  csv     Generate CSV files
  viz     Only generate Views and no Spreadsheet or CSV files
  stats   Only report characteristics of the STEP file, no output files are generated
  noopen  Do not open the Spreadsheet or View after it has been generated
  nolog   Do not generate a Log file
  file    Name of custom options file, e.g., C:/mydir/myoptions.dat  This file should
          be similar to STEP-File-Analyzer-options.dat in your home directory.

 Most options last used in the GUI version are used in this program unless the 'file'
 option is used.

 If 'myfile.stp' has spaces, put quotes around the file name \"C:/mydir/my file.stp\"

 You should run the GUI version of the software first.  If not already installed, the
 IFCsvr toolkit used to read STEP files will be installed the first time this software
 is run.
 
 Existing Spreadsheets and View files are always overwritten.  An Internet connection
 is required to show View files in the web browser.

 When the STEP file is opened, errors and warnings might appear in the output between
 the 'Begin ST-Developer output' and 'End ST-Developer output' messages.

Disclaimers
 This software was developed at the National Institute of Standards and Technology by
 employees of the Federal Government in the course of their official duties.  Pursuant
 to Title 17 Section 105 of the United States Code this software is not subject to
 copyright protection and is in the public domain.  This software is an experimental
 system.  NIST assumes no responsibility whatsoever for its use by other parties, and
 makes no guarantees, expressed or implied, about its quality, reliability, or any
 other characteristic.  NIST Disclaimer: https://www.nist.gov/disclaimer
 
Credits
- Generating spreadsheets:        Microsoft Excel (https://products.office.com/excel)
- Reading and parsing STEP files: IFCsvr ActiveX Component, Copyright \u00A9 1999, 2005 SECOM Co., Ltd. All Rights Reserved
                                  IFCsvr has been modified by NIST to include STEP schemas.
                                  The license agreement can be found in C:\\Program Files (x86)\\IFCsvrR300\\doc
- Viewing part geometry:          Developed by Soonjo Kwon at NIST based on the Open CASCADE STEP Processor
                                  Open CASCADE License https://www.opencascade.com/content/licensing"

if {$argc == 1} {set arg [string tolower [lindex $argv 0]]}
if {$argc == 0 || ($argc == 1 && ($arg == "help" || $arg == "-help" || $arg == "-h" || $arg == "-v"))} {
  puts $helpText
  exit
}

# -----------------------------------------------------------------------------------------------------
# set drive, myhome, mydocs, mydesk
setHomeDir

# set program files, environment variables will be in the correct language
set pf32 "C:\\Program Files (x86)"
if {[info exists env(ProgramFiles)]} {set pf32 $env(ProgramFiles)}
set pf64 ""
if {[info exists env(ProgramW6432)]} {set pf64 $env(ProgramW6432)}

# detect if NIST version
set nistVersion 1
#foreach item $auto_path {if {[string first "sfa-cl" $item] != -1} {set nistVersion 1}}

# get STEP file name
set localName [lindex $argv 0]
if {[string first ":" $localName] == -1} {set localName [file join [pwd] $localName]}
set localName [file nativename $localName]

# check for zipped file
set opt(LOGFILE) 0
if {[string first ".stpz" [string tolower $localName]] != -1} {unzipFile}  

if {![file exists $localName]} {
  puts "\n*** STEP file not found: [truncFileName $localName]"
  puts $helpText
  exit
}

# -----------------------------------------------------------------------------------------------------
# initialize variables, set opt to 1
foreach id { \
  LOGFILE PMIGRF PMISEM PR_STEP_AP242 PR_STEP_COMM PR_STEP_COMP PR_STEP_FEAT PR_STEP_KINE PR_STEP_PRES PR_STEP_QUAN \
  PR_STEP_REPR PR_STEP_SHAP PR_STEP_TOLR VALPROP VIZFEABC VIZFEADS VIZFEALV VIZPRTEDGE VIZPRTWIRE XL_OPEN \
} {set opt($id) 1}

# set opt to 0
foreach id { \
  DEBUG1 DEBUGINV DEBUGX3D HIDELINKS indentGeometry indentStyledItem INVERSE PMIGRFCOV PMISEMDIM PR_STEP_CPNT PR_STEP_GEOM PR_USER \
  SHOWALLPMI SYNCHK VIZPRT VIZPRTONLY VIZFEA VIZFEADSntail VIZFEALVS VIZPMI VIZTPG VIZTPGMSH writeDirType XL_FPREC XL_SORT \
} {set opt($id) 0}

set opt(gpmiColor) 0
set opt(x3dQuality) 7
set opt(XL_ROWLIM) 1003
set opt(XLSCSV) Excel

set coverageSTEP 0
set dispCmd ""
set dispCmds {}
set excelVersion 1000
set filesProcessed 0
set lastX3DOM ""
set lastXLS  ""
set lastXLS1 ""
set openFileList {}
set sfaVersion 0
set upgrade 0
set x3dFileName ""
set x3dStartFile 1

set fileDir  $mydocs
set fileDir1 $mydocs
set userWriteDir $mydocs
set writeDir $userWriteDir

set developer 0
if {$env(USERDOMAIN) == "NIST" || $env(USERDOMAIN) == "Cassie"} {set developer 1}

# initialize other data
initData
initDataInverses
getOpenPrograms

# -----------------------------------------------------------------------------------------------------
# check for custom options file
set optionsFile [file nativename [file join $fileDir STEP-File-Analyzer-options.dat]]
set customFile ""
for {set i 1} {$i <= 10} {incr i} {
  set arg [lindex $argv $i]
  set arg1 [string tolower $arg]
  if {$arg != "" && $arg1 != "csv" && $arg1 != "viz" && $arg1 != "noopen" && $arg1 != "stats" && $arg1 != "nolog"} {
    if {[file exists $arg]} {
      set customFile [file nativename $arg]
      puts "Using custom options file: [truncFileName $customFile]"
      set optionsFile $customFile
    } else {
      puts "\n*** Bad command-line argument: $arg"
    }
  }
}

# check for options file and read (source)
if {[file exists $optionsFile]} {
  if {[catch {
    source $optionsFile
    puts "Reading options file: [truncFileName $optionsFile]"
  } emsg]} {
    puts "\n*** Error reading options file [truncFileName $optionsFile]: $emsg"
  }
} else {
  puts "\n*** No options file was found.  Default options will be used."
}

# adjust some variables
if {[info exists userEntityFile]} {
  if {![file exists $userEntityFile]} {
    set userEntityFile ""
    set opt(PR_USER) 0
  }
}

#-------------------------------------------------------------------------------
# install IFCsvr
set ifcsvrDir [file join $pf32 IFCsvrR300 dll]
installIFCsvr 1

# get command line options
for {set i 1} {$i <= 10} {incr i} {
  set arg [string tolower [lindex $argv $i]]
  if {$arg != ""} {
    if {[string first "noo" $arg] == 0} {set opt(XL_OPEN) 0}                              
    if {[string first "csv" $arg] == 0} {
      if {[lsearch [string tolower $argv] "viz"] == -1} {set opt(XLSCSV) "CSV"}
    }                              
    if {[string first "viz" $arg] == 0} {
      set opt(XLSCSV) "None"
      foreach id {VIZPRT VIZFEA VIZFEABC VIZFEADS VIZFEALV VIZPMI VIZTPG} {set opt($id) 1}
      foreach id {PMIGRF PMISEM VALPROP VIZFEADSntail VIZFEALVS VIZTPGMSH} {set opt($id) 0}
    }
    if {[string first "sta" $arg] == 0} {set statsOnly 1}
    if {[string first "nol" $arg] == 0} {set opt(LOGFILE) 0}
  }
}

# update version in options file
if {$sfaVersion < [getVersion]} {
  set sfaVersion [getVersion]
  saveState
}

# set process id used to check memory usage for AP209 files
set sfaPID [twapi::get_process_ids -name "sfa-cl.exe"]

# -----------------------------------------------------------------------------------------------------
# generate spreadsheet or CSV files
genExcel
