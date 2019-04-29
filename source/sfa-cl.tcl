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
# The STEP File Analyzer and Viewer can only be built with Tcl 8.5.15 or earlier
# More recent versions are incompatibile with the IFCsvr toolkit that is used to read STEP files
# ----------------------------------------------------------------------------------------------
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
puts "NIST STEP File Analyzer and Viewer (v[getVersion] - Updated: [string trim [clock format $progtime -format "%e %b %Y"]])"

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
set helpText "\nUsage: sfa-cl.exe myfile.stp \[csv\] \[viz\] \[noopen\] \[file\]

Optional command line settings:
  csv     Generate CSV files
  viz     Only generate Views and no Spreadsheet or CSV files
  stats   Only report characteristics of the STEP file
  noopen  Do not open the Spreadsheet or View after it has been generated
  file    Name of custom options file, e.g., C:/mydir/myoptions.dat
          This file should be similar to STEP-File-Analyzer-options.dat in your home directory.

If 'myfile.stp' has spaces, put quotes around the file name \"C:/mydir/my file.stp\"

It is recommended to run the GUI version of the software first.  If not already
installed, the IFCsvr toolkit (used to read STEP files) and STEP schema files will
be installed the first time this software is run.

Disclaimers:

This software was developed at the National Institute of Standards and Technology
by employees of the Federal Government in the course of their official duties.
Pursuant to Title 17 Section 105 of the United States Code this software is not
subject to copyright protection and is in the public domain.  This software is an
experimental system.  NIST assumes no responsibility whatsoever for its use by
other parties, and makes no guarantees, expressed or implied, about its quality,
reliability, or any other characteristic.

This software uses Microsoft Excel and IFCsvr that are covered by their own
End-User License Agreements.  The B-rep part geometry viewer is based on software
from OpenCascade and pythonOCC.

See the NIST Disclaimer at: https://www.nist.gov/disclaimer"

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

# check for IFCsvr toolkit
set sfaType "CL"
set ifcsvrDir [file join $pf32 IFCsvrR300 dll]
if {![file exists [file join $ifcsvrDir IFCsvrR300.dll]]} {installIFCsvr; exit} 

# -----------------------------------------------------------------------------------------------------
# initialize variables, set opt to 1
foreach id { \
  DISPGUIDE1 FIRSTTIME LOGFILE PMIGRF PMISEM \
  PR_STEP_AP242 PR_STEP_COMM PR_STEP_COMP PR_STEP_FEAT PR_STEP_KINE \
  PR_STEP_PRES PR_STEP_QUAN PR_STEP_REPR PR_STEP_SHAP PR_STEP_TOLR \
  VALPROP VIZFEABC VIZFEADS VIZFEALV \
  XL_LINK1 XL_OPEN \
} {set opt($id) 1}

# set opt to 0
foreach id { \
  CRASH DEBUG1 DEBUGINV indentGeomtry indentStyledItem INVERSE \
  PR_STEP_CPNT PR_STEP_GEOM PR_USER VIZBRP VIZFEA VIZFEADSntail \
  VIZFEALVS VIZPMI VIZPMIVP VIZTPG VIZTPGMSH \
  writeDirType XL_FPREC XL_KEEPOPEN XL_SORT \
} {set opt($id) 0}

set opt(gpmiColor) 3
set opt(XL_ROWLIM) 1003
set opt(XLSBUG1) 30
set opt(XLSCSV) Excel

set coverageSTEP 0
set dispCmd ""
set dispCmds {}
set firsttime 1
set excelVersion 12
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
if {$env(USERDOMAIN) == "NIST"} {set developer 1}

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
  if {$arg != "" && $arg1 != "csv" && $arg1 != "viz" && $arg1 != "noopen" && $arg1 != "stats"} {
    if {[file exists $arg]} {
      set customFile [file nativename $arg]
      puts "\n*** Using custom options file: [truncFileName $customFile]"
      append endMsg "\nA custom options file was used: [truncFileName $customFile]"
      set optionsFile $customFile
    } else {
      puts "\n*** Bad command-line argument: $arg"
      append endMsg "Bad command-line argument: $arg"
    }
  }
}

# check for options file and read (source)
if {[file exists $optionsFile]} {
  if {[catch {
    source $optionsFile
    puts "Reading options file: [truncFileName $optionsFile]"
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
    if {[string first "csv" $arg] == 0} {
      if {[lsearch [string tolower $argv] "viz"] == -1} {set opt(XLSCSV) "CSV"}
    }                              
    if {[string first "viz" $arg] == 0} {
      set opt(XLSCSV) "None"
      foreach id {VIZBRP VIZFEA VIZFEABC VIZFEADS VIZFEALV VIZPMI VIZTPG} {set opt($id) 1}
      foreach id {PMIGRF PMISEM VALPROP VIZFEADSntail VIZFEALVS VIZPMIVP VIZTPGMSH} {set opt($id) 0}
    }
    if {[string first "sta" $arg] == 0} {set statsOnly 1}
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

# set process id used to check memory usage for AP209 files
set sfaPID [twapi::get_process_ids -name "sfa-cl.exe"]

# -----------------------------------------------------------------------------------------------------
# generate spreadsheet or CSV files
genExcel

# repeat options file messages
if {[info exists endMsg]} {
  puts "\n*** Options file messages"
  puts $endMsg
}
