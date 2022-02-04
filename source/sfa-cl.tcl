# This is the main routine for the STEP File Analyzer and Viewer command-line version

# Website - https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer
# NIST Disclaimer - https://www.nist.gov/disclaimer
# Source code - https://github.com/usnistgov/SFA

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

puts "\n[string repeat "-" 87]"
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
  if {[string first "couldn't load library" $emsg] != -1} {
    append emsg "\n\nAlthough the message above indicates that a library is missing, that is NOT the cause of the problem.  The problem is sometimes related to the directory where you are running the software.\n\n   [file nativename $dir]"
    append emsg "\n\n1 - The directory has accented, non-English, or symbol characters"
    append emsg "\n2 - The directory is on a different computer"
    append emsg "\n3 - No permissions to run the software in the directory"
    append emsg "\n4 - Other computer configuration problems"
    append emsg "\n\nTry these workarounds to run the software:"
    append emsg "\n1 - From a directory without any special characters in the pathname, or from your home directory, or desktop"
    append emsg "\n2 - Installed on your local computer"
    append emsg "\n3 - As Administrator"
    append emsg "\n4 - On a different computer"
  }
  puts "\nError: $emsg"
  exit
}

catch {
  lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/vfs1.4.2
  package require vfs::zip
}

# text color
catch {set clTextColor [lindex [twapi::get_console_screen_buffer_info stdout -textattr] 1]}

# no arguments, no file, print help, and exit
set helpText "\nUsage: sfa-cl.exe myfile.stp \[csv\] \{\[view\]|\[stats\]|\[syntax\]\} \[noopen\] \[nolog\] \[file\]

Optional command line settings:
  csv     Generate CSV files
  view    Only run the Viewer and do not generate a Spreadsheet
  stats   Only report characteristics of the STEP file, no output files are generated
  syntax  Only run the Syntax Checker
  noopen  Do not open the Spreadsheet or Viewer file after it has been generated
  nolog   Do not generate a Log file
  file    Name of custom options file, e.g., C:/mydir/myoptions.dat  This file should
          be similar to STEP-File-Analyzer-options.dat in your home directory.

 Most options last used in the GUI version are used in this program unless the 'file'
 option is used.  If 'myfile.stp' has spaces, put double quotes around the file name
 \"C:/my dir/my file.stp\"

 You should run the GUI version of the software first.  If not already installed, the
 IFCsvr toolkit will be installed the first time this software is run.

 When the STEP file is opened, errors and warnings might appear in the output between
 the 'Begin ST-Developer output' and 'End ST-Developer output' messages.  Existing
 Spreadsheets and Viewer files are always overwritten.

Disclaimers
 NIST Disclaimer: https://www.nist.gov/disclaimer

 This software uses IFCsvr, Microsoft Excel, and software based on Open CASCADE that
 are covered by their own Software License Agreements.  If you are using this software
 in your own application, please explicitly acknowledge NIST as the source of the
 software.

Credits
- Reading and parsing STEP files:
   IFCsvr ActiveX Component, Copyright \u00A9 1999, 2005 SECOM Co., Ltd. All Rights Reserved
   IFCsvr has been modified by NIST to include STEP schemas.
   The license agreement can be found in C:\\Program Files (x86)\\IFCsvrR300\\doc
- Translating STEP to X3D:
   Developed by Soonjo Kwon, former NIST Associate
   https://www.nist.gov/services-resources/software/step-x3d-translator
- Some Tcl code is based on: CAWT http://www.cawt.tcl3d.org/"

if {$argc == 1} {set arg [string tolower [lindex $argv 0]]}
if {$argc == 0 || ($argc == 1 && ($arg == "help" || $arg == "-help" || $arg == "-h" || $arg == "-v"))} {
  puts $helpText
  exit
}

# -----------------------------------------------------------------------------------------------------
# set program files, environment variables will be in the correct language
set pf32 "C:\\Program Files (x86)"
if {[info exists env(ProgramFiles)]} {set pf32 $env(ProgramFiles)}
set pf64 ""
if {[info exists env(ProgramW6432)]} {set pf64 $env(ProgramW6432)}

# set drive, myhome, mydocs, mydesk
setHomeDir

# NIST version
set nistVersion 1

# get STEP file name
set localName [lindex $argv 0]
if {[string first ":" $localName] == -1} {set localName [file join [pwd] $localName]}
set localName [file nativename $localName]

# check for zipped file
set opt(logFile) 0
if {[string first ".stpz" [string tolower $localName]] != -1} {unzipFile}

if {![file exists $localName]} {
  errorMsg "STEP file not found: [truncFileName $localName]"
  exit
}

# -----------------------------------------------------------------------------------------------------
# initialize variables, set opt to 1
foreach id {logFile outputOpen PMIGRF PMISEM stepADDM stepAP242 stepCOMM stepCOMP stepCONS stepFEAT stepKINE stepPRES stepQUAL stepQUAN \
  stepREPR stepSHAP stepTOLR valProp} {set opt($id) 1}

# set opt to 0
foreach id { \
  feaBounds feaDisp feaDispNoTail feaLoads feaLoadScale indentGeometry indentStyledItem INVERSE partEdges partNormals partOnly partSketch \
  PMIGRFCOV PMISEMDIM PMISEMDT PMISEMRND SHOWALLPMI stepCPNT stepGEOM stepUSER syntaxChecker tessPartMesh viewFEA viewPart viewPMI viewPMIAR viewPMIVP \
  viewTessPart writeDirType xlHideLinks xlNoRound xlSort xlUnicode x3dSave DEBUG1 DEBUGINV DEBUGNOXL DEBUGVP DEBUGX3D \
} {set opt($id) 0}

set opt(gpmiColor) 3
set opt(partQuality) 7
set opt(xlMaxRows) 1003
set opt(xlFormat) Excel

set coverageSTEP 0
set dispCmd ""
set dispCmds {}
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
  if {$arg != "" && $arg1 != "csv" && [string first "vi" $arg1] == -1 && [string first "noo" $arg1] == -1 && [string first "sta" $arg1] == -1 && [string first "nol" $arg1] == -1 && [string first "syn" $arg1] == -1} {
    if {[file exists $arg]} {
      set customFile [file nativename $arg]
      puts "Using custom options file: [truncFileName $customFile]"
      set optionsFile $customFile
    } else {
      errorMsg "Unexpected command-line argument: $arg"
    }
  }
}

# check for options file and read (source)
if {[file exists $optionsFile]} {
  if {[catch {
    source $optionsFile
    puts "Reading options file: [truncFileName $optionsFile]"

# rename and unset old opt variables
    foreach pair [list {HIDELINKS xlHideLinks} {LOGFILE logFile} {SYNCHK syntaxChecker} {VALPROP valProp} {VIZBRP VIZPRT} {VIZFEA viewFEA} {VIZFEABC feaBounds} \
      {VIZFEADS feaDisp} {VIZFEADSntail feaDispNoTail} {VIZFEALV feaLoads} {VIZFEALVS feaLoadScale} {VIZPMI viewPMI} {VIZPRT viewPart} {VIZPRTEDGE partEdges} \
      {VIZPRTNORMAL partNormals} {VIZPRTONLY partOnly} {VIZPRTWIRE partSketch} {VIZTES viewTessPart} {VIZTESMSH tessPartMesh} {VIZTPG viewTessPart} \
      {VIZTPGMSH tessPartMesh} {x3dQuality partQuality} {XL_FPREC xlNoRound} {XL_OPEN outputOpen} {XL_ROWLIM xlMaxRows} {XL_SORT xlSort} {XLSCSV xlFormat} \
    ] {
      set old [lindex $pair 0]
      set new [lindex $pair 1]
      if {[info exists opt($old)]} {set opt($new) $opt($old); unset opt($old)}
    }
    foreach id [array names opt] {foreach str {EX_ PR_ XL_ VIZ} {if {[string first $str $id] == 0} {unset opt($id)}}}
  } emsg]} {
    errorMsg "Error reading options file [truncFileName $optionsFile]: $emsg"
  }
} else {
  errorMsg "No options file was found.  Default options will be used."
}

# adjust some variables
if {[info exists userEntityFile]} {
  if {![file exists $userEntityFile]} {
    set userEntityFile ""
    set opt(stepUSER) 0
  }
}
if {$opt(partQuality) == 9} {set opt(partQuality) 10}

#-------------------------------------------------------------------------------
# install IFCsvr
installIFCsvr 1

# -----------------------------------------------------------------------------------------------------
# get command line options
for {set i 1} {$i <= 10} {incr i} {
  set arg [string tolower [lindex $argv $i]]
  if {$arg != ""} {
# noopen
    if {[string first "noo" $arg] == 0} {set opt(outputOpen) 0}
# csv
    if {[string first "csv" $arg] == 0 && [lsearch [string tolower $argv] "vi"] == -1} {set opt(xlFormat) "CSV"}
# view
    if {[string first "vi" $arg] == 0} {
      set opt(xlFormat) "None"
      set gen(Excel) 0
      set gen(CSV) 0
      set gen(View) 1
      set allNone -1
      foreach id {feaBounds feaDisp feaLoads viewFEA viewPart viewPMI viewTessPart} {set opt($id) 1}
      foreach id {feaDispNoTail feaLoadScale PMIGRF PMISEM tessPartMesh valProp} {set opt($id) 0}
      checkValues
    }
# stats
    if {[string first "sta" $arg] == 0} {set statsOnly 1}
# nolog
    if {[string first "nol" $arg] == 0} {set opt(logFile) 0}
# syntax, run syntax checker and exit
    if {[string first "syn" $arg] == 0} {syntaxChecker $localName; exit}
  }
}

# update version in options file
if {$sfaVersion < [getVersion]} {
  set sfaVersion [getVersion]
  saveState
}

# -----------------------------------------------------------------------------------------------------
# generate spreadsheet or CSV files
genExcel
