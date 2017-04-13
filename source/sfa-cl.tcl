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

# ----------------------------------------------------------------------------------------------
# The STEP File Analyzer can only be built with Tcl 8.5.15 or earlier
# More recent versions are incompatibile with the IFCsvr toolkit that is used to read STEP files
# ----------------------------------------------------------------------------------------------

# This is the main routine for the STEP File Analyzer command-line version

global env

set wdir [file dirname [info script]]
set auto_path [linsert $auto_path 0 $wdir]

# for freeWrap the following lappend commands add package locations to auto_path, must be before package commands
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/tcom3.9
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/twapi3.0.32
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Tclx8.4

# Tcl packages
package require tcom
package require twapi
package require Tclx

catch {
  lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/vfs1.4.2
  package require vfs::zip
}

set nistVersion 0
foreach item $auto_path {if {[string first "STEP-File-Analyzer" $item] != -1} {set nistVersion 1}}

foreach id {XL_OPEN XL_KEEPOPEN XL_LINK1 XL_FPREC XL_SORT \
            VALPROP PMIGRF PMISEM VIZPMI VIZFEA INVERSE DEBUG1 DEBUG2 \
            PR_STEP_AP242 PR_USER PR_STEP_KINE PR_STEP_COMP PR_STEP_COMM PR_STEP_GEOM PR_STEP_QUAN \
            PR_STEP_FEAT PR_STEP_PRES PR_STEP_TOLR PR_STEP_REPR PR_STEP_CPNT PR_STEP_SHAP} {set opt($id) 1}

set opt(DEBUG1) 0
set opt(DEBUG2) 0
set opt(DEBUGINV) 0
set opt(VIZPMI) 0
set opt(VIZFEA) 0
set opt(PR_STEP_CPNT) 0
set opt(PR_STEP_GEOM)  0
set opt(PR_USER) 0
set opt(XL_ROWLIM) 10000000
set opt(XL_SORT) 0
set opt(writeDirType) 0
set opt(XL_KEEPOPEN) 0
set opt(XLSCSV) Excel

set coverageSTEP 0
set dispCmd ""
set dispCmds {}
set excelYear ""
set firsttime 1
set lastXLS  ""
set lastXLS1 ""
set openFileList {}
set pointLimit 2
set sfaVersion 0
set upgrade 0
set userXLSFile ""
set x3domFileName ""
set x3domFileOpen 1

set developer 0
if {$env(USERNAME) == "lipman"} {set developer 1}

# -----------------------------------------------------------------------------------------------------
# set drive, myhome, mydocs, mydesk
setHomeDir

set fileDir  $mydocs
set fileDir1 $mydocs
set userWriteDir $mydocs
set writeDir $userWriteDir

# set program files
set programfiles "C:/Program Files"
set pf64 ""
if {[info exists env(ProgramFiles)]} {set programfiles $env(ProgramFiles)}
if {[info exists env(ProgramW6432)]} {set pf64 $env(ProgramW6432)}

# default installation directory for IFCsvr toolkit
set ifcsvrdir [file join $programfiles IFCsvrR300 dll]

# -----------------------------------------------------------------------------------------------------
# initialize data
initData
initDataInverses

# set options file name
set optionsFile1 [file nativename [file join $fileDir STEP_Excel_options.dat]]
set optionsFile2 [file nativename [file join $fileDir STEP-File-Analyzer-options.dat]]

if {(![file exists $optionsFile1] && ![file exists $optionsFile2]) || \
     [file exists $optionsFile2]} {
  set optionsFile $optionsFile2
} else {
  catch {
    file copy -force $optionsFile1 $optionsFile2
    file delete -force $optionsFile1
    set optionsFile $optionsFile2
  } optionserr
}

# check for options file and read
set optionserr ""
if {[file exists $optionsFile]} {
  catch {source $optionsFile} optionserr
  if {[string first "+" $optionserr] == 0} {set optionserr ""}
} else {
  puts "\n*** RUN THE GUI VERSION FIRST BEFORE RUNNING THE COMMAND-LINE VERSION ***"
}

# adjust some variables
if {[info exists userEntityFile]} {
  if {![file exists $userEntityFile]} {
    set userEntityFile ""
    set opt(PR_USER) 0
  }
}

#-------------------------------------------------------------------------------
# start 
set progtime 0
foreach item {sfa-cl sfa-data sfa-dimtol sfa-ent sfa-gen sfa-geotol sfa-grafpmi sfa-proc sfa-step sfa-valprop} {
  set fname [file join $wdir $item.tcl]
  set mtime [file mtime $fname]
  if {$mtime > $progtime} {set progtime $mtime}
}

puts "\n--------------------------------------------------------------------------------"
set str ""
if {$nistVersion} {set str "NIST "}
puts "$str\STEP File Analyzer (v[getVersion] - Updated: [string trim [clock format $progtime -format "%e %b %Y"]])"

#-------------------------------------------------------------------------------

# check for IFCsvr
if {![file exists [file join $programfiles IFCsvrR300 dll IFCsvrR300.dll]]} {
  puts "\n*** RUN THE GUI VERSION FIRST BEFORE RUNNING THE COMMAND-LINE VERSION ***\n*** IFCsvr needs to be installed ***"
  exit
} 

# no arguments, no file, print help, and exit

if {$argc == 1} {set arg [string tolower [lindex $argv 0]]}
if {$argc == 0 || ($argc == 1 && ($arg == "help" || $arg == "-help" || $arg == "-h" || $arg == "-v"))} {
  if {$nistVersion} {
    puts "\nUsage: STEP-File-Analyzer-CL.exe myfile.stp \[options ...\]"
  } else {
    puts "\nUsage: sfa-cl.exe myfile.stp \[options ...\]"
  }
  puts "\nWhere options include:\n"
  puts "  csv       Generate CSV files"                                                                                        
  puts "  noopen    Do not open spreadsheet after it has been generated"                                                                                        

  puts "\nOptions last used in the GUI version are used in this program."

if {$nistVersion} { 
puts "\n\nDisclaimers:
   
This software was developed at the National Institute of Standards and Technology
by employees of the Federal Government in the course of their official duties.  
Pursuant to Title 17 Section 105 of the United States Code this software is not
subject to copyright protection and is in the public domain.  This software is an
experimental system.  NIST assumes no responsibility whatsoever for its use by
other parties, and makes no guarantees, expressed or implied, about its quality,
reliability, or any other characteristic.

This software uses Microsoft Excel and IFCsvr which are covered by their own
EULAs (End-User License Agreements)."
}
  exit
}

# get arguments and initialize variables
for {set i 1} {$i <= 100} {incr i} {
  set arg [string tolower [lindex $argv $i]]
  if {$arg != ""} {
    lappend larg $arg
    if {[string first "noopen" $arg] == 0} {set opt(XL_OPEN) 0}                              
    if {[string first "csv"    $arg] == 0} {set opt(XLSCSV) "CSV"}                              
  }
}

# options used from GUI version
puts "\nOptions last used in the GUI version are being used.  Some of them are:"
if {$opt(PMISEM)}  {puts " PMI Representation Report"}
if {$opt(PMIGRF)}  {puts " PMI Presentation Report"}
if {$opt(VALPROP)} {puts " Validation Properties Report"}
if {$opt(INVERSE)} {puts " Inverse Relationships"}

set localName [lindex $argv 0]
if {[string first ":" $localName] == -1} {set localName [file join [pwd] $localName]}
set localName [file nativename $localName]
set remoteName $localName

if {[file exists $localName]} {
  genExcel
} else {
  outputMsg "File not found: [truncFileName $localName]"
}

