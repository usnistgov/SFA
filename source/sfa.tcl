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

# This is the main routine for the STEP File Analyzer

global env tcl_platform

set wdir [file dirname [info script]]
set auto_path [linsert $auto_path 0 $wdir]

# for freeWrap the following lappend commands add package locations to auto_path, must be before package commands
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/tcom3.9
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/twapi3.0.32
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Tclx8.4
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Itk3.4
lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Itcl3.4
lappend auto_path C:/Tcl/lib/teapot/package/tcl/lib/Iwidgets4.0.2

# Tcl packages
package require tcom
package require twapi
package require Tclx
package require Iwidgets 4.0.2

catch {
  lappend auto_path C:/Tcl/lib/teapot/package/tcl/lib/tooltip1.4.4
  package require tooltip
}

set nistVersion 0
foreach item $auto_path {if {[string first "STEP-File-Analyzer" $item] != -1} {set nistVersion 1}}

foreach id {XL_OPEN XL_KEEPOPEN XL_LINK1 XL_FPREC \
            VALPROP PMIGRF PMISEM GENX3DOM DEBUG1 DEBUG2 \
            INVERSE SORT PR_USER \
            PR_STEP_AP242 PR_STEP_AP242_QUAL PR_STEP_AP242_CONS \
            PR_STEP_AP242_MATH PR_STEP_AP242_KINE PR_STEP_AP242_GEOM \
            PR_STEP_AP203 PR_STEP_AP209 PR_STEP_AP210 PR_STEP_AP214 PR_STEP_AP238 \
            PR_STEP_OTHER PR_STEP_GEO PR_STEP_QUAN PR_STEP_PRES PR_STEP_TOLR PR_STEP_REP PR_STEP_CPNT PR_STEP_ASPECT} {set opt($id) 1}

set opt(SORT)  0
set opt(PR_USER) 0
set opt(PR_STEP_GEO)  0
set opt(PR_STEP_CPNT) 0

set opt(XLSCSV) Excel

set opt(XL_KEEPOPEN) 0

set opt(DEBUGINV) 0
set opt(DEBUG1) 0
set opt(DEBUG2) 0
set pointLimit 2
set opt(gpmiColor) 2

set x3domFileOpen 1
set x3domFileName ""

set coverageSTEP 0
set coverageValues 0

set edmWriteToFile 0
set edmWhereRules 0
set eeWriteToFile  0
set opt(indentStyledItem) 0
set opt(indentGeometry) 0
set opt(CRASH) 0
set opt(DISPGUIDE1) 1

set developer 0
if {$env(USERNAME) == "lipman"} {set developer 1}

# -----------------------------------------------------------------------------------------------------
# set drive, myhome, mydocs, mydesk
setHomeDir

# -----------------------------------------------------------------------------------------------------
# initialize data
initData
initDataInverses

set userWriteDir $mydocs
set writeDir ""
set opt(writeDirType) 0
set opt(ROWLIM) 10000000

set openFileList {}
set fileDir  $mydocs
set fileDir1 $mydocs
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

set filemenuinc 4
set lenlist 25
set upgrade 0
set excelYear ""

set writeDir $userWriteDir

set userXLSFile ""

set dispCmd ""
set dispCmds {}

set opt(FIRSTTIME) 1
set lastXLS  ""
set lastXLS1 ""
set sfaVersion 0

# check for options file and source
set optionserr ""
if {[file exists $optionsFile]} {
  catch {source $optionsFile} optionserr
  if {[string first "+" $optionserr] == 0} {set optionserr ""}

# check for old variable names
  if {[info exists verite]} {set sfaVersion $verite; unset verite}

  if {[info exists opt(PMIVRML)]} {set opt(GENX3DOM) $opt(PMIVRML)}
  if {[info exists opt(PMIPROP)]} {set opt(PMIGRF)   $opt(PMIPROP)}
  if {[info exists opt(SEMPROP)]} {set opt(PMISEM)   $opt(SEMPROP)}

  if {[info exists indentStyledItem]} {set opt(indentStyledItem) $indentStyledItem; unset indentStyledItem}
  if {[info exists indentGeometry]}   {set opt(indentGeometry)   $indentGeometry;   unset indentGeometry}
  if {[info exists writeDirType]}     {set opt(writeDirType)     $writeDirType;     unset writeDirType}

  if {[info exists gpmiColor]} {set opt(gpmiColor) $gpmiColor; unset gpmiColor}
  if {[info exists row_limit]} {set opt(ROWLIM)    $row_limit; unset row_limit}
  if {[info exists firsttime]} {set opt(FIRSTTIME) $firsttime; unset firsttime}
  if {[info exists ncrash]}    {set opt(CRASH)     $ncrash;    unset ncrash}

  if {[info exists flag(CRASH)]}      {set opt(CRASH)      $flag(CRASH);      unset flag(CRASH)}
  if {[info exists flag(FIRSTTIME)]}  {set opt(FIRSTTIME)  $flag(FIRSTTIME);  unset flag(FIRSTTIME)}
  if {[info exists flag(DISPGUIDE1)]} {set opt(DISPGUIDE1) $flag(DISPGUIDE1); unset flag(DISPGUIDE1)}

  foreach item {PR_STEP_BAD PR_STEP_UNIT PR_TYPE XL_XLSX COUNT EX_A2P3D FN_APPEND XL_LINK2 XL_LINK3 XL_ORIENT \
                XL_SCROLL PMIVRML PMIPROP SEMPROP PMIP EX_ANAL EX_ARBP EX_LP VPDBG} {
    catch {unset opt($item)}
  }
}

if {[info exists userWriteDir]} {if {![file exists $userWriteDir]} {set userWriteDir $mydocs}}
if {[info exists fileDir]}      {if {![file exists $fileDir]}      {set fileDir      $mydocs}}
if {[info exists fileDir1]}     {if {![file exists $fileDir1]}     {set fileDir1     $mydocs}}
if {[info exists userEntityFile]} {
  if {![file exists $userEntityFile]} {
    set userEntityFile ""
    set opt(PR_USER) 0
  }
}

if {[string index $opt(ROWLIM) end] == 1} {set opt(ROWLIM) [expr {$opt(ROWLIM)+2}]}

# set program files
set programfiles "C:/Program Files"
set pf64 ""
if {[info exists env(ProgramFiles)]} {set programfiles $env(ProgramFiles)}
if {[info exists env(ProgramW6432)]} {set pf64 $env(ProgramW6432)}

# default installation directory for IFCsvr toolkit
set ifcsvrdir [file join $programfiles IFCsvrR300 dll]

#-------------------------------------------------------------------------------
# get programs that can display STEP files
getDisplayPrograms

#-------------------------------------------------------------------------------
# user interface
guiStartWindow

# top menu
set Menu [menu .menubar]
. config -men .menubar
foreach m {File Websites Help} {
  set $m [menu .menubar.m$m -tearoff 1]
  .menubar add cascade -label $m -menu .menubar.m$m
}

# check if menu font is Segoe UI
catch {
  if {$tcl_platform(osVersion) >= 6.0} {
    set ff [join [$File cget -font]]
    if {[string first "Segoe" $ff] == -1} {
      $File     configure -font [list {Segoe UI}]
      $Websites configure -font [list {Segoe UI}]
      $Help     configure -font [list {Segoe UI}]
    }
  }
}

# File menu
guiFileMenu

# What's New
set progtime 0
foreach item {sfa sfa-data sfa-data1 sfa-dimtol sfa-ent sfa-entcsv sfa-gen sfa-geotol \
              sfa-grafpmi sfa-gui sfa-indent sfa-inv sfa-multi sfa-proc sfa-step sfa-valprop} {
  set fname [file join $wdir $item.tcl]
  set mtime [file mtime $fname]
  if {$mtime > $progtime} {set progtime $mtime}
}

proc whatsNew {} {
  global progtime sfaVersion
  
  if {$sfaVersion > 0 && $sfaVersion < [getVersion]} {outputMsg "\nThe previous version of the STEP File Analyzer was: $sfaVersion" red}

outputMsg "\nWhat's New (Version: [getVersion]  Updated: [string trim [clock format $progtime -format "%e %b %Y"]])" blue
outputMsg "- Support for CSV files (Options tab)
- Support for other STEP application protocols (Help > Other STEP APs)
- Source code on GitHub (Websites menu)
- Updated User's Guide Version 3
- PMI Representation Coverage worksheet Color-coded by expected PMI in NIST CAD models (Help > Coverage Analysis)
- Improved processing of Datum Targets and AP209 files
- Improved reporting of Associated Geometry for Tolerances and Annotations
- Feature Count (e.g. 4X) removed from Tolerances, can be derived from Associated Geometry
- Fixed bug that slowed processing when Inverse Relationships was selected"

  .tnb select .tnb.status
  update idletasks
}

# Help and Websites menu
guiHelpMenu
guiWebsitesMenu

# tabs
set nb [ttk::notebook .tnb]
pack $nb -fill both -expand true

# status tab
guiStatusTab

# options tab
guiProcessAndReports

# inverse relationships
guiInverse

# display option
guiDisplayResult
pack $fopt -side top -fill both -expand true -anchor nw

# spreadsheet tab
guiSpreadsheet

# generate logo, progress bars
guiButtons

#-------------------------------------------------------------------------------
# resize tabbed notebook in 8.5 because iwidgets messagebox is too big
update idletasks
catch {
  set maxx [expr {max([winfo reqwidth $fopt],[winfo reqwidth $fxls])}]
  $nb configure -width $maxx
  set maxy [expr {max([winfo reqheight $fopt],[winfo reqheight $fxls])}]
  $nb configure -height [expr {$maxy*1.5}]
}

#-------------------------------------------------------------------------------
# first time user
set copyrose 0
set ask 0
if {$opt(FIRSTTIME)} {
  #helpOverview
  whatsNew
  if {$nistVersion} {displayDisclaimer}
  
  set sfaVersion [getVersion]
  set opt(FIRSTTIME) 0
  
  after 1000
  displayGuide
  set opt(DISPGUIDE1) 0
  
  saveState
  set copyrose 1
  setShortcuts
  
  outputMsg " "
  errorMsg "Use function keys F6 and F5 to change this font size."
  saveState

# what's new message
} elseif {$sfaVersion < [getVersion]} {
  whatsNew
  if {$sfaVersion < 1.60} {
    errorMsg "- Version 3 of the User's Guide is now available"
    displayGuide
  }
  set sfaVersion [getVersion]
  saveState
  set copyrose 1
  setShortcuts
}

if {[info exists env(ROSE)]} {
  set n [string range $env(ROSE) end-2 end-1]
  set stdir [file join $programfiles "STEP Tools" "ST-Runtime $n" schemas]
  if {[file exists $stdir]} {set copyrose 1}
}

if {$developer} {set copyrose 1}

#-------------------------------------------------------------------------------
# crash recovery message
if {$opt(CRASH) < 2} {
  displayCrashRecovery
  incr opt(CRASH)
  saveState
}

#-------------------------------------------------------------------------------
# check for update every 30 days
if {$nistVersion} {
  if {$upgrade > 0} {
    set lastupgrade [expr {round(([clock seconds] - $upgrade)/86400.)}]
    if {$lastupgrade > 30} {
      set choice [tk_messageBox -type yesno -default yes -title "Check for Update" \
        -message "Do you want to check for a newer version of the STEP File Analyzer?\n \nThe last check for an update was $lastupgrade days ago.\n \nYou can always check for an update with Help > Check for Update" -icon question]
      if {$choice == "yes"} {
        set os $tcl_platform(osVersion)
        if {$pf64 != ""} {append os ".64"}
        set url "http://ciks.cbt.nist.gov/cgi-bin/ctv/sfa_upgrade.cgi?version=[getVersion]&auto=$lastupgrade&os=$os"
        if {[info exists excelYear]} {if {$excelYear != ""} {append url "&yr=[expr {$excelYear-2000}]"}}
        displayURL $url
      }
      set upgrade [clock seconds]
      saveState
    }
  } else {
    set upgrade [clock seconds]
    saveState
  }
}

# display user's guide if it hasn't already
if {$opt(DISPGUIDE1)} {
  displayGuide
  set opt(DISPGUIDE1) 0
  saveState
}

#-------------------------------------------------------------------------------
# install IFCsvr
installIFCsvr

focus .

# check command line arguments or drag-and-drop
if {$argv != ""} {
  set localName [lindex $argv 0]
  if {[file dirname $localName] == "."} {
    set localName [file join [pwd] $localName]
  }
  if {$localName != ""} {
    set localNameList [list $localName]
    outputMsg "Ready to process: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)" blue
    if {[info exists buttons(appDisplay)]} {$buttons(appDisplay) configure -state normal}
    if {[info exists buttons(genExcel)]} {
      $buttons(genExcel) configure -state normal
      focus $buttons(genExcel)
    }
  }
}

set writeDir $userWriteDir
checkValues

if {[string length $optionserr] > 5} {
  errorMsg "ERROR reading options file: $optionsFile\n $optionserr"
  errorMsg "Some previously saved options might be lost."
  .tnb select .tnb.status
}

set pid2 [twapi::get_process_ids -name "STEP-File-Analyzer.exe"]
set pid2 [concat $pid2 [twapi::get_process_ids -name "sfa.exe"]]

if {[llength $pid2] > 1} {
  set msg "There are ([expr {[llength $pid2]-1}]) other instances of the STEP File Analyzer already running.\nThe windows for the other instances might not be visible but will show up in the Windows Task Manager as STEP-File-Analyzer.exe"
  append msg "\n\nDo you want to close the other instances of the STEP File Analyzer?"
  set choice [tk_messageBox -type yesno -default yes -message $msg -icon question -title "Close other the STEP File Analyzer?"]
  if {$choice == "yes"} {
    foreach pid $pid2 {
      if {$pid != [pid]} {catch {twapi::end_process $pid -force}}
    }
    outputMsg "Other STEP File Analyzers closed" red
  }
}

# copy schema rose files that are in the Tcl Virtual File System (VFS) or STEPtools runtime to the IFCsvr dll directory
if {$copyrose} {copyRoseFiles}

if {$opt(writeDirType) == 1} {
  outputMsg " "
  errorMsg "Spreadsheets will be written to a user-defined file name (Spreadsheet tab)"
} elseif {$opt(writeDirType) == 2} {
  outputMsg " "
  errorMsg "Spreadsheets will be written to a user-defined directory (Spreadsheet tab)"
}

# set window minimum size
update idletasks
if {![info exists mingeo]} {
  set mingeo "[winfo width .]x[winfo height .]"
  wm minsize . [winfo width .] [winfo height .]
  saveState
} else {
  set wg [split $mingeo "x"]
  wm minsize . [lindex $wg 0] [lindex $wg 1]
}

#outputMsg [info script]
#outputMsg [file system [info script]]
