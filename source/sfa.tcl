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

# This is the main routine for the STEP File Analyzer and Viewer GUI version

global env tcl_platform

set scriptName [info script]
set wdir [file dirname $scriptName]
set auto_path [linsert $auto_path 0 $wdir]
set contact [getContact]

# for building your own version with freewrap, uncomment and modify C:/Tcl/lib/teapot directory if necessary
# the lappend commands add package locations to auto_path, must be before package commands below
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/tcom3.9
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/twapi3.0.32
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Tclx8.4
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Itk3.4
#lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Itcl3.4
#lappend auto_path C:/Tcl/lib/teapot/package/tcl/lib/Iwidgets4.0.2

# Tcl packages, check if they will load
if {[catch {
  package require tcom
  package require twapi
  package require Tclx
  package require Iwidgets 4.0.2
} emsg]} {
  set dir $wdir
  set c1 [string first [file tail [info nameofexecutable]] $dir]
  if {$c1 != -1} {set dir [string range $dir 0 $c1-1]}
  if {[string first "couldn't load library" $emsg] != -1} {
    append emsg "\n\nThere might be a problem running this software from a directory with accented, non-English, or symbol characters in the pathname."
    append emsg "\n     [file nativename $dir]\nTry running the software from a directory without any of the special characters in the pathname."
  }
  append emsg "\n\nContact [lindex $contact 0] ([lindex $contact 1]) if you cannot run the STEP File Analyzer and Viewer."
  set choice [tk_messageBox -type ok -icon error -title "ERROR starting the STEP File Analyzer and Viewer" -message $emsg]
  exit
}

catch {
  #lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/vfs1.4.2
  package require vfs::zip
}

catch {
  #lappend auto_path C:/Tcl/lib/teapot/package/tcl/lib/tooltip1.4.4
  package require tooltip
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
set nistVersion 0
foreach item $auto_path {if {[string first "STEP-File-Analyzer" $item] != -1} {set nistVersion 1}}

# -----------------------------------------------------------------------------------------------------
# initialize variables, set opt to 1
foreach id { \
  DELCOVROWS DISPGUIDE1 FIRSTTIME LOGFILE PMIGRF PMISEM PR_STEP_AP242 PR_STEP_COMM PR_STEP_COMP PR_STEP_FEAT PR_STEP_KINE PR_STEP_PRES PR_STEP_QUAN \
  PR_STEP_REPR PR_STEP_SHAP PR_STEP_TOLR VALPROP VIZBRP VIZFEA VIZFEABC VIZFEADS VIZFEALV VIZPMI VIZTPG XL_LINK1 XL_OPEN \
} {set opt($id) 1}

# set opt to 0
foreach id { \
  CRASH DEBUG1 DEBUGINV indentGeomtry indentStyledItem INVERSE PMIGRFCOV PMISEMDIM PR_STEP_CPNT PR_STEP_GEOM \
  PR_USER VIZFEADSntail VIZFEALVS VIZPMIVP VIZTPGMSH writeDirType XL_FPREC XL_SORT \
} {set opt($id) 0}

set opt(gpmiColor) 3
set opt(XL_ROWLIM) 1003
set opt(XLSBUG1) 30
set opt(XLSCSV) Excel

set coverageSTEP 0
set dispCmd "Default"
set dispCmds {}
set edmWhereRules 0
set edmWriteToFile 0
set eeWriteToFile  0
set excelVersion 12
set lastX3DOM ""
set lastXLS  ""
set lastXLS1 ""
set openFileList {}
set pointLimit 2
set sfaVersion 0
set upgrade 0
set upgradeIFCsvr 0
set userXLSFile ""
set x3dFileName ""
set x3dStartFile 1

set fileDir  $mydocs
set fileDir1 $mydocs
set userWriteDir $mydocs
set writeDir $userWriteDir

set developer 0
if {$env(USERDOMAIN) == "NIST"} {set developer 1}

# initialize data
initData
initDataInverses

# -----------------------------------------------------------------------------------------------------
# check for options file and read (source)
set optionsFile [file nativename [file join $fileDir STEP-File-Analyzer-options.dat]]
if {[file exists $optionsFile]} {
  if {[catch {
    source $optionsFile

# rename and unset old variable names from old options file
    if {[info exists opt(VIZTES)]}    {set opt(VIZTPG)    $opt(VIZTES);    unset opt(VIZTES)}
    if {[info exists opt(VIZTESMSH)]} {set opt(VIZTPGMSH) $opt(VIZTESMSH); unset opt(VIZTESMSH)}

    if {[info exists verite]} {set sfaVersion $verite; unset verite}
    if {[info exists indentStyledItem]} {set opt(indentStyledItem) $indentStyledItem; unset indentStyledItem}
    if {[info exists indentGeometry]}   {set opt(indentGeometry)   $indentGeometry;   unset indentGeometry}
    if {[info exists writeDirType]}     {set opt(writeDirType)     $writeDirType;     unset writeDirType}
  
    if {[info exists gpmiColor]} {set opt(gpmiColor) $gpmiColor; unset gpmiColor}
    if {[info exists row_limit]} {set opt(XL_ROWLIM) $row_limit; unset row_limit}
    if {[info exists firsttime]} {set opt(FIRSTTIME) $firsttime; unset firsttime}
    if {[info exists ncrash]}    {set opt(CRASH)     $ncrash;    unset ncrash}
  
    if {[info exists flag(CRASH)]}      {set opt(CRASH)      $flag(CRASH);      unset flag(CRASH)}
    if {[info exists flag(FIRSTTIME)]}  {set opt(FIRSTTIME)  $flag(FIRSTTIME);  unset flag(FIRSTTIME)}
    if {[info exists flag(DISPGUIDE1)]} {set opt(DISPGUIDE1) $flag(DISPGUIDE1); unset flag(DISPGUIDE1)}
  
    foreach item {COUNT EX_A2P3D EX_ANAL EX_ARBP EX_LP feaNodeType FN_APPEND GENX3DOM PMIP PMIPROP PMIVRML PR_STEP_AP203 PR_STEP_AP209 PR_STEP_AP210 \
                  PR_STEP_AP214 PR_STEP_AP238 PR_STEP_AP239 PR_STEP_AP242_CONS PR_STEP_AP242_GEOM PR_STEP_AP242_KINE PR_STEP_AP242_MATH PR_STEP_AP242_OTHER \
                  PR_STEP_AP242_QUAL PR_STEP_ASPECT PR_STEP_BAD PR_STEP_GEO PR_STEP_OTHER PR_STEP_REP PR_STEP_UNIT PR_TYPE ROWLIM SEMPROP SORT VIZ209 \
                  VIZBRPmsg VIZFEADStail VPDBG XL_KEEPOPEN XL_LINK2 XL_LINK3 XL_ORIENT XL_SCROLL XL_XLSX XLSBUG} {
      catch {unset opt($item)}
    }
  } emsg]} {
    set endMsg "Error reading options file: [truncFileName $optionsFile]\n $emsg\nFix or delete the file."
  }
}

# check some directory variables
if {[info exists userWriteDir]} {if {![file exists $userWriteDir]} {set userWriteDir $mydocs}}
if {[info exists fileDir]}      {if {![file exists $fileDir]}      {set fileDir      $mydocs}}
if {[info exists fileDir1]}     {if {![file exists $fileDir1]}     {set fileDir1     $mydocs}}
if {[info exists userEntityFile]} {
  if {![file exists $userEntityFile]} {
    set userEntityFile ""
    set opt(PR_USER) 0
  }
}

# fix row limit
if {$opt(XL_ROWLIM) < 103 || ([string range $opt(XL_ROWLIM) end-1 end] != "03" && \
   [string range $opt(XL_ROWLIM) end-1 end] != "76" && [string range $opt(XL_ROWLIM) end-1 end] != "36")} {set opt(XL_ROWLIM) 103}

# for output format buttons
set ofExcel 0
set ofCSV 0
set ofNone 0
switch -- $opt(XLSCSV) {
  Excel   {set ofExcel 1}
  CSV     {set ofExcel 1; set ofCSV 1}
  None    {set ofNone 1}
  default {set ofExcel 1}
}

# -------------------------------------------------------------------------------
# get programs that can open STEP files
getOpenPrograms

# -------------------------------------------------------------------------------
# user interface
guiStartWindow

# top menu
set Menu [menu .menubar]
. config -men .menubar
foreach m {File Websites Examples Help} {
  set $m [menu .menubar.m$m -tearoff 1]
  .menubar add cascade -label $m -menu .menubar.m$m
}

# check if menu font is Segoe UI for windows 7 or greater
catch {
  if {$tcl_platform(osVersion) >= 6.0} {
    set ff [join [$File cget -font]]
    if {[string first "Segoe" $ff] == -1} {
      $File     configure -font [list {Segoe UI}]
      $Websites configure -font [list {Segoe UI}]
      $Examples configure -font [list {Segoe UI}]
      $Help     configure -font [list {Segoe UI}]
    }
  }
}

# File menu
guiFileMenu

# What's New
set progtime 0
foreach fname [glob -nocomplain -directory $wdir *.tcl] {
  set mtime [file mtime $fname]
  if {$mtime > $progtime} {set progtime $mtime}
}

# -------------------------------------------------------------------------------
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

# open option, output format
guiOpenSTEPFile
pack $fopt -side top -fill both -expand true -anchor nw

# spreadsheet tab
guiSpreadsheet

# generate logo, progress bars
guiButtons

# switch to options tab (any text output will switch back to the status tab)
.tnb select .tnb.options

#-------------------------------------------------------------------------------
# first time user
set ask 0

if {$opt(FIRSTTIME)} {
  whatsNew
  if {$nistVersion} {showDisclaimer}
  
  set sfaVersion [getVersion]
  set opt(FIRSTTIME) 0
  
  after 1000
  showFileURL UserGuide
  set opt(DISPGUIDE1) 0
  
  saveState
  setShortcuts
  
  outputMsg " "
  errorMsg "Use F8 and F9 to change the font size."
  saveState

# what's new message
} elseif {$sfaVersion < [getVersion]} {
  whatsNew
  set sfaVersion [getVersion]
  saveState
  setShortcuts

} elseif {$sfaVersion > [getVersion]} {
  set sfaVersion [getVersion]
  saveState
}

#-------------------------------------------------------------------------------
# crash recovery message
if {$opt(CRASH) < 2} {
  showCrashRecovery
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
        -message "Do you want to check for a new version of the STEP File Analyzer and Viewer?\n\nThe last check for an update was $lastupgrade days ago.\n\nYou can always check for an update with Help > Check for Update" -icon question]
      if {$choice == "yes"} {
        set url "https://concrete.nist.gov/cgi-bin/ctv/sfa_upgrade.cgi?version=[getVersion]&auto=$lastupgrade"
        openURL $url
      }
      set upgrade [clock seconds]
      saveState
    }
  } else {
    set upgrade [clock seconds]
    saveState
  }
}

# open user guide if it hasn't already
if {$opt(DISPGUIDE1)} {
  showFileURL UserGuide
  set opt(DISPGUIDE1) 0
  saveState
}

#-------------------------------------------------------------------------------
# install IFCsvr
set ifcsvrDir [file join $pf32 IFCsvrR300 dll]
if {![file exists [file join $ifcsvrDir IFCsvrR300.dll]]} {
  installIFCsvr

# or reinstall IFCsvr
} elseif {$nistVersion} {
  set ifcsvrTime [file mtime [file join $wdir exe ifcsvrr300_setup_1008_en-update.msi]]
  if {$ifcsvrTime > $upgradeIFCsvr} {installIFCsvr 1}
}

focus .

# check command line arguments or drag-and-drop
if {$argv != ""} {
  set localName [lindex $argv 0]
  if {[file dirname $localName] == "."} {
    set localName [file join [pwd] $localName]
  }
  if {$localName != ""} {
    .tnb select .tnb.status
    if {[file exists $localName]} {
      set localNameList [list $localName]
      outputMsg "Ready to process: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)" green
      if {[info exists buttons(appOpen)]} {$buttons(appOpen) configure -state normal}
      if {[info exists buttons(genExcel)]} {
        $buttons(genExcel) configure -state normal
        focus $buttons(genExcel)
        if {$editorCmd != ""} {
          bind . <Key-F5> {
            if {[file exists $localName]} {
              outputMsg "\nOpening STEP file: [file tail $localName]"
              exec $editorCmd [file nativename $localName] &
            }
          }
        }
      }
    } else {
      errorMsg "File not found: [truncFileName [file nativename $localName]]"
    }
  }
}

set writeDir $userWriteDir
checkValues

# other STEP File Analyzer and Viewers already running
set pid2 [twapi::get_process_ids -name "STEP-File-Analyzer.exe"]
set pid2 [concat $pid2 [twapi::get_process_ids -name "sfa.exe"]]

if {[llength $pid2] > 1} {
  set msg "There are at least ([expr {[llength $pid2]-1}]) other instances of the STEP File Analyzer and Viewer already running.\n\nDo you want to close them?"
  set choice [tk_messageBox -type yesno -default yes -message $msg -icon question -title "Close?"]
  if {$choice == "yes"} {
    foreach pid $pid2 {
      if {$pid != [pid]} {catch {twapi::end_process $pid -force}}
    }
    outputMsg "Other STEP File Analyzer and Viewers closed" red
    .tnb select .tnb.status
  }
}

# set process id used to check memory usage for AP209 files
set sfaPID [twapi::get_process_ids -name "STEP-File-Analyzer.exe"]

# warn if spreadsheets not written to default directory
if {$opt(writeDirType) == 1} {
  outputMsg " "
  errorMsg "Spreadsheets will be written to a user-defined file name (Spreadsheet tab)"
  .tnb select .tnb.status
} elseif {$opt(writeDirType) == 2} {
  outputMsg " "
  errorMsg "Spreadsheets will be written to a user-defined directory (Spreadsheet tab)"
  .tnb select .tnb.status
}

# warn about output type
if {$opt(XLSCSV) == "CSV"} {
  outputMsg " "
  errorMsg "CSV files will be generated (Options tab)"
  .tnb select .tnb.status
} elseif {$opt(XLSCSV) == "None"} {
  outputMsg " "
  errorMsg "No Spreadsheet will be generated, only Views (Options tab)"
  .tnb select .tnb.status
}

# error messages from before GUI was available
if {[info exists endMsg]} {
  outputMsg " "
  errorMsg $endMsg
  .tnb select .tnb.status
}
  
# set window minimum size
update idletasks
wm minsize . [winfo reqwidth .] [expr {int([winfo reqheight .]*1.05)}]

# debug
#compareLists "AP242" $ap242all $ap242e1

#set all [lrmdups [concat $ap203all $ap214all $ap242all]] 
#foreach idx [array names entCategory] {compareLists "$idx" $all $entCategory($idx); outputMsg "--------------"}

#set apcat {}
#foreach idx [array names entCategory] {set apcat [concat $apcat $entCategory($idx)]}
#compareLists "cat" $apcat [lrmdups [concat $ap203all $ap214all $ap242all]]
#foreach idx [array names entCategory] {if {[llength $entCategory($idx)] != [llength [lrmdups $entCategory($idx)]]} {outputMsg $idx}}
