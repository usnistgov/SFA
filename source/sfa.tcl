# This software was developed at the National Institute of Standards and Technology by employees of
# the Federal Government in the course of their official duties.  Pursuant to Title 17 Section 105 of
# the United States Code this software is not subject to copyright protection and is in the public
# domain.  This software is an experimental system.  NIST assumes no responsibility whatsoever for
# its use by other parties, and makes no guarantees, expressed or implied, about its quality,
# reliability, or any other characteristic.

# This software is provided by NIST as a public service.  You may use, copy and distribute copies of
# the software in any medium, provided that you keep intact this entire notice.  You may improve,
# modify and create derivative works of the software or any portion of the software, and you may copy
# and distribute such modifications or works.  Modified works should carry a notice stating that you
# changed the software and should note the date and nature of any such change.  Please explicitly
# acknowledge NIST as the source of the software.

# See the NIST Disclaimer at https://www.nist.gov/disclaimer
# The latest version of the source code is available at: https://github.com/usnistgov/SFA

# This is the main routine for the STEP File Analyzer and Viewer GUI version

global env

set scriptName [info script]
set wdir [file dirname $scriptName]
set auto_path [linsert $auto_path 0 $wdir]
set contact [getContact]

# get path for command-line version
set path [split $wdir "/"]
set sfacl {}
foreach item $path {
  if {$item != "STEP-File-Analyzer.exe"} {
    lappend sfacl $item
  } else {
    break
  }
}
lappend sfacl "sfa-cl.exe"
set sfacl [join $sfacl "/"]

# for building your own version with freewrap, uncomment and modify C:/Tcl/lib/teapot directory if necessary
# the lappend commands add package locations to auto_path, must be before package commands below
# see 30 lines below for two more lappend commands
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
    append emsg "\n\nAlthough the message above indicates that a library is missing, that is NOT the cause of the problem.  The problem is usually related to the directory where the software is installed.\n\n[file nativename $dir]"
    append emsg "\n\n1 - The directory has accented, non-English, or symbol characters"
    append emsg "\n2 - The directory is on a different computer"
    append emsg "\n3 - No permissions to run the software in the directory"
    append emsg "\n4 - Other computer configuration problems"
    append emsg "\n\nTry the following workarounds to run the software:"
    append emsg "\n\n1 - From a directory without any special characters in the pathname, or from your home directory, or desktop"
    append emsg "\n2 - Installed on your local computer"
    append emsg "\n3 - As Administrator"
    append emsg "\n4 - On a different computer"
  }
  append emsg "\n\nPlease send a screenshot of this window to [lindex $contact 0] ([lindex $contact 1]) if you cannot fix the problem."
  set choice [tk_messageBox -type ok -icon error -title "ERROR running the STEP File Analyzer and Viewer" -message $emsg]
  exit
}

# for building your own version with freewrap, also uncomment and modify the lappend commands
catch {
  #lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/vfs1.4.2
  package require vfs::zip
}

catch {
  #lappend auto_path C:/Tcl/lib/teapot/package/tcl/lib/tooltip1.4.4
  package require tooltip
}

# -----------------------------------------------------------------------------------------------------
# set program files, environment variables will be in the correct language
set pf32 "C:\\Program Files (x86)"
if {[info exists env(ProgramFiles)]} {set pf32 $env(ProgramFiles)}
set pf64 ""
if {[info exists env(ProgramW6432)]} {set pf64 $env(ProgramW6432)}

# set drive, myhome, mydocs, mydesk
setHomeDir

# detect if NIST version
set nistVersion 0
foreach item $auto_path {if {[string first "STEP-File-Analyzer" $item] != -1} {set nistVersion 1}}

# -----------------------------------------------------------------------------------------------------
# initialize variables, set opt to 1
foreach id { \
  logFile outputOpen partEdges partSketch PMIGRF PMISEM stepCOMM stepPRES stepQUAN stepREPR stepSHAP stepTOLR valProp \
  viewPart viewPMI \
} {set opt($id) 1}

# set opt to 0
foreach id { \
  feaBounds feaDisp feaDispNoTail feaLoads feaLoadScale indentGeometry indentStyledItem INVERSE partNormals partOnly \
  PMIGRFCOV PMISEMDIM PMISEMRND SHOWALLPMI stepAP242 stepCOMP stepCPNT stepFEAT stepGEOM stepKINE stepUSER syntaxChecker \
  tessPartMesh viewFEA viewTessPart writeDirType xlHideLinks xlNoRound xlSort xlUnicode x3dSave DEBUG1 DEBUGINV DEBUGX3D \
} {set opt($id) 0}

set opt(gpmiColor) 3
set opt(partQuality) 7
set opt(xlMaxRows) 1003
set opt(xlFormat) Excel
set gen(AllAnalysis) 0
set gen(AllView) 0
set gen(View) 1

set coverageSTEP 0
set dispCmd "Default"
set dispCmds {}
set edmWhereRules 0
set edmWriteToFile 0
set stepToolsWriteToFile  0
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

    if {[info exists opt(CRASH)]}   {set filesProcessed 1}
    if {[info exists gpmiColor]}    {set opt(gpmiColor) $gpmiColor}
    if {[info exists row_limit]}    {set opt(xlMaxRows) $row_limit}
    if {[info exists writeDirType]} {set opt(writeDirType) $writeDirType}
    if {$opt(writeDirType) == 1}    {set opt(writeDirType) 0}

# unset old unused opt variables
    foreach item {COUNT CRASH DELCOVROWS DISPGUIDE1 feaNodeType FIRSTTIME FN_APPEND indentGeomtry GENX3DOM \
      PMIP PMIPROP PMIVRML ROWLIM SEMPROP SORT feaDisptail VPDBG XLSBUG XLSBUG1} {catch {unset opt($item)}
    }
    foreach id [array names opt] {foreach str {EX_ PR_ XL_ VIZ} {if {[string first $str $id] == 0} {unset opt($id)}}}
  } emsg]} {
    set endMsg "Error reading Options file [truncFileName $optionsFile]: $emsg"
  }
}

# check that directories exist
if {[info exists userWriteDir]} {if {![file exists $userWriteDir]} {set userWriteDir $mydocs}}
if {[info exists fileDir]}      {if {![file exists $fileDir]}      {set fileDir      $mydocs}}
if {[info exists fileDir1]}     {if {![file exists $fileDir1]}     {set fileDir1     $mydocs}}
if {[info exists userEntityFile]} {
  if {![file exists $userEntityFile]} {
    set userEntityFile ""
    set opt(stepUSER) 0
  }
}

# fix row limit
if {$opt(xlMaxRows) < 103 || ([string range $opt(xlMaxRows) end-1 end] != "03" && \
   [string range $opt(xlMaxRows) end-1 end] != "76" && [string range $opt(xlMaxRows) end-1 end] != "36")} {set opt(xlMaxRows) 103}

# for generate buttons
set gen(Excel) 0
set gen(CSV) 0
set gen(None) 0
switch -- $opt(xlFormat) {
  Excel   {set gen(Excel) 1}
  CSV     {set gen(Excel) 1; set gen(CSV) 1}
  None    {set gen(None) 1}
  default {set gen(Excel) 1}
}
set gen(View1) $gen(View)
set gen(Excel1) $gen(Excel)

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

# check if menu font is Segoe UI
catch {
  set ff [join [$File cget -font]]
  if {[string first "Segoe" $ff] == -1} {
    $File     configure -font [list {Segoe UI}]
    $Websites configure -font [list {Segoe UI}]
    $Examples configure -font [list {Segoe UI}]
    $Help     configure -font [list {Segoe UI}]
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
guiOptionsTab
guiOpenSTEPFile
pack $fopt -side top -fill both -expand true -anchor nw

# spreadsheet tab
guiSpreadsheet

# generate logo, progress bars
guiButtons

# switch to options tab (any text output will switch back to the status tab)
.tnb select .tnb.options

if {$developer} {if {$filesProcessed > 0} {outputMsg $filesProcessed} else {errorMsg $filesProcessed}}

# error messages from before GUI was available
if {[info exists endMsg]} {
  outputMsg " "
  errorMsg $endMsg
  .tnb select .tnb.status
}

# non-NIST version
if {!$nistVersion} {
  outputMsg "\nThis is a user-built version of the NIST STEP File Analyzer and Viewer."
  .tnb select .tnb.status
}

#-------------------------------------------------------------------------------
# first time user
set save 0
if {$sfaVersion == 0} {
  whatsNew
  set sfaVersion [getVersion]
  setShortcuts
  showFileURL UserGuide
  showCrashRecovery
  set save 1

# what's new message
} elseif {$sfaVersion < [getVersion]} {
  whatsNew
  set sfaVersion [getVersion]
  setShortcuts
  set save 1

} elseif {$sfaVersion > [getVersion]} {
  set sfaVersion [getVersion]
  set save 1
}

# save the variables
if {$save} {saveState}

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

#-------------------------------------------------------------------------------
# install IFCsvr
installIFCsvr
set ifcsvrDir [file join $pf32 IFCsvrR300 dll]

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
      outputMsg "Ready to process: [file tail $localName] ([fileSize $localName])" green

# check for STL file
      if {[string tolower [file extension $localName]] == ".stl"} {
        set opt(partOnly) 0
        set opt(xlFormat) None
        set opt(viewTessPart) 1
        set gen(Excel) 0
        set gen(CSV) 0
        set gen(None) 1
        set allNone -1
        checkValues
      }

      set fileDir [file dirname $localName]
      if {$fileDir == $drive} {outputMsg "There might be problems processing a STEP file directly in the $fileDir directory." red}

      if {[info exists buttons(appOpen)]} {$buttons(appOpen) configure -state normal}
      if {[info exists buttons(generate)]} {
        $buttons(generate) configure -state normal
        focus $buttons(generate)
        if {$editorCmd != ""} {
          bind . <Key-F5> {
            if {[file exists $localName]} {
              outputMsg "\nOpening STEP file: [file tail $localName]"
              exec $editorCmd [file nativename $localName] &
            }
          }
        }
        bind . <Shift-F5> {
          if {[file exists $localName]} {
            set dir [file nativename [file dirname $localName]]
            outputMsg "\nOpening STEP file directory: [truncFileName $dir]"
            catch {exec C:/Windows/explorer.exe $dir &}
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
set pids {}
catch {
  foreach proc [list "STEP-File-Analyzer.exe"] {
    foreach id [twapi::get_process_ids -name $proc] {
      if {$id != [pid]} {if {[string first "unknown" [twapi::get_process_info $id -commandline]] == -1} {lappend pids $id}}
    }
  }
}

if {[llength $pids] > 0} {
  set msg "There are ([llength $pids]) other STEP File Analyzer and Viewers already running.  Do you want to close them?"
  set choice [tk_messageBox -type yesno -default yes -message $msg -icon question -title "Close?"]
  if {$choice == "yes"} {
    foreach pid $pids {catch {twapi::end_process $pid -force}}
    outputMsg "Closed other STEP File Analyzer and Viewers" red
    .tnb select .tnb.status
  }
}

# warn if spreadsheets not written to default directory
if {$opt(writeDirType) == 2} {
  outputMsg " "
  errorMsg "All output files will be written to a User-defined directory (Spreadsheet tab)"
  .tnb select .tnb.status
}

# set window minimum size
update idletasks
set rw [winfo reqwidth .]
set rh [expr {int([winfo reqheight .]*1.05)}]
if {$rh > [winfo screenheight  .]} {set rh [winfo screenheight .]}
wm minsize . $rw $rh

# debug lists of entities in sfa-data.tcl
#debugData
