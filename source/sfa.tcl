# This is the main routine for the STEP File Analyzer and Viewer GUI version

# Website - https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer
# NIST Disclaimer - https://www.nist.gov/disclaimer
# Source code - https://github.com/usnistgov/SFA

global env

set scriptName [info script]
set wdir [file dirname $scriptName]
set auto_path [linsert $auto_path 0 $wdir]

# detect if NIST version
set nistVersion 0
foreach item $auto_path {if {[string first "STEP-File-Analyzer" $item] != -1} {set nistVersion 1}}

# for building your own version with freewrap, the following are explicitly added to auto_path
# change C:/Tcl if Tcl is installed in a different directory
if {!$nistVersion} {
  lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/tcom3.9
  lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/tdom0.8.3
  lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/twapi3.0.32
  lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Tclx8.4
  lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Itk3.4
  lappend auto_path C:/Tcl/lib/teapot/package/win32-ix86/lib/Itcl3.4
  lappend auto_path C:/Tcl/lib/teapot/package/tcl/lib/Iwidgets4.0.2
}

# Tcl packages, check if they will load
if {[catch {
  package require tcom
  package require tdom
  package require twapi
  package require Tclx
  package require Iwidgets 4.0.2
  if {$nistVersion} {package require tooltip; package require vfs::zip}
} emsg]} {
  set dir $wdir
  set c1 [string first [file tail [info nameofexecutable]] $dir]
  if {$c1 != -1} {set dir [string range $dir 0 $c1-1]}
  if {[string first "couldn't load library" $emsg] != -1} {
    append emsg "\n\nAlthough the message above indicates that a library is missing, that is NOT the cause of the problem.  The problem is sometimes related to the directory where you are running the software.\n\n   [file nativename $dir]"
    append emsg "\n\n1 - The directory has accented, non-English, or symbol characters"
    append emsg "\n2 - The directory is on a remote computer"
    append emsg "\n3 - No permissions to run the software in the directory"
    append emsg "\n\nTry these workarounds to run the software:"
    append emsg "\n\n1 - From a directory without any special characters in the pathname, or from your home directory, or desktop"
    append emsg "\n2 - From a directory on your local computer"
    append emsg "\n3 - As Administrator"
    append emsg "\n4 - On a different computer"
  }
  set choice [tk_messageBox -type ok -icon error -title "Error running the STEP File Analyzer and Viewer" -message $emsg]
  exit
}

# -----------------------------------------------------------------------------------------------------
# initialize all data
initData
initDataInverses
set edmWhereRules 0
set edmWriteToFile 0

# -----------------------------------------------------------------------------------------------------
# check for options file
set optionsFile [file nativename [file join $fileDir STEP-File-Analyzer-options.dat]]

# copy file from old location if not the same as new location (on OneDrive)
if {![file exists $optionsFile]} {
  set oldFile [file join $myhome Documents STEP-File-Analyzer-options.dat]
  if {[file exists $oldFile]} {
    file copy -force $oldFile [file dirname $optionsFile]
    file delete -force $oldFile
  }
}

# read (source) options file
if {[file exists $optionsFile]} {
  if {[catch {
    source $optionsFile
    checkVariables
  } emsg]} {
    set endMsg "Error reading Options file [truncFileName $optionsFile]: $emsg"
  }
}

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

# generate tab
guiGenerateTab
guiOpenSTEPFile
pack $fopt -side top -fill both -expand true -anchor nw

# more tab
guiMoreTab

# generate logo, progress bars
guiButtons

# switch to generate tab (any text output will switch back to the status tab)
.tnb select .tnb.generate

if {$developer} {if {$filesProcessed > 0} {outputMsg $filesProcessed} else {errorMsg $filesProcessed}}

# error messages from before GUI was available
if {[info exists endMsg]} {
  outputMsg " "
  errorMsg $endMsg
  .tnb select .tnb.status
}

#-------------------------------------------------------------------------------
# first time user
if {$sfaVersion == 0} {
  whatsNew
  setShortcuts
  openUserGuide
  showCrashRecovery
  saveState

# what's new message
} elseif {$sfaVersion < [getVersion]} {
  whatsNew
  setShortcuts
  saveState

} elseif {$sfaVersion > [getVersion]} {
  set sfaVersion [getVersion]
  saveState
}

#-------------------------------------------------------------------------------
# check for update every 90 days
if {$nistVersion} {
  if {$upgrade > 0} {
    set lastupgrade [expr {round(([clock seconds] - $upgrade)/86400.)}]
    if {$lastupgrade > 90} {
      set str ""
      if {$lastupgrade > 365} {set str ".  Welcome Back!"}
      outputMsg "The last check for an update was $lastupgrade days ago$str\nTo check for an updated version, go to Websites > STEP File Analyzer and Viewer\nTo see what is new in an updated version, go to Help > Release Notes" red
      .tnb select .tnb.status
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
      checkFileSize

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

# warning messages
set warning {}
if {$opt(writeDirType) == 2} {lappend warning "Output files will be written to a User-Defined directory (More tab)"}
if {$opt(PMISEMRND)}         {lappend warning "Rounding semantic PMI dimensions and tolerances (More tab)"}
if {$opt(tessPartOld)}       {lappend warning "Using alternative tessellated geometry processing (More tab)"}
if {$opt(brepAlt)}           {lappend warning "Using alternative b-rep geometry processing (More tab)"}
catch {
  set winver [registry get {HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion} {ProductName}]
  if {[string first "Server" $winver] != -1} {lappend warning "$winver is not supported."}
}
if {[llength $warning] > 0}  {foreach item $warning {errorMsg $item red}; .tnb select .tnb.status}

# set window minimum size
update idletasks
set rw [winfo reqwidth .]
set rh [expr {int([winfo reqheight .]*1.05)}]
if {$rh > [winfo screenheight  .]} {set rh [winfo screenheight .]}
wm minsize . $rw $rh

# debug lists of entities in sfa-data.tcl
#debugData
