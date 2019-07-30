# version numbers, software and user guide, contact
# user guide URLs are below in showFileURL

proc getVersion {}   {return 3.50}
proc getVersionUG {} {return 3.0}
proc getContact {}   {return [list "Robert Lipman" "robert.lipman@nist.gov"]}

# -------------------------------------------------------------------------------
proc whatsNew {} {
  global progtime sfaVersion

  if {$sfaVersion > 0 && $sfaVersion < [getVersion]} {outputMsg "\nThe previous version of the STEP File Analyzer and Viewer was: $sfaVersion" red}

outputMsg "\nWhat's New (Version: [getVersion]  Updated: [string trim [clock format $progtime -format "%e %b %Y"]])
- New features and bug fixes are now documented in the changelog, go to Help > Changelog" blue

if {$sfaVersion > 0 && $sfaVersion <= 2.60} {
  outputMsg "\nRenamed output files:\n Spreadsheets from  myfile_stp.xlsx  to  myfile-sfa.xlsx\n Views from  myfile-x3dom.html  to  myfile-sfa.html" red
}

  .tnb select .tnb.status
  update idletasks
}

#-------------------------------------------------------------------------------
# open local file or URL
proc showFileURL {type} {

  switch -- $type {
    UserGuide {
# update for new versions, local and online
      set localFile "SFA-User-Guide-v5.pdf"
      set URL https://doi.org/10.6028/NIST.AMS.200-6

# extra message if user guide is out-of-date, versions defined at the top of this file
      if {[getVersion] > [expr {[getVersionUG]+0.25}]} {
        outputMsg " "
        errorMsg "The User Guide is based on version [getVersionUG] of the STEP File Analyzer and Viewer.\n See Help > Changelog for changes to the software."
        outputMsg " "
        .tnb select .tnb.status
      }
    }

    Changelog {
# local changelog file should also be on amazon
      set localFile "STEP-File-Analyzer-changelog.xlsx"
      set URL https://s3.amazonaws.com/nist-el/mfg_digitalthread/STEP-File-Analyzer-changelog.xlsx
    }
  }

# open file, local file is assumed to be in same directory as executable, if not open URL
  set byURL 1
  set fname [file nativename [file join [file dirname [info nameofexecutable]] $localFile]]
  if {[file exists $fname]} {

# open spreadsheet
    if {[file extension $fname] == ".xlsx"} {
      openXLS $fname
      set byURL 0

# open other types of files
    } else {
      if {[catch {
        exec {*}[auto_execok start] "" $fname
        set byURL 0
      } emsg]} {
        if {[string first "UNC" $emsg] != -1} {set byURL 0}
      }
    }
  }

# open file by url if local file not opened
  if {$byURL} {openURL $URL}
}

#-------------------------------------------------------------------------------
# start window, bind keys
proc guiStartWindow {} {
  global fout editorCmd lastX3DOM lastXLS lastXLS1 localName localNameList wingeo winpos

  wm title . "STEP File Analyzer and Viewer  (v[getVersion])"
  wm protocol . WM_DELETE_WINDOW {exit}

# check that the saved window dimensions do not exceed the screen size
  if {[info exists wingeo]} {
    set gwid [lindex [split $wingeo "x"] 0]
    set ghgt [lindex [split $wingeo "x"] 1]
    if {$gwid > [winfo screenwidth  .]} {set gwid [winfo screenwidth  .]}
    if {$ghgt > [winfo screenheight .]} {set ghgt [winfo screenheight .]}
    set wingeo "$gwid\x$ghgt"
  }

# check that the saved window position is on the screen
  if {[info exists winpos]} {
    set pwid [lindex [split $winpos "+"] 1]
    set phgt [lindex [split $winpos "+"] 2]
    if {$pwid > [winfo screenwidth  .] || $pwid < -10} {set pwid 300}
    if {$phgt > [winfo screenheight .] || $phgt < -10} {set phgt 200}
    set winpos "+$pwid+$phgt"
  }

# check that the saved window position keeps the entire window on the screen
  if {[info exists wingeo] && [info exists winpos]} {
    if {[expr {$pwid+$gwid}] > [winfo screenwidth  .]} {
      set pwid [expr {[winfo screenwidth  .]-$gwid-40}]
      if {$pwid < 0} {set pwid 300}
    }
    if {[expr {$phgt+$ghgt}] > [winfo screenheight  .]} {
      set phgt [expr {[winfo screenheight  .]-$ghgt-40}]
      if {$phgt < 0} {set phgt 200}
    }
    set winpos "+$pwid+$phgt"
  }

# set the window position and dimensions
  if {[info exists winpos]} {catch {wm geometry . $winpos}}
  if {[info exists wingeo]} {catch {wm geometry . $wingeo}}

# yellow background color
  set bgcolor  "#ffffbb"
  catch {option add *Frame.background       $bgcolor}
  catch {option add *Label.background       $bgcolor}
  catch {option add *Checkbutton.background $bgcolor}
  catch {option add *Radiobutton.background $bgcolor}

  ttk::style configure TCheckbutton -background $bgcolor
  ttk::style map       TCheckbutton -background [list disabled $bgcolor]
  ttk::style configure TRadiobutton -background $bgcolor
  ttk::style map       TRadiobutton -background [list disabled $bgcolor]
  ttk::style configure TLabelframe       -background $bgcolor

  font create fontBold {*}[font configure TkDefaultFont]
  font configure fontBold -weight bold
  ttk::style configure TLabelframe.Label -background $bgcolor -font fontBold

# key bindings
  bind . <Control-o> {openFile}
  bind . <Control-d> {openMultiFile}
  bind . <Control-q> {exit}

  bind . <Key-F1> {
    .tnb select .tnb.status
    set localName [getFirstFile]
    if {$localName != ""} {
      set localNameList [list $localName]
      genExcel
    }
  }

  bind . <Key-F2> {if {$lastXLS   != ""} {set lastXLS  [openXLS $lastXLS  1]}}
  bind . <Key-F3> {if {$lastX3DOM != ""} {openX3DOM $lastX3DOM}}
  bind . <Key-F6> {openMultiFile 0}
  bind . <Key-F7> {if {$lastXLS1  != ""} {set lastXLS1 [openXLS $lastXLS1 1]}}
  bind . <Key-F12> {if {$lastX3DOM != "" && [file exists $lastX3DOM]} {exec $editorCmd [file nativename $lastX3DOM] &}}

  bind . <MouseWheel> {[$fout.text component text] yview scroll [expr {-%D/30}] units}
  bind . <Up>     {[$fout.text component text] yview scroll -1 units}
  bind . <Down>   {[$fout.text component text] yview scroll  1 units}
  bind . <Left>   {[$fout.text component text] xview scroll -10 units}
  bind . <Right>  {[$fout.text component text] xview scroll  10 units}
  bind . <Prior>  {[$fout.text component text] yview scroll -30 units}
  bind . <Next>   {[$fout.text component text] yview scroll  30 units}
  bind . <Home>   {[$fout.text component text] yview scroll -100000 units}
  bind . <End>    {[$fout.text component text] yview scroll  100000 units}
}

#-------------------------------------------------------------------------------
# buttons and progress bar
proc guiButtons {} {
  global buttons ftrans mytemp nistVersion nprogBarEnts nprogBarFiles opt wdir

  set ftrans [frame .ftrans1 -bd 2 -background "#F0F0F0"]
  set butstr "Spreadsheet"
  if {$opt(XLSCSV) == "CSV"} {set butstr "CSV Files"}
  set buttons(genExcel) [ttk::button $ftrans.generate1 -text "Generate $butstr" -padding 4 -state disabled -command {
    saveState
    if {![info exists localNameList]} {
      set localName [getFirstFile]
      if {$localName != ""} {
        set localNameList [list $localName]
        genExcel
      }
    } elseif {[llength $localNameList] == 1} {
      genExcel
    } else {
      openMultiFile 2
    }
  }]
  pack $ftrans.generate1 -side left -padx 10

# NIST logo and icon
  if {$nistVersion} {
    catch {
      set l3 [label $ftrans.l3 -relief flat -bd 0]
      $l3 config -image [image create photo -file [file join $wdir images nist.gif]]
      pack $l3 -side right -padx 10
      bind $l3 <ButtonRelease-1> {openURL https://www.nist.gov}
      tooltip::tooltip $l3 "Click here to learn more about NIST"
    }
    catch {[file copy -force -- [file join $wdir images NIST.ico] [file join $mytemp NIST.ico]]}
  }

  pack $ftrans -side top -padx 10 -pady 10 -fill x

# progress bars
  set fbar [frame .fbar -bd 2 -background "#F0F0F0"]
  set nprogBarEnts 0
  set buttons(pgb) [ttk::progressbar $fbar.pgb -mode determinate -variable nprogBarEnts]
  pack $fbar.pgb -side top -padx 10 -fill x

  set nprogBarFiles 0
  set buttons(pgb1) [ttk::progressbar $fbar.pgb1 -mode determinate -variable nprogBarFiles]
  pack forget $buttons(pgb1)
  pack $fbar -side bottom -padx 10 -pady {0 10} -fill x

# NIST icon bitmap
  if {$nistVersion} {
    catch {wm iconbitmap . -default [file join $wdir images NIST.ico]}
  }
}

#-------------------------------------------------------------------------------
# status tab
proc guiStatusTab {} {
  global fout nb outputWin statusFont tcl_platform wout

  set wout [ttk::panedwindow $nb.status -orient horizontal]
  $nb add $wout -text " Status " -padding 2
  set fout [frame $wout.fout -bd 2 -relief sunken -background "#E0DFE3"]

  set outputWin [iwidgets::messagebox $fout.text -maxlines 500000 -hscrollmode dynamic -vscrollmode dynamic -background white]
  pack $fout.text -side top -fill both -expand true
  pack $fout -side top -fill both -expand true

  $outputWin type add black -foreground black -background white
  $outputWin type add red -foreground "#bb0000" -background white
  $outputWin type add green -foreground "#00aa00" -background white
  $outputWin type add magenta -foreground "#990099" -background white
  $outputWin type add cyan -foreground "#00dddd" -background white
  $outputWin type add blue -foreground blue -background white
  $outputWin type add error -foreground black -background "#ffff99"
  $outputWin type add syntax -foreground black -background "#ff9999"

# windows 7 or greater
  if {$tcl_platform(osVersion) >= 6.0} {
    if {![info exists statusFont]} {
      set statusFont [$outputWin type cget black -font]
    }
    if {[string first "Courier" $statusFont] != -1} {
      regsub "Courier" $statusFont "Consolas" statusFont
      regsub "120" $statusFont "140" statusFont
      saveState
    }
  }

  if {[info exists statusFont]} {
    foreach typ {black red green magenta cyan blue error syntax} {$outputWin type configure $typ -font $statusFont}
  }

  bind . <Key-F9> {
    set statusFont [$outputWin type cget black -font]
    for {set i 210} {$i >= 100} {incr i -10} {regsub -all $i $statusFont [expr {$i+10}] statusFont}
    foreach typ {black red green magenta cyan blue error syntax} {$outputWin type configure $typ -font $statusFont}
  }
  bind . <Control-KeyPress-=> {
    set statusFont [$outputWin type cget black -font]
    for {set i 210} {$i >= 100} {incr i -10} {regsub -all $i $statusFont [expr {$i+10}] statusFont}
    foreach typ {black red green magenta cyan blue error syntax} {$outputWin type configure $typ -font $statusFont}
  }

  bind . <Key-F8> {
    set statusFont [$outputWin type cget black -font]
    for {set i 110} {$i <= 220} {incr i 10} {regsub -all $i $statusFont [expr {$i-10}] statusFont}
    foreach typ {black red green magenta cyan blue error syntax} {$outputWin type configure $typ -font $statusFont}
  }
  bind . <Control-KeyPress--> {
    set statusFont [$outputWin type cget black -font]
    for {set i 110} {$i <= 220} {incr i 10} {regsub -all $i $statusFont [expr {$i-10}] statusFont}
    foreach typ {black red green magenta cyan blue error syntax} {$outputWin type configure $typ -font $statusFont}
  }
}

#-------------------------------------------------------------------------------
# file menu
proc guiFileMenu {} {
  global File lastX3DOM lastXLS lastXLS1 openFileList

  $File add command -label "Open STEP File(s)..." -accelerator "Ctrl+O" -command openFile
  $File add command -label "Open Multiple STEP Files in a Directory..." -accelerator "Ctrl+D, F6" -command {openMultiFile}
  set newFileList {}
  foreach fo $openFileList {if {[file exists $fo]} {lappend newFileList $fo}}
  set openFileList $newFileList

  set llen [llength $openFileList]
  $File add separator
  if {$llen > 0} {
    for {set fi 0} {$fi < $llen} {incr fi} {
      set fo [lindex $openFileList $fi]
      if {$fi != 0} {
        $File add command -label [truncFileName [file nativename $fo] 1] -command [list openFile $fo]
      } else {
        $File add command -label [truncFileName [file nativename $fo] 1] -command [list openFile $fo] -accelerator "F1"
      }
    }
  }
  $File add separator
  $File add command -label "Open Spreadsheet" -accelerator "F2" -command {if {$lastXLS != ""}   {set lastXLS [openXLS $lastXLS 1]}}
  $File add command -label "Open View File"   -accelerator "F3" -command {if {$lastX3DOM != ""} {openX3DOM $lastX3DOM}}
  $File add command -label "Open Multiple File Summary Spreadsheet" -accelerator "F7" -command {if {$lastXLS1 != ""} {set lastXLS1 [openXLS $lastXLS1 1]}}
  $File add command -label "Exit" -accelerator "Ctrl+Q" -command exit
}

#-------------------------------------------------------------------------------
# options tab, process and report
proc guiProcessAndReports {} {
  global allNone buttons cb entCategory fopt fopta nb opt

  set cb 0
  set wopt [ttk::panedwindow $nb.options -orient horizontal]
  $nb add $wopt -text " Options " -padding 2
  set fopt [frame $wopt.fopt -bd 2 -relief sunken]
  set fopta [ttk::labelframe $fopt.a -text " Process "]

# option to process user-defined entities
  guiUserDefinedEntities

  set fopta1 [frame $fopta.1 -bd 0]
  foreach item {{" Common"         opt(PR_STEP_COMM)} \
                {" Presentation"   opt(PR_STEP_PRES)} \
                {" Representation" opt(PR_STEP_REPR)} \
                {" Tolerance"      opt(PR_STEP_TOLR)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta1.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    set tt [string range $idx 3 end]
    if {[info exists entCategory($tt)]} {
      set mostSome "most"
      if {$tt == "PR_STEP_TOLR"} {set mostSome "some"}
      set ttmsg "There are [llength $entCategory($tt)] [string trim [lindex $item 0]] entities.  These entities are found in $mostSome STEP APs."
      if {$tt != "PR_STEP_COMM"} {append ttmsg "\nEntities marked with an asterisk (*) are only in AP242.  Some are only in AP242 edition 2."}
      append ttmsg "\nSee Help > Supported STEP APs  and  Websites > STEP Format and Schemas\n\n"
      if {$tt != "PR_STEP_COMM"} {
        set ttmsg [guiToolTip $ttmsg $tt]
      } else {
        append ttmsg "All AP-specific entities from APs other than AP203, AP214, and AP242 are always processed,\nincluding AP209, AP210, AP238, and AP239.\n\nThe entity categories are used to group and color-code entities on the File Summary worksheet.\n\nSee Help > User Guide (section 4.4.2)"
      }
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta1 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta2 [frame $fopta.2 -bd 0]
  foreach item {{" Measure"      opt(PR_STEP_QUAN)} \
                {" Shape Aspect" opt(PR_STEP_SHAP)} \
                {" Geometry"     opt(PR_STEP_GEOM)} \
                {" Coordinates"  opt(PR_STEP_CPNT)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta2.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    set tt [string range $idx 3 end]
    if {[info exists entCategory($tt)]} {
      set ttmsg "There are [llength $entCategory($tt)] [string trim [lindex $item 0]] entities."
      if {$tt != "PR_STEP_CPNT"} {
        append ttmsg "  These entities are found in most STEP APs.\nEntities marked with an asterisk (*) are only in AP242.  Some are only in AP242 edition 2."
      } elseif {$tt == "PR_STEP_CPNT"} {
        append ttmsg "  coordinates_list is only in AP242."
      }
      append ttmsg "\nSee Help > Supported STEP APs  and  Websites > STEP Format and Schemas\n\n"
      set ttmsg [guiToolTip $ttmsg $tt]
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta2 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta3 [frame $fopta.3 -bd 0]
  foreach item {{" AP242"      opt(PR_STEP_AP242)} \
                {" Features"   opt(PR_STEP_FEAT)} \
                {" Composites" opt(PR_STEP_COMP)} \
                {" Kinematics" opt(PR_STEP_KINE)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta3.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    set tt [string range $idx 3 end]
    if {[info exists entCategory($tt)]} {
      set ttmsg "There are [llength $entCategory($tt)] [string trim [lindex $item 0]] entities."
      if {$tt != "PR_STEP_AP242"} {
        append ttmsg "  These entities are found in some STEP APs.\nEntities marked with an asterisk (*) are only in AP242.  Some are only in AP242 edition 2."
        append ttmsg "\nSee Help > Supported STEP APs  and  Websites > STEP Format and Schemas\n\n"
        set ttmsg [guiToolTip $ttmsg $tt]
      } else {
        append ttmsg "\n\nThese entities are only in AP242 and not in AP203 or AP214.  Some entities are only in AP242 Edition 2."
        append ttmsg "\nSee Websites > AP242"
        append ttmsg "\nSee Help > Supported STEP APs  and  Websites > STEP Format and Schemas"
      }
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta3 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta4 [frame $fopta.4 -bd 0]
  set anbut [list {"All" 0} {"None" 1} {"For Analysis" 2} {"For Views" 3}]
  foreach item $anbut {
    set bn "allNone[lindex $item 1]"
    set buttons($bn) [ttk::radiobutton $fopta4.$cb -variable allNone -text [lindex $item 0] -value [lindex $item 1] \
      -command {
        if {$allNone == 0} {
          foreach item [array names opt] {
            if {[string first "PR_STEP" $item] == 0 && $item != "PR_STEP_GEOM" && $item != "PR_STEP_CPNT"} {set opt($item) 1}
          }
        } elseif {$allNone == 1} {
          foreach item [array names opt] {if {[string first "PR_STEP" $item] == 0} {set opt($item) 0}}
          foreach item {VIZBRP VIZFEA VIZPMI VIZTPG PMISEM PMIGRF VALPROP INVERSE PR_USER} {set opt($item) 0}
          set opt(PR_STEP_COMM) 1
        } elseif {$allNone == 2} {
          foreach item {PMISEM PMIGRF VALPROP} {set opt($item) 1}
        } elseif {$allNone == 3} {
          foreach item {VIZBRP VIZFEA VIZPMI VIZTPG} {set opt($item) 1}
        }
        checkValues
      }]
    pack $buttons($bn) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  catch {
    tooltip::tooltip $buttons(allNone0) "Selects most Process categories\nSee Help > User Guide (section 4.4.2)"
    tooltip::tooltip $buttons(allNone1) "Deselects most Process categories and all Analyze and View options\nSee Help > User Guide (section 4.4.2)"
    tooltip::tooltip $buttons(allNone2) "Selects all Analyze options and associated Process categories\nSee Help > User Guide (section 5)"
    tooltip::tooltip $buttons(allNone3) "Selects all View options and associated Process categories\nSee Help > User Guide (section 7)"
  }
  pack $fopta4 -side left -anchor w -pady 0 -padx 0 -fill y
  pack $fopta -side top -anchor w -pady {5 2} -padx 10 -fill both

#-------------------------------------------------------------------------------
# report
  set foptrv [frame $fopt.rv -bd 0]
  set foptd [ttk::labelframe $foptrv.1 -text " Analyze "]

  set foptd1 [frame $foptd.1 -bd 0]
  foreach item {{" AP242 PMI Representation (Semantic PMI)" opt(PMISEM)}} {
  regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptd1.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptd1 -side top -anchor w -pady 0 -padx 0 -fill y

  set foptd2 [frame $foptd.2 -bd 0]
  foreach item {{" Only Dimensions" opt(PMISEMDIM)}} {
  regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptd2.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptd2 -side top -anchor w -pady 0 -padx 25 -fill y

  set foptd3 [frame $foptd.3 -bd 0]
  foreach item {{" PMI Presentation (Graphical PMI)"  opt(PMIGRF)} \
                {" Validation Properties"             opt(VALPROP)}} {
  regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptd3.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptd3 -side top -anchor w -pady 0 -padx 0 -fill y

  pack $foptd -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true
  catch {
    tooltip::tooltip $buttons(optPMISEM)  "The analysis of PMI Representation information is shown on\ndimension, tolerance, datum target, datum, and hole (AP242e2)\nentities.  Semantic PMI is found mainly in STEP AP242 files.\n\nSee Help > Analyze > PMI Representation\nSee Help > User Guide (section 5.1)\nSee Help > Syntax Errors\nSee Examples > Spreadsheet - PMI Representation\nSee Examples > Sample STEP Files\nSee Websites > AP242"
    tooltip::tooltip $buttons(optPMIGRF)  "The analysis of PMI Presentation information is\nshown on 'annotation occurrence' entities.\n\nSee Help > Analyze > PMI Presentation\nSee Help > User Guide (section 5.2)\nSee Help > Syntax Errors\nSee Examples > PMI Presentation, Validation Properties\nSee Examples > View Part with PMI\nSee Examples > AP242 Tessellated Part with PMI\nSee Examples > Sample STEP Files"
    tooltip::tooltip $buttons(optVALPROP) "The analysis of Validation Properties and other properties\nis shown on the 'property_definition' entity.\n\nSee Help > Analyze > Validation Properties\nSee Help > User Guide (section 5.3)\nSee Help > Syntax Errors\nSee Examples > PMI Presentation, Validation Properties"
    tooltip::tooltip $buttons(optPMISEMDIM)  "Analyze only dimensional tolerances and no\ngeometric tolerances, datums, or datum targets."
  }

#-------------------------------------------------------------------------------
# view
  set foptv [ttk::labelframe $foptrv.9 -text " View "]
  set foptv20 [frame $foptv.20 -bd 0]
  foreach item {{" Part Geometry" opt(VIZBRP)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptv20.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv20 -side top -anchor w -pady 0 -padx 0 -fill y

  set foptv3 [frame $foptv.3 -bd 0]
  foreach item {{" Graphical PMI" opt(VIZPMI)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptv3.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv3 -side top -anchor w -pady 0 -padx 0 -fill y

  set foptv4 [frame $foptv.4 -bd 0]
  set buttons(linecolor) [label $foptv4.l3 -text "PMI Color:"]
  pack $foptv4.l3 -side left -anchor w -padx 0 -pady 0 -ipady 0
  set gpmiColorVal {{"By View" 3} {"Random" 2} {"From File" 0} {"Black" 1}}
  foreach item $gpmiColorVal {
    set bn "gpmiColor[lindex $item 1]"
    set buttons($bn) [ttk::radiobutton $foptv4.$cb -variable opt(gpmiColor) -text [lindex $item 0] -value [lindex $item 1]]
    pack $buttons($bn) -side left -anchor w -padx 2 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv4 -side top -anchor w -pady 0 -padx 25 -fill y

  set foptv6 [frame $foptv.6 -bd 0]
  foreach item {{" AP242 Tessellated Part Geometry" opt(VIZTPG)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptv6.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  foreach item {{"Generate Wireframe" opt(VIZTPGMSH)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptv6.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 8 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv6 -side top -anchor w -pady 0 -padx 0 -fill y

  set foptv7 [frame $foptv.7 -bd 0]
  foreach item {{" AP209 Finite Element Model" opt(VIZFEA)} \
                {"Boundary conditions" opt(VIZFEABC)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptv7.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv7 -side top -anchor w -pady 0 -padx 0 -fill y

  set foptv8 [frame $foptv.8 -bd 0]
  foreach item {{"Loads" opt(VIZFEALV)} \
                {"Scale loads   " opt(VIZFEALVS)} \
                {"Displacements" opt(VIZFEADS)} \
                {"No vector tail" opt(VIZFEADSntail)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptv8.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 2 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv8 -side top -anchor w -pady 0 -padx 25 -fill y

  pack $foptv -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true
  pack $foptrv -side top -anchor w -pady 0 -fill x
  catch {
    tooltip::tooltip $buttons(optVIZBRP) "Views are shown in the default web browser.  Older versions of web\nbrowsers are not supported.  An Internet connection is required to\nshow View files.\n\nViews can be generated without generating a spreadsheet or CSV\nfiles.  See the Output Format option below.\n\nMost boundary representation (b-rep) part geometry can be viewed.\nMultiple and overriding part colors are ignored.\nSupplemental geometry and holes (AP242e2) are also shown.\nViews for very large STEP files might take 10-20 minutes to generate.\n\nSee Help > View > Part Geometry\nSee Help > View > Supplemental Geometry\nSee Examples > View Part with PMI\nSee Websites > STEP File Viewers (for other part geometry viewers)"
    tooltip::tooltip $buttons(optVIZPMI) "Graphical PMI is supported in AP242, AP203, and AP214 files.\n\nSee Help > View > Graphical PMI\nSee Help > User Guide (section 7.1.1)\nSee Examples > View Part with PMI\nSee Examples > AP242 Tessellated Part with PMI\nSee Examples > Sample STEP Files"
    tooltip::tooltip $buttons(optVIZTPG) "** Parts in an assembly might have the wrong\nposition and orientation or be missing. **\n\nTessellated edges (lines) are also shown.  Faces\nin tessellated shells are outlined in black.\n\nSee Help > View > AP242 Tessellated Part Geometry\nSee Help > User Guide (section 7.1.2, 7.1.3)\nSee Examples > AP242 Tessellated Part with PMI"
    tooltip::tooltip $buttons(optVIZTPGMSH) "Generate a wireframe mesh based on the tessellated faces and surfaces."
    tooltip::tooltip $buttons(optVIZFEALVS) "The length of load vectors can be scaled by their magnitude.\nLoad vectors are always colored by their magnitude."
    tooltip::tooltip $buttons(optVIZFEADSntail) "The length of displacement vectors with a tail are scaled by\ntheir magnitude.  Vectors without a tail are not.\nDisplacement vectors are always colored by their magnitude.\nLoad vectors always have a tail."
    tooltip::tooltip $foptv4 "For 'By View' PMI colors, each Saved View is assigned a different color.\nIf there are one or no Saved Views, then 'Random' PMI colors are used.\n\nFor 'Random' PMI colors, each 'annotation occurrence' is assigned a\ndifferent color to help differentiate one from another."
    set tt "FEM nodes, elements, boundary conditions,\nloads, and displacements are shown.\n\nSee Help > View > AP209 Finite Element Model\nSee Help > User Guide (section 7.1.4)\nSee Examples > AP209 Finite Element Model"
    tooltip::tooltip $foptv7 $tt
    tooltip::tooltip $foptv8 $tt
    #tooltip::tooltip $buttons(optVIZPMIVP) "PMI Viewpoints are experimental.\n\nViewpoints usually have the correct orientation but are not centered.\nUse pan and zoom to center the PMI."
  }
}

#-------------------------------------------------------------------------------
# user-defined list of entities
proc guiUserDefinedEntities {} {
  global buttons cb opt fileDir fopta userEntityFile userEntityList

  set fopta6 [frame $fopta.6 -bd 0]
  foreach item {{" User-Defined List: " opt(PR_USER)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta6.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }

  set buttons(userentity) [ttk::entry $fopta6.entry -width 50 -textvariable userEntityFile]
  pack $fopta6.entry -side left -anchor w

  set buttons(userentityopen) [ttk::button $fopta6.$cb -text " Browse " -command {
    set typelist {{"All Files" {*}}}
    set uef [tk_getOpenFile -title "Select File of STEP Entities" -filetypes $typelist -initialdir $fileDir]
    if {$uef != "" && [file isfile $uef]} {
      set userEntityFile [file nativename $uef]
      outputMsg "User-defined STEP list: [truncFileName $userEntityFile]" blue
      set fileent [open $userEntityFile r]
      set userEntityList {}
      while {[gets $fileent line] != -1} {
        set line [split [string trim $line] " "]
        foreach ent1 $line {lappend userEntityList $ent1}
      }
      close $fileent
      set llist [llength $userEntityList]
      if {$llist > 0} {
        outputMsg " ($llist) $userEntityList"
      } else {
        outputMsg "File does not contain any STEP entity names" red
        set opt(PR_USER) 0
        checkValues
      }
      .tnb select .tnb.status
    }
    checkValues
  }]
  pack $fopta6.$cb -side left -anchor w -padx 10
  incr cb
  catch {
    foreach item {optPR_USER userentity userentityopen} {
      tooltip::tooltip $buttons($item) "A User-Defined List is a plain text file with one STEP entity name per line.\n\nThis allows for more control to process only the required entity types,\nrather than process the board categories of entities above.\n\nIt is also useful when processing large files that might crash the software."
    }
  }
  pack $fopta6 -side bottom -anchor w -pady 5 -padx 0 -fill y
}

#-------------------------------------------------------------------------------
# inverse relationships
proc guiInverse {} {
  global buttons cb fopt inverses opt

  set foptc [ttk::labelframe $fopt.3 -text " Inverse Relationships "]
  set txt " Show Inverses and Backwards References (Used In) for PMI, Shape Aspect, Representation, Tolerance, and more"

  regsub -all {[\(\)]} opt(INVERSE) "" idx
  set buttons($idx) [ttk::checkbutton $foptc.$cb -text $txt -variable opt(INVERSE) -command {checkValues}]
  pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
  incr cb

  pack $foptc -side top -anchor w -pady {5 2} -padx 10 -fill both
  set ttmsg "Inverse Relationships and Backwards References (Used In) are reported for some attributes for the following entities.\nInverse or Used In values are shown in additional columns highlighted in light blue and purple.\n\nSee Help > User Guide (section 4.4.5)\n\n"
  set lent ""
  set ttlen 0
  if {[info exists inverses]} {
    foreach item [lsort $inverses] {
    set ent [lindex [split $item " "] 0]
      if {$ent != $lent} {
        set str "[formatComplexEnt $ent]    "
        append ttmsg $str
        incr ttlen [string length $str]
        if {$ttlen > 140} {
          if {[string index $ttmsg end] != "\n"} {append ttmsg "\n"}
          set ttlen 0
        }
      }
      set lent $ent
    }
    catch {tooltip::tooltip $foptc $ttmsg}
  }
}

#-------------------------------------------------------------------------------
# open STEP file and output format
proc guiOpenSTEPFile {} {
  global appName appNames buttons cb developer dispApps dispCmds edmWhereRules edmWriteToFile eeWriteToFile
  global fopt foptf useXL xlInstalled

  set foptf [ttk::labelframe $fopt.f -text " Open STEP File in "]

  set buttons(appCombo) [ttk::combobox $foptf.spinbox -values $appNames -width 40]
  pack $foptf.spinbox -side left -anchor w -padx 7 -pady {0 3}
  bind $buttons(appCombo) <<ComboboxSelected>> {
    set appName [$buttons(appCombo) get]

# Jotne EDM Model Checker
    if {$developer} {
      catch {
        if {[string first "EDM Model Checker" $appName] == 0} {
          pack $buttons(edmWriteToFile) -side left -anchor w -padx 5
          pack $buttons(edmWhereRules) -side left -anchor w -padx 5
        } else {
          pack forget $buttons(edmWriteToFile)
          pack forget $buttons(edmWhereRules)
        }
      }
    }
# STEP Tools
    catch {
      if {[string first "Conformance Checker" $appName] != -1} {
        pack $buttons(eeWriteToFile) -side left -anchor w -padx 5
      } else {
        pack forget $buttons(eeWriteToFile)
      }
    }
# file tree view
    catch {
      if {$appName == "Tree View (for debugging)"} {
        pack $buttons(indentStyledItem) -side left -anchor w -padx 5
        pack $buttons(indentGeometry) -side left -anchor w -padx 5
      } else {
        pack forget $buttons(indentStyledItem)
        pack forget $buttons(indentGeometry)
      }
    }

# set the app command
    foreach cmd $dispCmds {
      if {$appName == $dispApps($cmd)} {
        set dispCmd $cmd
      }
    }

# put the app name at the top of the list
    for {set i 0} {$i < [llength $dispCmds]} {incr i} {
      if {$dispCmd == [lindex $dispCmds $i]} {
        set dispCmds [lreplace $dispCmds $i $i]
        set dispCmds [linsert $dispCmds 0 $dispCmd]
      }
    }
    set appNames {}
    foreach cmd $dispCmds {
      if {[info exists dispApps($cmd)]} {lappend appNames $dispApps($cmd)}
    }
    $foptf.spinbox configure -values $appNames
  }

  set buttons(appOpen) [ttk::button $foptf.$cb -text " Open " -state disabled -command {
    runOpenProgram
    saveState
  }]
  pack $foptf.$cb -side left -anchor w -padx {10 0} -pady {0 3}
  incr cb

# Jotne EDM Model Checker
  if {$developer} {
    foreach item $appNames {
      if {[string first "EDM Model Checker" $item] == 0} {
        foreach item {{" Write results to file" edmWriteToFile}} {
          regsub -all {[\(\)]} [lindex $item 1] "" idx
          set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
          pack forget $buttons($idx)
          incr cb
        }
        foreach item {{" Check rules" edmWhereRules}} {
          regsub -all {[\(\)]} [lindex $item 1] "" idx
          set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
          pack forget $buttons($idx)
          incr cb
        }
      }
    }
  }

# Express Engine
  if {[lsearch -glob $appNames "*Conformance Checker*"] != -1} {
    foreach item {{" Write results to file" eeWriteToFile}} {
      regsub -all {[\(\)]} [lindex $item 1] "" idx
      set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
      pack forget $buttons($idx)
      incr cb
    }
  }

# built-in file tree view
  if {[lsearch $appNames "Tree View (for debugging)"] != -1} {
    foreach item {{" Include Styled_item" indentStyledItem} \
                  {" Include Geometry" indentGeometry}} {
      regsub -all {[\(\)]} [lindex $item 1] "" idx
      set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] -variable opt([lindex $item 1]) -command {checkValues}]
      pack forget $buttons($idx)
      incr cb
    }
  }

  catch {tooltip::tooltip $foptf "This option is a convenient way to open a STEP file in other applications.\nThe pull-down menu contains some applications that can open a STEP\nfile such as STEP viewers and browsers, however, only if they are installed in\ntheir default location.\n\nSee Help > Open STEP File in Apps\nSee Websites > STEP File Viewers\n\nThe 'Tree View (for debugging)' option rearranges and indents the\nentities to show the hierarchy of information in a STEP file.  The 'tree view'\nfile (myfile-sfa.txt) is written to the same directory as the STEP file or to the\nsame user-defined directory specified in the Spreadsheet tab.  Including\nGeometry or Styled_item can make the 'tree view' file very large.  The\n'tree view' might not process /*comments*/ in a STEP file correctly.\n\nThe 'Default STEP Viewer' option opens the STEP file in whatever\napplication is associated with STEP (.stp, .step, .p21) files.\n\nUse F5 to open the STEP file in a text editor."}
  pack $foptf -side top -anchor w -pady {5 2} -padx 10 -fill both

# output format
  set foptk [ttk::labelframe $fopt.k -text " Output Format "]

  set txt " Open Output files after they have been generated"
  regsub -all {[\(\)]} opt(XL_OPEN) "" idx
  set buttons($idx) [ttk::checkbutton $foptk.$cb -text $txt -variable opt(XL_OPEN)]
  pack $buttons($idx) -side bottom -anchor w -padx 5 -pady 0 -ipady 0
  incr cb

# checkbuttons are used for pseudo-radiobuttons
  foreach item {{" Spreadsheet" ofExcel} \
                {" CSV Files" ofCSV} \
                {" View Only" ofNone}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptk.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {
      if {![info exists useXL]} {set useXL 1}
      if {[info exists xlInstalled]} {
        if {!$xlInstalled} {set useXL 0}
      } else {
        set xlInstalled 1
      }

      if {$ofNone && $opt(XLSCSV) != "None"} {
        set ofExcel 0
        set ofCSV 0
        set opt(XLSCSV) "None"
        if {$useXL && $xlInstalled} {$buttons(ofExcel) configure -state normal}
      }
      if {$ofExcel && $opt(XLSCSV) != "Excel"} {
        set ofNone 0
        if {$useXL} {
          set ofCSV 0
          set opt(XLSCSV) "Excel"
        } else {
          set ofExcel 0
          set ofCSV 1
          set opt(XLSCSV) "CSV"
        }
      }
      if {$ofCSV} {
        if {$useXL} {
          set ofExcel 1
          $buttons(ofExcel) configure -state disabled
        }
        if {$opt(XLSCSV) != "CSV"} {
          set ofNone 0
          set opt(XLSCSV) "CSV"
        }
      } elseif {$xlInstalled} {
        $buttons(ofExcel) configure -state normal
      }
      if {!$ofExcel && !$ofCSV && !$ofNone} {
        if {$useXL} {
          set ofExcel 1
          set opt(XLSCSV) "Excel"
          $buttons(ofExcel) configure -state normal
        } else {
          set ofCSV 1
          set opt(XLSCSV) "CSV"
          $buttons(ofExcel) configure -state disabled
        }
      }
      checkValues
    }]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }

  pack $foptk -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {tooltip::tooltip $foptk "If Excel is installed, then Spreadsheets and CSV files can be\ngenerated.  If CSV Files is selected, the Spreadsheet is also\ngenerated.\n\nIf Excel is not installed, only CSV files can be generated.\nOptions for Analyze and Inverse Relationships are disabled.\n\nCSV files do not contain any cell colors, comments, or links.\nGD&T symbols will look correct only with Excel 2016 or newer.\n\nView Only does not generate any Spreadsheets or CSV files.\nAll options except View are disabled.\n\nIf Output files are not opened after they have been generated,\nuse F2 to open a Spreadsheet and F3 to open a View.  Use F7\nto open the File Summary Spreadsheet when processing\nmultiple files.\n\nIf possible, existing Spreadsheets, CSV files, and View files are\nalways overwritten by new files.\n\nSee Help > User Guide (section 4.4.1)"}

# log file
  set foptm [ttk::labelframe $fopt.m -text " Log File "]
  set txt " Generate a Log File of the text in the Status tab"
  regsub -all {[\(\)]} opt(LOGFILE) "" idx
  set buttons($idx) [ttk::checkbutton $foptm.$cb -text $txt -variable opt(LOGFILE)]
  pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
  pack $foptm -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {tooltip::tooltip $buttons(optLOGFILE)  "The Log file is written to myfile-sfa.log\nUse F4 to open the Log file.\n\nSee Help > Syntax Errors"}
}

#-------------------------------------------------------------------------------
# spreadsheet tab
proc guiSpreadsheet {} {
  global buttons cb developer excelVersion extXLS fileDir fxls mydocs nb opt userWriteDir userXLSFile writeDir

  set wxls [ttk::panedwindow $nb.xls -orient horizontal]
  $nb add $wxls -text " Spreadsheet " -padding 2
  set fxls [frame $wxls.fxls -bd 2 -relief sunken]

  set fxlsz [ttk::labelframe $fxls.z -text " Tables "]
  foreach item {{" Generate Tables for Sorting and Filtering" opt(XL_SORT)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fxlsz.$cb -text [lindex $item 0] -variable [lindex $item 1]]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsz -side top -anchor w -pady {5 2} -padx 10 -fill both
  set msg "Worksheets can be sorted by column values.\nThe worksheet with Properties is always sorted.\n\nSee Help > User Guide (section 4.5.1)"
  catch {tooltip::tooltip $fxlsz $msg}

  set fxlsa [ttk::labelframe $fxls.a -text " Number Format "]
  foreach item {{" Do not round Real Numbers" opt(XL_FPREC)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fxlsa.$cb -text [lindex $item 0] -variable [lindex $item 1]]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsa -side top -anchor w -pady {5 2} -padx 10 -fill both
  set msg "Excel rounds real numbers if there are more than 11 characters in the number string.  For example,\nthe number 0.1249999999997 in the STEP file is shown as 0.125\n\nClicking in a cell with a rounded number shows all of the digits in the formula bar.\n\nThis option shows most real numbers exactly as they appear in the STEP file.  This applies\nonly to single real numbers.  Lists of real numbers, such as cartesian point coordinates, are\nalways shown exactly as they appear in the STEP file.\n\nSee Help > User Guide (section 4.5.2)"
  catch {tooltip::tooltip $fxlsa $msg}

  set fxlsb [ttk::labelframe $fxls.b -text " Maximum Rows for any worksheet"]
  set rlimit {{" 100" 103} {" 500" 503} {" 1000" 1003} {" 5000" 5003} {" 10000" 10003} {" 50000" 50003} {" 100000" 100003} {" Maximum" 1048576}}
  if {$excelVersion < 12} {
    set rlimit [lrange $rlimit 0 5]
    lappend rlimit {" Maximum" 65536}
  }
  foreach item $rlimit {
    pack [ttk::radiobutton $fxlsb.$cb -variable opt(XL_ROWLIM) -text [lindex $item 0] -value [lindex $item 1]] -side left -anchor n -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsb -side top -anchor w -pady 5 -padx 10 -fill both
  set msg "This option limits the number of rows (entities) written to any one worksheet or CSV file.\nIf the maximum number of rows is exceeded, the number of entities processed will be\nreported as, for example, 'property_definition (100 of 147)'.\n\nFor large STEP files, setting a low maximum can speed up processing at the expense of\nnot processing all of the entities.  This is useful when processing Geometry entities.\n\nSyntax Errors might be missed if some entities are not processed due to a small value\nof maximum rows.  Maximum rows does not affect generating Views.  The maximum\nnumber of rows depends on the version of Excel.\n\nSee Help > User Guide (section 4.5.3)"
  catch {tooltip::tooltip $fxlsb $msg}

  set fxlsd [ttk::labelframe $fxls.d -text " Write Spreadsheet to "]
  set buttons(fileDir) [ttk::radiobutton $fxlsd.$cb -text " Same directory as the STEP file" -variable opt(writeDirType) -value 0 -command checkValues]
  pack $fxlsd.$cb -side top -anchor w -padx 5 -pady 2
  incr cb

  set fxls1 [frame $fxlsd.1]
  ttk::radiobutton $fxls1.$cb -text " User-defined directory:  " -variable opt(writeDirType) -value 2 -command {
    checkValues
    if {[file exists $userWriteDir] && [file isdirectory $userWriteDir]} {
      set writeDir $userWriteDir
    } else {
      set userWriteDir $mydocs
      tk_messageBox -type ok -icon error -title "Invalid Directory" \
        -message "The user-defined directory to write the Spreadsheet to is not valid.\nIt has been set to $userWriteDir"
    }
    focus $buttons(userdir)
  }
  pack $fxls1.$cb -side left -anchor w -padx {5 0}
  catch {tooltip::tooltip $fxls1.$cb "Use this option when the directory containing the STEP file is\nprotected (read-only) and the Spreadsheet cannot be written to it."}
  incr cb

  set buttons(userentry) [ttk::entry $fxls1.entry -width 38 -textvariable userWriteDir]
  pack $fxls1.entry -side left -anchor w -pady 2
  set buttons(userdir) [ttk::button $fxls1.button -text " Browse " -command {
    set uwd [tk_chooseDirectory -title "Select directory"]
    if {[file isdirectory $uwd]} {
      set userWriteDir $uwd
      set writeDir $userWriteDir
    }
  }]
  pack $fxls1.button -side left -anchor w -padx 10 -pady 2
  pack $fxls1 -side top -anchor w

  set fxls2 [frame $fxlsd.2]
  ttk::radiobutton $fxls2.$cb -text " User-defined file name:  " -variable opt(writeDirType) -value 1 -command {
    checkValues
    focus $buttons(userfile)
  }
  pack $fxls2.$cb -side left -anchor w -padx {5 0}
  incr cb

  set buttons(userentry1) [ttk::entry $fxls2.entry -width 35 -textvariable userXLSFile]
  pack $fxls2.entry -side left -anchor w -pady 2
  set buttons(userfile) [ttk::button $fxls2.button -text " Browse " -command {
    if {$extXLS == "xls"} {
      set typelist {{"Excel Files" {".xls"}}}
    } else {
      set typelist {{"Excel Files" {".xlsx"}}}
    }
    set uxf [tk_getSaveFile -title "Save Spreadsheet to" -filetypes $typelist -initialdir $fileDir -defaultextension ".$extXLS"]
    if {$uxf != ""} {
      set userXLSFile $uxf
    } else {
      errorMsg "No file selected"
    }
  }]
  pack $fxls2.button -side left -anchor w -padx 10 -pady 2
  pack $fxls2 -side top -anchor w
  pack $fxlsd -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {tooltip::tooltip $fxlsd "If possible, existing Spreadsheets, CSV files, and View files are\nalways overwritten by new files."}

  set fxlsc [ttk::labelframe $fxls.c -text " Other "]
  set items [list {" On File Summary worksheet, create links to STEP files and spreadsheets (see File > Open Multiple ...)" opt(XL_LINK1)}]
  foreach item $items {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fxlsc.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
  }
  pack $fxlsc -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {tooltip::tooltip $fxlsc.$cb "Deselecting this option is useful when sharing a Spreadsheet with another user."}
  incr cb

  if {$developer} {
    set fxlsx [ttk::labelframe $fxls.x -text " Debug "]
    foreach item {{" Analysis" opt(DEBUG1)} \
                  {" Inverses" opt(DEBUGINV)}} {
      regsub -all {[\(\)]} [lindex $item 1] "" idx
      set buttons($idx) [ttk::checkbutton $fxlsx.$cb -text [lindex $item 0] -variable [lindex $item 1]]
      pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
      incr cb
    }
    pack $fxlsx -side top -anchor w -pady {5 2} -padx 10 -fill both
    catch {tooltip::tooltip $fxlsx "These features are only available on NIST computers."}
  }

  pack $fxls -side top -fill both -expand true -anchor nw
}

#-------------------------------------------------------------------------------
# help menu
proc guiHelpMenu {} {
  global contact defaultColor Examples excelVersion Help ifcsvrDir mytemp nistVersion opt stepAPs virtualDir

  $Help add command -label "User Guide" -command {showFileURL UserGuide}
  $Help add command -label "What's New" -command {whatsNew}
  $Help add command -label "Changelog"  -command {showFileURL Changelog}

  if {$nistVersion} {
    $Help add command -label "Check for Update" -command {
      .tnb select .tnb.status
      set lastupgrade [expr {round(([clock seconds] - $upgrade)/86400.)}]
      outputMsg "The last check for an update was $lastupgrade days ago." red
      set url "https://concrete.nist.gov/cgi-bin/ctv/sfa_upgrade.cgi?version=[getVersion]&auto=-$lastupgrade"
      openURL $url
      set upgrade [clock seconds]
      saveState
    }
  }

  $Help add separator
  $Help add command -label "Overview" -command {
outputMsg "\nOverview -------------------------------------------------------------------" blue
outputMsg "The STEP File Analyzer and Viewer (SFA) reads a STEP file and generates an Excel spreadsheet or
CSV files.  One worksheet or CSV file is generated for each entity type in the STEP file.  Each
worksheet or CSV file lists every entity instance and its attributes.  The types of entities that
are Processed can be selected in the Options tab.  Other options are available that add to or
modify the information written to the spreadsheet or CSV files.

Spreadsheets or CSV files can be selected in the Options tab.  CSV files are automatically
generated if Excel is not installed.  To generate a spreadsheet or CSV files, select a STEP file
from the File menu above and click the Generate button below.  If possible, existing spreadsheets,
CSV files, and view files are always overwritten by new files.

For spreadsheets, a Summary worksheet shows the Count of each entity.  Links on the Summary and
entity worksheets can be used to navigate to other worksheets.

Views can also be generated with or without generating a spreadsheet or CSV files.

Multiple STEP files can be selected or an entire directory structure of STEP files can also be
processed from the File menu. If multiple STEP files are translated, then a separate File Summary
spreadsheet is also generated.

Tooltip help is available for the selections in the tabs.  Hold the mouse over text in the tabs
until a tooltip appears."
    .tnb select .tnb.status
  }

# options help
  $Help add command -label "Options" -command {
outputMsg "\nOptions --------------------------------------------------------------------" blue
outputMsg "See Help > User Guide (sections 4.4, 4.5, 5, and 7)
process: Select which types of entities are processed.  The tooltip help lists all the entities
associated with that type.  Selectively process only the entities or views relevant to your
analysis.  Entity types and views can also be selected with the All, None, For Analysis, and
For Views buttons.

PMI Representation Analysis: Dimensional tolerances, geometric tolerances, and datum features
are reported on various entities indicated by PMI Representation on the Summary worksheet.

PMI Presentation Analysis: Geometric entities used for PMI Presentation annotations are reported.
Associated Saved Views, Validation Properties, and Geometry are also reported.

Validation Properties Analysis: Geometric, assembly, PMI, annotation, attribute, and tessellated
validation properties are reported.

View: Graphical PMI annotations, tessellated part geometry in AP242 files, and AP209 finite
element models can be viewed in a web browser.

Inverse Relationships: For some entities, Inverse relationships and backwards references
(Used In) are shown on the worksheets.

Output Format: Generate Excel spreadsheets, CSV files, or only Views.  If Excel is not installed,
CSV files are automatically generated.  Some options are not available with CSV files.  The Views
Only option does not generate spreadsheets or CSV files.

Table: Generate tables for each spreadsheet to facilitate sorting and filtering (Spreadsheet tab).

Number Format: Option to not round real numbers.

Maximum Rows: The maximum number of rows for any worksheet can be set lower than the normal
limits for Excel."
    .tnb select .tnb.status
  }

# open Function Keys help
  $Help add command -label "Function Keys" -command {
outputMsg "\nFunction Keys --------------------------------------------------------------" blue
outputMsg "Function keys can be used as shortcuts for several commands:

F1 - Generate Spreadsheets and/or Views from the current or last STEP file
F2 - Open current or last Spreadsheet
F3 - Open current or last View file
F4 - Open Log file
F5 - Open STEP file in a text editor  (See also Help > Open STEP File in Apps)

F6 - Generate Speadsheets and/or Views from current or last set of multiple STEP files
F7 - Open current or last File Summary Spreadsheet from set of multiple STEP files

F8 - Decrease this font size
F9 - Increase this font size

For F1, F2, F3, F6, and F7 the last STEP file, Spreadsheet, and View are remembered between
sessions.  In other words, F1 can process the last STEP file from a previous session without
having to select the file.  F2 and F3 function similarly for Spreadsheets and Views."
    .tnb select .tnb.status
  }

# open STEP files help
  $Help add command -label "Open STEP File in Apps" -command {
outputMsg "\nOpen STEP File in Apps -----------------------------------------------------" blue
outputMsg "STEP files can be opened in other applications.  If applications are installed in their default
directory, then the pull-down menu in the Options tab will contain applications that can open a
STEP file such as STEP viewers, browsers, and conformance checkers.

See Help > User Guide (section 4.4.6)

The 'Tree View (for debugging)' option rearranges and indents the entities to show the
hierarchy of information in a STEP file.  The 'tree view' file (myfile-sfa.txt) is written to the
same directory as the STEP file or to the same user-defined directory specified in the Spreadsheet
tab.  It is useful for debugging STEP files but is not recommended for large STEP files.

The 'Default STEP Viewer' option opens the STEP file in whatever application is associated with
STEP files.  A text editor always appear in the menu.  Use F5 to open the STEP file in the text
editor."
    .tnb select .tnb.status
  }

# multiple files help
  $Help add command -label "Multiple STEP Files" -command {
outputMsg "\nMultiple STEP Files --------------------------------------------------------" blue
outputMsg "Multiple STEP files can be selected in the Open File(s) dialog by holding down the control or
shift key when selecting files or an entire directory of STEP files can be selected with 'Open
Multiple STEP Files in a Directory'.  Files in subdirectories of the selected directory can also
be processed.

See Help > User Guide (section 9)
See Examples > PMI Coverage Analysis

When processing multiple STEP files, a File Summary spreadsheet is generated in addition to
individual spreadsheets for each file.  The File Summary spreadsheet shows the entity count and
totals for all STEP files. The File Summary spreadsheet also links to the individual spreadsheets
and STEP files.

If only the File Summary spreadsheet is needed, it can be generated faster by deselecting most
Process categories and options in the Options tab.

If the reports for PMI Representation or Presentation are selected, then Coverage Analysis
worksheets are also generated."
    .tnb select .tnb.status
  }

  $Help add separator
  $Help add cascade -label "Analyze" -menu $Help.0
  set helpAnalyze [menu $Help.0 -tearoff 1]

# validation properties, PMI presentation, conformance checking help
  $helpAnalyze add command -label "PMI Representation (Semantic PMI)" -command {
outputMsg "\nPMI Representation ---------------------------------------------------------" blue
outputMsg "PMI Representation (Semantic PMI) includes all information necessary to represent geometric and
dimensional tolerances (GD&T) without any graphical presentation elements.  PMI Representation is
associated with CAD model geometry and is computer-interpretable to facilitate automated
consumption by downstream applications for manufacturing, measurement, inspection, and other
processes.

See Help > User Guide (section 5.1)
See Help > Analyze > PMI Coverage Analysis
See Examples > Spreadsheet - PMI Representation
See Examples > Sample STEP Files

PMI Representation is found mainly in AP242 files and is defined by the CAx-IF Recommended Practice
for Representation and Presentation of Product Manufacturing Information (AP242)
See Websites > Recommended Practices to access documentation.

Worksheets for the analysis of PMI Representation show a visual recreation of the representation
for Dimensional Tolerances, Geometric Tolerances, and Datum Features.  The results are in columns,
highlighted in yellow and green, on the relevant worksheets.  The GD&T is recreated as best as
possible given the constraints of Excel.

All of the visual recreation of Datum Systems, Dimensional Tolerances, and Geometric Tolerances
that are reported on individual worksheets are collected on one PMI Representation Summary
worksheet.

If STEP files from the NIST CAD models (Websites > PMI Validation Testing) are processed,
then the PMI the visual recreation of the PMI Representation is color-coded by the expected PMI
in each CAD model.  See Help > Analyze > NIST CAD Models.

Dimensional Tolerances are reported on the dimensional_characteristic_representation worksheet.
The dimension name, representation name, length/angle, length/angle name, and plus minus bounds
are reported.  The relevant section in the Recommended Practice is shown in the column headings.

Datum Features are reported on datum_* entities.  Datum_system will show the complete Datum
Reference Frame.  Datum Targets are reported on placed_datum_target_feature.

Geometric Tolerances are reported on *_tolerance entities by showing the complete Feature Control
Frame (FCF), and possible Dimensional Tolerance and Datum Feature.  The FCF should contain the
geometry tool, tolerance zone, datum reference frame, and associated modifiers.

If a Dimensional Tolerance refers to the same geometric element as a Geometric Tolerance, then it
will be shown above the FCF.  If a Datum Feature refers to the same geometric face as a Geometric
Tolerance, then it will be shown below the FCF.  If an expected Dimensional Tolerance is not shown
above a Geometric Tolerance then the tolerances do not reference the same geometric element.  For
example, referencing the edge of a hole versus the surfaces of a hole.

The association of the Datum Feature with a Geometric Tolerance is based on each referring to the
same geometric element.  However, the PMI Presentation might show the Geometric Tolerance and
Datum Feature as two separate annotations with leader lines attached to the same geometric element.

Some syntax errors that indicate non-conformance to a CAx-IF Recommended Practices related to PMI
Representation are also reported in the Status tab and the relevant worksheet cells.  Syntax
errors are highlighted in red.  See Help > Syntax Errors.

A PMI Representation Coverage Analysis worksheet is generated."
    .tnb select .tnb.status
  }

  $helpAnalyze add command -label "PMI Presentation (Graphical PMI)" -command {
outputMsg "\nPMI Presentation -----------------------------------------------------------" blue
outputMsg "PMI Presentation (Graphical PMI) consists of geometric elements such as lines and arcs
preserving the exact appearance (color, shape, positioning) of the geometric and dimensional
tolerance (GD&T) annotations.  PMI Presentation is not intended to be computer-interpretable and
does not carry any representation information, although it can be linked to its corresponding PMI
Representation.

See Help > User Guide (sections 5.2)
See Help > View > Graphical PMI
See Examples > View Part with PMI
See Examples > PMI Presentation, Validation Properties
See Examples > Sample STEP Files

The analysis of Graphical PMI on annotation_curve_occurrence, annotation_curve,
annotation_fill_area_occurrence, and tessellated_annotation_occurrence entities is supported.
Geometric entities used for PMI Presentation annotations are reported in columns, highlighted in
yellow and green, on those worksheets.  Presentation Style, Saved Views, Validation Properties,
Annotation Plane, Associated Geometry, and Associated Representation are also reported.

PMI Presentation is defined by the CAx-IF Recommended Practices for Representation and Presentation
of Product Manufacturing Information (AP242) and PMI Polyline Presentation (AP203/AP242)
See Websites > Recommended Practices to access documentation.

The Summary worksheet indicates on which worksheets PMI Presentation is reported.  Some syntax
errors related to PMI Presentation are also reported in the Status tab and the relevant worksheet
cells.  Syntax errors are highlighted in red.  See Help > Syntax Errors.

A PMI Presentation Coverage Analysis worksheet is generated.
See Help > Analyze > PMI Coverage Analysis."
    .tnb select .tnb.status
  }

# coverage analysis help
  $helpAnalyze add command -label "PMI Coverage Analysis" -command {
outputMsg "\nPMI Coverage Analysis ------------------------------------------------------" blue
outputMsg "PMI Coverage Analysis worksheets are generated when processing single or multiple files and when
reports for PMI Representation or Presentation are selected.

See Help > Analyze > PMI Representation
See Help > User Guide (sections 5.1.7 and 5.2.1)
See Examples > PMI Coverage Analysis

PMI Representation Coverage Analysis (semantic PMI) counts the number of PMI Elements found in a
STEP file for tolerances, dimensions, datums, modifiers, and CAx-IF Recommended Practices for PMI
Representation.  On the Coverage Analysis worksheet, some PMI Elements show their associated
symbol, while others show the relevant section in the Recommended Practice.  PMI Elements without
a section number do not have a Recommended Practice for their implementation.  The PMI Elements are
grouped by features related tolerances, tolerance zones, dimensions, dimension modifiers, datums,
datum targets, and other modifiers.  The number of some modifiers, e.g., maximum material condition,
does not differentiate whether they appear in the tolerance zone definition or datum reference frame.

Some PMI Elements might not be exported to a STEP file by your CAD system.  Some PMI Elements are
only in AP242 edition 2.

If STEP files from the NIST CAD models (Websites > PMI Validation Testing) are processed, then
the PMI Representation Coverage Analysis worksheet is color-coded by the expected number of PMI
elements in each CAD model.  See Help > Analyze > NIST CAD Models.

PMI Presentation Coverage Analysis (graphical PMI) counts the occurrences of a name attribute
defined in the CAx-IF Recommended Practice for PMI Representation and Presentation of PMI (AP242) or
PMI Polyline Presentation (AP203/AP242).  The name attribute is associated with the graphic elements
used to draw a PMI annotation.  There is no semantic meaning to the name attributes."
    .tnb select .tnb.status
  }

  $helpAnalyze add command -label "Validation Properties" -command {
outputMsg "\nValidation Properties ------------------------------------------------------" blue
outputMsg "Geometric, assembly, PMI, annotation, attribute, tessellated, composite, and FEA validation
properties are reported.  The property values are reported in columns highlighted in yellow and
green on the property_definition worksheet.  The worksheet can also be sorted and filtered.

See Help > User Guide (section 5.3)
See Examples > PMI Presentation, Validation Properties

Other properties and User-Defined Attributes are also reported.

Syntax errors related to validation property attribute values are also reported in the Status tab
and the relevant worksheet cells.  Syntax errors are highlighted in red.  See Help > Syntax Errors.

Clicking on the plus '+' symbols above the columns shows other columns that contain the entity ID
and attribute name of the validation property value.  All of the other columns can be shown or
hidden by clicking the '1' or '2' in the upper right corner of the spreadsheet.

The Summary worksheet indicates on the property_definition entity if properties are reported.

Validation properties are defined by the CAx-IF.  See Websites > Recommended Practices to access
documentation."
    .tnb select .tnb.status
  }

# NIST CAD model help
  $helpAnalyze add command -label "NIST CAD Models" -command {
outputMsg "\nNIST CAD Models ------------------------------------------------------------" blue
outputMsg "If a STEP file from a NIST CAD model is processed, then the PMI found in the STEP file is
automatically checked against the expected PMI in the corresponding NIST test case.  The PMI
Representation Coverage and Summary worksheets are color-coded by the expected PMI in each NIST test
case.  The color-coding only works if the STEP file name can be recognized as having been generated
from one of the NIST CAD models.

See Help > User Guide (section 8)
See Websites > PMI Validation Testing
See Examples > Spreadsheet - PMI Representation

* PMI Representation Summary *
This worksheet is color-coded by the expected PMI annotations in a test case drawing.
- Green is an exact match to an expected PMI annotation in the test case drawing.
- Cyan is a partial match.
- Yellow is a possible match.
- Red is no match.
For partial and possible matches, the best Similar PMI match is shown.  Missing PMI annotations are
also shown.

Trailing and leading zeros are ignored when matching a PMI annotation.  Matches also only consider
the current capabilities of PMI annotations in STEP AP242 and CAx-IF Recommended Practices.  For
example, PMI annotations for hole features such as counterbore, countersink, and depth are ignored.

Some causes of partial and possible matches are:
- missing associations of a geometric tolerance with a datum feature or dimension
- missing diameter and radius symbols
- wrong feature counts for repetitive dimensions
- wrong dimension or tolerance zone values
- missing or wrong values for dimension tolerances
- missing or wrong datum reference frames
- missing datum features
- missing or incorrect modifiers for dimensions, tolerance zones, and datum reference frames
- missing composite tolerances

* PMI Representation Coverage Analysis *
This worksheet is color-coded by the expected number of PMI elements in a test case drawing.  The
expected results were determined by manually counting the number of PMI elements in each drawing.
Counting of some modifiers, e.g. maximum material condition, does not differentiate whether they
appear in the tolerance zone definition or datum reference frame.
- A green cell is a match to the expected number of PMI elements. (3/3)
- Cyan means that more were found than expected. (4/3)
- Yellow means that less were found than expected. (2/3)
- Red means that no instances of an expected PMI element were found. (0/3)
- Magenta means that some PMI elements were found when none were expected. (3/0)

* Missing PMI *
Missing PMI annotations on the Summary worksheet or PMI elements on the Coverage worksheet might
mean that the CAD system or translator:
- did not or cannot correctly create in the CAD model a PMI annotation defined in a NIST test case
- did not follow CAx-IF Recommended Practices for PMI (See Websites > Recommended Practices)
- has not implemented exporting a PMI element to a STEP file
- mapped an internal PMI element to the wrong STEP PMI element

Some of the NIST test cases have complex PMI annotations that are not commonly used.  There might
be ambiguities in counting the number of PMI elements, particularly for dimensions."
    .tnb select .tnb.status
  }

  $Help add cascade -label "View" -menu $Help.1
  set helpView [menu $Help.1 -tearoff 1]

  $helpView add command -label "Part Geometry" -command {
outputMsg "\nPart Geometry --------------------------------------------------------------" blue
outputMsg "Views are shown in the default web browser.  Older versions of web browsers are not supported.
All Views are written to: myfile-sfa.html  An Internet connection is required to show View files.

Part geometry (b-rep) is shown for any STEP file where the geometry is modeled with
advanced_brep_shape_representation, manifold_surface_shape_representation, manifold_solid_brep, or
shell_based_surface_model entities.

Supplemental geometry (axes, points, lines, circles, planes, cylinders) are shown.
See Help > View > Supplemental Geometry

Counterbore and countersink holes in AP242 edition 2 are also shown.  Countbore holes are green and
countersink holes are cyan.  Both types of holes are shown with black dot at the entry point for drilling.

Part colors are ignored if multiple colors are specified.  Overriding style colors are also ignored.

Part geometry might also include supplemental geometry for planes.  In some cases, curved surfaces might
appear jagged or incomplete.  Some part geometry cannot be processed.  Views for very large STEP files
might take 10-20 minutes to generate.

See Examples > View Part with PMI
See Websites > STEP File Viewers

Some other STEP file viewers cannot view PMI, tessellated part geometry, and finite element models.
However, those viewers usually have better capabilities for viewing and measuring part geometry.

The part geometry view is based on OpenCascade and pythonOCC.  See Help > About"
    .tnb select .tnb.status
  }

  $helpView add command -label "Graphical PMI" -command {
outputMsg "\nGraphical PMI --------------------------------------------------------------" blue
outputMsg "Graphical PMI (PMI Presentation) annotations composed of polylines, lines, circles, and tessellated
geometry are supported for viewing.  The color of the annotations can be modified.  Filled
characters are not filled.  PMI associated with Saved Views can be switched on and off.  Some
Graphical PMI might not have equivalent or any Semantic PMI in the STEP file.  Some STEP files with
Semantic PMI might not have any Graphical PMI.

See Help > User Guide (sections 7.1.1)
See Help > Analyze > PMI Presentation
See Examples > View Part with PMI
See Examples > Sample STEP Files"
    .tnb select .tnb.status
  }

  $helpView add command -label "Supplemental Geometry" -command {
outputMsg "\nSupplemental Geometry ------------------------------------------------------" blue
outputMsg "Supplemental geometry is shown only if part or PMI is also viewed.  Supplemental geometry is not
associated with Saved Views.

The following types of supplemental geometry and associated text are supported.
- Coordinate System: red/green/blue axes or by color assigned to axes
- Plane: blue transparent outlined square
- Cylinder: blue transparent cylinder
- Line/Circle: purple line/circle
- Point: black dot
- Tessellated Surface: faces outlined in black

Lines and circles that are trimmed by cartesian_point will not be trimmed.  Bounding edges for
planes and cylinders are ignored.  All bounded and unbounded planes are shown with a fixed size."
    .tnb select .tnb.status
  }

  $helpView add command -label "AP242 Tessellated Part Geometry" -command {
outputMsg "\nAP242 Tessellated Part Geometry --------------------------------------------" blue
outputMsg "Tessellated part geometry is supported by AP242 and is usually supplementary to part geometry.

** Parts in an assembly might have the wrong position and orientation or be missing. **

Faces in a tessellated shell are outlined in black.  Lines generated from tessellated edges are also
shown.  A wireframe mesh, outlining the facets of the tessellated surfaces can also be shown.  If
both are present, tessellated edges might be obscured by the wireframe mesh.  [string totitle [lindex $defaultColor 1]] is used as the
color assigned to tessellated solids, shells, or faces that do not have colors assigned to them.

See Help > User Guide (section 7.1.3)
See Examples > AP242 Tessellated Part with PMI"
    .tnb select .tnb.status
  }

  $helpView add command -label "AP209 Finite Element Model" -command {
outputMsg "\nAP209 Finite Element Model -------------------------------------------------" blue
outputMsg "All AP209 entities are always processed and written to a spreadsheet unless a User-defined list is
used.

The AP209 finite element model composed of bodes, mesh, elements, boundary conditions, loads, and
displacments are shown and can be toggled on and off in the viewer.  Internal faces for solid
elements are not shown.  Elements can be made transparent although it is only approximate.

Nodal loads and element surface pressures are shown.  Load vectors are colored by their magnitude.
The length of load vectors can be scaled by their magnitude.  Forces use a single-headed arrow.
Moments use a double-headed arrow.

Displacement vectors are colored by their magnitude.  The length of displacement vectors can be
scaled by their magnitude depending on if they have a tail.  The finite element mesh is not deformed.

Boundary conditions for translation DOF are shown with a red, green, or blue line along the
X, Y, or Z axes depending on the constrained DOF.  Boundary conditions for rotation DOF are shown
with a red, green, or blue circle around the X, Y, or Z axes depending on the constrained DOF.  A
gray box is used when all six DOF are constrained.  A gray pyramid is used when all three
translation DOF are constrained.  A gray sphere is used when all three rotation DOF are constrained.

Stresses and strains are not shown.  Multiple coordinate systems are not considered.

Setting Maximum Rows (Spreadsheet tab) does not affect the view.  For large AP209 files, there
might be insufficient memory to process all of the elements, loads, displacements, and boundary
conditions.

See Help > User Guide (section 7.1.4)
See Examples > AP209 Finite Element Model
See Websites > AP209 FEA"
    .tnb select .tnb.status
  }

  $Help add separator
  $Help add command -label "Syntax Errors" -command {
outputMsg "\nSyntax Errors --------------------------------------------------------------" blue
outputMsg "Syntax error information and other errors can be used to debug a STEP file.  Most syntax errors are
generated when Analysis related to Semantic PMI, Graphical PMI, and Validation Properties are
selected.  Analysis syntax errors and some other errors are shown in the Status tab and highlighted
in red or yellow.

Analysis syntax errors for Analysis are related to CAx-IF Recommended Practices and usually refer to
a specific section, figure, or table in a Recommended Practice.  Some references to section, figure,
and table numbers in a recommended practice might be to a newer version of a recommended practice
that has not been released to the public.  Specific section, figure, and table numbers might be wrong
relative to the publicly available recommended practice.

Most entity types that have syntax errors or some other errors are highlighted in gray in column A
on the File Summary worksheet.  A comment indicating that there are errors is shown with a small red
triangle in the upper right corner of the cell.

On an entity worksheet, most syntax errors are highlighted in red and have a cell comment with the
text of the syntax error that was shown in the Status tab.

See Help > User Guide (section 6)
See Websites > Recommended Practices

Log File
  All text in the Status tab can be written to a Log File when a STEP file is processed (Options tab).
  The log file is written to myfile-sfa.log.  In a log file, error messages are highlighted by ***.
  Use F4 to open the log file.

Command-line version
  Some syntax errors and warning are reported only by the command-line version sfa-cl.exe when reading
  the STEP file.  Use the 'stats' option to only check for errors and warnings without generating a
  spreadsheet."
    .tnb select .tnb.status
  }

  $Help add command -label "Crash Recovery" -command {
outputMsg "\nCrash Recovery -------------------------------------------------------------" blue
outputMsg "Sometimes the STEP File Analyzer and Viewer crashes after a STEP file has been successfully
opened and the processing of entities has started.  Popup dialogs might appear that say
\"Runtime Error!\" or \"ActiveState Basekit has stopped working\".

See Help > User Guide (sections 2.4 and 10)

A crash is most likely due to syntax errors in the STEP file or sometimes due to limitations of
the toolkit used to read STEP files.  To see which type of entity caused the error, check the
Status tab to see which type of entity was last processed.  A crash might also be caused by a very
large STEP file.

Workarounds for this problem:

1 - This software keeps track of the last entity type processed when it crashed.  Simply restart
the STEP File Analyzer and Viewer and use F1 to process the last STEP file or use F6 if processing
multiple files.  The type of entity that caused the crash will be skipped.  The list of bad entity
types that will not be processed is stored in myfile-skip.dat.

2 - Deselect all Analyze and Inverse Relationships options.  If one of these options caused the
crash, then the *-skip.dat file is still created as described above and might need to be deleted.

3 - Processing of the type of entity that caused the error can be deselected in the Options tab
under Process.  However, this will prevent processing of other entities that do not cause a crash.

4 - Run the command-line version 'sfa-cl.exe' in a command prompt window.  Use the 'stats' option
so that no spreadsheet or view is generated.  The output from reading the STEP file might show
error and warning messages that might have caused the software to crash.  Those messages will be
between the '*** Begin ST-Developer output' and '*** End ST-Developer output' messages.

NOTE - If syntax errors related to the bad entities are corrected, then delete the *-skip.dat file
so that the corrected entities are processed.  When the STEP file is processed, the list of
specific entities that are not processed is reported."
  .tnb select .tnb.status
}

# large files help
  $Help add command -label "Large STEP Files" -command {
outputMsg "\nLarge STEP Files -----------------------------------------------------------" blue
outputMsg "To reduce the amount of time to process large STEP files and to reduce the size of the resulting
spreadsheet, several options are available:
- In the Process section, deselect entity types Geometry and Coordinates
- In the Process section, select only a User-Defined List of required entities
- In the Spreadsheet tab, select a smaller value for the Maximum Rows
- In the Options tab, deselect Analyze options and Inverse Relationships

The STEP File Analyzer and Viewer might also crash when processing very large STEP files.  Popup
dialogs might appear that say 'Unable to alloc xxx bytes'.  See the Help > Crash Recovery."
    .tnb select .tnb.status
  }
  $Help add command -label "Supported STEP APs" -command {
    if {[llength [glob -nocomplain -directory $ifcsvrDir *.rose]] < 12} {copyRoseFiles}

    outputMsg "\nSupported STEP APs ----------------------------------------------------------" blue
    outputMsg "The following STEP Application Protocols (AP) and other schemas are supported.\nThe name of the AP is found on the FILE_SCHEMA entity in the HEADER section of a STEP file.\nThe 'e1' notation after an AP number below refers to an older version of that AP.\n"

    set nschema 0
    catch {file delete -force -- [file join $ifcsvrDir ap214e3_2010.rose]}

    set schemas {}
    foreach match [lsort [glob -nocomplain -directory $ifcsvrDir *.rose]] {
      set schema [string toupper [file rootname [file tail $match]]]
      if {[string first "HEADER_SECTION" $schema] == -1 && [string first "KEYSTONE" $schema] == -1 && [string range $schema end-2 end] != "MIM"} {
        if {[info exists stepAPs($schema)]} {
          if {[string first "CONFIGURATION" $schema] != 0} {
            lappend schemas "$stepAPs($schema) - $schema"
          } else {
            lappend schemas $schema
          }
        } elseif {[string first "AP2" $schema] == 0} {
          lappend schemas "[string range $schema 0 4] - $schema"
        } elseif {[string first "IFC" $schema] == -1} {
          lappend schemas $schema
        } elseif {$schema == "IFC2X3" || $schema == "IFC4"} {
          lappend schemas $schema
        }
        incr nschema
      }
    }

    set n 0
    foreach item [lsort $schemas] {
      set c1 [string first "-" $item]
      if {$c1 == -1} {
        if {$n == 0} {
          incr n
          outputMsg "\nOther Schemas"
        }
        outputMsg "  [string toupper $item]"
      } else {
        outputMsg "  [string range $item 0 $c1][string toupper [string range $item $c1+1 end]]"
      }
    }

    if {$nschema == 0} {errorMsg "No Supported STEP APs were found.\nThere was a problem copying STEP schema files (*.rose) to the IFCsvr/dll directory."}
    outputMsg "\nSee the Websites menu for information about the STEP Format, EXPRESS Schemas, AP242, and more."

    .tnb select .tnb.status
  }

  $Help add separator
  if {"$nistVersion"} {
    $Help add command -label "Disclaimers" -command {showDisclaimer}
    $Help add command -label "NIST Disclaimer" -command {openURL https://www.nist.gov/disclaimer}
  }
  $Help add command -label "About" -command {
    outputMsg "\nSTEP File Analyzer and Viewer ---------------------------------------------------------" blue
    outputMsg "Version:  [getVersion]"
    outputMsg "Updated:  [string trim [clock format $progtime -format "%e %b %Y"]]"
    if {"$nistVersion"} {
      outputMsg "Contact:  [lindex $contact 0], [lindex $contact 1]

The STEP File Analyzer and Viewer was first released in April 2012 and is developed at
NIST in the Systems Integration Division of the Engineering Laboratory.  Click the NIST
logo below for the NIST website.

See Help > Disclaimer and NIST Disclaimer

Credits
- Generating spreadsheets:        Microsoft Excel (https://products.office.com/excel)
- Reading and parsing STEP files: IFCsvr (https://groups.yahoo.com/neo/groups/ifcsvr-users/info)
                                  License agreement C:\\Program Files (x86)\\IFCsvrR300\\doc
                                  IFCsvr ActiveX Component, Copyright \u00A9 1999, 2005 SECOM Co., Ltd. All Rights Reserved
- Viewing B-rep part geometry:    OpenCascade (https://www.opencascade.com/) and
                                  pythonOCC (https://github.com/tpaviot/pythonocc)
                                  See Websites > STEP Software"
    } else {
      outputMsg "\nThis version was built from the NIST STEP File Analyzer and Viewer source\ncode available on GitHub.  https://github.com/usnistgov/SFA"
    }

  # debug
    if {$opt(XL_ROWLIM) == 100003 || $env(USERDOMAIN) == "NIST"} {
      outputMsg " "
      outputMsg "Environment variables" red
      foreach id [lsort [array names env]] {
        foreach id1 [list HOME Program System USER TEMP TMP APP ROSE EDM] {
          if {[string first $id1 $id] == 0} {outputMsg " $id   $env($id)"; break}
        }
      }
      outputMsg "Registry values" red
      catch {outputMsg " Personal  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Personal}]"}
      catch {outputMsg " Desktop   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]"}
      catch {outputMsg " Programs  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Programs}]"}
      catch {outputMsg " AppData   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local AppData}]"}
      catch {outputMsg " Browser   [registry get {HKEY_CURRENT_USER\Software\Classes\http\shell\open\command} {}]"}
      outputMsg "SFA variables" red
      catch {outputMsg " Drive $drive"}
      catch {outputMsg " Home  $myhome"}
      catch {outputMsg " Docs  $mydocs"}
      catch {outputMsg " Desk  $mydesk"}
      catch {outputMsg " Menu  $mymenu"}
      catch {outputMsg " Temp  $mytemp  ([file exists $mytemp])"}
      catch {outputMsg " ifcsvrDir  [file nativename $ifcsvrDir]"}
      if {[info exists virtualDir]} {outputMsg " virtualDir  $virtualDir"}
      outputMsg " pf32  $pf32"
      if {$pf64 != ""} {outputMsg " pf64  $pf64"}
      outputMsg "Other variables" red
      outputMsg " Tcl [info patchlevel]"
      outputMsg " twapi [package versions twapi]"
      outputMsg " $tcl_platform(os) $tcl_platform(osVersion)"
      outputMsg " Excel $excelVersion"
    }
    .tnb select .tnb.status
  }

# examples menu
  $Examples add command -label "Sample STEP Files (zip)" -command {openURL https://s3.amazonaws.com/nist-el/mfg_digitalthread/NIST_CTC_STEP_PMI.zip}
  $Examples add command -label "STEP File Library"       -command {openURL https://www.cax-if.org/library/}
  $Examples add command -label "AP203e2 Archive"         -command {openURL http://web.archive.org/web/20160812122922/http://www.steptools.com/support/stdev_docs/stpfiles/ap203e2/index.html}
  $Examples add command -label "AP203 Archive"           -command {openURL http://web.archive.org/web/20160812122922/http://www.steptools.com/support/stdev_docs/stpfiles/ap203/index.html}
  $Examples add command -label "AP214 Archive"           -command {openURL http://web.archive.org/web/20160903141712/http://www.steptools.com/support/stdev_docs/stpfiles/ap214/index.html}
  $Examples add separator
  $Examples add command -label "View Part with PMI"              -command {openURL https://pages.nist.gov/CAD-PMI-Testing/graphical-pmi-viewer.html}
  $Examples add command -label "AP242 Tessellated Part with PMI" -command {openURL https://pages.nist.gov/CAD-PMI-Testing/tessellated-part-geometry.html}
  $Examples add command -label "AP209 Finite Element Model"      -command {openURL https://pages.nist.gov/CAD-PMI-Testing/ap209-viewer.html}
  $Examples add separator
  $Examples add command -label "Spreadsheet - PMI Representation"        -command {openURL https://s3.amazonaws.com/nist-el/mfg_digitalthread/STEP-File-Analyzer-PMI-Representation_stp.xlsx}
  $Examples add command -label "PMI Presentation, Validation Properties" -command {openURL https://s3.amazonaws.com/nist-el/mfg_digitalthread/STEP-File-Analyzer_stp.xlsx}
  $Examples add command -label "PMI Coverage Analysis"                   -command {openURL https://s3.amazonaws.com/nist-el/mfg_digitalthread/STEP-File-Analyzer-Coverage.xlsx}
}

#-------------------------------------------------------------------------------
# Websites menu
proc guiWebsitesMenu {} {
  global Websites

  $Websites add command -label "STEP File Analyzer and Viewer"             -command {openURL https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer}
  $Websites add command -label "NIST Journal of Research"                  -command {openURL https://www.nist.gov/publications/step-file-analyzer-software}
  $Websites add command -label "Conformance Checking of PMI in STEP Files" -command {openURL https://www.nist.gov/publications/conformance-checking-pmi-representation-cad-model-step-data-exchange-files}
  $Websites add command -label "PMI Validation Testing"                    -command {openURL https://www.nist.gov/el/systems-integration-division-73400/mbe-pmi-validation-and-conformance-testing-project/download}
  $Websites add command -label "Digital Thread for Smart Manufacturing"    -command {openURL https://www.nist.gov/el/systems-integration-division-73400/enabling-digital-thread-smart-manufacturing}
  $Websites add command -label "STEP: The Grand Experience"                -command {openURL https://www.nist.gov/publications/step-grand-experience}

  $Websites add separator
  $Websites add command -label "CAx Implementor Forum (CAx-IF)" -command {openURL https://www.cax-if.org}
  $Websites add command -label "STEP File Viewers"              -command {openURL https://www.cax-if.org/step_viewers.html}
  $Websites add command -label "Recommended Practices"          -command {openURL https://www.cax-if.org/joint_testing_info.html#recpracs}
  $Websites add command -label "CAD Implementations"            -command {openURL https://www.cax-if.org/vendor_info.php}

  $Websites add separator
  $Websites add command -label "CAE-IF" -command {openURL http://afnet.fr/dotank/sps/plm-committee/cae-if/}
  $Websites add command -label "PDM-IF" -command {openURL http://www.pdm-if.org/}

  $Websites add separator
  $Websites add cascade -label "AP242" -menu $Websites.0
  set Websites0 [menu $Websites.0 -tearoff 1]
  $Websites0 add command -label "AP242 Project"           -command {openURL http://www.ap242.org}
  $Websites0 add command -label "Paper"             -command {openURL https://www.nist.gov/publications/portrait-iso-step-tolerancing-standard-enabler-smart-manufacturing-systems}
  $Websites0 add command -label "Presentation"      -command {openURL https://www.nist.gov/document-2058}
  $Websites0 add command -label "Benchmark Testing" -command {openURL http://www.asd-ssg.org/step-ap242-benchmark}
  $Websites0 add command -label "Edition 2"         -command {openURL http://www.ap242.org/edition-2}

  $Websites add command -label "AP209 FEA"                -command {openURL http://www.ap209.org}

  $Websites add separator
  $Websites add cascade -label "STEP Format and Schemas" -menu $Websites.2
  set Websites2 [menu $Websites.2 -tearoff 1]
  $Websites2 add command -label "STEP Format"                 -command {openURL https://www.loc.gov/preservation/digital/formats/fdd/fdd000448.shtml}
  $Websites2 add command -label "ISO 10303 Part 21"           -command {openURL https://en.wikipedia.org/wiki/ISO_10303-21}
  $Websites2 add command -label "ISO 10303 Part 21 Edition 3" -command {openURL https://www.steptools.com/stds/step/}
  $Websites2 add separator
  $Websites2 add command -label "EXPRESS Schemas"             -command {openURL https://www.cax-if.org/joint_testing_info.html#schemas}
  $Websites2 add command -label "More EXPRESS Schemas"        -command {openURL http://web.archive.org/web/20160322005246/www.steptools.com/support/stdev_docs/express/}
  $Websites2 add command -label "ISO 10303 Part 11 EXPRESS"   -command {openURL https://www.loc.gov/preservation/digital/formats/fdd/fdd000449.shtml}
  $Websites2 add separator
  $Websites2 add command -label "AP235 Properties"            -command {openURL http://www.ap235.org}
  $Websites2 add command -label "AP238 Machining"             -command {openURL http://www.ap238.org}
  $Websites2 add command -label "AP239 PLCS"                  -command {openURL http://www.ap239.org}
  $Websites2 add command -label "AP243 MoSSEC"                -command {openURL http://www.mossec.org/}

  $Websites add cascade -label "STEP Software" -menu $Websites.4
  set Websites4 [menu $Websites.4 -tearoff 1]
  $Websites4 add command -label "STEP File Analyzer and Viewer source code" -command {openURL https://github.com/usnistgov/SFA}
  $Websites4 add command -label "Digital Manufacturing Certificate Toolkit" -command {openURL https://github.com/usnistgov/DT4SM/tree/master/DMC-Toolkit}
  $Websites4 add separator
  $Websites4 add command -label "STEP Tools Software"                       -command {openURL https://github.com/steptools}
  $Websites4 add command -label "OpenCascade STEP Processor"                -command {openURL https://www.opencascade.com/doc/occt-7.0.0/overview/html/occt_user_guides__step.html}
  $Websites4 add command -label "pythonOCC"                                 -command {openURL https://github.com/tpaviot/pythonocc}
  $Websites4 add command -label "STEP to X3D Translation"                   -command {openURL http://www.web3d.org/wiki/index.php/STEP_X3D_Translation}
  $Websites4 add command -label "STEP Class Library (STEPcode)"             -command {openURL https://www.nist.gov/services-resources/software/step-class-library-scl}
  $Websites4 add command -label "Express Engine"                            -command {openURL http://exp-engine.sourceforge.net/}
  #$Websites4 add command -label "STEP Engine"                               -command {openURL http://rdf.bg/product-list/step-engine/}

  $Websites add cascade -label "STEP Related Organizations" -menu $Websites.3
  set Websites3 [menu $Websites.3 -tearoff 1]
  $Websites3 add command -label "PDES, Inc. (U.S.)"                         -command {openURL http://pdesinc.org}
  $Websites3 add command -label "prostep ivip (Germany)"                    -command {openURL https://www.prostep.org/en/projects/}
  $Websites3 add command -label "AFNeT (France)"                            -command {openURL http://afnet.fr/dotank/sps/plm-committee/}
  $Websites3 add separator
  $Websites3 add command -label "ISO TC184/SC4"                             -command {openURL https://www.iso.org/committee/54158.html}
  $Websites3 add command -label "LOTAR (LOng Term Archiving and Retrieval)" -command {openURL http://www.lotar-international.org}
  $Websites3 add command -label "ASD Strategic Standardisation Group"       -command {openURL http://www.asd-ssg.org/}
  $Websites3 add command -label "KStep (Korea)"                             -command {openURL http://www.kstep.or.kr/}
}

#-------------------------------------------------------------------------------
proc showDisclaimer {} {
  global nistVersion

  if {$nistVersion} {
    outputMsg "\nDisclaimer -------------------------------------------------------------" blue
    outputMsg "This software was developed at the National Institute of Standards and Technology by employees of
the Federal Government in the course of their official duties. Pursuant to Title 17 Section 105 of
the United States Code this software is not subject to copyright protection and is in the public
domain.  This software is an experimental system.  NIST assumes no responsibility whatsoever for
its use by other parties, and makes no guarantees, expressed or implied, about its quality,
reliability, or any other characteristic.

The Examples menu of this software provides links to several sources of STEP files.  This software
and other software might indicate that there are errors in some of the STEP files.  NIST assumes
no responsibility whatsoever for the use of the STEP files by other parties, and makes no
guarantees, expressed or implied, about their quality, reliability, or any other characteristic.

Any mention of commercial products or references to web pages in this software is for information
purposes only; it does not imply recommendation or endorsement by NIST.  For any of the web links
in this software, NIST does not necessarily endorse the views expressed, or concur with the facts
presented on those web sites.

This software uses Microsoft Excel and IFCsvr that are covered by their own Software License
Agreements.  The IFCsvr agreement is in C:\\Program Files (x86)\\IFCsvrR300\\doc  The B-rep part
geometry viewer is based on software from OpenCascade and pythonOCC.

See Help > NIST Disclaimer and Help > About"
    .tnb select .tnb.status

set txt "This software was developed at the National Institute of Standards and Technology by employees of the Federal Government in the course of their official duties. Pursuant to Title 17 Section 105 of the United States Code this software is not subject to copyright protection and is in the public domain.  This software is an experimental system.  NIST assumes no responsibility whatsoever for its use by other parties, and makes no guarantees, expressed or implied, about its quality, reliability, or any other characteristic.

The Examples menu of this software provides links to several sources of STEP files.  This software and other software might indicate that there are errors in some of the STEP files.  NIST assumes no responsibility whatsoever for the use of the STEP files by other parties, and makes no guarantees, expressed or implied, about their quality, reliability, or any other characteristic.

Any mention of commercial products or references to web pages in this software is for information purposes only; it does not imply recommendation or endorsement by NIST.  For any of the web links in this software, NIST does not necessarily endorse the views expressed, or concur with the facts presented on those web sites.

This software uses Microsoft Excel and IFCsvr that are covered by their own Software License Agreements.  The IFCsvr agreement is in C:\\Program Files (x86)\\IFCsvrR300\\doc  The B-rep part geometry viewer is based on software from OpenCascade and pythonOCC.

See Help > NIST Disclaimer and Help > About"

    tk_messageBox -type ok -icon info -title "Disclaimers" -message $txt
  }
}

#-------------------------------------------------------------------------------
# crash recovery dialog
proc showCrashRecovery {} {

set txt "Sometimes the STEP File Analyzer and Viewer crashes AFTER a file has been successfully opened and the processing of entities has started.

A crash is most likely due to syntax errors in the STEP file or sometimes due to limitations of the toolkit used to read STEP files.

If this happens, simply restart the STEP File Analyzer and Viewer and process the same STEP file again by using function key F1 or if processing multiple STEP files use F6.  Also deselect, Analyze and Inverse Relationships options.

The STEP File Analyzer and Viewer keeps track of which entity type caused the error for a particular STEP file and won't process that type again.  The bad entities types are stored in a file *-skip.dat  If syntax errors related to the bad entities are corrected, then delete the *-skip.dat file so that the corrected entities are processed.

The software might also crash when processing very large STEP files.  In this case, deselect some entity types to process in Options tab or use a User-Defined List of entities to process.

More details about recovering from a crash are explained in Help > Crash Recovery and in the User Guide.

Please report other types of crashes to the software developer."

  tk_messageBox -type ok -icon error -title "What to do if the STEP File Analyzer and Viewer crashes?" -message $txt
}

#-------------------------------------------------------------------------------
proc guiToolTip {ttmsg tt} {
  global ap242only entCategory

  set ttlen 0
  set lchar ""
  set r1 0
  set ttlim 120
  if {$tt == "PR_STEP_PRES" || $tt == "PR_STEP_GEOM"} {set ttlim 150}

  foreach item [lsort $entCategory($tt)] {
    if {[string range $item 0 $r1] != $lchar && $lchar != ""} {
      if {[string index $ttmsg end] != "\n"} {append ttmsg "\n"}
      set ttlen 0
    }
    set ent $item

# check for ap242 entities
    if {[lsearch $ap242only $ent] != -1} {append ent "*"}

    append ttmsg "$ent   "
    incr ttlen [string length $ent]
    if {$ttlen > $ttlim} {
      if {[string index $ttmsg end] != "\n"} {append ttmsg "\n"}
      set ttlen 0
    }
    set lchar [string range $ent 0 $r1]
  }
  return $ttmsg
}

#-------------------------------------------------------------------------------
proc getOpenPrograms {} {
  global dispApps dispCmds dispCmd appNames appName env
  global drive editorCmd developer myhome pf32 pf64

# Including any of the CAD viewers and software below does not imply a recommendation or endorsement of them by NIST https://www.nist.gov/disclaimer
# For more STEP viewers, go to https://www.cax-if.org/step_viewers.html

  regsub {\\} $pf32 "/" p32
  lappend pflist $p32
  if {$pf64 != "" && $pf64 != $pf32} {
    regsub {\\} $pf64 "/" p64
    lappend pflist $p64
  }
  set lastver 0

# Jotne EDM Model Checker
  if {$developer} {
    set edms [glob -nocomplain -directory [file join $drive edm] -join edm* bin Edms.exe]
    foreach match $edms {
      set name "EDM Model Checker"
      if {[string first "edm5" $match] != -1} {
        set num 5
      } elseif {[string first "edmsix" $match] != -1} {
        set num 6
      }
      set dispApps($match) "$name $num"
    }
  }

# STEP Tools apps
  foreach pf $pflist {
    if {[file isdirectory [file join $pf "STEP Tools"]]} {
      set applist [list \
        [list ap203checkgui.exe "AP203 Conformance Checker"] \
        [list ap209checkgui.exe "AP209 Conformance Checker"] \
        [list ap214checkgui.exe "AP214 Conformance Checker"] \
        [list apconformgui.exe "AP Conformance Checker"] \
        [list stepbrws.exe "STEP File Browser"] \
        [list stepcleangui.exe "STEP File Cleaner"] \
        [list stpcheckgui.exe "STEP Check and Browse"] \
        [list stview.exe "ST-Viewer"] \
      ]
      foreach app $applist {
        set stmatch ""
        foreach match [glob -nocomplain -directory $pf -join "STEP Tools" "ST-Developer*" bin [lindex $app 0]] {
          if {$stmatch == ""} {
            set stmatch $match
            set lastver [lindex [split [file nativename $match] [file separator]] 3]
          } else {
            set ver [lindex [split [file nativename $match] [file separator]] 3]
            if {$ver > $lastver} {set stmatch $match}
          }
        }
        if {$stmatch != ""} {
          if {![info exists dispApps($stmatch)]} {set dispApps($stmatch) [lindex $app 1]}
        }
      }
    }

# other STEP file apps
    set applist [list \
      [list {*}[glob -nocomplain -directory [file join $pf] -join "Afanche3D*" "Afanche3D*.exe"] Afanche3D] \
      [list {*}[glob -nocomplain -directory [file join $pf "Common Files"] -join "eDrawings*" eDrawings.exe] eDrawings] \
      [list {*}[glob -nocomplain -directory [file join $pf "SOLIDWORKS Corp"] -join "eDrawings (*)" eDrawings.exe] eDrawings] \
      [list {*}[glob -nocomplain -directory [file join $pf "Stratasys Direct Manufacturing"] -join "SolidView Pro RP *" bin SldView.exe] SolidView] \
      [list {*}[glob -nocomplain -directory [file join $pf "TransMagic Inc"] -join "TransMagic *" System code bin TransMagic.exe] TransMagic] \
      [list {*}[glob -nocomplain -directory [file join $pf Actify SpinFire] -join "*" SpinFire.exe] SpinFire] \
      [list {*}[glob -nocomplain -directory [file join $pf CADSoftTools] -join "ABViewer*" ABViewer.exe] ABViewer] \
      [list {*}[glob -nocomplain -directory [file join $pf Kubotek] -join "KDisplayView*" KDisplayView.exe] "K-Display View"] \
      [list {*}[glob -nocomplain -directory [file join $pf Kubotek] -join "Spectrum*" Spectrum.exe] Spectrum] \
      [list {*}[glob -nocomplain -directory [file join $pf] -join "3D-Tool V*" 3D-Tool.exe] 3D-Tool] \
      [list {*}[glob -nocomplain -directory [file join $pf] -join "VariCADViewer *" bin varicad-x64.exe] "VariCAD Viewer"] \
      [list {*}[glob -nocomplain -directory [file join $pf] -join ZWSOFT "CADbro *" CADbro.exe] CADbro] \
    ]
    if {$pf64 == ""} {
      lappend applist [list {*}[glob -nocomplain -directory [file join $pf] -join "VariCADViewer *" bin varicad-i386.exe] "VariCAD Viewer (32-bit)"]
    }

    foreach app $applist {
      if {[llength $app] == 2} {
        set match [join [lindex $app 0]]
        if {$match != "" && ![info exists dispApps($match)]} {
          set dispApps($match) [lindex $app 1]
        }
      }
    }

    set applist [list \
      [list [file join $pf "3DJuump X64" 3DJuump.exe] "3DJuump"] \
      [list [file join $pf "CAD Assistant" CADAssistant.exe] "CAD Assistant"] \
      [list [file join $pf "CAD Exchanger" bin Exchanger.exe] "CAD Exchanger"] \
      [list [file join $pf "SOLIDWORKS Corp" eDrawings eDrawings.exe] "eDrawings"] \
      [list [file join $pf "STEP Tools" "STEP-NC Machine Personal Edition" STEPNCExplorer.exe] "STEP-NC Machine"] \
      [list [file join $pf "STEP Tools" "STEP-NC Machine Personal Edition" STEPNCExplorer_x86.exe] "STEP-NC Machine"] \
      [list [file join $pf "STEP Tools" "STEP-NC Machine" STEPNCExplorer.exe] "STEP-NC Machine"] \
      [list [file join $pf "STEP Tools" "STEP-NC Machine" STEPNCExplorer_x86.exe] "STEP-NC Machine"] \
      [list [file join $pf "Tekla BIMsight" BIMsight.exe] "Tekla BIMsight"] \
      [list [file join $pf CadFaster QuickStep QuickStep.exe] QuickStep] \
      [list [file join $pf gCAD3D gCAD3D.bat] gCAD3D] \
      [list [file join $pf Glovius Glovius glovius.exe] Glovius] \
      [list [file join $pf IFCBrowser IfcQuickBrowser.exe] IfcQuickBrowser] \
      [list [file join $pf Kisters 3DViewStation 3DViewStation.exe] 3DViewStation] \
      [list [file join $pf STPViewer STPViewer.exe] "STP Viewer"] \
    ]
    foreach app $applist {
      if {[file exists [lindex $app 0]]} {
        set name [lindex $app 1]
        set dispApps([lindex $app 0]) $name
      }
    }

# FreeCAD
    foreach app [list {*}[glob -nocomplain -directory [file join $pf] -join "FreeCAD *" bin FreeCAD.exe] FreeCAD] {
      set ver [lindex [split [file nativename $app] [file separator]] 2]
      if {$pf64 != "" && [string first "x86" $app] != -1} {append ver " (32-bit)"}
      set dispApps($app) $ver
    }

# Tetra4D in Adobe Acrobat
    for {set i 40} {$i > 9} {incr i -1} {
      if {$i > 11} {
        set j "20$i"
      } else {
        set j "$i.0"
      }
      foreach match [glob -nocomplain -directory $pf -join Adobe "Acrobat $j" Acrobat Acrobat.exe] {
        if {[file exists [file join $pf Adobe "Acrobat $j" Acrobat plug_ins 3DPDFConverter 3DPDFConverter.exe]]} {
          if {![info exists dispApps($match)]} {
            set name "Tetra4D Converter"
            set dispApps($match) $name
          }
        }
      }
      set match [file join $pf Adobe "Acrobat $j" Acrobat plug_ins 3DPDFConverter 3DReviewer.exe]
      if {![info exists dispApps($match)]} {
        set name "Tetra4D Reviewer"
        set dispApps($match) $name
      }
    }
  }

# others
  set b1 [file join $myhome AppData Local IDA-STEP ida-step.exe]
  if {[file exists $b1]} {
    set name "IDA-STEP Viewer"
    set dispApps($b1) $name
  }
  set b1 [file join $drive CCELabs EnSuite-View Bin EnSuite-View.exe]
  if {[file exists $b1]} {
    set name "EnSuite-View"
    set dispApps($b1) $name
  } else {
    set b1 [file join $drive CCE EnSuite-View Bin EnSuite-View.exe]
    if {[file exists $b1]} {
      set name "EnSuite-View"
      set dispApps($b1) $name
    }
  }

#-------------------------------------------------------------------------------
# default viewer
  set dispApps(Default) "Default STEP Viewer"

# file tree view
  set dispApps(Indent) "Tree View (for debugging)"

#-------------------------------------------------------------------------------
# set text editor command and name
  set editorCmd ""
  set editorName ""

# Notepad++ or Notepad
  set editorCmd [file join $pf32 Notepad++ notepad++.exe]
  if {[file exists $editorCmd]} {
    set editorName "Notepad++"
    set dispApps($editorCmd) $editorName
  } elseif {[info exists env(windir)]} {
    set editorCmd [file join $env(windir) system32 Notepad.exe]
    set editorName "Notepad"
    set dispApps($editorCmd) $editorName
  }

#-------------------------------------------------------------------------------
# remove cmd that do not exist in dispCmds and non-executables
  set dispCmds1 {}
  foreach app $dispCmds {
    if {([file exists $app] || [string first "Default" $app] == 0 || [string first "Indent" $app] == 0) && \
         [file tail $app] != "NotePad.exe"} {
      lappend dispCmds1 $app
    }
  }
  set dispCmds $dispCmds1

# check for cmd in dispApps that does not exist in dispCmds and add to list
  foreach app [array names dispApps] {
    if {[file exists $app] || [string first "Default" $app] == 0 || [string first "Indent" $app] == 0} {
      set notInCmds 1
      foreach cmd $dispCmds {if {[string tolower $cmd] == [string tolower $app]} {set notInCmds 0}}
      if {$notInCmds} {lappend dispCmds $app}
    }
  }

# remove duplicates in dispCmds
  if {[llength $dispCmds] != [llength [lrmdups $dispCmds]]} {set dispCmds [lrmdups $dispCmds]}

# clean up list of app viewer commands
  if {[info exists dispCmd]} {
    if {([file exists $dispCmd] || [string first "Default" $dispCmd] == 0 || [string first "Indent" $dispCmd] == 0)} {
      if {[lsearch $dispCmds $dispCmd] == -1 && $dispCmd != ""} {lappend dispCmds $dispCmd}
    } else {
      if {[llength $dispCmds] > 0} {
        foreach dispCmd $dispCmds {
          if {([file exists $dispCmd] || [string first "Default" $dispCmd] == 0 || [string first "Indent" $dispCmd] == 0)} {break}
        }
      } else {
        set dispCmd ""
      }
    }
  } else {
    if {[llength $dispCmds] > 0} {
      set dispCmd [lindex $dispCmds 0]
    }
  }
  for {set i 0} {$i < [llength $dispCmds]} {incr i} {
    if {![file exists [lindex $dispCmds $i]] && [string first "Default" [lindex $dispCmds $i]] == -1 && [string first "Indent" [lindex $dispCmds $i]] == -1} {set dispCmds [lreplace $dispCmds $i $i]}
  }

# put dispCmd at beginning of dispCmds list
  if {[info exists dispCmd]} {
    for {set i 0} {$i < [llength $dispCmds]} {incr i} {
      if {$dispCmd == [lindex $dispCmds $i]} {
        set dispCmds [lreplace $dispCmds $i $i]
        set dispCmds [linsert $dispCmds 0 $dispCmd]
      }
    }
  }

# remove duplicates in dispCmds, again
  if {[llength $dispCmds] != [llength [lrmdups $dispCmds]]} {set dispCmds [lrmdups $dispCmds]}

# set list of STEP viewer names, appNames
  set appNames {}
  set appName  ""
  foreach cmd $dispCmds {
    if {[info exists dispApps($cmd)]} {
      lappend appNames $dispApps($cmd)
    } else {
      set name [file rootname [file tail $cmd]]
      lappend appNames  $name
      set dispApps($cmd) $name
    }
  }
  if {$dispCmd != ""} {
    if {[info exists dispApps($dispCmd)]} {set appName $dispApps($dispCmd)}
  }
}

# -------------------------------------------------------------------------------------------------
proc getFirstFile {} {
  global editorCmd openFileList buttons

  set localName [lindex $openFileList 0]
  if {$localName != ""} {
    outputMsg "\nReady to process: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)" green
    if {[info exists buttons(appOpen)]} {
      $buttons(appOpen) configure -state normal
      if {$editorCmd != ""} {
        bind . <Key-F5> {
          if {[file exists $localName]} {
            outputMsg "\nOpening STEP file: [file tail $localName]"
            exec $editorCmd [file nativename $localName] &
          }
        }
      }
    }
  }
  return $localName
}

#-------------------------------------------------------------------------------
proc addFileToMenu {} {
  global openFileList localName File buttons

  set lenlist 25
  set filemenuinc 4

  if {![info exists buttons]} {return}

# change backslash to forward slash, if necessary
  regsub -all {\\} $localName "/" localName

# remove duplicates
  set newlist {}
  set dellist {}
  for {set i 0} {$i < [llength $openFileList]} {incr i} {
    set name [lindex $openFileList $i]
    set ifile [lsearch -all $openFileList $name]
    if {[llength $ifile] == 1 || [lindex $ifile 0] == $i} {
      lappend newlist $name
    } else {
      lappend dellist $i
    }
  }
  set openFileList $newlist

# check if file name is already in the menu, if so, delete
  set ifile [lsearch $openFileList $localName]
  if {$ifile > 0} {
    set openFileList [lreplace $openFileList $ifile $ifile]
    $File delete [expr {$ifile+$filemenuinc}] [expr {$ifile+$filemenuinc}]
  }

# insert file name at top of list
  set fext [string tolower [file extension $localName]]
  if {$ifile != 0 && ($fext == ".stp" || $fext == ".step" || $fext == ".p21")} {
    set openFileList [linsert $openFileList 0 $localName]
    $File insert $filemenuinc command -label [truncFileName [file nativename $localName] 1] \
      -command [list openFile $localName] -accelerator "F1"
    catch {$File entryconfigure 5 -accelerator {}}
  }

# check length of file list, delete from the end of the list
  if {[llength $openFileList] > $lenlist} {
    set openFileList [lreplace $openFileList $lenlist $lenlist]
    $File delete [expr {$lenlist+$filemenuinc}] [expr {$lenlist+$filemenuinc}]
  }

# compare file list and menu list
  set llen [llength $openFileList]
  for {set i 0} {$i < $llen} {incr i} {
    set f1 [file tail [lindex $openFileList $i]]
    set f2 ""
    catch {set f2 [file tail [lindex [$File entryconfigure [expr {$i+$filemenuinc}] -label] 4]]}
  }

# save the state so that if the program crashes the file list will be already saved
  saveState
  return
}
