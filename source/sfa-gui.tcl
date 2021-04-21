# SFA version number
proc getVersion {} {return 4.40}

# version of SFA that the User Guide is based on
proc getVersionUG {} {return 4.2}

# IFCsvr version, depends on string entered when IFCsvr is repackaged for new STEP schemas
proc getVersionIFCsvr {} {return 20201208}

proc getContact {} {return [list "Robert Lipman" "robert.lipman@nist.gov"]}

# -------------------------------------------------------------------------------
proc whatsNew {} {
  global progtime sfaVersion upgrade

  if {$sfaVersion > 0 && $sfaVersion < [getVersion]} {
    outputMsg "\nThe previous version of the STEP File Analyzer and Viewer was: $sfaVersion" red
    set upgrade [clock seconds]
  }

# new user welcome message
  if {$sfaVersion == 0} {
    outputMsg "\nWelcome to the NIST STEP File Analyzer and Viewer\n" blue
    outputMsg "Please take a few minutes to read some the Help text so that you understand the options
available with the software.  Also explore the Examples and Websites menus.  The User
Guide is based on version [getVersionUG] of the software.  New and updated features are documented
in the Changelog and Help menu.

You will be prompted to install the IFCsvr toolkit which is required to read STEP files.
After the toolkit is installed, you are ready to process a STEP file.  Go to the File
menu, select a STEP file, and click the Generate Spreadsheet and View button.  If you
only want to generate a View of the STEP file, go to the Output Format section on the
Options tab and check View Only.

Use F9 and F10 to change the font size here.  See Help > Function Keys"
  }

  outputMsg "\nWhat's New (Version: [getVersion]  Updated: [string trim [clock format $progtime -format "%e %b %Y"]])" blue

# messages if SFA has already been run
  if {$sfaVersion > 0} {

# update the version number when IFCsvr is repackaged to include updated STEP schemas
    if {$sfaVersion < 4.32} {outputMsg "- The IFCsvr toolkit might need to be reinstalled.  Please follow the directions carefully." red}

    if {$sfaVersion < [getVersionUG]} {
      outputMsg "- A new User Guide is available based on version [getVersionUG] of this software."
      showFileURL UserGuide
    }
    if {$sfaVersion < 4.37} {outputMsg "- Updated Sample STEP Files on the Examples menu"}
    if {$sfaVersion < 4.12} {outputMsg "- Viewer for part geometry is faster and supports color, transparency, edges, sketch geometry, normals, and nested assemblies.  See Help > Viewer"}
    if {$sfaVersion < 3.80} {outputMsg "- Run the Syntax Checker with function key F8 or the Options tab selection.  See Help > Syntax Checker"}
    if {$sfaVersion < 2.62} {outputMsg "- Renamed output files: Spreadsheets from 'myfile_stp.xlsx' to 'myfile-sfa.xlsx' and Views from 'myfile-x3dom.html' to 'myfile-sfa.html'"}
    if {$sfaVersion < 2.30} {outputMsg "- Command-line version has been renamed: sfa-cl.exe  The old version STEP-File-Analyzer-CL.exe can be deleted."}
  }
  outputMsg "- All new features and bug fixes are listed in the Changelog.  See Help > Changelog"

  .tnb select .tnb.status
  update idletasks
}

#-------------------------------------------------------------------------------
# open local file or URL
proc showFileURL {type} {
  global sfaVersion

  switch -- $type {
    UserGuide {
# update for new versions, local and online
      set localFile "SFA-User-Guide-v6.pdf"
      set URL https://doi.org/10.6028/NIST.AMS.200-10
      if {$sfaVersion >= [expr {[getVersionUG]+0.1}]} {
        outputMsg "\nThe User Guide is based on version [getVersionUG] of this software.  See Help > Text Strings for information\nthat supplements the User Guide section 5.5 on Unicode Characters."
        .tnb select .tnb.status
      }
    }

    Changelog {
# local changelog file should also be on amazonaws
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

  wm title . "STEP File Analyzer and Viewer [getVersion]"
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
  ttk::style configure TLabelframe  -background $bgcolor
  ttk::style map       TLabelframe  -background [list disabled $bgcolor]

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

  bind . <Key-F2> {if {$lastXLS   != ""} {set lastXLS [openXLS $lastXLS  1]}}
  bind . <Key-F3> {if {$lastX3DOM != ""} {openX3DOM $lastX3DOM}}
  bind . <Key-F6> {openMultiFile 0}
  bind . <Key-F7> {if {$lastXLS1  != ""} {set lastXLS1 [openXLS $lastXLS1 1]}}
  bind . <Key-F8> {if {[info exists localName]} {if {[file exists $localName]} {syntaxChecker $localName}}}
  bind . <Key-F12> {if {$lastX3DOM != "" && [file exists $lastX3DOM]} {exec $editorCmd [file nativename $lastX3DOM] &}}

# scrolling status tab
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
  global buttons ftrans mytemp nprogBarEnts nprogBarFiles opt wdir

# generate button
  set ftrans [frame .ftrans1 -bd 2 -background "#F0F0F0"]
  set butstr "Spreadsheet"
  if {$opt(xlFormat) == "CSV"} {set butstr "CSV Files"}
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
  catch {
    set l3 [label $ftrans.l3 -relief flat -bd 0]
    $l3 config -image [image create photo -file [file join $wdir images nist.gif]]
    pack $l3 -side right -padx 10
    bind $l3 <ButtonRelease-1> {openURL https://www.nist.gov}
    tooltip::tooltip $l3 "Click here to learn more about NIST"
  }
  catch {[file copy -force -- [file join $wdir images NIST.ico] [file join $mytemp NIST.ico]]}

  pack $ftrans -side top -padx 10 -pady 10 -fill x

# progress bars
  set fbar [frame .fbar -bd 2 -background "#F0F0F0"]
  set nprogBarEnts 0
  set buttons(progressBar) [ttk::progressbar $fbar.pgb -mode determinate -variable nprogBarEnts]
  pack $fbar.pgb -side top -padx 10 -fill x

  set nprogBarFiles 0
  set buttons(progressBarMulti) [ttk::progressbar $fbar.pgb1 -mode determinate -variable nprogBarFiles]
  pack forget $buttons(progressBarMulti)
  pack $fbar -side bottom -padx 10 -pady {0 10} -fill x

# NIST icon bitmap
  catch {wm iconbitmap . -default [file join $wdir images NIST.ico]}
}

#-------------------------------------------------------------------------------
# status tab
proc guiStatusTab {} {
  global fout nb outputWin statusFont wout

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

# font type
  if {![info exists statusFont]} {
    set statusFont [$outputWin type cget black -font]
  }
  if {[string first "Courier" $statusFont] != -1} {
    regsub "Courier" $statusFont "Consolas" statusFont
    regsub "120" $statusFont "140" statusFont
  }

  if {[info exists statusFont]} {
    foreach typ {black red green magenta cyan blue error syntax} {$outputWin type configure $typ -font $statusFont}
  }

# function key bindings
  bind . <Key-F10> {
    set statusFont [$outputWin type cget black -font]
    for {set i 210} {$i >= 100} {incr i -10} {regsub -all $i $statusFont [expr {$i+10}] statusFont}
    foreach typ {black red green magenta cyan blue error syntax} {$outputWin type configure $typ -font $statusFont}
  }
  bind . <Control-KeyPress-=> {
    set statusFont [$outputWin type cget black -font]
    for {set i 210} {$i >= 100} {incr i -10} {regsub -all $i $statusFont [expr {$i+10}] statusFont}
    foreach typ {black red green magenta cyan blue error syntax} {$outputWin type configure $typ -font $statusFont}
  }

  bind . <Key-F9> {
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

  $File add command -label "Open STEP File(s)..." -accelerator "Ctrl+O" -command {openFile}
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
  global allNone buttons cb entCategory fopt fopta nb opt recPracNames

  set cb 0
  set wopt [ttk::panedwindow $nb.options -orient horizontal]
  $nb add $wopt -text " Options " -padding 2
  set fopt [frame $wopt.fopt -bd 2 -relief sunken]
  set fopta [ttk::labelframe $fopt.a -text " Process "]

# option to process user-defined entities
  guiUserDefinedEntities

  set fopta1 [frame $fopta.1 -bd 0]
  foreach item {{" Common"         opt(stepCOMM)} \
                {" Presentation"   opt(stepPRES)} \
                {" Representation" opt(stepREPR)} \
                {" Tolerance"      opt(stepTOLR)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $fopta1.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    if {[info exists entCategory($idx)]} {
      set str "most STEP APs."
      if {$idx == "stepTOLR"} {set str "AP214 and AP242."}
      set ttmsg "[string trim [lindex $item 0]] entities are found in $str"
      append ttmsg "  See Help > Supported STEP APs  and  Websites > STEP Format and Schemas\n\n"
      if {$idx != "stepCOMM"} {
        set ttmsg [guiToolTip $ttmsg $idx]
      } else {
        append ttmsg "Entity types from any selected Process category that are found in a STEP file are written to the Spreadsheet.  All AP-specific\nentities from APs other than AP203, AP214, and AP242 are always written to the Spreadsheet, including AP209, AP210, AP238,\nand AP239.  The Process categories are used to group and color-code entities on the Summary worksheet.\n\nSee Help > User Guide (section 3.4.2)"
      }
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta1 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta2 [frame $fopta.2 -bd 0]
  foreach item {{" Measure"      opt(stepQUAN)} \
                {" Shape Aspect" opt(stepSHAP)} \
                {" Geometry"     opt(stepGEOM)} \
                {" Coordinates"  opt(stepCPNT)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $fopta2.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    if {[info exists entCategory($idx)]} {
      set ttmsg "[string trim [lindex $item 0]] entities"
      if {$idx != "stepCPNT"} {append ttmsg " are found in most STEP APs."}
      append ttmsg "  See Help > Supported STEP APs  and  Websites > STEP Format and Schemas\n\n"
      set ttmsg [guiToolTip $ttmsg $idx]
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta2 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta3 [frame $fopta.3 -bd 0]
  foreach item {{" AP242"      opt(stepAP242)} \
                {" Composites" opt(stepCOMP)} \
                {" Features"   opt(stepFEAT)} \
                {" Kinematics" opt(stepKINE)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $fopta3.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    if {[info exists entCategory($idx)]} {
      if {$idx != "stepAP242"} {
        set ttmsg "[string trim [lindex $item 0]] entities"
        append ttmsg " are found in"
        if {$idx == "stepCOMP"} {
          append ttmsg " AP203 and AP242"
        } else {
          append ttmsg " AP214 and AP242"
        }
        append ttmsg ".  See Help > Supported STEP APs  and  Websites > STEP Format and Schemas\n\n"
        set ttmsg [guiToolTip $ttmsg $idx]
      } else {
        set ttmsg "These entities are only in AP242, however, the more commonly used AP242\nentities are found in the other Process categories."
        append ttmsg "\n\nSee Websites > AP242\nSee Help > Supported STEP APs  and  Websites > STEP Format and Schemas"
      }
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta3 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta4 [frame $fopta.4 -bd 0]
  set anbut [list {"All" 0} {"For Analysis" 2} {"For Views" 3} {"Reset" 1}]
  foreach item $anbut {
    set bn "allNone[lindex $item 1]"
    set buttons($bn) [ttk::radiobutton $fopta4.$cb -variable allNone -text [lindex $item 0] -value [lindex $item 1] \
      -command {
        if {$allNone == 0} {
          foreach item [array names opt] {
            if {[string first "step" $item] == 0 && $item != "stepGEOM" && $item != "stepCPNT" && $item != "stepUSER"} {set opt($item) 1}
          }
        } elseif {$allNone == 1} {
          foreach item [array names opt] {if {[string first "step" $item] == 0} {set opt($item) 0}}
          foreach item {INVERSE PMIGRF PMISEM valProp viewFEA viewPart viewPMI viewTessPart stepUSER x3dSave} {set opt($item) 0}
          set opt(stepCOMM) 1
          set ofNone 0
          set ofExcel 1
          if {$opt(xlFormat) == "None"} {set opt(xlFormat) "Excel"}
          if {!$ofCSV} {$buttons(ofExcel) configure -state normal}
        } elseif {$allNone == 2} {
          foreach item {PMISEM PMIGRF valProp} {set opt($item) 1}
        } elseif {$allNone == 3} {
          foreach item {viewFEA viewPart viewPMI viewTessPart} {set opt($item) 1}
        }
        checkValues
      }]
    pack $buttons($bn) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  catch {
    tooltip::tooltip $buttons(allNone0) "Selects most Process categories to write to the Spreadsheet\nSee Help > User Guide (section 3.4.2)"
    tooltip::tooltip $buttons(allNone1) "Deselects most Process categories and all Analyze and View options\nSee Help > User Guide (section 3.4.2)"
    tooltip::tooltip $buttons(allNone2) "Selects all Analyze options and associated Process categories to write to the Spreadsheet\nSee Help > User Guide (section 6)"
    tooltip::tooltip $buttons(allNone3) "Selects all View options and associated Process categories\nSee View Only in Output Format below\nSee Help > User Guide (section 4)"
  }
  pack $fopta4 -side left -anchor w -pady 0 -padx 0 -fill y
  pack $fopta -side top -anchor w -pady {5 2} -padx 10 -fill both

#-------------------------------------------------------------------------------
# report
  set foptRV [frame $fopt.rv -bd 0]
  set foptd [ttk::labelframe $foptRV.1 -text " Analyze "]

  set foptd1 [frame $foptd.1 -bd 0]
  foreach item {{" Validation Properties" opt(valProp)} \
                {" AP242 PMI Representation (Semantic PMI)" opt(PMISEM)} \
                {" Round dimensions and geometric tolerances" opt(PMISEMRND)} \
                {" Only Dimensions" opt(PMISEMDIM)} \
                {" PMI Presentation (Graphical PMI)" opt(PMIGRF)} \
                {" Generate Presentation Coverage worksheet" opt(PMIGRFCOV)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $foptd1.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    if {$idx == "PMISEMDIM" || $idx == "PMIGRFCOV" || $idx == "PMISEMRND"} {
      pack $buttons($idx) -side top -anchor w -padx {26 10} -pady 0 -ipady 0
    } else {
      pack $buttons($idx) -side top -anchor w -padx {5 10} -pady 0 -ipady 0
    }
    incr cb
  }
  pack $foptd1 -side top -anchor w -pady 0 -padx 0 -fill y

  pack $foptd -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true
  catch {
    tooltip::tooltip $buttons(valProp) "Geometric, assembly, PMI, annotation, attribute, tessellated, composite, and FEA\nvalidation properties, and semantic text are reported.  Properties are shown on\nthe 'property_definition' and other entities.  Some properties are reported only if\nthe analysis for Semantic PMI is selected.  Some properties might not be shown\ndepending on the value of Maximum Rows (Spreadsheet tab).\n\nSee Help > Analyze > Validation Properties\nSee Help > User Guide (section 6.3)\nSee Help > Analyze > Syntax Errors\nSee Examples > PMI Presentation, Validation Properties"
    tooltip::tooltip $buttons(PMISEM)  "Semantic PMI is the information necessary to represent geometric\nand dimensional tolerances without any graphical PMI.  It is shown\non dimension, tolerance, datum target, and datum entities.\nSemantic PMI is found mainly in STEP AP242 files.\n\nSee Help > Analyze > PMI Representation\nSee Help > User Guide (section 6.1)\nSee Help > Analyze > Syntax Errors\nSee Examples > Spreadsheet - PMI Representation\nSee Examples > Sample STEP Files\nSee Websites > AP242"
    tooltip::tooltip $buttons(PMIGRF)  "Graphical PMI is the geometric elements necessary to draw annotations.\nThe information is shown on 'annotation occurrence' entities.\n\nSee Help > Analyze > PMI Presentation\nSee Help > User Guide (section 6.2)\nSee Help > Analyze > Syntax Errors\nSee Examples > PMI Presentation, Validation Properties\nSee Examples > Part with PMI\nSee Examples > AP242 Tessellated Part with PMI\nSee Examples > Sample STEP Files"
    tooltip::tooltip $buttons(PMIGRFCOV) "The PMI Presentation Coverage worksheet counts the number of recommended names used from the\nRecommended Practice for Representation and Presentation of PMI (AP242), Section 8.4.  The names\ndo not have any semantic PMI meaning.  This worksheet was always generated before version 3.62\nwhen PMI Presentation was selected.\n\nSee Help > Analyze > PMI Coverage Analysis"
    tooltip::tooltip $buttons(PMISEMRND) "The number of decimal places for dimensions and geometric tolerances can be specified with\nvalue_format_type_qualifier in the STEP file.  By definition the qualifier always truncates the value.  This\noption rounds the value instead.\n\nFor example with the value 0.5625, the qualifier 'NR2 1.3' will truncate it to 0.562  However, rounding\nwill show 0.563\n\nRounding values might result in a better match to graphical PMI shown by the Viewer or to expected\nPMI in the NIST models (FTC 7, 8).\n\nSee Websites > Recommended Practice for $recPracNames(pmi242), section 5.4"
    tooltip::tooltip $buttons(PMISEMDIM) "Analyze only dimensional tolerances and no\ngeometric tolerances, datums, or datum targets."
  }

#-------------------------------------------------------------------------------
# view
  set foptv [ttk::labelframe $foptRV.9 -text " View "]

# part geometry
  set foptv20 [frame $foptv.20 -bd 0]
  foreach item {{" Part Geometry" opt(viewPart)} \
                {" Edges" opt(partEdges)} \
                {" Sketch" opt(partSketch)} \
                {" Normals" opt(partNormals)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $foptv20.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv20 -side top -anchor w -pady 0 -padx 0 -fill y

# part quality
  set foptv21 [frame $foptv.21 -bd 0]
  set buttons(partqual) [label $foptv21.l3 -text "Quality:"]
  pack $foptv21.l3 -side left -anchor w -padx 0 -pady 0 -ipady 0
  foreach item {{Low 4} {Normal 7} {High 9}} {
    set bn "partQuality[lindex $item 1]"
    set buttons($bn) [ttk::radiobutton $foptv21.$cb -variable opt(partQuality) -text [lindex $item 0] -value [lindex $item 1]]
    pack $buttons($bn) -side left -anchor w -padx 2 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv21 -side top -anchor w -pady 0 -padx {26 10} -fill y

# graphical pmi
  set foptv3 [frame $foptv.3 -bd 0]
  set item {" Graphical PMI" opt(viewPMI)}
  set idx [string range [lindex $item 1] 4 end-1]
  set buttons($idx) [ttk::checkbutton $foptv3.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
  incr cb
  pack $foptv3 -side top -anchor w -pady 0 -padx 0 -fill y

  set foptv4 [frame $foptv.4 -bd 0]
  set buttons(linecolor) [label $foptv4.l3 -text "PMI Color:"]
  pack $foptv4.l3 -side left -anchor w -padx 0 -pady 0 -ipady 0
  set gpmiColorVal {{"From File" 0} {"Black" 1} {"By View" 3} {"Random" 2}}
  foreach item $gpmiColorVal {
    set bn "gpmiColor[lindex $item 1]"
    set buttons($bn) [ttk::radiobutton $foptv4.$cb -variable opt(gpmiColor) -text [lindex $item 0] -value [lindex $item 1]]
    pack $buttons($bn) -side left -anchor w -padx 2 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv4 -side top -anchor w -pady 0 -padx {26 10} -fill y

# tessellated geometry
  set foptv6 [frame $foptv.6 -bd 0]
  foreach item {{" AP242 Tessellated Part Geometry" opt(viewTessPart)} \
                {"Generate Wireframe" opt(tessPartMesh)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $foptv6.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv6 -side top -anchor w -pady 0 -padx 0 -fill y

# finite element model
  set foptv7 [frame $foptv.7 -bd 0]
  foreach item {{" AP209 Finite Element Model" opt(viewFEA)} \
                {"Boundary conditions" opt(feaBounds)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $foptv7.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv7 -side top -anchor w -pady 0 -padx 0 -fill y

  set foptv8 [frame $foptv.8 -bd 0]
  foreach item {{"Loads" opt(feaLoads)} \
                {"Scale loads" opt(feaLoadScale)} \
                {"Displacements" opt(feaDisp)} \
                {"No vector tail" opt(feaDispNoTail)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $foptv8.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 2 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv8 -side top -anchor w -pady 0 -padx {26 10} -fill y

  pack $foptv -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true
  pack $foptRV -side top -anchor w -pady 0 -fill x
  catch {
    tooltip::tooltip $foptv20 "The view for part geometry supports color, transparency, edges, and\nsketch geometry.  The viewer does not support measurements.\n\nNormals improve the default smooth shading at the expense of slower\nprocessing and display.  Using High Quality and Normals results in the\nbest appearance for part geometry.\n\nSee Help > Viewer\n\nViews are shown in the default web browser.\nViews can be generated without generating a spreadsheet or CSV files.\nSee the Output Format option below.\n\nSee Help > View for other viewing features\nSee Examples > View Box Assembly and others\nSee Websites > STEP File Viewers"
    tooltip::tooltip $buttons(viewPMI) "Graphical PMI is supported in AP242, AP203, and AP214 files.\n\nSee Help > View > Graphical PMI\nSee Help > Viewer\nSee Help > User Guide (section 4.2)\nSee Examples > Part with PMI\nSee Examples > AP242 Tessellated Part with PMI\nSee Examples > Sample STEP Files"
    tooltip::tooltip $buttons(viewTessPart) "Tessellated part geometry is typically written to an AP242 file instead of\nor in addition to b-rep part geometry.  ** Parts in an assembly might\nhave the wrong position and orientation or be missing. **\n\nSee Help > View > AP242 Tessellated Part Geometry\nSee Help > Viewer\nSee Help > User Guide (section 4.3)\nSee Examples > AP242 Tessellated Part with PMI"
    tooltip::tooltip $buttons(tessPartMesh) "Generate a wireframe mesh based on the tessellated faces and surfaces."
    tooltip::tooltip $buttons(feaLoadScale) "The length of load vectors can be scaled by their magnitude.\nLoad vectors are always colored by their magnitude."
    tooltip::tooltip $buttons(feaDispNoTail) "The length of displacement vectors with a tail are scaled by\ntheir magnitude.  Vectors without a tail are not.\nDisplacement vectors are always colored by their magnitude.\nLoad vectors always have a tail."
    tooltip::tooltip $foptv21 "Quality controls the number of facets used for curved surfaces.\nFor example, the higher the quality the more facets around the\ncircumference of a cylinder.  Also, the higher the quality the longer\nit takes to generate the view and show in a web browser."
    tooltip::tooltip $foptv4  "For 'By View' PMI colors, each Saved View is set to a different color.  If there\nis only one or no Saved Views, then 'Random' PMI colors are used.\nFor 'Random' PMI colors, each 'annotation occurrence' is set to a different\ncolor to help differentiate one from another."
    set tt "FEM nodes, elements, boundary conditions, loads, and\ndisplacements found in AP209 files are shown.\n\nSee Help > View > AP209 Finite Element Model\nSee Help > Viewer\nSee Help > User Guide (section 4.4)\nSee Examples > AP209 Finite Element Model"
    tooltip::tooltip $foptv7 $tt
    tooltip::tooltip $foptv8 $tt
  }
}

#-------------------------------------------------------------------------------
# user-defined list of entities
proc guiUserDefinedEntities {} {
  global buttons cb opt fileDir fopta userEntityFile userEntityList

  set fopta6 [frame $fopta.6 -bd 0]
  set item {" User-Defined List: " opt(stepUSER)}
  set idx [string range [lindex $item 1] 4 end-1]
  set buttons($idx) [ttk::checkbutton $fopta6.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
  incr cb

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
        set opt(stepUSER) 0
        checkValues
      }
      .tnb select .tnb.status
    }
    checkValues
  }]
  pack $fopta6.$cb -side left -anchor w -padx 10
  incr cb
  catch {tooltip::tooltip $fopta6 "A User-Defined List is a plain text file with one STEP entity name per line.\n\nThis allows for more control to process only the required entity types,\nrather than process the board categories of entities above.\n\nIt is also useful when processing large files that might crash the software."}
  pack $fopta6 -side bottom -anchor w -pady 5 -padx 0 -fill y
}

#-------------------------------------------------------------------------------
# inverse relationships
proc guiInverse {} {
  global buttons cb fopt inverses opt

  set foptc [ttk::labelframe $fopt.c -text " Inverse Relationships "]
  set item {" Show Inverses and Backwards References (Used In) for PMI, Shape Aspect, Representation, Tolerance, and more" opt(INVERSE)}
  set idx [string range [lindex $item 1] 4 end-1]
  set buttons($idx) [ttk::checkbutton $foptc.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
  incr cb

  pack $foptc -side top -anchor w -pady {5 2} -padx 10 -fill both
  set ttmsg "Inverse Relationships and Backwards References (Used In) are reported for some attributes for the following entities.\nInverse or Used In values are shown in additional columns highlighted in light blue and purple.\n\nSee Help > User Guide (section 5.6.1)\n\n"
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
  global appName appNames buttons cb developer dispApps dispCmds edmWhereRules edmWriteToFile stepToolsWriteToFile
  global fopt foptf useXL xlInstalled

  set foptOP [frame $fopt.op -bd 0]
  set foptf [ttk::labelframe $foptOP.f -text " Open STEP File in App"]

  set buttons(appCombo) [ttk::combobox $foptf.spinbox -values $appNames -width 30]
  pack $foptf.spinbox -side left -anchor w -padx 7 -pady {0 3}
  bind $buttons(appCombo) <<ComboboxSelected>> {
    set appName [$buttons(appCombo) get]

# Jotne EDM Model Checker
    if {$developer} {
      catch {
        if {[string first "EDM Model Checker" $appName] == 0 || [string first "EDMsdk" $appName] != -1} {
          pack $buttons(edmWriteToFile) -side left -anchor w -padx {5 0}
          pack $buttons(edmWhereRules) -side left -anchor w -padx {5 0}
        } else {
          pack forget $buttons(edmWriteToFile)
          pack forget $buttons(edmWhereRules)
        }
      }
    }
# STEP Tools
    catch {
      if {[string first "Conformance Checker" $appName] != -1} {
        pack $buttons(stepToolsWriteToFile) -side left -anchor w -padx {5 0}
      } else {
        pack forget $buttons(stepToolsWriteToFile)
      }
    }
# file tree view
    catch {
      if {$appName == "Tree View (for debugging)"} {
        pack $buttons(indentStyledItem) -side left -anchor w -padx {5 0}
        pack $buttons(indentGeometry) -side left -anchor w -padx {5 0}
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
      if {[string first "EDM Model Checker" $item] == 0 || [string first "EDMsdk" $item] != -1} {
        foreach item {{"Check rules" edmWhereRules} \
                      {"Write to file" edmWriteToFile}} {
          set idx [lindex $item 1]
          set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
          pack forget $buttons($idx)
          incr cb
        }
      }
    }
  }

# STEP Tools
  if {[lsearch -glob $appNames "*Conformance Checker*"] != -1} {
    foreach item {{"Write to file" stepToolsWriteToFile}} {
      regsub -all {[\(\)]} [lindex $item 1] "" idx
      set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
      pack forget $buttons($idx)
      incr cb
    }
  }

# built-in file tree view
  if {[lsearch $appNames "Tree View (for debugging)"] != -1} {
    foreach item {{"Include Geometry" indentGeometry} \
                  {"Include styled_item" indentStyledItem}} {
      set idx [lindex $item 1]
      set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] -variable opt([lindex $item 1]) -command {checkValues}]
      pack forget $buttons($idx)
      incr cb
    }
  }

  catch {tooltip::tooltip $buttons(appCombo) "This option is a convenient way to open a STEP file in other apps.  The\npull-down menu contains some apps that can open a STEP file such as\nSTEP viewers and browsers, however, only if they are installed in their\ndefault location.\n\nSee Help > Open STEP File in App\nSee Websites > STEP File Viewers\n\nThe 'Tree View (for debugging)' option rearranges and indents the entities\nto show the hierarchy of information in a STEP file.  The 'tree view' file\n(myfile-sfa.txt) is written to the same directory as the STEP file or to the\nsame user-defined directory specified in the Spreadsheet tab.  Including\nGeometry or Styled_item can make the 'tree view' file very large.  The\n'tree view' might not process /*comments*/ in a STEP file correctly.\n\nThe 'Default STEP Viewer' option opens the STEP file in whatever\napp is associated with STEP (.stp, .step, .p21) files.\n\nUse F5 to open the STEP file in a text editor."}
  pack $foptf -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true

#-------------------------------------------------------------------------------
# syntax checker
  set foptl [ttk::labelframe $foptOP.l -text " Syntax Checker "]
  set txt " Run Syntax Checker"
  set idx syntaxChecker
  set buttons($idx) [ttk::checkbutton $foptl.$cb -text $txt -variable opt($idx) -command {checkValues}]
  pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
  incr cb
  pack $foptl -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true
  catch {tooltip::tooltip $foptl "Select this option to run the Syntax Checker when generating a Spreadsheet\nor View.  The Syntax Check can also be run with function key F8.\n\nIt checks for basic syntax errors and warnings in the STEP file related to\nmissing or extra attributes, incompatible and unresolved\ entity references,\nselect value types, illegal and unexpected characters, and other problems\nwith entity attributes.\n\nSee Help > Syntax Checker\nSee Help > User Guide (section 7)"}
  pack $foptOP -side top -anchor w -pady 0 -fill x

#-------------------------------------------------------------------------------
# output format
  set foptOF [frame $fopt.of -bd 0]
  set foptk [ttk::labelframe $foptOF.k -text " Output Format "]
  set idx outputOpen
  set buttons($idx) [ttk::checkbutton $foptk.$cb -text " Open Output Files" -variable opt(outputOpen)]
  pack $buttons($idx) -side bottom -anchor w -padx 5 -pady 0 -ipady 0
  incr cb

# checkbuttons are used for pseudo-radiobuttons
  foreach item {{" Spreadsheet" ofExcel} \
                {" CSV Files  " ofCSV} \
                {" View Only" ofNone}} {
    set idx [lindex $item 1]
    set buttons($idx) [ttk::checkbutton $foptk.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {
      if {![info exists useXL]} {set useXL 1}
      if {[info exists xlInstalled]} {
        if {!$xlInstalled} {set useXL 0}
      } else {
        set xlInstalled 1
      }

      if {$ofNone && $opt(xlFormat) != "None"} {
        set ofExcel 0
        set ofCSV 0
        set opt(xlFormat) "None"
        set allNone -1
        if {$useXL && $xlInstalled} {$buttons(ofExcel) configure -state normal}
      }
      if {$ofExcel && $opt(xlFormat) != "Excel"} {
        set ofNone 0
        if {$useXL} {
          set ofCSV 0
          set opt(xlFormat) "Excel"
        } else {
          set ofExcel 0
          set ofCSV 1
          set opt(xlFormat) "CSV"
        }
      }
      if {$ofCSV} {
        if {$useXL} {
          set ofExcel 1
          $buttons(ofExcel) configure -state disabled
        }
        if {$opt(xlFormat) != "CSV"} {
          set ofNone 0
          set opt(xlFormat) "CSV"
        }
      } elseif {$xlInstalled} {
        $buttons(ofExcel) configure -state normal
      }
      if {!$ofExcel && !$ofCSV && !$ofNone} {
        if {$useXL} {
          set ofExcel 1
          set opt(xlFormat) "Excel"
          $buttons(ofExcel) configure -state normal
        } else {
          set ofCSV 1
          set opt(xlFormat) "CSV"
          $buttons(ofExcel) configure -state disabled
        }
      }
      checkValues
    }]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }

# part only
  set item {" Part Only" opt(partOnly)}
  set idx [string range [lindex $item 1] 4 end-1]
  set buttons($idx) [ttk::checkbutton $foptk.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
  incr cb
  pack $foptk -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true

  catch {tooltip::tooltip $foptk "If Excel is installed, then Spreadsheets and CSV files can be generated.  If CSV Files\nis selected, the Spreadsheet is also generated.  CSV files do not contain any cell\ncolors, comments, or links.  GD&T symbols in CSV files are only supported with\nExcel 2016 or newer.\n\nIf Excel is not installed, only CSV files can be generated.  Options for Analyze and\nInverse Relationships are disabled.\n\nView Only does not generate any Spreadsheets or CSV files.  All options except\nthose for View are disabled.  Part Only generates only Part Geometry.  This is\nuseful when no other View features of the software are needed and for large STEP\nfiles.\n\nIf output files are not opened after they have been generated, use F2 to open a\nSpreadsheet and F3 to open a View.  Use F7 to open the File Summary spreadsheet \nwhen processing multiple files.\n\nIf possible, existing output files are always overwritten by new files.  Output files\ncan be written to a user-defined directory.  See Spreadsheet tab.\n\nSee Help > User Guide (section 3.4.1)"}

# log file
  set foptm [ttk::labelframe $foptOF.m -text " Log File "]
  set txt " Generate a Log File of the text in the Status tab"
  set idx logFile
  set buttons($idx) [ttk::checkbutton $foptm.$cb -text $txt -variable opt($idx)]
  pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
  pack $foptm -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true
  pack $foptOF -side top -anchor w -pady 0 -fill x
  catch {tooltip::tooltip $buttons(optlogFile)  "The Log file is written to myfile-sfa.log  Use F4 to open the Log file.\nSyntax Checker results are written to myfile-sfa-err.log  See Help > Syntax Checker\nAll text in the Status tab can be saved by right-clicking and selecting Save."}
}

#-------------------------------------------------------------------------------
# spreadsheet tab
proc guiSpreadsheet {} {
  global buttons cb developer excelVersion fileDir fxls mydocs nb opt pmiElementsMaxRows userWriteDir writeDir

  set wxls [ttk::panedwindow $nb.xls -orient horizontal]
  $nb add $wxls -text " Spreadsheet " -padding 2
  set fxls [frame $wxls.fxls -bd 2 -relief sunken]

# tables for sorting
  set fxlsz [ttk::labelframe $fxls.z -text " Tables "]
  set item {" Generate Tables for Sorting and Filtering" opt(xlSort)}
  set idx [string range [lindex $item 1] 4 end-1]
  set buttons($idx) [ttk::checkbutton $fxlsz.$cb -text [lindex $item 0] -variable [lindex $item 1]]
  pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
  incr cb
  pack $fxlsz -side top -anchor w -pady {5 2} -padx 10 -fill both
  set msg "Worksheets can be sorted by column values.\nThe worksheet with Properties is always sorted.\n\nSee Help > User Guide (section 5.6.2)"
  catch {tooltip::tooltip $fxlsz $msg}

# number format
  set fxlsa [ttk::labelframe $fxls.a -text " Number Format "]
  set item {" Do not round real numbers in spreadsheet cells" opt(xlNoRound)}
  set idx [string range [lindex $item 1] 4 end-1]
  set buttons($idx) [ttk::checkbutton $fxlsa.$cb -text [lindex $item 0] -variable [lindex $item 1]]
  pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
  incr cb
  pack $fxlsa -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {
    tooltip::tooltip $buttons(xlNoRound) "Excel rounds real numbers if there are more than 11 characters in the number string.\nFor example, the number 0.1249999999997 in a STEP file is shown as 0.125 in a\nspreadsheet cell.  Clicking in a cell with a rounded number shows all of the digits\nin the formula bar.\n\nSelecting this option will show most single real numbers exactly as they appear in\nthe STEP file.  Numbers not rounded are left-justified in a cell.  Lists of real numbers,\nsuch as cartesian points, are always shown exactly as they appear in the STEP file.\n\nSee Help > User Guide (section 5.6.3)"
  }

# maximum rows
  set fxlsb [ttk::labelframe $fxls.b -text " Maximum Rows for any worksheet"]
  set rlimit {{" 100" 103} {" 500" 503} {" 1000" 1003} {" 5000" 5003} {" 10000" 10003} {" 50000" 50003} {" 100000" 100003} {" Maximum" 1048576}}
  if {$excelVersion < 12} {
    set rlimit [lrange $rlimit 0 5]
    lappend rlimit {" Maximum" 65536}
  }
  foreach item $rlimit {
    pack [ttk::radiobutton $fxlsb.$cb -variable opt(xlMaxRows) -text [lindex $item 0] -value [lindex $item 1]] -side left -anchor n -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsb -side top -anchor w -pady 5 -padx 10 -fill both
  set msg "This option limits the number of rows (entities) written to any one worksheet or CSV file.\nIf the maximum number of rows is exceeded, the number of entities processed will be\nreported as, for example, 'property_definition (100 of 147)'.\n\nFor large STEP files, setting a low maximum can speed up processing at the expense of\nnot processing all of the entities.  This is useful when processing Geometry entities.\n\nSyntax Errors might be missed if some entities are not processed due to a smaller value\nof maximum rows.  Maximum rows does not affect generating Views.  The maximum\nnumber of rows depends on the version of Excel.\n\nSee Help > User Guide (section 5.6.4)"
  catch {tooltip::tooltip $fxlsb $msg}

# output directory (opt(writeDirType) = 1 no longer used)
  set fxlsd [ttk::labelframe $fxls.d -text " Write Output to "]
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
  catch {tooltip::tooltip $fxls1 "This option is useful when the directory containing the STEP file is\nprotected (read-only) and none of the output can be written to it.\nDo not use a directory name containing bracket \[\] characters."}
  incr cb

  set buttons(userentry) [ttk::entry $fxls1.entry -width 50 -textvariable userWriteDir]
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

  pack $fxlsd -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {tooltip::tooltip $fxlsd "If possible, existing output files are always overwritten by new files.\nIf spreadsheets cannot be overwritten, a number is appended to the\nfile name: myfile-sfa-1.xlsx"}

# some other options
  set fxlsc [ttk::labelframe $fxls.c -text " Other "]
  foreach item {{" Process text strings with symbols and non-English characters" opt(xlUnicode)} \
                {" Show all PMI Elements on PMI Representation Coverage worksheets" opt(SHOWALLPMI)} \
                {" Do not generate links to STEP files and spreadsheets on File Summary worksheet for multiple files" opt(xlHideLinks)} \
                {" Save X3D file generated by the Viewer" opt(x3dSave)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $fxlsc.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsc -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {
    tooltip::tooltip $buttons(xlUnicode)   "See Help > Text Strings\n\nOnly use this option if there are non-English characters\nencoded with the \\X2\\ control directive in the STEP file.\nThis option can slow down the processing of the STEP file."
    tooltip::tooltip $buttons(SHOWALLPMI)  "The complete list of [expr {$pmiElementsMaxRows-3}] PMI Elements, including those that are not found in\nthe STEP file, will be shown on the PMI Representation Coverage worksheet.\n\nSee Help > Analyze > PMI Coverage Analysis\nSee Help > User Guide (section 6.1.7)"
    tooltip::tooltip $buttons(x3dSave)     "The X3D file can be shown in an X3D viewer or imported to other software.\nSee Help > Viewer"
    tooltip::tooltip $buttons(xlHideLinks) "This option is useful when sharing a Spreadsheet with another user."
  }

# developer only options
  if {$developer} {
    set fxlsx [ttk::labelframe $fxls.x -text " Debug "]
    foreach item {{" View" opt(DEBUGX3D)} \
                  {" Analysis" opt(DEBUG1)} \
                  {" Inverses" opt(DEBUGINV)}} {
      set idx [string range [lindex $item 1] 4 end-1]
      set buttons($idx) [ttk::checkbutton $fxlsx.$cb -text [lindex $item 0] -variable [lindex $item 1]]
      pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
      incr cb
    }
    pack $fxlsx -side top -anchor w -pady {5 2} -padx 10 -fill both
    catch {tooltip::tooltip $fxlsx "Debug options are only available on computers in the NIST domain."}
  }

  pack $fxls -side top -fill both -expand true -anchor nw
}

#-------------------------------------------------------------------------------
# help menu
proc guiHelpMenu {} {
  global contact defaultColor Examples excelVersion filesProcessed Help ifcsvrDir ifcsvrVer
  global mytemp nistVersion opt recPracNames scriptName stepAPs webCmd

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
outputMsg "\nOverview ------------------------------------------------------------------------------------------" blue
outputMsg "The STEP File Analyzer and Viewer (SFA) opens a STEP file (ISO 10303 - STandard for Exchange of
Product model data) Part 21 file (.stp or .step or .p21 file extension) and

1 - generates an Excel spreadsheet or CSV files of all entity and attribute information,
2 - creates a visualization (view) of part geometry, graphical PMI, and other features that is
    displayed in a web browser,
3 - reports and analyzes validation properties, semantic PMI, and graphical PMI and for conformance
    to recommended practices, and
4 - checks for basic syntax errors.

The four different types of output can be selected in the Options tab.  If you are interested in
only using the Viewer and not generating a spreadsheet, select View Only on the Options tab.

Help is available here, in the User Guide, and in tooltip help.  New features might not be
described in the User Guide.  Check the Changelog for recent updates to the software."
    .tnb select .tnb.status
  }

# options help
  $Help add command -label "Options" -command {
outputMsg "\nOptions -------------------------------------------------------------------------------------------" blue
outputMsg "See Help > User Guide (sections 3.4, 3.5, 4, and 6)

Process: Select which types of entities are processed.  The tooltip help lists all the entities
associated with that type.  Selectively process only the entities or views relevant to your
analysis.  Entity types and views can also be selected with the All, For Analysis, and For Views
buttons.

The CAx-IF Recommended Practices are checked for any of the Analyze options.
- PMI Representation Analysis: Dimensional tolerances, geometric tolerances, and datum features are
  reported on various entities indicated by PMI Representation on the Summary worksheet.
- PMI Presentation Analysis: Geometric entities used for PMI Presentation annotations are reported.
  Associated Saved Views, Validation Properties, and Geometry are also reported.
- Validation Properties Analysis: Geometric, assembly, PMI, annotation, attribute, and tessellated
  validation properties are reported.

View: Part geometry, graphical PMI annotations, tessellated part geometry in AP242 files, and AP209
finite element models can be shown in a web browser.

Inverse Relationships: For some entities, Inverse relationships and backwards references (Used In)
are shown on the worksheets.

Output Format: Generate Excel spreadsheets, CSV files, or only Views.  If Excel is not installed,
CSV files are automatically generated.  Some options are not supported with CSV files.  The View
Only option does not generate spreadsheets or CSV files.  The Syntax Checker can also be run when
processing a STEP file.

Spreadsheet tab:
- Table: Generate tables for each spreadsheet to facilitate sorting and filtering.
- Number Format: Option to not round real numbers.
- Maximum Rows: The maximum number of rows for any worksheet can be set lower than the normal
  limits for Excel.

All text in the Status tab can be written to a Log File when a STEP file is processed (Options tab).
The log file is written to myfile-sfa.log.  In the log file, syntax errors are highlighted by ***
and warnings and other messages are highlighted by **.  Use F4 to open the log file."
    .tnb select .tnb.status
  }

# general viewer help
  $Help add command -label "Viewer" -command {
outputMsg "\nViewer --------------------------------------------------------------------------------------------" blue
outputMsg "A View is written to an HTML file 'myfile-sfa.html' and is shown in the default web browser.
x3dom (x3dom.org) is used to show and navigate 3D models in a web browser.  x3dom requires an
Internet connection.  The HTML file is self-contained and can be shared with other users, including
those on non-Windows systems.  The viewer does not support measuring the model.

Views can be generated without generating a spreadsheet or CSV files.  See the Output Format on the
Options tab.  The Part Only option is useful when no other View features of the software are needed
and for large STEP files.

The viewer supports part geometry with color, transparency, part edges, sketch geometry, and nested
assemblies.  Part geometry viewer features:

- Part edges are shown in black.  Use the transparency slider to show only edges.  Transparency
  for parts is only approximate.  Parts inside of assemblies may not be visible.  This is a
  limitation of x3dom.  If a part is defined to be completely transparent in the STEP file and
  edges are not selected, then the part will not be visible in the viewer.  See Examples >
  View Box Assembly for an example of part transparency.

- Sketch geometry is supplemental lines created when generating a CAD model.  Sketch geometry is
  also known as construction, auxiliary, support, or reference geometry.  To view only sketch
  geometry in the viewer, turn off edges and make the part completely transparent.  Sketch geometry
  is not same as supplemental geometry.  See Help > View > Supplemental Geometry

- Nested assemblies have one STEP file that contains the assembly structure with external file
  references to individual assembly components that contain part geometry.
  See Examples > STEP File Library > External References

- Normals improve the default smooth shading by explicitly computing surface normals to improve the
  appearance of curved surfaces.  Computing normals takes more time to process and show in the web
  browser.

- Quality controls the number of facets used for curved surfaces.  Higher quality uses more facets
  around the circumference of a cylinder.  Using High Quality and the Normals options results in
  the best appearance for part geometry.

- Most assemblies and parts can be switched on and off depending on the assembly structure.  An
  alphabetic list of part and assembly names is shown on the right.  Parts with the same shape are
  usually grouped with the same checkbox.  Clicking on the model shows the part name in the upper
  left.  The part name shown may not be in the list of assemblies and parts.  The part might be
  contained in a higher-level assembly that is in the list.  Some names in the list might have an
  underscore and number appended to their name.  Processing sketch geometry might also affect the
  list of names.  Some assemblies have no unique names assigned to parts, therefore there is no
  list of part names.

- The part bounding box min and max XYZ coordinates are based on the faceted geometry being shown
  and not the exact geometry in the STEP file.  There might be a variation in the coordinates
  depending on the Quality option.  The bounding box also accounts for any sketch geometry if it is
  displayed but not graphical PMI and supplemental geometry.  The bounding box can be shown in the
  viewer to confirm that the min and max coordinates are correct.  If the part is too large to
  rotate smoothly in the viewer, turn off the part and rotate the bounding box.

- See Help > Text Strings for how non-English characters are handled in the Viewer.

- The origin of the model at '0 0 0' is shown with a small XYZ coordinate axis that can be switched
  off.  The background color can be changed between white, blue, gray, and black.

In the web browser, use Page Down to switch between front, side, top, and orthographic viewpoints.
Use key 'a' to view all and 'r' to restore to the original view.  The function of other keys is
described in the link 'Use the mouse'.  Navigation uses the Examine Mode.

Sometimes a part might be located far away from the origin and not visible.  In this case, turn off
the Origin and Sketch Geometry.  Then use 'a' to view all.

For very large STEP files, it might take several minutes to process the STEP part geometry.  In the
View section on the Options tab, uncheck Edges and Sketch, and select Quality Low.  For Output
Format, select View Only and Part Only.  The resulting HTML file might also take several minutes to
display in the web browser.  Select 'Wait' if the web browser prompts that it is running slowly
when opening the HTML file.

The viewer generates an X3D file that is embedded in the HTML file, thus creating the x3dom file
that is shown in the web browser.  Select 'Save X3D ...' on the Spreadsheet tab to save the X3D
file so that it can be shown in an X3D viewer or imported to other software.  Part geometry
including tessellated geometry and graphical PMI is supported.

The viewer can also be used with ASCII STL files used for 3D printing.  The STL file is first
converted to a STEP file containing AP242 tessellated geometry and then processed by the viewer.

See Help > User Guide (section 4)
See Help > View for more information about viewing:
- Supplemental Geometry
- Datum Targets
- Graphical PMI
- AP242 Tessellated Geometry
- AP209 Finite Element Models

See Examples > View Box Assembly and others

The viewer for part geometry is based on the NIST STEP to X3D Translator.
See Websites > STEP Software

Other STEP file viewers are available.  See Websites > STEP File Viewers.  Some of the other
viewers have better features for viewing and measuring part geometry.  However, many of the other
viewers cannot view graphical PMI, sketch geometry, supplemental geometry, datum targets, AP242
tessellated part geometry, and AP209 finite element models."
    .tnb select .tnb.status
  }

  $Help add separator
  $Help add cascade -label "View" -menu $Help.1
  set helpView [menu $Help.1 -tearoff 1]

  $helpView add command -label "Graphical PMI" -command {
outputMsg "\nGraphical PMI -------------------------------------------------------------------------------------" blue
outputMsg "Graphical PMI (PMI Presentation) annotations composed of polylines, lines, circles, and tessellated
geometry are supported for viewing.  The color of the annotations can be modified.  PMI associated
with Saved Views can be switched on and off.  Some Graphical PMI might not have equivalent or any
Semantic PMI in the STEP file.  Some STEP files with Semantic PMI might not have any Graphical PMI.

See Help > User Guide (section 4.2)
See Help > Analyze > PMI Presentation
See Examples > Part with PMI
See Examples > Sample STEP Files
See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242,
  PMI Polyline Presentation for AP203 and AP214)"
    .tnb select .tnb.status
  }

  $helpView add command -label "Supplemental Geometry" -command {
outputMsg "\nSupplemental Geometry -----------------------------------------------------------------------------" blue
outputMsg "Supplemental geometry is shown only if Part Geometry or Graphical PMI is also viewed.  Supplemental
geometry is not associated with graphical PMI Saved Views.

The following types of supplemental geometry and associated text are supported.
- Coordinate System: red/green/blue axes or by axes color
- Plane: blue transparent outlined surface
- Cylinder: blue transparent cylinder
- Line/Circle/Ellipse: purple line/circle/ellipse
- Point: black dot
- Tessellated Surface: assigned color

Trimming lines and circles with cartesian_point is not supported.  Unbounded planes are with shown
with a square surface.  Supplemental geometry can be switched on and off in the viewer.

See Websites > CAx Recommended Practices (Supplemental Geometry)"
    .tnb select .tnb.status
  }

  $helpView add command -label "Datum Targets" -command {
outputMsg "\nDatum Targets -------------------------------------------------------------------------------------" blue
outputMsg "Datum targets are shown only if a spreadsheet is generated with Analyze for Semantic PMI selected
and Part Geometry or Graphical PMI is also viewed.  There are two methods to represent and view the
position, orientation, and dimensions of a datum target.

1 - The position, orientation, and target length, width, and diameter are specified with the
placed_datum_target_feature entity.  Point, line, circle, circular curve, and rectangle datum
targets are supported.  A small coordinate axes is shown at the origin of a datum target except
for point datum targets.

2 - The shape and location of arbitrarily shaped area and curve datum targets is specified with
geometric entities referred to by the datum_target entity.  Supported geometric entities, that lie
in a plane, are line, circle, trimmed_curve, and advanced_face bounded by lines, circles, or
ellipses.  If other geometric entities are used, then either the datum target will not be shown or
some of the edges of the datum targets will be missing.  Datum targets defined by multiple types of
curves are not supported.

Both types of datum targets are shown in red and can be switched on and off in the viewer.

Datum target feature geometry (feature_for_datum_target_relationship), also specified with
geometric entities similar to the second method, is shown in green.

See Examples > Part with PMI
See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242, Sec. 6.6)"
    .tnb select .tnb.status
  }

  $helpView add command -label "AP242 Tessellated Part Geometry" -command {
outputMsg "\nAP242 Tessellated Part Geometry -------------------------------------------------------------------" blue
outputMsg "Tessellated part geometry is supported by AP242 and is usually supplementary to part geometry.

** Parts in an assembly might have the wrong position and orientation or be missing. **

Lines generated from tessellated edges are also shown.  A wireframe mesh, outlining the facets of
the tessellated surfaces can also be shown.  If both are present, tessellated edges might be
obscured by the wireframe mesh.  [string totitle [lindex $defaultColor 1]] is used for tessellated solids, shells, or faces that do not
have colors specified.  Clicking on a part with show the part name.

See Help > User Guide (section 4.3)
See Examples > AP242 Tessellated Part with PMI
See Websites > CAx Recommended Practices (Tessellated 3D Geometry)"
    .tnb select .tnb.status
  }

  $helpView add command -label "AP209 Finite Element Model" -command {
outputMsg "\nAP209 Finite Element Model ------------------------------------------------------------------------" blue
outputMsg "All AP209 entities are always processed and written to a spreadsheet unless a User-defined list is
used.

The AP209 finite element model composed of nodes, mesh, elements, boundary conditions, loads, and
displacements are shown and can be toggled on and off in the viewer.  Internal faces for solid
elements are not supported.

Nodal loads and element surface pressures are shown.  Load vectors are colored by their magnitude.
The length of load vectors can be scaled by their magnitude.  Forces use a single-headed arrow.
Moments use a double-headed arrow.

Displacement vectors are colored by their magnitude.  The length of displacement vectors can be
scaled by their magnitude depending on if they have a tail.  The finite element mesh is not
deformed.

Boundary conditions for translation DOF are shown with a red, green, or blue line along the X, Y,
or Z axes depending on the constrained DOF.  Boundary conditions for rotation DOF are shown with a
red, green, or blue circle around the X, Y, or Z axes depending on the constrained DOF.  A gray box
is used when all six DOF are constrained.  A gray pyramid is used when all three translation DOF
are constrained.  A gray sphere is used when all three rotation DOF are constrained.

Stresses, strains, and multiple coordinate systems are not support.

Setting Maximum Rows (Spreadsheet tab) does not affect the view.  For large AP209 files, there
might be insufficient memory to process all of the elements, loads, displacements, and boundary
conditions.

See Help > User Guide (section 4.4)
See Examples > AP209 Finite Element Model
See Websites > AP209 FEA"
    .tnb select .tnb.status
  }

  $helpView add command -label "Holes" -command {
outputMsg "\nHoles ---------------------------------------------------------------------------------------------" blue
outputMsg "Hole features, including basic round, counterbore, and countersink holes, and spotface are
supported in AP242 edition 2 but have not been widely implemented.  Semantic information related to
holes is reported on *_hole_definition and basic_round_hole worksheets.

If the Analyze report for Semantic PMI is not generated, then holes are shown only with a drill
entry point.  If the report is generated, then cylindrical or conical surfaces are used to view the
depth and diameter of the hole, counterbore, and countersink.  If there is no depth associated with
the hole (a through hole), then a very thin cylindrical surface with the correct diameter is shown.
The bottom of a hole is also shown if the hole is not a through hole and the hole has a depth.
Usually, only the counterbore or countersink is shown for through holes.

Holes can be switched on and off in the viewer.  Cylindrical surfaces are green and conical
surfaces are blue.  Holes are viewed regardless of whether they were explicitly modeled in the part
geometry.

In the Process section on the Options tab, Features is automatically selected when holes are found
in the STEP file."
    .tnb select .tnb.status
  }

  $Help add cascade -label "Analyze" -menu $Help.0
  set helpAnalyze [menu $Help.0 -tearoff 1]

# validation properties, PMI, conformance checking help
  $helpAnalyze add command -label "Validation Properties" -command {
outputMsg "\nValidation Properties -----------------------------------------------------------------------------" blue
outputMsg "Geometric, assembly, PMI, annotation, attribute, tessellated, composite, and FEA validation
properties, and semantic text are reported.  The property values are reported in columns
highlighted in yellow and green on the property_definition worksheet.  The worksheet can also be
sorted and filtered.  All properties might not be shown depending on the Maximum Rows set on the
Spreadsheet tab.

Validation properties are also reported on their associated annotation, dimension, geometric
tolerance, and shape aspect entities.  The report includes the validation property name and names
of the properties.  Some properties are reported only if the analysis for Semantic PMI is selected.
Other properties and user defined attributes are also reported.

Syntax errors related to validation property attribute values are also reported in the Status tab
and the relevant worksheet cells.  Syntax errors are highlighted in red.
See Help > Analyze > Syntax Errors

Clicking on the plus '+' symbols above the columns shows other columns that contain the entity ID
and attribute name of the validation property value.  All of the other columns can be shown or
hidden by clicking the '1' or '2' in the upper right corner of the spreadsheet.

The Summary worksheet indicates if properties are reported on property_definition and other
entities.

See Help > User Guide (section 6.3)
See Examples > PMI Presentation, Validation Properties
See Websites > CAx Recommended Practices (Geometric and Assembly Validation Properties,
  User Defined Attributes, Representation and Presentation of PMI for AP242,
  Tessellated 3D Geometry)"
    .tnb select .tnb.status
  }

  $helpAnalyze add command -label "PMI Representation (Semantic PMI)" -command {
outputMsg "\nPMI Representation --------------------------------------------------------------------------------" blue
outputMsg "PMI Representation (Semantic PMI) includes all information necessary to represent geometric and
dimensional tolerances (GD&T) without any graphical presentation elements.  PMI Representation is
associated with CAD model geometry and is computer-interpretable to facilitate automated
consumption by downstream applications for manufacturing, measurement, inspection, and other
processes.  PMI Representation is found mainly in AP242 files.

Worksheets for the analysis of PMI Representation show a visual recreation of the representation
for Dimensional Tolerances, Geometric Tolerances, and Datum Features.  The results are in columns,
highlighted in yellow and green, on the relevant worksheets.  The GD&T is recreated as best as
possible given the constraints of Excel.

All of the visual recreation of Datum Systems, Dimensional Tolerances, and Geometric Tolerances
that are reported on individual worksheets are collected on one PMI Representation Summary
worksheet.

If STEP files from the NIST CAD models (Websites > MBE PMI Validation Testing) are processed, then
the PMI the visual recreation of the PMI Representation is color-coded by the expected PMI in each
CAD model.  See Help > Analyze > NIST CAD Models

Datum Features are reported on datum_* entities.  Datum_system will show the complete Datum
Reference Frame.  Datum Targets are reported on placed_datum_target_feature.

Dimensional Tolerances are reported on the dimensional_characteristic_representation worksheet.
The dimension name, representation name, length/angle, length/angle name, and plus minus bounds are
reported.  The relevant section in the Recommended Practice is shown in the column headings.
Dimensional tolerances for holes are reported on *_hole_definition worksheets.

Geometric Tolerances are reported on *_tolerance entities by showing the complete Feature Control
Frame (FCF), and possible Dimensional Tolerance and Datum Feature.  The FCF should contain the
geometry tool, tolerance zone, datum reference frame, and associated modifiers.

If a Dimensional Tolerance refers to the same geometric element as a Geometric Tolerance, then it
will be shown above the FCF.  If a Datum Feature refers to the same geometric face as a Geometric
Tolerance, then it will be shown below the FCF.  If an expected Dimensional Tolerance is not shown
above a Geometric Tolerance, then the tolerances do not reference the same geometric element.  For
example, referencing the edge of a hole versus the surfaces of a hole.

The association of the Datum Feature with a Geometric Tolerance is based on each referring to the
same geometric element.  However, the PMI Presentation might show the Geometric Tolerance and
Datum Feature as two separate annotations with leader lines attached to the same geometric element.

The number of decimal places for dimension and geometric tolerance values can be specified in the
STEP file.  By definition the value is always truncated, however, the values can be rounded instead.
For example with the value 0.5625, the qualifier 'NR2 1.3' will truncate it to 0.562  Rounding will
show 0.563  Rounding values might result in a better match to graphical PMI shown by the Viewer or
to expected PMI in the NIST models.  See Options tab > Analyze > Round ...

Some syntax errors that indicate non-conformance to a CAx-IF Recommended Practices related to PMI
Representation are also reported in the Status tab and the relevant worksheet cells.  Syntax errors
are highlighted in red.  See Help > Analyze > Syntax Errors

A PMI Representation Coverage Analysis worksheet is also generated.

See Help > User Guide (section 6.1)
See Help > Analyze > PMI Coverage Analysis
See Examples > Spreadsheet - PMI Representation
See Examples > Sample STEP Files
See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242)"
    .tnb select .tnb.status
  }

  $helpAnalyze add command -label "PMI Presentation (Graphical PMI)" -command {
outputMsg "\nPMI Presentation ----------------------------------------------------------------------------------" blue
outputMsg "PMI Presentation (Graphical PMI) consists of geometric elements such as lines and arcs preserving
the exact appearance (color, shape, positioning) of the geometric and dimensional tolerance (GD&T)
annotations.  PMI Presentation is not intended to be computer-interpretable and does not have any
representation information, although it can be linked to its corresponding PMI Representation.

The analysis of Graphical PMI on annotation_curve_occurrence, annotation_curve,
annotation_fill_area_occurrence, and tessellated_annotation_occurrence entities is supported.
Geometric entities used for PMI Presentation annotations are reported in columns, highlighted in
yellow and green, on those worksheets.  Presentation Style, Saved Views, Validation Properties,
Annotation Plane, Associated Geometry, and Associated Representation are also reported.

The Summary worksheet indicates on which worksheets PMI Presentation is reported.  Some syntax
errors related to PMI Presentation are also reported in the Status tab and the relevant worksheet
cells.  Syntax errors are highlighted in red.  See Help > Analyze > Syntax Errors

An optional PMI Presentation Coverage Analysis worksheet can be generated.

See Help > User Guide (section 6.2)
See Help > View > Graphical PMI
See Help > Analyze > PMI Coverage Analysis
See Examples > Part with PMI
See Examples > PMI Presentation, Validation Properties
See Examples > Sample STEP Files
See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242,
  PMI Polyline Presentation for AP203 and AP214)"
    .tnb select .tnb.status
  }

# coverage analysis help
  $helpAnalyze add command -label "PMI Coverage Analysis" -command {
outputMsg "\nPMI Coverage Analysis -----------------------------------------------------------------------------" blue
outputMsg "PMI Coverage Analysis worksheets are generated when processing single or multiple files and when
reports for PMI Representation or Presentation are selected.

PMI Representation Coverage Analysis (semantic PMI) counts the number of PMI Elements found in a
STEP file for tolerances, dimensions, datums, modifiers, and CAx-IF Recommended Practices for PMI
Representation.  On the Coverage Analysis worksheet, some PMI Elements show their associated
symbol, while others show the relevant section in the Recommended Practice.  PMI Elements without
a section number do not have a Recommended Practice for their implementation.  The PMI Elements are
grouped by features related tolerances, tolerance zones, dimensions, dimension modifiers, datums,
datum targets, and other modifiers.  The number of some modifiers, e.g., maximum material
condition, does not differentiate whether they appear in the tolerance zone definition or datum
reference frame.  Rows with no count of a PMI Element can be shown, see Spreadsheet tab.

Some PMI Elements might not be exported to a STEP file by your CAD system.  Some PMI Elements are
only in AP242 edition 2.

If STEP files from the NIST CAD models (Websites > MBE PMI Validation Testing) are processed, then
the PMI Representation Coverage Analysis worksheet is color-coded by the expected number of PMI
elements in each CAD model.  See Help > Analyze > NIST CAD Models

The optional PMI Presentation Coverage Analysis (graphical PMI) counts the occurrences of the
recommended name attribute defined in the CAx-IF Recommended Practice for PMI Representation and
Presentation of PMI (AP242) or PMI Polyline Presentation (AP203/AP242).  The name attribute is
associated with the graphic elements used to draw a PMI annotation.  There is no semantic PMI
meaning to the name attributes.

See Help > Analyze > PMI Representation
See Help > User Guide (sections 6.1.7 and 6.2.1)
See Examples > PMI Coverage Analysis"
    .tnb select .tnb.status
  }

  $helpAnalyze add command -label "Syntax Errors" -command {
outputMsg "\nSyntax Errors -------------------------------------------------------------------------------------" blue
outputMsg "Syntax Errors are generated when an Analyze option related to Semantic PMI, Graphical PMI, and
Validation Properties is selected.  The errors refer to specific sections, figures, or tables in
the relevant CAx-IF Recommended Practice.  Errors should be fixed so that the STEP file can
interoperate with other CAx software.  See Websites > CAx Recommended Practices

Syntax errors are highlighted in red in the Status tab.  Other informative warnings are highlighted
in yellow.  Syntax errors that refer to section, figure, and table numbers might use numbers that
are in a newer version of a recommended practice that has not been publicly released.

Some syntax errors use abbreviations for STEP entities:
 GISU - geometric_item_specific_usage
 IIRU - identified_item_representation_usage

On the Summary worksheet in column A, most entity types that have syntax errors are colored gray.
A comment indicating that there are errors is also shown with a small red triangle in the upper
right corner of a cell in column A.

On an entity worksheet, most syntax errors are highlighted in red and have a cell comment with the
text of the syntax error that was shown in the Status tab.  Syntax errors are highlighted by *** in
the log file.

NOTE - Syntax Errors related to CAx-IF Recommended Practices are unrelated to errors detected with
the Syntax Checker.  See Help > Syntax Checker

See Help > User Guide (section 6.4)"
    .tnb select .tnb.status
  }

# NIST CAD model help
  $helpAnalyze add command -label "NIST CAD Models" -command {
outputMsg "\nNIST CAD Models -----------------------------------------------------------------------------------" blue
outputMsg "If a STEP file from a NIST CAD model is processed, then the PMI found in the STEP file is
automatically checked against the expected PMI in the corresponding NIST test case.  The PMI
Representation Coverage and Summary worksheets are color-coded by the expected PMI in each NIST
test case.  The color-coding only works if the STEP file name can be recognized as having been
generated from one of the NIST CAD models.

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
example, PMI annotations for hole features such as counterbore, countersink, and depth are not
supported.

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
- Yellow, orange, and yellow-green means that less were found than expected. (2/3)
- Red means that no instances of an expected PMI element were found. (0/3)
- Cyan means that more were found than expected. (4/3)
- Magenta means that some PMI elements were found when none were expected. (3/0)

From the PMI Representation Summary results, color-coded percentages of exact, partial, and
possible matches and missing PMI is shown in a table below the PMI Representation Coverage Analysis.
The Total PMI on which the percentages are based on is also shown.  Coverage Analysis is only based
on individual PMI elements.  The PMI Representation Summary is based on the entire PMI feature
control frame and provides a better understanding of the PMI.  The Coverage Analysis might show
that there is an exact match for all of the PMI elements, however, the Representation
Summary might show less than exact matches.

* Missing PMI *
Missing PMI annotations on the Summary worksheet or PMI elements on the Coverage worksheet might
mean that the CAD system or translator:
- did not or cannot correctly create in the CAD model a PMI annotation defined in a NIST test case
- did not follow CAx-IF Recommended Practices for PMI (See Websites > CAx Recommended Practices)
- has not implemented exporting a PMI element to a STEP file
- mapped an internal PMI element to the wrong STEP PMI element

NOTE - Some of the NIST test cases have complex PMI annotations that are not commonly used.  There
might be ambiguities in counting the number of PMI elements.

See Help > User Guide (section 6.5)
See Websites > MBE PMI Validation Testing
See Examples > Spreadsheet - PMI Representation"
    .tnb select .tnb.status
  }
  $Help add separator

# open Function Keys help
  $Help add command -label "Function Keys" -command {
outputMsg "\nFunction Keys -------------------------------------------------------------------------------------" blue
outputMsg "Function keys can be used as shortcuts for several commands:

F1 - Generate Spreadsheets and/or Views from the current or last STEP file
F2 - Open current or last Spreadsheet
F3 - Open current or last View file
F4 - Open Log file
F5 - Open STEP file in a text editor  (See Help > Open STEP File in App)
Shift-F5 - Open STEP file directory

F6 - Generate Speadsheets and/or Views from current or last set of multiple STEP files
F7 - Open current or last File Summary Spreadsheet generated from a set of multiple STEP files

F8 - Run the Syntax Checker (See Help > Syntax Checker)

F9  - Decrease this font size
F10 - Increase this font size

For F1, F2, F3, F6, and F7 the last STEP file, Spreadsheet, and View are remembered between
sessions.  In other words, F1 can process the last STEP file from a previous session without having
to select the file.  F2 and F3 function similarly for Spreadsheets and Views."
    .tnb select .tnb.status
  }

  $Help add command -label "Syntax Checker" -command {
outputMsg "\nSyntax Checker ------------------------------------------------------------------------------------" blue
outputMsg "The Syntax Checker checks for basic syntax errors and warnings in the STEP file related to missing
or extra attributes, incompatible and unresolved entity references, select value types, illegal and
unexpected characters, and other problems with entity attributes.  Some errors might prevent this
software and others from processing a STEP file.  Characters that are identified as illegal or
unexpected might not be shown in a spreadsheet or in the viewer.  See Help > Text Strings

There should not be any of these types of syntax errors in a STEP file.  Errors should be fixed to
ensure that the STEP file conforms to the STEP schema and can interoperate with other software.

There are other validation rules defined by STEP schemas (where, uniqueness, and global rules,
inverses, derived attributes, and aggregates) that are NOT checked.  Conforming to the validation
rules is also important for interoperability with STEP files.
See Websites > STEP Format and Schemas

The Syntax Checker can be run with function key F8 or when a Spreadsheet or View is generated.  The
Status tab might be grayed out when the Syntax Checker is running.

Syntax checker results appear in the Status tab.  If the Log File option is selected, the results
are also written to a log file (myfile-sfa-err.log).  The syntax checker errors and warnings are
not reported in the spreadsheet.  If errors and warnings are reported, the number in parentheses is
the line number in the STEP file where the error or warning was detected.

The Syntax Checker works with any supported STEP schema.  See Help > Supported STEP APs

NOTE - Syntax Checker errors and warnings are unrelated to those detected when CAx-IF Recommended
Practices are checked with one of the Analyze options.  See Help > Analyze > Syntax Errors"
    .tnb select .tnb.status
  }

# open STEP files help
  $Help add command -label "Open STEP File in App" -command {
outputMsg "\nOpen STEP File in App -----------------------------------------------------------------------------" blue
outputMsg "STEP files can be opened in other apps.  If apps are installed in their default directory, then the
pull-down menu in the Options tab will contain apps that can open a STEP file such as STEP viewers,
browsers, and conformance checkers.

The 'Tree View (for debugging)' option rearranges and indents the entities to show the hierarchy of
information in a STEP file.  The 'tree view' file (myfile-sfa.txt) is written to the same directory
as the STEP file or to the same user-defined directory specified in the Spreadsheet tab.  It is
useful for debugging STEP files but is not recommended for large STEP files.

The 'Default STEP Viewer' option opens the STEP file in whatever app is associated with STEP files.
A text editor always appear in the menu.  Use F5 to open the STEP file in the text editor.

See Help > User Guide (section 3.4.5)"
    .tnb select .tnb.status
  }

# multiple files help
  $Help add command -label "Multiple STEP Files" -command {
outputMsg "\nMultiple STEP Files -------------------------------------------------------------------------------" blue
outputMsg "Multiple STEP files can be selected in the Open File(s) dialog by holding down the control or shift
key when selecting files or an entire directory of STEP files can be selected with 'Open Multiple
STEP Files in a Directory'.  Files in subdirectories of the selected directory can also be
processed.

When processing multiple STEP files, a File Summary spreadsheet is generated in addition to
individual spreadsheets for each file.  The File Summary spreadsheet shows the entity count and
totals for all STEP files. The File Summary spreadsheet also links to the individual spreadsheets
and STEP files.

If only the File Summary spreadsheet is needed, it can be generated faster by deselecting most
Process categories and options in the Options tab.

If the reports for PMI Representation or Presentation are selected, then Coverage Analysis
worksheets are also generated.

In some rare cases an error will be reported with an entity when processing multiple files that is
not an error when processing it as a single file.  Reporting the error is a bug.

See Help > User Guide (section 8)
See Examples > PMI Coverage Analysis"
    .tnb select .tnb.status
  }

  $Help add command -label "Supported STEP APs" -command {
    outputMsg "\nSupported STEP APs ----------------------------------------------------------------------------" blue
    outputMsg "The following STEP Application Protocols (AP) and other schemas are supported for generating
spreadsheets.  The Viewer does not support all schemas.

The name of the AP is found on the FILE_SCHEMA entity in the HEADER section of a STEP file.
The 'e1' notation after an AP number below refers to an older version of that AP.  There are
multiple editions of AP242 all with the same name.  The edition is identified when a file is read.\n"

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

    if {[llength $schemas] <= 1} {
      errorMsg "No Supported STEP APs were found."
      if {[llength $schemas] == 1} {errorMsg "- Manually uninstall the existing IFCsvrR300 ActiveX Component 'App'."}
      errorMsg "- Restart the software to install the new IFCsvr toolkit."
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
    outputMsg "\nSee the Websites menu for information about the STEP Format, EXPRESS Schemas, AP242, and more."
    .tnb select .tnb.status
  }

  $Help add command -label "Text Strings" -command {
outputMsg "\nText Strings --------------------------------------------------------------------------------------" blue
outputMsg "The following supplements the User Guide section 5.5 on Unicode Characters.

Text strings in STEP files might use non-English characters or symbols.  Some examples are accented
characters in European languages (for example é), and Asian languages that use different characters
sets such as Russian or Chinese.  Text strings with non-English characters or symbols are usually
found on descriptive measure or product related entities with name, description, or id attributes.

According to ISO 10303 Part 21 section 6.4.3, Unicode can be used for non-English characters and
symbols with the control directives \\X\\ and \\X2\\.  For example, \\X\\E9 or \\X2\\00E9\\X0\\ is used for
the accented character é.  Definitions of Unicode characters, such as E9, can be found at
www.unicode.org/charts  Some CAD software do not support these control directives when exporting or
importing a STEP file.

For a spreadsheet, the \\X\\ and \\S\\ control directives are supported by default.  To support
non-English characters using the \\X2\\ control directive, use the option on the Spreadsheets tab.
Processing a STEP file takes longer with the option selected.  In some cases the option will be
automatically selected based on the file schema or size.  There is a warning message if \\X2\\ is
detected in the STEP file and the option is not selected.  In this case the \\X2\\ characters are
ignored and will be missing in the spreadsheet.

For the viewer, all control directives are supported for part and assembly names.  Non-English
characters are supported depending STEP file encoding, e.g., UTF-8 or ANSI.

Some non-English characters might cause the software to crash or prevent a view from being
generated.  See Help > Crash Recovery

The Syntax Checker identifies non-English characters as 'illegal characters'.  You should test your
CAD software to see if it supports non-English characters or control directives.

See Websites > STEP Format and Schemas > ISO 10303 Part 21 Standard"
    .tnb select .tnb.status
  }

# large files help
  $Help add command -label "Large STEP Files" -command {
outputMsg "\nLarge STEP Files ----------------------------------------------------------------------------------" blue
outputMsg "The Status tab might be grayed out when a large STEP file is being read.

See Help > Viewer for viewing large STEP files.

To reduce the amount of time to process large STEP files and to reduce the size of the resulting
spreadsheet, several options are available:
- In the Process section, deselect entity types Geometry and Coordinates
- In the Process section, select only a User-Defined List of required entities
- In the Spreadsheet tab, select a smaller value for the Maximum Rows
- In the Options tab, deselect Analyze options and Inverse Relationships

The STEP File Analyzer and Viewer might also crash when processing very large STEP files.  Popup
dialogs might appear that say 'Unable to alloc xxx bytes'.  See the Help > Crash Recovery."
    .tnb select .tnb.status
  }

  $Help add command -label "Crash Recovery" -command {
outputMsg "\nCrash Recovery ------------------------------------------------------------------------------------" blue
outputMsg "Sometimes the STEP File Analyzer and Viewer crashes after a STEP file has been successfully opened
and the processing of entities has started.  Popup dialogs might appear that say
\"ActiveState Basekit has stopped working\" or \"Runtime Error!\".

A crash is most likely due to syntax errors in the STEP file, a very large STEP file, or due to
limitations of the toolkit used to read STEP files.  Run the Syntax Checker with function key F8 or
the option on the Options tab to check for errors with entities that might have caused the crash.
See Help > Syntax Checker

The software keeps track of the last entity type processed when it crashed.  The list of bad entity
types is stored in myfile-skip.dat.  Simply restart the STEP File Analyzer and Viewer and use F1 to
process the last STEP file or use F6 if processing multiple files.  The entity types listed in
myfile-skip.dat that caused the crash will be skipped.

NOTE - When the STEP file is processed again, the list of specific entities that are not processed
is reported.  If syntax errors related to the bad entities are corrected, then delete or edit the
*-skip.dat file so that the corrected entities are processed.

See Help > User Guide (sections 2.4 and 9)
See Help > Large STEP Files

Other fixes:

1 - STEP files generated by Creo or Pro/Engineer might cause the software to crash.  This is due to
a problem with some of the PRESENTATION_STYLE_ASSIGNMENT entities.  To fix the problem, change all
PRESENTATION_STYLE_ASSIGNMENT((.NULL.)); to PRESENTATION_STYLE_ASSIGNMENT((NULL_STYLE(.NULL.))); in
the STEP file.  Then follow the instructions in the NOTE above.  See Websites > CAx-IF Recommended
Practices for the Recommended Practice for $recPracNames(pmi242), sec. 9.1.

2 - Processing of the type of entity that caused the error can be deselected in the Options tab
under Process.  However, this will prevent processing of other entities that do not cause a crash."
    .tnb select .tnb.status
  }

  $Help add separator
  $Help add command -label "Disclaimers"     -command {showDisclaimer}
  $Help add command -label "NIST Disclaimer" -command {openURL https://www.nist.gov/disclaimer}
  $Help add command -label "About" -command {
    set sysvar "System:   $tcl_platform(os) $tcl_platform(osVersion)"
    if {$excelVersion < 1000} {append sysvar ", Excel $excelVersion"}
    catch {append sysvar ", IFCsvr [registry get $ifcsvrVer {DisplayVersion}]"}
    append sysvar ", Files processed: $filesProcessed"
    if {$opt(xlMaxRows) != 100003} {append sysvar "\n          For more System variables, set Maximum Rows to 100000 and repeat Help > About"}

    outputMsg "\nSTEP File Analyzer and Viewer ---------------------------------------------------------------------" blue
    outputMsg "Version:  [getVersion]"
    outputMsg "Updated:  [string trim [clock format $progtime -format "%e %b %Y"]]"
    outputMsg "Contact:  [lindex $contact 0], [lindex $contact 1]\n$sysvar

The STEP File Analyzer and Viewer was first released in April 2012 and is developed at
NIST in the Systems Integration Division of the Engineering Laboratory.  Click the logo
below for information about NIST.

See Help > Disclaimers and NIST Disclaimer

Credits
- Generating spreadsheets:        Microsoft Excel  https://products.office.com/excel
- Reading and parsing STEP files: IFCsvr ActiveX Component, Copyright \u00A9 1999, 2005 SECOM Co., Ltd. All Rights Reserved
                                  IFCsvr has been modified by NIST to include STEP schemas.
                                  The license agreement can be found in  C:\\Program Files (x86)\\IFCsvrR300\\doc
- Translating STEP to X3D:        Developed by Soonjo Kwon (former NIST Guest Researcher)
                                  See Websites > STEP Software"

# debug
    if {$opt(xlMaxRows) == 100003} {
      outputMsg " "
      outputMsg "SFA variables" red
      catch {outputMsg " Drive $drive"}
      catch {outputMsg " Home  $myhome"}
      catch {outputMsg " Docs  $mydocs"}
      catch {outputMsg " Desk  $mydesk"}
      catch {outputMsg " Menu  $mymenu"}
      catch {outputMsg " Temp  $mytemp  ([file exists $mytemp])"}
      outputMsg " pf32  $pf32"
      if {$pf64 != ""} {outputMsg " pf64  $pf64"}
      catch {outputMsg " webCmd  $webCmd"}
      catch {outputMsg " scriptName $scriptName"}
      outputMsg " Tcl [info patchlevel], twapi [package versions twapi]"
      outputMsg " S [winfo screenwidth  .]x[winfo screenheight  .], M [winfo reqwidth .]x[expr {int([winfo reqheight .]*1.05)}]"

      outputMsg "Registry values" red
      catch {outputMsg " Personal  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Personal}]"}
      catch {outputMsg " Desktop   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]"}
      catch {outputMsg " Programs  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Programs}]"}
      catch {outputMsg " AppData   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local AppData}]"}

      outputMsg "Environment variables" red
      foreach id [lsort [array names env]] {
        foreach id1 [list HOME Program System USER TEMP TMP APP ROSE EDM] {if {[string first $id1 $id] == 0} {outputMsg " $id  $env($id)"; break}}
      }
    }
    .tnb select .tnb.status
  }

# examples menu
  $Examples add command -label "Sample STEP Files (zip)" -command {openURL https://s3.amazonaws.com/nist-el/mfg_digitalthread/NIST_CTC_STEP_PMI.zip}
  $Examples add command -label "STEP File Library"       -command {openURL https://www.cax-if.org/cax/cax_stepLib.php}

  $Examples add cascade -label "Archived Sample STEP Files" -menu $Examples.0
  set Examples0 [menu $Examples.0 -tearoff 1]
  $Examples0 add command -label "AP203e2 Files" -command {openURL http://web.archive.org/web/20160812122922/http://www.steptools.com/support/stdev_docs/stpfiles/ap203e2/index.html}
  $Examples0 add command -label "AP203e1 Files" -command {openURL http://web.archive.org/web/20160812122922/http://www.steptools.com/support/stdev_docs/stpfiles/ap203/index.html}
  $Examples0 add command -label "AP214e1 Files" -command {openURL http://web.archive.org/web/20160903141712/http://www.steptools.com/support/stdev_docs/stpfiles/ap214/index.html}

  $Examples add separator
  $Examples add command -label "View Box Assembly"               -command {openURL https://pages.nist.gov/CAD-PMI-Testing/step-file-viewer.html}
  $Examples add command -label "Part with PMI"                   -command {openURL https://pages.nist.gov/CAD-PMI-Testing/graphical-pmi-viewer.html}
  $Examples add command -label "AP242 Tessellated Part with PMI" -command {openURL https://pages.nist.gov/CAD-PMI-Testing/tessellated-part-geometry.html}
  $Examples add command -label "AP209 Finite Element Model"      -command {openURL https://pages.nist.gov/CAD-PMI-Testing/ap209-viewer.html}
  $Examples add separator
  $Examples add command -label "Spreadsheet - PMI Representation"        -command {openURL https://s3.amazonaws.com/nist-el/mfg_digitalthread/STEP-File-Analyzer-PMI-Representation-sfa.xlsx}
  $Examples add command -label "PMI Presentation, Validation Properties" -command {openURL https://s3.amazonaws.com/nist-el/mfg_digitalthread/STEP-File-Analyzer-sfa.xlsx}
  $Examples add command -label "PMI Coverage Analysis"                   -command {openURL https://s3.amazonaws.com/nist-el/mfg_digitalthread/STEP-File-Analyzer-Coverage.xlsx}
}

#-------------------------------------------------------------------------------
# Websites menu
proc guiWebsitesMenu {} {
  global Websites

  $Websites add command -label "STEP File Analyzer and Viewer"              -command {openURL https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer}
  $Websites add command -label "Conformance Checking of PMI in STEP Files"  -command {openURL https://www.nist.gov/publications/conformance-checking-pmi-representation-cad-model-step-data-exchange-files}
  $Websites add command -label "MBE PMI Validation and Comformance Testing" -command {openURL https://www.nist.gov/el/systems-integration-division-73400/mbe-pmi-validation-and-conformance-testing-project/download}
  $Websites add command -label "SMS Test Bed: Technical Data Packages"      -command {openURL https://smstestbed.nist.gov/tdp/}

  $Websites add separator
  $Websites add command -label "CAx Interoperability Forum (CAx-IF)" -command {openURL https://www.cax-if.org/cax/cax_introduction.php}
  $Websites add command -label "CAx Recommended Practices"           -command {openURL https://www.cax-if.org/cax/cax_recommPractice.php}
  $Websites add command -label "CAD Implementations"                 -command {openURL https://www.cax-if.org/cax/vendor_info.php}
  $Websites add command -label "STEP File Viewers"                   -command {openURL https://www.cax-if.org/step_viewers.php}

  $Websites add separator
  $Websites add cascade -label "AP242" -menu $Websites.0
  set Websites0 [menu $Websites.0 -tearoff 1]
  $Websites0 add command -label "AP242 Project"           -command {openURL http://www.ap242.org}
  $Websites0 add command -label "AP203 vs AP214 vs AP242" -command {openURL https://www.capvidia.com/blog/best-step-file-to-use-ap203-vs-ap214-vs-ap242}
  $Websites0 add command -label "Schema Documentation"    -command {openURL https://www.cax-if.org/documents/AP242ed2_HTML/AP242ed2.htm}
  $Websites0 add command -label "EXPRESS Schema"          -command {openURL https://www.cax-if.org/documents/ap242ed2_mim_lf_v1.101.exp}
  $Websites0 add command -label "ISO 10303-242"           -command {openURL https://www.iso.org/standard/66654.html}
  $Websites0 add separator
  $Websites0 add command -label "Journal Article"    -command {openURL https://www.nist.gov/publications/portrait-iso-step-tolerancing-standard-enabler-smart-manufacturing-systems}
  $Websites0 add command -label "Presentation (pdf)" -command {openURL https://s3.amazonaws.com/nist-el/mfg_digitalthread/16_aBarnardFeeney.pdf}
  $Websites0 add command -label "Benchmark Testing"  -command {openURL http://www.asd-ssg.org/step-ap242-benchmark}

  $Websites add cascade -label "STEP Format and Schemas" -menu $Websites.2
  set Websites2 [menu $Websites.2 -tearoff 1]
  $Websites2 add command -label "STEP Format"                 -command {openURL https://www.loc.gov/preservation/digital/formats/fdd/fdd000448.shtml}
  $Websites2 add command -label "ISO 10303 Part 21"           -command {openURL https://en.wikipedia.org/wiki/ISO_10303-21}
  $Websites2 add command -label "ISO 10303 Part 21 Edition 3" -command {openURL https://www.steptools.com/stds/step/}
  $Websites2 add command -label "ISO 10303 Part 21 Standard"  -command {openURL http://www.steptools.com/stds/step/IS_final_p21e3.html}
  $Websites2 add separator
  $Websites2 add command -label "EXPRESS Schemas"                -command {openURL https://www.cax-if.org/cax/cax_express.php}
  $Websites2 add command -label "Archived EXPRESS Schemas"       -command {openURL http://web.archive.org/web/20160322005246/www.steptools.com/support/stdev_docs/express/}
  $Websites2 add command -label "STEP AIM Model"                 -command {openURL https://www.steptools.com/stds/stp_expg/aim.html}
  $Websites2 add command -label "ISO 10303 Part 11 EXPRESS"      -command {openURL https://www.loc.gov/preservation/digital/formats/fdd/fdd000449.shtml}
  $Websites2 add command -label "EXPRESS data modeling language" -command {openURL https://en.wikipedia.org/wiki/EXPRESS_(data_modeling_language)}
  $Websites2 add separator
  $Websites2 add command -label "AP238 Machining"  -command {openURL http://www.ap238.org}
  $Websites2 add command -label "AP239 PLCS"       -command {openURL http://www.ap239.org}
  $Websites2 add command -label "PDM-IF"           -command {openURL http://www.pdm-if.org/}
  $Websites2 add command -label "AP209 FEA"        -command {openURL http://www.ap209.org}
  $Websites2 add command -label "CAE-IF"           -command {openURL https://www.cax-if.org/cae/cae_introduction.php}
  $Websites2 add command -label "AP243 MoSSEC"     -command {openURL http://www.mossec.org/}
  $Websites2 add command -label "AP235 Properties" -command {openURL http://www.ap235.org}
  $Websites2 add separator
  $Websites2 add command -label "AP203 Recommended Practice (1998)" -command {openURL https://www.oasis-open.org/committees/download.php/11728/recprac8.pdf}
  $Websites2 add command -label "STEP Application Handbook (2006)"  -command {openURL http://pdesinc.org/wp-content/uploads/2020/04/STEP-application-handbook-63006-BF.pdf}

  $Websites add cascade -label "STEP Software" -menu $Websites.3
  set Websites3 [menu $Websites.3 -tearoff 1]
  $Websites3 add command -label "Source Code"            -command {openURL https://github.com/usnistgov/SFA}
  $Websites3 add command -label "STEP to X3D Translator" -command {openURL https://www.nist.gov/services-resources/software/step-x3d-translator}
  $Websites3 add command -label "STEP to OWL Translator" -command {openURL https://github.com/usnistgov/stp2owl}
  $Websites3 add command -label "STEP Class Library"     -command {openURL https://www.nist.gov/services-resources/software/step-class-library-scl}
  $Websites3 add command -label "Express Engine"         -command {openURL https://sourceforge.net/projects/exp-engine/}

  $Websites add cascade -label "STEP Related Organizations" -menu $Websites.4
  set Websites4 [menu $Websites.4 -tearoff 1]
  $Websites4 add command -label "PDES, Inc. (US)"        -command {openURL https://pdesinc.org/}
  $Websites4 add command -label "prostep ivip (Germany)" -command {openURL https://www.prostep.org/en/projects/}
  $Websites4 add command -label "AFNeT (France)"         -command {openURL http://afnet.fr/dotank/sps/plm-committee/}
  $Websites4 add command -label "KStep (Korea)"          -command {openURL https://www.kstep.or.kr/}
  $Websites4 add separator
  $Websites4 add command -label "LOTAR (LOng Term Archiving and Retrieval)" -command {openURL https://lotar-international.org/}
  $Websites4 add command -label "ASD Strategic Standardisation Group"       -command {openURL http://www.asd-ssg.org/}
  $Websites4 add command -label "ISO TC184/SC4"                             -command {openURL https://committee.iso.org/home/tc184sc4}
}

#-------------------------------------------------------------------------------
proc showDisclaimer {} {
  global sfaVersion

# text disclaimer
  if {$sfaVersion > 0} {
    outputMsg "\nDisclaimer ----------------------------------------------------------------------------------------" blue
    outputMsg "This software was developed at the National Institute of Standards and Technology by employees of
the Federal Government in the course of their official duties.  Pursuant to Title 17 Section 105 of
the United States Code this software is not subject to copyright protection and is in the public
domain.  This software is an experimental system.  NIST assumes no responsibility whatsoever for
its use by other parties, and makes no guarantees, expressed or implied, about its quality,
reliability, or any other characteristic.

This software is provided by NIST as a public service.  You may use, copy and distribute copies of
the software in any medium, provided that you keep intact this entire notice.  You may improve,
modify and create derivative works of the software or any portion of the software, and you may copy
and distribute such modifications or works.  Modified works should carry a notice stating that you
changed the software and should note the date and nature of any such change.  Please explicitly
acknowledge NIST as the source of the software.

The Examples menu of this software provides links to several sources of STEP files.  This software
and other software might indicate that there are errors in some of the STEP files.  NIST assumes
no responsibility whatsoever for the use of the STEP files by other parties, and makes no
guarantees, expressed or implied, about their quality, reliability, or any other characteristic.

Any mention of commercial products or references to web pages in this software is for information
purposes only; it does not imply recommendation or endorsement by NIST.  For any of the web links
in this software, NIST does not necessarily endorse the views expressed, or concur with the facts
presented on those web sites.

This software uses Microsoft Excel, IFCsvr, and software based on Open CASCADE that are covered by
their own Software License Agreements.

See Help > NIST Disclaimer and Help > About"
    .tnb select .tnb.status

# dialog box disclaimer
  } else {
set txt "This software was developed at the National Institute of Standards and Technology by employees of the Federal Government in the course of their official duties. Pursuant to Title 17 Section 105 of the United States Code this software is not subject to copyright protection and is in the public domain.  This software is an experimental system.  NIST assumes no responsibility whatsoever for its use by other parties, and makes no guarantees, expressed or implied, about its quality, reliability, or any other characteristic.

This software is provided by NIST as a public service.  You may use, copy and distribute copies of the software in any medium, provided that you keep intact this entire notice.  You may improve, modify and create derivative works of the software or any portion of the software, and you may copy and distribute such modifications or works.  Modified works should carry a notice stating that you changed the software and should note the date and nature of any such change.  Please explicitly acknowledge NIST as the source of the software.

The Examples menu of this software provides links to several sources of STEP files.  This software and other software might indicate that there are errors in some of the STEP files.  NIST assumes no responsibility whatsoever for the use of the STEP files by other parties, and makes no guarantees, expressed or implied, about their quality, reliability, or any other characteristic.

Any mention of commercial products or references to web pages in this software is for information purposes only; it does not imply recommendation or endorsement by NIST.  For any of the web links in this software, NIST does not necessarily endorse the views expressed, or concur with the facts presented on those web sites.

This software uses Microsoft Excel, IFCsvr and software based on Open CASCADE that are covered by their own Software License Agreements.

See Help > NIST Disclaimer and Help > About"

    tk_messageBox -type ok -icon info -title "Disclaimers" -message $txt
  }
}

#-------------------------------------------------------------------------------
# crash recovery dialog
proc showCrashRecovery {} {

set txt "Sometimes the STEP File Analyzer and Viewer crashes AFTER a file has been successfully opened and the processing of entities has started.

A crash is most likely due to syntax errors in the STEP file or sometimes due to limitations of the toolkit used to read STEP files.  Run the Syntax Checker with the Output Format option on the Options tab or function key F8 to check for errors with entities that might have caused the crash.  See Help > Syntax Checker

You can also restart the software and process the same STEP file again by using function key F1 or if processing multiple STEP files use F6.  The software keeps track of which entity type caused the error for a particular STEP file and won't process that type again.  The bad entities types are stored in a file *-skip.dat  If syntax errors related to the bad entities are corrected, then delete the *-skip.dat file so that the corrected entities are processed.

The software might also crash when processing very large STEP files.  See Help > Large STEP Files

More details about recovering from a crash are explained in Help > Crash Recovery and in the User Guide."

  tk_messageBox -type ok -icon error -title "What to do if the STEP File Analyzer and Viewer crashes?" -message $txt
}

#-------------------------------------------------------------------------------
proc guiToolTip {ttmsg tt} {
  global ap242only entCategory

  set ttlim 120
  if {$tt == "stepPRES" || $tt == "stepGEOM"} {set ttlim 150}

  foreach type {ap203 ap242} {
    set ttlen 0
    foreach item [lsort $entCategory($tt)] {
      set ok 0
      set ent $item
      switch -- $type {
        ap203 {if {[lsearch $ap242only $ent] == -1} {set ok 1}}
        ap242 {if {[lsearch $ap242only $ent] != -1} {set ok 1}}
      }
      if {$ok} {
        append ttmsg "$ent   "
        incr ttlen [string length $ent]
        if {$ttlen > $ttlim} {
          if {[string index $ttmsg end] != "\n"} {append ttmsg "\n"}
          set ttlen 0
        }
      }
    }
    if {$type == "ap203"} {append ttmsg "\n\nThe following entities are found only in AP242.\n\n"}
  }
  return $ttmsg
}

#-------------------------------------------------------------------------------
proc getOpenPrograms {} {
  global env dispApps dispCmds dispCmd appNames appName
  global drive editorCmd developer myhome pf32 pf64 pflist

# Including any of the CAD viewers and software does not imply a recommendation or endorsement of them by NIST https://www.nist.gov/disclaimer
# For more STEP viewers, go to https://www.cax-if.org/step_viewers.php

  regsub {\\} $pf32 "/" p32
  lappend pflist $p32
  if {$pf64 != "" && $pf64 != $pf32} {
    regsub {\\} $pf64 "/" p64
    lappend pflist $p64
  }
  set lastver 0

# Jotne EDM Model Checker
  if {$developer} {
    foreach match [glob -nocomplain -directory [file join $drive edm] -join edm* bin Edms.exe] {set dispApps($match) "EDM Model Checker 5"}
    foreach match [glob -nocomplain -directory [file join $pf64 Jotne] -join EDMsdk6* bin edms.exe] {
      set ver [lindex [split $match "/"] 3]
      set dispApps($match) [string range $ver 0 [string last "." $ver]-1]
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
      [list {*}[glob -nocomplain -directory [file join $pf "Common Files"] -join "eDrawings*" eDrawings.exe] "eDrawings Viewer"] \
      [list {*}[glob -nocomplain -directory [file join $pf "SOLIDWORKS Corp"] -join "eDrawings (*)" eDrawings.exe] "eDrawings Pro"] \
      [list {*}[glob -nocomplain -directory [file join $pf "Stratasys Direct Manufacturing"] -join "SolidView Pro RP *" bin SldView.exe] SolidView] \
      [list {*}[glob -nocomplain -directory [file join $pf "TransMagic Inc"] -join "TransMagic *" System code bin TransMagic.exe] TransMagic] \
      [list {*}[glob -nocomplain -directory [file join $pf Actify SpinFire] -join "*" SpinFire.exe] SpinFire] \
      [list {*}[glob -nocomplain -directory [file join $pf CADSoftTools] -join "ABViewer*" ABViewer.exe] ABViewer] \
      [list {*}[glob -nocomplain -directory [file join $pf Kubotek] -join "KDisplayView*" KDisplayView.exe] "K-Display View"] \
      [list {*}[glob -nocomplain -directory [file join $pf Kubotek] -join "Spectrum*" Spectrum.exe] Spectrum] \
      [list {*}[glob -nocomplain -directory [file join $pf] -join "3D-Tool V*" 3D-Tool.exe] 3D-Tool] \
      [list {*}[glob -nocomplain -directory [file join $pf] -join "VariCADViewer *" bin varicad-x64.exe] "VariCAD Viewer"] \
      [list {*}[glob -nocomplain -directory [file join $pf] -join Clari3D "Lite*" lite.exe] "Clari3D Lite"] \
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
          set c1 [string first "eDrawings20" $match]
          if {$c1 != -1} {set dispApps($match) "[lindex $app 1] [string range $match $c1+9 $c1+12]"}
        }
      }
    }

    set applist [list \
      [list [file join $pf "3DJuump X64" 3DJuump.exe] "3DJuump"] \
      [list [file join $pf "C3D Labs" "C3D Viewer" bin c3dviewer.exe] "C3D Viewer"] \
      [list [file join $pf "CAD Assistant" CADAssistant.exe] "CAD Assistant"] \
      [list [file join $pf "CAD Exchanger" bin Exchanger.exe] "CAD Exchanger"] \
      [list [file join $pf "SOLIDWORKS Corp" eDrawings eDrawings.exe] "eDrawings Pro"] \
      [list [file join $pf "SOLIDWORKS Corp" "eDrawings X64 Edition" eDrawings.exe] "eDrawings Pro"] \
      [list [file join $pf "STEP Tools" "STEP-NC Machine" STEPNCExplorer.exe] "STEP-NC Machine"] \
      [list [file join $pf "STEP Tools" "STEP-NC Machine" STEPNCExplorer_x86.exe] "STEP-NC Machine"] \
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
    foreach app [list {*}[glob -nocomplain -directory [file join $myhome AppData Local] -join "FreeCAD *" bin FreeCAD.exe] FreeCAD] {
      set ver [lindex [split [file nativename $app] [file separator]] 5]
      if {$pf64 != "" && [string first "x86" $app] != -1} {append ver " (32-bit)"}
      set dispApps($app) $ver
    }
  }

# others
  set b1 [file join $myhome AppData Local IDA-STEP ida-step.exe]
  if {[file exists $b1]} {
    set dispApps($b1) "IDA-STEP Viewer"
  }
  set b1 [file join $drive CCELabs EnSuite-View Bin EnSuite-View.exe]
  if {[file exists $b1]} {
    set dispApps($b1) "EnSuite-View"
  } else {
    set b1 [file join $drive CCE EnSuite-View Bin EnSuite-View.exe]
    if {[file exists $b1]} {
      set dispApps($b1) "EnSuite-View"
    }
  }

#-------------------------------------------------------------------------------
# default viewer
  set dispApps(Default) "Default STEP Viewer"

# file tree view
  set dispApps(Indent) "Tree View (for debugging)"

#-------------------------------------------------------------------------------
# set text editor command and name for Notepad++ or Notepad
  set editorCmd ""
  foreach pf $pflist {
    set cmd [file join $pf Notepad++ notepad++.exe]
    if {[file exists $cmd]} {
      set editorCmd $cmd
      set dispApps($editorCmd) "Notepad++"
      break
    }
  }
  if {$editorCmd == ""} {
    if {[info exists env(windir)]} {
      set cmds [list [file join $env(windir) system32 notepad.exe] [file join $env(windir) notepad.exe]]
    } else {
      set cmds [list [file join C:/ Windows system32 notepad.exe] [file join C:/ Windows notepad.exe]]
    }
    foreach cmd $cmds {
      if {[file exists $cmd]} {
        set editorCmd $cmd
        set dispApps($editorCmd) "Notepad"
        break
      }
    }
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
