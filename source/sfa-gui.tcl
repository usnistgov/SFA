# SFA version
proc getVersion {} {return 5.25}

# see proc installIFCsvr in sfa-proc.tcl for the IFCsvr version
# see below (line 37) for the sfaVersion when IFCsvr was updated

# -------------------------------------------------------------------------------
proc whatsNew {} {
  global progtime sfaVersion

  if {$sfaVersion > 0 && $sfaVersion < [getVersion]} {
    outputMsg "\nThe previous version of the STEP File Analyzer and Viewer was: $sfaVersion" red
  }

# new user welcome message
  if {$sfaVersion == 0} {
    outputMsg "\nWelcome to the NIST STEP File Analyzer and Viewer [getVersion]\n" blue
    outputMsg "You will be prompted to install the IFCsvr toolkit which is required to read STEP files.
After the toolkit is installed, you are ready to process a STEP file.  Go to the File
menu, select a STEP file, and click the Generate button below.  If you only need to view
the STEP part geometry, go to the Generate tab and check View and Part Only.

Please take a few minutes to read some of the Help and the User Guide so that you
understand the options available and the resulting output.  Explore the Examples and
Websites menus.  Read the Disclaimers at the end of the Help menu.

Use F9 and F10 to change the font size here.  See Help > Function Keys"
  }

  outputMsg "\nWhat's New (Version: [getVersion]  Updated: [string trim [clock format $progtime -format "%e %b %Y"]])" blue

# messages if SFA has already been run
  if {$sfaVersion > 0} {

# update the version number when IFCsvr is repackaged to include updated STEP schemas
    if {$sfaVersion < 5.24} {outputMsg "- The IFCsvr toolkit might need to be reinstalled.  Please follow the directions carefully." red}

    if {$sfaVersion < 4.60} {
      outputMsg "- User Guide (Update 7) is based on version 4.60 (October 2021)"
      openUserGuide
    }
  }
  if {$sfaVersion < 5.10} {outputMsg "- Faster processing of B-rep and AP242 tessellated part geometry for the Viewer, see Release Notes" red}
  if {$sfaVersion < 5.0}  {outputMsg "- Renamed 'Options' and 'Spreadsheet' tabs to 'Generate' and 'More'"}
  if {$sfaVersion < 5.20} {outputMsg "- Renamed 'PMI Representation' to 'Semantic PMI' and 'PMI Presentation' to' Graphic PMI'"}
  if {$sfaVersion < 5.06} {outputMsg "- New 'Other' Entity Types category on the Generate tab"}
  if {$sfaVersion < 5.03} {outputMsg "- Hidden checkboxes and sliders in the Viewer"}
  if {$sfaVersion < 5.02} {outputMsg "- Help > Viewer > Viewpoints, and Help > Viewer > New Features"}
  if {$sfaVersion < 5.10} {outputMsg "- Updated Sample STEP Files, see Examples menu"}
  if {$sfaVersion < 5.16} {outputMsg "- The CAx-IF website and Recommended Practices for PMI in AP242 have been updated, see Websites"}
  outputMsg "- See Help > Release Notes for all new features and bug fixes"
  set sfaVersion [getVersion]

  .tnb select .tnb.status
  update idletasks
}

#-------------------------------------------------------------------------------
# open User Guide
proc openUserGuide {} {
  global sfaVersion

# update for new versions, local and online
  if {$sfaVersion > 4.6} {
    outputMsg "\nThe User Guide (Update 7) is based on version 4.60 (October 2021)" blue
    outputMsg "- See Help > Release Notes for software updates\n- New and updated features are documented in the Help menu\n- The Options and Spreadsheet tabs have been renamed Generate and More\n- PMI Representation and PMI Presentation have been renamed Semantic PMI and Graphic PMI"
    .tnb select .tnb.status
  }
  set fname [file nativename [file join [file dirname [info nameofexecutable]] "SFA-User-Guide-v7.pdf"]]
  set URL https://doi.org/10.6028/NIST.AMS.200-12
  if {![file exists $fname]} {set fname $URL}
  openURL $fname
}

#-------------------------------------------------------------------------------
# start window, bind keys
proc guiStartWindow {} {
  global fout editorCmd lastX3DOM lastXLS lastXLS1 localName localNameList

  wm title . "STEP File Analyzer and Viewer [getVersion]"
  wm protocol . WM_DELETE_WINDOW {exit}

# yellow background color
  set bgcolor  "#ffffbb"
  option add *Frame.background $bgcolor
  option add *Label.background $bgcolor

  ttk::style configure TCheckbutton -background $bgcolor
  ttk::style configure TRadiobutton -background $bgcolor
  ttk::style configure TLabelframe  -background $bgcolor

  font create fontBold {*}[font configure TkDefaultFont]
  font configure fontBold -weight bold
  ttk::style configure TLabelframe.Label -background $bgcolor -font fontBold

# key bindings
  bind . <Control-o> {openFile}
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
  global buttons ftrans mytemp nprogBarEnts nprogBarFiles opt wdir wingeo winpos

# generate button
  set ftrans [frame .ftrans1 -bd 2 -background "#F0F0F0"]
  set butstr "Spreadsheet"
  if {$opt(xlFormat) == "CSV"} {set butstr "CSV Files"}
  set buttons(generate) [ttk::button $ftrans.generate1 -text "Generate $butstr" -padding 4 -state disabled -command {
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
    tooltip::tooltip $l3 "Learn more about NIST"
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

# set the window position and dimensions
  catch {wm geometry . $winpos}
  catch {wm geometry . $wingeo}
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
  global File openFileList

  $File add command -label "Open File(s)..." -accelerator "Ctrl+O" -command {openFile}
  $File add command -label "Open Multiple Files in a Directory..." -accelerator "F6" -command {openMultiFile}
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
  $File add command -label "Exit" -accelerator "Ctrl+Q" -command exit
}

#-------------------------------------------------------------------------------
# generate tab (formerly options)
proc guiGenerateTab {} {
  global allNone buttons cb entCategory fopt fopta nb opt recPracNames useXL xlInstalled

  set cb 0
  set wopt [ttk::panedwindow $nb.generate -orient horizontal]
  $nb add $wopt -text " Generate " -padding 2
  set fopt [frame $wopt.fopt -bd 2 -relief sunken]

#-------------------------------------------------------------------------------
# generate section
  set foptOF [frame $fopt.of -bd 0]
  set foptk [ttk::labelframe $foptOF.k -text " Generate "]

# checkbuttons are used for pseudo-radiobuttons
  foreach item {{" Spreadsheet" gen(Excel)} {" CSV Files" gen(CSV)} {" View" gen(View)}} {
    set idx "gen[string range [lindex $item 1] 4 end-1]"
    set buttons($idx) [ttk::checkbutton $foptk.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {
      if {![info exists useXL]} {set useXL 1}
      if {[info exists xlInstalled]} {
        if {!$xlInstalled} {set useXL 0}
      } else {
        set xlInstalled 1
      }

# toggle spreadsheet and view
      if {$gen(Excel) == 0 && $gen(View) == 0 && $gen(Excel1) == 0 && $gen(View1) == 1} {
        set gen(Excel) 1
      } else {
        if {$gen(Excel) == 0} {
          set gen(View) 1
          set gen(View1) 1
          set gen(None) 1
        }
        if {$gen(View) == 0} {
          set gen(Excel) 1
          set gen(Excel1) 1
          set gen(None) 0
        }
      }

# part only
      if {$gen(None) && $opt(xlFormat) != "None"} {
        set gen(Excel) 0
        set gen(CSV) 0
        set opt(xlFormat) "None"
        set allNone -1
        if {$useXL && $xlInstalled} {$buttons(genExcel) configure -state normal}
      }

# spreadsheet
      if {$gen(Excel) && $opt(xlFormat) != "Excel"} {
        set gen(None) 0
        set opt(partOnly) 0
        if {$useXL} {
          set opt(xlFormat) "Excel"
        } else {
          set gen(Excel) 0
          set gen(CSV) 1
          set opt(xlFormat) "CSV"
        }
      }

# CSV
      if {$gen(CSV)} {
        set opt(partOnly) 0
        if {$useXL} {
          set gen(Excel) 1
          $buttons(genExcel) configure -state disabled
        }
        if {$opt(xlFormat) != "CSV"} {
          set gen(None) 0
          set opt(xlFormat) "CSV"
        }
      } elseif {$xlInstalled} {
        $buttons(genExcel) configure -state normal
      }

# none of the above
      if {!$gen(Excel) && !$gen(CSV) && !$gen(None)} {
        if {$useXL} {
          set gen(Excel) 1
          set opt(xlFormat) "Excel"
          $buttons(genExcel) configure -state normal
        } else {
          set gen(CSV) 1
          set opt(xlFormat) "CSV"
          $buttons(genExcel) configure -state disabled
        }
      }
      checkValues
      set gen(Excel1) $gen(Excel)
      set gen(View1) $gen(View)
    }]

    pack $buttons($idx) -side left -anchor w -padx {5 0} -pady {0 3} -ipady 0
    incr cb

    if {$idx == "genCSV"} {
      pack [ttk::separator $foptk.$cb -orient vertical] -side left -anchor w -padx {10 5} -pady {0 3} -ipady 9
      incr cb
    }
  }

# part only
  foreach item {{" Part Only" opt(partOnly)} {" BOM" opt(BOM)} {" Syntax Checking" opt(syntaxChecker)} {" Log File" opt(logFile)} {" Open Files  " opt(outputOpen)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $foptk.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx {5 0} -pady {0 3} -ipady 0
    incr cb

    if {$idx != "outputOpen"} {
      pack [ttk::separator $foptk.$cb -orient vertical] -side left -anchor w -padx {10 5} -pady {0 3} -ipady 9
      incr cb
    }
  }

  pack $foptk -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true

  set txt "Spreadsheets contain one worksheet for each STEP entity type.  The categories below\ncontrol which STEP entity types are written to the Spreadsheet.  Analyzer options\nbelow also write information to the Spreadsheet.  See the More tab for more options.\n\nIf Excel is installed, then Spreadsheets and CSV files can be generated.  If CSV Files\nis selected, the Spreadsheet is also generated.  CSV files do not contain any cell\ncolors, comments, or links.  GD&T symbols in CSV files are only supported with\nExcel 2016 or newer.\n\nIf Excel is not installed, only CSV files can be generated.  Analyzer options are disabled."
  catch {tooltip::tooltip $buttons(genExcel) $txt}
  catch {tooltip::tooltip $buttons(genCSV) $txt}
  set txt "The Viewer supports b-rep and tessellated part geometry, graphic PMI, sketch\ngeometry, supplemental geometry, datum targets, finite element models, and more.\nUse the Viewer options below to control what features of the STEP file are shown.\n\nPart Only generates only Part Geometry.  This is useful when no other Viewer\nfeatures are needed and for large STEP files.\n\nSee Help > Viewer"
  catch {tooltip::tooltip $buttons(genView) $txt}
  catch {tooltip::tooltip $buttons(partOnly) $txt}
  catch {tooltip::tooltip $buttons(BOM) "Generate a Bill of Materials (BOM) of parts and assemblies\n\nSee Help > Bill of Materials\nSee Examples > Bill of Materials"}

  catch {tooltip::tooltip $buttons(logFile) "Status tab text can be written to a Log file myfile-sfa.log\nUse F4 to open the Log file.\nSyntax Checker results are written to myfile-sfa-err.log\n\nAll text in the Status tab can be saved by right-clicking\nand selecting Save."}
  catch {tooltip::tooltip $buttons(syntaxChecker) "Use this option to run the Syntax Checker when generating a Spreadsheet\nor View.  The Syntax Checker can also be run with function key F8.\n\nThis checks for basic syntax errors and warnings in the STEP file related to\nmissing or extra attributes, incompatible and unresolved\ entity references,\nselect value types, illegal and unexpected characters, and other problems\nwith entity attributes.\n\nSee Help > Syntax Checker\nSee Help > User Guide (section 7)"}
  catch {tooltip::tooltip $buttons(outputOpen) "If output files are not opened after they have been generated,\nthey can be opened with functions keys.  See Help > Function Keys\n\nIf possible, existing output files are always overwritten by new files.\nOutput files can be written to a user-defined directory (More tab)."}
  pack $foptOF -side top -anchor w -pady 0 -fill x

#-------------------------------------------------------------------------------
# process section
  set fopta [ttk::labelframe $fopt.a -text " Entity Types "]

# option to process user-defined entities
  guiUserDefinedEntities

# entity categories
  set fopta1 [frame $fopta.1 -bd 0]
  foreach item {{" Common" opt(stepCOMM)} {" Presentation" opt(stepPRES)} {" Representation" opt(stepREPR)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $fopta1.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    if {[info exists entCategory($idx)]} {
      set ttmsg ""
      if {$idx == "stepCOMM"} {
        append ttmsg "Entity Type categories control which entities from AP203, AP214, and AP242 are written to the Spreadsheet.  The categories\nare used to group and color-code entities on the Summary worksheet.  All entities specific to other APs are always written\nto the Spreadsheet.  See Help > Supported STEP APs\n\n"
      }
      append ttmsg "[llength $entCategory($idx)] [string trim [lindex $item 0]] entities are supported in most STEP APs."
      set ttmsg [guiToolTip $ttmsg $idx [string trim [lindex $item 0]]]
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta1 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta2 [frame $fopta.2 -bd 0]
  foreach item {{" Measure" opt(stepQUAN)} {" Shape Aspect" opt(stepSHAP)} {" Tolerance" opt(stepTOLR)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $fopta2.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    if {[info exists entCategory($idx)]} {
      set str "most"
      if {$idx == "stepTOLR"} {set str "some"}
      set ttmsg "[llength $entCategory($idx)] [string trim [lindex $item 0]] entities are supported in $str STEP APs."
      set ttmsg [guiToolTip $ttmsg $idx [string trim [lindex $item 0]]]
      if {$idx == "stepTOLR"} {
        append ttmsg "\n\nSee Websites > Recommended Practice for $recPracNames(pmi242)"
        append ttmsg "\nTolerance entities are based on ISO 10303 Part 47 - Shape variation tolerances"
      }
      if {$idx == "stepSHAP"} {append ttmsg "\n\nOther Shape Aspect entities are in the Tolerance and Features categories."}
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta2 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta3 [frame $fopta.3 -bd 0]
  foreach item {{" Geometry" opt(stepGEOM)} {" Coordinates" opt(stepCPNT)} {" Features" opt(stepFEAT)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $fopta3.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    if {[info exists entCategory($idx)]} {
      if {$idx != "stepCPNT"} {
        set ttmsg "[llength $entCategory($idx)] [string trim [lindex $item 0]] entities are supported in"
        if {$idx != "stepFEAT"} {
          append ttmsg " most STEP APs."
        } else {
          append ttmsg " AP214 and AP242."
        }
      } else {
        set ttmsg "[llength $entCategory($idx)] [string trim [lindex $item 0]] entities"
      }
      set ttmsg [guiToolTip $ttmsg $idx [string trim [lindex $item 0]]]
      if {$idx == "stepGEOM"} {append ttmsg "\n\nGeometry entities are based on ISO 10303 Part 42 - Geometric and topological representation"}
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta3 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta4 [frame $fopta.4 -bd 0]
  foreach item {{" Kinematics" opt(stepKINE)} {" Composites" opt(stepCOMP)} {" AP242" opt(stepAP242)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $fopta4.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    if {[info exists entCategory($idx)]} {
      set ttmsg "[llength $entCategory($idx)] [string trim [lindex $item 0]] entities"
      if {$idx == "stepAP242"} {
        append ttmsg " are supported in AP242.  Commonly used AP242 entities are in the other Entity Type categories.\nSuperscript indicates edition of AP242"
      } else {
        append ttmsg " are supported in"
        if {$idx == "stepCOMP"} {append ttmsg " AP203 and"}
        if {$idx == "stepKINE"} {append ttmsg " AP214 and"}
        append ttmsg " AP242."
      }
      set ttmsg [guiToolTip $ttmsg $idx [string trim [lindex $item 0]]]
      if {$idx == "stepKINE"} {append ttmsg "\n\nKinematics is now implemented with the AP242 Domain Model XML.  See Websites > CAx Recommended Practices"}
      if {$idx == "stepCOMP"} {append ttmsg "\n\nSee Websites > Recommended Practices for Composite Materials"}
      if {$idx == "stepAP242"} {append ttmsg "\n\nAssembly Structure is also supported by the AP242 Domain Model XML.  See Websites > CAx Recommended Practices"}
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta4 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta5 [frame $fopta.5 -bd 0]
  foreach item {{" Quality" opt(stepQUAL)} {" Constraint" opt(stepCONS)} {" Other" opt(stepOTHR)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $fopta5.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    if {[info exists entCategory($idx)]} {
      set ttmsg "[llength $entCategory($idx)] [string trim [lindex $item 0]] entities are supported in "
      if {$idx != "stepOTHR"} {
        append ttmsg "AP242.  "
      } else {
        append ttmsg "some STEP APs.  "
      }
      if {$idx == "stepQUAL"} {append ttmsg "\n"}
      if {$idx != "stepOTHR"} {append ttmsg "Superscript indicates edition of AP242."}
      set ttmsg [guiToolTip $ttmsg $idx [string trim [lindex $item 0]]]
      if {$idx == "stepQUAL"} {append ttmsg "\n\nQuality entities are based on ISO 10303 Part 59 - Quality of product shape data"}
      if {$idx == "stepCONS"} {append ttmsg "\n\nConstraint entities are based on ISO 10303 Parts 108 and 109"}
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta5 -side left -anchor w -pady 0 -padx 0 -fill y

  set fopta9 [frame $fopta.9 -bd 0]
  foreach item {{"All" 0} {"Reset" 1}} {
    set bn "allNone[lindex $item 1]"
    set buttons($bn) [ttk::radiobutton $fopta9.$cb -variable allNone -text [lindex $item 0] -value [lindex $item 1] \
      -command {
        if {$allNone == 0} {
          foreach item [array names opt] {
            if {[string first "step" $item] == 0 && $item != "stepUSER"} {set opt($item) 1}
          }
        } elseif {$allNone == 1} {
          foreach item [array names opt] {if {[string first "step" $item] == 0} {set opt($item) 0}}
          foreach item {BOM INVERSE PMIGRF PMISEM valProp stepUSER x3dSave} {set opt($item) 0}
          set opt(stepCOMM) 1
          set gen(None) 0
          set gen(Excel) 1
          if {$opt(xlFormat) == "None"} {set opt(xlFormat) "Excel"}
          if {!$gen(CSV)} {$buttons(genExcel) configure -state normal}
        }
        checkValues
      }]
    pack $buttons($bn) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fopta9 -side left -anchor w -pady 0 -padx 15 -fill y
  pack $fopta -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {tooltip::tooltip $fopta "Entity Type categories control which entities from AP203, AP214, and AP242\nare written to the Spreadsheet.  The categories are used to group and\ncolor-code entities on the Summary worksheet.  All entities specific to other\nAPs are always written to the Spreadsheet.\n\nSee Help > Supported STEP APs"}

#-------------------------------------------------------------------------------
# analyzer section
  set foptRV [frame $fopt.rv -bd 0]
  set foptd [ttk::labelframe $foptRV.1 -text " Analyzer "]
  set foptd1 [frame $foptd.1 -bd 0]

  foreach item {{" Validation Properties" opt(valProp)} \
                {" AP242 Semantic Representation PMI" opt(PMISEM)} \
                {" Graphic Presentation PMI" opt(PMIGRF)} \
                {" Inverse Relationships and Backwards References" opt(INVERSE)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $foptd1.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx {5 10} -pady {0 5} -ipady 0
    incr cb
  }
  pack $foptd1 -side top -anchor w -pady 0 -padx 0 -fill y
  pack $foptd -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true

  catch {
    tooltip::tooltip $buttons(valProp) "Geometric, assembly, PMI, annotation, attribute, tessellated, composite, and FEA\nvalidation properties, and semantic text are reported.  Properties are shown on\nthe 'property_definition' and other entities.  Some properties are reported only if\nAnalyzer option for Semantic PMI is selected.  Some properties might not be\nshown depending on the value of Maximum Rows (More tab).\n\nSee Help > Analyzer > Validation Properties\nSee Help > User Guide (section 6.3)\nSee Help > Analyzer > Syntax Errors\n\nValidation properties must conform to recommended practices.\nSee Websites > CAx Recommended Practices"
    tooltip::tooltip $buttons(PMISEM)  "Semantic PMI is the information necessary to represent geometric\nand dimensional tolerances without any graphic PMI.  It is shown\non dimension, tolerance, datum target, and datum entities.\nSemantic PMI is mainly in STEP AP242 files.  See the More tab for\nmore options.\n\nSee Help > Analyzer > Semantic Representation PMI\nSee Help > User Guide (section 6.1)\nSee Help > Analyzer > Syntax Errors\nSee Websites > AP242\n\nSemantic PMI must conform to recommended practices.\nSee Websites > CAx Recommended Practices"
    tooltip::tooltip $buttons(PMIGRF)  "Graphic PMI is the geometric elements necessary to draw annotations.\nThe information is shown on 'annotation occurrence' entities.\n\nSee Help > Analyzer > Graphic Presentation PMI\nSee Help > User Guide (section 6.2)\nSee Help > Analyzer > Syntax Errors\n\nGraphic PMI must conform to recommended practices.\nSee Websites > CAx Recommended Practices"

    set ttmsg "Inverse Relationships and Backwards References (Used In) are reported for some attributes for these entities in\nadditional columns highlighted in light blue and purple.  This option is useful for debugging some Syntax Errors\nand finding missing relationships and references.  See Help > User Guide (section 6.4)"
    set ttmsg [guiToolTip $ttmsg "inverses" "Inverse"]
    tooltip::tooltip $buttons(INVERSE) $ttmsg
  }

#-------------------------------------------------------------------------------
# viewer section
  set foptv [ttk::labelframe $foptRV.9 -text " Viewer "]
  set foptv20 [frame $foptv.20 -bd 0]

# part geometry
  foreach item {{" Part Geometry" opt(viewPart)} {" Edges" opt(partEdges)} {" Sketch" opt(partSketch)} {" Supplemental" opt(partSupp)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $foptv20.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv20 -side top -anchor w -pady 0 -padx 0 -fill y

# part quality, normals
  set foptv21 [frame $foptv.21 -bd 0]
  set buttons(labelPartQuality) [label $foptv21.l3 -text "Quality: "]
  pack $foptv21.l3 -side left -anchor w -padx 0 -pady 0 -ipady 0
  foreach item {{Low 4} {Normal 7} {High 10}} {
    set bn "partQuality[lindex $item 1]"
    set buttons($bn) [ttk::radiobutton $foptv21.$cb -variable opt(partQuality) -text [lindex $item 0] -value [lindex $item 1]]
    pack $buttons($bn) -side left -anchor w -padx 2 -pady 0 -ipady 0
    incr cb
  }

  set item {" Normals" opt(partNormals)}
  set idx [string range [lindex $item 1] 4 end-1]
  set buttons($idx) [ttk::checkbutton $foptv21.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side left -anchor w -padx 10 -pady 0 -ipady 0
  incr cb
  pack $foptv21 -side top -anchor w -pady {0 5} -padx {26 10} -fill y

# graphic PMI
  set foptv3 [frame $foptv.3 -bd 0]
  set item {" Graphic PMI" opt(viewPMI)}
  set idx [string range [lindex $item 1] 4 end-1]
  set buttons($idx) [ttk::checkbutton $foptv3.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
  pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
  incr cb
  pack $foptv3 -side top -anchor w -pady 0 -padx 0 -fill y

  set foptv4 [frame $foptv.4 -bd 0]
  set buttons(labelPMIcolor) [label $foptv4.l3 -text "PMI Color: "]
  pack $foptv4.l3 -side left -anchor w -padx 0 -pady 0 -ipady 0
  foreach item {{"From File " 0} {"Black " 1} {"By View " 3} {"Random" 2}} {
    set bn "gpmiColor[lindex $item 1]"
    set buttons($bn) [ttk::radiobutton $foptv4.$cb -variable opt(gpmiColor) -text [lindex $item 0] -value [lindex $item 1]]
    pack $buttons($bn) -side left -anchor w -padx 2 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv4 -side top -anchor w -pady {0 5} -padx {26 10} -fill y

# finite element model
  set foptv7 [frame $foptv.7 -bd 0]
  foreach item {{" AP209 Finite Element Model" opt(viewFEA)} {" Boundary conditions" opt(feaBounds)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $foptv7.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptv7 -side top -anchor w -pady 0 -padx 0 -fill y

  set foptv8 [frame $foptv.8 -bd 0]
  foreach item {{"Loads" opt(feaLoads)} {"Scale loads  " opt(feaLoadScale)} {"Displacements" opt(feaDisp)} {"No vector tail" opt(feaDispNoTail)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $foptv8.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 2 -pady {0 3} -ipady 0
    incr cb
  }
  pack $foptv8 -side top -anchor w -pady 0 -padx {26 10} -fill y

  pack $foptv -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true
  pack $foptRV -side top -anchor w -pady 0 -fill x
  catch {
    tooltip::tooltip $foptv20 "The viewer supports b-rep and AP242 tessellated part geometry, color,\ntransparency, edges, sketch and supplemental geometry, and clipping planes.\nTesselated part geometry is typically written to an AP242 file instead of or in\naddition to b-rep part geometry.\n\nThe viewer uses the default web browser.  An Internet connection is required.\nThe viewer does not support measurements.\n\nSee the More tab for more Viewer options.\nSee Help > Viewer > Overview and other topics"
    tooltip::tooltip $buttons(viewPMI) "Graphic PMI for annotations is supported in AP242, AP203, and AP214 files.\nPMI (annotation) placeholders are supported in AP242.\n\nA Saved View is a subset of graphic PMI which has its own viewpoint position\nand orientation.  Use PageDown in the viewer to cycle through saved views to\nswitch to the associated viewpoint and subset of graphic PMI.\n\nSee the options related to viewpoints on the More tab.\nSee Help > Viewer > Graphic PMI\nSee Help > Viewer > PMI Placeholders\nSee Help > User Guide (section 4.2)\n\nGraphic PMI and placeholders must conform to recommended practices.\nSee Websites > CAx Recommended Practices"
    tooltip::tooltip $buttons(feaLoadScale) "The length of load vectors can be scaled by their magnitude.\nLoad vectors are always colored by their magnitude."
    tooltip::tooltip $buttons(feaDispNoTail) "The length of displacement vectors with a tail are scaled by\ntheir magnitude.  Vectors without a tail are not.\nDisplacement vectors are always colored by their magnitude.\nLoad vectors always have a tail."
    tooltip::tooltip $foptv21 "Quality controls the number of facets used for curved surfaces.\nFor example, the higher the quality the more facets around the\ncircumference of a cylinder.\n\nNormals improve the default smooth shading of surfaces.  Using\nHigh Quality and Normals results in the best appearance for b-rep\npart geometry.  Quality and normals do not apply to tessellated\npart geometry.\n\nIf curved surfaces for b-rep part geometry look wrong even with\nQuality set to High, then use the Alternative B-rep Geometry\nProcessing method on the More tab."
    tooltip::tooltip $foptv4  "For 'By View' PMI colors, each Saved View is set to a different color.  If there\nis only one or no Saved Views, then 'Random' PMI colors are used.\n\nFor 'Random' PMI colors, each annotation is set to a different color to help\ndifferentiate one from another.\n\nPMI color does not apply to annotation placeholders which are always black."
    set tt "FEM nodes, elements, boundary conditions, loads, and\ndisplacements in AP209 files are shown.\n\nSee Help > Viewer > AP209 Finite Element Model\nSee Help > User Guide (section 4.4)"
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
  catch {tooltip::tooltip $fopta6 "A User-Defined List is a plain text file with one STEP entity name per line.\nThis allows for more control to process only the required entities, rather\nthan process the broad categories of entity types above."}
  pack $fopta6 -side bottom -anchor w -pady 5 -padx 0 -fill y
}

#-------------------------------------------------------------------------------
# open STEP file
proc guiOpenSTEPFile {} {
  global appName appNames buttons cb developer dispApps dispCmds edmWhereRules edmWriteToFile
  global fopt foptf

  set foptOP [frame $fopt.op -bd 0]
  set foptf [ttk::labelframe $foptOP.f -text " Open STEP File in App "]

  set buttons(appCombo) [ttk::combobox $foptf.spinbox -values $appNames -width 30]
  pack $foptf.spinbox -side left -anchor w -padx 7 -pady {0 3}

  bind $buttons(appCombo) <<ComboboxSelected>> {
    set appName [$buttons(appCombo) get]

# Jotne EDM Model Checker
    if {$developer} {
      catch {
        if {[string first "EDMsdk" $appName] != -1} {
          pack $buttons(edmWriteToFile) -side left -anchor w -padx {5 0}
          pack $buttons(edmWhereRules) -side left -anchor w -padx {5 0}
        } else {
          pack forget $buttons(edmWriteToFile)
          pack forget $buttons(edmWhereRules)
        }
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
      if {[string first "EDMsdk" $item] != -1} {
        foreach item {{"Check rules" edmWhereRules} {"Write to file" edmWriteToFile}} {
          set idx [lindex $item 1]
          set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
          pack forget $buttons($idx)
          incr cb
        }
      }
    }
  }

# built-in file tree view
  if {[lsearch $appNames "Tree View (for debugging)"] != -1} {
    foreach item {{"Include Geometry" indentGeometry} {"Include styled_item" indentStyledItem}} {
      set idx [lindex $item 1]
      set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] -variable opt([lindex $item 1]) -command {checkValues}]
      pack forget $buttons($idx)
      incr cb
    }
  }

  catch {tooltip::tooltip $buttons(appCombo) "This option is a convenient way to open a STEP file in other apps.  The\npull-down menu contains some apps that can open a STEP file including\nSTEP viewers and browsers, however, only if they are installed in their\ndefault location.\n\nSee Help > Open STEP File in App\nSee Websites > STEP > STEP File Viewers\n\nThe 'Tree View (for debugging)' option rearranges and indents the entities\nto show the hierarchy of information in a STEP file.  The 'tree view' file\n(myfile-sfa.txt) is written to the same directory as the STEP file or to the\nsame user-defined directory specified in the More tab.  Including\nGeometry or Styled_item can make the 'tree view' file very large.  The\n'tree view' might not process /*comments*/ in a STEP file correctly.\n\nThe 'Default STEP Viewer' option opens the STEP file in whatever app is\nassociated with STEP (.stp, .step, .p21) files.\n\nUse F5 to open the STEP file in a text editor."}
  pack $foptf -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true
  pack $foptOP -side top -anchor w -pady 0 -fill x
}

#-------------------------------------------------------------------------------
# more tab
proc guiMoreTab {} {
  global buttons cb developer fileDir fxls mydocs nb opt pmiElementsMaxRows recPracNames userWriteDir writeDir

  set wxls [ttk::panedwindow $nb.xls -orient horizontal]
  $nb add $wxls -text " More " -padding 2
  set fxls [frame $wxls.fxls -bd 2 -relief sunken]

# spreadsheet formatting options
  set fxlsb [ttk::labelframe $fxls.b -text " Spreadsheet "]

# maximum rows
  set fxlsb0 [frame $fxlsb.0 -bd 0]
  set buttons(labelMaxRows) [label $fxlsb0.l1 -text "Maximum Rows: "]
  pack $fxlsb0.l1 -side left -anchor w -padx {3 0} -pady 0 -ipady 0
  set rlimit {{"100" 103} {"500" 503} {"1000" 1003} {"5000" 5003} {"10000" 10003} {"50000" 50003} {"100000" 100003} {"Maximum" 1048576}}
  set n 0
  foreach item $rlimit {
    set idx "maxrows$n"
    set buttons($idx) [ttk::radiobutton $fxlsb0.$cb -variable opt(xlMaxRows) -text [lindex $item 0] -value [lindex $item 1]]
    pack $buttons($idx) -side left -anchor n -padx {5 3} -pady 0 -ipady 0
    incr cb
    incr n
  }
  pack $fxlsb0 -side bottom -anchor w -pady {8 3} -padx 0 -fill y
  set msg "Maximum rows limits the number of rows (entities) written to any one worksheet.\nIf the maximum number of rows is exceeded, the number of entities processed will be reported\nas, for example, 'property_definition (100 of 147)'.  For large STEP files, setting a low maximum\ncan speed up processing at the expense of not processing all of the entities.\n\nMaximum rows is increased to 5000 for entities where Analyzer results are reported.  Syntax\nErrors might be missed if some entities are not processed due to a low value of maximum rows.\nMaximum rows does not affect the Viewer.\n\nSee Help > User Guide (section 5.5.1)"
  catch {tooltip::tooltip $fxlsb0 $msg}

# checkboxes
  set fxlsb1 [frame $fxlsb.1 -bd 0]
  set fxlsb2 [frame $fxlsb.2 -bd 0]
  set n 0
  foreach item {{" Process text strings with non-English characters" opt(xlUnicode)} \
                {" Generate tables for sorting and filtering" opt(xlSort)} \
                {" Do not round real numbers in spreadsheet cells" opt(xlNoRound)} \
                {" Do not generate links on File Summary worksheet" opt(xlHideLinks)}} {
    incr n
    set frm $fxlsb1
    if {$n > 2} {set frm $fxlsb2}
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $frm.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsb1 -side left -anchor w -pady {0 5} -padx 0 -fill y
  pack $fxlsb2 -side left -anchor w -pady {0 5} -padx 6 -fill y
  pack $fxlsb -side top -anchor w -pady {5 2} -padx 10 -fill both

# viewer options
  set fxlsd [ttk::labelframe $fxls.d -text " Viewer "]
  set fxlsda [frame $fxlsd.a -bd 0]
  set fxlsd1 [frame $fxlsda.1 -bd 0]
  set fxlsd2 [frame $fxlsda.2 -bd 0]
  set items [list {" Use parallel projection viewpoints defined in file" opt(viewParallel)} \
                  {" Show viewpoints without graphic PMI" opt(viewNoPMI)} \
                  {" Correct for older viewpoint implementations" opt(viewCorrect)} \
                  {" Debug saved view camera model viewpoint" opt(debugVP)}]
  set n 0
  foreach item $items {
    incr n
    set frm $fxlsd1
    if {$n > 2} {set frm $fxlsd2}
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $frm.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsd1 -side left -anchor w -pady {0 5} -padx 0 -fill y
  pack $fxlsd2 -side left -anchor w -pady {0 5} -padx 6 -fill y
  pack $fxlsda -side top -anchor w -pady 0 -padx 0 -fill both

  set fxlsdb [frame $fxlsd.b -bd 0]
  set fxlsd3 [frame $fxlsdb.3 -bd 0]
  set fxlsd4 [frame $fxlsdb.4 -bd 0]
  set items [list {" Generate capped surfaces for clipping planes" opt(partCap)} \
                  {" Do not group identical parts in an assembly" opt(partNoGroup)} \
                  {" Save X3D file generated by the Viewer" opt(x3dSave)} \
                  {" Alternative processing of tessellated part geometry" opt(tessPartOld)} \
                  {" Alternative processing of b-rep part geometry" opt(brepAlt)}]
  set n 0
  foreach item $items {
    incr n
    set frm $fxlsd3
    if {$n > 3} {set frm $fxlsd4}
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $frm.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsd3 -side left -anchor w -pady {0 5} -padx 0 -fill y
  pack $fxlsd4 -side left -anchor w -pady {0 5} -padx 6 -fill y
  pack $fxlsdb -side top -anchor w -pady 0 -padx 0 -fill both
  pack $fxlsd -side top -anchor w -pady {10 2} -padx 10 -fill both

# other analyzer options
  set fxlsa [ttk::labelframe $fxls.a -text " Analyzer "]
  foreach item {{" Round dimensions and geometric tolerances for semantic PMI" opt(PMISEMRND)} \
                {" Show all PMI Elements on Semantic PMI Coverage worksheet" opt(SHOWALLPMI)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $fxlsa.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx {5 10} -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsa -side top -anchor w -pady {10 2} -padx 10 -fill both

  catch {
    tooltip::tooltip $fxlsd                "These Viewer options should be selected only if necessary.\nRead the tooltips for each individual option."
    tooltip::tooltip $buttons(xlUnicode)   "Use this option if there are non-English characters or symbols\nencoded with the \\X2\\ control directive in the STEP file.\n\nSee Help > Text Strings and Numbers\nSee User Guide (section 5.5.2)"
    tooltip::tooltip $buttons(xlSort)      "Worksheets can be sorted by column values.\nWorksheets related to Analyzer options are always sorted.\n\nSee Help > User Guide (section 5.5.3)"
    tooltip::tooltip $buttons(xlNoRound)   "See Help > User Guide (section 5.5.4)"
    tooltip::tooltip $buttons(SHOWALLPMI)  "The complete list of [expr {$pmiElementsMaxRows-3}] PMI Elements, including those that are not in the\nSTEP file, will be shown on the Semantic PMI Coverage worksheet.\n\nSee Help > Analyzer > PMI Coverage Analysis\nSee Help > User Guide (section 6.1.7)"
    tooltip::tooltip $buttons(xlHideLinks) "This option is useful when sharing a Spreadsheet with another user."
    tooltip::tooltip $buttons(PMISEMRND)   "Rounding values might result in a better match to graphic PMI shown in\nthe viewer or to expected PMI in the NIST CAD models (FTC/STC 7, 8, 11).\n\nSee User Guide (section 6.1.3.1)\nSee Websites > Recommended Practice for\n   $recPracNames(pmi242), Section 5.4"
    tooltip::tooltip $buttons(viewParallel) "Use parallel projection defined in the STEP file for saved view viewpoints,\ninstead of the default perspective projection.  Pan and zoom might not\nwork with parallel projection.  See Help > Viewer > Viewpoints"
    tooltip::tooltip $buttons(viewCorrect) "Correct for older implementations of camera models that\nmight not conform to current recommended practices.\nThe corrected viewpoint might fix the orientation but\nmaybe not the position.\n\nSee Help > Viewer > Viewpoints\nSee the CAx-IF Recommended Practice for\n $recPracNames(pmi242), Sec. 9.4.2.6"
    tooltip::tooltip $buttons(viewNoPMI)   "If the model has viewpoints with and without graphic PMI,\nthen also show the viewpoints without graphic PMI.  Those\nviewpoints are typically top, front, side, etc."
    tooltip::tooltip $buttons(debugVP)     "Debug viewpoint orientation defined by a camera model\nby showing the view frustum in the viewer.\n\nSee Help > Viewer > Viewpoints\nSee the CAx-IF Recommended Practice for\n $recPracNames(pmi242), Sec. 9.4.2.6"
    tooltip::tooltip $buttons(partCap)     "Generate capped surfaces for section view clipping planes.  Capped\nsurfaces might take a long time to generate or look wrong for parts\nin an assembly.  See Help > Viewer > New Features"
    tooltip::tooltip $buttons(brepAlt)     "If curved surfaces for Part Geometry look wrong even with\nQuality set to High, use an alternative b-rep geometry\nprocessing algorithm.  It will take longer to process the STEP\nfile and the resulting Viewer file will be larger."
    tooltip::tooltip $buttons(tessPartOld) "Process AP242 tessellated part geometry with the old method in SFA < 5.10.\nIt is not recommended for assemblies or large STEP files."
    tooltip::tooltip $buttons(x3dSave)     "The X3D file can be shown in an X3D viewer or imported to other software.\nUse this option if an Internet connection is not available for the Viewer.\nSee Help > Viewer"
    tooltip::tooltip $buttons(partNoGroup) "This option might create a very long list of parts names in the viewer.\nIdentical parts have a underscore and number appended to their name.\nSee Help > Assemblies"
  }

# output directory
  set fxlse [ttk::labelframe $fxls.e -text " Write Output to "]
  set buttons(fileDir) [ttk::radiobutton $fxlse.$cb -text " Same directory as STEP file" -variable opt(writeDirType) -value 0 -command checkValues]
  pack $fxlse.$cb -side left -anchor w -padx 5 -pady 2
  incr cb

  set fxls1 [frame $fxlse.1]
  ttk::radiobutton $fxls1.$cb -text " User-defined directory: " -variable opt(writeDirType) -value 2 -command {
    checkValues
    if {[file exists $userWriteDir] && [file isdirectory $userWriteDir]} {
      set writeDir $userWriteDir
    } else {
      set userWriteDir $mydocs
      tk_messageBox -type ok -icon error -title "Bad Directory" \
        -message "The user-defined directory to write the Spreadsheet is bad.\nIt has been set to $userWriteDir"
    }
    focus $buttons(userdir)
  }
  pack $fxls1.$cb -side left -anchor w -padx {5 0}
  catch {tooltip::tooltip $fxls1 "This option is useful when the directory containing the STEP file is\nprotected (read-only) and none of the output can be written to it.\nDo not use a directory name containing bracket \[\] characters."}
  incr cb

  set buttons(userentry) [ttk::entry $fxls1.entry -width 35 -textvariable userWriteDir]
  pack $fxls1.entry -side left -anchor w -pady 2
  set buttons(userdir) [ttk::button $fxls1.button -text "Browse" -command {
    set uwd [tk_chooseDirectory -title "Select directory"]
    if {[file isdirectory $uwd]} {
      set userWriteDir $uwd
      set writeDir $userWriteDir
    }
  }]
  pack $fxls1.button -side left -anchor w -padx 5 -pady 2
  pack $fxls1 -side top -anchor w

  pack $fxlse -side top -anchor w -pady {10 2} -padx 10 -fill both
  catch {tooltip::tooltip $fxlse "If possible, existing output files are always overwritten by new files.\nIf spreadsheets cannot be overwritten, a number is appended to the\nfile name: myfile-sfa-1.xlsx"}

# developer only options
  if {$developer} {
    set fxlsx [ttk::labelframe $fxls.x -text " Developer "]
    foreach item {{" Viewer" opt(debugX3D)} {" Analyzer" opt(DEBUG1)} {" Assoc Geom" opt(debugAG)} {" Inverses" opt(DEBUGINV)} {" Dimensions" opt(PMISEMDIM)} {" Datum Targets" opt(PMISEMDT)} {" No Excel" opt(debugNOXL)}} {
      set idx [string range [lindex $item 1] 4 end-1]
      set buttons($idx) [ttk::checkbutton $fxlsx.$cb -text [lindex $item 0] -variable [lindex $item 1]]
      pack $buttons($idx) -side left -anchor w -padx {5 0} -pady 0 -ipady 0
      incr cb
    }
    pack $fxlsx -side top -anchor w -pady {1 0} -padx 10 -fill both
    catch {
      tooltip::tooltip $buttons(debugX3D)  "Debug running stp2x3d and processing the resulting x3d file for the Viewer"
      tooltip::tooltip $buttons(DEBUG1)    "Debug Analyzer reports"
      tooltip::tooltip $buttons(debugAG)   "Debug computing Associated Geometry for any PMI"
      tooltip::tooltip $buttons(DEBUGINV)  "Debug Inverse Relationships"
      tooltip::tooltip $buttons(PMISEMDIM) "Process ONLY Dimensional Tolerances"
      tooltip::tooltip $buttons(PMISEMDT)  "Process ONLY Datum Targets"
      tooltip::tooltip $buttons(debugNOXL) "Simulate if Excel is not installed forcing only CSV files to be generated"
    }
  }

  pack $fxls -side top -fill both -expand true -anchor nw
}


#-------------------------------------------------------------------------------
# help menu
proc guiHelpMenu {} {
  global developer Examples filesProcessed Help ifcsvrDir ifcsvrVer mytemp opt scriptName stepAPs

  $Help add command -label "User Guide" -command {openUserGuide}
  $Help add command -label "Release Notes" -command {openURL https://www.nist.gov/document/sfa-release-notes}

  $Help add separator
  $Help add command -label "Overview" -command {
outputMsg "\nOverview ------------------------------------------------------------------------------------------" blue
outputMsg "The STEP File Analyzer and Viewer (SFA) opens a STEP file (ISO 10303 - STandard for Exchange of
Product model data) Part 21 file (.stp or .step or .p21 file extension) and

1 - generates an Excel spreadsheet or CSV files of all entity and attribute information,
2 - creates a visualization (view) of part geometry, graphic PMI, and other features that is
    displayed in a web browser,
3 - reports and analyzes validation properties, semantic PMI, and graphic PMI, and checks them for
    conformance to recommended practices, and
4 - checks for basic syntax errors.

Compressed STEP files (.stpZ) and STEP archive files (.stpA) are supported.
AP238 STEP-NC files (.stpnc) are supported by renaming the file extension to .stp
AP242 Domain Model XML files (.stpx) are not supported.

Help is available in this menu, in the User Guide, and in tooltip help.  New features are listed in
the Release Notes and described in some Help.  Help in the menu, tooltips, and spreadsheet comments
are more up-to-date than the User Guide."
    .tnb select .tnb.status
  }

# options help
  $Help add command -label "Options" -command {
outputMsg "\nOptions -------------------------------------------------------------------------------------------" blue
outputMsg "See Help > User Guide (sections 3.4, 3.5, 4, and 6)

Generate: Generate Excel spreadsheets, CSV files, and/or Views.  If Excel is not installed, CSV
files are automatically generated.  Some options are not supported with CSV files.  The Syntax
Checker can also be run when processing a STEP file.

All text in the Status tab can be written to a Log File when a STEP file is processed.  The log
file is written to myfile-sfa.log.  Syntax errors, warnings, and other messages are highlighted by
asterisks *.  Use F4 to open the log file.

Entity Types: Select which types of entities are processed from AP203, AP214, and AP242 for the
Spreadsheet.  All entities specific to other APs are always written to the Spreadsheet such as
AP209, AP210, and AP238.  The categories are used to group and color-code entities on the Summary
worksheet.  The tooltip help lists all the entities associated with that type.

Analyzer options report PMI and check for conformance to recommended practices.
- Semantic Representation PMI: Dimensional tolerances, geometric tolerances, and datum features are
  reported on various entities indicated by Semantic PMI on the Summary worksheet.
- Graphic Presentation PMI: Geometric entities used for Graphic PMI annotations are reported.
  Associated Saved Views, Validation Properties, and Geometry are also reported.
- Validation Properties: Geometric, assembly, PMI, annotation, attribute, and tessellated
  validation properties are reported.
- Inverse Relationships: For some entities, Inverse relationships and backwards references (Used In)
  are shown on the worksheets.

Viewer: Part geometry, graphic PMI annotations, tessellated part geometry in AP242 files, and AP209
finite element models can be shown in a web browser.

More tab: Spreadsheet formatting and other Analyzer and Viewer options."
    .tnb select .tnb.status
  }

  $Help add cascade -label "Viewer" -menu $Help.1
  set helpView [menu $Help.1 -tearoff 1]

# general viewer help
  $helpView add command -label "Overview" -command {
outputMsg "\nViewer Overview -----------------------------------------------------------------------------------" blue
outputMsg "The viewer generates an HTML file 'myfile-sfa.html' that is shown in the default web browser.  An
Internet connection is required.  The HTML file is self-contained and can be shared with other
users including those on non-Windows systems.  The viewer does not support measurements.

The viewer can be used without generating a spreadsheet.  See Generate on the Options tab.  The
Part Only option is useful when no other Viewer features are needed and for large STEP files.

The viewer supports boundary representation (b-rep) and AP242 tessellated part geometry, both with
color, transparency, part edges, sketch geometry, and assemblies.  Polyhedral b-rep geometry is not
supported.  B-rep geometry is also known as exact geometry.  Part geometry viewer features:

- Part edges are shown in black.  Use the transparency slider to show only edges.  Some parts might
  not be affected by the transparency slider.  If a part is completely transparent and edges are
  not selected, then the part will not be visible.  In some cases transparency might look wrong for
  assemblies with many parts.

- Sketch geometry is supplemental lines created when generating a CAD model.  Sketch geometry is
  also known as construction, auxiliary, support, or reference geometry.  To show only sketch
  geometry, turn off edges and make the part completely transparent.  Sometimes processing sketch
  geometry will affect the behavior of the transparency slider.  Sketch geometry is not same as
  supplemental geometry.  See Help > Viewer > Supplemental Geometry

- Normals improve the default smooth shading by explicitly computing surface normals to improve the
  appearance of curved surfaces.

- Quality controls the number of facets used for curved surfaces.  Higher quality uses more facets
  around the circumference of a cylinder.  Using High Quality and the Normals options results in
  the best appearance for part geometry.  See the new feature below for the Alternative B-rep
  Geometry Processing.

- AP242 tessellated part geometry is typically written to a STEP file in addition to or instead of
  b-rep part geometry.  A wireframe mesh, outlining the facets of the tessellated surfaces can be
  shown with the Edges option.  In some cases, parts in an assembly might have the wrong position and
  orientation or be missing.  Quality and normals do not apply to tessellated part geometry.
  See Websites > CAx Recommended Practices (Tessellated 3D Geometry)

- The bounding box min and max XYZ coordinates are based on the faceted geometry being shown and
  not the exact geometry in the STEP file.  There might be a variation in the coordinates depending
  on the Quality option.  The bounding box also accounts for any sketch geometry if it is displayed
  but not graphic PMI and supplemental geometry.  The bounding box can be shown to confirm that the
  min and max coordinates are correct.  If the part is too large to rotate smoothly, turn off the
  part and rotate the bounding box.

- The origin of the model at '0 0 0' is shown with a small XYZ coordinate axis that can be switched
  off.  The background color can be changed between white, blue, gray, and black.

- See Help > Text Strings and Numbers for how non-English characters are supported.

For very large STEP files it might take several minutes to process the STEP part geometry.  To
speed up the process, select View and Part Only on the Generate tab.  The resulting HTML file might
also take several minutes to display in the web browser.  Select 'Wait' if the web browser prompts
that it is running slowly when opening the HTML file.

The viewer generates an X3D file that is embedded in the HTML file that is displayed in the default
web browser.  Select 'Save X3D ...' on the More tab to save the X3D file so that it can be shown in
an X3D viewer or imported to other software.  Part geometry including tessellated geometry and
graphic PMI is supported.  Use this option if an Internet connection is not available for the
Viewer.

See Help > User Guide (section 4)
See Help > Viewer for other topics

The viewer for part geometry is based on the NIST STEP to X3D Translator and only runs on 64-bit
computers.  It runs a separate program stp2x3d-part.exe from [file nativename $mytemp]
See Websites > STEP

Other STEP file viewers are available.  See Websites > STEP > STEP File Viewers.  Some of the
viewers are faster and have better features for viewing and measuring part geometry.  This viewer
supports many features that other viewers do not, including: graphic PMI, sketch geometry,
supplemental geometry, datum targets, viewpoints, clipping planes, point clouds, composite rosettes,
hole features, AP242 tessellated part geometry, and AP209 finite element models and results.  Try
the Open STEP Viewer with AP242 Domain Model XML files (.stpx)."
    .tnb select .tnb.status
  }

  $helpView add command -label "New Features" -command {
outputMsg "\nNew Features --------------------------------------------------------------------------------------" blue
outputMsg "These Viewer features are not documented in the User Guide (Update 7).

1 - Hidden buttons and sliders

Some checkboxes and sliders on the right side of the viewer might be hidden.  If hidden, they can
be shown by clicking on the buttons for More Options, Saved View Graphic PMI, and others.

2 - Cloud of points and point clouds

The cloud of points (COPS) geometric validation property are sampling points generated by the CAD
system on the surfaces and edges of a part.  The points are used to check the deviation of surfaces
from those points in an importing system.  The report for Validation Properties must be generated
to show the COPS.  See Websites > CAx Recommended Practices (Geometric and Assembly Validation Properties)

3D scanning point clouds are supported in AP242.  Point cloud colors, intensities, and normals are
not supported in the Viewer.

Points are shown with a blue dot.  In both cases, the exact points might not appear on part
surfaces because part geometry in the viewer is only a faceted approximation.  For parts in an
assembly, the COPS might have the wrong position and orientation.

3 - Section view clipping planes

Part geometry can be clipped by section view clipping planes defined in the STEP file.  Part Only
must not be selected on the Generate tab.  The planes are shown with a black square that might not
be centered on the model.  Checkboxes show the names of each clipping plane.  If there are
duplicate clipping plane names, a number in parentheses is appended to the name.  You have to
manually select the clipping plane that is associated with a viewpoint or saved view graphic PMI.

Use the option on the More tab to generate capped surfaces in the plane of a clipping plane.
Capped surfaces, in the plane of the black square, are generated when there is only one clipping
plane per section view.  Capped surfaces for parts in an assembly might be in the wrong position.
Sometimes capped surfaces are not generated.  Switching off parts in an assembly does not turn off
their capped surfaces.

4 - Parallel projection viewpoints

Use the option on the More tab to use parallel projection for saved view viewpoints defined in the
STEP file, instead of the default perspective projection.  See Help > Viewer > Viewpoints

5 - PMI placeholders

PMI placeholders provide information about the position, orientation, and organization of an
annotation without the graphic presentation of numeric values and symbols for geometric or
dimensional tolerances.  See Help > Viewer > PMI Placeholders

6 - B-rep part geometry processing

If curved surfaces for Part Geometry look wrong even with Quality set to High, select the
Alternative B-rep Geometry Processing method on the More tab.  It will take longer to process the
STEP file and the resulting Viewer file will be larger.

7 - Tessellated supplemental geometry

Supplemental geometry represented as tessellated geometry is supported.
See Help > Viewer > Supplemental Geometry

8 - Convert STL to AP242

STL files can be converted to STEP AP242 tessellated geometry that can be shown in the viewer.
In the Open File(s) dialog, change the 'Files of type' to 'STL (*.stl)'.  ASCII and binary STL
files are supported.  Tessellated geometry is not exact b-rep surfaces and may not be supported in
some CAD software.

9 - Composite rosettes defined by cartesian points and curves are shown in the viewer."
    .tnb select .tnb.status
  }

    $helpView add command -label "Viewpoints" -command {
outputMsg "\nViewpoints ----------------------------------------------------------------------------------------" blue
outputMsg "Use PageDown to switch between viewpoints in the viewer window.  Viewpoint names are shown in the
upper left corner of the viewer.  User-defined viewpoints are used with saved view graphic PMI.

If there are no user-defined viewpoints (saved views) in the STEP file, then front, side, top, and
isometric viewpoints are generated.  Since the default orientation of the part is not known, the
viewpoints might not correspond to the actual front, side, and top of the model.  The isometric
viewpoint might not be centered.  All of the viewpoints use perspective except for an additional
front parallel projection.

---------------------------------------------------------------------------------------------------
If there are user-defined viewpoints (saved views) in the STEP file, then in addition to the saved
views from the file, two additional front viewpoints named 'Front (SFA)' are generated, one
perspective and the other a parallel projection.  Pan and zoom might not work with parallel
projection.

If there are duplicate saved view names, then a number in parentheses is appended to the name.  For
example, two viewpoints named MBD_A will appear as MBD_A (1) and MBD_A (2) for the Viewpoint name
in the upper left corner of the viewer when cycling through the viewpoints with PageDown.

Saved view names with non-English characters (Unicode) are supported in the viewer if a spreadsheet
is also generated.

On the More tab, parallel projection viewpoints as defined in the STEP file can be used instead of
the default perspective.  Also, if the model has viewpoints with and without graphic PMI, then the
viewpoints without graphic PMI can also be shown.  Those viewpoints are usually top, front, and
side viewpoints.

If there is graphic PMI associated with saved views, then the PMI is automatically switched
on/off when using PageDown if 'Saved View Viewpoints' is checked on the Generate tab.  If there are
duplicate saved view names as described above, then the list of Saved View Graphic PMI will append,
after a slash, the saved view name to the PMI name.

---------------------------------------------------------------------------------------------------
In the viewer, use key 'a' to view all and 'r' to restore to the original view.  The function of
other keys is described in the link 'Use the mouse'.  Navigation uses the Examine Mode.

Sometimes a part is located far from the origin and not visible.  In this case, turn off the Origin
and Sketch Geometry and then use 'a' to view all.

Older implementations of saved views might not conform to current recommended practices.  The
resulting model orientation will look wrong.  Use the option on the More tab to correct for the
wrong orientation.  The position of the model might still be wrong.

The More tab option 'Show save view camera model viewpoints' can be used to debug the camera model
(view frustum) viewpoint geometry.

Saved views are ignored with Part Only.  Part visibility in saved views is not supported."
    .tnb select .tnb.status
  }

  $helpView add command -label "Assemblies" -command {
outputMsg "\nAssemblies ----------------------------------------------------------------------------------------" blue
outputMsg "Assemblies are related to part geometry, graphic PMI, and supplemental geometry.

Part Geometry for assemblies with b-rep or tessellated geometry is supported in the viewer.  Most
assemblies and parts can be switched on and off depending on the assembly structure.  An alphabetic
list of part and assembly names is shown on the right.

Parts with the same shape are usually grouped with the same checkbox.  Some names in the list might
have an underscore and number appended to their name.  Grouping parts and assemblies with the same
shape can be disabled with the option on the Spreadsheet tab.  In this case, parts with the same
shape will have an underscore and number appended to their name.  This might create a very long
list of part names.

Clicking on the model shows the part name in the upper left.  The part name shown may not be in the
list of assemblies and parts.  The part might be contained in a higher-level assembly that is in
the list.

Processing sketch geometry might also affect the list of names.  Some assemblies have no unique
names assigned to parts, therefore there is no list of part names.  See Help > Text Strings and
Numbers for how non-English characters in part names are supported.

Nested assemblies are also supported where one file contains the assembly structure with external
file references to individual assembly components that contain part geometry.

NOTE: Graphic PMI, supplemental geometry, capped surfaces for clipping planes, and cloud of points
on parts in an assembly is supported, however, it has not been thoroughly tested and might have the
wrong position and orientation.

Transparency for assemblies with AP242 tessellated geometry might look wrong.  In some rare cases,
parts in an assembly using tessellated geometry might have the wrong position and orientation or be
missing.

Assembly Structure is also supported by the AP242 Domain Model XML. See Websites > CAx Recommended Practices"
    .tnb select .tnb.status
  }

  $helpView add command -label "Supplemental Geometry" -command {
outputMsg "\nSupplemental Geometry -----------------------------------------------------------------------------" blue
outputMsg "Supplemental geometry is geometrical elements created in the CAD system that do not belong to the
manufactured part.  It is usually used to create other geometric shapes.  Supplemental geometry is
also known as construction, auxiliary, design, support, or reference geometry.

B-rep and tessellated geometry are supported for supplemental geometry. These types of supplemental
geometry and associated text are supported.  Colors defined in the STEP file override the default
colors below.

- Coordinate System: X axis red, Y axis green, Z axis blue
- Plane: blue transparent outlined surface (unbounded planes are with shown with a square surface)
- Cylinder: blue transparent cylinder
- Line/Circle/Ellipse: purple line/circle/ellipse (trimming with cartesian_point is not supported)
- Point: black dot
- Tessellated Surface: defined color

Supplemental geometry:
- can be optionally generated
- can be switched on and off
- is not associated with graphic PMI Saved Views
- in assemblies might have the wrong position and orientation (See Help > Assemblies)
- is counted on the PMI Coverage Analysis worksheet if a Viewer file is generated

See Websites > CAx Recommended Practices (Supplemental Geometry)"
    .tnb select .tnb.status
  }

  $helpView add command -label "Graphic PMI" -command {
outputMsg "\nGraphic PMI ---------------------------------------------------------------------------------------" blue
outputMsg "Graphic Presentation PMI annotations for geometric dimensioning and tolerancing composed of
polylines, lines, circles, and tessellated geometry are supported.  On the Generate tab, the color
of the annotations can be modified.  PMI associated with saved views can be switched on and off.

Some graphic PMI might not have equivalent or any semantic PMI in the STEP file.  Some STEP files
with semantic PMI might not have any graphic PMI.

Only graphic PMI defined in recommended practices is supported.  Older implementations of saved
view viewpoints might not conform to current recommended practices.
See Websites > CAx Recommended Practices
 (Representation and Presentation of PMI for AP242, PMI Polyline Presentation for AP203 and AP214)

Graphic PMI on parts in an assembly might have the wrong position and orientation.

See Help > User Guide (section 4.2)
See Help > Analyzer > Graphic Presentation PMI
See Examples > Viewer
See Examples > Sample STEP Files

---------------------------------------------------------------------------------------------------
Datum targets are shown only if a spreadsheet is generated with the Analyzer option for Semantic
PMI, and Part Geometry or Graphic PMI selected.

See Help > User Guide (section 4.2.2)
See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242, Sec. 6.6)"
    .tnb select .tnb.status
  }

  $helpView add command -label "PMI Placeholders" -command {
outputMsg "\nPMI Placeholders ----------------------------------------------------------------------------------" blue
outputMsg "PMI (annotation) placeholders provide information about the position, orientation, and organization
of an annotation without the graphic presentation of numeric values, symbols, and text for
geometric or dimensional tolerances.  Placeholders are associated with saved views if there is also
graphic PMI.  Placeholders are supported in AP242 editions >=2.  They are not documented in the
User Guide.

Placeholder coordinate systems are shown with an axes triad, gray sphere, and text label with the
name of the placeholder.  Leader lines and a rectangle for the annotation are shown with black
lines.  To identify which annotation a leader line is associated with, the first and last points of
a leader line have a text label.  Leader line symbols show their type and position with blue text.
Some implemetations of placeholders might not contain all of the graphic elements.

See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242, Sec. 7.2)"
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

Stresses, strains, and multiple coordinate systems are not supported.

Setting Maximum Rows (More tab) does not affect the view.  For large AP209 files, there might be
insufficient memory to process all of the elements, loads, displacements, and boundary conditions.

See Help > User Guide (section 4.4)
See Examples > Viewer
See Websites > STEP > AP209 FEA"
    .tnb select .tnb.status
  }

  $Help add cascade -label "Analyzer" -menu $Help.0
  set helpAnalyze [menu $Help.0 -tearoff 1]

# analyzer overview
  $helpAnalyze add command -label "Overview" -command {
outputMsg "\nAnalyzer Overview ---------------------------------------------------------------------------------" blue
outputMsg "The Analyzer reports information related to validation properties, semantic PMI, and graphic PMI,
and checks them for conformance to recommended practices.  Syntax Errors are reported for
nonconformance.  Entities that report this information are highlighted on the File Summary
worksheet.

Inverse relationships and Backwards References show the relationship between some entities through
other entities.

PMI Coverage Analysis shows the distribution of specific semantic PMI elements related to geometric
dimensioning and tolerancing.

If a STEP AP242 file is processed that is generated from one of the NIST CAD models, the semantic
PMI Analyzer report is color-coded by the expected PMI.

See Help > Analyzer for other topics
See Help > User Guide (section 6)
See Websites > CAx Recommended Practices
See Examples > NIST CAD Models"
    .tnb select .tnb.status
  }

# validation properties, PMI, conformance checking help
  $helpAnalyze add command -label "Validation Properties" -command {
outputMsg "\nValidation Properties -----------------------------------------------------------------------------" blue
outputMsg "Geometric, assembly, PMI, annotation, attribute, tessellated, composite, and FEA validation
properties are reported.  The property values are reported in columns highlighted in yellow and
green on the property_definition worksheet.  The worksheet can also be sorted and filtered.  All
properties might not be shown depending on the Maximum Rows set on the More tab.

The name or description attribute of the entity referred to by the property_definition definition
attribute is included in brackets.

Validation properties are also reported on their associated annotation, dimension, geometric
tolerance, and shape aspect entities.  The report includes the validation property name and names
of the properties.  Some properties are reported only if the Semantic PMI Analyzer report is
selected.  Other properties and user defined attributes are also reported.

Another type of validation property is known as Semantic Text where explicit text strings in the
STEP file can be associated with part surfaces, similar to semantic PMI.  The semantic text will
appear in the spreadsheet on shape_aspect and other related entities.  The shape aspects can be
related to their corresponding dimensional or geometric tolerance entities.  A message will
indicate when semantic text is added to entities.

The PMI validation property Equivalent Unicode String is shown on worksheets for semantic and
graphic PMI with that validation property.  The sampling points for the Cloud of Points validation
property are shown in the viewer.  See Help > Viewer > New Features.  Neither of these features are
documented in the User Guide.

Syntax errors related to validation property attribute values are also reported in the Status tab
and the relevant worksheet cells.  Syntax errors are highlighted in red.  See Help > Analyzer > Syntax Errors

Clicking on the plus '+' symbols above the columns shows other columns that contain the entity ID
and attribute name of the validation property value.  All of the other columns can be shown or
hidden by clicking the '1' or '2' in the upper right corner of the spreadsheet.

The Summary worksheet indicates if properties are reported on property_definition and other
entities.

See Help > User Guide (section 6.3)
See Examples > Graphic PMI, Validation Properties

Validation properties must conform to recommended practices.
 See Websites > CAx Recommended Practices (Geometric and Assembly Validation Properties,
  User Defined Attributes, Representation and Presentation of PMI for AP242,
  Tessellated 3D Geometry)"
    .tnb select .tnb.status
  }

  $helpAnalyze add command -label "Semantic Representation PMI" -command {
outputMsg "\nSemantic Representation PMI -----------------------------------------------------------------------" blue
outputMsg "Semantic Representation PMI includes all information necessary to represent geometric and
dimensional tolerances (GD&T) without any graphic presentation elements.  Semantic PMI is
associated with CAD model geometry and is computer-interpretable to facilitate automated
consumption by downstream applications for manufacturing, measurement, inspection, and other
processes.  Semantic Representation PMI is mainly in AP242 files.

Worksheets for the Semantic Representation PMI Analyzer report show a visual recreation of the
representation for Dimensional Tolerances, Geometric Tolerances, and Datum Features.  The results
are in columns, highlighted in yellow and green, on the relevant worksheets.  The GD&T is recreated
as best as possible given the constraints of Excel.

All of the visual recreation of Datum Systems, Dimensional Tolerances, and Geometric Tolerances
that are reported on individual worksheets are collected on one Semantic PMI Summary worksheet.

If STEP files from the NIST CAD models (Examples > NIST CAD Models) are processed, then the
Semantic PMI Summary is color-coded by the expected PMI in each CAD model.
See Help > Analyzer > NIST CAD Models

Datum Features are reported on datum_* entities.  Datum_system will show the complete Datum
Reference Frame.  Datum Targets are reported on placed_datum_target_feature.

Dimensional Tolerances are reported on the dimensional_characteristic_representation worksheet.
The dimension name, representation name, length/angle, length/angle name, plus minus bounds, and
other associated information are reported.  The relevant section in the recommended practice is
shown in the column headings.

Geometric Tolerances are reported on *_tolerance entities by showing the complete Feature Control
Frame (FCF), and possible Dimensional Tolerance and Datum Feature.  The FCF should contain the
geometry tool, tolerance zone, datum reference frame, and associated modifiers.

---------------------------------------------------------------------------------------------------
If a Dimensional Tolerance refers to the same geometric element as a Geometric Tolerance, then it
will be shown above the FCF.  If a Datum Feature refers to the same geometric face as a Geometric
Tolerance, then it is shown below the FCF.  If an expected Dimensional Tolerance is not shown above
a Geometric Tolerance, then the tolerances do not reference the same geometric element.  For
example, the dimensional tolerance referencing the curved edge of a hole and the geometric
tolerance referencing the cylindrical surfaces of the same hole.

The association of the Datum Feature with a Geometric Tolerance is based on each referring to the
same geometric element.  However, the Graphic PMI might show the Geometric Tolerance and Datum
Feature as two separate annotations with leader lines attached to the same geometric element.

The number of decimal places for dimension and geometric tolerance values can be specified in the
STEP file.  By definition the value is always truncated, however, the values can be rounded instead.
For example with the value 0.5625, the qualifier 'NR2 1.3' will truncate it to 0.562  Rounding will
show 0.563  Rounding values might result in a better match to graphic PMI shown in the viewer or to
expected PMI in the NIST CAD models.  See More tab > Analyzer > Round ...

Some syntax errors that indicate non-conformance to a CAx-IF Recommended Practices related to PMI
Representation are also reported in the Status tab and the relevant worksheet cells.  Syntax errors
are highlighted in red.  See Help > Analyzer > Syntax Errors

A Semantic PMI Coverage Analysis worksheet is also generated.

See Help > User Guide (section 6.1)
See Help > Analyzer > PMI Coverage Analysis
See Examples > Spreadsheets - Semantic PMI
See Examples > Sample STEP Files

Semantic PMI must conform to recommended practices.
 See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242)"
    .tnb select .tnb.status
  }

  $helpAnalyze add command -label "Graphic Presentation PMI" -command {
outputMsg "\Graphic Presentation PMI ---------------------------------------------------------------------------" blue
outputMsg "Graphic Presentation PMI consists of geometric elements including lines and curves preserving the
exact appearance (color, shape, positioning) of the geometric and dimensional tolerance (GD&T)
annotations.  Graphic PMI is not intended to be computer-interpretable and does not have any
representation information, although it can be linked to its corresponding Semantic PMI.

The Analyzer report for Graphic PMI supports annotation_curve_occurrence, annotation_curve,
annotation_fill_area_occurrence, and tessellated_annotation_occurrence entities.  Geometric
entities used for Graphic PMI annotations are reported in columns, highlighted in yellow and green,
on those worksheets.  Presentation Style, Saved Views, Validation Properties, Annotation Plane,
Associated Geometry, and Associated Semantic PMI are also reported.

The Summary worksheet indicates on which worksheets Graphic PMI is reported.  Some syntax errors
related to Graphic PMI are also reported in the Status tab and the relevant worksheet
cells.  Syntax errors are highlighted in red.  See Help > Analyzer > Syntax Errors

A Graphic PMI Coverage Analysis worksheet is also generated.

See Help > User Guide (section 6.2)
See Help > Viewer > Graphic PMI
See Help > Analyzer > PMI Coverage Analysis
See Examples > Viewer
See Examples > Graphic PMI, Validation Properties
See Examples > Sample STEP Files

Graphic PMI must conform to recommended practices.
 See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242,
  PMI Polyline Presentation for AP203 and AP214)"
    .tnb select .tnb.status
  }

# coverage analysis help
  $helpAnalyze add command -label "PMI Coverage Analysis" -command {
outputMsg "\nPMI Coverage Analysis -----------------------------------------------------------------------------" blue
outputMsg "PMI Coverage Analysis worksheets are generated when processing single or multiple files and when
reports for Semantic Representation PMI or Graphic Presentation PMI are selected.

Semantic PMI Coverage Analysis counts the number of PMI Elements in a STEP file for tolerances,
dimensions, datums, modifiers, and CAx-IF Recommended Practices for PMI Representation.  On the
Coverage Analysis worksheet, some PMI Elements show their associated symbol, while others show the
relevant section in the recommended practice.  PMI Elements without a section number do not have a
recommended practice for their implementation.  The PMI Elements are grouped by features related
tolerances, tolerance zones, dimensions, dimension modifiers, datums, datum targets, and other
modifiers.  The number of some modifiers, e.g., maximum material condition, does not differentiate
whether they appear in the tolerance zone definition or datum reference frame.  Rows with no count
of a PMI Element can be shown, see More tab.

Some PMI Elements might not be exported to a STEP file by your CAD system.  Some PMI Elements are
only in AP242 editions > 1.

If STEP files from the NIST CAD models (Examples > NIST CAD Models) are processed, then the
Semantic PMI Coverage Analysis worksheet is color-coded by the expected number of PMI elements in
each CAD model.  See Help > Analyzer > NIST CAD Models

The Graphic PMI Coverage Analysis counts the occurrences of the recommended name attribute defined
in the CAx-IF Recommended Practice for PMI Representation and Presentation of PMI (AP242).  The
name attribute is associated with the graphic elements used to draw a PMI annotation or placeholder.
There is no semantic PMI meaning to the name attributes.

See Help > Analyzer > Semantic Representation PMI
See Help > User Guide (sections 6.1.7 and 6.2.1)
See Examples > PMI Coverage Analysis"
    .tnb select .tnb.status
  }

  $helpAnalyze add command -label "Syntax Errors" -command {
outputMsg "\nSyntax Errors -------------------------------------------------------------------------------------" blue
outputMsg "Syntax Errors are generated when an Analyzer option related to Semantic PMI, Graphic PMI, and
Validation Properties is selected.  The errors refer to specific sections, figures, or tables in
the relevant CAx-IF Recommended Practice.  Errors should be fixed so that the STEP file can
interoperate with other CAx software and conform to recommended practices.
See Websites > CAx Recommended Practices

Syntax errors are highlighted in red in the Status tab.  Other informative warnings are highlighted
in yellow.  Syntax errors that refer to section, figure, and table numbers might use numbers that
are in a newer version of a recommended practice that has not been publicly released.

Some syntax errors use abbreviations for STEP entities:
 GISU - geometric_item_specific_usage
 IIRU - identified_item_representation_usage

On the Summary worksheet in column A, most entities that have syntax errors are colored gray.  A
comment indicating that there are errors is also shown with a small red triangle in the upper right
corner of a cell in column A.

On an entity worksheet, most syntax errors are highlighted in red and have a cell comment with the
text of the syntax error that was shown in the Status tab.  Syntax errors are highlighted by *** in
the log file.

The Inverse Relationships option in the Analyzer section might be useful to debug Syntax Errors.

NOTE - Syntax Errors related to CAx-IF Recommended Practices are unrelated to errors detected with
the Syntax Checker.  See Help > Syntax Checker

See Help > User Guide (section 6.5)"
    .tnb select .tnb.status
  }

# NIST CAD model help
  $helpAnalyze add command -label "NIST CAD Models" -command {
outputMsg "\nNIST CAD Models -----------------------------------------------------------------------------------" blue
outputMsg "If a STEP file from a NIST CAD model (CTC/FTC/STC) is processed, then the PMI in the STEP file is
automatically checked against the expected PMI in the corresponding NIST test case.  The Semantic
PMI Coverage and Summary worksheets are color-coded by the expected PMI in each NIST test case.

The color-coding only works if the STEP file name can be recognized as having been generated from
one of the NIST CAD models.  For example, nist_ctc_02-some-text.stp would recognize the STEP file
as being generated from CTC 2.  nist_ftc_06-more-text.stp would be for FTC 6 and
nist_stc_10-whatever.stp would be for STC 10.

---------------------------------------------------------------------------------------------------
* Semantic PMI Summary *
This worksheet is color-coded by the Expected PMI annotations in a test case drawing.
- Green is an Exact match to an expected PMI annotation in the test case drawing
- Green (lighter shade) is an Exact match with Exceptions
- Cyan is a Partial match
- Yellow is a Possible match
- Red is No match

The following Exceptions are ignored when considering an Exact match:
- repetitive dimensions 'nX'
- different, missing, or unexpected dimensional tolerances in a Feature Control Frame (FCF)
- some datum features associated with geometric tolerances
- some modifiers in an FCF
- all around symbol

Some causes of Partial and Possible matches are, missing or wrong:
- diameter and radius symbols
- numeric values for dimensions and tolerances
- datum features and datum reference frames
- modifiers for dimensions, tolerance zones, and datum reference frames
- composite tolerances

On the Summary worksheet the column for Similar PMI and Exceptions shows the most closely matching
Expected PMI for Partial and Possible matches and the reason for an Exact match with Exceptions.

Trailing and leading zeros are ignored when matching a PMI annotation.  Matches also only consider
the current capabilities of PMI annotations in STEP AP242 and CAx-IF Recommended Practices.  For
example, PMI annotations for hole features including counterbore, countersink, and depth are not
supported.

---------------------------------------------------------------------------------------------------
* Graphic PMI Coverage Analysis *
This worksheet is color-coded by the expected number of PMI elements in a test case drawing.  The
expected results were determined by manually counting the number of PMI elements in each drawing.
Counting of some modifiers, e.g., maximum material condition, does not differentiate whether they
appear in the tolerance zone definition or datum reference frame.
- A green cell is a match to the expected number of PMI elements. (3/3)
- Yellow, orange, and yellow-green means that less were found than expected. (2/3)
- Red means that no instances of an expected PMI element were found. (0/3)
- Cyan means that more were found than expected. (4/3)
- Magenta means that some PMI elements were found when none were expected. (3/0)

---------------------------------------------------------------------------------------------------
From the Semantic PMI Summary results, color-coded percentages of Exact, Partial, Possible and
Missing matches is shown in a table below the Semantic PMI Coverage Analysis.  Exceptions are
counted as an Exact match and do not affect the percentage, except one or two points are deducted
when the percentage would be 100.

The Total PMI on which the percentages are based on is also shown.  Coverage Analysis is only based
on individual PMI elements.  The Semantic PMI Summary is based on the entire PMI feature control
frame and provides a better understanding of the PMI.  The Coverage Analysis might show that there
is an exact match for all of the PMI elements, however, the Semantic PMI Summary might show less
than exact matches.

---------------------------------------------------------------------------------------------------
* Missing PMI *
Missing PMI refers to semantic PMI and does not imply that the corresponding graphic annotation is
missing.  Missing PMI annotations on the Summary worksheet or PMI elements on the Coverage
worksheet might mean that the CAD system or translator:
- PMI annotation defined in a NIST test case is not in the CAD model
- did not follow CAx-IF Recommended Practices for PMI (See Websites > CAx Recommended Practices)
- has not implemented exporting a PMI element to a STEP file
- mapped an internal PMI element to the wrong STEP PMI element

* User-defined Expected PMI *
A user-defined file of expected PMI can also be used.  The file must be named SFA-EPMI-yourstepfilename.xlsx
Contact the developer for more information.  This feature is not documented in the User Guide.

NOTE - Some of the NIST test cases have complex PMI annotations that are not commonly used.  There
might be ambiguities in counting the number of PMI elements.

See Help > User Guide (section 6.6)
See Examples > NIST CAD Models
See Examples > Spreadsheets - Semantic PMI"
    .tnb select .tnb.status
  }
  $Help add separator

  $Help add command -label "Syntax Checker" -command {
outputMsg "\nSyntax Checker ------------------------------------------------------------------------------------" blue
outputMsg "The Syntax Checker checks for basic syntax errors and warnings in the STEP file related to missing
or extra attributes, incompatible and unresolved entity references, select value types, illegal and
unexpected characters, and other problems with entity attributes.  Some errors might prevent this
software and others from processing a STEP file.  Characters that are identified as illegal or
unexpected might not be shown in a spreadsheet or in the viewer.  See Help > Text Strings and Numbers

Entities in the STEP file that are not in the STEP AP schema are reported as ignored entities. They
will not appear in the spreadsheet.  Attributes on other entities that refer to ignored entities
will be blank.  Analyzer reports and the Viewer might be affected.

If errors and warnings are reported, the number in parentheses is the line number in the STEP file
where the error or warning was detected.  There should not be any of these types of syntax errors
in a STEP file.  Errors should be fixed to ensure that the STEP file conforms to the STEP schema
and can interoperate with other software.

There are other validation rules defined by STEP schemas (where, uniqueness, and global rules,
inverses, derived attributes, and aggregates) that are NOT checked.  Conforming to the validation
rules is also important for interoperability with STEP files.  See Websites > STEP

---------------------------------------------------------------------------------------------------
The Syntax Checker can be run with function key F8 or when a Spreadsheet or View is generated.  All
entities are always checked and is not affected by the Maximum Rows setting.  The Status tab might
be grayed out when the Syntax Checker is running.

Syntax checker results appear in the Status tab.  If the Log File option is selected, the results
are also written to a log file (myfile-sfa-err.log).  The syntax checker errors and warnings are
not reported in the spreadsheet.

The Syntax Checker can also be run from the command-line version with the command-line argument
'syntax'.  For example: sfa-cl.exe myfile.stp syntax

The Syntax Checker works with any supported schema.  See Help > Supported STEP APs and
Help > Large STEP files

NOTE - Syntax Checker errors and warnings are unrelated to those detected when CAx-IF Recommended
Practices are checked with one of the Analyzer options.  See Help > Analyzer > Syntax Errors"
    .tnb select .tnb.status
  }

  $Help add command -label "Bill of Materials" -command {
outputMsg "\nBill of Materials ---------------------------------------------------------------------------------" blue
outputMsg "Select BOM on the Generate tab to generate a Bill of Materials.  The next_assembly_usage_occurrence
entity shows the assembly and component names for the relating and related products in an assembly.
If there are no next_assembly_usage_occurrence entities, then the Bill of Materials (BOM) cannot
be generated.

The BOM worksheet (third worksheet) lists the quantities of parts and assemblies in two tables.
Assemblies also show their components which can be parts or other assemblies.  A STEP file might
not contain all the necessary information to generate a complete BOM.  Parts do not have to be
contained in an assembly, therefore some BOMs will not have a list of assemblies and some parts
might not be listed as a component of an assembly.  See Examples > Bill of Materials

Generate the Analyzer report for Validation Properties to see possible properties associated with
Parts.

Bill of Materials are not documented in the User Guide.  See Examples > Bill of Materials"
    .tnb select .tnb.status
  }
  $Help add separator

# open Function Keys help
  $Help add command -label "Function Keys" -command {
outputMsg "\nFunction Keys -------------------------------------------------------------------------------------" blue
outputMsg "Function keys can be used as shortcuts for several commands:

F1 - Generate Spreadsheet and/or run the Viewer with the current or last STEP file
F2 - Open current or last Spreadsheet
F3 - Open current or last Viewer file in web browser
F4 - Open Log file
F5 - Open STEP file in a text editor  (See Help > Open STEP File in App)
Shift-F5 - Open STEP file directory

F6 - Generate Spreadsheets and/or run the Viewer with the current or last set of multiple STEP files
F7 - Open current or last File Summary Spreadsheet generated from a set of multiple STEP files

F8 - Run the Syntax Checker (See Help > Syntax Checker)

F9  - Decrease this font size (also ctrl -)
F10 - Increase this font size (also ctrl +)

F12 - Open Viewer file in text editor

For F1, F2, F3, F6, and F7 the last STEP file, Spreadsheet, and Viewer file are remembered between
sessions.  In other words, F1 can process the last STEP file from a previous session without having
to select the file.  F2 and F3 function similarly for Spreadsheets and the Viewer."
    .tnb select .tnb.status
  }

  $Help add command -label "Supported STEP APs" -command {
outputMsg "\nSupported STEP APs --------------------------------------------------------------------------------" blue
outputMsg "These STEP Application Protocols (AP) and other schemas are supported for generating spreadsheets.
The Viewer works with STEP AP242, AP203, AP214, AP238, and AP209.  See Websites > STEP
AP238 STEP-NC files (.stpnc) are supported by renaming the file extension to '.stp'.

The name of the AP is on the FILE_SCHEMA entity in the HEADER section of a STEP file.  The 'e1'
notation below, after an AP number, refers to an older Edition of that AP.  Some APs have multiple
editions with the same name.  AP242 editions 1-4 were released in 2014, 2020, 2022, and 2025.
See Websites > AP242\n"

    set schemas {}
    set ifcschemas {}
    foreach match [lsort [glob -nocomplain -directory $ifcsvrDir *.rose]] {
      set schema [string toupper [file rootname [file tail $match]]]
      if {[string first "HEADER_SECTION" $schema] == -1 && [string first "KEYSTONE" $schema] == -1 && \
          [string first "ENGINEERING_MIM_LF-OLD" $schema] == -1 && [string range $schema end-2 end] != "MIM"} {
        if {[info exists stepAPs($schema)] && $schema != "STRUCTURAL_FRAME_SCHEMA"} {
          if {[string first "CONFIGURATION" $schema] != 0} {
            set str $stepAPs($schema)
            if {[string first "e1" $str] == -1} {append str "  "}
            lappend schemas "$str - $schema"
          } else {
            lappend schemas $schema
          }
        } elseif {[string first "AP2" $schema] == 0} {
          lappend schemas "[string range $schema 0 4]   - $schema"
        } elseif {[string first "IFC" $schema] == -1} {
          lappend schemas $schema
        } elseif {$schema == "IFC2X3" || [string first "IFC4" $schema] == 0 || [string first "IFC5" $schema] == 0} {
          lappend ifcschemas [string range $schema 3 end]
        }
      }
    }
    if {[llength $ifcschemas] > 0} {lappend schemas "IFC ($ifcschemas)"}

    if {[llength $schemas] <= 1} {
      errorMsg "No Supported STEP APs were found."
      if {[llength $schemas] == 1} {errorMsg "- Manually uninstall the existing IFCsvrR300 ActiveX Component 'App'."}
      errorMsg "- Restart this software to install the new IFCsvr toolkit."
    }

    set n 0
    foreach item [lsort $schemas] {
      set c1 [string first "-" $item]
      if {$c1 == -1} {
        if {$n == 0} {
          incr n
          outputMsg "\nOther Schemas"
        }
        set txt [string toupper $item]
        if {$txt == "CUTTING_TOOL_SCHEMA_ARM"} {append txt " (ISO 13399)"}
        if {[string first "ISO13584_25" $txt] == 0} {append txt " (Supplier library)"}
        if {[string first "ISO13584_42" $txt] == 0} {append txt " (Parts library)"}
        if {$txt == "STRUCTURAL_FRAME_SCHEMA"} {append txt " (CIS/2)"}
        outputMsg "  $txt"
      } else {
        set txt "[string range $item 0 $c1][string toupper [string range $item $c1+1 end]]"
        outputMsg "  $txt"
      }
    }
    .tnb select .tnb.status
  }

  $Help add command -label "Text Strings and Numbers" -command {
outputMsg "\nText Strings and Numbers --------------------------------------------------------------------------" blue
outputMsg "Text strings in STEP files might use non-English characters or symbols.  Some examples are accented
characters in European languages (for example é), and Asian languages that use different characters
sets such as Japanese or Cyrillic.  Text strings with non-English characters or symbols are usually
on descriptive measure or product related entities with name, description, or id attributes.

According to ISO 10303 Part 21 section 6.4.3, Unicode can be used for non-English characters and
symbols with the control directives \\X2\\ and \\X0\\.  For example, \\X2\\00E9\\X0\\ is used for the
accented character é.  Some CAD software does not support these control directives when exporting
or importing a STEP file.

---------------------------------------------------------------------------------------------------
Spreadsheet - Use the option on the More tab to support non-English characters using the \\X2\\
control directive.  In some cases the option will be automatically selected based on the file size
or schema.  There is a warning message if \\X2\\ is detected in the STEP file and the option is not
selected.  In this case the \\X2\\ characters are ignored and will be missing in the spreadsheet.
Non-English characters that do not use the control directives might be missing in the spreadsheet.
Control directives are supported only if Excel is installed.

Unicode characters for GD&T symbols are used by Equivalent Unicode Strings reported on the
descriptive_representation_item worksheet and worksheets for semantic and graphic PMI where there
is an associated PMI validation property.  Equivalent Unicode Strings are not documented in the
User Guide.  See the Recommended Practice for PMI Unicode String Specification.

Unicode characters using \\X2\\ control directives are not processed for attribute strings on
Geometry entities and some Presentation entities.  The \\X\\ and \\S\\ control directives are supported
by default although they are no longer implemented in STEP files.

---------------------------------------------------------------------------------------------------
Viewer - All control directives are supported for part and assembly names.  Non-English characters
that do not use the control directives might be shown with the wrong characters.

Support for non-English characters that do not use the control directives can be improved by
converting the encoding of the STEP file to UTF-8 with the Notepad++ text editor or other software.
Regardless, some non-English characters might cause a crash or prevent the viewer from running.
See Help > Crash Recovery

The Syntax Checker identifies non-English characters as 'illegal characters'.  You should test your
CAD software to see if it supports non-English characters or control directives.

---------------------------------------------------------------------------------------------------
Numbers in a STEP file use a period '.' as the decimal separator.  Some non-English language
versions of Excel use a comma ',' as a decimal separator.  This might cause some real numbers to be
formatted as a date in a spreadsheet.  For example, the number 1.5 might appear as 1-Mai.

To check if the formatting is a problem, process the STEP file nist_ctc_05.stp included with the
SFA zip file and select the Geometry Entity Type category.  Check the 'radius' attribute on the
resulting 'circle' worksheet.

To change the formatting in Excel, go to the Excel File menu > Options > Advanced.  Uncheck
'Use system separators' and change 'Decimal separator' to a period . and 'Thousands separator' to a
comma ,

This change applies to ALL Excel spreadsheets on your computer.  Change the separators back to
their original values when finished.  You can always check the STEP file to see the actual value of
the number.

See Help > User Guide (section 5.5.2)"
    .tnb select .tnb.status
  }

# open STEP files help
  $Help add command -label "Open STEP File in App" -command {
outputMsg "\nOpen STEP File in App -----------------------------------------------------------------------------" blue
outputMsg "STEP files can be opened in other apps.  If apps are installed in their default directory, then the
pull-down menu on the Generate tab will contain apps including STEP viewers and browsers.

The 'Tree View (for debugging)' option rearranges and indents the entities to show the hierarchy of
information in a STEP file.  The 'tree view' file (myfile-sfa.txt) is written to the same directory
as the STEP file or to the same user-defined directory specified in the More tab.  It is useful for
debugging STEP files but is not recommended for large STEP files.

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
Entity Types and options on the Generate tab.

If the reports for Semantic or Graphic PMI are selected, then Coverage Analysis worksheets are also
generated.

In some rare cases an error will be reported with an entity when processing multiple files that is
not an error when processing it as a single file.  Reporting the error is a bug.

See Help > User Guide (section 8)
See Examples > PMI Coverage Analysis"
    .tnb select .tnb.status
  }

# large files help
  $Help add command -label "Large STEP Files" -command {
outputMsg "\nLarge STEP Files ----------------------------------------------------------------------------------" blue
outputMsg "The largest STEP file that can be processed for a Spreadsheet or the Syntax Checker is
approximately 430 MB.  Processing a larger STEP file might cause a crash.  A popup dialog might
appear that says 'unable to realloc xxx bytes'.  See Help > Crash Recovery

Some workarounds are available to (1) prevent a crash with a smaller file that can be processed,
(2) to reduce the processing time, or (3) to reduce the size of the spreadsheet:
- Deselect Entity Types that might not need to be processed such as Geometry and Coordinates
- Use a User-Defined List of only the required entity types
- Use a smaller value for the Maximum Rows
- Deselect Analyzer options and Inverse Relationships

The Status tab might be grayed out when a large STEP file is being read.

---------------------------------------------------------------------------------------------------
To use only the Viewer with a large STEP file, select View and Part Only.  A 1.5 GB STEP file has
been successfully tested this way in the Viewer."
    .tnb select .tnb.status
  }

  $Help add command -label "Crash Recovery" -command {
outputMsg "\nCrash Recovery ------------------------------------------------------------------------------------" blue
outputMsg "Sometimes this software crashes after a STEP file has been successfully opened and the processing
of entities has started.  Popup dialogs might appear that say 'Runtime Error!' or 'ActiveState
Basekit has stopped working'.

A crash is most likely due to syntax errors in the STEP file, a very large STEP file, the options
selected for the software, or due to limitations of the toolkit used to read STEP files.  Run the
Syntax Checker with function key F8 or the option on the Generate tab to check for errors with
entities that might have caused the crash.  See Help > Syntax Checker

The software keeps track of the last entity processed when it crashed.  The list of bad entity
types is stored in myfile-skip.dat.  Simply restart this software and use F1 to process the last
STEP file or use F6 if processing multiple files.  The entities listed in myfile-skip.dat that
caused the crash will be skipped.

When the STEP file is processed again, the list of specific entities that are not processed is
reported.  If syntax errors related to the bad entities are corrected, then delete or edit the
*-skip.dat file so that the corrected entities are processed.

There are two other workarounds.  On the Generate tab deselect the Entity Type that caused the
error.  However, this will prevent processing of other entities that do not cause a crash.  A
User-Defined List can be used to process only the required entities for a spreadsheet.

For errors similar to 'unable to realloc xxx bytes', see Help > Large STEP Files

See Help > User Guide (section 2.4)

---------------------------------------------------------------------------------------------------
If the software crashes the first time you run it, there might be a problem with the installation
of the IFCsvr toolkit.  First uninstall the IFCsvr toolkit.  Then run SFA as Administrator and when
prompted, install the IFCsvr toolkit for Everyone, not Just Me.  Subsequently, SFA does not have to
be run as Administrator."
    .tnb select .tnb.status
  }

  $Help add separator
  $Help add command -label "Disclaimers" -command {
outputMsg "\nDisclaimers ---------------------------------------------------------------------------------------" blue
outputMsg "Please see Help > NIST Disclaimer for the Software Disclaimer.

The Examples menu provides links to several sources of STEP files.  This software and others might
indicate that there are errors in some of the STEP files.  NIST assumes no responsibility
whatsoever for the use of the STEP files by other parties, and makes no guarantees, expressed or
implied, about their quality, reliability, or any other characteristic.

Any mention of commercial products or references to web pages is for information purposes only; it
does not imply recommendation or endorsement by NIST.  For any of the web links, NIST does not
necessarily endorse the views expressed, or concur with the facts presented on those web sites.

This software uses IFCsvr, Microsoft Excel, and software based on Open Cascade that are covered by
their own Software License Agreements.

If you are using this software in your own application, please explicitly acknowledge NIST as the
source of the software."
  .tnb select .tnb.status
  }

  $Help add command -label "NIST Disclaimer" -command {openURL https://www.nist.gov/disclaimer}
  $Help add command -label "About" -command {
    outputMsg "\nSTEP File Analyzer and Viewer ---------------------------------------------------------------------" blue
    set sysvar "Version: [getVersion] ([string trim [clock format $progtime -format "%e %b %Y"]])"
    catch {append sysvar ", IFCsvr [registry get $ifcsvrVer {DisplayVersion}]"}
    catch {append sysvar ", stp2x3d ([string trim [clock format [file mtime [file join $mytemp stp2x3d-part.exe]] -format "%e %b %Y"]])"}
    append sysvar "\nFiles processed: $filesProcessed"
    outputMsg $sysvar
    outputMsg "\nThis software was first released in April 2012 and developed in the NIST Engineering Laboratory.

Credits
- Reading and parsing STEP files
   IFCsvr ActiveX Component, Copyright \u00A9 1999, 2005 SECOM Co., Ltd. All Rights Reserved
   IFCsvr has been modified by NIST to include STEP schemas
   The license agreement is in C:\\Program Files (x86)\\IFCsvrR300\\doc
- Viewer for b-rep and tessellated part geometry
   STEP to X3D Translator (stp2x3d) developed by Soonjo Kwon, former NIST Associate
   See Websites > STEP
- Some Tcl code is based on CAWT https://www.tcl3d.org/cawt/

See Help > Disclaimers and NIST Disclaimer"

    set winver ""
    if {[catch {
      set winver [registry get {HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion} {ProductName}]
    } emsg]} {
      set winver "Windows [expr {int($tcl_platform(osVersion))}]"
    }
    catch {
      set build [string range [lindex [split [::twapi::get_os_description] " "] end] 0 end-1]
      if {[string is integer $build] && $build >=22000} {regsub "10" $winver "11" winver}
    }
    if {[string first "Server" $winver] != -1 || $tcl_platform(osVersion) < 6.1} {
      outputMsg "\n$winver is not supported." red
    } elseif {$tcl_platform(osVersion) < 10.0} {
      outputMsg "\nThis software is no longer tested in $winver." red
    }

# debug
    if {$opt(xlMaxRows) == 100003} {
      outputMsg " "
      outputMsg "SFA variables" red
      catch {outputMsg " Drive [file nativename $drive]"}
      catch {outputMsg " Home  $myhome"}
      catch {outputMsg " Docs  $mydocs"}
      catch {outputMsg " Desk  $mydesk"}
      catch {outputMsg " Menu  $mymenu"}
      catch {outputMsg " Temp  $mytemp"}
      outputMsg " pf32  $pf32"
      if {$pf64 != ""} {outputMsg " pf64  $pf64"}
      outputMsg " S [winfo screenwidth  .]x[winfo screenheight  .], M [winfo reqwidth .]x[expr {int([winfo reqheight .]*1.05)}]"
      outputMsg " $winver"
      catch {outputMsg " scriptName [file nativename $scriptName]"}

      outputMsg "Registry values" red
      catch {outputMsg " Personal  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Personal}]"}
      catch {outputMsg " Desktop   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]"}
      catch {outputMsg " Programs  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Programs}]"}
      catch {outputMsg " AppData   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local AppData}]"}

      outputMsg "Environment variables" red
      foreach id [lsort [array names env]] {
        foreach id1 [list HOME USER APP ROSE_SCHEMAS] {if {[string first $id1 $id] == 0} {catch {outputMsg " $id  $env($id)"; break}}}
      }
    }
    .tnb select .tnb.status
  }

# examples menu
  $Examples add command -label "Viewer"                   -command {openURL https://pages.nist.gov/CAD-PMI-Testing/}
  $Examples add command -label "Spreadsheets - AP242 PMI" -command {openURL https://www.nist.gov/document/sfa-semantic-pmi-spreadsheet}
  $Examples add command -label "- AP203 PMI"              -command {openURL https://www.nist.gov/document/sfa-spreadsheet}
  $Examples add command -label "- PMI Coverage Analysis"  -command {openURL https://www.nist.gov/document/sfa-multiple-files-spreadsheet}
  $Examples add command -label "- Bill of Materials"      -command {openURL https://www.nist.gov/document/sfa-bill-materials-spreadsheet}
  $Examples add separator
  $Examples add command -label "Sample STEP Files (zip)" -command {openURL https://www.nist.gov/document/nist-pmi-step-files}
  $Examples add command -label "NIST CAD Models"         -command {openURL https://www.nist.gov/ctl/smart-connected-systems-division/smart-connected-manufacturing-systems-group/mbe-pmi-0}
}

#-------------------------------------------------------------------------------
# Websites menu
proc guiWebsitesMenu {} {
  global Websites

  $Websites add command -label "STEP File Analyzer and Viewer"              -command {openURL https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer}
  $Websites add separator
  $Websites add command -label "CAx Interoperability Forum (CAx-IF)" -command {openURL https://www.mbx-if.org/home/cax/}
  $Websites add command -label "CAx Recommended Practices"           -command {openURL https://www.mbx-if.org/home/cax/recpractices/}
  $Websites add command -label "CAD Implementations"                 -command {openURL https://www.mbx-if.org/home/cax/implementation-coverage/}
  $Websites add command -label "PDM-IF"                              -command {openURL https://www.mbx-if.org/home/pdm/}

  $Websites add separator
  $Websites add cascade -label "AP242" -menu $Websites.0
  set Websites0 [menu $Websites.0 -tearoff 1]
  $Websites0 add command -label "AP242 Project"           -command {openURL http://www.ap242.org}
  $Websites0 add command -label "AP203 vs AP214 vs AP242" -command {openURL https://www.capvidia.com/blog/best-step-file-to-use-ap203-vs-ap214-vs-ap242}
  $Websites0 add command -label "Benchmark Testing"       -command {openURL http://www.asd-ssg.org/step-ap242-benchmark.html}
  $Websites0 add command -label "Domain Model XML"        -command {openURL https://www.mbx-if.org/home/pdm/recpractices/}

  $Websites0 add separator
  $Websites0 add command -label "ISO 10303-242"           -command {openURL https://www.iso.org/standard/84300.html}
  $Websites0 add command -label "STEP in 3D PDF"          -command {openURL https://www.iso.org/standard/77686.html}
  $Websites0 add command -label "STEP Geometry Services"  -command {openURL https://www.iso.org/standard/84820.html}

  $Websites add cascade -label "STEP" -menu $Websites.2
  set Websites2 [menu $Websites.2 -tearoff 1]
  $Websites2 add command -label "STEP File Viewers"      -command {openURL https://www.mbx-if.org/home/mbx/resources/}
  $Websites2 add command -label "STEP to X3D Translator" -command {openURL https://www.nist.gov/services-resources/software/step-x3d-translator}
  $Websites2 add command -label "STEP File Format"  -command {openURL https://www.loc.gov/preservation/digital/formats/fdd/fdd000448.shtml}
  $Websites2 add command -label "STEP ISO 10303-21" -command {openURL https://en.wikipedia.org/wiki/ISO_10303-21}

  $Websites2 add separator
  $Websites2 add command -label "AP209 FEA"     -command {openURL http://www.ap209.org}
  $Websites2 add command -label "AP238 STEP-NC" -command {openURL https://www.ap238.org}
  $Websites2 add command -label "AP239 PLCS"    -command {openURL http://www.ap239.org}
  $Websites2 add separator
  $Websites2 add command -label "Learning EXPRESS"               -command {openURL https://www.expresslang.org/learn/}
  $Websites2 add command -label "easyEXPRESS"                    -command {openURL https://www.nist.gov/news-events/news/2023/12/nist-releases-easyexpress-its-first-official-visual-studio-code-extension}
  $Websites2 add command -label "EXPRESS ISO 10303-11"           -command {openURL https://www.loc.gov/preservation/digital/formats/fdd/fdd000449.shtml}
  $Websites2 add command -label "EXPRESS data modeling language" -command {openURL https://en.wikipedia.org/wiki/EXPRESS_(data_modeling_language)}
  $Websites2 add command -label "EXPRESS Schemas"                -command {openURL https://www.mbx-if.org/home/mbx/resources/express-schemas/}

  $Websites add cascade -label "Organizations" -menu $Websites.4
  set Websites4 [menu $Websites.4 -tearoff 1]
  $Websites4 add command -label "STEP at NIST"           -command {openURL https://www.nist.gov/ctl/smart-connected-systems-division/smart-connected-manufacturing-systems-group/step-nist}
  $Websites4 add command -label "PDES, Inc. (U.S.)"      -command {openURL https://pdesinc.org/}
  $Websites4 add command -label "prostep ivip (Germany)" -command {openURL https://www.prostep.org/en/projects/mbx-interoperability-forum-mbx-if}
  $Websites4 add command -label "AFNeT (France)"         -command {openURL https://atlas.afnet.fr/en/domaines/plm/}
  $Websites4 add command -label "KStep (Korea)"          -command {openURL https://www.kstep.or.kr}
  $Websites4 add separator
  $Websites4 add command -label "MBx Interoperability Forum (MBx-IF)"       -command {openURL https://www.mbx-if.org/home/}
  $Websites4 add command -label "LOTAR - LOng Term Archiving and Retrieval" -command {openURL https://lotar-international.org}
  $Websites4 add command -label "ISO/TC 184/SC 4 - Industrial Data"         -command {openURL https://committee.iso.org/home/tc184sc4}
  $Websites4 add command -label "JT-IF"                                     -command {openURL https://www.prostep.org/en/projects/jt-project-groups-jt-wf-jt-if-jt-bm}
}

#-------------------------------------------------------------------------------
# crash recovery dialog
proc showCrashRecovery {} {

set txt "Sometimes this software crashes AFTER a file has been successfully opened and the processing of entities has started.

A crash is most likely due to syntax errors in the STEP file or sometimes due to limitations of the toolkit used to read STEP files.  Run the Syntax Checker with function key F8 to check for errors with entities that might have caused the crash.  See Help > Syntax Checker

Processing very large STEP files might also cause a crash.  See Help > Large STEP Files

More details about recovering from a crash are explained in Help > Crash Recovery.

If the software crashes the first time you run it, there might be a problem with the installation of the IFCsvr toolkit.  First uninstall the IFCsvr toolkit.  Then run SFA as Administrator and when prompted, install the IFCsvr toolkit for Everyone, not Just Me.  Subsequently, SFA does not have to be run as Administrator."

  tk_messageBox -type ok -icon error -title "What to do if the STEP File Analyzer and Viewer crashes?" -message $txt
}

#-------------------------------------------------------------------------------
proc guiToolTip {ttmsg tt {name ""}} {
  global ap203all ap214all ap242all ap242only entCategory inverses

# two different types of subsets of entities
  if {$tt == "stepCOMM" || $tt == "stepPRES" || $tt == "stepREPR" || $tt == "stepKINE" || $tt == "inverses"} {
    set ents {}
    set prefix {}

    if {$tt != "inverses"} {
      set entities $entCategory($tt)
    } else {
      set entities {}
      foreach item [lsort $inverses] {
        set ent [lindex $item 0]
        if {[lsearch $ap203all $ent] != -1 || [lsearch $ap214all $ent] != -1 || [lsearch $ap242all $ent] != -1} {lappend entities $ent}
      }
    }

    set n 0
    foreach ent $entities {
      set c1 [string first "_" $ent]
      if {$c1 == -1} {
        lappend ents $ent
      } else {
        set c2 [string first "_" [string range $ent $c1+1 end]]
        if {$c2 != -1} {
          set pre [string range $ent 0 $c1+$c2]
        } else {
          set pre $ent
        }
        if {[lsearch $prefix $pre] == -1} {
          incr n
          set ok 0
          if {$tt != "stepCOMM"} {
            set ok 1
          } elseif {[expr {$n%3}] != 0} {
            set ok 1
          }
          lappend prefix $pre
          if {$ok} {lappend ents $ent}
        }
      }
    }
    append ttmsg "  This is a subset of [llength $ents] entities."

# another type of subset
  } elseif {$tt == "stepAP242" || $tt == "stepQUAL" || $tt == "stepQUAN" || $tt == "stepGEOM" || $tt == "stepOTHR"} {
    set ents {}
    set prefix {}
    set n 0
    foreach ent $entCategory($tt) {
      if {$tt != "stepCOMM" || [lsearch $ap242all $ent] != -1} {
        set c1 [string first "_" $ent]
        if {$c1 != -1} {
          set pre [string range $ent 0 3]
          if {[lsearch $prefix $pre] == -1} {
            incr n
            set ok 0
            if {$tt != "stepAP242" && $tt != "stepOTHR"} {
              set ok 1
            } elseif {[expr {$n%3}] != 0} {
              set ok 1
            }
            lappend prefix $pre
            if {$ok} {lappend ents $ent}
          }
        }
      }
    }
    append ttmsg "  This is a subset of [llength $ents] entities."

# all entities, no subset
  } else {
    set ents $entCategory($tt)
  }

  set space 2
  set ttlim 120
  if {$tt == "stepQUAL" || $tt == "stepCONS" || $tt == "stepOTHR"} {set ttlim 110}
  append ttmsg "\n\n"

  foreach type {ap203 ap242} {
    set ttlen 0
    foreach item $ents {
      set ok 0
      set ent $item
      switch -- $type {
        ap203 {if {[lsearch $ap242only(all) $ent] == -1} {set ok 1}}
        ap242 {if {[lsearch $ap242only(all) $ent] != -1} {set ok 1}}
      }
      if {$tt == "inverses"} {
        switch -- $type {
          ap203 {set ok 1}
          ap242 {set ok 0}
        }
      }
      if {$ok} {

# add superscript for AP242 edition
        if {$type == "ap242" && $tt != "stepKINE"} {
          if {[lsearch $ap242only(e1) $ent] != -1} {append ent "\u00B9"}
          if {[lsearch $ap242only(e2) $ent] != -1} {append ent "\u00B2"}
          if {[lsearch $ap242only(e3) $ent] != -1} {append ent "\u00B3"}
          if {[lsearch $ap242only(e4) $ent] != -1} {append ent "\u2074"}
          catch {if {[lsearch $ap242only(e5) $ent] != -1} {append ent "\u2075"}}
          catch {if {[lsearch $ap242only(e6) $ent] != -1} {append ent "\u2076"}}
        }
        incr ttlen [expr {[string length $ent]+$space}]
        if {$ttlen <= $ttlim} {
          append ttmsg "$ent[string repeat " " $space]"
        } else {
          if {[string index $ttmsg end] != "\n"} {set ttmsg "[string range $ttmsg 0 end-$space]\n$ent[string repeat " " $space]"}
          set ttlen [expr {[string length $ent]+$space}]
        }
      }
    }
    if {$type == "ap203" && $tt != "stepCOMM" && $tt != "stepAP242" && $tt != "stepQUAL" && $tt != "stepCONS" && $tt != "inverses"} {
      if {$tt == "stepCPNT"} {append ttmsg "is supported in most STEP APs."}
      if {$tt != "stepOTHR"} {append ttmsg "\n\nThe following entities are supported only in AP242."}
      if {$tt != "stepKINE" && $tt != "stepCPNT" && $tt != "stepOTHR"} {append ttmsg "  Superscript indicates edition of AP242."}
      if {$tt == "stepCPNT"} {append ttmsg "\nSuperscript indicates edition of AP242."}
      if {$tt != "stepOTHR"} {append ttmsg "\n\n"}
    }
  }
  return $ttmsg
}

#-------------------------------------------------------------------------------
proc getOpenPrograms {} {
  global env dispApps dispCmds dispCmd appNames appName
  global drive editorCmd developer myhome pf32 pf64 pflist

# Including any of the CAD viewers and software does not imply a recommendation or endorsement of them by NIST https://www.nist.gov/disclaimer
# For more STEP viewers, go to https://www.mbx-if.org/home/mbx/resources/

  regsub {\\} $pf32 "/" p32
  lappend pflist $p32
  if {$pf64 != "" && $pf64 != $pf32} {
    regsub {\\} $pf64 "/" p64
    lappend pflist $p64
  }
  set lastver 0

# Jotne EDM Model Checker
  if {$developer} {
    foreach match [glob -nocomplain -directory [file join $pf64 Jotne] -join EDMsdk6* bin edms.exe] {
      set ver [lindex [split $match "/"] 3]
      set dispApps($match) [string range $ver 0 [string last "." $ver]-1]
    }
  }

# STEP file viewers, use * when the directory or name has a number
  foreach pf $pflist {
    set applist [list \
      [list {*}[glob -nocomplain -directory [file join $pf "CAD Assistant"] -join CADAssistant.exe] "CAD Assistant"] \
      [list {*}[glob -nocomplain -directory [file join $pf "CAD Exchanger" bin] -join Exchanger.exe] "CAD Exchanger"] \
      [list {*}[glob -nocomplain -directory [file join $pf "C3D Labs"] -join "C3D Viewer" Bin c3dviewer.exe] "C3D Viewer"] \
      [list {*}[glob -nocomplain -directory [file join $pf "Common Files"] -join "eDrawings*" eDrawings.exe] "eDrawings Viewer"] \
      [list {*}[glob -nocomplain -directory [file join $pf "SOLIDWORKS Corp" "eDrawings X64 Edition"] -join eDrawings.exe] "eDrawings Pro"] \
      [list {*}[glob -nocomplain -directory [file join $pf "SOLIDWORKS Corp" eDrawings] -join eDrawings.exe] "eDrawings Pro"] \
      [list {*}[glob -nocomplain -directory [file join $pf "SOLIDWORKS Corp"] -join "eDrawings (*)" eDrawings.exe] "eDrawings Pro"] \
      [list {*}[glob -nocomplain -directory [file join $pf "Stratasys Direct Manufacturing"] -join "SolidView Pro RP *" bin SldView.exe] SolidView] \
      [list {*}[glob -nocomplain -directory [file join $pf "TransMagic Inc"] -join "TransMagic *" System code bin TransMagic.exe] TransMagic] \
      [list {*}[glob -nocomplain -directory [file join $pf Actify SpinFire] -join "*" SpinFire.exe] SpinFire] \
      [list {*}[glob -nocomplain -directory [file join $pf Asitus] -join asiExe.exe] "Analysis Situs"] \
      [list {*}[glob -nocomplain -directory [file join $pf CADSoftTools "CST CAD Navigator"] -join cstCadNavigator.exe] "CST CAD Navigator"] \
      [list {*}[glob -nocomplain -directory [file join $pf CADSoftTools] -join "ABViewer*" ABViewer.exe] ABViewer] \
      [list {*}[glob -nocomplain -directory [file join $pf Fougue] -join "Mayo" mayo.exe] Mayo] \
      [list {*}[glob -nocomplain -directory [file join $pf gCAD3D] -join gCAD3D.bat] gCAD3D] \
      [list {*}[glob -nocomplain -directory [file join $pf Glovius Glovius] -join glovius.exe] Glovius] \
      [list {*}[glob -nocomplain -directory [file join $pf Gstarsoft 3DFastView] -join "V*" 3DFastView.exe] 3DFastView] \
      [list {*}[glob -nocomplain -directory [file join $pf IFCBrowser] -join IfcQuickBrowser.exe] IfcQuickBrowser] \
      [list {*}[glob -nocomplain -directory [file join $pf Kisters 3DViewStation] -join 3DViewStation.exe] 3DViewStation] \
      [list {*}[glob -nocomplain -directory [file join $pf Kubotek] -join "KDisplayView*" KDisplayView.exe] "K-Display View"] \
      [list {*}[glob -nocomplain -directory [file join $pf Kubotek] -join "Spectrum*" Spectrum.exe] Spectrum] \
      [list {*}[glob -nocomplain -directory [file join $pf ODA] -join "Open STEP Viewer*" OpenSTEPViewer.exe] "Open STEP Viewer"] \
      [list {*}[glob -nocomplain -directory [file join $pf STPViewer] -join STPViewer.exe] "STP Viewer"] \
      [list {*}[glob -nocomplain -directory [file join $pf ZWSOFT] -join "CADbro *" CADbro.exe] CADbro] \
      [list {*}[glob -nocomplain -directory [file join $pf] -join "3D-Tool V*" 3D-Tool.exe] 3D-Tool] \
      [list {*}[glob -nocomplain -directory [file join $pf] -join "Afanche3D*" "Afanche3D*.exe"] Afanche3D] \
      [list {*}[glob -nocomplain -directory [file join $pf] -join "VariCAD*" bin varicad-x64.exe] "VariCAD Viewer"] \
    ]

# add version number for some
    foreach app $applist {
      if {[llength $app] == 2} {
        set match [join [lindex $app 0]]
        if {$match != "" && ![info exists dispApps($match)]} {
          set dispApps($match) [lindex $app 1]
          set c1 [string first "eDrawings20" $match]
          if {$c1 != -1} {set dispApps($match) "[lindex $app 1] [string range $match $c1+9 $c1+12]"}
          set c1 [string first "Open STEP Viewer" $match]
          if {$c1 != -1} {
            set str [string range $match $c1+17 $c1+22]
            set str [string range $str 0 [string last "." $str]-1]
            set dispApps($match) "[lindex $app 1] $str"
          }
        }
      }
    }

# FreeCAD
    foreach app [list {*}[glob -nocomplain -directory [file join $pf] -join "FreeCAD *" bin FreeCAD.exe] FreeCAD] {
      set ver [lindex [split [file nativename $app] [file separator]] 2]
      set dispApps($app) $ver
    }
    foreach app [list {*}[glob -nocomplain -directory [file join $myhome AppData Local] -join "FreeCAD *" bin FreeCAD.exe] FreeCAD] {
      set ver [lindex [split [file nativename $app] [file separator]] 5]
      set dispApps($app) $ver
    }
  }

# others
  set b1 [file join $myhome AppData Local IDA-STEP ida-step.exe]
  if {[file exists $b1]} {set dispApps($b1) "IDA-STEP Viewer"}
  set b1 [file join $drive CCELabs EnSuite-View Bin EnSuite-View.exe]
  if {[file exists $b1]} {
    set dispApps($b1) "EnSuite-View"
  } else {
    set b1 [file join $drive CCE EnSuite-View Bin EnSuite-View.exe]
    if {[file exists $b1]} {set dispApps($b1) "EnSuite-View"}
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
    if {[file exists $app] || [string first "Default" $app] == 0 || [string first "Indent" $app] == 0} {
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

# remove old commands
  set oldcmd [list 3DJuump 3DPDFConverter 3DReviewer avwin BIMsight Magics QuickStep roamer \
                   apconformgui checkgui stepbrws stepcleangui STEPNCExplorer_x86 STEPNCExplorer stview]
  foreach cmd $dispCmds {if {[string first "notepad++.exe" $cmd] != -1} {lappend oldcmd "notepad"}}

  set ndcs {}
  foreach cmd $dispCmds {
    set ok 1
    foreach bcmd $oldcmd {
      append bcmd ".exe"
      if {[string first $bcmd $cmd] != -1} {set ok 0}
    }
    if {$ok} {lappend ndcs $cmd}
  }
  set dispCmds $ndcs

# set list of STEP viewer names, appNames
  set appNames {}
  set appName  ""
  foreach cmd $dispCmds {
    if {[info exists dispApps($cmd)]} {
      lappend appNames $dispApps($cmd)
    } else {
      set name [file rootname [file tail $cmd]]
      lappend appNames $name
      set dispApps($cmd) $name
    }
  }
  if {$dispCmd != ""} {
    if {[info exists dispApps($dispCmd)]} {set appName $dispApps($dispCmd)}
  }
}
#-------------------------------------------------------------------------------
# turn on/off values and enable/disable buttons depending on values
proc checkValues {} {
  global allNone appName appNames buttons developer edmWhereRules edmWriteToFile gen opt userEntityList useXL

  set butNormal {}
  set butDisabled {}

  if {[info exists buttons(appCombo)]} {
    set ic [lsearch $appNames $appName]
    if {$ic < 0} {set ic 0}
    catch {$buttons(appCombo) current $ic}

# Jotne EDM Model Checker
    if {$developer} {
      catch {
        if {[string first "EDMsdk" $appName] != -1} {
          pack $buttons(edmWhereRules) -side left -anchor w -padx {5 0}
          pack $buttons(edmWriteToFile) -side left -anchor w -padx {5 0}
        } else {
          pack forget $buttons(edmWriteToFile)
          pack forget $buttons(edmWhereRules)
        }
      }
    }

    catch {
      if {$appName == "Tree View (for debugging)"} {
        pack $buttons(indentGeometry) -side left -anchor w -padx {5 0}
        pack $buttons(indentStyledItem) -side left -anchor w -padx {5 0}
      } else {
        pack forget $buttons(indentGeometry)
        pack forget $buttons(indentStyledItem)
      }
    }
  }

# view
  if {$gen(View)} {
    lappend butNormal viewFEA viewPMI viewPart partOnly partCap partNoGroup x3dSave viewParallel viewCorrect viewNoPMI
    if {!$opt(viewFEA) && !$opt(viewPMI) && !$opt(viewPart)} {set opt(viewPart) 1}
    if {$developer} {lappend butNormal debugX3D}
  } else {
    set opt(x3dSave) 0
    lappend butDisabled viewFEA viewPMI viewPart partOnly partCap partNoGroup x3dSave viewParallel viewCorrect viewNoPMI
    lappend butDisabled gpmiColor0 gpmiColor1 gpmiColor2 gpmiColor3 labelPMIcolor
    lappend butDisabled partEdges partSketch partSupp partNormals labelPartQuality partQuality4 partQuality7 partQuality10 tessPartMesh
    lappend butDisabled feaBounds feaLoads feaLoadScale feaDisp feaDispNoTail
    if {$developer} {lappend butDisabled debugX3D; set opt(debugX3D) 0}
  }

# part only
  if {$opt(partOnly)} {
    set opt(viewPart) 1
    set opt(viewParallel) 0
    set opt(viewCorrect) 0
    set opt(viewNoPMI) 0
    set opt(xlFormat) "None"
    set gen(Excel) 0
    set gen(Excel1) 0
    set gen(CSV) 0
    set opt(tessPartOld) 0
    lappend butNormal genExcel
    lappend butDisabled partCap debugVP tessPartOld viewParallel viewCorrect viewNoPMI
  } else {
    lappend butNormal partCap debugVP tessPartOld viewParallel viewCorrect viewNoPMI
  }

  if {$gen(View) && $opt(viewPart)} {
    lappend butNormal brepAlt
    if {!$opt(partOnly)} {lappend butNormal tessPartOld}
  } else {
    lappend butDisabled brepAlt tessPartOld
    set opt(brepAlt) 0
    set opt(tessPartOld) 0
  }

  if {$opt(tessPartOld)} {
    set tessSolid 0
    set opt(viewTessPart) 1
    set opt(tessPartMesh) 1
  }

  if {!$gen(Excel)} {
    lappend butDisabled labelMaxRows xlHideLinks xlUnicode xlSort xlNoRound
    if {$opt(xlFormat) == "None"} {for {set i 0} {$i < 8} {incr i} {lappend butDisabled "maxrows$i"}}
    set opt(xlHideLinks) 0
  } else {
    lappend butNormal labelMaxRows xlHideLinks xlUnicode xlSort xlNoRound
    for {set i 0} {$i < 8} {incr i} {lappend butNormal "maxrows$i"}
  }
  if {$gen(Excel) && $gen(CSV)} {lappend butDisabled genExcel}

# configure generate button
  if {![info exists useXL]} {set useXL 1}
  set btext "Generate "
  if {$opt(xlFormat) == "Excel"} {
    append btext "Spreadsheet"
  } elseif {$opt(xlFormat) == "CSV"} {
    if {$gen(CSV) && $useXL} {append btext "Spreadsheet and "}
    append btext "CSV Files"
  } elseif {$gen(View) && $opt(xlFormat) == "None"} {
    append btext "View"
  }
  if {$gen(View) && $opt(xlFormat) != "None" && ($opt(viewPart) || $opt(viewFEA) || $opt(viewPMI))} {
    append btext " and View"
  }
  catch {$buttons(generate) configure -text $btext}

# no Excel
  if {!$useXL} {
    foreach item {BOM INVERSE PMIGRF PMISEM valProp} {set opt($item) 0}
    foreach item [array names opt] {
      if {[string first "step" $item] == 0} {lappend butNormal $item}
    }
    lappend butDisabled xlHideLinks xlUnicode xlSort xlNoRound BOM INVERSE PMIGRF PMISEM valProp genExcel
    lappend butNormal viewFEA viewPMI viewPart
    lappend butNormal allNone0 allNone1 stepUSER

# Excel
  } else {
    foreach item [array names opt] {
      if {[string first "step" $item] == 0} {lappend butNormal $item}
    }
    lappend butNormal xlHideLinks xlUnicode xlSort xlNoRound BOM INVERSE PMIGRF PMISEM valProp
    lappend butNormal viewFEA viewPMI viewPart
    lappend butNormal allNone0 allNone1 stepUSER
  }

# view only
  if {$opt(xlFormat) == "None"} {
    foreach item [array names opt] {
      if {[string first "step" $item] == 0} {lappend butDisabled $item}
    }
    lappend butDisabled PMIGRF PMISEM PMISEMDIM PMISEMDT PMISEMRND valProp stepUSER BOM INVERSE
    lappend butDisabled allNone0
    lappend butDisabled userentity userentityopen
    set userEntityList {}
    if {!$opt(viewFEA) && !$opt(viewPMI) && !$opt(viewPart)} {set opt(viewPart) 1}
  }

# part geometry
  if {$opt(viewPart)} {
    lappend butNormal partOnly partEdges partSketch partNormals partCap partNoGroup labelPartQuality partQuality4 partQuality7 partQuality10
    if {$opt(partOnly) && $opt(xlFormat) == "None"} {
      lappend butDisabled syntaxChecker viewFEA viewPMI partSupp
      foreach item {syntaxChecker viewFEA viewPMI partSupp partCap debugVP} {set opt($item) 0}
    } else {
      lappend butNormal syntaxChecker viewFEA viewPMI partSupp
    }
  } else {
    lappend butDisabled partEdges partSketch partNormals partCap partNoGroup labelPartQuality partQuality4 partQuality7 partQuality10
  }
  if {$opt(viewPart)} {
    lappend butNormal partSupp
  } else {
    lappend butDisabled partSupp
  }

# graphic PMI report
  if {$opt(PMIGRF)} {
    if {$opt(xlFormat) != "None"} {
      foreach b {stepPRES stepREPR stepSHAP} {
        set opt($b) 1
        lappend butDisabled $b
      }
    }
  } else {
    lappend butNormal stepPRES
    if {!$opt(valProp)} {lappend butNormal stepQUAN}
    if {!$opt(PMISEM)}  {lappend butNormal stepSHAP stepREPR}
  }

# validation properties
  if {$opt(valProp)} {
    foreach b {stepQUAN stepREPR stepSHAP} {
      set opt($b) 1
      lappend butDisabled $b
    }
  } elseif {!$opt(PMIGRF)} {
    lappend butNormal stepQUAN
  }

# graphic PMI view
  if {$opt(viewPMI)} {
    lappend butNormal gpmiColor0 gpmiColor1 gpmiColor2 gpmiColor3 labelPMIcolor viewNoPMI
  } else {
    lappend butDisabled gpmiColor0 gpmiColor1 gpmiColor2 gpmiColor3 labelPMIcolor viewNoPMI
  }

  if {$gen(View)} {
    lappend butNormal debugVP
  } else {
    lappend butDisabled debugVP
  }

# FEM view
  if {$opt(viewFEA)} {
    lappend butNormal feaBounds feaLoads feaDisp
    if {$opt(feaLoads)} {
      lappend butNormal feaLoadScale
    } else {
      lappend butDisabled feaLoadScale
    }
    if {$opt(feaDisp)} {
      lappend butNormal feaDispNoTail
    } else {
      lappend butDisabled feaDispNoTail
    }
  } else {
    lappend butDisabled feaBounds feaLoads feaLoadScale feaDisp feaDispNoTail
  }

# semantic PMI report
  if {$opt(PMISEM)} {
    foreach b {stepREPR stepSHAP stepTOLR stepQUAN} {
      set opt($b) 1
      lappend butDisabled $b
    }
    lappend butNormal PMISEMDIM PMISEMDT PMISEMRND
  } else {
    lappend butNormal stepREPR stepTOLR
    if {!$opt(PMIGRF)} {
      if {!$opt(valProp)} {lappend butNormal stepQUAN}
      lappend butNormal stepSHAP
    }
    lappend butDisabled PMISEMDIM PMISEMDT PMISEMRND
  }
  if {$opt(PMISEM) && $gen(Excel)} {
    lappend butNormal SHOWALLPMI
  } else {
    set opt(SHOWALLPMI) 0
    lappend butDisabled SHOWALLPMI
  }

# BOM
  if {$opt(BOM)} {
    set opt(stepCOMM) 1
    lappend butDisabled stepCOMM
  } else {
    lappend butNormal stepCOMM
  }

# common entities
  if {$opt(valProp) || $opt(PMISEM) || $opt(PMIGRF)} {
    set opt(stepCOMM) 1
    lappend butDisabled stepCOMM
  }

# not part geometry view
  if {!$opt(viewPart) && !$opt(PMISEM)} {lappend butNormal stepPRES}
  catch {if {!$opt(PMISEM)} {lappend butNormal stepPRES}}

# user-defined entity list
  if {$opt(stepUSER)} {
    lappend butNormal userentity userentityopen
  } else {
    lappend butDisabled userentity userentityopen
    set userEntityList {}
  }

  if {$developer} {
    if {$opt(INVERSE) && $gen(Excel)} {
      lappend butNormal DEBUGINV
    } else {
      lappend butDisabled DEBUGINV
      set opt(DEBUGINV) 0
    }
    if {$gen(Excel)} {
      lappend butNormal debugNOXL
      if {$opt(PMISEM) || $opt(PMIGRF) || $opt(valProp)} {
        lappend butNormal DEBUG1
        if {$opt(PMISEM) || $opt(PMIGRF)} {
          lappend butNormal debugAG
        } else {
          lappend butDisable debugAG
        }
      } else {
        lappend butDisabled DEBUG1 debugAG
      }
    } else {
      lappend butDisabled DEBUG1 debugAG debugNOXL
      set opt(DEBUG1) 0
      set opt(debugNOXL) 0
      set opt(debugAG) 0
    }
  }

# user-defined directory text entry and browse button
  if {$opt(writeDirType) == 0} {
    lappend butDisabled userdir userentry
  } elseif {$opt(writeDirType) == 2} {
    lappend butNormal userdir userentry
  }

# make sure there is some entity type to process
  set nopt 0
  foreach idx [lsort [array names opt]] {
    if {[string first "step" $idx] == 0 || $idx == "valProp" || $idx == "PMIGRF" || $idx == "PMISEM"} {
      incr nopt $opt($idx)
    }
  }
  if {$nopt == 0 && $opt(xlFormat) != "None"} {set opt(stepCOMM) 1}

# configure buttons
  if {[llength $butNormal]   > 0} {foreach but $butNormal   {catch {$buttons($but) configure -state normal}}}
  if {[llength $butDisabled] > 0} {foreach but $butDisabled {catch {$buttons($but) configure -state disabled}}}

# configure all, reset, 'all' view and analyzer buttons
  if {[info exists allNone]} {
    if {$allNone == 1} {
      foreach item [array names opt] {
        if {[string first "step" $item] == 0 && $item != "stepCOMM"} {
          if {$opt($item) == 1} {set allNone -1; break}
        }
        if {[string length $item] == 6 && ([string first "PMI" $item] == 0)} {
          if {$opt($item) == 1} {set allNone -1; break}
        }
      }
    } elseif {$allNone == 0} {
      foreach item [array names opt] {
        if {[string first "step" $item] == 0} {
          if {$item != "stepGEOM" && $item != "stepCPNT" && $item != "stepUSER"} {if {$opt($item) == 0} {set allNone -1}}
        }
      }
    }
  }
}
