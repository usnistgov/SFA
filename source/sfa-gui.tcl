# SFA version
proc getVersion {} {return 5.33}

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
      outputMsg "- User Guide is based on version 4.60"
      openUserGuide
    }
  }
  if {$sfaVersion < 5.10} {outputMsg "- Faster processing of B-rep and AP242 tessellated part geometry for the Viewer"}
  if {$sfaVersion < 5.0}  {outputMsg "- Renamed 'Options' and 'Spreadsheet' tabs to 'Generate' and 'More'"}
  if {$sfaVersion < 5.20} {outputMsg "- Renamed 'PMI Representation' and 'PMI Presentation' to 'Semantic PMI' and 'Graphic PMI'"}
  if {$sfaVersion < 5.27} {outputMsg "- Help > Viewer > Hole Features"}
  if {$sfaVersion < 5.14} {outputMsg "- Help > Viewer > PMI Placeholders"}
  if {$sfaVersion < 5.02} {outputMsg "- Help > Viewer > Viewpoints"}
  if {$sfaVersion < 5.16} {outputMsg "- The CAx-IF website and Recommended Practices for PMI in AP242 have been updated, see Websites"}
  outputMsg "- See Help > Release Notes for all new features, updates, and bug fixes"
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
    outputMsg "\nThe User Guide is based on version 4.60" blue
    outputMsg "- See Help > Release Notes for updates\n- New and updated features are documented in the Help menu\n- The Options and Spreadsheet tabs have been renamed Generate and More\n- PMI Representation and PMI Presentation have been renamed Semantic PMI and Graphic PMI"
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
# generate tab
proc guiGenerateTab {} {
  global allNone buttons cb entCategory fopt fopta nb lastPartOnly opt optSave recPracNames useXL xlInstalled

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
  foreach item {{" Part Only" opt(partOnly)} {" BOM" opt(BOM)} {" Syntax Checker" opt(syntaxChecker)} {" Log File" opt(logFile)} {" Open Files  " opt(outputOpen)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $foptk.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {

# save and restore options if Part Only changes
      set opts [list partCap partSupp syntaxChecker tessPartOld viewCorrect viewFEA viewNoPMI viewParallel viewPMI]
      if {[info exists lastPartOnly]} {
        if {$opt(partOnly) != $lastPartOnly} {
          if {$opt(partOnly) == 1} {
            foreach i $opts {set optSave($i) $opt($i)}
          } elseif {$opt(partOnly) == 0} {
            foreach i $opts {catch {set opt($i) $optSave($i)}}
          }
        }
      } else {
        foreach i $opts {set optSave($i) $opt($i)}
      }
      set lastPartOnly $opt(partOnly)
      checkValues
    }]
    pack $buttons($idx) -side left -anchor w -padx {5 0} -pady {0 3} -ipady 0
    incr cb

    if {$idx != "outputOpen"} {
      pack [ttk::separator $foptk.$cb -orient vertical] -side left -anchor w -padx {10 5} -pady {0 3} -ipady 9
      incr cb
    }
  }

  pack $foptk -side left -anchor w -pady {5 2} -padx 10 -fill both -expand true

  set txt "Spreadsheets contain one worksheet for each STEP entity type.  The categories below\ncontrol which STEP entity types are written to the Spreadsheet.  Analyzer options\nbelow also write information to the Spreadsheet.  See the More tab for more options.\n\nIf Excel is installed, then Spreadsheets and CSV files can be generated.  If CSV Files\nis selected, the Spreadsheet is also generated.  CSV files do not contain any cell\ncolors, comments, or links.\n\nIf Excel is not installed, only CSV files can be generated.  Analyzer options are disabled."
  catch {tooltip::tooltip $buttons(genExcel) $txt}
  catch {tooltip::tooltip $buttons(genCSV) $txt}
  set txt "The Viewer supports b-rep and tessellated part geometry, graphic PMI, sketch\ngeometry, supplemental geometry, datum targets, finite element models, and more.\nUse the Viewer options below to control what features of the STEP file are shown.\n\nPart Only generates only Part Geometry.  This is useful when no other Viewer\nfeatures are needed and for large STEP files greater than about 430 MB.\n\nSee Help > Viewer"
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
        append ttmsg "Entity Type categories control which entities from AP242, AP203, and AP214 are written to the Spreadsheet.  The categories\nare used to group and color-code entities on the Summary worksheet.  All entities specific to other APs are always written\nto the Spreadsheet.  See Help > Supported STEP APs\n\n"
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
          append ttmsg " AP242 and AP214."
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
  foreach item {{" AP242" opt(stepAP242)} {" Composites" opt(stepCOMP)} {" Kinematics" opt(stepKINE)}} {
    set idx [string range [lindex $item 1] 4 end-1]
    set buttons($idx) [ttk::checkbutton $fopta4.$cb -text [lindex $item 0] -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    if {[info exists entCategory($idx)]} {
      set ttmsg "[llength $entCategory($idx)] [string trim [lindex $item 0]] entities"
      if {$idx == "stepAP242"} {
        append ttmsg " are supported in AP242.  Commonly used AP242 entities are in the other Entity Type categories.\nSuperscript indicates edition of AP242"
      } else {
        append ttmsg " are supported in AP242"
        if {$idx == "stepCOMP"} {append ttmsg " and AP203."}
        if {$idx == "stepKINE"} {append ttmsg " and AP214."}
      }
      set ttmsg [guiToolTip $ttmsg $idx [string trim [lindex $item 0]]]
      if {$idx == "stepAP242"} {append ttmsg "\n\nAssembly Structure is also supported by the AP242 Domain Model XML.  See Websites > CAx Recommended Practices"}
      if {$idx == "stepCOMP"} {append ttmsg "\n\nSee Websites > Recommended Practices for Composite Materials"}
      if {$idx == "stepKINE"} {append ttmsg "\n\nKinematics is also supported by the AP242 Domain Model XML.  See Websites > CAx Recommended Practices"}
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
  catch {tooltip::tooltip $fopta "Entity Type categories control which entities from AP242, AP203, and AP214\nare written to the Spreadsheet.  The categories are used to group and\ncolor-code entities on the Summary worksheet.  All entities specific to other\nAPs are always written to the Spreadsheet.\n\nSee Help > Supported STEP APs"}

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
    tooltip::tooltip $foptv20 "The Viewer supports b-rep and AP242 tessellated part geometry, color,\ntransparency, edges, sketch and supplemental geometry, and clipping planes.\nTesselated part geometry is typically written to an AP242 file instead of or in\naddition to b-rep part geometry.\n\nThe Viewer uses the default web browser.  An Internet connection is required.\nThe Viewer does not support measurements.\n\nSee the More tab for more Viewer options.\nSee Help > Viewer > Overview and other topics"
    tooltip::tooltip $buttons(viewPMI) "Graphic PMI for annotations is supported in AP242, AP203, and AP214 files.\nPMI (annotation) placeholders are supported in AP242.\n\nA Saved View is a subset of graphic PMI which has its own viewpoint position\nand orientation.  Use PageDown in the Viewer to cycle through saved views to\nswitch to the associated viewpoint and subset of graphic PMI.\n\nSee the options related to viewpoints on the More tab.\nSee Help > Viewer > Graphic PMI\nSee Help > Viewer > PMI Placeholders\nSee Help > User Guide (section 4.2)\n\nGraphic PMI and placeholders must conform to recommended practices.\nSee Websites > CAx Recommended Practices"
    tooltip::tooltip $buttons(feaLoadScale) "The length of load vectors can be scaled by their magnitude.\nLoad vectors are always colored by their magnitude."
    tooltip::tooltip $buttons(feaDispNoTail) "The length of displacement vectors with a tail are scaled by\ntheir magnitude.  Vectors without a tail are not.\nDisplacement vectors are always colored by their magnitude.\nLoad vectors always have a tail."
    tooltip::tooltip $foptv21 "Quality controls the number of facets used for curved surfaces.\nFor example, the higher the quality the more facets around the\ncircumference of a cylinder.\n\nNormals improve the default smooth shading of surfaces.  Using\nHigh Quality and Normals results in the best appearance for b-rep\npart geometry.  Quality and normals do not apply to tessellated\npart geometry.\n\nIf curved surfaces for b-rep part geometry look wrong even with\nQuality set to High, then use the Alternative B-rep Geometry\nProcessing method on the More tab."
    tooltip::tooltip $foptv4  "For 'By View' PMI colors, all of the annotations in a Saved View are set to the\nsame color.  If there is only one or no Saved Views, then 'Random' PMI colors\nare used.  For 'Random' PMI colors, each annotation is set to a different color.\nPMI color does not apply to annotation placeholders which are always black."
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
  global buttons cb developer fileDir fxls mydocs nb opt pmiElementsMaxRows recPracNames userWriteDir writeDir xlsManual

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
  set msg "Maximum rows limits the number of rows (entities) written to any one worksheet.\nIf the maximum number of rows is exceeded, the number of entities processed will be reported\nas, for example, 'cartesian_point (100 of 147)'.  For large STEP files, setting a low maximum\ncan speed up processing at the expense of not processing all of the entities.\n\nMaximum rows is ignored for entities where Analyzer results are reported.  Syntax Errors\nmight be missed if some entities are not processed due to a low value of maximum rows.\nMaximum rows does not affect the Viewer.\n\nSee Help > User Guide (section 5.5.1)"
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
    tooltip::tooltip $buttons(PMISEMRND)   "Rounding values might result in a better match to graphic PMI shown in\nthe Viewer or to expected PMI in the NIST CAD models (FTC/STC 7, 8, 11).\n\nSee User Guide (section 6.1.3.1)\nSee Websites > Recommended Practice for\n   $recPracNames(pmi242), Section 5.4"
    tooltip::tooltip $buttons(viewParallel) "Use parallel projection defined in the STEP file for saved view viewpoints,\ninstead of the default perspective projection.  Pan and zoom might not\nwork with parallel projection.  See Help > Viewer > Viewpoints"
    tooltip::tooltip $buttons(viewCorrect) "Correct for older implementations of camera models that\nmight not conform to current recommended practices.\nThe corrected viewpoint might fix the orientation but\nmaybe not the position.\n\nSee Help > Viewer > Viewpoints\nSee the CAx-IF Recommended Practice for\n $recPracNames(pmi242), Sec. 9.4.2.6"
    tooltip::tooltip $buttons(viewNoPMI)   "If the model has viewpoints with and without graphic PMI,\nthen also show the viewpoints without graphic PMI.  Those\nviewpoints are typically top, front, side, etc."
    tooltip::tooltip $buttons(debugVP)     "Debug viewpoint orientation defined by a camera model\nby showing the view frustum in the Viewer.\n\nSee Help > Viewer > Viewpoints\nSee the CAx-IF Recommended Practice for\n $recPracNames(pmi242), Sec. 9.4.2.6"
    tooltip::tooltip $buttons(partCap)     "Generate capped surfaces for section view clipping planes.  Capped\nsurfaces might take a long time to generate or look wrong for parts\nin an assembly.  See Help > Viewer > Other Features"
    tooltip::tooltip $buttons(brepAlt)     "If curved surfaces for Part Geometry look wrong even with\nQuality set to High, use an alternative b-rep geometry\nprocessing algorithm.  It will take longer to process the STEP\nfile and the resulting Viewer file will be larger."
    tooltip::tooltip $buttons(tessPartOld) "Process AP242 tessellated part geometry with the old method in SFA < 5.10.\nIt is not recommended for assemblies or large STEP files."
    tooltip::tooltip $buttons(x3dSave)     "The X3D file can be shown in an X3D viewer or imported to other software.\nUse this option if an Internet connection is not available for the Viewer.\nSee Help > Viewer"
    tooltip::tooltip $buttons(partNoGroup) "This option might create a very long list of parts names in the Viewer.\nIdentical parts have a underscore and number appended to their name.\nSee Help > Assemblies"
  }

# output directory
  set fxlse [ttk::labelframe $fxls.e -text " Write Output to "]
  set fxls1 [frame $fxlse.1]

# writeDirType = 0
  set buttons(fileDir) [ttk::radiobutton $fxls1.$cb -text " Same directory as STEP file" -variable opt(writeDirType) -value 0 -command checkValues]
  pack $fxls1.$cb -side left -anchor w -padx 5 -pady 2
  incr cb
  catch {tooltip::tooltip $buttons(fileDir) "If possible, existing output files are always overwritten by new files.\nIf spreadsheets cannot be overwritten, a number is appended to the\nfile name: myfile-sfa-1.xlsx"}

# writeDirType = 2
  set buttons(userDir) [ttk::radiobutton $fxls1.$cb -text " User-defined directory: " -variable opt(writeDirType) -value 2 -command {
    checkValues
    if {[file exists $userWriteDir] && [file isdirectory $userWriteDir]} {
      set writeDir $userWriteDir
    } else {
      set userWriteDir $mydocs
      tk_messageBox -type ok -icon error -title "Invalid Directory" \
        -message "The user-defined directory is not valid.\nIt has been set to $userWriteDir"
    }
    focus $buttons(userBrowse)
  }]
  pack $fxls1.$cb -side left -anchor w -padx {5 0}
  incr cb

  set buttons(userEntry) [ttk::entry $fxls1.entry -width 35 -textvariable userWriteDir]
  pack $fxls1.entry -side left -anchor w -pady 2
  set buttons(userBrowse) [ttk::button $fxls1.button -text "Browse" -command {
    set uwd [tk_chooseDirectory -title "Select directory"]
    if {[file isdirectory $uwd]} {
      set userWriteDir $uwd
      set writeDir $userWriteDir
    }
  }]
  pack $fxls1.button -side top -anchor w -padx 5 -pady 2
  pack $fxls1 -side top -anchor w
  set ttmsg "This option is useful when the directory containing the STEP file is\nprotected (read-only) and none of the output can be written to it."
  catch {
    tooltip::tooltip $buttons(userDir)    $ttmsg
    tooltip::tooltip $buttons(userEntry)  $ttmsg
    tooltip::tooltip $buttons(userBrowse) $ttmsg
  }

# manually save spreadsheet
  set buttons(xlsManual) [ttk::checkbutton $fxlse.$cb -text " Manually save Excel spreadsheet" -variable xlsManual -command checkValues]
  pack $fxlse.$cb -side top -anchor w -padx 5 -pady 2
  incr cb
  catch {tooltip::tooltip $buttons(xlsManual) "Use this option if other Excel options need to be set when saving\nthe spreadsheet or to set a user-defined file name.  Does not work\nwhen processing multiple STEP files."}

  pack $fxlse -side top -anchor w -pady {10 2} -padx 10 -fill both

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
# Examples and Websites menus
proc guiWebsitesMenu {} {
  global Examples Websites

  $Websites add command -label "STEP File Analyzer and Viewer" -command {openURL https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer}
  $Websites add command -label "- on GitHub"                   -command {openURL https://github.com/usnistgov/SFA}
  $Websites add command -label "- on CAx-IF"                   -command {openURL https://www.mbx-if.org/home/cax/resources/sfa/}
  $Websites add separator
  $Websites add command -label "CAx Interoperability Forum (CAx-IF)" -command {openURL https://www.mbx-if.org/home/cax/}
  $Websites add command -label "CAx Recommended Practices"           -command {openURL https://www.mbx-if.org/home/cax/recpractices/}
  $Websites add command -label "CAD Implementations"                 -command {openURL https://www.mbx-if.org/home/cax/implementation-coverage/}
  $Websites add command -label "PDM-IF"                              -command {openURL https://www.mbx-if.org/home/pdm/}

  $Websites add separator
  $Websites add cascade -label "AP242" -menu $Websites.0
  set Websites0 [menu $Websites.0 -tearoff 1]
  $Websites0 add command -label "AP242 Project"           -command {openURL http://www.ap242.org}
  $Websites0 add command -label "Benchmark Testing"       -command {openURL http://www.asd-ssg.org/step-ap242-benchmark.html}
  $Websites0 add command -label "Domain Model XML"        -command {openURL https://www.mbx-if.org/home/pdm/recpractices/}

  $Websites0 add separator
  $Websites0 add command -label "ISO 10303-242"           -command {openURL https://www.iso.org/standard/84300.html}
  $Websites0 add command -label "STEP in 3D PDF"          -command {openURL https://www.iso.org/standard/77686.html}
  $Websites0 add command -label "STEP Geometry Services"  -command {openURL https://www.iso.org/standard/84820.html}

  $Websites add cascade -label "STEP" -menu $Websites.2
  set Websites2 [menu $Websites.2 -tearoff 1]
  $Websites2 add command -label "STEP File Format"       -command {openURL https://www.loc.gov/preservation/digital/formats/fdd/fdd000448.shtml}
  $Websites2 add command -label "STEP ISO 10303-21"      -command {openURL https://en.wikipedia.org/wiki/ISO_10303-21}
  $Websites2 add command -label "STEP File Viewers"      -command {openURL https://www.mbx-if.org/home/mbx/resources/}
  $Websites2 add command -label "STEP to X3D Translator" -command {openURL https://www.nist.gov/services-resources/software/step-x3d-translator}

  $Websites2 add separator
  $Websites2 add command -label "AP209 FEA"     -command {openURL http://www.ap209.org}
  $Websites2 add command -label "AP238 STEP-NC" -command {openURL https://www.ap238.org}
  $Websites2 add command -label "AP239 PLCS"    -command {openURL http://www.ap239.org}
  $Websites2 add separator
  $Websites2 add command -label "EXPRESS ISO 10303-11"           -command {openURL https://www.loc.gov/preservation/digital/formats/fdd/fdd000449.shtml}
  $Websites2 add command -label "EXPRESS data modeling language" -command {openURL https://en.wikipedia.org/wiki/EXPRESS_(data_modeling_language)}
  $Websites2 add command -label "EXPRESS Schemas"                -command {openURL https://www.mbx-if.org/home/mbx/resources/express-schemas/}
  $Websites2 add command -label "Learning EXPRESS"               -command {openURL https://www.expresslang.org/learn/}
  $Websites2 add command -label "easyEXPRESS"                    -command {openURL https://marketplace.visualstudio.com/items?itemName=usnistgov.easyEXPRESS}

  $Websites add cascade -label "Organizations" -menu $Websites.4
  set Websites4 [menu $Websites.4 -tearoff 1]
  $Websites4 add command -label "PDES, Inc. (U.S.)"      -command {openURL https://pdesinc.org/}
  $Websites4 add command -label "prostep ivip (Germany)" -command {openURL https://www.prostep.org/en/projects/mbx-interoperability-forum-mbx-if}
  $Websites4 add command -label "AFNeT (France)"         -command {openURL https://atlas.afnet.fr/en/domaines/plm/}
  $Websites4 add command -label "KStep (Korea)"          -command {openURL https://www.kstep.or.kr}
  $Websites4 add separator
  $Websites4 add command -label "MBx Interoperability Forum (MBx-IF)"       -command {openURL https://www.mbx-if.org/home/}
  $Websites4 add command -label "LOTAR - LOng Term Archiving and Retrieval" -command {openURL https://lotar-international.org}
  $Websites4 add command -label "ISO/TC 184/SC 4 - Industrial Data"         -command {openURL https://committee.iso.org/home/tc184sc4}
  $Websites4 add command -label "JT-IF"                                     -command {openURL https://www.prostep.org/en/projects/jt-project-groups-jt-wf-jt-if-jt-bm}

  $Examples add command -label "Viewer"                   -command {openURL https://pages.nist.gov/CAD-PMI-Testing/}
  $Examples add command -label "Spreadsheets - AP242 PMI" -command {openURL https://www.nist.gov/document/sfa-semantic-pmi-spreadsheet}
  $Examples add command -label "- AP203 PMI"              -command {openURL https://www.nist.gov/document/sfa-spreadsheet}
  $Examples add command -label "- PMI Coverage Analysis"  -command {openURL https://www.nist.gov/document/sfa-multiple-files-spreadsheet}
  $Examples add command -label "- Bill of Materials"      -command {openURL https://www.nist.gov/document/sfa-bill-materials-spreadsheet}
  $Examples add separator
  $Examples add command -label "Sample STEP Files (zip)" -command {openURL https://www.nist.gov/document/nist-pmi-step-files}
  $Examples add command -label "NIST CAD Models"         -command {openURL https://www.nist.gov/ctl/smart-connected-systems-division/smart-connected-manufacturing-systems-group/mbe-pmi-0}
  $Examples add command -label "- on CAx-IF"             -command {openURL https://www.mbx-if.org/home/cax/resources/}
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
  }

  if {$opt(tessPartOld)} {
    set tessSolid 0
    set opt(viewTessPart) 1
    set opt(tessPartMesh) 1
  }

  if {!$gen(Excel)} {
    lappend butDisabled labelMaxRows xlHideLinks xlUnicode xlSort xlNoRound xlsManual
    if {$opt(xlFormat) == "None"} {for {set i 0} {$i < 8} {incr i} {lappend butDisabled "maxrows$i"}}
  } else {
    lappend butNormal labelMaxRows xlHideLinks xlUnicode xlSort xlNoRound xlsManual
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
    lappend butDisabled xlHideLinks xlUnicode xlSort xlNoRound BOM INVERSE PMIGRF PMISEM valProp genExcel xlsManual
    lappend butNormal viewFEA viewPMI viewPart
    lappend butNormal allNone0 allNone1 stepUSER

# Excel
  } else {
    foreach item [array names opt] {
      if {[string first "step" $item] == 0} {lappend butNormal $item}
    }
    lappend butNormal xlHideLinks xlUnicode xlSort xlNoRound BOM INVERSE PMIGRF PMISEM valProp xlsManual
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
    lappend butDisabled userBrowse userEntry
  } elseif {$opt(writeDirType) == 2} {
    lappend butNormal userBrowse userEntry
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
