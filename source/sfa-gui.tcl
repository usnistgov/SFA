#-------------------------------------------------------------------------------
# start window, bind keys

proc guiStartWindow {} {
  global winpos wingeo localName localNameList lastXLS lastXLS1 fout mingeo
  
  wm title . "STEP File Analyzer  (v[getVersion])"
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

# check that the saved window geometry keeps the entire window on the screen and isn't smaller than the minimum
    if {[expr {$pwid+$gwid}] > [winfo screenwidth  .]} {
      set gwid [expr {[winfo screenwidth  .]-$pwid-40}]
      set mwid [lindex [split $mingeo "x"] 0]
      if {$gwid < $mwid} {set gwid $mwid}
    }
    if {[expr {$phgt+$ghgt}] > [winfo screenheight  .]} {
      set ghgt [expr {[winfo screenheight  .]-$phgt-40}]
      set mhgt [lindex [split $mingeo "x"] 1]
      if {$ghgt < $mhgt} {set ghgt $mhgt}
    }
    set wingeo "$gwid\x$ghgt"
  }

# set the window position and dimensions
  if {[info exists winpos]} {catch {wm geometry . $winpos}}
  if {[info exists wingeo]} {catch {wm geometry . $wingeo}}

# fonts
  #set fontfam {MS San Serif}
  #set normalfont [list $fontfam]
  #set boldfont   [list $fontfam 0 bold]
  #
  #catch {option add *Button.font      $normalfont}
  #catch {option add *Checkbutton.font $normalfont}
  #catch {option add *Entry.font       $normalfont}
  #catch {option add *Label.font       $normalfont}
  #catch {option add *Radiobutton.font $normalfont}

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

# control o,q
  bind . <Control-o> {openFile}
  bind . <Control-d> {openMultiFile}
  bind . <Key-F4>    {openMultiFile 0}
  bind . <Control-q> {exit}

  bind . <Key-F1> {
    set localName [getFirstFile]
    if {$localName != ""} {
      set localNameList [list $localName]
      genExcel
    }
  }

  bind . <Key-F2> {set lastXLS [openXLS $lastXLS 1]}
  if {$lastXLS1 != ""} {bind . <Key-F3> {set lastXLS1 [openXLS $lastXLS1 1]}}

  bind . <MouseWheel> {[$fout.text component text] yview scroll [expr {-%D/30}] units}
  bind . <Up>     {[$fout.text component text] yview scroll -1 units}
  bind . <Down>   {[$fout.text component text] yview scroll  1 units}
  #bind . <Left>   {[$fout.text component text] xview scroll -1 units}
  #bind . <Right>  {[$fout.text component text] xview scroll  1 units}
  bind . <Prior>  {[$fout.text component text] yview scroll -30 units}
  bind . <Next>   {[$fout.text component text] yview scroll  30 units}
  bind . <Home>   {[$fout.text component text] yview scroll -100000 units}
  bind . <End>    {[$fout.text component text] yview scroll  100000 units}
}

#-------------------------------------------------------------------------------
# buttons and progress bar

proc guiButtons {} {
  global buttons wdir nline nprogfile ftrans mytemp opt nistVersion
  
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

# NIST logo
  if {$nistVersion} {
    catch {
      set l3 [label $ftrans.l3 -relief flat -bd 0]
      $l3 config -image [image create photo -file [file join $wdir images nist.gif]]
      pack $l3 -side right -padx 10
      bind $l3 <ButtonRelease-1> {displayURL http://www.nist.gov/el/}
      tooltip::tooltip $l3 "Click here"
    }
  }

  pack $ftrans -side top -padx 10 -pady 10 -fill x

  set fbar [frame .fbar -bd 2 -background "#F0F0F0"]
  set nline 0
  set buttons(pgb) [ttk::progressbar $fbar.pgb -mode determinate -variable nline]
  pack $fbar.pgb -side top -padx 10 -fill x

  set nprogfile 0
  set buttons(pgb1) [ttk::progressbar $fbar.pgb1 -mode determinate -variable nprogfile]
  pack forget $buttons(pgb1)
  #pack $fbar.pgb1 -side top -padx 10 -pady {5 0} -expand true -fill x
  pack $fbar -side bottom -padx 10 -pady {0 10} -fill x
  
# NIST icon bitmap
  if {$nistVersion} {
    catch {wm iconbitmap . -default [file join $wdir images NIST.ico]}
  }
}

#-------------------------------------------------------------------------------
# status tab

proc guiStatusTab {} {
  global nb wout fout outputWin statusFont tcl_platform developer

  set wout [ttk::panedwindow $nb.status -orient horizontal]
  $nb add $wout -text " Status " -padding 2
  set fout [frame $wout.fout -bd 2 -relief sunken -background "#E0DFE3"]

  set outputWin [iwidgets::messagebox $fout.text -maxlines 500000 \
    -hscrollmode dynamic -vscrollmode dynamic -background white]
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
  
  if {$tcl_platform(osVersion) >= 6.0} {
    if {![info exists statusFont]} {
      set statusFont [$outputWin type cget black -font]
      #set newsize [expr {int((508./[winfo screenmmwidth .])*120.)}]
      #if {[string index $newsize 2] != 0} {
      #  set newsize [expr {round($newsize/10.)*10}]
      #  for {set i 210} {$i >= 100} {incr i -10} {regsub -all $i $statusFont $newsize statusFont}
      #}
    }
    if {[string first "Courier" $statusFont] != -1} {
      regsub "Courier" $statusFont "Consolas" statusFont
      saveState
    }
  }
  
  if {[info exists statusFont]} {
    foreach typ {black red green magenta cyan blue error syntax} {
      $outputWin type configure $typ -font $statusFont
    }
  }
  
  bind . <Key-F6> {
    set statusFont [$outputWin type cget black -font]
    for {set i 210} {$i >= 100} {incr i -10} {regsub -all $i $statusFont [expr {$i+10}] statusFont}
    foreach typ {black red green magenta cyan blue error syntax} {
      $outputWin type configure $typ -font $statusFont
    }
    #if {$developer} {outputMsg $statusFont}
  }
  bind . <Control-KeyPress-=> {
    set statusFont [$outputWin type cget black -font]
    for {set i 210} {$i >= 100} {incr i -10} {regsub -all $i $statusFont [expr {$i+10}] statusFont}
    foreach typ {black red green magenta cyan blue error syntax} {
      $outputWin type configure $typ -font $statusFont
    }
    #if {$developer} {outputMsg $statusFont}
  }
  
  bind . <Key-F5> {
    set statusFont [$outputWin type cget black -font]
    for {set i 110} {$i <= 220} {incr i 10} {regsub -all $i $statusFont [expr {$i-10}] statusFont}
    foreach typ {black red green magenta cyan blue error syntax} {
      $outputWin type configure $typ -font $statusFont
    }
    #if {$developer} {outputMsg $statusFont}
  }
  bind . <Control-KeyPress--> {
    set statusFont [$outputWin type cget black -font]
    for {set i 110} {$i <= 220} {incr i 10} {regsub -all $i $statusFont [expr {$i-10}] statusFont}
    foreach typ {black red green magenta cyan blue error syntax} {
      $outputWin type configure $typ -font $statusFont
    }
    #if {$developer} {outputMsg $statusFont}
  }
  
  #bind . <Key-F7> {
  #  set statusFont [$outputWin type cget black -font]
  #  if {[string first "Courier" $statusFont] != -1} {
  #    regsub "Courier" $statusFont "Consolas" statusFont
  #  } else {
  #    regsub "Consolas" $statusFont "Courier" statusFont
  #  }
  #  foreach typ {black red green magenta cyan blue error syntax} {
  #    $outputWin type configure $typ -font $statusFont
  #  }
  #}

# excel colors
  #$outputWin type add excel12 -foreground black -background "#808000"
  #$outputWin type add excel15 -foreground black -background "#C0C0C0"
  #$outputWin type add excel16 -foreground black -background "#808080"
  #$outputWin type add excel17 -foreground black -background "#9999FF"
  #$outputWin type add excel19 -foreground black -background "#FFFFCC"
  #$outputWin type add excel20 -foreground black -background "#CCFFFF"
  #$outputWin type add excel22 -foreground black -background "#FF8080"
  #$outputWin type add excel24 -foreground black -background "#CCCCFF"
  #$outputWin type add excel34 -foreground black -background "#CCFFFF"
  #$outputWin type add excel35 -foreground black -background "#CCFFCC"
  #$outputWin type add excel36 -foreground black -background "#FFFF99"
  #$outputWin type add excel37 -foreground black -background "#99CCFF"
  #$outputWin type add excel38 -foreground black -background "#FF99CC"
  #$outputWin type add excel39 -foreground black -background "#CC99FF"
  #$outputWin type add excel42 -foreground black -background "#33CCCC"
  #$outputWin type add excel43 -foreground black -background "#99CC00"
  #$outputWin type add excel44 -foreground black -background "#FFCC00"
  #$outputWin type add excel48 -foreground black -background "#969696"
}

#-------------------------------------------------------------------------------
# file menu
proc guiFileMenu {} {
  global File openFileList lastXLS lastXLS1

  $File add command -label "Open STEP File(s)..." -accelerator "Ctrl+O" -command openFile
  $File add command -label "Open Multiple STEP Files in a Directory..." -accelerator "Ctrl+D, F4" -command {openMultiFile}
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
  $File add command -label "Open Last Spreadsheet" -accelerator "F2" -command {set lastXLS [openXLS $lastXLS 1]}
  if {$lastXLS1 != ""} {
    $File add command -label "Open Last Multiple File Summary Spreadsheet" -accelerator "F3" -command {set lastXLS1 [openXLS $lastXLS1 1]}
  }
  $File add command -label "Exit" -accelerator "Ctrl+Q" -command exit
}

#-------------------------------------------------------------------------------
# options tab, process and report
proc guiProcessAndReports {} {
  global fopt fopta nb opt cb buttons entCategory developer userentlist userEntityFile

  set cb 0
  set wopt [ttk::panedwindow $nb.opt -orient horizontal]
  $nb add $wopt -text " Options " -padding 2
  set fopt [frame $wopt.fopt -bd 2 -relief sunken]
  
  set fopta [ttk::labelframe $fopt.a -text " Process "]
  
  # option to process user-defined entities
  guiUserDefinedEntities
  
  set fopta1 [frame $fopta.1 -bd 0]
  foreach item {{" AP242" opt(PR_STEP_AP242)} \
                {" AP203" opt(PR_STEP_AP203)} \
                {" AP214" opt(PR_STEP_AP214)} \
                {" AP209" opt(PR_STEP_AP209)} \
                {" AP210" opt(PR_STEP_AP210)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta1.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    set tt [string range $idx 3 end]
    if {[info exists entCategory($tt)]} {
      if {$tt == "PR_STEP_AP203" || $tt == "PR_STEP_AP214"} {
        set ttmsg "[lindex $item 0] entities ([llength $entCategory($tt)])"
        append ttmsg "  Some entities are also found in AP242"
      } else {
        set ttmsg "These entities are unique to[lindex $item 0] ([llength $entCategory($tt)])"
      }
      set ttmsg1 $ttmsg
      append ttmsg "\n\n"
      set ttmsg [processToolTip $ttmsg $tt 210]
      append ttmsg "\n\n$ttmsg1"
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta1 -side left -anchor w -pady 0 -padx 0 -fill y
  
  set fopta2 [frame $fopta.2 -bd 0]
  foreach item {{" AP242 Geometry"     opt(PR_STEP_AP242_GEOM)} \
                {" AP242 Kinematics"   opt(PR_STEP_AP242_KINE)} \
                {" AP242 Data Quality" opt(PR_STEP_AP242_QUAL)} \
                {" AP242 Constraint"   opt(PR_STEP_AP242_CONS)} \
                {" AP242 Math"         opt(PR_STEP_AP242_MATH)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta2.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    set tt [string range $idx 3 end]
    if {[info exists entCategory($tt)]} {
      set ttmsg "[lindex $item 0] ([llength $entCategory($tt)])  These entities are unique to AP242"
      append ttmsg "\n\n"
      set ttmsg [processToolTip $ttmsg $tt]
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta2 -side left -anchor w -pady 0 -padx 0 -fill y
  
  set fopta3 [frame $fopta.3 -bd 0]
  foreach item {{" Common"         opt(PR_STEP_OTHER)} \
                {" Shape Aspect"   opt(PR_STEP_ASPECT)} \
                {" GD&T"           opt(PR_STEP_TOLR)} \
                {" Presentation"   opt(PR_STEP_PRES)} \
                {" Representation" opt(PR_STEP_REP)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta3.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {
      checkValues
    }]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    set tt [string range $idx 3 end]
    if {[info exists entCategory($tt)]} {
      set ttmsg "[lindex $item 0] entities ([llength $entCategory($tt)])\n\n"
      set ttmsg [processToolTip $ttmsg $tt 120]
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  pack $fopta3 -side left -anchor w -pady 0 -padx 0 -fill y
  
  set fopta4 [frame $fopta.4 -bd 0]
  foreach item {{" Geometry"        opt(PR_STEP_GEO)} \
                {" Cartesian Point" opt(PR_STEP_CPNT)} \
                {" Measure"         opt(PR_STEP_QUAN)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta4.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
    set tt [string range $idx 3 end]
    if {[info exists entCategory($tt)]} {
      set ttmsg "[lindex $item 0] entities ([llength $entCategory($tt)])\n\n"
      if {$tt == "PR_STEP_GEO" || $tt == "PR_STEP_CPNT"} {
        append ttmsg "For large STEP files, this option can slow down the processing of the file and increase the size of the spreadsheet.\nUse Maximum Rows options to possibly speed up the processing of these entities.\n\n"
      }
      set ttmsg [processToolTip $ttmsg $tt]
      catch {tooltip::tooltip $buttons($idx) $ttmsg}
    }
  }
  
  set anbut {{"All" 1} {"For Reports" 0}}
  foreach item $anbut {
    set bn "anbut[lindex $item 1]"            
    set buttons($bn) [ttk::radiobutton $fopta4.$cb -variable allnone -text [lindex $item 0] -value [lindex $item 1] \
      -command {
        if {$allnone} {
          foreach item [array names opt] {
            if {[string first "PR_STEP" $item] == 0} {
                set opt($item) $allnone
              
            }
          }
        } else {
          if {$opt(PMISEM) == 0 && $opt(PMIGRF) == 0 && $opt(VALPROP) == 0} {
            set opt(PMISEM) 1
            set opt(PMIGRF) 1
            set opt(VALPROP) 1
          }
          set opt(PR_STEP_AP203) 1
          set opt(PR_STEP_AP214) 1
          set opt(PR_STEP_AP238) 1
        }
        set opt(PR_STEP_GEO)  0
        set opt(PR_STEP_CPNT) 0
        checkValues
      }]
    pack $buttons($bn) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  catch {
    tooltip::tooltip $buttons(anbut0) "Selects all of the Reports below and the required entities for the reports"
    tooltip::tooltip $buttons(anbut1) "Selects all Entity types except Geometry and Cartesian Point"
  }
  pack $fopta4 -side left -anchor w -pady 0 -padx 0 -fill y
  
  pack $fopta -side top -anchor w -pady {5 2} -padx 10 -fill both
  
#-------------------------------------------------------------------------------
# report
  set foptd [ttk::labelframe $fopt.1 -text " Report "]
  set foptd1 [frame $foptd.1 -bd 0]
  
  set foptd4 [frame $foptd.4 -bd 0]
  foreach item {{" PMI Representation (Semantic PMI)" opt(PMISEM)}} {
  regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptd4.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {
      if {$opt(PMISEM)} {set opt(INVERSE) 1}
      checkValues
    }]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptd4 -side top -anchor w -pady 0 -padx 0 -fill y
  
  set foptd3 [frame $foptd.3 -bd 0]
  foreach item {{" PMI Presentation (Graphical PMI)" opt(PMIGRF)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptd3.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {
      checkValues
    }]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptd3 -side top -anchor w -pady 0 -padx 0 -fill y
  
  set foptd5 [frame $foptd.5 -bd 0]
  foreach item {{" Visualize PMI Presentation" opt(GENX3DOM)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptd5.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {
      checkValues
    }]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  set buttons(linecolor) [label $foptd5.l3 -text "Line Color:"]
  pack $foptd5.l3 -side left -anchor n -padx 5 -pady 0 -ipady 0
  set gpmiColorVal {{" From file" 0} {" Black" 1} {" Random" 2}}
  foreach item $gpmiColorVal {
    set bn "gpmiColor[lindex $item 1]"            
    set buttons($bn) [ttk::radiobutton $foptd5.$cb -variable opt(gpmiColor) -text [lindex $item 0] -value [lindex $item 1]]
    pack $buttons($bn) -side left -anchor n -padx 1 -pady 0 -ipady 0
    incr cb
  }
  pack $foptd5 -side top -anchor w -pady 0 -padx 22 -fill y
  
  foreach item {{" Validation Properties" opt(VALPROP)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $foptd1.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {
      checkValues
    }]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptd1 -side top -anchor w -pady 0 -padx 0 -fill y
  
  pack $foptd -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {
    tooltip::tooltip $buttons(optPMISEM) "PMI Representation includes all information necessary to represent GD&T without any graphical presentation elements.\nPMI Representation is associated with CAD model geometry and is computer-interpretable to facilitate automated\nconsumption by downstream applications for manufacturing, measurement, inspection, and other processes.\n\nPMI Representation information is defined in a CAx-IF Recommended Practices\nand is reported for Dimensional Tolerances, Geometric Tolerances, and Datum Features.\nThe results are reported on various entities as indicated by PMI Representation on the Summary worksheet.\n\nSelecting this report automatically selects the required entities above for the report.\n\nSee Help > PMI Representation"
    tooltip::tooltip $buttons(optPMIGRF) "PMI Presentation (also known as graphical PMI) consists of geometric elements such as\nlines and arcs preserving the exact appearance (color, shape, positioning) of the GD&T\nannotations.  PMI Presentation is not intended to be computer-interpretable and does not\ncarry any representation information, although it can be linked to its corresponding\nPMI Representation.\n\nPMI Presentation annotations are defined in CAx-IF Recommended Practices.\nThe PMI Presentation information is reported in columns highlighted in yellow and green\non the Annotation_*_occurrence worksheets.\n\nAssociated presentation style, saved views, and PMI validation properties are also reported.\nA PMI coverage analysis worksheet is also generated.\nAn X3DOM file (WebGL) of the PMI Presentation can also be generated.\n\nSelecting this report automatically selects the required entities above for the report.\n\nSee Help > PMI Presentation"
    tooltip::tooltip $buttons(optGENX3DOM) "PMI Presentation annotations can be visualized with an\nX3DOM (WebGL) file that can be displayed in most web browsers.\n\nTessellated annotations are not supported.\nThe color of the line segments can be modified.\n\nSee Help > PMI Presentation"
    tooltip::tooltip $buttons(optVALPROP) "Validation properties for geometry, assemblies, PMI, annotations,\nattributes, and tessellations are defined in CAx-IF Recommended Practices.\nThe property values are reported in columns highlighted in yellow and green\non the Property_definition worksheet.  The worksheet can also be sorted and filtered.\n\nSelecting this report automatically selects the required entities above for the report.\n\nSee Help > Validation Properties"
  }
}

#-------------------------------------------------------------------------------
# user-defined list of entities
proc guiUserDefinedEntities {} {
  global fopta opt cb buttons fileDir userEntityFile userentlist
  
  set fopta6 [frame $fopta.6 -bd 0]
  foreach item {{" User-Defined List: " opt(PR_USER)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fopta6.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }

  set buttons(userentity) [ttk::entry $fopta6.entry -width 40 -textvariable userEntityFile]
  pack $fopta6.entry -side left -anchor w

  set buttons(userentityopen) [ttk::button $fopta6.$cb -text " Browse " -command {
    set typelist {{"All Files" {*}}}
    set uef [tk_getOpenFile -title "Select File of STEP Entities" -filetypes $typelist -initialdir $fileDir]
    if {$uef != "" && [file isfile $uef]} {
      set userEntityFile [file nativename $uef]
      outputMsg "User-defined STEP list: [truncFileName $userEntityFile]" blue
      set fileent [open $userEntityFile r]
      set userentlist {}
      while {[gets $fileent line] != -1} {
        set line [split [string trim $line] " "]
        foreach ent1 $line {lappend userentlist $ent1}
      }
      close $fileent
      set llist [llength $userentlist]
      if {$llist > 0} {
        outputMsg " ($llist) $userentlist"
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
      tooltip::tooltip $buttons($item) "The User-Defined List is defined in a text file with one STEP entity name per line.\nThis allows for more control to process only the required entity types.\nIt is also useful when processing large files that might crash the software."
    }
  }
  pack $fopta6 -side bottom -anchor w -pady 5 -padx 0 -fill y
}

#-------------------------------------------------------------------------------
# inverse relationships
proc guiInverse {} {
  global buttons cb fopt inverses opt developer
  
  set foptc [ttk::labelframe $fopt.3 -text " Inverse Relationships "]
  set txt " Display some Inverse Relationships and Backwards References (Used In) for\n PMI, Shape Aspect, Draughting Model, Annotations, Analysis"

  regsub -all {[\(\)]} opt(INVERSE) "" idx
  set buttons($idx) [ttk::checkbutton $foptc.$cb -text $txt \
    -variable opt(INVERSE) -command {
      checkValues
    }]
  pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
  incr cb

  pack $foptc -side top -anchor w -pady {5 2} -padx 10 -fill both
  set ttmsg "Inverse Relationships\n"
  set lent ""
  set litem ""
  foreach item [lsort $inverses] {
    set ok 1
    if {[string first "geometric_tolerance_with" $item] != -1} {set ok 0}
    if {[string first "related" $item] < [string first "relating" $item]} {set ok 0}
    if {$ok} {
      set ilist [split $item " "]
      set ent [lindex $ilist 0]
      if {$ent != $lent} {
        if {$litem != ""} {append ttmsg \n$litem}
        regsub " " $item "  (" item
        append item ")"
        set litem $item
        set lent [lindex $ent 0]
      } else {
        append litem "  ([lindex $ilist 1] [lindex $ilist 2])"
      }
    }
  }
  append ttmsg \n$litem
  append ttmsg "\n\nInverse Relationships are displayed on the entity worksheets.  The Inverse values are\ndisplayed in additional columns of the worksheets that are highlighted in light blue."
  catch {tooltip::tooltip $foptc $ttmsg}
}

#-------------------------------------------------------------------------------
# display result
proc guiDisplayResult {} {
  global buttons cb fopt appNames dispCmds appName dispApps foptf
  global edmWriteToFile edmWhereRules eeWriteToFile
  
  set foptf [ttk::labelframe $fopt.f -text " Display STEP File in "]

  set buttons(appCombo) [ttk::combobox $foptf.spinbox -values $appNames -width 35]
  pack $foptf.spinbox -side left -anchor w -padx 7 -pady {0 3}
  bind $buttons(appCombo) <<ComboboxSelected>> {
    set appName [$buttons(appCombo) get]
    catch {
      if {[string first "EDM Model Checker" $appName] == 0} {
        pack $buttons(edmWriteToFile)  -side left -anchor w -padx 5
        pack $buttons(edmWhereRules) -side left -anchor w -padx 5
      } else {
        pack forget $buttons(edmWriteToFile)
        pack forget $buttons(edmWhereRules)
      }
    }
    catch {
      if {[string first "Conformance Checker" $appName] != -1} {
        pack $buttons(eeWriteToFile) -side left -anchor w -padx 5
      } else {
        pack forget $buttons(eeWriteToFile)
      }
    }
    catch {
      if {$appName == "Indent STEP File (for debugging)"} {
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

  set buttons(appDisplay) [ttk::button $foptf.$cb -text " Display " -state disabled -command {
    displayResult
    saveState
  }]
  pack $foptf.$cb -side left -anchor w -padx {10 0} -pady {0 3}
  incr cb
  
  foreach item $appNames {
    if {[string first "EDM Model Checker" $item] == 0} {
      foreach item {{" Write results to a file" edmWriteToFile}} {
        regsub -all {[\(\)]} [lindex $item 1] "" idx
        set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] \
          -variable [lindex $item 1] -command {checkValues}]
        pack forget $buttons($idx)
        incr cb
      }
      foreach item {{" Check WHERE, UNIQUENESS rules" edmWhereRules}} {
        regsub -all {[\(\)]} [lindex $item 1] "" idx
        set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] \
          -variable [lindex $item 1] -command {checkValues}]
        pack forget $buttons($idx)
        incr cb
      }
    }
  }
  if {[lsearch -glob $appNames "*Conformance Checker*"] != -1} {
    foreach item {{" Write results to a file" eeWriteToFile}} {
      regsub -all {[\(\)]} [lindex $item 1] "" idx
      set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] \
        -variable [lindex $item 1] -command {checkValues}]
      pack forget $buttons($idx)
      incr cb
    }
  }
  if {[lsearch $appNames "Indent STEP File (for debugging)"] != -1} {
    foreach item {{" Include Geometry" indentGeometry} \
                  {" Include Styled_item" indentStyledItem}} {
      regsub -all {[\(\)]} [lindex $item 1] "" idx
      set buttons($idx) [ttk::checkbutton $foptf.$cb -text [lindex $item 0] \
        -variable opt([lindex $item 1]) -command {checkValues}]
      pack forget $buttons($idx)
      incr cb
    }
  }
  
  catch {tooltip::tooltip $foptf "This option is a convenient way to display a STEP file in other applications.\nThe pull-down menu will contain applications that can display a STEP file\nsuch as STEP viewers, browsers, and conformance checkers only if they are\ninstalled in their default location.\n\nThe 'Indent STEP File (for debugging)' option rearranges and indents the\nentities to show the hierarchy of information in a STEP file.  The 'indented'\nfile is written to the same directory as the STEP file or to the same\nuser-defined directory specified in the Spreadsheet tab.  Including Geometry\nor Styled_item can make the 'indented' file very large.\n\nThe 'Default STEP Viewer' option will open the STEP file in whatever\napplication is associated with STEP files.\n\nA text editor will always appear in the menu."}
  pack $foptf -side top -anchor w -pady {5 2} -padx 10 -fill both

  set foptk [ttk::labelframe $fopt.k -text " Output Format "]
  foreach item {{" Excel" Excel} {" CSV" CSV}} {
    pack [ttk::radiobutton $foptk.$cb -variable opt(XLSCSV) -text [lindex $item 0] -value [lindex $item 1] -command {checkValues}] -side left -anchor n -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $foptk -side top -anchor w -pady {5 2} -padx 10 -fill both
  catch {tooltip::tooltip $foptk "Microsoft Excel is required to generate spreadsheets.  CSV files will be generated if Excel is not installed.\n\nOne CSV file is generated for each entity type.  Reports, Inverse Relationships, and some of the\nSpreadsheet tab options are not available with CSV files."}
}

#-------------------------------------------------------------------------------
# spreadsheet tab
proc guiSpreadsheet {} {
  global buttons cb env extXLS fileDir fxls mydocs nb opt developer
  global userWriteDir userXLSFile writeDir
  
  set wxls [ttk::panedwindow $nb.xls -orient horizontal]
  $nb add $wxls -text " Spreadsheet " -padding 2
  set fxls [frame $wxls.fxls -bd 2 -relief sunken]

  set fxlsz [ttk::labelframe $fxls.z -text " Tables "]
  foreach item {{" Generate Tables for Sorting and Filtering" opt(SORT)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fxlsz.$cb -text [lindex $item 0] \
      -variable [lindex $item 1]]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsz -side top -anchor w -pady {5 2} -padx 10 -fill both

  set fxlsa [ttk::labelframe $fxls.a -text " Number Format "]
  foreach item {{" Do not round Real Numbers (See Help > Number Format)" opt(XL_FPREC)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fxlsa.$cb -text [lindex $item 0] \
      -variable [lindex $item 1]]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsa -side top -anchor w -pady {5 2} -padx 10 -fill both
  
  set fxlsb [ttk::labelframe $fxls.b -text " Maximum Rows for any worksheet"]
  set rlimit {{" No limit" 10000000} {" 100" 103} {" 500" 503} {" 1000" 1003} {" 5000" 5003} {" 10000" 10003} {" 50000" 50003} {" 100000" 100003}}
  foreach item $rlimit {
    pack [ttk::radiobutton $fxlsb.$cb -variable opt(ROWLIM) -text [lindex $item 0] -value [lindex $item 1]] -side left -anchor n -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsb -side top -anchor w -pady 5 -padx 10 -fill both
  set msg "This option will limit the number of rows (entities) written to any one worksheet.\nWithout setting a maximum, the row maximums are:\n\nExcel 2007 and later:  1,048,576 rows\nExcel 2003 and earlier:    65,536 rows\n\nFor large STEP files, setting a low maximum can speed up processing at the expense\nof not processing all of the entities.  This is useful when processing Geometry entities."
  catch {tooltip::tooltip $fxlsb $msg}

  set fxlsc [ttk::labelframe $fxls.c -text " Excel Options "]
  foreach item {{" Open spreadsheet after it has been generated" opt(XL_OPEN)} \
                {" Keep spreadsheet open while it is being generated (slow)" opt(XL_KEEPOPEN)} \
                {" Create links to STEP files and spreadsheets with multiple files" opt(XL_LINK1)}} {
    regsub -all {[\(\)]} [lindex $item 1] "" idx
    set buttons($idx) [ttk::checkbutton $fxlsc.$cb -text [lindex $item 0] \
      -variable [lindex $item 1] -command {checkValues}]
    pack $buttons($idx) -side top -anchor w -padx 5 -pady 0 -ipady 0
    incr cb
  }
  pack $fxlsc -side top -anchor w -pady {5 2} -padx 10 -fill both

  set fxlsd [ttk::labelframe $fxls.d -text " Write Spreadsheet to "]
  set buttons(fileDir) [ttk::radiobutton $fxlsd.$cb \
    -text " Same directory as the STEP file" \
    -variable opt(writeDirType) -value 0 -command checkValues]
  pack $fxlsd.$cb -side top -anchor w -padx 5 -pady 2
  incr cb

  set fxls1 [frame $fxlsd.1]
  ttk::radiobutton $fxls1.$cb -text " User-defined directory:  " \
    -variable opt(writeDirType) -value 2 -command {
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
  catch {tooltip::tooltip $fxls1.$cb "This option can be used when the directory containing the STEP file is\nprotected (read-only) and the Spreadsheet cannot be written to it."}
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
  ttk::radiobutton $fxls2.$cb -text " User-defined file name:  " \
    -variable opt(writeDirType) -value 1 -command {
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

  if {$developer} {
    set fxlsx [ttk::labelframe $fxls.x -text " Debug "]
    foreach item {{" Everything" opt(DEBUG1)} \
                  {" Inverses" opt(DEBUGINV)} \
                  {" Dimtol > geotol path" opt(DEBUG2)}} {
      regsub -all {[\(\)]} [lindex $item 1] "" idx
      set buttons($idx) [ttk::checkbutton $fxlsx.$cb -text [lindex $item 0] -variable [lindex $item 1]]
      pack $buttons($idx) -side left -anchor w -padx 5 -pady 0 -ipady 0
      incr cb
    }
    pack $fxlsx -side top -anchor w -pady {5 2} -padx 10 -fill both
  }

  pack $fxls -side top -fill both -expand true -anchor nw
}

#-------------------------------------------------------------------------------
# help menu
proc guiHelpMenu {} {
  global Help opt nistVersion mytemp

  $Help add command -label "User's Guide (pdf)" -command {displayGuide}
  $Help add command -label "What's New" -command {whatsNew}

  if {$nistVersion} {
    $Help add command -label "Check for Update" -command {
      set lastupgrade [expr {round(([clock seconds] - $upgrade)/86400.)}]
      outputMsg "The last check for an update was $lastupgrade days ago." red
      set os $tcl_platform(osVersion)
      if {$pf64 != ""} {append os ".64"}
      set url "http://ciks.cbt.nist.gov/cgi-bin/ctv/sfa_upgrade.cgi?version=[getVersion]&auto=-$lastupgrade&os=$os"
      if {[info exists excelYear]} {if {$excelYear != ""} {append url "&yr=[expr {$excelYear-2000}]"}}
      displayURL $url
    }
  }

  $Help add separator
  $Help add command -label "Sample STEP Files (zip)"                          -command {displayURL http://www.nist.gov/el/msid/infotest/upload/NIST_MBE_PMI_CTC_STEP_PMI.zip}
  $Help add cascade -label "Sample Output" -menu $Help.5
  set Help5 [menu $Help.5 -tearoff 1]
  $Help5 add command -label "Spreadsheet - PMI Representation"          -command {displayURL http://www.nist.gov/el/msid/infotest/upload/STEP-File-Analyzer-PMI-Representation_stp.xlsx}
  $Help5 add command -label "Spreadsheet - PMI Presentation, ValProps"  -command {displayURL http://www.nist.gov/el/msid/infotest/upload/STEP-File-Analyzer_stp.xlsx}
  $Help5 add command -label "Spreadsheet - Coverage Analysis"           -command {displayURL http://www.nist.gov/el/msid/infotest/upload/STEP-File-Analyzer-Coverage.xlsx}
  $Help5 add command -label "X3DOM (WebGL) file - PMI Presentation"     -command {displayURL http://www.nist.gov/el/msid/infotest/upload/STEP-File-Analyzer_x3dom.html}

  $Help add separator
  $Help add command -label "Overview" -command {
outputMsg "\nOverview -------------------------------------------------------------------" blue
outputMsg "The STEP File Analyzer reads a STEP file and generates an Excel spreadsheet or CSV files.  One
worksheet or CSV file is generated for each entity type in the STEP file.  Each worksheet or CSV
file lists every entity instance and its attributes.  The types of entities that are Processed
can be selected in the Options tab.  Other options are available that add to or modify the
information written to the spreadsheet or CSV files.
  
For spreadsheets, a Summary worksheet shows the Count of each entity.  Links on the Summary and
entity worksheets can be used to navigate to other worksheets.

Spreadsheets or CSV files can be selected in the Options tab.  CSV files are automatically
generated if Excel is not installed.
  
To generate a spreadsheet or CSV files, select a STEP file from the File menu above and click
the Generate button below.  Existing spreadsheets or CSV files are always overwritten.

Multiple STEP files can be selected or an entire directory structure of STEP files can also be
processed from the File menu. If multiple STEP files are translated, then a separate File Summary
spreadsheet is also generated.

Tooltip help is available for the selections in the tabs.  Hold the mouse over text in the tabs
until a tooltip appears.

Function keys F5 and F6 change the font size in this tab."

    .tnb select .tnb.status
    update idletasks
  }

# options help
  $Help add command -label "Options" -command {
outputMsg "\nOptions --------------------------------------------------------------------" blue
outputMsg "*Process: Select which types of entities are processed.  The tooltip help lists all the entities
associated with that type.  Selectively process only the entities relevant to your analysis.
Entity types can also be selected with the All and None buttons.  The None button will select all
Reports and associated entities."

outputMsg "\n*Report PMI Representation: Dimensional tolerances, geometric tolerances, and datum features are
reported on various entities indicated by PMI Representation on the Summary worksheet.  Values are
reported in columns highlighted in yellow and green on those worksheets."

outputMsg "\n*Report PMI Presentation: Geometric entities used for PMI Presentation annotations are reported
in columns highlighted in yellow and green on Annotation_*_occurrence worksheets.  Associated
Saved Views, Validation Properties, and Geometry are also reported.  PMI Presentation annotations
can also be visualized with an X3DOM (WebGL) file."

outputMsg "\n*Report Validation Properties: Geometric, assembly, PMI, annotation, attribute, and tessellated
validation properties are reported.  The property values are reported in columns highlighted in
yellow and green on the Property_definition worksheet."

outputMsg "\n*Inverse Relationships: For some entities, Inverse relationships and backwards references
(Used In) are displayed on the worksheets.  The values are displayed in additional columns of
entity worksheets that are highlighted in light blue and purple."

outputMsg "\n*Output Format: Generate Excel spreadsheets or CSV files.  If Excel is not installed,
CSV files are automatically generated.  Some options are not available with CSV files."

outputMsg "\n*Table: Generate tables for each spreadsheet to facilitate sorting and filtering (Spreadsheet tab)."

outputMsg "\n*Number Format: Option to not round real numbers."

outputMsg "\n*Maximum Rows: The maximum number of rows for any worksheet can be set lower than the normal
limits for Excel.  This is useful for very large STEP files at the expense of not processing some
entities."

    .tnb select .tnb.status
    update idletasks
  }

# validation properties, PMI presentation, conformance checking help
  $Help add command -label "PMI Representation (Semantic PMI)" -command {
outputMsg "\nPMI Representation ---------------------------------------------------------" blue

outputMsg "PMI Representation (aka Semantic PMI) includes all information necessary to represent GD&T without
any graphical presentation elements. PMI Representation is associated with CAD model geometry and
is computer-interpretable to facilitate automated consumption by downstream applications for
manufacturing, measurement, inspection, and other processes.

Worksheets with PMI Representation show a visual recreation of the representation for Dimensional
Tolerances, Geometric Tolerances, and Datum Features.  The results are reported on various entity
worksheets as indicated by 'PMI Representation' on the Summary Worksheet.  The results are in
columns, highlighted in yellow and green, on the relevant worksheets.  The GD&T is recreated as
best as possible given the constraints of Excel.

Dimensional Tolerances are reported on the dimensional_characteristic_representation worksheet.
The dimension name, representation name, length/angle, length/angle name, and plus minus bounds
are reported.  The relevant section in the Recommended Practice is shown in the column headings.
The resulting Dimensional Tolerance is reported in the last column.

Datum Features are reported on datum_* entities.  Datum_system will show the complete Datum
Reference Frame.  Datum Targets are reported on placed_datum_target_feature.

Geometric Tolerances are reported on *_tolerance entities by showing the complete Feature Control
Frame (FCF), and possible Dimensional Tolerance and Datum Feature.  The FCF should contain the
geometry tool, tolerance zone, datum reference frame, and associated modifiers.

If a Dimensional Tolerance refers to the same geometric face as a Geometric Tolerance, then it
will be shown above the FCF.  If a Datum Feature refers to the same geometric face as a Geometric
Tolerance, then it will be shown below the FCF.

The association of the Datum Feature with a Geometric Tolerance is based on each referring to the
same geometric face in the STEP file.  However, the PMI Presentation might show the Geometric
Tolerance and Datum Feature as two separate annotations with leader lines attached to the same
geometric face.

Sometimes an expected association between a Dimensional and Geometric Tolerance is not found.

A Feature Count Modifier, such as ‘8X’ might be displayed in the visual presentation of the PMI
Representation although it might not appear in the PMI Presentation.

All of the Datum Systems, Dimensional Tolerances, and Geometric Tolerances that are reported on
individual worksheets are collected on one PMI Representation Summary worksheet.

Some syntax errors that indicate non-conformance to a CAx-IF Recommended Practices related to PMI
Representation are also reported in the Status tab and the relevant worksheet cells.  Syntax
errors are highlighted in red.

Inverse Relationships are automatically selected when reporting PMI Representation because they
contain information about the relationship between entities used to model PMI Representation.

PMI Representation is defined by the CAx-IF Recommended Practice for:
  Representation and Presentation of Product Manufacturing Information (AP242)
  
Go to Help > Websites > CAx-IF Recommended Practices to access documentation."
  
    .tnb select .tnb.status
    update idletasks
  }
  
  $Help add command -label "PMI Presentation (Graphical PMI)" -command {
outputMsg "\nPMI Presentation -----------------------------------------------------------" blue
outputMsg "PMI Presentation (aka Graphical PMI) consists of geometric elements such as lines and arcs
preserving the exact appearance (color, shape, positioning) of the GD&T annotations.  PMI
Presentation is not intended to be computer-interpretable and does not carry any representation
information, although it can be linked to its corresponding PMI Representation.

Geometric entities used for PMI Presentation annotations are reported in columns, highlighted in
yellow and green, on Annotation_*_occurrence worksheets.  PMI Presentation annotations are used to
specify GD&T (Geometric Dimensioning and Tolerancing).  The Summary worksheet will indicate on the
Annotation_*_occurrence row if PMI Presentation is reported.

Some syntax errors related to PMI Presentation are also reported in the Status tab and the
relevant worksheet cells.  Syntax errors are highlighted in red.

Presentation Style, Saved Views, Validation Properties, Associated Geometry, and Associated
Representation are also reported.

A PMI Presentation Coverage Analysis worksheet is generated.

PMI Presentation annotations can be visualized with an X3DOM (WebGL) file that can be displayed in
most web browsers.  The X3DOM file is only of the annotations, not the model geometry.  The
resulting X3DOM file is named mystepfile_x3dom.html  Tessellated annotations are not supported.

PMI Presentation annotations are found in the STEP file by inspecting the Annotation_*_occurrence
entities and associated geometric_curve_set or annotation_fill_area entities.  Polyline,
trimmed_curve, composite_curve, and line entities used for PMI Presentation are checked.  Circles
are ignored.

PMI Presentation is defined by the CAx-IF Recommended Practices for:
  Representation and Presentation of Product Manufacturing Information (AP242)
  PMI Polyline Presentation (AP203/AP242)
  
Section numbers in the PMI Presentation worksheet refer to the first recommended
practice related to AP242.
  
Go to Help > Websites > CAx-IF Recommended Practices to access documentation."
  
    .tnb select .tnb.status
    update idletasks
  }

# coverage analysis help
  $Help add command -label "Coverage Analysis" -command {
outputMsg "\nCoverage Analysis ----------------------------------------------------------" blue
outputMsg "Coverage Analysis worksheets are generated when processing single or multiple files and when
reports for PMI Representation or Presentation are selected.  

PMI Representation Coverage Analysis (semantic PMI) counts the number of PMI elements found in a
STEP file for tolerances, dimensions, datums, modifiers, and CAx-IF Recommended Practices for PMI
Representation.  On the coverage analysis worksheet, some PMI elements show their associated
symbol, while others show the relevant section in the Recommended Practice.  The PMI elements are
grouped by features related tolerances, dimensions, datums, tolerance zones, common modifiers, and
other modifiers.

PMI Presentation Coverage Analysis (graphical PMI) counts the occurrences of a name attribute
defined in the CAx-IF Recommended Practice for PMI Representation and Presentation of PMI (AP242),
section 8.4, table 14.  The name attribute is associated with the graphic elements used to draw a
PMI annotation."

    .tnb select .tnb.status
    update idletasks

    if {$nistVersion || [file exists [file nativename [file join $mytemp SFA-semantic-coverage.xlsx]]]} {
outputMsg "\n----------
If the STEP file was generated from a NIST CAD model (http://go.usa.gov/mGVm) and the file can be
recognized as having been generated from one of the CAD models, then the PMI Representation
Coverage Analysis worksheet is color-coded by the expected number of PMI elements.  The expected
results were determined by manually counting the number of PMI elements found in each test case
drawing.  Counting of some modifiers, e.g. maximum material condition, does not differentiate by
whether they appear in the tolerance zone definition or datum reference frame.

Green is a match to the expected number of STEP PMI elements.  Yellow means that more were
found.  Red means that less were found.  A magneta cell means that none of an expected PMI element
were found in the STEP file.

Red, magenta, or yellow cells might mean that the CAD system or translator: (1) mapped an internal
PMI element to the wrong STEP PMI element, (2) has not implemented exporting a PMI element to a
STEP file, (3) does not support a type of PMI element, or (4) did not correctly model a PMI
element defined in a test case.  For example, in the STEP file, a required spherical radius might
appear as a radius.  In the coverage analysis worksheet, the missing spherical radius would appear
red or magenta, while the extra radius would appear yellow.

Gray means that a PMI element is in a test case definition but there is no CAx-IF Recommended
Practice to model it.  For example, there is no recommended practice for hole depth,
counterbore, and countersink.  This means that the dimensions of a hole are not associated
with a dimension type such as diameter and radius in the STEP file although the dimension value
is still represented semantically.  This does not affect the PMI Presentation (graphics) for
those PMI elements."
  
      .tnb select .tnb.status
      update idletasks
    }
  }
    
  $Help add command -label "Validation Properties" -command {
outputMsg "\nValidation Properties ------------------------------------------------------" blue
outputMsg "Geometric, assembly, PMI, annotation, attribute, and tessellated validation properties are
reported.  The property values are reported in columns highlighted in yellow and green on the
Property_definition worksheet.  The worksheet can also be sorted and filtered.

Syntax errors related to validation property attribute values are also reported in the Status tab
and the relevant worksheet cells.  Syntax errors are highlighted in red.

Clicking on the plus '+' symbols above the columns will show other columns that contain the entity
ID and attribute name of the validation property value.  All of the other columns can be shown or
hidden by clicking the '1' or '2' in the upper right corner of the spreadsheet.

The Summary worksheet will indicate on the property_definition row if validation properties are
reported.

The number of Cartesian Points reported for smooth or sharp sampling points can also be limited.

Other non-validation properties such as material properties are also reported.

Validation properties are found in the STEP file by inspecting the Representation entities related
to Property_definition through Property_definition_representation.  Property values other than
validation properties are also reported if the relationship is found.

Validation properties are defined by the CAx-IF.
Go to Help > Websites > CAx-IF Recommended Practices to access documentation."
  
    .tnb select .tnb.status
    update idletasks
  }
  $Help add separator

# display files help
  $Help add command -label "Display STEP Files" -command {
outputMsg "\nDisplay STEP Files ---------------------------------------------------------" blue
outputMsg "This option is a convenient way to display a STEP file in other applications. The pull-down menu
will contain applications that can display a STEP file such as STEP viewers, browsers, and
conformance checkers.  If applications are installed in their default location, then they will
appear in the pull-down menu.  

The 'Indent STEP File (for debugging)' option rearranges and indents the entities to show the
hierarchy of information in a STEP file.  The 'indented' file is written to the same directory as
the STEP file or to the same user-defined directory specified in the Spreadsheet tab.  It is
useful for debugging STEP files but is not recommended for large STEP files.

The 'Default STEP Viewer' option will open the STEP file in whatever application is associated
with STEP files.  A text editor will always appear in the menu."

    .tnb select .tnb.status
    update idletasks
  }

# multiple files help
  $Help add command -label "Multiple STEP Files" -command {
outputMsg "\nMultiple STEP Files --------------------------------------------------------" blue
outputMsg "Multiple STEP files can be selected in the Open File(s) dialog by holding down the control or
shift key when selecting files or an entire directory of STEP files can be selected with 'Open
Multiple STEP Files in a Directory'.  Files in subdirectories of the selected directory can also
be processed.

When processing multiple STEP files, a File Summary spreadsheet is generated in addition to
individual spreadsheets for each file.  The File Summary spreadsheet shows the entity count and
totals for all STEP files. The File Summary spreadsheet also links to the individual spreadsheets
and the STEP file.

If only the File Summary spreadsheet is needed, it can be generated faster by turning off
Processing of most of the entity types and options in the Options tab.

If the reports for PMI Representation or Presentation are selected, then Coverage Analysis
worksheets are also generated."

    .tnb select .tnb.status
    update idletasks
  }

# number format help
  $Help add command -label "Number Format" -command {
outputMsg "\nNumber Format --------------------------------------------------------------" blue
outputMsg "Excel rounds real numbers if there are more than 11 characters in the number string.  For example,
the number 0.12499999999999997 in the STEP file will be displayed as 0.125  However, double
clicking in a cell with a rounded number will show all of the digits.

This option will display most real numbers exactly as they appear in the STEP file.  This applies
only to single real numbers.  Lists of real numbers, such as cartesian point coordinates, are
always displayed exactly as they appear in the STEP file."

    .tnb select .tnb.status
    update idletasks
  }
  $Help add separator
  
  $Help add command -label "Conformance Checking" -command {
outputMsg "\nConformance Checking -------------------------------------------------------" blue
outputMsg "STEP AP203 and AP214 files can be checked for conformance with the free ST-Developer Personal
Edition.  If installed, then it will show up in the 'Display STEP File in' pull-down menu in the
Options tab. A spreadsheet does not have be generated to run it.  If the option to 'Write results
to a file' is selected, then the output file will be named mystepfile_stdev.log  To download
ST-Developer Personal Edition, go to: http://www.steptools.com/products/stdev/personal.html  This
software also includes some useful STEP utility programs.

To check for some missing entity references, in the Options tab under Display STEP File, use
Indent STEP File (for debugging)."
  
    .tnb select .tnb.status
    update idletasks
  }
  
  $Help add command -label "Other STEP APs" -command {
outputMsg "\nOther STEP APs -------------------------------------------------------------" blue
outputMsg "Processing of STEP files from other STEP APs can be enabled by installing the free ST-Developer
Personal Edition.  To download ST-Developer Personal Edition, go to:
http://www.steptools.com/products/stdev/personal.html  This software also includes some useful
STEP utility programs."
  
    .tnb select .tnb.status
    update idletasks
  }

# large files help
  $Help add command -label "Large STEP Files" -command {
outputMsg "\nLarge STEP Files -----------------------------------------------------------" blue
outputMsg "To reduce the amount of time to process large STEP files and to reduce the size of the resulting
spreadsheet, several options are available:
- In the Process section, deselect entity types Geometry and Cartesian Point
- In the Spreadsheet tab, select a value for the Maximum Rows
- In the Options tab, deselect Reports and Inverse Relationships

The STEP File Analyzer might also crash when processing very large STEP files.  Popup dialogs
might appear that say \"unable to alloc xxx bytes\".  In this case, deselect some entity types to
process in Options tab or use a User-Defined List of entities to process.  See the Help for Crash
Recovery."

    .tnb select .tnb.status
    update idletasks
  }

  $Help add command -label "Crash Recovery" -command {
outputMsg "\nCrash Recovery -------------------------------------------------------------" red
outputMsg "Sometimes the STEP File Analyzer will crash after a STEP file has been successfully opened and the
processing of entities has started.  Popup dialogs might appear that say \"ActiveState Basekit has
stopped working\". 

A crash is most likely due to syntax errors in the STEP file or sometimes due to limitations of
the toolkit used to read STEP files.  To see which type of entity caused the error, check the
Status tab to see which type of entity was last processed.

Workarounds for this problem:

- The program keeps track of the last entity type processed when it crashed.  Simply restart the
STEP File Analyzer and hit F1 to process the last STEP file or F4 if processing multiple files.
The type of entity that caused the crash will be skipped.  The list of bad entity types that will
not be processed is stored in a file myfile_fix.dat.  If syntax errors related to the bad entities
are corrected, then delete the *_fix.dat file so that the corrected entities are processed.  When
the STEP file is processed, the list of specific entities that are not processed is reported.

- Deselect Inverse Relationships and all Reports in the Options tab.  If one of these features
caused the crash, then the *_fix.dat file is still created as described above and might need to be
deleted.

- Processing of the type of entity that caused the error can be deselected in the Options tab
under Process.  However, this will prevent processing of other entities that do not cause a crash.

The STEP File Analyzer might also crash when processing very large STEP files.  Popup dialogs
might appear that say \"unable to alloc xxx bytes\".  In this case, deselect some entity types to
process in Options tab or use a User-Defined List of entities to process."

  .tnb select .tnb.status
  update idletasks
}
  
  $Help add command -label "Errors" -command {
outputMsg "\nErrors ---------------------------------------------------------------------" blue
outputMsg "If sufficient memory is not available to process a very large STEP file, then the STEP File
Analyzer will stop with an error message that might say \"Fatal Error in Wish - unable to alloc
123456 bytes\".  Try processing the STEP file on a computer with more memory or deselect some
categories of entities to Process in the Options tab.

After stopping the program when a large STEP file has been processed, sometimes the
STEP-File-Analyzer.exe and/or EXCEL.EXE processes will still be running.  The Windows Task Manager
can be used to kill those processes."

    .tnb select .tnb.status
    update idletasks
  }

  $Help add separator
  if {"$nistVersion"} {
    $Help add command -label "Disclaimer" -command {displayDisclaimer}
    $Help add command -label "NIST Disclaimer" -command {displayURL http://www.nist.gov/public_affairs/disclaimer.cfm}
  }
  $Help add command -label "About" -command {
    outputMsg "\nSTEP File Analyzer ---------------------------------------------------------" blue
    outputMsg "Version:  [getVersion]"
    outputMsg "Updated:  [string trim [clock format $progtime -format "%e %b %Y"]]"
    if {"$nistVersion"} {
      outputMsg "Contact:  Robert Lipman, robert.lipman@nist.gov"
      outputMsg "\nThe STEP File Analyzer was first released in 2012."
    } else {
      outputMsg "\nThis version was built from the NIST STEP File Analyzer source\ncode available on GitHub.  https://github.com/usnistgov/SFA"
    }
  
  # debug
    if {$opt(ROWLIM) == 100003} {
      outputMsg "\nDebug Messages below" red
      foreach id [lsort [array names env]] {outputMsg "$id   $env($id)"}
      outputMsg \nUSERPROFILE-$env(USERPROFILE)
      catch {outputMsg registry_personal-[registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Personal}]}
      catch {outputMsg registry_desktop-[registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]}
      outputMsg \ndrive-$drive\nmyhome-$myhome\nmydocs-$mydocs\nmytemp-$mytemp
      catch {outputMsg mydesk-$mydesk}
      catch {outputMsg mymenu-$mymenu}
      outputMsg programfiles-$programfiles\npf64-$pf64
      outputMsg os-$tcl_platform(os)-$tcl_platform(osVersion)
      outputMsg [::twapi::get_os_description]
      outputMsg [::twapi::get_os_version]
      outputMsg "Debug Messages above" red
    }
  
    .tnb select .tnb.status
    update idletasks
  }
}

#-------------------------------------------------------------------------------
# Websites menu
proc guiWebsitesMenu {} {
  global Websites

  $Websites add command -label "STEP File Analyzer"     -command {displayURL http://www.nist.gov/el/msid/infotest/step-file-analyzer.cfm}
  $Websites add command -label "Source code on GitHub"  -command {displayURL https://github.com/usnistgov/SFA}
  $Websites add command -label "Journal of CAD Article" -command {displayURL http://www.nist.gov/manuscript-publication-search.cfm?pub_id=917105}
  $Websites add command -label "MBE PMI Validation Testing (free CAD models and STEP files)" -command {displayURL http://www.nist.gov/el/msid/infotest/mbe-pmi-validation.cfm}
  $Websites add command -label "Enabling the Digital Thread for Smart Manufacturing"         -command {displayURL http://www.nist.gov/el/msid/infotest/digital-thread-manufacturing.cfm}
  
  $Websites add separator
  $Websites add command -label "CAx Implementor Forum (CAx-IF)" -command {displayURL https://www.cax-if.org/}
  $Websites add command -label "Implementation Coverage"        -command {displayURL https://www.cax-if.org/vendor_info.php}
  $Websites add command -label "STEP File Viewers"              -command {displayURL https://www.cax-if.org/step_viewers.html}
  $Websites add command -label "STEP File Library"              -command {displayURL https://www.cax-if.org/library/index.html}
  $Websites add command -label "Recommended Practices"          -command {displayURL https://www.cax-if.org/joint_testing_info.html#recpracs}
  
  $Websites add separator
  $Websites add command -label "STEP AP242 Project" -command {displayURL http://www.ap242.org/}
  $Websites add command -label "EXPRESS Schemas"    -command {displayURL https://www.cax-if.org/joint_testing_info.html#schemas}
  $Websites add command -label "AP242 Schema"       -command {displayURL http://www.steptools.com/support/stdev_docs/express/ap242/html/index.html}
  $Websites add command -label "AP203e2 Schema"     -command {displayURL http://www.steptools.com/support/stdev_docs/express/ap203e2/html/index.html}
  $Websites add command -label "AP214e3 Schema"     -command {displayURL http://www.steptools.com/support/stdev_docs/express/ap214/html/index.html}
  
  $Websites add separator
  $Websites add command -label "PDES, Inc."   -command {displayURL https://pdesinc.org/}
  $Websites add command -label "ProSTEP iViP" -command {displayURL http://www.prostep.org/en/projects.html}
  $Websites add command -label "LOTAR"        -command {displayURL http://www.lotar-international.org/}
  
  $Websites add separator
  $Websites add cascade -label "Research" -menu $Websites.3
  set Websites3 [menu $Websites.3 -tearoff 1]
  $Websites3 add command -label "An ISO STEP Tolerancing Standard as an Enabler of Smart Manufacturing Systems"  -command {displayURL http://www.nist.gov/manuscript-publication-search.cfm?pub_id=915430}
  $Websites3 add command -label "Standardized STEP Composite Structure Design and Manufacturing Information"     -command {displayURL http://www.nist.gov/manuscript-publication-search.cfm?pub_id=913466}
  $Websites3 add command -label "A Strategy for Testing Product Conformance to GD&T Standards"                   -command {displayURL http://www.nist.gov/manuscript-publication-search.cfm?pub_id=911123}
  $Websites3 add command -label "The Role of Science in the Evolution of Dimensioning and Tolerancing Standards" -command {displayURL http://www.nist.gov/manuscript-publication-search.cfm?pub_id=910523}
  $Websites3 add command -label "Model Based Enterprise for Manufacturing"                                       -command {displayURL http://www.nist.gov/manuscript-publication-search.cfm?pub_id=908343}
  $Websites3 add command -label "MBE Standardization and Validation"                                             -command {displayURL http://www.nist.gov/manuscript-publication-search.cfm?pub_id=908106}
}

#-------------------------------------------------------------------------------

proc displayDisclaimer {} {
  global nistVersion

  if {$nistVersion} {
set txt "This software was developed at the National Institute of Standards and Technology by employees of the Federal Government in the course of their official duties. Pursuant to Title 17 Section 105 of the United States Code this software is not subject to copyright protection and is in the public domain.  This software is an experimental system.  NIST assumes no responsibility whatsoever for its use by other parties, and makes no guarantees, expressed or implied, about its quality, reliability, or any other characteristic.

Any mention of commercial products or references to web pages in this software is for information purposes only; it does not imply recommendation or endorsement by NIST.  For any of the web links in this software, NIST does not necessarily endorse the views expressed, or concur with the facts presented on those web sites.

This software uses Microsoft Excel and IFCsvr that are covered by their own EULAs (End-User License Agreements)."
  
    tk_messageBox -type ok -icon info -title "Disclaimers for STEP File Analyzer" -message $txt
  }
}

#-------------------------------------------------------------------------------
# crash recovery dialog
proc displayCrashRecovery {} {

set txt "Sometimes the STEP File Analyzer will crash AFTER a file has been successfully opened and the processing of entities has started.

A crash is most likely due to syntax errors in the STEP file or sometimes due to limitations of the toolkit used to read STEP files.

If this happens, simply restart the STEP File Analyzer and process the same STEP file again by using function key F1 or F4 if processing multiple STEP files.  Also deselect, Reports and Inverse Relationships in the Options tab.

The STEP File Analyzer keeps track of which entity type caused the error for a particular STEP file and won't process that type again.  The bad entities types are stored in a file *_fix.dat  If syntax errors related to the bad entities are corrected, then delete the *_fix.dat file so that the corrected entities are processed.

The software might also crash when processing very large STEP files.  In this case, deselect some entity types to process in Options tab or use a User-Defined List of entities to process.

More details about recovering from a crash are explained in the User's Guide and in Help > Crash Recovery

Please report other types of crashes to the software developer."
  
  tk_messageBox -type ok -icon error -title "What to do if the STEP File Analyzer crashes?" -message $txt
}

#-------------------------------------------------------------------------------
# display user guide
proc displayGuide {} {
  
  set ugName [file nativename [file join [file dirname [info nameofexecutable]] SFA-Users-Guide.pdf]]
  if {[file exists $ugName]} {
    exec {*}[auto_execok start] "" $ugName
  } else {
    displayURL http://dx.doi.org/10.6028/NIST.IR.8122
  }
}
 
#-------------------------------------------------------------------------------
proc processToolTip {ttmsg tt {ttlim 120}} {
  global entCategory
 
  set ttlen 0
  set lchar ""
  set r1 0
  if {$tt == "PR_STEP_OTHER"} {set ttlim 160}

  foreach item [lsort $entCategory($tt)] {
    if {[string range $item 0 $r1] != $lchar && $lchar != ""} {
      if {[string index $ttmsg end] != "\n"} {append ttmsg "\n"}
      set ttlen 0
    }
    append ttmsg "$item   "
    incr ttlen [string length $item]
    if {$ttlen > $ttlim} {
      if {[string index $ttmsg end] != "\n"} {append ttmsg "\n"}
      set ttlen 0
      set ok 0
    }
    set lchar [string range $item 0 $r1]
  }
  return $ttmsg
}
