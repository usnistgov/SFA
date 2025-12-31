# generate an Excel spreadsheet and/or view from a STEP file
proc genExcel {{numFile 0}} {
  global allEntity aoEntTypes ap203all ap214all ap242all ap242only ap242ed badAttributes brepGeomEntTypes buttons cadSystem cameraModels cells cells1
  global col col1 commaSeparator count csvdirnam csvfile csvinhome currLogFile developer dim draughtingModels driUnicode entCategories entCategory
  global entColorIndex entCount entityCount entsIgnored entsWithErrors env epmi epmiUD errmsg equivUnicodeStringErr excel fcsv
  global feaFirstEntity feaLastEntity File fileEntity fileItems filesProcessed fileSumRow gen gpmiTypesInvalid gpmiTypesPerFile guiSFA idRow idxColor
  global ifcsvrDir ifcsvrVer iloldscr inverses lastXLS lenfilelist localName localNameList logFile matrixList multiFile multiFileDir mydocs mytemp nistCoverageLegend
  global nistName nistPMIexpected nistPMImaster noFontFile nprogBarEnts opt pf32 p21e3Section pmiCol resetRound roseSchemas row rowmax savedViewButtons
  global savedViewName savedViewNames scriptName sheetLast skipEntities skipFileName spmiEntity spmiSumName spmiSumRow spmiTypesPerFile
  global startrow statsOnly stepAP stepAPreport sumHeaderRow syntaxErr tessBrep tessColor tessEnts tessMinMax tessSolid thisEntType timeStamp tlast tolNames tolStandard tolStandards
  global totalEntity unicodeActual unicodeAttr unicodeAttributes unicodeEnts unicodeInFile unicodeNumEnts unicodeString unicodeStringCM userEntityFile userEntityList
  global userWriteDir useXL uuidCount uuidEnts valRounded viz wdir workbook workbooks worksheet worksheet1 worksheets writeDir wsCount wsNames x3dAxes
  global x3dColor x3dColorFile x3dColors x3dFileName x3dIndex x3dMax x3dMin x3dMsg x3dMsgColor x3dStartFile x3dViewOK xlFileName xlFileNames xlInstalled xlsManual
  global objDesign

# check a few variables
  set writeDir $userWriteDir
  if {![info exists gen(Excel)]} {set gen(Excel) 1}
  set guiSFA 1
  if {[string first "sfa-cl.exe" $scriptName] != -1} {set guiSFA 0}
  if {[info exists errmsg]} {set errmsg ""}

# check for permissions to write to the same directory as the STEP file
  set dirname [file dirname $localName]
  if {$opt(writeDirType) == 2} {set dirname $writeDir}
  if {[string first $mydocs [file nativename $dirname]] == -1} {
    if {[catch {
      set tfile [file join $dirname test.txt]
      set tf [open $tfile w]
      close $tf
      catch {file delete -force -- $tfile}
    } emsg]} {
      set msg "Error opening Output files in: $dirname"
      if {[string first "permission denied" $emsg] != -1} {
        append msg "\n Copy the STEP file to a different directory or write the Output to a different User-defined directory (More tab)"
      }
      errorMsg $msg
      .tnb select .tnb.status
      return
    }
  }

# generate STEP AP242 tessellated geometry from STL file
  if {[string tolower [file extension $localName]] == ".stl"} {
    STL2STEP
    if {$localName == ""} {return}
  }

# -------------------------------------------------------------------------------------------------
# initialize for x3dom geometry
  set x3dViewOK 0
  set x3dMsg {}
  set tessBrep 0
  set tessMinMax 0
  if {$gen(View)} {
    set x3dMsgColor green
    if {$opt(viewPMI) || $opt(viewTessPart) || $opt(viewFEA) || $opt(viewPart)} {
      set tessEnts 0
      set x3dStartFile 1
      set x3dAxes 1
      set x3dFileName ""
      set x3dColor ""
      set x3dColors {}
      foreach idx {x y z} {set x3dMax($idx) -1.e8; set x3dMin($idx) 1.e8}
      catch {unset tessColor}
      catch {unset tessSolid}
      catch {unset x3dColorFile}
    }
  }

# multiFile
  set multiFile 0
  if {$numFile > 0} {set multiFile 1}

  if {[info exists buttons]} {
    $buttons(generate) configure -state disabled
    .tnb select .tnb.status
  }
  set lasttime [clock clicks -milliseconds]

# open log file
  if {$opt(logFile)} {
    set lfile [file rootname $localName]
    if {$opt(writeDirType) == 2} {set lfile [file join $writeDir [file rootname [file tail $localName]]]}
    append lfile "-sfa.log"
    set logFile [open $lfile w]
    puts $logFile "NIST STEP File Analyzer and Viewer [getVersion] ([string trim [clock format [clock seconds]]])"
  }

# -------------------------------------------------------------------------------------------------
# view part geometry only, does not require opening STEP file with IFCsvr
  if {$opt(partOnly) && ![info exists statsOnly]} {
    outputMsg "\nGenerating View"
    foreach var {cadSystem entCount stepAP timeStamp} {if {[info exists $var]} {unset -- $var}}

# add file name to menu
    set ok 0
    if {$numFile <= 1} {set ok 1}
    if {[info exists localNameList]} {if {[llength $localNameList] > 1} {set ok 1}}
    if {$ok} {addFileToMenu}

# initialize
    foreach idx {COMPOSITES DTMTAR EDGE FEA HOLE PMI PLACE POINTS SUPPGEOM TESSEDGE TESSPART} {set viz($idx) 0}
    set viz(PART) 1
    set x3dMsgColor blue
    set lasttime [clock clicks -milliseconds]

# generate x3d
    x3dFileStart

# done
    if {$x3dViewOK} {
      x3dFileEnd

# processing time
      set cc [clock clicks -milliseconds]
      set proctime [expr {($cc - $lasttime)/1000}]
      if {$proctime <= 60} {set proctime [expr {(($cc - $lasttime)/100)/10.}]}
      outputMsg "Processing time: $proctime seconds"

# view file
      if {$viz(PART)} {openX3DOM "" $numFile}

# save log file
      if {[info exists logFile]} {
        saveLogFile $lfile
        unset lfile
      }

# clean up and return
      incr filesProcessed
      saveState
    }

    foreach var {cadSystem stepAP timeStamp x3dCoord x3dFile x3dFileName x3dIndex x3dMax x3dMin x3dStartFile} {if {[info exists $var]} {unset -- $var}}
    if {[info exists buttons]} {$buttons(generate) configure -state normal}
    return
  }

# -------------------------------------------------------------------------------------------------
# check if IFCsvr is installed
  if {![info exists ifcsvrDir]} {set ifcsvrDir [file join $pf32 IFCsvrR300 dll]}
  if {![file exists [file join $ifcsvrDir IFCsvrR300.dll]]} {
    if {[info exists buttons]} {$buttons(generate) configure -state disabled}
    installIFCsvr
    return
  }

# run syntax checker too
  if {$opt(syntaxChecker) && [info exists buttons]} {syntaxChecker $localName}

# -------------------------------------------------------------------------------------------------
# connect to IFCsvr, cannot redirect createobject output to suppress it
  if {[catch {
    outputMsg "Connecting to IFCsvr"
    update idletasks
    set objIFCsvr [::tcom::ref createobject IFCsvr.R300]

# set environment variable that is sometimes necessary
    set roseSchemas ""
    catch {if {$env(ROSE_SCHEMAS) != $ifcsvrDir} {set roseSchemas $env(ROSE_SCHEMAS)}}
    set env(ROSE_SCHEMAS) [file nativename $ifcsvrDir]

# error
  } emsg]} {
    errorMsg "Error connecting to the IFCsvr toolkit that is used to read STEP files: $emsg"
    if {[string first "could not be found" $emsg] != -1} {
      catch {
        if {[string first "." [registry get $ifcsvrVer {DisplayVersion}]] != -1} {
          errorMsg " Manually uninstall the IFCsvr toolkit, then run this software to reinstall it."
        }
      }
    }
    catch {raise .}
    return 0
  }

# -------------------------------------------------------------------------------------------------
# open STEP file
  if {[catch {
    set openStage 1
    set nprogBarEnts 0
    set fname $localName
    set stepAP [getStepAP $fname]
    if {$stepAP == ""} {return}
    for {set i 1} {$i < 10} {incr i} {set ap242ed($i) {}}

# stepAPreport controls which APs support Analyzer reports
    set stepAPreport 0
    set ap [string range $stepAP 0 4]
    if {$ap == "AP203" || $ap == "AP209" || $ap == "AP210" || $ap == "AP214" || $ap == "AP238" || $ap == "AP242"} {set stepAPreport 1}

# check for Part 21 edition 3 files and strip out sections
    set fname [checkP21e3 $fname]

# add file name and size to multi file summary
    if {$numFile != 0 && [info exists cells1(Summary)]} {
      set dlen [expr {[string length [truncFileName $multiFileDir]]+1}]
      set fn [string range [file nativename [truncFileName $fname]] $dlen end]
      set fn1 [split $fn "\\"]
      set fn2 [lindex $fn1 end]
      set idx [string first $fn2 $fn]
      if {[string length $fn2] > 40} {
        set div [expr {int([string length $fn2]/2)}]
        set fn2 [string range $fn2 0 $div][format "%c" 10][string range $fn2 [expr {$div+1}] end]
        set fn  [file nativename [string range $fn 0 $idx-1]$fn2]
      }
      regsub -all {\\} $fn [format "%c" 10] fn

      set colsum [expr {$col1(Summary)+1}]
      set range [$worksheet1(Summary) Range [cellRange 4 $colsum]]
      $cells1(Summary) Item 4 $colsum $fn
    }

# open file with IFCsvr
    set str "STEP"
    if {[string first "AP" $stepAP] == 0} {
      set str "STEP [string range $stepAP 0 4]"
    } elseif {$stepAP != ""} {
      set str $stepAP
    }
    outputMsg "\nOpening $str file"

    set openStage 2
    if {![info exists buttons]} {outputMsg "\n<Reading STEP file and checking for syntax errors>"}
    set objDesign [$objIFCsvr OpenDesign [file nativename $fname]]
    if {![info exists buttons]} {
      update idletasks
      outputMsg "<Done>\n"
    }

# CountEntities causes the error if the STEP file cannot be opened because objDesign is null
    set entityCount [$objDesign CountEntities "*"]

# get stats
    set openStage 3
    if {![info exists gen(CSV)]} {set gen(CSV) 0}
    if {$entityCount > 0} {
      outputMsg " $entityCount entities"
      set entityTypeNames [$objDesign EntityTypeNames [expr 2]]

# set which types of characteristics are in the file
      set characteristics {}
      foreach entType $entityTypeNames {
        set ecount [$objDesign CountEntities "$entType"]
        if {$ecount > 0} {

# complex entities
          set c1 [string first "_and_" $entType]
          set ent1 ""
          set ent2 ""
          if {$c1 != -1} {
            set ent1 [string range $entType 0 $c1-1]
            set ent2 [string range $entType 0 $c1+5]
          }

          if {$entType == "dimensional_characteristic_representation" || $entType == $iloldscr} {
            lappend characteristics "Dimensions"
          } elseif {$entType == "datum"} {
            lappend characteristics "Datums"
          } elseif {$entType == "placed_datum_target_feature" || $entType == "datum_target"} {
            lappend characteristics "Datum targets"

          } elseif {$entType == "tessellated_annotation_occurrence"} {
            lappend characteristics "Graphic PMI (tessellated)"
          } elseif {$entType == "annotation_occurrence" || [string first "annotation_curve_occurrence" $entType] != -1 || \
                    $entType == "annotation_fill_area_occurrence" || $entType == "annotation_occurrence_and_characterized_object"} {
            lappend characteristics "Graphic PMI (polyline)"
          } elseif {$entType == "annotation_placeholder_occurrence" || $entType == "annotation_placeholder_occurrence_with_leader_line"} {
            lappend characteristics "Placeholder PMI"
          } elseif {$entType == "external_image_placement_in_callout"} {
            lappend characteristics "External image"

          } elseif {$entType == "constructive_geometry_representation"} {
            lappend characteristics "Supplemental geometry"
          } elseif {$entType == "tessellated_constructive_geometry_representation"} {
            lappend characteristics "Supplemental geometry (tessellated)"
          } elseif {$entType == "property_definition_representation" || $entType == "material_property_representation"} {
            lappend characteristics "Properties"

          } elseif {[lsearch $entCategory(stepCOMP) $entType] != -1} {
            lappend characteristics "Composites"
          } elseif {[lsearch $entCategory(stepKINE) $entType] != -1} {
            lappend characteristics "Kinematics"
          } elseif {[lsearch $entCategory(stepFEAT) $entType] != -1 || [lsearch $entCategory(stepFEAT) $ent1] != -1 || [lsearch $entCategory(stepFEAT) $ent2] != -1} {
            lappend characteristics "Features"
          } else {
            foreach tol $tolNames {if {[string first $tol $entType] != -1} {lappend characteristics "Geometric tolerances"}}
          }
          if {[string first "uuid" $entType] != -1} {lappend characteristics "UUIDs"}

# make sure some entity types are always processed
          if {$opt(xlFormat) != "None"} {
            foreach cat {stepCOMP stepKINE stepFEAT stepAP242 stepQUAL stepCONS stepOTHR} {
              if {$opt($cat) == 0 && [lsearch $entCategory($cat) $entType] != -1} {
                set opt($cat) 1
                checkValues
              }
            }
          }
        }

# check for entities in different editions of AP242
        for {set i 1} {$i < 9} {incr i} {
          if {[info exists ap242only(e$i)]} {if {[lsearch $ap242only(e$i) $entType] != -1} {lappend ap242ed($i) $entType}}
        }
      }
      if {[llength $characteristics] > 0} {set characteristics [lrmdups $characteristics]}

# check for type of part geometry
      catch {
        set str "Part geometry ("
        set gtypes {}
        foreach entType [list closed_shell open_shell surface_of_linear_extrusion surface_of_revolution tessellated_solid tessellated_shell tessellated_closed_shell tessellated_open_shell] {
          set num [$objDesign CountEntities "$entType"]
          if {$num > 0} {
            if {$entType == "closed_shell"} {lappend gtypes "B-rep solid"}
            if {$entType == "open_shell" || [string first "surface_of_" $entType] != -1} {lappend gtypes "B-rep surface"}
            if {$entType == "tessellated_solid"} {lappend gtypes "tessellated solid"}
            if {$entType == "tessellated_shell"} {lappend gtypes "tessellated surface"}
            if {$entType == "tessellated_closed_shell" || $entType == "tessellated_open_shell"} {lappend gtypes "polyhedral B-rep"}
          }
        }
        if {[llength $gtypes] > 0} {
          set gtypes "Part geometry ([join [lrmdups $gtypes] ", "])"
          set characteristics [concat [list $gtypes] $characteristics]
        }
      }

# report characteristics
      set fileItems ""
      if {[llength $characteristics] > 0} {
        set str  ""
        set str1 ""
        foreach item $characteristics {
          append str1 "$item, "
          if {[string length $str1] > 100} {
            append str " $str1\n"
            set str1 ""
          }
        }
        if {[string length $str1] > 0} {append str " $str1"}
        set str [string trim $str]
        if {[string index $str end] == ","} {set str [string range $str 0 end-1]}
        if {$str != "Part geometry (B-rep solid)"} {
          if {![info exists buttons]} {outputMsg \n}
          outputMsg "\nThis file contains:" blue
          outputMsg " $str"
          set fileItems $str

          if {[string range $stepAP 0 4] == "AP203" || [string range $stepAP 0 4] == "AP214"} {
            foreach item [list Datums Dimensions "Geometric tolerances" "Graphic PMI"] {
              if {[string first $item $fileItems] != -1} {
                set msg " Graphic and Semantic PMI is best supported in AP242"
                if {$gen(Excel) && ($opt(PMISEM) || $opt(PMIGRF))} {append msg  ".  Analyzer reports are based on AP242 Recommended Practices."}
                errorMsg $msg red
              }
            }
          }
        }
      }
    } else {
      errorMsg "The number of entities could not be counted or there are no entities in the STEP file.\n See Examples menu for sample STEP files.\n See Help > Syntax Checker"
    }
    outputMsg " "

# exit if stats only from command-line version
    if {[info exists statsOnly]} {
      if {[info exists logFile]} {saveLogFile $lfile}
      if {$roseSchemas != "" && $roseSchemas != $ifcsvrDir} {
        unset env(ROSE_SCHEMAS)
        if {$roseSchemas != ""} {set env(ROSE_SCHEMAS) $roseSchemas}
      }
      exit
    }

# -------------------------------------------------------------------------------------------------
# add AP, file size, entity count to multi file summary
    if {$numFile != 0 && [info exists cells1(Summary)]} {
      set ap $stepAP

# fix ISO13584 schemas to fit
      if {[string first "ISO13584" $ap] == 0} {
        foreach str {"_2_LONG_FORM_SCHEMA" "_IEC61360_5_LIBRARY_IMPLICIT_SCHEMA"} {
          set c1 [string first $str $ap]
          if {$c1 != -1} {set ap [string range $ap 0 $c1-1]}
        }
      }
      if {$ap == "CUTTING_TOOL_SCHEMA_ARM"} {set ap "ISO13399"}

      $cells1(Summary) Item [expr {$startrow-2}] $colsum $ap
      $cells1(Summary) Item [expr {$startrow-1}] $colsum [fileSize $fname]
      $cells1(Summary) Item $startrow $colsum $entityCount
    }

# -------------------------------------------------------------------------------------------------
# open file of entities (-skip.dat) not to process (skipEntities)
    set skipEntities {}
    set skipFileName [file rootname $fname]
    if {$opt(writeDirType) == 2} {set skipFileName [file join $writeDir [file rootname [file tail $fname]]]}
    append skipFileName "-skip.dat"
    if {[file exists $skipFileName]} {
      set skipFile [open $skipFileName r]
      while {[gets $skipFile line] >= 0} {
        if {[lsearch $skipEntities $line] == -1 && $line != "" && ![info exists badAttributes($line)]} {
          lappend skipEntities $line
        }
      }
      close $skipFile
    }
    catch {if {$developer && ([string toupper $cadSystem] == "CREO" || $cadSystem == "Pro/E")} {set badAttributes(presentation_style_assignment) {styles}}}

# check if a file generated from a NIST test case (and some other files) is being processed
    set nistName [nistGetName]

# error opening file
  } emsg]} {
    if {$roseSchemas != "" && $roseSchemas != $ifcsvrDir} {
      unset env(ROSE_SCHEMAS)
      if {$roseSchemas != ""} {set env(ROSE_SCHEMAS) $roseSchemas}
    }

    if {$openStage == 2} {
      errorMsg "Error opening STEP file: $emsg"

      set fext [string tolower [file extension $fname]]
      if {$fext != ".stp" && $fext != ".step" && $fext != ".p21" && $fext != ".stpz" && $fext != ".stpa" && $fext != ".ifc"} {
        if {$fext != ""} {errorMsg " File extension '[file extension $fname]' is not supported." red}
      } else {
        set fs [getSchemaFromFile $fname 1]
        set c1 [string first "\{" $fs]
        if {$c1 != -1} {set fs [string trim [string range $fs 0 $c1-1]]}

# check for a bad schema
        set okSchema 0
        foreach match [lsort [glob -nocomplain -directory $ifcsvrDir *.rose]] {
          set schema [string toupper [file rootname [file tail $match]]]
          if {$fs == $schema} {set okSchema 1; break}
        }
        if {!$okSchema} {
          if {[string first "," $fs] != -1} {
            set msg "\nMultiple schemas are not supported: $fs"
          } elseif {[string first "_MIM" $fs] != -1 && [string first "_MIM_LF" $fs] == -1} {
            set msg "\nThe STEP AP (schema) should end with _MIM_LF: $fs"
          } else {
            set msg "\nThe STEP AP (schema) is not supported: $fs\n Check View and Part Only to use the Viewer"
          }
          if {[info exists buttons]} {append msg "\n See Help > Supported STEP APs"}
          errorMsg $msg red

# other possible errors
        } else {
          set msg "\nPossible causes of the error:"
          if {[file size $localName] > 429000000} {append msg "\n- The STEP file is too large to open to generate a Spreadsheet.  The limit is about 430 MB.\n   For the Viewer use Part Only on the Generate tab."}
          append msg "\n- Syntax errors in the STEP file"
          append msg "\n   Use F8 to run the Syntax Checker to check for errors in the STEP file.  See Help > Syntax Checker"
          append msg "\n   Try opening the file in another STEP viewer.  See Websites > STEP > STEP File Viewers"
          append msg "\n- File or directory name contains accented, non-English, or symbol characters."
          append msg "\n   [file nativename $fname]"
          append msg "\n   Change the file name or directory name"
          errorMsg $msg red
        }
      }

      if {[info exists errmsg]} {unset errmsg}
      catch {$objDesign Delete}
      catch {unset objDesign}
      catch {unset objIFCsvr}
      catch {raise .}
      return 0

# other errors
    } elseif {$openStage == 1} {
      if {$emsg != ""} {errorMsg "Error before opening STEP file: $emsg"}
      return
    } elseif {$openStage == 3} {
      errorMsg "Error after opening STEP file: $emsg"
    }
  }

# -------------------------------------------------------------------------------------------------
# connect to Excel
  set useXL 1
  set xlInstalled 1
  set csvinhome 0
  if {$opt(xlFormat) != "None"} {
    if {[catch {
      set pid1 [checkForExcel $multiFile]
      set xlapp "Excel.Application"
      if {$opt(debugNOXL)} {set xlapp "null"}
      set excel [::tcom::ref createobject $xlapp]
      set pidExcel [lindex [intersect3 $pid1 [twapi::get_process_ids -name "EXCEL.EXE"]] 2]
      [$excel ErrorCheckingOptions] TextDate False
      set excelVersion [expr {int([$excel Version])}]

# file format, max rows
      set extXLS "xlsx"
      set xlFormat [expr 51]
      set rowmax [expr {2**20}]

# generate with Excel but save as CSV
      set saveCSV 0
      if {$opt(xlFormat) == "CSV"} {
        set saveCSV 1
        catch {$buttons(genExcel) configure -state disabled}
      } else {
        catch {$buttons(genExcel) configure -state normal}
      }

# turning off ScreenUpdating
      $excel Visible 0
      catch {$excel ScreenUpdating 0}

      set rowmax [expr {$rowmax-2}]
      if {$opt(xlMaxRows) < $rowmax} {set rowmax $opt(xlMaxRows)}

# no Excel, use CSV instead
    } emsg]} {
      set useXL 0
      set xlInstalled 0
      if {$opt(xlFormat) == "Excel"} {
        errorMsg "Excel is not installed or cannot be started: $emsg\n CSV files will be generated instead of a spreadsheet.  Analyzer options are disabled."
        outputMsg " "
        set opt(xlFormat) "CSV"
        catch {raise .}
      }
      set gen(Excel) 0
      set gen(CSV) 1
      checkValues
      catch {$buttons(genExcel) configure -state disabled}
    }

# view only
  } else {
    set useXL 0
  }

# -------------------------------------------------------------------------------------------------
# start worksheets
  if {$useXL} {
    if {[catch {
      set workbooks  [$excel Workbooks]
      set workbook   [$workbooks Add]
      set worksheets [$workbook Worksheets]

# load custom color theme that only changes the hyperlink color
      catch {
        file copy -force -- [file join $wdir images SFA-excel-theme.xml] [file join $mytemp SFA-excel-theme.xml]
        [[[$excel ActiveWorkbook] Theme] ThemeColorScheme] Load [file nativename [file join $mytemp SFA-excel-theme.xml]]
      }

# delete all but one worksheet
      catch {$excel DisplayAlerts False}
      set sheetCount [$worksheets Count]
      for {set n $sheetCount} {$n > 1} {incr n -1} {[$worksheets Item [expr $n]] Delete}
      set sheetLast [$worksheets Item [$worksheets Count]]
      catch {$excel DisplayAlerts True}
      [$excel ActiveWindow] TabRatio [expr 0.7]

# check decimal separator
      if {[$excel UseSystemSeparators] == 1 && [$excel DecimalSeparator] == ","} {
        if {![info exists commaSeparator]} {
          set cmsg "Numbers in a STEP file use a period \".\" as the decimal separator.  Your version of Excel uses a comma \",\" as a decimal separator.  This might cause some real numbers to be formatted as a date in a spreadsheet.  For example, 1.5 might appear as 1-Mai.\n\nTo change the formatting in Excel, go to the Excel File menu > Options > Advanced.  Uncheck 'Use system separators' and change 'Decimal separator' to a period \".\" and 'Thousands separator' to a comma \",\"\n\nWarning: this applies to ALL Excel spreadsheets on your computer.  Change the separators back to their original values when finished.\n\nYou can always check the STEP file to see the actual value of the number."
          if {[info exists buttons]} {
            append cmsg "\n\nSee the section about Numbers at the end of Help > Text Strings and Numbers."
            tk_messageBox -title "Decimal Separator" -type ok -default ok -icon warning -message $cmsg
          } else {
            errorMsg $cmsg
          }
          set commaSeparator 1
        }
      }

# errors
    } emsg]} {
      errorMsg "Error opening Excel workbooks and worksheets: $emsg"
      catch {raise .}
      return 0
    }

# CSV files or view only
  } else {
    set rowmax [expr {2**20}]
    if {$opt(xlMaxRows) < $rowmax} {set rowmax $opt(xlMaxRows)}
  }

# -------------------------------------------------------------------------------------------------
# security classifications
  set secClass {}
  catch {
    ::tcom::foreach sc [$objDesign FindObjects [string trim security_classification]] {
      set scn {}
      set val [string trim [[[$sc Attributes] Item [expr 1]] Value]]
      if {$val != ""} {lappend scn $val}
      set val [string trim [[[$sc Attributes] Item [expr 2]] Value]]
      if {$val != ""} {lappend scn $val}
      set scl [[[$sc Attributes] Item [expr 3]] Value]
      set val [string trim [[[$scl Attributes] Item [expr 1]] Value]]
      if {$val != ""} {lappend scn $val}
      set scn [join $scn " | "]
      if {[lsearch $secClass $scn] == -1} {lappend secClass $scn}
    }
    if {[llength $secClass] > 0} {
      outputMsg "Security classification(s):"
      foreach sc $secClass {outputMsg " $sc" red}
      outputMsg " "
    }
  }

# -------------------------------------------------------------------------------------------------
# add header worksheet, for CSV files create directory and header file
  addHeaderWorksheet $numFile $fname

# -------------------------------------------------------------------------------------------------
# set Excel spreadsheet name, delete file if already exists

# user-defined file name
  if {$useXL} {
    set xlsmsg ""

# same directory as file
    set xlFileName "[file nativename [file join [file dirname $fname] [file rootname [file tail $fname]]]]-sfa.$extXLS"

# user-defined directory
    if {$opt(writeDirType) == 2} {
      set xlFileName "[file nativename [file join $writeDir [file rootname [file tail $fname]]]]-sfa.$extXLS"
    }

# file name too long
    if {[string length $xlFileName] > 218} {
      append xlsmsg "Spreadsheet file name is too long for Excel ([string length $xlFileName])."
      set xlFileName "[file nativename [file join $mydocs [file rootname [file tail $fname]]]]-sfa.$extXLS"
      if {[string length $xlFileName] < 219} {
        append xlsmsg "  Spreadsheet written to the home directory."
      }
    }

# delete existing file
    if {[file exists $xlFileName]} {
      if {[catch {
        file delete -force -- $xlFileName
      } emsg]} {
        if {[string length $xlsmsg] > 0} {append xlsmsg "\n"}
        append xlsmsg "Existing Spreadsheet will not be overwritten: [file tail $xlFileName]"
        catch {raise .}
      }
    }
  }

# add file name to menu
  set ok 0
  if {$numFile <= 1} {set ok 1}
  if {[info exists localNameList]} {if {[llength $localNameList] > 1} {set ok 1}}
  if {$ok} {addFileToMenu}

# set types of entities to process
  set entCategories {}
  foreach pr [array names entCategory] {
    set ok 1
    if {[info exists opt($pr)] && [info exists entCategory($pr)] && $ok} {
      if {$opt($pr)} {set entCategories [concat $entCategories $entCategory($pr)]}
    }
  }

# -------------------------------------------------------------------------------------------------
# set which entities are processed and which are not
  set entsToProcess {}
  set entsToIgnore {}
  set numEnts 0
  set noFontFile 0

# user-defined entity list
  catch {set userEntityList {}}
  if {$opt(stepUSER) && [llength $userEntityList] == 0 && [info exists userEntityFile]} {
    if {$userEntityFile != ""} {
      set userEntityList {}
      set fileUserEnt [open $userEntityFile r]
      while {[gets $fileUserEnt line] != -1} {
        set line [string trim $line]
        if {[string first " " $line] != -1} {set line [split $line " "]}
        foreach ent $line {lappend userEntityList [string trim [string tolower $ent]]}
      }
      close $fileUserEnt
      if {[llength $userEntityList] == 0} {
        set opt(stepUSER) 0
        checkValues
      }
    } else {
      errorMsg "No file is selected for the User-Defined List on the Generate tab."
      checkValues
    }
  }

# get entity types from ANCHOR section IDs
  set anchorEnts {}
  if {[info exists p21e3Section]} {
    foreach line $p21e3Section {
      set c2 [string first ";" $line]
      if {$c2 != -1} {set line [string range $line 0 $c2-1]}
      set c1 [string first "\#" $line]
      if {$c1 != -1} {
        set anchorID [string range $line $c1+1 end]
        if {[string is integer $anchorID]} {
          catch {
            set objValue  [$objDesign FindObjectByP21Id [expr {int($anchorID)}]]
            set anchorEnt [$objValue Type]
            lappend anchorEnts $anchorEnt
          }
        }
      }
    }
    set anchorEnts [lrmdups $anchorEnts]
  }

# get totals of each entity in file
  set skipList {}
  if {![info exists objDesign]} {return}
  catch {unset entCount}

  if {![info exists entityTypeNames]} {
    set msg "The STEP file cannot be processed."
    if {!$opt(syntaxChecker) && $guiSFA} {append msg "\n Use F8 to run the Syntax Checker to check for errors in the STEP file.\n See Help > Syntax Checker"}
    errorMsg $msg
    return
  }

# for all entity types, check for which ones to process
  set msgcount ""
  foreach entType $entityTypeNames {
    set entCount($entType) [$objDesign CountEntities "$entType"]

    if {$entCount($entType) > 0} {
      if {$numFile != 0} {
        set idx [setColorIndex $entType]
        if {$idx == -2} {set idx 99}
        lappend allEntity "$idx$entType"
        lappend fileEntity($numFile) "$entType $entCount($entType)"
        if {![info exists totalEntity($entType)]} {
          set totalEntity($entType) $entCount($entType)
        } else {
          incr totalEntity($entType) $entCount($entType)
        }
      }

# check entity count vs. max rows
      if {$gen(Excel) && $guiSFA && [lsearch $entCategories $entType] != -1} {
        if {$entCount($entType) > 10003 && $opt(xlMaxRows) > 10003} {
          if {$msgcount == ""} {set msgcount "The number of entities > 10000 for some types of entities.  Consider using a smaller value for Maximum Rows."}
          if {[lsearch $entCategory(stepGEOM) $entType] != -1 || [lsearch $entCategory(stepCPNT) $entType] != -1} {
            if {[string first "Geometry" $msgcount] == -1} {append msgcount "\nYou might want to uncheck the Geometry and/or Coordinates Entity Types on the Generate tab."}
          }
        }
      }

# user-defined entities
      set ok 0
      if {$opt(stepUSER) && [lsearch $userEntityList $entType] != -1} {set ok 1}

# STEP entities that are translated depending on the options
      set ok1 [setEntsToProcess $entType]
      if {$ok == 0} {set ok $ok1}

# entities in unsupported APs that are not AP203, AP214, AP242 - if not using a user-defined list or not generating a spreadsheet
      if {[string first "AP203" $stepAP] == -1 && [string first "AP214" $stepAP] == -1 && [string first "AP242" $stepAP] == -1} {
        if {!$opt(stepUSER) || $opt(xlFormat) == "None"} {
          set et $entType
          set c1 [string first "_and_" $et]
          if {$c1 != -1} {set et [string range $et 0 $c1-1]}
          if {[lsearch $ap203all $et] == -1 && [lsearch $ap214all $et] == -1 && [lsearch $ap242all $et] == -1} {
            if {$c1 == -1} {
              set ok 1
            } else {
              if {[lsearch $ap203all $entType] == -1 && [lsearch $ap214all $entType] == -1 && [lsearch $ap242all $entType] == -1} {
                set ok 1
              }
            }
          }

# user-defined list and AP209 views are not allowed when generating a spreadsheet
        } elseif {[string first "AP209" $stepAP] != -1 && $opt(viewFEA) && $opt(xlFormat) != "None"} {
          outputMsg " "
          errorMsg "Viewing the AP209 FEM is not allowed when a User-Defined List is selected on the Generate tab."
          set opt(viewFEA) 0
          checkValues
        }
      }

# check for composite entities with "_11"
      if {$opt(stepCOMP) && [string first "_11" $entType] != -1} {set ok 1}

# check for descriptive_representation_item
      if {$entType == "descriptive_representation_item"} {
        if {$opt(PMIGRF)} {set ok 1}

# check for font file with GD&T symbols (ARIALUNI.TTF), needed only for certain Unicode characters
        if {![file exists [file nativename C:/Windows/Fonts/ARIALUNI.TTF]] && \
            ![file exists [file join $env(USERPROFILE) AppData Local Microsoft Windows Fonts ARIALUNI.TTF]]} {
          if {![file exists [file join $mytemp ARIALUNI.TTF]]} {catch {[file copy -force -- [file join $wdir images ARIALUNI.TTF] [file join $mytemp ARIALUNI.TTF]]}}
          set noFontFile 1
        }
      }

# check for entities referred to in the ANCHOR section
      if {[lsearch $anchorEnts $entType] != -1} {set ok 1}

# handle '_and_' due to a complex entity, entType_1 is the first part before the '_and_'
      set entType_1 $entType
      set c1 [string first "_and_" $entType_1]
      if {$c1 != -1} {set entType_1 [string range $entType_1 0 $c1-1]}

# check for entities that cause crashes
      set noSkip 1
      if {[info exists skipEntities]} {if {[lsearch $skipEntities $entType] != -1} {set noSkip 0}}

# add to list of entities to process (entsToProcess), uses color index to set the order
      set cidx [setColorIndex $entType]
      if {([lsearch $entCategories $entType_1] != -1 || $ok)} {
        if {$noSkip} {
          lappend entsToProcess "$cidx$entType"
          incr numEnts $entCount($entType)
        } else {
          lappend skipList $entType
          lappend entsToIgnore $entType
          set entsIgnored($cidx$entType) $entCount($entType)
        }
      } elseif {[lsearch $entCategories $entType] != -1} {
        if {$noSkip} {
          lappend entsToProcess "$cidx$entType"
          incr numEnts $entCount($entType)
        } else {
          lappend skipList $entType
          lappend entsToIgnore $entType
          set entsIgnored($cidx$entType) $entCount($entType)
        }
      } else {
        lappend entsToIgnore $entType
        set entsIgnored($cidx$entType) $entCount($entType)
      }
    }
  }
  if {$msgcount != ""} {
    outputMsg " "
    append msgcount "\nSee Help > Large STEP Files"
    errorMsg $msgcount
  }

# -------------------------------------------------------------------------------------------------
# decide how to process tessellated geometry by SFA (original) or by stp2x3d
  set brep 0
  set tessEnts 0
  foreach item $brepGeomEntTypes {if {[info exists entCount($item)] && $entCount($item) > 0} {set brep 1}}
  foreach item [list tessellated_solid tessellated_shell tessellated_closed_shell tessellated_open_shell] {
    if {[info exists entCount($item)] && $entCount($item) > 0} {set tessEnts 1}
  }
  if {[info exists entCount(tessellated_brep_shape_representation)] && $entCount(tessellated_brep_shape_representation) > 0} {
    set tessBrep 1
  }

# setting for SFA original
  set tessSolid 0
  set opt(viewTessPart) 1
  set viz(TESSMESH) 1
  if {$tessEnts} {set viz(TESSPART) 1}

# use new stp2x3d for tessellated geometry, except if there is also B-rep or not using SFA original method
  if {$tessEnts && $brep == 0 && $opt(tessPartOld) == 0 && $tessBrep == 0} {
    set tessSolid 1
    set opt(viewTessPart) 0
    set opt(tessPartMesh) 0
    set opt(partNormals) 0
    set viz(TESSPART) 0
    set viz(TESSMESH) 0
  }

# -------------------------------------------------------------------------------------------------
# check if there is anything to view
  foreach typ {PMI PLACE TESSPART FEA} {set viz($typ) 0}
  if {$gen(View)} {
    if {$opt(viewPMI)} {
      foreach ao $aoEntTypes {
        if {[info exists entCount($ao)] && $entCount($ao) > 0} {set viz(PMI) 1}
        set ao1 "$ao\_and_characterized_object"
        if {[info exists entCount($ao1)] && $entCount($ao1) > 0} {set viz(PMI) 1}
        set ao1 "$ao\_and_geometric_representation_item"
        if {[info exists entCount($ao1)] && $entCount($ao1) > 0} {set viz(PMI) 1}
        if {[string first "placeholder" $ao] != -1} {
          if {[info exists entCount($ao)] && $entCount($ao) > 0} {set viz(PLACE) 1}
        }
      }
    }
    if {$opt(viewTessPart)} {
      if {[info exists entCount(tessellated_solid)] || [info exists entCount(tessellated_shell)] || \
          [info exists entCount(tessellated_closed_shell)] || [info exists entCount(tessellated_open_shell)]} {set viz(TESSPART) 1}
    }
    if {$opt(viewFEA) && [string first "AP209" $stepAP] == 0} {set viz(FEA) 1}
  }

# read expected PMI worksheet (once) if semantic PMI and correct file name
  set epmiUD ""
  if {$opt(PMISEM) && [string first "AP242" $stepAP] == 0 && $opt(xlFormat) != "None"} {

# NIST test case
    if {![info exists nistName]} {set nistName ""}
    if {$nistName != ""} {
      set tols [concat $tolNames [list dimensional_characteristic_representation datum datum_feature datum_reference_compartment datum_reference_element datum_system placed_datum_target_feature]]
      foreach tol $tols {if {[info exists entCount($tol)]} {set ok 1; break}}
      if {$ok && ![info exists nistPMImaster($nistName)]} {nistReadExpectedPMI}

# user-defined expected PMI
    } else {
      set lname [file tail [file rootname $localName]]
      for {set i 3} {$i < [string length $lname]} {incr i} {
        set ln [string range $lname 0 $i]
        set epmiFile [file join [file dirname $localName] SFA-EPMI-$ln.xlsx]
        if {[file exists $epmiFile]} {
          set epmiUD [file tail [file rootname $localName]]
          nistReadExpectedPMI $epmiFile
          break
        }
      }
    }
  }

# filter inverse relationships to check only by entities in file
  if {$opt(INVERSE)} {
    if {$entityTypeNames != ""} {
      initDataInverses
      set invNew {}
      foreach item $inverses {
        if {[lsearch $entityTypeNames [lindex $item 0]] != -1} {lappend invNew $item}
      }
      set inverses $invNew
    }
  }

# check draughting model entities for PMI saved views
  set draughtingModels {}
  foreach dm [list draughting_model \
                   draughting_model_and_tessellated_shape_representation \
                   characterized_object_and_draughting_model \
                   characterized_representation_and_draughting_model \
                   characterized_representation_and_draughting_model_and_representation \
                   characterized_representation_and_draughting_model_and_tessellated_shape_representation] {
    if {[info exists entCount($dm)] && $entCount($dm) > 0} {lappend draughtingModels $dm}
  }

# -------------------------------------------------------------------------------------------------
# list entities not processed based on skip file
  if {[llength $skipList] > 0} {
    if {[file exists $skipFileName]} {
      outputMsg " "
      if {$opt(xlFormat) != "None"} {
        set msg "Worksheets"
        if {!$useXL} {set msg "CSV files"}
        append msg " will not be generated for the entities listed in"
      } else {
        set msg "The Viewer might not generate anything because of the entities listed in:"
      }
      append msg " [truncFileName [file nativename $skipFileName]]"
      errorMsg $msg
      set str ""
      foreach item [lsort $skipList] {append str " [formatComplexEnt $item],"}
      outputMsg [string range $str 0 end-1] red
      if {$guiSFA} {errorMsg "Use F8 to run the Syntax Checker and See Help > Crash Recovery"}
    }
  }

# sort entsToProcess by color index
  set entsToProcess [lsort $entsToProcess]

# -------------------------------------------------------------------------------------------------
# process datum* and dimensional* entities before specific *_tolerance entities
  if {$opt(PMISEM)} {
    if {[info exists entCount(angularity_tolerance)] || \
        [info exists entCount(circular_runout_tolerance)] || \
        [info exists entCount(coaxiality_tolerance)] || \
        [info exists entCount(concentricity_tolerance)] || \
        [info exists entCount(cylindricity_tolerance)]} {
      set entsToProcessTmp(0) {}
      set entsToProcessTmp(1) {}
      set entsToProcessDatum {}
      set itmp 0
      for {set i 0} {$i < [llength $entsToProcess]} {incr i} {
        set str1 [lindex $entsToProcess $i]
        set tc [string range [lindex $entsToProcess $i] 0 1]
        if {$tc == $entColorIndex(stepTOLR)} {set itmp 1}
        if {[string first $entColorIndex(stepTOLR) $str1] == 0 && \
          ([string first "datum" $str1] == 2 || [string first "dimensional" $str1] == 2 || [string first $iloldscr $str1] == 2)} {
          lappend entsToProcessDatum $str1
        } else {
          lappend entsToProcessTmp($itmp) $str1
        }
      }
      if {$itmp && [llength $entsToProcessDatum] > 0} {
        set entsToProcess [concat $entsToProcessTmp(0) $entsToProcessDatum $entsToProcessTmp(1)]
      }
    }

# move dimensional_characteristic_representation to the beginning
    foreach entdim {$iloldscr dimensional_characteristic_representation} {
      if {[info exists entCount($entdim)]} {
        set dcr "$entColorIndex(stepTOLR)$entdim"
        set c1 [lsearch $entsToProcess $dcr]
        set entsToProcess [lreplace $entsToProcess $c1 $c1]
        set entsToProcess [linsert $entsToProcess 0 $dcr]
      }
    }
  }

# -------------------------------------------------------------------------------------------------
# move some entities to end of AP209 entities
  if {$gen(View) && $viz(FEA)} {
    set ok  1
    set ok1 0
    set etp {}

# order is important, first 2 nodal loads, 3rd displacements (like a load), 4th boundary conditions
    set ent209 [list nodal_freedom_action_definition \
                     surface_3d_element_boundary_constant_specified_surface_variable_value \
                     volume_3d_element_boundary_constant_specified_variable_value \
                     nodal_freedom_values \
                     single_point_constraint_element_values]
    foreach ent $entsToProcess {
      if {$ok && [string range $ent 0 1] > 19} {
        foreach ent1 $ent209 {
          if {[info exists entCount($ent1)]} {
            lappend etp "19$ent1"
            set ok1 1
          }
        }
        set ok 0
      }
      if {$ent != "19nodal_freedom_action_definition" && \
          $ent != "19surface_3d_element_boundary_constant_specified_surface_variable_value" && \
          $ent != "19volume_3d_element_boundary_constant_specified_variable_value" && \
          $ent != "19nodal_freedom_values" && \
          $ent != "19single_point_constraint_element_values"} {
        lappend etp $ent
      }
    }
    if {!$ok1} {
      foreach ent1 $ent209 {if {[info exists entCount($ent1)]} {lappend etp "19$ent1"}}
    }
    set entsToProcess $etp

# find last entity type that will be processed, order is very important
    set ents [list curve_3d_element_representation surface_3d_element_representation volume_3d_element_representation]
    if {$opt(feaLoads)} {
      lappend ents "nodal_freedom_action_definition"
      lappend ents "surface_3d_element_boundary_constant_specified_surface_variable_value"
      lappend ents "volume_3d_element_boundary_constant_specified_variable_value"
    }
    if {$opt(feaDisp)} {lappend ents "nodal_freedom_values"}
    if {$opt(feaBounds)} {lappend ents "single_point_constraint_element_values"}
    foreach ent $ents {if {[info exists entCount($ent)]} {set feaFirstEntity $ent; break}}
    foreach ent $ents {if {[info exists entCount($ent)]} {set feaLastEntity $ent}}
  }

# then strip off the color index
  for {set i 0} {$i < [llength $entsToProcess]} {incr i} {
    lset entsToProcess $i [string range [lindex $entsToProcess $i] 2 end]
  }

# -------------------------------------------------------------------------------------------------
# max progress bar - number of entities or finite elements
  if {[info exists buttons]} {
    $buttons(progressBar) configure -maximum $numEnts
    if {[string first "AP209" $stepAP] == 0 && $opt(xlFormat) == "None"} {
      set n 0
      foreach elem {curve_3d_element_representation surface_3d_element_representation volume_3d_element_representation} {
        if {[info exists entCount($elem)]} {incr n $entCount($elem)}
      }
      $buttons(progressBar) configure -maximum $n
    }
  }

# -------------------------------------------------------------------------------------------------
# check for ISO/ASME standards on product_definition_formation, product
  set tolStandard(type) ""
  set tolStandard(num)  ""
  set tolStandards {}
  foreach item {product_definition_formation product} {
    ::tcom::foreach thisEnt [$objDesign FindObjects $item] {
      ::tcom::foreach attr [$thisEnt Attributes] {
        if {[$attr Name] == "id"} {
          set val [$attr Value]
          if {([string first "ISO" $val] == 0 || [string first "ASME" $val] == 0) && [string first "NIST" [string toupper $val]] == -1} {
            if {[string first "ISO" $val] == 0} {
              set tolStandard(type) "ISO"
              if {[string first "1101" $val] != "" || [string first "16792" $val] != ""} {if {[string first $val $tolStandard(num)] == -1} {append tolStandard(num) "$val    "}}
            }
            if {[string first "ASME" $val] == 0 && [string first "NIST" [string toupper $val]] == -1} {
              set tolStandard(type) "ASME"
              if {[string first "Y14." $val] != ""} {if {[string first $val $tolStandard(num)] == -1} {append tolStandard(num) "$val    "}}
            }
            set ok 1
            foreach std $tolStandards {if {[string first $val $std] != -1} {set ok 0}}
            if {$ok} {lappend tolStandards $val}

            if {$item == "product_definition_formation" && $opt(PMISEM)} {
              if {[string first "Y14.5" $val]  != -1 || [string first "1101" $val]  != -1} {lappend spmiTypesPerFile "$tolStandard(type) dimensioning standard"}
              if {[string first "Y14.41" $val] != -1 || [string first "16792" $val] != -1} {lappend spmiTypesPerFile "$tolStandard(type) modeling standard"}
            }
          }
        }
      }
    }
  }
  if {[llength $tolStandards] > 0} {
    outputMsg "\nStandards:" blue
    foreach std [lsort $tolStandards] {outputMsg " $std"}
  }
  if {$tolStandard(type) == "ISO"} {
    set fn [string toupper [file tail $localName]]
    if {[string first "NIST_" $fn] == 0 && [string first "ASME" $fn] != -1} {errorMsg "All of the NIST CAD models use the ASME Y14.5 tolerance standard."}
  }

# check for entities in unicodeAttributes that might have Unicode strings, complex entities require special exceptions in proc unicodeStrings
  catch {unset unicodeString}
  catch {unset unicodeStringCM}
  set unicodeEnts {}
  if {$opt(xlUnicode) && (($opt(xlFormat) != "None" && $useXL) || $gen(View))} {
    set unicodeNumEnts 0
    if {[info exists entsToProcess]} {
      foreach ent [array names unicodeAttributes] {
        if {[lsearch $entsToProcess $ent] != -1} {
          if {([string first "AP2" $stepAP] == 0 && $unicodeInFile) || [string first "ISO13" $stepAP] == 0 || [string first "CUTTING_TOOL_" $stepAP] == 0} {
            lappend unicodeEnts [string toupper $ent]
            incr unicodeNumEnts $entCount($ent)
            set unicodeAttr($ent) $unicodeAttributes($ent)
          }
        }
      }
    }
    if {$gen(View)} {
      if {$opt(xlFormat) == "None"} {
        set unicodeEnts {}
        set unicodeNumEnts 0
      }
      foreach cment [list camera_model_d3 camera_model_d3_multi_clipping] {
        if {[info exists entCount($cment)] && $entCount($cment) > 0} {
          if {[lsearch $unicodeEnts [string toupper $cment]] == -1} {
            lappend unicodeEnts [string toupper $cment]
            incr unicodeNumEnts $entCount($cment)
            set unicodeAttr($cment) $unicodeAttributes($cment)
          }
        }
      }
    }
    if {[llength $unicodeEnts] > 0} {
      unicodeStrings $unicodeEnts
      set unicodeEnts {}
      foreach item $unicodeActual {lappend unicodeEnts [string toupper $item]}
    }
  }

# -------------------------------------------------------------------------------------------------
# generate worksheet for each entity or a view
  outputMsg " "
  if {$useXL} {
    outputMsg "Generating STEP Entity worksheets" blue
  } elseif {$opt(xlFormat) == "CSV"} {
    outputMsg "Generating STEP Entity CSV files" blue
  } elseif {$opt(xlFormat) == "None"} {
    outputMsg "Generating View"
  }

# initialize variables
  if {[catch {
    set nistCoverageLegend 0
    set entsWithErrors {}
    set gpmiTypesInvalid {}
    set idxColor(0) 0
    set idxColor(1) 0
    set inverseEnts {}
    set lastEnt ""
    set nprogBarEnts 0
    set savedViewName {}
    set savedViewNames {}
    set savedViewButtons {}
    set spmiEntity {}
    set spmiSumRow 1
    set stat 1
    set valRounded 0
    set wsCount 0
    foreach f {elements mesh meshIndex faceIndex} {catch {file delete -force -- [file join $mytemp $f.txt]}}

    if {[info exists dim]} {unset dim}
    set dim(prec,max) 0
    set dim(unit) ""
    set dim(unitOK) 1

# no entities to process
    if {[llength $entsToProcess] == 0 && $gen(Excel)} {
      errorMsg "For a Spreadsheet, select more Entity Types on the Generate tab and try again."
      catch {unset entsIgnored}
      if {!$gen(View)} {break}
    }
    set tlast [clock clicks -milliseconds]

# find camera models used in draughting model items
    if {$gen(View) || ($gen(Excel) && $opt(PMIGRF))} {
      foreach cms $cameraModels {if {[info exists entCount($cms)]} {pmiGetCameras; break}}
    }

# get validation properties related to graphic or semantic PMI
    if {$opt(PMIGRF) || $opt(PMISEM)} {getValProps}

# -------------------------------------------------------------------------------------------------
# loop over list of entities in file
    foreach entType $entsToProcess {
      if {$opt(xlFormat) != "None"} {
        set nerr1 0
        set lastEnt $entType

# increase maximum rows for Analyzer options
        set newmax 5003
        set rmax $rowmax
        if {$stepAPreport && $rowmax < $newmax} {
          if {$opt(PMISEM)} {
            foreach item [list "angular" "datum" "dimension" "limits_and_fits" "runout" "tolerance"] {
              if {[string first $item $entType] != -1} {set rmax $newmax; break}
            }
          }
          if {$opt(PMIGRF)} {if {[string first "annotation" $entType] != -1 && [string first "plane" $entType] == -1} {set rmax $newmax}}
          if {$opt(valProp)} {if {$entType == "property_definition" || $entType == "material_property"} {set rmax $newmax}}
          if {$entType == "descriptive_representation_item" && [info exists driUnicode]} {set rmax $newmax; unset driUnicode}
          if {$opt(PMIUUID)} {if {[string first "uuid" $entType] != -1} {set rmax $newmax}}
        }

# decide if inverses should be checked for this entity type
        set checkInv 0
        if {$opt(INVERSE)} {set checkInv [invSetCheck $entType]}
        if {$checkInv} {lappend inverseEnts $entType}

# check for bad attributes
        set badAttr [info exists badAttributes($entType)]

# check for Unicode strings
        set unicodeCheck 0
        if {[llength $unicodeEnts] > 0} {if {[lsearch $unicodeEnts [string toupper $entType]] != -1} {set unicodeCheck 1}}

# process the entity type
        catch {unset matrixList}
        ::tcom::foreach objEntity [$objDesign FindObjects [join $entType]] {
          if {$entType == [$objEntity Type]} {
            incr nprogBarEnts
            if {[expr {$nprogBarEnts%1000}] == 0} {update}

            if {[catch {
              if {$useXL} {
                set stat [getEntity $objEntity $rmax $checkInv $badAttr $unicodeCheck]
              } else {
                set stat [getEntityCSV $objEntity $badAttr]
              }
            } emsg1]} {

# process errors with entity
              if {$stat != 1} {break}

              set msg "Error processing "
              if {[info exists objEntity]} {
                if {[string first "handle" $objEntity] != -1} {
                  append msg "\#[$objEntity P21ID]=[$objEntity Type] (row [expr {$row($thisEntType)+2}]): $emsg1"

# handle specific errors
                  if {[string first "Unknown error" $emsg1] != -1} {
                    errorMsg $msg
                    catch {raise .}
                    incr nerr1
                    if {$nerr1 > 20} {
                      errorMsg "Processing of [formatComplexEnt $entType] entities has stopped" red
                      set nprogBarEnts [expr {$nprogBarEnts + $entCount($thisEntType) - $count($thisEntType)}]
                      break
                    }

                  } elseif {[string first "Insufficient memory to perform operation" $emsg1] != -1} {
                    errorMsg $msg
                    errorMsg "Several options are available to reduce memory usage:\nUse the option to limit the Maximum Rows"
                    if {$opt(INVERSE)} {errorMsg "Turn off Inverse Relationships and process the file again" red}
                    catch {raise .}
                    break
                  }
                  errorMsg $msg
                  catch {raise .}
                }
              }
            }

# max rows exceeded
            if {$stat != 1} {
              set ok 1
              if {[string first "element_representation" $thisEntType] != -1 && $opt(viewFEA)} {set ok 0}
              if {$ok} {set nprogBarEnts [expr {$nprogBarEnts + $entCount($thisEntType) - $count($thisEntType)}]}
              break
            }
          }
        }

# write matrix of values to the worksheet for this entity, matrixList is from getEntity
        if {$useXL} {
          if {[catch {
            set range [$worksheet($thisEntType) Range [cellRange 4 1] [cellRange [expr {[llength $matrixList]+3}] [llength [lindex $matrixList 0]]]]
            $range Value2 $matrixList
          } emsg3]} {
            errorMsg "Error writing worksheet cells for $thisEntType: $emsg3"
          }

# close CSV file
        } else {
          catch {close $fcsv}
        }
      }

# check for reports (properties, semantic and graphic PMI, AP209 FEM)
      checkForReports $entType

# report errors related to descriptive_representation_item equivalent Unicode strings
      if {$entType == "descriptive_representation_item" && [info exists equivUnicodeStringErr]} {
        outputMsg " Warnings for 'equivalent unicode string': [join [lrmdups $equivUnicodeStringErr] "\; "]" red
        unset equivUnicodeStringErr
      }
    }

# check for entities in the STEP file that are not in the schema
    if {$opt(checkEntities)} {syntaxChecker $localName 1}

# generate tessellated geometry for viewer if using old SFA method
    if {$gen(View) && ($opt(tessPartOld) || $opt(viewTessPart))} {
      set tp 0
      foreach item [list tessellated_solid tessellated_shell tessellated_closed_shell tessellated_open_shell tessellated_wire] {
        if {[info exists entCount($item)] && $entCount($item) > 0} {tessPart $item; set tp 1}
      }
      if {$tp == 0} {
        set item "triangulated_face"
        if {[info exists entCount($item)] && $entCount($item) > 0} {tessPart $item}
      }
    }

# other errors
  } emsg2]} {
    catch {raise .}
    if {[llength $entsToProcess] > 0} {
      set msg "Error processing STEP file"
      if {[info exists objEntity]} {if {[string first "handle" $objEntity] != -1} {append msg " with entity \#[$objEntity P21ID]=[string toupper [$objEntity Type]]"}}
      append msg ": $emsg2\nProcessing of the STEP file has stopped"
      errorMsg $msg
    } else {
      return
    }
  }

# -------------------------------------------------------------------------------------------------
# check skip file
  if {[info exists skipFileName]} {
    set skiptmp {}
    if {[file exists $skipFileName]} {
      set skipFile [open $skipFileName r]
      while {[gets $skipFile line] >= 0} {
        if {[lsearch $skiptmp $line] == -1 && $line != $lastEnt} {lappend skiptmp $line}
      }
      close $skipFile
    }

    if {[join $skiptmp] == ""} {
      catch {file delete -force -- $skipFileName}
    } else {
      set skipFile [open $skipFileName w]
      foreach item $skiptmp {puts $skipFile $item}
      close $skipFile
    }
  }

# -------------------------------------------------------------------------------------------------
# generate B-rep part geometry if no other viz exists
  set vizprt 0
  if {$gen(View)} {
    if {$opt(viewPart) && !$viz(PMI) && !$viz(FEA) && !$viz(TESSPART) && ![info exists statsOnly]} {
      x3dFileStart
      set vizprt 1
    }

# generate B-rep part geom, set viewpoints, and close x3dom geometry file
    if {($viz(PMI) || $viz(FEA) || $viz(TESSPART) || $vizprt) && $x3dFileName != ""} {x3dFileEnd}
  }

# -------------------------------------------------------------------------------------------------
# add validation properties to some worksheets that are not associated with any PMI Analyzer report
  if {$opt(xlFormat) == "Excel" && [lsearch $characteristics "Properties"] != -1} {
    set ok 0
    if {$opt(PMISEM)} {
      foreach item [list Dimensions Datums "Datum Targets" "Geometric Tolerances"] {if {[lsearch $characteristics $item] != -1} {set ok 1}}
    }
    if {$opt(PMIGRF)} {
      foreach item $characteristics {if {[string first "Graphic PMI" $item] != -1} {set ok 1}}
    }
    if {[lsearch $characteristics "Composites"] != -1} {set ok 1}
    if {$ok} {reportValProps}
  }

# -------------------------------------------------------------------------------------------------
# add persistent IDs from uuid_attribute entities (AP242 Edition 4)
  if {$opt(xlFormat) == "Excel" || ($opt(xlFormat) == "CSV" && $useXL)} {
    set entsUUID {}
    set uuidEnts {}
    set totalUUID 0
    if {$opt(PMIUUID)} {
      foreach ent [list HASH_BASED_V5_UUID_ATTRIBUTE V4_UUID_ATTRIBUTE V5_UUID_ATTRIBUTE UUID_ATTRIBUTE_WITH_APPROXIMATE_LOCATION] {
        set entlc [string tolower $ent]
        if {[info exists entCount($entlc)] && $entCount($entlc) > 0} {
          set totalUUID [expr {$totalUUID+$entCount($entlc)}]
          lappend entsUUID $ent
        }
      }
      if {$totalUUID > 0} {uuidGetAttributes $totalUUID $entsUUID}
    }
  }

# -------------------------------------------------------------------------------------------------
# add summary worksheet
  if {$useXL && [llength $entsToProcess] > 0} {
    set tmp [sumAddWorksheet]
    set sumLinks  [lindex $tmp 0]
    set sheetSort [lindex $tmp 1]
    set sumRow    [lindex $tmp 2]
    set sum "Summary"

# add file name and other info to top of Summary
    set sumHeaderRow [sumAddFileName $sum $sumLinks]

# freeze panes (must be before adding color and hyperlinks below)
    [$worksheet($sum) Range "A[expr {$sumHeaderRow+3}]"] Select
    catch {[$excel ActiveWindow] FreezePanes [expr 1]}
    [$worksheet($sum) Range "A1"] Select

# -------------------------------------------------------------------------------------------------
# format cells on each entity worksheets
    formatWorksheets $sheetSort $sumRow $inverseEnts

# add Summary color and hyperlinks
    sumAddColorLinks $sum $sumHeaderRow $sumLinks $sheetSort $sumRow

# -------------------------------------------------------------------------------------------------
# add Semantic PMI Coverage Analysis worksheet for a single file
    if {$opt(PMISEM) && $stepAPreport} {

# check for datum and datum_system
      if {!$opt(PMISEMDIM) && !$opt(PMISEMDT)} {
        if {[info exists entCount(datum)]} {
          for {set i 0} {$i < $entCount(datum)} {incr i} {lappend spmiTypesPerFile "datum (6.5)"}
        }
        if {[info exists entCount(datum_system)] && [lsearch $spmiTypesPerFile "datum system"] == -1} {
          for {set i 0} {$i < $entCount(datum_system)} {incr i} {lappend spmiTypesPerFile "datum system"}
        }
      }

      if {[info exists spmiTypesPerFile]} {
        set ok 0

# do not generate if only certain PMI types were counted
        foreach type [lrmdups $spmiTypesPerFile] {
          if {[string first "annotation placeholder" $type] == -1 && [string first "editable text" $type] == -1 && \
              [string first "saved views" $type] == -1 && [string first "section views" $type] == -1 && \
              [string first "supplemental geometry" $type] == -1 && [string first "document identification" $type] == -1 && \
              [string first "standard" $type] == -1 && [string first "default tolerance decimal places" $type] == -1} {set ok 1; break}
        }

        if {$ok} {
          set spmiCoverageWS "Semantic PMI Coverage"
          if {![info exists worksheet($spmiCoverageWS)]} {
            outputMsg "Generating Semantic PMI Coverage worksheet" blue
            spmiCoverageStart 0
            spmiCoverageWrite "" "" 0
            spmiCoverageFormat "" 0
          }
        } else {
          unset spmiTypesPerFile
        }
      }

# format Semantic PMI Summary worksheet
      if {[info exists spmiSumName]} {
        set name $nistName
        if {[info exists epmiFile]} {if {$epmiFile != ""} {set name $epmiUD}}
        if {$name != "" && [info exists nistPMIexpected($name)]} {nistPMISummaryFormat $name}
        [$worksheet($spmiSumName) Columns] AutoFit
        [$worksheet($spmiSumName) Rows] AutoFit
      }
      catch {unset spmiSumName}
    }

# add Graphic PMI Coverage Analysis worksheet for a single file
    if {$opt(PMIGRF) && $opt(xlFormat) != "None" && $stepAPreport} {
      if {[info exists gpmiTypesPerFile]} {
        set gpmiCoverageWS "Graphic PMI Coverage"
        if {![info exists worksheet($gpmiCoverageWS)]} {
          outputMsg "Generating Graphic PMI Coverage worksheet" blue
          gpmiCoverageStart 0
          gpmiCoverageWrite "" "" 0
          gpmiCoverageFormat "" 0
        }
      }
    }

# reset rounding
    if {[info exists resetRound]} {
      set opt(PMISEMRND) $resetRound
      unset resetRound
    }

# -------------------------------------------------------------------------------------------------
# add UUID summary
    if {[info exists uuidCount]} {uuidSummary}

# add ANCHOR and other sections from Part 21 Edition 3
    if {[info exists p21e3Section]} {if {[llength $p21e3Section] > 0} {uuidReportAttributes 1}}

# -------------------------------------------------------------------------------------------------
# generate bill of materials (BOM)
    if {$opt(BOM) && ([string first "AP203" $stepAP] != -1 || [string first "AP214" $stepAP] != -1 || [string first "AP242" $stepAP] != -1)} {generateBOM}

# -------------------------------------------------------------------------------------------------
# select the first tab
    [$worksheets Item [expr 1]] Select
    [$excel ActiveWindow] ScrollRow [expr 1]
  }

# -------------------------------------------------------------------------------------------------
# quit IFCsvr
  if {[catch {
    $objDesign Delete
    unset objDesign
    unset objIFCsvr
    if {$roseSchemas != "" && $roseSchemas != $ifcsvrDir} {
      unset env(ROSE_SCHEMAS)
      if {$roseSchemas != ""} {set env(ROSE_SCHEMAS) $roseSchemas}
    }

# errors
  } emsg]} {
    errorMsg "Error closing IFCsvr: $emsg"
    catch {raise .}
  }

# processing time
  set cc [clock clicks -milliseconds]
  set proctime [expr {($cc - $lasttime)/1000}]
  if {$proctime <= 60} {set proctime [expr {(($cc - $lasttime)/100)/10.}]}
  set clr "black"
  if {!$useXL} {set clr "blue"}
  outputMsg "Processing time: $proctime seconds" $clr
  incr filesProcessed
  update

# -------------------------------------------------------------------------------------------------
# save spreadsheet
  set csvOpenDir 0
  if {$useXL && [llength $entsToProcess] > 0} {
    if {[catch {
      outputMsg " "
      if {$xlsmsg != ""} {outputMsg $xlsmsg red}

      if {!$xlsManual || $numFile != 0} {
        set xlFileName [checkFileName $xlFileName]
        set xlfn $xlFileName

# create new file name if spreadsheet already exists, delete new file name spreadsheets if possible
        if {[file exists $xlfn]} {set xlfn [incrFileName $xlfn]}

# always save as spreadsheet
        outputMsg "Saving Spreadsheet to:"
        outputMsg " [truncFileName $xlfn 1]" blue
        if {[catch {
          catch {$excel DisplayAlerts False}
          if {$xlFormat == 51} {
            $workbook -namedarg SaveAs Filename $xlfn FileFormat $xlFormat
          } else {
            $workbook -namedarg SaveAs Filename $xlfn
          }
          catch {$excel DisplayAlerts True}
          set lastXLS $xlfn
          lappend xlFileNames $xlfn
        } emsg1]} {
          errorMsg "Error Saving Spreadsheet: $emsg1"
        }

# user to manually save spreadsheet
      } else {
        $excel Visible 1
        outputMsg "You have to manually save the Excel spreadsheet" red
      }

# save worksheets as CSV files
      if {$saveCSV} {
        if {[catch {

# set directory for CSV files
          set csvdirnam "[file join [file dirname $localName] [file rootname [file tail $localName]]]-sfa-csv"
          if {$opt(writeDirType) == 2} {set csvdirnam [file join $writeDir [file rootname [file tail $localName]]-sfa-csv]}
          if {[string first "\[" [file dirname $localName]] != -1 || [string first "\]" [file dirname $localName]] != -1} {
            set csvdirnam "[file join $mydocs [file rootname [file tail $localName]]]-sfa-csv"
          }
          if {[string first "\[" $csvdirnam] != -1 || [string first "\]" $csvdirnam] != -1} {
            regsub -all {\[} $csvdirnam "(" csvdirnam
            regsub -all {\]} $csvdirnam ")" csvdirnam
          }
          file mkdir $csvdirnam
          outputMsg "Saving Spreadsheet as multiple CSV files to:"
          outputMsg " [truncFileName [file nativename $csvdirnam]]" blue
          set csvFormat [expr 6]
          if {$excelVersion > 15} {set csvFormat [expr 62]}

          set csvinhome 0
          set nprogBarEnts 0
          for {set i 1} {$i <= [$worksheets Count]} {incr i} {
            set ws [$worksheets Item [expr $i]]
            set wsn [$ws Name]
            if {[info exists wsNames($wsn)]} {
              set wsname $wsNames($wsn)
            } else {
              set wsname $wsn
            }

            $worksheet($wsname) Activate
            regsub -all " " $wsname "-" wsname
            set csvfname [file nativename [file join $csvdirnam $wsname.csv]]
            if {[string length $csvfname] > 218} {
              set csvfname [file nativename [file join $mydocs $wsname.csv]]
              errorMsg " Some CSV files are saved in the home directory." red
              set csvinhome 1
            }
            catch {file delete -force -- $csvfname}
            if {[file exists $csvfname]} {set csvfname [incrFileName $csvfname]}
            set csvfname [checkFileName $csvfname]

            $workbook -namedarg SaveAs Filename [file rootname $csvfname] FileFormat $csvFormat
            incr nprogBarEnts
            update
          }
        } emsg2]} {
          errorMsg "Error Saving CSV files: $emsg2"
        }
      }

      catch {$excel ScreenUpdating 1}

# close Excel if not manually saving spreadsheet
      set openxl 0
      if {!$xlsManual || $numFile != 0} {
        $excel Quit
        set openxl 1
        catch {unset excel}
        catch {if {[llength $pidExcel] == 1} {twapi::end_process $pidExcel -force}}
      }

# add Link(n) text to multi file summary
      if {$numFile != 0 && [info exists cells1(Summary)]} {
        set colsum [expr {$col1(Summary)+1}]
        if {!$opt(xlHideLinks)} {
          $cells1(Summary) Item 3 $colsum "Link ($numFile)"
          set range [$worksheet1(Summary) Range [cellRange 3 $colsum]]
          regsub -all {\\} $xlFileName "/" xls
        } else {
          $cells1(Summary) Item 3 $colsum "$numFile"
        }
      }
      update idletasks

# errors
    } emsg]} {
      errorMsg "Error: $emsg"
      catch {raise .}
      set openxl 0
    }

# -------------------------------------------------------------------------------------------------
# open spreadsheet or directory of CSV files
    set ok 0
    if {$openxl && $opt(outputOpen)} {
      if {$numFile == 0} {
        set ok 1
      } elseif {[info exists lenfilelist]} {
        if {$lenfilelist == 1} {set ok 1}
      }
    }

# open spreadsheet
    if {$useXL} {
      if {$ok} {
        openXLS $xlfn
      } elseif {!$opt(outputOpen) && $numFile == 0 && $guiSFA && !$xlsManual} {
        outputMsg " Use F2 to open the Spreadsheet" red
      }
    }

# CSV files generated too
    if {$saveCSV} {set csvOpenDir 1}

# open directory of CSV files
  } elseif {$opt(xlFormat) != "None" && [llength $entsToProcess] > 0} {
    set csvOpenDir 1
    unset csvfile
    outputMsg "\nCSV files written to:"
    outputMsg " [truncFileName [file nativename $csvdirnam]]" blue
  }

  if {$opt(xlFormat) == "None"} {set useXL 1}

# open directory of CSV files
  if {$csvOpenDir} {
    set ok 0
    if {$opt(outputOpen)} {
      if {$numFile == 0} {
        set ok 1
      } elseif {[info exists lenfilelist]} {
        if {$lenfilelist == 1} {set ok 1}
      }
    }
    if {$ok} {
      outputMsg "Opening directory of CSV files"
      catch {
        exec C:/Windows/explorer.exe [file nativename $csvdirnam] &
        if {[info exists csvinhome]} {if {$csvinhome} {exec C:/Windows/explorer.exe $mydocs &}}
      }
    }
  }

# -------------------------------------------------------------------------------------------------
# open x3dom file for views
  if {$x3dViewOK} {openX3DOM "" $numFile}

# save log file
  if {[info exists logFile]} {
    update idletasks
    saveLogFile $lfile
    unset lfile
  } elseif {[info exists buttons]} {
    if {[info exists currLogFile]} {unset currLogFile}
    bind . <Key-F4> {}
  }

# -------------------------------------------------------------------------------------------------
# save state
  if {[info exists errmsg]} {unset errmsg}
  saveState
  if {!$multiFile && [info exists buttons]} {$buttons(generate) configure -state normal}
  update idletasks

# unset variables
  foreach var {anchorSum assemTransformPMI brepScale cells cgrObjects cmNameID colColor commasep count currx3dPID datumEntType datumGeom datumIDs datumSymbol datumSystem datumSystemPDS defComment dimrep dimrepID dimtolEnt dimtolEntID dimtolGeom draughtingModels draftModelCameraNames draftModelCameras driPropID entCount entName entsIgnored epmi epmiUD equivUnicodeString feaDOFR feaDOFT feaNodes fileSumRow fontErr gpmiID gpmiIDRow gpmiRow heading idRow invCol invGroup noFontFile npart nrep numx3dPID placeAxes placeAxesDef placeCoords placeSphereDef pmiCol pmiColumns pmiStartCol pmivalprop propDefID propDefIDRow propDefName propDefOK propDefRow ptzError savedsavedViewNames savedViewFile savedViewFileName savedViewItems savedViewNames savedViewpoint savedViewVP shapeRepName spmiSumRowID srNames suppGeomEnts syntaxErr taoLastID tessCoord tessCoordID tessCoordName tessIndex tessIndexCoord tessPlacement tessReadOnce tessRepo trimVal unicode unicodeActual unicodeNumEnts unicodeString viz vpEnts workbook workbooks worksheet worksheets x3dCoord x3dFile x3dFileName x3dIndex x3dMax x3dMin x3dStartFile} {
    catch {global $var}
    if {[info exists $var]} {unset $var}
  }
  if {!$multiFile} {foreach var {gpmiTypesPerFile spmiTypesPerFile} {catch {global $var}; if {[info exists $var]} {unset $var}}}

# delete leftover text files in temp directory
  foreach f [glob -nocomplain -directory $mytemp *.txt] {catch {file delete -force -- $f}}
  catch {file delete -force -- [file join $mytemp "gunzip.exe"]}
  update idletasks
  return 1
}

# -------------------------------------------------------------------------------------------------
proc addHeaderWorksheet {numFile fname} {
  global objDesign
  global ap242ed cadApps cadSystem cells cells1 col1 csvdirnam developer excel excel1 fileItems fileSchema legendColor
  global localName opt recPracVersions row spmiTypesPerFile stepAP timeStamp useXL writeDir worksheet worksheet1 worksheets

  if {[catch {
    set cadSystem ""
    set timeStamp ""

    set hdr "Header"
    if {$useXL} {
      outputMsg "Generating Header worksheet" blue
      set worksheet($hdr) [$worksheets Item [expr 1]]
      $worksheet($hdr) Activate
      $worksheet($hdr) Name $hdr
      set cells($hdr) [$worksheet($hdr) Cells]

# create directory for CSV files
    } elseif {$opt(xlFormat) != "None"} {
      outputMsg "Generating Header CSV file" blue
      foreach var {csvdirnam csvfname fcsv} {catch {unset $var}}
      set csvdirnam "[file join [file dirname $localName] [file rootname [file tail $localName]]]-sfa-csv"
      if {$opt(writeDirType) == 2} {set csvdirnam "[file join $writeDir [file rootname [file tail $localName]]]-sfa-csv"}
      file mkdir $csvdirnam
      set csvfname [file join $csvdirnam $hdr.csv]
      if {[file exists $csvfname]} {file delete -force -- $csvfname}
      set fcsv [open $csvfname w]
    }

    set row($hdr) 0
    foreach attr {Name FileDirectory FileDescription FileImplementationLevel FileTimeStamp FileAuthor \
                  FileOrganization FilePreprocessorVersion FileOriginatingSystem FileAuthorisation SchemaName} {
      incr row($hdr)
      if {$useXL} {
        $cells($hdr) Item $row($hdr) 1 $attr
      } elseif {$opt(xlFormat) != "None"} {
        set csvstr $attr
      }
      set objAttr [string trim [join [$objDesign $attr]]]

# FileDirectory
      if {$attr == "FileDirectory"} {
        if {$useXL} {
          $cells($hdr) Item $row($hdr) 2 [$objDesign $attr]
        } elseif {$opt(xlFormat) != "None"} {
          append csvstr ",[$objDesign $attr]"
          puts $fcsv $csvstr
        }
        outputMsg "$attr:  [$objDesign $attr]"

# SchemaName
      } elseif {$attr == "SchemaName"} {
        set sn $fileSchema
        if {$useXL} {
          $cells($hdr) Item $row($hdr) 2 $sn
        } elseif {$opt(xlFormat) != "None"} {
          append csvstr ",$sn"
          puts $fcsv $csvstr
        }
        set str "$attr:  $sn"

# check edition of AP242 (schema identifier)
        set c1 [string first "1 0 10303 442" $sn]
        set simsg ""
        if {$c1 != -1} {
          set id [lindex [split [string range $sn $c1+14 end] " "] 0]
          if {$id == 1} {
            append str " (Edition 1, 2014)"
            if {[llength $ap242ed(2)] > 0 || [llength $ap242ed(3)] > 0 || [llength $ap242ed(4)] > 0} {
              if {$simsg != ""} {append simsg \n}
              append simsg " The STEP file contains entities ([join [lrmdups [concat $ap242ed(2) $ap242ed(3) $ap242ed(4)]]]) from a newer edition of AP242, however, the file is identified as Edition 1.  See Websites > STEP > EXPRESS Schemas"
            }
          } elseif {$id == 2 || $id == 3} {
            append str " (Edition 2, 2020)"
            if {[llength $ap242ed(3)] > 0 || [llength $ap242ed(4)] > 0} {
              set simsg " The STEP file contains entities ([join [lrmdups [concat $ap242ed(3) $ap242ed(4)]]]) from a newer edition of AP242, however, the file is identified as Edition 2.  See Websites > STEP > EXPRESS Schemas"
            }
          } elseif {$id == 4} {
            append str " (Edition 3, 2022)"
            if {[llength $ap242ed(4)] > 0} {
              set simsg " The STEP file contains entities ([join $ap242ed(4)]) found in AP242 Edition 4, however, the file is identified as Edition 3.  See Websites > STEP > EXPRESS Schemas"
            }
          } elseif {$id >= 5 && $id <= 7} {
            if {$id != 7} {set simsg " AP242 Edition 4 should be identified with '\{1 0 10303 442 7 1 4\}'  See Websites > STEP > EXPRESS Schemas"}
            append str " (Edition 4, 2025)"
          } elseif {$id > 7} {
            set simsg " Unsupported AP242 Schema Identifier '\{... $id 1 4\}'  See Websites > STEP > EXPRESS Schemas"
          }
          if {$developer} {foreach i {4 5} {if {[llength $ap242ed($i)] > 0} {regsub -all " " [join $ap242ed($i)] ", " str1; outputMsg " AP242e$i: $str1" red}}}
        } elseif {[string first "AP242" $sn] == 0} {
          set simsg " SchemaName is missing the Schema Identifier '\{1 0 10303 442 ...\}' that specifies the edition of AP242.  See Websites > STEP > EXPRESS Schemas"
        }

# check edition of AP214 (schema identifier)
        set c1 [string first "1 0 10303 214" $sn]
        if {$c1 != -1} {
          set id [lindex [split [string range $sn $c1+14 end] " "] 0]
          if {$id == 1 || $id == 3} {append str " (Edition $id)"}
        }

# check for IFC files
        if {[string first "IFC" $sn] == 0} {append str "  (Use the NIST IFC File Analyzer)"}

# schema name
        outputMsg $str blue
        if {$simsg != ""} {errorMsg $simsg red}

# check for multiple schemas
        if {[string first "," $sn] != -1} {
          errorMsg "Multiple FILE_SCHEMA names are not supported.  See Header worksheet."
          if {$useXL} {[[$worksheet($hdr) Range B11] Interior] Color $legendColor(red)}
        }

# other File attributes
      } else {
        if {$attr == "FileDescription" || $attr == "FileAuthor" || $attr == "FileOrganization"} {
          set str1 "$attr:  "
          set str2 ""
          foreach item [$objDesign $attr] {
            append str1 "[string trim $item], "
            if {$useXL} {
              append str2 "[string trim $item][format "%c" 10]"
            } elseif {$opt(xlFormat) != "None"} {
              append str2 ",[string trim $item]"
            }
          }
          outputMsg [string range $str1 0 end-2]
          if {$useXL} {
            $cells($hdr) Item $row($hdr) 2 "'[string trim $str2]"
            set range [$worksheet($hdr) Range "$row($hdr):$row($hdr)"]
            $range VerticalAlignment [expr -4108]
          } elseif {$opt(xlFormat) != "None"} {
            append csvstr [string trim $str2]
            puts $fcsv $csvstr
          }
        } else {
          outputMsg "$attr:  $objAttr"
          if {$useXL} {
            $cells($hdr) Item $row($hdr) 2 "'$objAttr"
            set range [$worksheet($hdr) Range "$row($hdr):$row($hdr)"]
            $range VerticalAlignment [expr -4108]
          } elseif {$opt(xlFormat) != "None"} {
            append csvstr ",$objAttr"
            puts $fcsv $csvstr
          }
        }

# check implementation level
        if {$attr == "FileImplementationLevel"} {
          if {[string first "\;" $objAttr] == -1} {
            errorMsg "FileImplementationLevel should be '2\;1'.  See Header worksheet."
            if {$useXL} {
              [[$worksheet($hdr) Range B4] Interior] Color $legendColor(red)
              addCellComment "Header" 4 2 "FileImplementationLevel should be '2\;1'."
            }
          }
        }

# check and add time stamp to multi file summary
        if {$attr == "FileTimeStamp"} {
          if {([string first "-" $objAttr] == -1 || [string first "T" $objAttr] == -1 || [string length $objAttr] < 17 || [string length $objAttr] > 26) && $objAttr != ""} {
            errorMsg "FileTimeStamp has the wrong format.  See Header worksheet."
            if {$useXL} {
              [[$worksheet($hdr) Range B5] Interior] Color $legendColor(red)
              addCellComment "Header" 5 2 "FileTimeStamp has the wrong format."
            }
          }
          set timeStamp $objAttr
          if {$numFile != 0 && [info exists cells1(Summary)] && $useXL} {
            set colsum [expr {$col1(Summary)+1}]
            set range [$worksheet1(Summary) Range [cellRange 5 $colsum]]
            catch {$cells1(Summary) Item 5 $colsum "'[string range $timeStamp 2 9]"}
          }
        }
      }
    }

# check for CAx-IF Recommended Practices in the file description
    set caxifrp {}
    set okcheck 0
    if {[string first "AP242" $stepAP] == 0 && $stepAP != "AP242e1" && $stepAP != "AP242e2" && $stepAP != "AP242"} {set okcheck 1}

    foreach fd [$objDesign "FileDescription"] {
      set c1 [string first "CAx-IF Rec." $fd]
      if {$c1 != -1} {lappend caxifrp [string trim [string range $fd $c1+20 end]]}
    }

    if {[llength $caxifrp] > 0} {
      outputMsg "\nCAx-IF Recommended Practices (See Websites):" blue
      foreach item $caxifrp {
        set str " $item"
        set c1 [string first "---" $item]
        if {$c1 != -1} {
          catch {
            set name [string trim [string range $item 0 $c1-1]]
            set c2 [string first "(" $name]
            if {$c2 != -1} {set name [string trim [string range $name 0 $c2-1]]}

# check version number
            if {[info exists recPracVersions($name)] && $okcheck} {
              set num [string range $item $c1+3 $c1+5]
              if {$num < $recPracVersions($name)} {set str " $item  (Current version is $recPracVersions($name))"}
            }
          }
        }
        outputMsg $str
        if {$opt(PMISEM)} {lappend spmiTypesPerFile "document identification"}
      }

# check if recommended practices are missing for recent versions of AP242
    } elseif {$okcheck} {
      set ok 0
      foreach item [list "Datum targets" Datums Dimensions "Geometric tolerances" "Supplemental geometry"] {
        if {[string first $item $fileItems] != -1} {
          errorMsg " FILE_DESCRIPTION is missing references to CAx-IF Recommended Practices (See Websites)" red
        }
      }
    }

# set the application from various file attributes, cadApps is a list of all apps defined in sfa-data.tcl, take the first one that matches
    set ok 0
    set app2 ""
    set fos [$objDesign FileOriginatingSystem]
    set fpv [$objDesign FilePreprocessorVersion]
    foreach attr {FileOriginatingSystem FilePreprocessorVersion FileOrganization} {
      foreach app $cadApps {
        set app1 $app
        if {$cadSystem == "" && [string first [string tolower $app] [string tolower [join [$objDesign $attr]]]] != -1} {
          set cadSystem [join [$objDesign $attr]]

# for multiple files, modify the app string to fit in file summary worksheet
          if {$app == "3D Evolution"}           {set app1 "CT 3D Evolution"}
          if {$app == "CoreTechnologie"}        {set app1 "CT 3D Evolution"}
          if {$app == "DATAKIT"}                {set app1 "Datakit"}
          if {$app == "EDMsix"}                 {set app1 "Jotne EDMsix"}
          if {$app == "PRO/ENGINEER"}           {set app1 "Pro/E"}
          if {$app == "SOLIDWORKS"}             {set app1 "SolidWorks"}
          if {$app == "SOLIDWORKS MBD"}         {set app1 "SolidWorks MBD"}
          if {$app == "SIEMENS PLM Software NX"} {set app1 "Siemens NX"}
          if {$app == "CREO"} {
            set app1 "Creo"
            if {[string index $cadSystem 28] == 2} {append app1 [string range $cadSystem 27 35]}
          }

          if {[string first "CATIA Version" $app] == 0} {set app1 "CATIA V[string range $app 14 end]"}
          if {$app == "3D EXPERIENCE"} {set app1 "3DX"}
          if {$app == "3DEXPERIENCE"}  {set app1 "3DX"}

          if {[string first "CATIA SOLUTIONS V4"      $fos] != -1} {set app1 "CATIA V4"}
          if {[string first "Autodesk Inventor"       $fos] != -1} {set app1 $fos}
          if {[string first "SolidWorks 2"            $fos] != -1} {set app1 $fos}
          if {[string first "MBDVidia"                $fos] != -1} {set app1 "MBDVidia"}
          if {[string first "SIEMENS PLM Software NX" $fos] ==  0} {set app1 "Siemens NX [string range $fos 24 end]"}
          if {[string first "THEOREM"                 $fpv] != -1} {set app1 "Theorem Solutions"}

# set caxifVendor based on CAx-IF vendor notation used in testing rounds, use for app if appropriate
          set caxifVendor [setCAXIFvendor]
          if {$caxifVendor != ""} {
            if {[string first [lindex [split $caxifVendor " "] 0] $app1] != -1} {
              if {[string length $caxifVendor] > [string length $app1]} {set app1 $caxifVendor}
            } elseif {[string first [lindex [split $app1 " "] 0] $caxifVendor] != -1} {
              if {[string length $caxifVendor] < [string length $app1]} {set app1 "$app1 ($caxifVendor)"}
            }
          }
          set ok 1
          set app2 $app1
          break
        }
      }
    }

# app name not found in cadApps
    if {$app2 == ""} {set app2 $fos}
    if {$app2 == "" || $app2 == "UNIX" || $app2 == "Windows" || $app2 == "Unspecified"} {set app2 $fpv}
    if {$app2 != ""} {set ok 1}

# add version number
    if {[string first "CATIA V5" $app2] == 0} {
      set c1 [string first "6 Release 20" $fpv]
      if {$c1 != -1} {append app2 " 6R20[string range $fpv $c1+12 $c1+14]"}
    } elseif {[string first "Datakit" $app2] == 0} {
      set c1 [string first "V20" $fpv]
      if {$c1 != -1} {append app2 " [string range $fpv $c1 $c1+6]"}
    } elseif {[string first "Elysium" $app2] == 0} {
      set c1 [string first "Translator v" $fos]
      if {$c1 != -1} {append app2 " [string range $fos $c1+11 end]"}
    } elseif {[string first "Kubotek" $app2] == 0} {
      set c1 [string first "Kosmos Version" $fpv]
      if {$c1 != -1} {
        append app2 " [string range $fpv $c1+15 end]"
      } else {
        set c1 [string first "Framework Version" $fpv]
        if {$c1 != -1} {append app2 " [string range $fpv $c1+18 end]"}
      }
    }

# add app2 to multiple file summary worksheet
    if {$numFile != 0 && $useXL && [info exists cells1(Summary)]} {
      if {$ok == 0} {set app2 [setCAXIFvendor]}
      set colsum [expr {$col1(Summary)+1}]
      if {$colsum > 16} {[$excel1 ActiveWindow] ScrollColumn [expr {$colsum-16}]}
      if {[string length $app2] > 35} {set app2 "[string range $app2 0 34]..."}
      regsub -all " " $app2 [format "%c" 10] app3
      $cells1(Summary) Item 6 $colsum [string trim $app3]
    }
    set cadSystem $app2
    if {$cadSystem == ""} {set cadSystem [setCAXIFvendor]}

# close csv file
    if {!$useXL && $opt(xlFormat) != "None"} {close $fcsv}

  } emsg]} {
    errorMsg "Error adding Header worksheet: $emsg"
    catch {raise .}
  }
}

#-------------------------------------------------------------------------------------------------
# add summary worksheet
proc sumAddWorksheet {} {
  global andEntAP209 cells col entCategory entCount entsIgnored equivUnicodeString excel fileSumRow
  global gpmiEnts iloldscr opt row sheetLast sheetSort spmiEntity stepAP sum uuidEnts vpEnts worksheet worksheets

  outputMsg "\nGenerating Summary worksheet" blue
  set sum "Summary"
  if {![info exists vpEnts]} {set vpEnts {}}

  set sheetSort {}
  foreach entType [lsort [array names worksheet]] {
    if {$entType != "Summary" && $entType != "Header" && $entType != "Section"} {
      lappend sheetSort "[setColorIndex $entType]$entType"
    }
  }
  set sheetSort [lsort $sheetSort]
  for {set i 0} {$i < [llength $sheetSort]} {incr i} {
    lset sheetSort $i [string range [lindex $sheetSort $i] 2 end]
  }

  if {[catch {
    set worksheet($sum) [$worksheets Add [::tcom::na] $sheetLast]
    $worksheet($sum) Activate
    $worksheet($sum) Name $sum
    set cells($sum) [$worksheet($sum) Cells]
    $cells($sum) Item 1 1 "Entity"
    $cells($sum) Item 1 2 "Count"
    set ncol 2
    set col($sum) $ncol
    set sumLinks [$worksheet($sum) Hyperlinks]

    set wsCount [$worksheets Count]
    [$worksheets Item [expr $wsCount]] -namedarg Move Before [$worksheets Item [expr 1]]

# Summary of entities in column 1 and count in column 2
    set x3dLink 1
    set row($sum) 1
    foreach entType $sheetSort {
      incr row($sum)
      set sumRow [expr {[lsearch $sheetSort $entType]+2}]
      set fileSumRow($entType) $sumRow

# check if entity is compound as opposed to an entity with '_and_'
      set ok 0

# no '_and_'
      if {[string first "_and_" $entType] == -1} {
        set ok 1

# check for explicit '_and_'
      } else {
        foreach item [array names entCategory] {if {[lsearch $entCategory($item) $entType] != -1} {set ok 1}}
        if {[string first "AP209" $stepAP] != -1} {foreach str $andEntAP209 {if {[string first $str $entType] != -1} {set ok 1}}}
      }

# no '_and_' or explicit '_and_'
      if {$ok} {
        $cells($sum) Item $sumRow 1 $entType

# add text strings
        set okao 0
        if {($entType == "property_definition" || $entType == "material_property") && $col($entType) > 4 && $opt(valProp)} {
          $cells($sum) Item $sumRow 1 "$entType  \[Properties\]"
        } elseif {$entType == "dimensional_characteristic_representation" && $col($entType) > 3 && $opt(PMISEM)} {
          $cells($sum) Item $sumRow 1 "dimensional_characteristic_representation  \[Semantic PMI\]"
        } elseif {$entType == $iloldscr && $col($entType) > 3 && $opt(PMISEM)} {
          $cells($sum) Item $sumRow 1 "$iloldscr  \[Semantic PMI\]"
        } elseif {[lsearch $spmiEntity $entType] != -1 && $opt(PMISEM)} {
          $cells($sum) Item $sumRow 1 "$entType  \[Semantic PMI\]"
        } elseif {([string first "annotation" $entType] != -1 || [string first "external_image" $entType] == 0) && $opt(PMIGRF)} {
          if {$gpmiEnts($entType) && $col($entType) > 5} {set okao 1}
        } elseif {[lsearch $vpEnts $entType] != -1} {
          $cells($sum) Item $sumRow 1 "$entType  \[Properties\]"
        } elseif {$entType == "next_assembly_usage_occurrence" && $opt(BOM)} {
          $cells($sum) Item $sumRow 1 "$entType  \[Assembly\]"
        } elseif {$entType == "descriptive_representation_item" && [info exists equivUnicodeString]} {
          $cells($sum) Item $sumRow 1 "$entType  \[Unicode String\]"
        }
        if {$okao} {$cells($sum) Item $sumRow 1 "$entType  \[Graphic PMI\]"}

# for '_and_' (complex entity) split on multiple lines
# '10' is the ascii character for a linefeed
      } else {
        regsub -all "_and_" $entType ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" entType_multiline
        set entType_multiline "($entType_multiline)"

# special case
        if {[[$cells($entType) Item 3 G] Value] == "datum_system" && [string first "(" [formatComplexEnt $entType]] == 0 && \
             [string first "with_datum_reference" $entType] == -1} {
           set entType_multiline "(geometric_tolerance_with_datum_reference)[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32]$entType_multiline"
        }
        $cells($sum) Item $sumRow 1 $entType_multiline

# add [Properties] or [Graphic PMI] text string
        set okao 0
        if {[string first "annotation" $entType] != -1} {
          if {$gpmiEnts($entType) && $col($entType) > 7} {set okao 1}
        } elseif {[lsearch $spmiEntity $entType] != -1} {
          $cells($sum) Item $sumRow 1 "$entType_multiline  \[Semantic PMI\]"
        } elseif {[lsearch $vpEnts $entType] != -1} {
          $cells($sum) Item $sumRow 1 "$entType_multiline  \[Properties\]"
        }
        if {$okao} {
          $cells($sum) Item $sumRow 1 "$entType_multiline  \[Graphic PMI\]"
        }
        set range [$worksheet($sum) Range $sumRow:$sumRow]
        $range VerticalAlignment [expr -4108]
      }

      if {[lsearch $uuidEnts $entType] != -1} {
        set val [[$cells($sum) Item $sumRow 1] Value]
        $cells($sum) Item $sumRow 1 "$val  \[UUID\]"
      }

# entity count in column 2
      $cells($sum) Item $sumRow 2 $entCount($entType)
    }

# entities not processed
    set rowIgnored [expr {[array size worksheet]+2}]
    $cells($sum) Item $rowIgnored 1 "Entity types not processed ([array size entsIgnored])"

    foreach ent [lsort [array names entsIgnored]] {
      set ent0 [string range $ent 2 end]
      set ok 0
      if {[string first "_and_" $ent] == -1} {
        set ok 1
      } else {
        foreach item [array names entCategory] {if {[lsearch $entCategory($item) $ent0] != -1} {set ok 1}}
      }
      if {$ok} {
        $cells($sum) Item [incr rowIgnored] 1 $ent0
      } else {
# '10' is the ascii character for a linefeed
        regsub -all "_and_" $ent0 ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" ent1
        $cells($sum) Item [incr rowIgnored] 1 "($ent1)"
        set range [$worksheet($sum) Range $rowIgnored:$rowIgnored]
        $range VerticalAlignment [expr -4108]
      }
      $cells($sum) Item $rowIgnored 2 $entsIgnored($ent)
    }
    set row($sum) $rowIgnored
    [$excel ActiveWindow] ScrollRow [expr 1]

# autoformat entire summary worksheet
    set range [$worksheet($sum) Range [cellRange 1 1] [cellRange $row($sum) $col($sum)]]
    $range AutoFormat

# name and link to program website that generated the spreadsheet
    $cells($sum) Item [expr {$row($sum)+2}] 1 "NIST STEP File Analyzer and Viewer [getVersion]"
    set anchor [$worksheet($sum) Range [cellRange [expr {$row($sum)+2}] 1]]
    [$worksheet($sum) Hyperlinks] Add $anchor [join "https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer"] [join ""] \
      [join "Link to NIST STEP File Analyzer and Viewer"]
    $cells($sum) Item [expr {$row($sum)+3}] 1 "[clock format [clock seconds]]"

# errors
  } emsg]} {
    errorMsg "Error adding Summary worksheet: $emsg"
    catch {raise .}
  }
  return [list $sumLinks $sheetSort $sumRow]
}

#-------------------------------------------------------------------------------------------------
# add file name and other info to top of Summary
proc sumAddFileName {sum sumLinks} {
  global cadSystem cells dim entityCount fileSchema localName opt stepAP sumHeaderRow timeStamp tolStandard worksheet xlFileName

  set sumHeaderRow 0
  if {[catch {
    $worksheet($sum) Activate
    [$worksheet($sum) Range "1:1"] Insert

    if {[info exists dim(unit)] && $dim(unit) != ""} {
      [$worksheet($sum) Range "1:1"] Insert
      $cells($sum) Item 1 1 "Dimension Units"
      $cells($sum) Item 1 2 "$dim(unit)"
      set range [$worksheet($sum) Range "B1:K1"]
      $range MergeCells [expr 1]
      incr sumHeaderRow
    }

    if {$tolStandard(type) != ""} {
      [$worksheet($sum) Range "1:1"] Insert
      $cells($sum) Item 1 1 "Standards"
      if {$tolStandard(num) != ""} {
        $cells($sum) Item 1 2 [string trim $tolStandard(num)]
      } else {
        $cells($sum) Item 1 2 [string trim $tolStandard(type)]
      }
      set range [$worksheet($sum) Range "B1:K1"]
      $range MergeCells [expr 1]
      incr sumHeaderRow
    }

    if {$stepAP != ""} {
      [$worksheet($sum) Range "1:1"] Insert
      $cells($sum) Item 1 1 "Schema"
      $cells($sum) Item 1 2 "'$stepAP"
      set range [$worksheet($sum) Range "B1:K1"]
      $range MergeCells [expr 1]
      set ap [string range $stepAP 0 4]
      set schemaLink ""
      if {$ap == "AP203" || $ap == "AP214" || $ap == "AP242" || $ap == "AP209"} {set schemaLink "https://www.mbx-if.org/home/mbx/resources/express-schemas/"}
      if {$schemaLink != ""} {
        set anchor [$worksheet($sum) Range "B1"]
        $sumLinks Add $anchor $schemaLink [join ""] [join "Link to $stepAP schema documentation"]
      }
      incr sumHeaderRow
    } else {
      [$worksheet($sum) Range "1:1"] Insert
      $cells($sum) Item 1 1 "Schema"
      $cells($sum) Item 1 2 "'$fileSchema"
      set range [$worksheet($sum) Range "B1:K1"]
      $range MergeCells [expr 1]
      incr sumHeaderRow
    }

    [$worksheet($sum) Range "1:1"] Insert
    $cells($sum) Item 1 1 "Total Entities"
    $cells($sum) Item 1 2 "'$entityCount"
    set range [$worksheet($sum) Range "B1:K1"]
    $range MergeCells [expr 1]
    incr sumHeaderRow

    if {$timeStamp != ""} {
      [$worksheet($sum) Range "1:1"] Insert
      $cells($sum) Item 1 1 "Timestamp"
      $cells($sum) Item 1 2 [join $timeStamp]
      set range [$worksheet($sum) Range "B1:K1"]
      $range MergeCells [expr 1]
      incr sumHeaderRow
    }

    if {$cadSystem != ""} {
      [$worksheet($sum) Range "1:1"] Insert
      $cells($sum) Item 1 1 "Application"
      $cells($sum) Item 1 2 [join $cadSystem]
      set range [$worksheet($sum) Range "B1:K1"]
      $range MergeCells [expr 1]
      incr sumHeaderRow
    }

    [$worksheet($sum) Range "1:1"] Insert
    $cells($sum) Item 1 1 "Excel File"
    if {[file dirname $localName] == [file dirname $xlFileName]} {
      $cells($sum) Item 1 2 [file tail $xlFileName]
    } else {
      $cells($sum) Item 1 2 [truncFileName $xlFileName]
    }
    set range [$worksheet($sum) Range "B1:K1"]
    $range MergeCells [expr 1]
    incr sumHeaderRow

    [$worksheet($sum) Range "1:1"] Insert
    $cells($sum) Item 1 1 "STEP File"
    $cells($sum) Item 1 2 [file tail $localName]
    set range [$worksheet($sum) Range "B1:K1"]
    $range MergeCells [expr 1]
    set anchor [$worksheet($sum) Range "B1"]
    if {!$opt(xlHideLinks) && [string first "#" $localName] == -1} {
      regsub -all {\\} $localName "/" ln
      $sumLinks Add $anchor [join $ln] [join ""] [join "Link to STEP file"]
    }
    incr sumHeaderRow

    [$worksheet($sum) Range "1:1"] Insert
    $cells($sum) Item 1 1 "STEP Directory"
    $cells($sum) Item 1 2 [file nativename [file dirname [truncFileName $localName]]]
    set range [$worksheet($sum) Range "B1:K1"]
    $range MergeCells [expr 1]
    incr sumHeaderRow

    set range [$worksheet($sum) Range [cellRange 1 1] [cellRange $sumHeaderRow 1]]
    [$range Font] Bold [expr 1]

  } emsg]} {
    errorMsg "Error adding File Names to Summary: $emsg"
    catch {raise .}
  }
  return $sumHeaderRow
}

#-------------------------------------------------------------------------------------------------
# add file name and other info to top of Summary
proc sumAddColorLinks {sum sumHeaderRow sumLinks sheetSort sumRow} {
  global cells col entCount entName entsIgnored entsWithErrors excel fileEntity nfile row worksheet

  if {[catch {
    set row($sum) [expr {$sumHeaderRow+2}]

# header worksheet
    set hdr "Header"
    set range [$worksheet($hdr) Range A1 A11]
    [$range Font] Bold [expr 1]
    [$worksheet($hdr) Columns] AutoFit
    [$worksheet($hdr) Rows] AutoFit

    foreach ent $sheetSort {
      update idletasks

      incr row($sum)
      set nrow [expr {20-$sumHeaderRow}]
      if {$row($sum) > $nrow} {[$excel ActiveWindow] ScrollRow [expr {$row($sum)-$nrow}]}

      set sumRow [expr {[lsearch $sheetSort $ent]+3+$sumHeaderRow}]

# link from summary to entity worksheet
      set anchor [$worksheet($sum) Range "A$sumRow"]
      set hlsheet $ent
      if {[string length $ent] > 31} {
        foreach item [array names entName] {
          if {$entName($item) == $ent} {set hlsheet $item}
        }
      }
      catch {$sumLinks Add $anchor [string trim ""] "$hlsheet!A1" "Go to $ent"}

# color cells
      set cidx [setColorIndex $ent]
      if {$cidx > 0} {

# color entities on summary if no errors or warnings and add comment that there are CAx-IF RP errors
        if {[lsearch $entsWithErrors [formatComplexEnt $ent]] == -1} {
          [$anchor Interior] ColorIndex [expr $cidx]

# color entities on summary gray and add comment that there are CAx-IF RP errors
        } else {
          [$anchor Interior] ColorIndex [expr 15]
          if {[info exists nfile]} {
            if {$nfile != 0} {
              set nent [lsearch $fileEntity($nfile) "$ent $entCount($ent)"]
              set fileEntity($nfile) [lreplace $fileEntity($nfile) $nent $nent "$ent -$entCount($ent)"]
            }
          }
          set comm "There are Errors or Warnings for at least one entity of this type.  See Help > Analyzer > Syntax Errors"
          if {$ent == "dimensional_characteristic_representation"} {append comm ".  Check for cell comments in the Associated Geometry column."}
          addCellComment $sum $sumRow 1 $comm
        }
        catch {foreach i {8 9} {[[$anchor Borders] Item $i] Weight [expr 1]}}
      }

# bold entities for reports
      if {[string first "\[" [$anchor Value]] != -1} {[$anchor Font] Bold [expr 1]}

      set ncol [expr {$col($sum)-1}]
    }

# add links for entsIgnored entities, find row where they start
    set i1 [expr {max([array size worksheet],9)}]
    for {set i $i1} {$i < 1000} {incr i} {
      if {[string first "Entity types" [[$cells($sum) Item $i 1] Value]] == 0} {
        set rowIgnored $i
        break
      }
    }
    set range [$worksheet($sum) Range "A$rowIgnored"]
    [$range Font] Bold [expr 1]

    set i1 0
    set range [$worksheet($sum) Range [cellRange $rowIgnored 1] [cellRange $rowIgnored [expr {$col($sum)+$i1}]]]
    catch {[[$range Borders] Item [expr 8]] Weight [expr -4138]}

    foreach ent [lsort [array names entsIgnored]] {
      incr rowIgnored
      set nrow [expr {20-$sumHeaderRow}]
      if {$rowIgnored > $nrow} {[$excel ActiveWindow] ScrollRow [expr {$rowIgnored-$nrow}]}
      set ncol [expr {$col($sum)-1}]

      set range [$worksheet($sum) Range [cellRange $rowIgnored 1]]
      set cidx [string range $ent 0 1]
      if {$cidx > 0} {[$range Interior] ColorIndex [expr $cidx]}
    }
    [$worksheet($sum) Columns] AutoFit
    [$worksheet($sum) Rows] AutoFit

  } emsg]} {
    errorMsg "Error adding Summary colors and links: $emsg"
    catch {raise .}
  }
}

#-------------------------------------------------------------------------------------------------
# format worksheets
proc formatWorksheets {sheetSort sumRow inverseEnts} {
  global buttons cells col count developer entCount entRows equivUnicodeString excel formattedEnts gpmiEnts idRow nprogBarEnts opt pmiStartCol
  global row spmiEnts stepAP stepAPreport sumHeaderRow syntaxErr thisEntType useXL uuidEnts viz vpEnts worksheet
  outputMsg "Formatting Worksheets" blue

  if {[info exists buttons]} {$buttons(progressBar) configure -maximum [llength $sheetSort]}
  set nprogBarEnts 0
  set nsort 0
  set okequiv 0
  set formattedEnts {}

  foreach thisEntType $sheetSort {
    incr nprogBarEnts
    update idletasks

    if {[catch {
      $worksheet($thisEntType) Activate
      [$excel ActiveWindow] ScrollRow [expr 1]

# move some worksheets to the correct position, originally moved to process semantic PMI data in the necessary order
      set moveWS 0
      if {$opt(PMISEM)} {
        foreach item {angularity_tolerance circular_runout_tolerance coaxiality_tolerance \
                      concentricity_tolerance cylindricity_tolerance dimensional_characteristic_representation} {
          if {[info exists entCount($item)] && $item == $thisEntType} {set moveWS 1}
        }
      }
      if {[string first "AP209" $stepAP] != -1 && $viz(FEA)} {
        foreach item {nodal_freedom_action_definition nodal_freedom_values \
                      surface_3d_element_boundary_constant_specified_surface_variable_value \
                      volume_3d_element_boundary_constant_specified_variable_value \
                      single_point_constraint_element_values} {
          if {[info exists entCount($item)] && $item == $thisEntType} {set moveWS 1}
        }
      }

      if {$moveWS} {
        if {[string first "dimensional_characteristic_repr" $thisEntType] == 0} {
          moveWorksheet [list dimensional_characteristic_repr dimensional_location dimensional_size]
        }
        foreach item {angularity_tolerance circular_runout_tolerance coaxiality_tolerance concentricity_tolerance cylindricity_tolerance} {
          if {$thisEntType == $item} {moveWorksheet [list $item datum]}
        }

        if {$thisEntType == "nodal_freedom_action_definition"} {moveWorksheet [list nodal_freedom_action_definition node]}
        if {$thisEntType == "nodal_freedom_values"} {moveWorksheet [list nodal_freedom_values node]}
        if {$thisEntType == "single_point_constraint_element_values"} {moveWorksheet [list single_point_constraint_element single_point_constraint_elemen1] After}
        if {$thisEntType == "surface_3d_element_boundary_constant_specified_surface_variable_value"} {moveWorksheet [list surface_3d_element_boundary_con surface_3d_element_descriptor]}
        if {$thisEntType == "volume_3d_element_boundary_constant_specified_variable_value"} {moveWorksheet [list volume_3d_element_boundary_cons volume_3d_element_descriptor]}
      }

# extent of columns and rows
      set rancol [[[$worksheet($thisEntType) UsedRange] Columns] Count]
      set ranrow $entRows($thisEntType)

# autoformat
      set range [$worksheet($thisEntType) Range [cellRange 3 1] [cellRange $ranrow $rancol]]
      $range AutoFormat

# freeze panes
      [$worksheet($thisEntType) Range "B4"] Select
      catch {[$excel ActiveWindow] FreezePanes [expr 1]}

# set A1 as default cell
      [$worksheet($thisEntType) Range "A1"] Select

# set column color, border, group for INVERSES and Used In
      if {$opt(INVERSE)} {if {[lsearch $inverseEnts $thisEntType] != -1} {invFormat $rancol}}

# properties
      if {($thisEntType == "property_definition" || $thisEntType == "material_property") && $opt(valProp)} {
        valPropFormat

# color STEP annotation occurrence (Graphic PMI)
      } elseif {$gpmiEnts($thisEntType) && $opt(PMIGRF) && $stepAPreport} {
        pmiFormatColumns "Graphic PMI"

# color STEP semantic PMI
      } elseif {$spmiEnts($thisEntType) && $opt(PMISEM) && $stepAPreport} {
        pmiFormatColumns "Semantic PMI"

# add Semantic PMI Summary worksheet
        if {$thisEntType != "datum_feature" && $stepAPreport} {spmiSummary}

# extra validation properties
      } elseif {[lsearch $vpEnts $thisEntType] != -1} {
        pmiFormatColumns "Validation Properties"

# equivalent Unicode string on descriptive_representation_item
      } elseif {$thisEntType == "descriptive_representation_item" && [info exists equivUnicodeString]} {
        outputMsg " $thisEntType"
        foreach id [array names equivUnicodeString] {
          if {[info exists idRow($thisEntType,$id)]} {
            catch {
              $cells($thisEntType) Item $idRow($thisEntType,$id) 4 $equivUnicodeString($id)
              set range [$worksheet($thisEntType) Range D$idRow($thisEntType,$id)]
              [$range Interior] ColorIndex 36
              [[$range Borders] Item [expr 8]] Weight [expr 1]
              [[$range Borders] Item [expr 9]] Weight [expr 1]
              set okequiv 1
            }
          }
        }
        if {$okequiv} {
          $cells($thisEntType) Item 3 4 "Equivalent Unicode String"
          set range [$worksheet($thisEntType) Range D3]
          [$range Interior] ColorIndex 36
          $range HorizontalAlignment [expr -4108]
          [$range Font] Bold [expr 1]
          [[$range Borders] Item [expr 8]] Weight [expr -4138]
          set range [$worksheet($thisEntType) Range D$ranrow]
          [[$range Borders] Item [expr 9]] Weight [expr -4138]
          [$worksheet($thisEntType) Columns] AutoFit
          [$worksheet($thisEntType) Rows] AutoFit
          addCellComment "descriptive_representation_item" 3 4 "The string interprets the characters '\\w' as ' | ' and '\\n' as a new line.  Unicode characters not supported by Windows fonts appear as a question mark.  When this column is sorted, the row height might need to be increased.  See Recommended Practice for PMI Unicode String Specification."
          incr rancol
        }
      }

# UUIDs
      if {[lsearch $uuidEnts $thisEntType] != -1} {
        set fthis [formatComplexEnt $thisEntType]
        if {[lsearch $formattedEnts $fthis] == -1} {outputMsg " $fthis"}
        uuidReportAttributes 2 $thisEntType
      }

# -------------------------------------------------------------------------------------------------
# link back to summary on entity worksheets
      set hlink [$worksheet($thisEntType) Hyperlinks]
      set txt "[formatComplexEnt $thisEntType]  "
      set row1 [expr {$row($thisEntType)-3}]
      if {$row1 == $count($thisEntType) && $row1 == $entCount($thisEntType)} {
        append txt "($row1)"
      } elseif {$row1 > $count($thisEntType) && $count($thisEntType) < $entCount($thisEntType)} {
        append txt "($count($thisEntType) of $entCount($thisEntType))"
      } elseif {$row1 < $entCount($thisEntType)} {
        if {$count($thisEntType) == $entCount($thisEntType)} {
          append txt "($row1 of $entCount($thisEntType))"
        } else {
          append txt "([expr {$row1-3}] of $count($thisEntType))"
        }
      }

# special case
      if {[[$cells($thisEntType) Item 3 G] Value] == "datum_system" && [string first "(" $txt] == 0 && \
           [string first "with_datum_reference" $txt] == -1} {
         set txt "(geometric_tolerance_with_datum_reference)$txt"
         if {$developer} {errorMsg " [formatComplexEnt $thisEntType] was missing (geometric_tolerance_with_datum_reference)" red}
      }
      $cells($thisEntType) Item 1 1 $txt

# set range of cells to merge with A1
      set c [[[$worksheet($thisEntType) UsedRange] Columns] Count]
      set okinv 0
      if {$opt(INVERSE)} {
        for {set i 1} {$i <= $c} {incr i} {
          set val [[$cells($thisEntType) Item 3 $i] Value]
          if {$val == "Used In" || [string first "INV-" $val] != -1} {
            set c [expr {$i-1}]
            set okinv 1
            break
          }
        }
      }
      if {!$okinv && [info exists pmiStartCol($thisEntType)]} {set c [expr {$pmiStartCol($thisEntType)-1}]}
      if {$thisEntType == "property_definition" || $thisEntType == "material_property"} {set c 4}
      if {$c > 8} {set c 8}
      if {$c == 1} {set c 2}
      set range [$worksheet($thisEntType) Range [cellRange 1 1] [cellRange 1 $c]]
      $range MergeCells [expr 1]

# link back to summary
      set anchor [$worksheet($thisEntType) Range "A1"]
      set sumRow [expr {[lsearch $sheetSort $thisEntType]+$sumHeaderRow+3}]
      catch {$hlink Add $anchor [string trim ""] "Summary!A$sumRow" "Return to Summary"}

# check width of columns, wrap text
      if {[catch {
        set widlim 400.
        for {set i 2} {$i <= $rancol} {incr i} {
          if {[[$cells($thisEntType) Item 3 $i] Value] != ""} {
            set wid [[$cells($thisEntType) Item 3 $i] Width]
            if {$wid > $widlim} {
              set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
              if {($thisEntType != "property_definition" && $thisEntType != "material_property") || $i != 9} {
                $range ColumnWidth [expr {[$range ColumnWidth]/$wid * $widlim}]
                $range WrapText [expr 1]
              }
            }
          }
        }
      } emsg]} {
        errorMsg "Error setting column widths for [formatComplexEnt $thisEntType]: $emsg"
        catch {raise .}
      }

# color red for syntax errors
      if {[info exists syntaxErr($thisEntType)]} {colorBadCells $thisEntType}

# -------------------------------------------------------------------------------------------------
# add table for sorting and filtering, always sort Analyzer worksheets and some others
      if {[catch {
        set oksort 0
        if {($opt(xlSort) && $useXL && $thisEntType != "property_definition")} {set oksort 1}
        if {$thisEntType == "descriptive_representation_item" && $okequiv} {set oksort 1}
        if {$spmiEnts($thisEntType) && $opt(PMISEM) && $stepAPreport} {set oksort 1}
        if {$gpmiEnts($thisEntType) && $opt(PMIGRF) && $stepAPreport} {set oksort 1}
        if {[string first "uuid_attribute" $thisEntType] != -1} {set oksort 1}

        if {$oksort && $ranrow > 7} {
          set range [$worksheet($thisEntType) Range [cellRange 3 1] [cellRange $ranrow $rancol]]
          set tname [string trim "TABLE-$thisEntType"]
          [[$worksheet($thisEntType) ListObjects] Add 1 $range] Name $tname
          [[$worksheet($thisEntType) ListObjects] Item $tname] TableStyle "TableStyleLight1"
        }
      } emsg]} {
        errorMsg "Error adding Tables for Sorting: $emsg"
        catch {raise .}
      }

# errors
    } emsg]} {
      errorMsg "Error formatting Spreadsheet for [formatComplexEnt $thisEntType]: $emsg"
      catch {raise .}
    }
  }
}

# -------------------------------------------------------------------------------------------------
proc moveWorksheet {items {where "Before"}} {
  global worksheets

  if {[catch {
    set n 0
    set p1 0
    set p2 1000
    foreach item $items {
      incr n
      for {set i 1} {$i <= [$worksheets Count]} {incr i} {
        if {$item == [[$worksheets Item [expr $i]] Name]} {
          if {$n == 1} {
            set p1 $i
          } else {
            set p2 [expr {min($p2,$i)}]
          }
        }
      }
    }

    if {$p1 != 0 && $p2 != 1000} {
      if {$where == "Before"} {
        [$worksheets Item [expr $p1]] -namedarg Move Before [$worksheets Item [expr $p2]]
      } else {
        [$worksheets Item [expr $p2]] -namedarg Move After [$worksheets Item [expr $p1]]
      }
    }
  } emsg]} {
    errorMsg "Error moving worksheet: $emsg"
  }
}
