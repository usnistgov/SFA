# generate an Excel spreadsheet from a STEP file
proc genExcel {{numFile 0}} {
  global allEntity aoEntTypes ap203all ap214all ap242all badAttributes brepEnts buttons cells cells1 col col1 count csvdirnam csvfile currLogFile developer
  global dim dimRepeatDiv editorCmd entCategories entCategory entColorIndex entCount entityCount entsIgnored entsWithErrors errmsg excel
  global excelVersion extXLS fcsv feaLastEntity File fileEntity filesProcessed gpmiTypesInvalid gpmiTypesPerFile idxColor ifcsvrDir inverses
  global lastXLS lenfilelist localName localNameList logFile multiFile multiFileDir mytemp nistCoverageLegend nistName nistPMIexpected nistPMImaster
  global nprogBarEnts nshape ofCSV ofExcel opt p21e3 p21e3Section row rowmax savedViewButtons savedViewName savedViewNames scriptName
  global sheetLast skipEntities skipPerm spmiEntity spmiSumName spmiSumRow spmiTypesPerFile startrow statsOnly stepAP tessColor thisEntType tlast
  global tolNames tolStandard tolStandards totalEntity userEntityFile userEntityList userXLSFile useXL viz workbook workbooks
  global worksheet worksheet1 worksheets writeDir wsCount wsNames x3dAxes x3dColor x3dColorFile x3dColors x3dColorsUsed x3dFileName x3dIndex
  global x3dMax x3dMin x3dMsg x3dStartFile xlFileName xlFileNames xlInstalled
  global objDesign

  if {[info exists errmsg]} {set errmsg ""}
  #outputMsg "genExcel" red

# initialize for X3DOM geometry
  if {$opt(VIZPMI) || $opt(VIZTPG) ||$opt(VIZFEA) || $opt(VIZBRP)} {
    set x3dStartFile 1
    set x3dAxes 1
    set x3dFileName ""
    set x3dColor ""
    set x3dColors  {}
    set x3dColorsUsed {}
    foreach idx {x y z} {
      set x3dMax($idx) -1.e10
      set x3dMin($idx)  1.e10
    }
    catch {unset tessColor}
    catch {unset x3dColorFile}
  }

# check if IFCsvr is installed
  if {![file exists [file join $ifcsvrDir IFCsvrR300.dll]]} {
    if {[info exists buttons]} {$buttons(genExcel) configure -state disable}
    installIFCsvr
    return
  }

  if {[info exists buttons]} {
    $buttons(genExcel) configure -state disable
    .tnb select .tnb.status
  }
  set lasttime [clock clicks -milliseconds]

  set multiFile 0
  if {$numFile > 0} {set multiFile 1}

# -------------------------------------------------------------------------------------------------
# connect to IFCsvr
  if {[catch {
    if {![info exists buttons]} {outputMsg "\n*** Begin ST-Developer output"}
    set objIFCsvr [::tcom::ref createobject IFCsvr.R300]
    if {![info exists buttons]} {outputMsg "*** End ST-Developer output"}

# error
  } emsg]} {
    errorMsg "\nERROR connecting to the IFCsvr software that is used to read STEP files: $emsg"
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

# open log file
    if {$opt(LOGFILE)} {
      set lfile [file rootname $fname]
      append lfile "-sfa.log"
      set logFile [open $lfile w]
      puts $logFile "NIST STEP File Analyzer and Viewer (v[getVersion])  [clock format [clock seconds]]"
    }

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
      if {$opt(XL_LINK1)} {[$worksheet1(Summary) Hyperlinks] Add $range [join $fname] [join ""] [join "Link to STEP file"]}
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
    if {![info exists buttons]} {outputMsg "\n*** Begin ST-Developer output\n*** Check for error or warning messages up to 'End ST-Developer output' below"}
    set objDesign [$objIFCsvr OpenDesign [file nativename $fname]]
    if {![info exists buttons]} {outputMsg "*** End ST-Developer output\n"}

# CountEntities causes the error if the STEP file cannot be opened because objDesign is null
    set entityCount [$objDesign CountEntities "*"]

# get stats
    set openStage 3
    set viz(PMIMSG) "The STEP file contains only Graphical PMI and no Semantic PMI."
    if {$entityCount > 0} {
      outputMsg " $entityCount entities"
      set entityTypeNames [$objDesign EntityTypeNames [expr 2]]
      set characteristics {}
      foreach entType $entityTypeNames {
        set ecount [$objDesign CountEntities "$entType"]
        if {$ecount > 0} {
          if {$entType == "dimensional_characteristic_representation"} {
            lappend characteristics "Dimensions"
            set viz(PMIMSG) "Some Graphical PMI might not have equivalent Semantic PMI in the STEP file."
          } elseif {$entType == "datum"} {
            lappend characteristics "Datums"
            set viz(PMIMSG) "Some Graphical PMI might not have equivalent Semantic PMI in the STEP file."

          } elseif {$entType == "tessellated_annotation_occurrence"} {
            lappend characteristics "Graphical PMI (tessellated)"
          } elseif {$entType == "annotation_occurrence" || [string first "annotation_curve_occurrence" $entType] != -1 || $entType == "annotation_fill_area_occurrence" || $entType == "annotation_occurrence_and_characterized_object"} {
            lappend characteristics "Graphical PMI (polyline)"

          } elseif {$entType == "advanced_brep_shape_representation" || $entType == "manifold_surface_shape_representation" || $entType == "manifold_solid_brep" || $entType == "shell_based_surface_model"} {
            lappend characteristics "B-rep geometry"
          } elseif {$entType == "constructive_geometry_representation"} {
            lappend characteristics "Supplemental geometry"
          } elseif {$entType == "tessellated_solid" || $entType == "tessellated_shell"} {
            lappend characteristics "Tessellated geometry"
          } elseif {$entType == "property_definition_representation"} {
            lappend characteristics "Properties"

          } elseif {[lsearch $entCategory(PR_STEP_COMP) $entType] != -1} {
            lappend characteristics "Composites"
          } elseif {[lsearch $entCategory(PR_STEP_KINE) $entType] != -1} {
            lappend characteristics "Kinematics"
          } elseif {[lsearch $entCategory(PR_STEP_FEAT) $entType] != -1} {
            lappend characteristics "Features"
          } else {
            foreach tol $tolNames {
              if {[string first $tol $entType] != -1} {
                lappend characteristics "Geometric tolerances"
                set viz(PMIMSG) "Some Graphical PMI might not have equivalent Semantic PMI in the STEP file."
              }
            }
          }
        }
      }
      if {[llength $characteristics] > 0} {
        set str ""
        foreach item [lrmdups $characteristics] {append str "$item, "}
        set str [string range $str 0 end-2]
        if {$str != "B-rep geometry"} {outputMsg "This file contains: $str" red}
      }
    } else {
      errorMsg "There are no entities in the STEP file."
    }
    outputMsg " "

# exit if stats only from command-line version
    if {[info exists statsOnly]} {
      if {[info exists logFile]} {
        update idletasks
        outputMsg "\nSaving Log file as:"
        outputMsg " [truncFileName [file nativename $lfile]]" blue
        close $logFile
        unset lfile
        unset logFile
      }
      exit
    }

# add AP, file size, entity count to multi file summary
    if {$numFile != 0 && [info exists cells1(Summary)]} {
      $cells1(Summary) Item [expr {$startrow-2}] $colsum $stepAP

      set fsize [expr {[file size $fname]/1024}]
      if {$fsize > 10240} {
        set fsize "[expr {$fsize/1024}] Mb"
      } else {
        append fsize " Kb"
      }
      $cells1(Summary) Item [expr {$startrow-1}] $colsum $fsize
      $cells1(Summary) Item $startrow $colsum $entityCount
    }

# open file of entities (-skip.dat) not to process (skipEntities), skipPerm are entities always to skip
    set cfile [file rootname $fname]
    append cfile "-skip.dat"
    set skipPerm {}
    set skipEntities $skipPerm
    if {[file exists $cfile]} {
      set skipFile [open $cfile r]
      while {[gets $skipFile line] >= 0} {
        if {[lsearch $skipEntities $line] == -1 && $line != "" && ![info exists badAttributes($line)]} {
          lappend skipEntities $line
        }
      }
      close $skipFile

# old skip file name (_fix.dat), delete
    } else {
      set cfile1 [file rootname $fname]
      append cfile1 "_fix.dat"
      if {[file exists $cfile1]} {
        set skipFile [open $cfile1 r]
        while {[gets $skipFile line] >= 0} {
          if {[lsearch $skipEntities $line] == -1 && $line != "" && ![info exists badAttributes($line)]} {
            lappend skipEntities $line
          }
        }
        close $skipFile
        file delete -force -- $cfile1
        errorMsg "File of entities to skip '[file tail $cfile1]' renamed to '[file tail $cfile]'."
      }
    }

# check if a file generated from a NIST test case (and some other files) is being processed
    set nistName [nistGetName]

# error opening file
  } emsg]} {
    if {$openStage == 2} {
      errorMsg "ERROR opening STEP file"

      if {!$p21e3} {
        set fext [string tolower [file extension $fname]]
        if {$fext != ".stp" && $fext != ".step" && $fext != ".p21" && $fext != ".stpz" && $fext != ".ifc"} {
          errorMsg "File extension not supported ([file extension $fname])" red
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
              set msg "\nThe STEP AP (schema) is not supported: $fs"
            }
            if {[info exists buttons]} {append msg "\n See Help > Supported STEP APs"}
            errorMsg $msg red

            if {[string first "IFC" $fs] == 0} {
              errorMsg "Use the IFC File Analyzer with IFC files."
              after 1000
              openURL https://www.nist.gov/services-resources/software/ifc-file-analyzer
            }

# other possible errors
          } else {
            set msg "\nPossible causes of the ERROR:"
            append msg "\n1 - Syntax errors in the STEP file"
            append msg "\n    The file must start with ISO-10303-21; and end with ENDSEC; END-ISO-10303-21;"
            append msg "\n    Try opening the file in a different STEP viewer, see Websites > STEP File Viewers"
            append msg "\n2 - File or directory name contains accented, non-English, or symbol characters"
            append msg "\n     [file nativename $fname]"
            append msg "\n    Change the file or directory name"
            append msg "\n3 - If the problem is not with the STEP file, then restart this software and try again,"
            append msg "\n    or run this software as administrator, or reboot your computer"
            append msg "\n\nFor other problems, contact: [join [getContact]]"
            errorMsg $msg red
          }
        }

# part 21 edition 3, but should not get to this point
      } else {
        outputMsg " "
        errorMsg "The STEP file uses ISO 10303 Part 21 Edition 3 and cannot be processed by this software.\n Edit the STEP file to delete the Edition 3 content such as the ANCHOR and REFERENCE sections."
      }

# open STEP file in editor
      if {$editorCmd != ""} {
        outputMsg " "
        errorMsg "Opening STEP file in text editor"
        exec $editorCmd [file nativename $localName] &
      }

      if {[info exists errmsg]} {unset errmsg}
      catch {$objDesign Delete}
      catch {unset objDesign}
      catch {unset objIFCsvr}
      catch {raise .}
      return 0

# other errors
    } elseif {$openStage == 1} {
      errorMsg "ERROR before opening STEP file: $emsg"
    } elseif {$openStage == 3} {
      errorMsg "ERROR after opening STEP file: $emsg"
    }
  }

# -------------------------------------------------------------------------------------------------
# connect to Excel
  set useXL 1
  set xlInstalled 1
  if {$opt(XLSCSV) != "None"} {
    if {[catch {
      set pid1 [checkForExcel $multiFile]
      set excel [::tcom::ref createobject Excel.Application]
      set pidExcel [lindex [intersect3 $pid1 [twapi::get_process_ids -name "EXCEL.EXE"]] 2]
      [$excel ErrorCheckingOptions] TextDate False
      set excelVersion [expr {int([$excel Version])}]

# file format, max rows
      set extXLS "xlsx"
      set xlFormat [expr 51]
      set rowmax [expr {2**20}]

# older Excel
      if {$excelVersion < 12} {
        set extXLS "xls"
        set xlFormat [expr 56]
        set rowmax [expr {2**16}]
        errorMsg "Some spreadsheet features are not compatible with older versions of Excel."
      }

# generate with Excel but save as CSV
      set saveCSV 0
      if {$opt(XLSCSV) == "CSV"} {
        set saveCSV 1
        catch {$buttons(ofExcel) configure -state disabled}
      } else {
        catch {$buttons(ofExcel) configure -state normal}
      }

# turning off ScreenUpdating, saves A LOT of time
      $excel Visible 0
      catch {$excel ScreenUpdating 0}

      set rowmax [expr {$rowmax-2}]
      if {$opt(XL_ROWLIM) < $rowmax} {set rowmax $opt(XL_ROWLIM)}

# no Excel, use CSV instead
    } emsg]} {
      set useXL 0
      set xlInstalled 0
      if {$opt(XLSCSV) == "Excel"} {
        errorMsg "Excel is not installed or cannot be started: $emsg\n CSV files will be generated instead of a spreadsheet.  See the Output Format option.  Some options are disabled."
        set opt(XLSCSV) "CSV"
        catch {raise .}
      }
      checkValues
      set ofExcel 0
      set ofCSV 1
      catch {$buttons(ofExcel) configure -state disabled}
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

# delete all but one worksheet
      catch {$excel DisplayAlerts False}
      set sheetCount [$worksheets Count]
      for {set n $sheetCount} {$n > 1} {incr n -1} {[$worksheets Item [expr $n]] Delete}
      set sheetLast [$worksheets Item [$worksheets Count]]
      catch {$excel DisplayAlerts True}
      [$excel ActiveWindow] TabRatio [expr 0.7]

# print errors
    } emsg]} {
      errorMsg "ERROR opening Excel workbooks and worksheets: $emsg"
      catch {raise .}
      return 0
    }

# CSV files or viz only
  } else {
    set rowmax [expr {2**20}]
    if {$opt(XL_ROWLIM) < $rowmax} {set rowmax $opt(XL_ROWLIM)}
  }

# -------------------------------------------------------------------------------------------------
# add header worksheet, for CSV files create directory and header file
  addHeaderWorksheet $numFile $fname

# -------------------------------------------------------------------------------------------------
# set Excel spreadsheet name, delete file if already exists

# user-defined file name
  if {$useXL} {
    set xlsmsg ""
    if {$opt(writeDirType) == 1} {
      if {$userXLSFile != ""} {
        set xlFileName [file nativename $userXLSFile]
      } else {
        append xlsmsg "User-defined Spreadsheet file name is not valid.  Spreadsheet directory and\n file name will be based on the STEP file. (Options tab)"
        set opt(writeDirType) 0
      }
    }

# same directory as file
    if {$opt(writeDirType) == 0} {
      set xlFileName "[file nativename [file join [file dirname $fname] [file rootname [file tail $fname]]]]-sfa.$extXLS"
      set xlFileNameOld "[file nativename [file join [file dirname $fname] [file rootname [file tail $fname]]]]_stp.$extXLS"

# user-defined directory
    } elseif {$opt(writeDirType) == 2} {
      set xlFileName "[file nativename [file join $writeDir [file rootname [file tail $fname]]]]-sfa.$extXLS"
      set xlFileNameOld "[file nativename [file join $writeDir [file rootname [file tail $fname]]]]_stp.$extXLS"
    }

# file name too long
    if {[string length $xlFileName] > 218} {
      if {[string length $xlsmsg] > 0} {append xlsmsg "\n\n"}
      append xlsmsg "Pathname of Spreadsheet file is too long for Excel ([string length $xlFileName])"
      set xlFileName "[file nativename [file join $writeDir [file rootname [file tail $fname]]]]-sfa.$extXLS"
      set xlFileNameOld "[file nativename [file join $writeDir [file rootname [file tail $fname]]]]_stp.$extXLS"
      if {[string length $xlFileName] < 219} {
        append xlsmsg "\nSpreadsheet file written to User-defined directory (Spreadsheet tab)"
      }
    }

# delete existing file
    if {[file exists $xlFileNameOld]} {catch {file delete -force -- $xlFileNameOld}}
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

# user-defined entity list
  catch {set userEntityList {}}
  if {$opt(PR_USER) && [llength $userEntityList] == 0 && [info exists userEntityFile]} {
    set userEntityList {}
    set fileUserEnt [open $userEntityFile r]
    while {[gets $fileUserEnt line] != -1} {
      set line [split [string trim $line] " "]
      foreach ent $line {lappend userEntityList [string tolower $ent]}
    }
    close $fileUserEnt
    if {[llength $userEntityList] == 0} {
      set opt(PR_USER) 0
      checkValues
    }
  }

# get totals of each entity in file
  set fixlist {}
  if {![info exists objDesign]} {return}
  catch {unset entCount}

# for all entity types, check for which ones to process
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

# user-defined entities
      set ok 0
      if {$opt(PR_USER) && [lsearch $userEntityList $entType] != -1} {set ok 1}

# STEP entities that are translated depending on the options
      set ok1 [setEntsToProcess $entType]
      if {$ok == 0} {set ok $ok1}

# entities in unsupported APs that are not AP203, AP214, AP242 - if not using a user-defined list or not generating a spreadsheet
      if {[string first "AP203" $stepAP] == -1 && [string first "AP214" $stepAP] == -1 && $stepAP != "AP242"} {
        if {!$opt(PR_USER) || $opt(XLSCSV) == "None"} {
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
        } elseif {[string first "AP209" $stepAP] != -1 && $opt(VIZFEA) && $opt(XLSCSV) != "None"} {
          outputMsg " "
          errorMsg "Viewing the AP209 FEM is not allowed when a User-Defined List is selected in the Options tab."
          set opt(VIZFEA) 0
          checkValues
        }
      }

# always process composite entities
      if {[lsearch $entCategory(PR_STEP_COMP) $entType] != -1} {
        set ok 1
        if {!$opt(PR_STEP_COMP)} {
          set opt(PR_STEP_COMP) 1
          outputMsg "\nComposites entities will be processed." red
        }
      }

# check for composite entities with "_11"
      if {$opt(PR_STEP_COMP) && $ok == 0} {if {[string first "_11" $entType] != -1} {set ok 1}}

# new AP242 entities in a ROSE file, but not yet in ap242all or any entity category, for testing new schemas
      #if {$developer} {if {$stepAP == "AP242" && [lsearch $ap242all $entType] == -1} {set ok 1}}

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
          lappend fixlist $entType
          lappend entsToIgnore $entType
          set entsIgnored($cidx$entType) $entCount($entType)
        }
      } elseif {[lsearch $entCategories $entType] != -1} {
        if {$noSkip} {
          lappend entsToProcess "$cidx$entType"
          incr numEnts $entCount($entType)
        } else {
          lappend fixlist $entType
          lappend entsToIgnore $entType
          set entsIgnored($cidx$entType) $entCount($entType)
        }
      } else {
        lappend entsToIgnore $entType
        set entsIgnored($cidx$entType) $entCount($entType)
      }
    }
  }

# -------------------------------------------------------------------------------------------------
# check if there is anything to view
  foreach typ {PMI TPG FEA} {set viz($typ) 0}
  if {$opt(VIZPMI)} {
    foreach ao $aoEntTypes {
      if {[info exists entCount($ao)]}  {if {$entCount($ao)  > 0} {set viz(PMI) 1}}
      set ao1 "$ao\_and_characterized_object"
      if {[info exists entCount($ao1)]} {if {$entCount($ao1) > 0} {set viz(PMI) 1}}
      set ao1 "$ao\_and_geometric_representation_item"
      if {[info exists entCount($ao1)]} {if {$entCount($ao1) > 0} {set viz(PMI) 1}}
    }
  }
  if {$opt(VIZTPG)} {if {[info exists entCount(tessellated_solid)] || [info exists entCount(tessellated_shell)]} {set viz(TPG) 1}}
  if {$opt(VIZFEA) && [string first "AP209" $stepAP] == 0} {set viz(FEA) 1}

# open expected PMI worksheet (once) if PMI representation and correct file name
  if {$opt(PMISEM) && $stepAP == "AP242" && $nistName != "" && $opt(XLSCSV) != "None"} {
    set tols [concat $tolNames [list dimensional_characteristic_representation datum datum_feature datum_reference_compartment datum_reference_element datum_system placed_datum_target_feature]]
    foreach tol $tols {if {[info exist entCount($tol)]} {set ok 1; break}}
    if {$ok && ![info exists nistPMImaster($nistName)]} {nistReadExpectedPMI}
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

# -------------------------------------------------------------------------------------------------
# list entities not processed based on fix file
  if {[llength $fixlist] > 0} {
    if {[file exists $cfile]} {
      set ok 0
      foreach item $fixlist {if {[lsearch $skipPerm $item] == -1} {set ok 1}}
    }
    if {$ok} {
      outputMsg " "
      if {$opt(XLSCSV) != "None"} {
        set msg "Worksheets"
        if {!$useXL} {set msg "CSV files"}
        append msg " will not be generated for the entity types listed in"
      } else {
        set msg "Views might not be generated because of the entity types listed in"
      }
      append msg " [truncFileName [file nativename $cfile]]"
      errorMsg $msg
      foreach item [lsort $fixlist] {outputMsg " [formatComplexEnt $item]" red}
      errorMsg "See Help > Crash Recovery"
    }
  }

# sort entsToProcess by color index
  set entsToProcess [lsort $entsToProcess]

# -------------------------------------------------------------------------------------------------
# for STEP process datum* and dimensional* entities before specific *_tolerance entities
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
        if {$tc == $entColorIndex(PR_STEP_TOLR)} {set itmp 1}
        if {[string first $entColorIndex(PR_STEP_TOLR) $str1] == 0 && ([string first "datum" $str1] == 2 || [string first "dimensional" $str1] == 2)} {
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
    if {[info exists entCount(dimensional_characteristic_representation)]} {
      set dcr "$entColorIndex(PR_STEP_TOLR)\dimensional_characteristic_representation"
      set c1 [lsearch $entsToProcess $dcr]
      set entsToProcess [lreplace $entsToProcess $c1 $c1]
      set entsToProcess [linsert $entsToProcess 0 $dcr]
    }
  }

# -------------------------------------------------------------------------------------------------
# move some entities to end of AP209 entities
  if {$viz(FEA)} {
    set ok  1
    set ok1 0
    set etp {}

# order is important, first 2 nodal loads, 3rd displacements (like a load), 4th boundary conditions
    set ent209 [list nodal_freedom_action_definition \
                     surface_3d_element_boundary_constant_specified_surface_variable_value \
                     volume_3d_element_boundary_constant_specified_variable_value \
                     nodal_freedom_values \
                     single_point_constraint_element_values]
    #set ent209 [list element_nodal_freedom_actions] not included
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
    if {$opt(VIZFEALV)} {
      lappend ents "nodal_freedom_action_definition"
      lappend ents "surface_3d_element_boundary_constant_specified_surface_variable_value"
      lappend ents "volume_3d_element_boundary_constant_specified_variable_value"
    }
    if {$opt(VIZFEADS)} {lappend ents "nodal_freedom_values"}
    if {$opt(VIZFEABC)} {lappend ents "single_point_constraint_element_values"}
    foreach ent $ents {if {[info exists entCount($ent)]} {set feaLastEntity $ent}}
  }

# then strip off the color index
  for {set i 0} {$i < [llength $entsToProcess]} {incr i} {
    lset entsToProcess $i [string range [lindex $entsToProcess $i] 2 end]
  }

# -------------------------------------------------------------------------------------------------
# max progress bar - number of entities or finite elements
  if {[info exists buttons]} {
    $buttons(pgb) configure -maximum $numEnts
    if {[string first "AP209" $stepAP] == 0 && $opt(XLSCSV) == "None"} {
      set n 0
      foreach elem {curve_3d_element_representation surface_3d_element_representation volume_3d_element_representation} {
        if {[info exists entCount($elem)]} {incr n $entCount($elem)}
      }
      $buttons(pgb) configure -maximum $n
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

            if {$item == "product_definition_formation"} {
              if {[string first "Y14.5" $val]  != -1 || [string first "1101" $val]  != -1} {lappend spmiTypesPerFile "dimensioning standard"}
              if {[string first "Y14.41" $val] != -1 || [string first "16792" $val] != -1} {lappend spmiTypesPerFile "modeling standard"}
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
    if {[string first "NIST_" $fn] == 0 && [string first "ASME" $fn] != -1} {errorMsg "All of the NIST models use the ASME Y14.5 tolerance standard."}
  }

# -------------------------------------------------------------------------------------------------
# generate worksheet for each entity
  outputMsg " "
  if {$useXL} {
    outputMsg "Generating STEP Entity worksheets" blue
  } elseif {$opt(XLSCSV) == "CSV"} {
    outputMsg "Generating STEP Entity CSV files" blue
  } elseif {$opt(XLSCSV) == "None"} {
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
    set nshape 0
    set ntable 0
    set savedViewName {}
    set savedViewNames {}
    set savedViewButtons {}
    set spmiEntity {}
    set spmiSumRow 1
    set stat 1
    set wsCount 0
    set x3dMsg {}
    foreach f {elements mesh meshIndex faceIndex} {catch {file delete -force -- [file join $mytemp $f.txt]}}

    if {[info exists dim]} {unset dim}
    set dim(prec,max) 0
    set dim(unit) ""
    set dim(unitOK) 1
    set dimRepeatDiv 2

# find camera models used in draughting model items and annotation_occurrence used in property_definition and datums
    if {$opt(PMIGRF) || $viz(PMI)} {pmiGetCamerasAndProperties}

# no entities to process
    if {[llength $entsToProcess] == 0} {
      if {$opt(XLSCSV) != "None"} {
        errorMsg "Select some other entity types to Process in the Options tab."
        catch {unset entsIgnored}
      } else {
        errorMsg "There is nothing in the STEP file to view based on the View selections (Options tab)."
      }
      break
    }
    set tlast [clock clicks -milliseconds]

# loop over list of entities in file
    foreach entType $entsToProcess {
      if {$opt(XLSCSV) != "None"} {
        set nerr1 0
        set lastEnt $entType

# decide if inverses should be checked for this entity type
        set checkInv 0
        if {$opt(INVERSE)} {set checkInv [invSetCheck $entType]}
        if {$checkInv} {lappend inverseEnts $entType}
        set badAttr [info exists badAttributes($entType)]

# process the entity type
        ::tcom::foreach objEntity [$objDesign FindObjects [join $entType]] {
          if {$entType == [$objEntity Type]} {
            incr nprogBarEnts
            if {[expr {$nprogBarEnts%1000}] == 0} {update}

            if {[catch {
              if {$useXL} {
                set stat [getEntity $objEntity $checkInv]
              } else {
                set stat [getEntityCSV $objEntity]
              }
            } emsg1]} {

# process errors with entity
              if {$stat != 1} {break}

              set msg "ERROR processing "
              if {[info exists objEntity]} {
                if {[string first "handle" $objEntity] != -1} {
                  append msg "\#[$objEntity P21ID]=[$objEntity Type] (row [expr {$row($thisEntType)+2}]): $emsg1"

# handle specific errors
                  if {[string first "Unknown error" $emsg1] != -1} {
                    errorMsg $msg
                    catch {raise .}
                    incr nerr1
                    if {$nerr1 > 20} {
                      errorMsg "Processing of $entType entities has stopped" red
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
              if {[string first "element_representation" $thisEntType] != -1 && $opt(VIZFEA)} {set ok 0}
              if {$ok} {set nprogBarEnts [expr {$nprogBarEnts + $entCount($thisEntType) - $count($thisEntType)}]}
              break
            }
          }
        }

# close CSV file
        if {!$useXL} {catch {close $fcsv}}
      }

# check for reports (validation properties, PMI presentation and representation, tessellated geometry, AP209 FEM)
      checkForReports $entType
    }

  } emsg2]} {
    catch {raise .}
    if {[llength $entsToProcess] > 0} {
      set msg "ERROR processing STEP file: "
      if {[info exists objEntity]} {if {[string first "handle" $objEntity] != -1} {append msg " \#[$objEntity P21ID]=[$objEntity Type]"}}
      append msg "\n $emsg2"
      append msg "\nProcessing of the STEP file has stopped"
      errorMsg $msg
    } else {
      return
    }
  }

# -------------------------------------------------------------------------------------------------
# check fix file
  if {[info exists cfile]} {
    set fixtmp {}
    if {[file exists $cfile]} {
      set skipFile [open $cfile r]
      while {[gets $skipFile line] >= 0} {
        if {[lsearch $fixtmp $line] == -1 && $line != $lastEnt} {lappend fixtmp $line}
      }
      close $skipFile
    }

    if {[join $fixtmp] == ""} {
      catch {file delete -force -- $cfile}
    } else {
      set skipFile [open $cfile w]
      foreach item $fixtmp {puts $skipFile $item}
      close $skipFile
    }
  }

# -------------------------------------------------------------------------------------------------
# generate b-rep part geometry if no other viz exists
  set vizbrp 0
  if {$opt(VIZBRP) && !$viz(PMI) && !$viz(FEA) && !$viz(TPG)} {
    set ok 0
    foreach item $brepEnts {if {[info exists entCount($item)]} {set ok 1}}
    if {$ok} {
      x3dFileStart
      set vizbrp 1
      update
    }
  }

# generate b-rep part geom, set viewpoints, and close X3DOM geometry file
  if {($viz(PMI) || $viz(FEA) || $viz(TPG) || $vizbrp) && $x3dFileName != ""} {x3dFileEnd}

# -------------------------------------------------------------------------------------------------
# add summary worksheet
  if {$useXL} {
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
# add PMI Rep. Coverage Analysis worksheet for a single file
    if {$opt(PMISEM)} {
      if {[info exists spmiTypesPerFile]} {
        set spmiCoverageWS "PMI Representation Coverage"
        if {![info exists worksheet($spmiCoverageWS)]} {
          outputMsg " Adding PMI Representation Coverage worksheet" blue
          spmiCoverageStart 0
          spmiCoverageWrite "" "" 0
          spmiCoverageFormat "" 0
        }
      }

# format PMI Representation Summary worksheet
      if {[info exists spmiSumName]} {
        if {$nistName != "" && [info exists nistPMIexpected($nistName)]} {nistPMISummaryFormat}
        [$worksheet($spmiSumName) Columns] AutoFit
        [$worksheet($spmiSumName) Rows] AutoFit
      }
      catch {unset spmiSumName}
    }

# add PMI Pres. Coverage Analysis worksheet for a single file
    if {$opt(PMIGRF) && $opt(XLSCSV) != "None" && $opt(PMIGRFCOV)} {
      if {[info exists gpmiTypesPerFile]} {
        set gpmiCoverageWS "PMI Presentation Coverage"
        if {![info exists worksheet($gpmiCoverageWS)]} {
          outputMsg " Adding PMI Presentation Coverage worksheet" blue
          gpmiCoverageStart 0
          gpmiCoverageWrite "" "" 0
          gpmiCoverageFormat "" 0
        }
      }
    }

# add ANCHOR and other sections from Part 21 Edition 3
    if {[info exists p21e3Section]} {
      if {[llength $p21e3Section] > 0} {addP21e3Section}
    }
# -------------------------------------------------------------------------------------------------
# select the first tab
    [$worksheets Item [expr 1]] Select
    [$excel ActiveWindow] ScrollRow [expr 1]
  }

# -------------------------------------------------------------------------------------------------
# quit IFCsvr, but not sure how to do it properly
  if {[catch {
    $objDesign Delete
    unset objDesign
    unset objIFCsvr

# errors
  } emsg]} {
    errorMsg "ERROR closing IFCsvr: $emsg"
    catch {raise .}
  }

# processing time
  set cc [clock clicks -milliseconds]
  set proctime [expr {($cc - $lasttime)/1000}]
  if {$proctime <= 60} {set proctime [expr {(($cc - $lasttime)/100)/10.}]}
  outputMsg "Processing time: $proctime seconds"
  update

  incr filesProcessed
  saveState
  
# -------------------------------------------------------------------------------------------------
# save spreadsheet
  set csvOpenDir 0
  if {$useXL} {
    if {[catch {
      outputMsg " "
      if {$xlsmsg != ""} {outputMsg $xlsmsg red}
      if {[string first "\[" $xlFileName] != -1} {
        regsub -all {\[} $xlFileName "(" xlFileName
        regsub -all {\]} $xlFileName ")" xlFileName
        outputMsg "In the spreadsheet file name, the characters \'\[\' and \'\]\' have been\n substituted by \'\(\' and \'\)\'" red
      }
      set xlfn $xlFileName

# create new file name if spreadsheet already exists, delete new file name spreadsheets if possible
      if {[file exists $xlfn]} {set xlfn [incrFileName $xlfn]}

# always save as spreadsheet
      outputMsg "Saving Spreadsheet as:"
      outputMsg " [truncFileName $xlfn 1]" blue
      if {[catch {
        catch {$excel DisplayAlerts False}
        $workbook -namedarg SaveAs Filename $xlfn FileFormat $xlFormat
        catch {$excel DisplayAlerts True}
        set lastXLS $xlfn
        lappend xlFileNames $xlfn
      } emsg1]} {
        errorMsg "ERROR Saving Spreadsheet: $emsg1"
      }

# save worksheets as CSV files
      if {$saveCSV} {
        if {[catch {
          set csvdirnam "[file join [file dirname $localName] [file rootname [file tail $localName]]]-sfa-csv"
          file mkdir $csvdirnam
          outputMsg "Saving Spreadsheet as multiple CSV files to directory:"
          outputMsg " [truncFileName [file nativename $csvdirnam]]" blue
          set csvFormat [expr 6]
          if {$excelVersion > 15} {set csvFormat [expr 62]}

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
            catch {file delete -force -- $csvfname}
            if {[file exists $csvfname]} {set csvfname [incrFileName $csvfname]}
            if {[string first "PMI-Representation" $csvfname] != -1 && $excelVersion < 16} {
              errorMsg "PMI symbols written to CSV files will look correct only with Excel 2016 or newer." red
            }
            $workbook -namedarg SaveAs Filename [file rootname $csvfname] FileFormat $csvFormat
            incr nprogBarEnts
            update
          }
        } emsg2]} {
          errorMsg "ERROR Saving CSV files: $emsg2"
        }
      }

      catch {$excel ScreenUpdating 1}

# close Excel
      $excel Quit
      set openxl 1
      catch {unset excel}
      catch {if {[llength $pidExcel] == 1} {twapi::end_process $pidExcel -force}}

# add Link(n) text to multi file summary
      if {$numFile != 0 && [info exists cells1(Summary)]} {
        set colsum [expr {$col1(Summary)+1}]
        if {$opt(XL_LINK1)} {
          $cells1(Summary) Item 3 $colsum "Link ($numFile)"
          set range [$worksheet1(Summary) Range [cellRange 3 $colsum]]
          regsub -all {\\} $xlFileName "/" xls
          [$worksheet1(Summary) Hyperlinks] Add $range [join $xls] [join ""] [join "Link to Spreadsheet"]
        } else {
          $cells1(Summary) Item 3 $colsum "$numFile"
        }
      }
      update idletasks

# errors
    } emsg]} {
      errorMsg "ERROR: $emsg"
      catch {raise .}
      set openxl 0
    }

# -------------------------------------------------------------------------------------------------
# open spreadsheet or directory of CSV files
    set ok 0
    if {$openxl && $opt(XL_OPEN)} {
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
      } elseif {!$opt(XL_OPEN) && $numFile == 0 && [string first "STEP-File-Analyzer.exe" $scriptName] != -1} {
        outputMsg " Use F2 to open the Spreadsheet (see Options tab, Help > Function Keys)" red
      }
    }

# CSV files generated too
    if {$saveCSV} {set csvOpenDir 1}

# open directory of CSV files
  } elseif {$opt(XLSCSV) != "None"} {
    set csvOpenDir 1
    unset csvfile
    outputMsg "\nCSV files written to:"
    outputMsg " [truncFileName [file nativename $csvdirnam]]" blue
  }

  if {$opt(XLSCSV) == "None"} {set useXL 1}

# open directory of CSV files
  if {$csvOpenDir} {
    set ok 0
    if {$opt(XL_OPEN)} {
      if {$numFile == 0} {
        set ok 1
      } elseif {[info exists lenfilelist]} {
        if {$lenfilelist == 1} {set ok 1}
      }
    }
    if {$ok} {
      set dir [file nativename $csvdirnam]
      if {[string first " " $dir] == -1} {
        outputMsg "Opening CSV file directory"
        if {[catch {
          exec {*}[auto_execok start] $dir
        } emsg]} {
          if {[string first "UNC" $emsg] != -1} {set emsg [fixErrorMsg $emsg]}
          if {$emsg != ""} {errorMsg "ERROR opening CSV file directory: $emsg"}
        }
      } else {
        exec C:/Windows/explorer.exe $dir &
      }
    }
  }

# -------------------------------------------------------------------------------------------------
# open X3DOM file of graphical PMI or FEM
  openX3DOM "" $numFile

# save log file
  if {[info exists logFile]} {
    update idletasks
    outputMsg "\nSaving Log file as:"
    outputMsg " [truncFileName [file nativename $lfile]]" blue
    close $logFile
    if {!$multiFile && [info exists buttons]} {
      set currLogFile $lfile
      bind . <Key-F4> {
        if {[file exists $currLogFile]} {
          outputMsg "\nOpening Log file: [file tail $currLogFile]"
          exec $editorCmd [file nativename $currLogFile] &
        }
      }
    }
    unset lfile
    unset logFile
  } elseif {[info exists buttons]} {
    if {[info exists currLogFile]} {unset currLogFile}
    bind . <Key-F4> {}
  }

# -------------------------------------------------------------------------------------------------
# save state
  if {[info exists errmsg]} {unset errmsg}
  saveState
  if {!$multiFile && [info exists buttons]} {$buttons(genExcel) configure -state normal}
  update idletasks

# unset variables to release memory and/or to reset them
  global cgrObjects colColor coordinatesList currx3dPID datumGeom datumIDs datumSymbol dimrep dimrepID dimtolEnt dimtolEntID dimtolGeom entName
  global feaDOFR feaDOFT feaNodes gpmiID gpmiIDRow gpmiRow heading invCol invGroup lineStrips nrep numx3dPID pmiColumns pmiStartCol
  global propDefID propDefIDRow propDefName propDefOK propDefRow savedsavedViewNames savedViewFile savedViewFileName shapeRepName
  global srNames suppGeomEnts syntaxErr tessPlacement tessRepo

  foreach var {cells cgrObjects colColor coordinatesList count currx3dPID datumGeom datumIDs datumSymbol dimrep dimrepID dimtolEnt dimtolEntID dimtolGeom \
               entName entsIgnored feaDOFR feaDOFT feaNodes gpmiID gpmiIDRow gpmiRow heading invCol invGroup lineStrips nrep numx3dPID \
               pmiCol pmiColumns pmiStartCol pmivalprop propDefID propDefIDRow propDefName propDefOK propDefRow savedsavedViewNames \
               savedViewFile savedViewFileName savedViewNames shapeRepName srNames suppGeomEnts syntaxErr tessPlacement tessRepo \
               workbook workbooks worksheet worksheets x3dCoord x3dFile x3dFileName x3dIndex x3dMax x3dMin x3dStartFile} {
    if {[info exists $var]} {unset $var}
  }
  if {!$multiFile} {foreach var {gpmiTypesPerFile spmiTypesPerFile} {if {[info exists $var]} {unset $var}}}

# delete leftover text files
  foreach f [glob -nocomplain -directory $mytemp *.txt] {catch {file delete -force -- $f}}
  update idletasks
  return 1
}

# -------------------------------------------------------------------------------------------------
proc addHeaderWorksheet {numFile fname} {
  global objDesign
  global ap242edition cadApps cadSystem cells cells1 col1 csvdirnam excel excel1 fileSchema legendColor
  global localName opt p21e3 row spmiTypesPerFile timeStamp useXL worksheet worksheet1 worksheets

  if {[catch {
    set cadSystem ""
    set timeStamp ""
    set p21e3 0

    set hdr "Header"
    if {$useXL} {
      outputMsg "Generating Header worksheet" blue
      set worksheet($hdr) [$worksheets Item [expr 1]]
      $worksheet($hdr) Activate
      $worksheet($hdr) Name $hdr
      set cells($hdr) [$worksheet($hdr) Cells]

# create directory for CSV files
    } elseif {$opt(XLSCSV) != "None"} {
      outputMsg "Generating Header CSV file" blue
      foreach var {csvdirnam csvfname fcsv} {catch {unset $var}}
      set csvdirnam "[file join [file dirname $localName] [file rootname [file tail $localName]]]-sfa-csv"
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
      } elseif {$opt(XLSCSV) != "None"} {
        set csvstr $attr
      }
      set objAttr [string trim [join [$objDesign $attr]]]

# FileDirectory
      if {$attr == "FileDirectory"} {
        if {$useXL} {
          $cells($hdr) Item $row($hdr) 2 [$objDesign $attr]
        } elseif {$opt(XLSCSV) != "None"} {
          append csvstr ",[$objDesign $attr]"
          puts $fcsv $csvstr
        }
        outputMsg "$attr:  [$objDesign $attr]"

# SchemaName
      } elseif {$attr == "SchemaName"} {
        set sn $fileSchema
        if {$useXL} {
          $cells($hdr) Item $row($hdr) 2 $sn
        } elseif {$opt(XLSCSV) != "None"} {
          append csvstr ",$sn"
          puts $fcsv $csvstr
        }
        outputMsg "$attr:  $sn" blue

# check edition of AP242
        set ap242edition 1
        if {[string first "1 0 10303 442" $sn] != -1} {
          if {[string first "1 0 10303 442 2 1 4" $sn] != -1} {
            errorMsg "This file uses the new AP242 Edition 2."
            set ap242edition 2
          } elseif {[string first "1 0 10303 442 3 1 4" $sn] != -1} {
            errorMsg "This file uses the new AP242 Edition 3."
            set ap242edition 3
          } elseif {[string first "1 0 10303 442 1 1 4" $sn] == -1} {
            errorMsg "This file uses an older or unknown version of AP242."
          }

# check version of AP203, AP214
        } elseif {[string first "CONFIG_CONTROL_DESIGN" $sn] == 0 || [string first "CONFIGURATION_CONTROL_3D_DESIGN" $sn] == 0} {
          errorMsg "This file uses an older version of STEP AP203.  See Help > Supported STEP APs"
        } elseif {[string first "AUTOMOTIVE_DESIGN_CC2" $sn] == 0} {
          errorMsg "This file uses an older version of STEP AP214.  See Help > Supported STEP APs"
        }

# check for IFC or CIS/2 files
        set fschema [string toupper [string range $objAttr 0 5]]
        if {[string first "IFC" $fschema] == 0} {
          errorMsg "Use the IFC File Analyzer with IFC files."
          after 1000
          openURL https://www.nist.gov/services-resources/software/ifc-file-analyzer
        } elseif {$objAttr == "STRUCTURAL_FRAME_SCHEMA"} {
          errorMsg "Use SteelVis to view CIS/2 files.  https://go.usa.gov/s8fm"
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
            } elseif {$opt(XLSCSV) != "None"} {
              append str2 ",[string trim $item]"
            }
          }
          outputMsg [string range $str1 0 end-2]
          if {$useXL} {
            $cells($hdr) Item $row($hdr) 2 "'[string trim $str2]"
            set range [$worksheet($hdr) Range "$row($hdr):$row($hdr)"]
            $range VerticalAlignment [expr -4108]
          } elseif {$opt(XLSCSV) != "None"} {
            append csvstr [string trim $str2]
            puts $fcsv $csvstr
          }
        } else {
          outputMsg "$attr:  $objAttr"
          if {$useXL} {
            $cells($hdr) Item $row($hdr) 2 "'$objAttr"
            set range [$worksheet($hdr) Range "$row($hdr):$row($hdr)"]
            $range VerticalAlignment [expr -4108]
          } elseif {$opt(XLSCSV) != "None"} {
            append csvstr ",$objAttr"
            puts $fcsv $csvstr
          }
        }

# check implementation level
        if {$attr == "FileImplementationLevel"} {
          if {[string first "\;" $objAttr] == -1} {
            errorMsg "FileImplementationLevel is usually '2\;1', see Header worksheet"
            if {$useXL} {[[$worksheet($hdr) Range B4] Interior] Color $legendColor(red)}
          } elseif {$objAttr == "4\;1"} {
            set p21e3 1
            if {[string first "p21e2" $fname] == -1} {errorMsg "This file uses ISO 10303 Part 21 Edition 3 with possible ANCHOR, REFERENCE, SIGNATURE, or TRACE sections."}
          }
        }

# check and add time stamp to multi file summary
        if {$attr == "FileTimeStamp"} {
          if {([string first "-" $objAttr] == -1 || [string length $objAttr] < 17 || [string length $objAttr] > 25) && $objAttr != ""} {
            errorMsg "FileTimeStamp has the wrong format, see Header worksheet"
            if {$useXL} {[[$worksheet($hdr) Range B5] Interior] Color $legendColor(red)}
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

    if {$useXL} {
      [[$worksheet($hdr) Range "A:A"] Font] Bold [expr 1]
      [$worksheet($hdr) Columns] AutoFit
      [$worksheet($hdr) Rows] AutoFit
    }

# check for CAx-IF Recommended Practices in the file description
    set caxifrp {}
    foreach fd [$objDesign "FileDescription"] {
      set c1 [string first "CAx-IF Rec." $fd]
      if {$c1 != -1} {lappend caxifrp [string trim [string range $fd $c1+20 end]]}
    }
    if {[llength $caxifrp] > 0} {
      outputMsg "\nCAx-IF Recommended Practices: (www.cax-if.org/joint_testing_info.html#recpracs)" blue
      foreach item $caxifrp {
        outputMsg " $item"
        lappend spmiTypesPerFile "document identification"
        if {[string first "AP242" $fschema] == -1 && [string first "Tessellated" $item] != -1} {
          errorMsg "  Error: Recommended Practices related to 'Tessellated' only apply to AP242 files."
        }
      }
    }

# set the application from various file attributes, cadApps is a list of all apps defined in sfa-data.tcl, take the first one that matches
    set ok 0
    set app2 ""
    set fos [$objDesign FileOriginatingSystem]
    set fpv [$objDesign FilePreprocessorVersion]
    foreach attr {FileOriginatingSystem FilePreprocessorVersion FileDescription FileAuthorisation FileOrganization} {
      foreach app $cadApps {
        set app1 $app
        if {$cadSystem == "" && [string first [string tolower $app] [string tolower [join [$objDesign $attr]]]] != -1} {
          set cadSystem [join [$objDesign $attr]]

# for multiple files, modify the app string to fit in file summary worksheet
          if {$app == "3D Evolution"}           {set app1 "CT 3D Evolution"}
          if {$app == "CoreTechnologie"}        {set app1 "CT 3D Evolution"}
          if {$app == "DATAKIT"}                {set app1 "Datakit"}
          if {$app == "EDMsix"}                 {set app1 "Jotne EDMsix"}
          if {$app == "Implementor Forum Team"} {set app1 "CAx-IF"}
          if {$app == "PRO/ENGINEER"}           {set app1 "Pro/E"}
          if {$app == "SOLIDWORKS"}             {set app1 "SolidWorks"}
          if {$app == "SOLIDWORKS MBD"}         {set app1 "SolidWorks MBD"}
          if {$app == "3D Reviewer"}            {set app1 "TechSoft3D 3D_Reviewer"}

          if {$app == "UGS - NX"}                {set app1 "UGS-NX"}
          if {$app == "UNIGRAPHICS"}             {set app1 "Unigraphics"}
          if {$app == "jt_step translator"}      {set app1 "Siemens NX"}
          if {$app == "SIEMENS PLM Software NX"} {set app1 "Siemens NX"}

          if {[string first "CATIA Version" $app] == 0} {set app1 "CATIA V[string range $app 14 end]"}
          if {$app == "3D EXPERIENCE"} {set app1 "3D Experience"}
          if {$app == "3DEXPERIENCE"}  {set app1 "3D Experience"}

          if {[string first "CATIA SOLUTIONS V4"      $fos] != -1} {set app1 "CATIA V4"}
          if {[string first "Autodesk Inventor"       $fos] != -1} {set app1 $fos}
          if {[string first "FreeCAD"                 $fos] != -1} {set app1 "FreeCAD"}
          if {[string first "SIEMENS PLM Software NX" $fos] ==  0} {set app1 "Siemens NX_[string range $fos 24 end]"}

          if {[string first "THEOREM"   $fpv] != -1} {set app1 "Theorem Solutions"}
          if {[string first "T-Systems" $fpv] != -1} {set app1 "T-Systems"}

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

# add app2 to multiple file summary worksheet
    if {$numFile != 0 && $useXL && [info exists cells1(Summary)]} {
      if {$ok == 0} {set app2 [setCAXIFvendor]}
      set colsum [expr {$col1(Summary)+1}]
      if {$colsum > 16} {[$excel1 ActiveWindow] ScrollColumn [expr {$colsum-16}]}
      regsub -all " " $app2 [format "%c" 10] app2
      $cells1(Summary) Item 6 $colsum [string trim $app2]
    }
    set cadSystem $app2
    if {$cadSystem == ""} {set cadSystem [setCAXIFvendor]}

# close csv file
    if {!$useXL && $opt(XLSCSV) != "None"} {close $fcsv}

  } emsg]} {
    errorMsg "ERROR adding Header worksheet: $emsg"
    catch {raise .}
  }
}

#-------------------------------------------------------------------------------------------------
# add summary worksheet
proc sumAddWorksheet {} {
  global andEntAP209 cells col entCategory entCount entsIgnored excel gpmiEnts nistVersion
  global opt row sheetLast sheetSort spmiEntity stepAP sum worksheet worksheets

  outputMsg "\nGenerating Summary worksheet" blue
  set sum "Summary"

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
  #set ws_nsort [lsort $sheetSort]

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

# for STEP add [Properties], [PMI Presentation], [PMI Representation] text string
        set okao 0
        if {$entType == "property_definition" && $col($entType) > 4 && $opt(VALPROP)} {
          $cells($sum) Item $sumRow 1 "property_definition  \[Properties\]"
        } elseif {$entType == "dimensional_characteristic_representation" && $col($entType) > 3 && $opt(PMISEM)} {
          $cells($sum) Item $sumRow 1 "dimensional_characteristic_representation  \[PMI Representation\]"
        } elseif {[lsearch $spmiEntity $entType] != -1 && $opt(PMISEM)} {
          $cells($sum) Item $sumRow 1 "$entType  \[PMI Representation\]"
        } elseif {[string first "annotation" $entType] != -1 && $opt(PMIGRF)} {
          if {$gpmiEnts($entType) && $col($entType) > 5} {set okao 1}
        }
        if {$okao} {
          $cells($sum) Item $sumRow 1 "$entType  \[PMI Presentation\]"
        }

# for '_and_' (complex entity) split on multiple lines
# '10' is the ascii character for a linefeed
      } else {
        regsub -all "_and_" $entType ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" entType_multiline
        set entType_multiline "($entType_multiline)"
        $cells($sum) Item $sumRow 1 $entType_multiline

# for STEP add [Properties] or [PMI Presentation] text string
        set okao 0
        if {[string first "annotation" $entType] != -1} {
          if {$gpmiEnts($entType) && $col($entType) > 7} {set okao 1}
        } elseif {[lsearch $spmiEntity $entType] != -1} {
          $cells($sum) Item $sumRow 1 "$entType_multiline  \[PMI Representation\]"
        }
        if {$okao} {
          $cells($sum) Item $sumRow 1 "$entType_multiline  \[PMI Presentation\]"
        }
        set range [$worksheet($sum) Range $sumRow:$sumRow]
        $range VerticalAlignment [expr -4108]
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
    set str "NIST "
    set url "https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer"
    if {!$nistVersion} {
      set str ""
      set url "https://github.com/usnistgov/SFA"
    }
    $cells($sum) Item [expr {$row($sum)+2}] 1 "$str\STEP File Analyzer and Viewer (v[getVersion])"
    set anchor [$worksheet($sum) Range [cellRange [expr {$row($sum)+2}] 1]]
    [$worksheet($sum) Hyperlinks] Add $anchor [join $url] [join ""] [join "Link to $str\STEP File Analyzer and Viewer"]
    $cells($sum) Item [expr {$row($sum)+3}] 1 "[clock format [clock seconds]]"

# print errors
  } emsg]} {
    errorMsg "ERROR adding Summary worksheet: $emsg"
    catch {raise .}
  }
  return [list $sumLinks $sheetSort $sumRow]
  #return [list $sumLinks $sumDocCol $sheetSort $sumRow]
}

#-------------------------------------------------------------------------------------------------
# add file name and other info to top of Summary
proc sumAddFileName {sum sumLinks} {
  global cadSystem cells dim entityCount fileSchema localName opt schemaLinks stepAP timeStamp tolStandard worksheet xlFileName

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
        #set range [$worksheet($sum) Range "1:1"]
        #$range VerticalAlignment [expr -4108]
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
      if {[info exists schemaLinks($stepAP)]} {
        set anchor [$worksheet($sum) Range "B1"]
        $sumLinks Add $anchor $schemaLinks($stepAP) [join ""] [join "Link to $stepAP schema documentation"]
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
    if {$opt(XL_LINK1)} {
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
    errorMsg "ERROR adding File Names to Summary: $emsg"
    catch {raise .}
  }
  return $sumHeaderRow
}

#-------------------------------------------------------------------------------------------------
# add file name and other info to top of Summary
proc sumAddColorLinks {sum sumHeaderRow sumLinks sheetSort sumRow} {
  global cells col entName entsIgnored entsWithErrors excel row worksheet xlFileName

  if {[catch {
    #outputMsg " Adding links on Summary to Entity worksheets"
    set row($sum) [expr {$sumHeaderRow+2}]

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
      $sumLinks Add $anchor $xlFileName "$hlsheet!A1" "Go to $ent"

# color cells
      set cidx [setColorIndex $ent]
      if {$cidx > 0} {

# color entities on summary if no errors or warnings and add comment that there are CAx-IF RP errors
        if {[lsearch $entsWithErrors [formatComplexEnt $ent]] == -1} {
          [$anchor Interior] ColorIndex [expr $cidx]

# color entities on summary gray and add comment that there are CAx-IF RP errors
        } else {
          [$anchor Interior] ColorIndex [expr 15]
          if {$ent != "dimensional_characteristic_representation"} {
            addCellComment $sum $sumRow 1 "There are errors or warnings for this entity based on CAx-IF Recommended Practices.  See Help > Syntax Errors."
          } else {
            addCellComment $sum $sumRow 1 "There are errors or warnings for this entity based on CAx-IF Recommended Practices.  Check for cell comments in the Associated Geometry column.  See Help > Syntax Errors."
          }
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
    errorMsg "ERROR adding Summary colors and links: $emsg"
    catch {raise .}
  }
}

#-------------------------------------------------------------------------------------------------
# format worksheets
proc formatWorksheets {sheetSort sumRow inverseEnts} {
  global buttons cells col count entCount excel excelVersion gpmiEnts nprogBarEnts opt pmiStartCol
  global row rowmax spmiEnts stepAP syntaxErr thisEntType viz worksheet xlFileName
  outputMsg "Formatting Worksheets" blue

  if {[info exists buttons]} {$buttons(pgb) configure -maximum [llength $sheetSort]}
  set nprogBarEnts 0
  set nsort 0

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
        #if {$thisEntType == "element_nodal_freedom_actions"} {moveWorksheet [list element_nodal_freedom_actions element_nodal_freedom_terms]}
      }

# find extent of columns and rows
      set rancol [[[$worksheet($thisEntType) UsedRange] Columns] Count]
      set ranrow [expr {$row($thisEntType)+2}]
      if {$ranrow > $rowmax} {set ranrow [expr {$rowmax+2}]}
      set ranrow [expr {$ranrow-2}]

# autoformat
      set range [$worksheet($thisEntType) Range [cellRange 3 1] [cellRange $ranrow $rancol]]
      $range AutoFormat

# freeze panes
      [$worksheet($thisEntType) Range "B4"] Select
      catch {[$excel ActiveWindow] FreezePanes [expr 1]}

# set A1 as default cell, was A4
      [$worksheet($thisEntType) Range "A1"] Select

# set column color, border, group for INVERSES and Used In
      if {$opt(INVERSE)} {if {[lsearch $inverseEnts $thisEntType] != -1} {invFormat $rancol}}

# STEP Property_definition (Validation Properties)
      if {$thisEntType == "property_definition" && $opt(VALPROP)} {
        valPropFormat

# color STEP annotation occurrence (Graphical PMI)
      } elseif {$gpmiEnts($thisEntType) && $opt(PMIGRF)} {
        pmiFormatColumns "PMI Presentation"

# color STEP semantic PMI
      } elseif {$spmiEnts($thisEntType) && $opt(PMISEM)} {
        pmiFormatColumns "PMI Representation"

# add PMI Representation Summary worksheet
        spmiSummary
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
      if {$thisEntType == "property_definition"} {set c 4}
      if {$c > 8} {set c 8}
      if {$c == 1} {set c 2}
      set range [$worksheet($thisEntType) Range [cellRange 1 1] [cellRange 1 $c]]
      $range MergeCells [expr 1]

# link back to summary
      set anchor [$worksheet($thisEntType) Range "A1"]
      $hlink Add $anchor $xlFileName "Summary!A$sumRow" "Return to Summary"

# check width of columns, wrap text
      if {[catch {
        set widlim 400.
        for {set i 2} {$i <= $rancol} {incr i} {
          if {[[$cells($thisEntType) Item 3 $i] Value] != ""} {
            set wid [[$cells($thisEntType) Item 3 $i] Width]
            if {$wid > $widlim} {
              set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
              $range ColumnWidth [expr {[$range ColumnWidth]/$wid * $widlim}]
              $range WrapText [expr 1]
            }
          }
        }
      } emsg]} {
        errorMsg "ERROR setting column widths: $emsg\n  $thisEntType"
        catch {raise .}
      }

# color red for syntax errors
      if {[info exists syntaxErr($thisEntType)]} {colorBadCells $thisEntType}

# -------------------------------------------------------------------------------------------------
# add table for sorting and filtering
      if {$excelVersion > 11} {
        if {[catch {
          if {$opt(XL_SORT) && $thisEntType != "property_definition"} {
            if {$ranrow > 8} {
              set range [$worksheet($thisEntType) Range [cellRange 3 1] [cellRange $ranrow $rancol]]
              set tname [string trim "TABLE-$thisEntType"]
              [[$worksheet($thisEntType) ListObjects] Add 1 $range] Name $tname
              [[$worksheet($thisEntType) ListObjects] Item $tname] TableStyle "TableStyleLight1"
              if {[incr ntable] == 1 && $opt(XL_SORT)} {outputMsg " Generating Tables for Sorting" blue}
            }
          }
        } emsg]} {
          errorMsg "ERROR adding Tables for Sorting: $emsg"
          catch {raise .}
        }
      }

# errors
    } emsg]} {
      errorMsg "ERROR formatting Spreadsheet for: $thisEntType\n$emsg"
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
    errorMsg "ERROR moving worksheet: $emsg"
  }
}

# -------------------------------------------------------------------------------------------------
proc addP21e3Section {} {
  global objDesign
  global cells legendColor p21e3Section worksheet worksheets

# look for three section types possible in Part 21 Edition 3
  foreach line $p21e3Section {
    if {$line == "ANCHOR" || $line == "REFERENCE" || $line == "SIGNATURE"} {
      set sect $line
      set worksheet($sect) [$worksheets Add [::tcom::na] [$worksheets Item [$worksheets Count]]]
      set n [$worksheets Count]
      [$worksheets Item [expr $n]] -namedarg Move Before [$worksheets Item [expr 3]]
      $worksheet($sect) Activate
      $worksheet($sect) Name $sect
      set cells($sect) [$worksheet($sect) Cells]
      set r 0
      outputMsg " Adding $line worksheet" green
    }

# add to worksheet
    incr r
    $cells($sect) Item $r 1 $line

# process anchor
    if {$sect == "ANCHOR"} {
      if {$r == 1} {$cells($sect) Item $r 2 "Entity"}
      set c2 [string first ";" $line]
      if {$c2 != -1} {set line [string range $line 0 $c2-1]}

      set c1 [string first "\#" $line]
      if {$c1 != -1} {
        set badEnt 0
        set anchorID [string range $line $c1+1 end]
        if {[string is integer $anchorID]} {
          set anchorEnt [[$objDesign FindObjectByP21Id [expr {int($anchorID)}]] Type]

# add anchor ID to entity worksheet
          if {$anchorEnt != ""} {
            $cells($sect) Item $r 2 $anchorEnt
            if {[info exists worksheet($anchorEnt)]} {
              set c3 [string first ">" $line]
              if {$c3 == -1} {set c3 [string first "=" $line]}
              set uuid [string range $line 1 $c3-1]
              if {![info exists urow($anchorEnt)]} {set urow($anchorEnt) [[[$worksheet($anchorEnt) UsedRange] Rows] Count]}
              if {![info exists ucol($anchorEnt)]} {set ucol($anchorEnt) [getNextUnusedColumn $anchorEnt]}
              for {set ur 4} {$ur <= $urow($anchorEnt)} {incr ur} {
                set id [[$cells($anchorEnt) Item $ur 1] Value]
                if {$id == $anchorID} {
                  $cells($anchorEnt) Item $ur $ucol($anchorEnt) $uuid
                  break
                }
              }
            } else {
              lappend noanchor $anchorEnt
            }
          } else {
            set badEnt 1
          }
        } else {
          set badEnt 1
        }

# bad ID in anchor section
        if {$badEnt} {
          [[$worksheet($sect) Range [cellRange $r 1] [cellRange $r 1]] Interior] Color $legendColor(red)
          errorMsg "Syntax Error: Bad format for entity ID in ANCHOR section."
        }
      }
    }

    if {$line == "ENDSEC"} {[$worksheet($sect) Columns] AutoFit}
  }

# format ID columns
  if {[llength [array names urow]] > 0} {
    set ids {}
    foreach item [lsort [array names urow]] {lappend ids [formatComplexEnt $item]}
    outputMsg " Adding ANCHOR IDs on: $ids" green
    if {[info exists noanchor]} {
      set ids {}
      foreach item [lrmdups $noanchor] {lappend ids [formatComplexEnt $item]}
      outputMsg "  ANCHOR IDs are also associated with: $ids" red
    }
  }
  foreach ent [array names urow] {
    $cells($ent) Item 3 $ucol($ent) "ANCHOR ID"
    set range [$worksheet($ent) Range [cellRange 3 $ucol($ent)] [cellRange $urow($ent) $ucol($ent)]]
    [$range Columns] AutoFit
    [$range Interior] ColorIndex [expr 40]
    for {set r 4} {$r <= $urow($ent)} {incr r} {
      set range [$worksheet($ent) Range [cellRange $r $ucol($ent)]]
      catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 1]}}
    }

    set range [$worksheet($ent) Range [cellRange 3 $ucol($ent)]]
    [[$range Borders] Item [expr 8]] Weight [expr 3]
    [$range Font] Bold [expr 1]
    $range HorizontalAlignment [expr -4108]
    [[[$worksheet($ent) Range [cellRange $urow($ent) $ucol($ent)]] Borders] Item [expr 9]] Weight [expr 3]
  }
}
