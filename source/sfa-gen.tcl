# generate an Excel spreadsheet from a STEP file

proc genExcel {{numFile 0}} {
  global localName localNameList programfiles env entCategories openFileList File buttons fileDir opt
  global worksheet worksheets wsCount sheetLast workbook workbooks worksheet1
  global cells row col heading thisEntType colColor count
  global propDefIDRow propDefID propDefRow propDefOK
  global gpmiIDRow gpmiID gpmiRow gpmiOK syntaxErr
  global excel excel1 entsIgnored nproc homedir filemenuinc lenlist startrow stepAP
  global writeDir entColorIndex entName errmsg lastXLS
  global ype nline entityCount userXLSFile comma colinv
  global cells1 col1 all_entity file_entity total_entity timeStamp xlFileNames xlFileName extXLS excelYear
  global multiFile spmiSumRow idxColor tlast developer
  global rowmax entCount userentlist userEntityFile flag
  global fixent fixprm lenfilelist badAttributes multiFileDir
  global ap203all ap209all ap210all ap214all ap238all ap242all
  global x3domFileOpen x3domFileName x3domCoord x3domIndex x3domFile x3domMin x3domMax x3domColor pmiCol spmiEntity
  global excelVersion recPracNames creo tolStandard dim coverageValues coverageLegend nistName
  global fcsv csvdirnam csvfile nistVersion
  
  if {[info exists errmsg]} {set errmsg ""}

# initialize for PMI X3DOM
  if {$opt(GENX3DOM)} {
    set x3domFileOpen 1
    set x3domFileName ""
    set x3domColor ""
    set x3domMax(x) -1.e10
    set x3domMax(y) -1.e10
    set x3domMax(z) -1.e10
    set x3domMin(x)  1.e10
    set x3domMin(y)  1.e10
    set x3domMin(z)  1.e10
  }

# check if IFCsvr is installed
  if {![file exists [file join $programfiles IFCsvrR300 dll IFCsvrR300.dll]]} {
    $buttons(genExcel) configure -state disable
    installIFCsvr
    return
  } 

# check for ROSE files
  if {![file exists [file join $programfiles IFCsvrR300 dll automotive_design.rose]]} {copyRoseFiles}

  set env(ROSE_RUNTIME) [file join $programfiles IFCsvrR300 dll]
  set env(ROSE_SCHEMAS) [file join $programfiles IFCsvrR300 dll]

  if {[info exists buttons]} {$buttons(genExcel) configure -state disable}
  catch {.tnb select .tnb.status}
  set lasttime [clock clicks -milliseconds]

  set multiFile 0
  if {$numFile > 0} {set multiFile 1}
  
# -------------------------------------------------------------------------------------------------
# connect to IFCsvr
  if {[catch {
    outputMsg "\nConnecting to IFCsvr" green
    set objIFCsvr [::tcom::ref createobject IFCsvr.R300]
    
# print errors
  } emsg]} {
    errorMsg "\nERROR connecting to IFCsvr: $emsg"
    catch {raise .}
    return 0
  }

# -------------------------------------------------------------------------------------------------
# open STEP file
  if {[catch {
    set nline 0
    outputMsg "Opening STEP file"
    set fname $localName  

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

# open file
    set objDesign [$objIFCsvr OpenDesign [file nativename $fname]]

# count entities
    set entityCount [$objDesign CountEntities "*"]
    outputMsg " $entityCount entities\n"
    if {$entityCount == 0} {errorMsg "There are no entities in the STEP file"}
    
# for STEP, get AP number, i.e. AP203   
    set stepAP [getStepAP $fname]

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
    
# open file of entities not to process (fixent)
    set cfile [file rootname $fname]
    append cfile "_fix.dat"
    set fixprm {}
    set fixent $fixprm
    if {[file exists $cfile]} {
      set fixfile [open $cfile r]
      while {[gets $fixfile line] >= 0} {
        if {[lsearch $fixent $line] == -1 && $line != "" && ![info exists badAttributes($line)]} {
          lappend fixent $line
        }
      }
      close $fixfile
    }

# check if a file generated from a NIST test case is being processed
    set nistName ""
    set ftail [string tolower [file tail $localName]]
    set ctcftc 0
    set filePrefix [list sp4_ sp5_ sp6_ sp7_ tgp1_ tgp2_ tgp3_ tgp4_ lsp_ lpp_ ltg_ ltp_]

    set ok  0
    set ok1 0
    foreach prefix $filePrefix {
      if {[string first $prefix $ftail] == 0 || [string first "nist" $ftail] != -1} {
        set tmp "nist_"
        foreach item {ctc ftc} {
          if {[string first $item $ftail] != -1} {
            append tmp "$item\_"
            set ctcftc 1
          }
        }

# find nist_ctc_01 directly        
        if {$ctcftc} {
          for {set i 1} {$i <= 11} {incr i} {
            set i1 $i
            if {$i < 10} {set i1 "0$i"}
            set tmp1 "$tmp$i1"
            if {[string first $tmp1 $ftail] != -1 && !$ok1} {
              set nistName $tmp1
              #outputMsg $nistName blue
              set ok1 1
            }
          }
        }

# find the number in the string            
        if {!$ok1} {
          for {set i 1} {$i <= 11} {incr i} {
            if {!$ok} {
              set i1 $i
              if {$i < 10} {set i1 "0$i"}
              set c {""}
              #outputMsg "$i1  [string first $i1 $ftail]  [string last $i1 $ftail]" blue
              if {[string first $i1 $ftail] != [string last $i1 $ftail]} {set c {"_" "-"}}
              foreach c1 $c {
                for {set j 0} {$j < 2} {incr j} {
                  if {$j == 0} {set i2 "$c1$i1"}
                  if {$j == 1} {set i2 "$i1$c1"}
                  #outputMsg "[string first $i2 $ftail]  $i2  $ftail" green
                  if {[string first $i2 $ftail] != -1 && !$ok} {
                    if {$ctcftc} {
                      append tmp $i1
                    } elseif {$i <= 5} {
                      append tmp "ctc_$i1"
                    } else {
                      append tmp "ftc_$i1"
                    }
                    set nistName $tmp
                    set ok 1
                    #outputMsg $nistName red
                  }
                }
              }
            }
          }
        }
      }
    }
    if {$developer && [string first "step-file-analyzer" $ftail] == 0} {set nistName "nist_ctc_01"}
    
# open tolerance coverage worksheet if PMI presentation and correct file name
    if {$opt(PMISEM) && !$coverageValues} {
      set ft [string tolower [file tail $fname]]
      if {([string first "sp" $ft] == 0 || [string first "lsp" $ft] == 0 || \
          [string first "nist" $ft] != -1 || [string first "ctc" $ft] != -1 || [string first "ftc" $ft] != -1 || \
          [string first "step-file-analyzer" $ft] != -1) && [string first "sp3" $ft] == -1} {
        spmiGetCoverageValues
        set coverageValues 1
      }
    }
    
# error opening file, report the schema
  } emsg]} {
    errorMsg "ERROR opening STEP file: $emsg"
    errorMsg "Possible causes of the error: (1) syntax errors in the STEP file, see Help > Conformance Checking\n (2) STEP schema is not supported or found, see Help > Other STEP APs,\n (3) wrong file extension, (4) file is not a STEP file." red
    getSchemaFromFile $fname 1
    if {!$nistVersion} {
      outputMsg " "
      errorMsg "You must process at least one STEP file with the NIST version of the STEP File Analyzer\n before using a user-built version."
    }
    outputMsg "\nClosing IFCsvr" green
    if {[info exists errmsg]} {unset errmsg}
    catch {
      $objDesign Delete
      unset objDesign
      unset objIFCsvr
    }
    catch {raise .}
    return 0
  }

# -------------------------------------------------------------------------------------------------
# connect to Excel
  if {$opt(XLSCSV) == "Excel"} {
    if {[catch {
      set pid1 [checkForExcel $multiFile]
      set excel [::tcom::ref createobject Excel.Application]
      set pidExcel [lindex [intersect3 $pid1 [twapi::get_process_ids -name "EXCEL.EXE"]] 2]
      [$excel ErrorCheckingOptions] TextDate False
      
      set excelVersion [expr {int([$excel Version])}]
      if {$excelVersion < 12} {
        set extXLS "xls"
        set rowmax [expr {2**16}]
        $excel DefaultSaveFormat [expr 56]
      } else {
        set extXLS "xlsx"
        set rowmax [expr {2**20}]
        $excel DefaultSaveFormat [expr 51]
      }
      set excelYear ""
      if {$excelVersion == 9} {
        set excelYear 2000
      } elseif {$excelVersion == 10} {
        set excelYear 2002
      } elseif {$excelVersion == 11} {
        set excelYear 2003
      } elseif {$excelVersion == 12} {
        set excelYear 2007
      } elseif {$excelVersion == 14} {
        set excelYear 2010
      } elseif {$excelVersion == 15} {
        set excelYear 2013
      } elseif {$excelVersion == 16} {
        set excelYear 2016
      }
      if {$excelVersion >= 2000 && $excelVersion < 2100} {set excelYear $excelVersion}
      outputMsg "Connecting to Excel $excelYear" green
  
      if {$excelVersion  < 12} {errorMsg " Some spreadsheet features are not available with this older version of Excel."}
  
# turning off ScreenUpdating saves A LOT of time
      if {$opt(XL_KEEPOPEN) && $numFile == 0} {
        $excel Visible 1
      } else {
        $excel Visible 0
        catch {$excel ScreenUpdating 0}
      }
      
      set rowmax [expr {$rowmax-2}]
      if {$opt(ROWLIM) < $rowmax} {set rowmax $opt(ROWLIM)}
      
# error with Excel, use CSV instead
    } emsg]} {
      errorMsg "ERROR connecting to Excel: $emsg"
      #if {[string first "Invalid class string" $emsg] != -1} {
        errorMsg "The STEP File will be written to CSV files.  See the option on the Spreadsheet tab."
        set opt(XLSCSV) "CSV"
        checkValues
        tk_messageBox -type ok -icon error -title "ERROR connecting to Excel" -message "Cannot connect to Excel or Excel is not installed.\nThe STEP file will be written to CSV files.\nSee the option on the Spreadsheet tab."
        catch {raise .}
        #return 0
      #}
    }

# -------------------------------------------------------------------------------------------------
# start worksheets
    if {$opt(XLSCSV) == "Excel"} {
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
  
# determine decimal separator
        set sheet [$worksheets Item [expr 1]]
        set cell  [$sheet Cells]
    
        set A1 12345,67890
        $cell Item 1 A $A1
        set range [$sheet Range "A1"]
        set comma 0
        if {[$range Value] == 12345.6789} {
          set comma 1
          errorMsg "Using comma \",\" as the decimal separator for numbers" red
        }
      
# print errors
      } emsg]} {
        errorMsg "ERROR opening Excel workbooks and worksheets: $emsg"
        catch {raise .}
        return 0
      }
    }
  }
  
# -------------------------------------------------------------------------------------------------
# add header worksheet, for CSV files create directory too and header file
  addHeaderWorksheet $objDesign $numFile $fname

# -------------------------------------------------------------------------------------------------
# set Excel spreadsheet name, delete file if already exists

# user-defined file name
  if {$opt(XLSCSV) == "Excel"} {
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
      set xlFileName "[file nativename [file join [file dirname $fname] [file rootname [file tail $fname]]]]_stp.$extXLS"
  
# user-defined directory
    } elseif {$opt(writeDirType) == 2} {
      set xlFileName "[file nativename [file join $writeDir [file rootname [file tail $fname]]]]_stp.$extXLS"
    }
    
# file name too long
    if {[string length $xlFileName] > 218} {
      if {[string length $xlsmsg] > 0} {append xlsmsg "\n\n"}
      append xlsmsg "Pathname of Spreadsheet file is too long for Excel ([string length $xlFileName])"
      set xlFileName "[file nativename [file join $writeDir [file rootname [file tail $fname]]]]_stp.$extXLS"
      if {[string length $xlFileName] < 219} {
        append xlsmsg "\nSpreadsheet file written to User-defined directory (Spreadsheet tab)"
      }
    }
  
# delete existing file
    if {[file exists $xlFileName]} {
      if {[catch {
        file delete -force $xlFileName
        #errorMsg "\nDeleting existing Spreadsheet: [truncFileName $xlFileName]" red
      } emsg]} {
        if {[string length $xlsmsg] > 0} {append xlsmsg "\n"}
        append xlsmsg "ERROR deleting existing Spreadsheet: [truncFileName $xlFileName]"
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
  global entCategory
  foreach pr [array names entCategory] {
    set ok 1
    if {[info exists opt($pr)] && $ok} {
      if {$opt($pr)} {set entCategories [concat $entCategories $entCategory($pr)]}
    }
  }
  
# -------------------------------------------------------------------------------------------------
# set which entities are processed and which are not
  set entsToProcess {}
  set entsToIgnore {}
  set numEnts 0
  
# user-defined entity list
  catch {set userentlist {}}
  if {$opt(PR_USER) && [llength $userentlist] == 0 && [info exists userEntityFile]} {
    set userentlist {}
    set fileUserEnt [open $userEntityFile r]
    while {[gets $fileUserEnt line] != -1} {
      set line [split [string trim $line] " "]
      foreach ent $line {
        if {[lsearch $ap203all $ent] != -1 || \
            [lsearch $ap209all $ent] != -1 || \
            [lsearch $ap210all $ent] != -1 || \
            [lsearch $ap214all $ent] != -1 || \
            [lsearch $ap238all $ent] != -1 || \
            [lsearch $ap242all $ent] != -1} {
          lappend userentlist $ent
        } elseif {[string first "_and_" $ent] != -1} {
          lappend userentlist $ent            
        }
      }
    }
    close $fileUserEnt
    if {[llength $userentlist] == 0} {
      set opt(PR_USER) 0
      checkValues
    }
  }
  
# get totals of each entity in file
  set fixlist {}
  if {![info exists objDesign]} {return}
  foreach entType [$objDesign EntityTypeNames [expr 2]] {
    set entCount($entType) [$objDesign CountEntities "$entType"]

    if {$entCount($entType) > 0} {
      if {$numFile != 0} {
        set idx [setColorIndex $entType 1]
        if {$idx == -2} {set idx 99}
        lappend all_entity "$idx$entType"
        lappend file_entity($numFile) "$entType $entCount($entType)"
        if {![info exists total_entity($entType)]} {
          set total_entity($entType) $entCount($entType)
        } else {
          incr total_entity($entType) $entCount($entType)
        }
      }

# user-defined entities
      set ok 0
      if {$opt(PR_USER) && [lsearch $userentlist $entType] != -1} {set ok 1}
      
# STEP entities that are translated depending on the options
      set ok1 [setEntsToProcess $entType $objDesign]
      if {$ok == 0} {set ok $ok1}
      
# entities in unsupported APs
      if {$stepAP == ""} {
        if {[lsearch $ap203all $entType] == -1 && \
            [lsearch $ap209all $entType] == -1 && \
            [lsearch $ap210all $entType] == -1 && \
            [lsearch $ap214all $entType] == -1 && \
            [lsearch $ap238all $entType] == -1 && \
            [lsearch $ap242all $entType] == -1} {
          set ok 1
        }
      }

# handle '_and_' due to a complex entity, entType_1 is the first part before the '_and_'
      set entType_1 $entType
      set c1 [string first "_and_" $entType_1]
      if {$c1 != -1} {set entType_1 [string range $entType_1 0 $c1-1]}
      
# check for entities that cause crashes
      set nofix 1
      if {[info exists fixent]} {if {[lsearch $fixent $entType] != -1} {set nofix 0}}
      if {$entType == "presentation_style_assignment" && $creo == 1} {set nofix 0}

# add to list of entities to process (entsToProcess), uses color index to set the order
      if {([lsearch $entCategories $entType_1] != -1 || $ok)} {
        if {$nofix} {
          lappend entsToProcess "[setColorIndex $entType]$entType"
          incr numEnts $entCount($entType)
        } else {
          if {$entType != "presentation_style_assignment" || $creo == 0} {lappend fixlist $entType}
          lappend entsToIgnore $entType
          set entsIgnored($entType) $entCount($entType)
        }
      } elseif {[lsearch $entCategories $entType] != -1} {
        if {$nofix} {
          lappend entsToProcess "[setColorIndex $entType]$entType"
          incr numEnts $entCount($entType)
        } else {
          if {$entType != "presentation_style_assignment" || $creo == 0} {lappend fixlist $entType}
          lappend entsToIgnore $entType
          set entsIgnored($entType) $entCount($entType)
        }
      } else {
        lappend entsToIgnore $entType
        set entsIgnored($entType) $entCount($entType)
      }
    }
  }
  
# list entities not processed based on fix file
  if {[llength $fixlist] > 0} {
    outputMsg " "
    if {[file exists $cfile]} {
      set ok 0
      foreach item $fixlist {if {[lsearch $fixprm $item] == -1} {set ok 1}}
      if {$ok} {errorMsg "Based on entities listed in [truncFileName [file nativename $cfile]]"}
    }
    errorMsg " Worksheets will not be generated for the following entities:"
    foreach item [lsort $fixlist] {outputMsg "  $item" red}
  }
  
# sort entsToProcess by color index
  set entsToProcess [lsort $entsToProcess]
  
# for STEP process datum* and dimensional* entities before specific *_tolerance entities
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
  
# then strip off the color index
  for {set i 0} {$i < [llength $entsToProcess]} {incr i} {
    lset entsToProcess $i [string range [lindex $entsToProcess $i] 2 end]
  }

  if {[info exists buttons]} {$buttons(pgb) configure -maximum $numEnts}
      
# check for ISO/ASME standards on product_definition_formation, document, product
  set tolStandard(type) ""
  set tolStandard(num)  ""
  set stds {}
  foreach item {product_definition_formation product} {
    ::tcom::foreach thisEnt [$objDesign FindObjects $item] {
      ::tcom::foreach attr [$thisEnt Attributes] {
        if {[$attr Name] == "id"} {
          set val [$attr Value]
          if {([string first "ISO" $val] != -1 || [string first "ASME" $val] != -1) && [string first "NIST" [string toupper $val]] == -1} {
            if {[string first "ISO"  $val] != -1} {
              set tolStandard(type) "ISO"
              if {[string first "1101" $val] != "" || [string first "16792" $val] != ""} {if {[string first $val $tolStandard(num)] == -1} {append tolStandard(num) "$val  "}}
            }
            if {[string first "ASME" $val] != -1 && [string first "NIST" [string toupper $val]] == -1} {
              set tolStandard(type) "ASME"
              if {[string first "Y14." $val] != ""} {if {[string first $val $tolStandard(num)] == -1} {append tolStandard(num) "$val  "}}
            }
            set ok 1
            foreach std $stds {if {[string first $val $std] != -1} {set ok 0}}
            if {$ok} {lappend stds $val}
          }
        }
      }
    }
  }
  if {[llength $stds] > 0} {
    outputMsg "\nStandards:"
    foreach std $stds {outputMsg " $std" blue}
  }
  if {$tolStandard(type) == "ISO"} {
    set fn [string toupper [file tail $localName]]
    if {[string first "NIST_" $fn] == 0 && [string first "ASME" $fn] != -1} {errorMsg "All of the NIST models use ASME Y14.5 tolerance standard."}
  }

# -------------------------------------------------------------------------------------------------
# generate worksheet for each entity
  if {$opt(XLSCSV) == "Excel"} {
    outputMsg "\nGenerating STEP Entity worksheets" green
  } else {
    outputMsg "\nGenerating STEP Entity CSV files" green
  }
  if {[catch {
    set inverseEnts {}
    set lastEnt ""
    set nline 0
    set nproc 0
    set wsCount 0
    set stat 1
    set spmiEntity {}
    set ntable 0
    set spmiSumRow 1
    set idxColor 0
    set coverageLegend 0

    if {[info exists dim]} {unset dim}
    set dim(prec,max) 0
    set dim(unit) ""
    #set dim(name) ""

# find camera models used in draughting model items andannotation_occurrence used in property_definition and datums
    if {$opt(PMIGRF)} {pmiGetCamerasAndProperties $objDesign}

    if {[llength $entsToProcess] == 0} {
      errorMsg "No STEP entities were found to Process as selected in the Options tab."
      break
    }
    set tlast [clock clicks -milliseconds]
    #getTiming "start entity processing"
    
# loop over list of entities in file
    foreach entType $entsToProcess {
      set nerr1 0
      set lastEnt $entType
      
# decide if inverses should be checked for this entity type
      set checkInv 0
      if {$opt(INVERSE)} {set checkInv [invSetCheck $entType]}
      if {$checkInv} {lappend inverseEnts $entType}
      set badAttr [info exists badAttributes($entType)]

# process the entity
      ::tcom::foreach objEntity [$objDesign FindObjects [join $entType]] {
        if {$entType == [$objEntity Type]} {
          incr nline
          if {[expr {$nline%1000}] == 0} {update idletasks}

          if {[catch {
            if {$opt(XLSCSV) == "Excel"} {
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
                if {[$objEntity Type] != "trimmed_curve"} {
                  append msg "\#[$objEntity P21ID]=[$objEntity Type] (row [expr {$row($thisEntType)+2}]): $emsg1"

# handle specific errors
                  if {[string first "Unknown error" $emsg1] != -1} {
                    errorMsg $msg
                    catch {raise .}
                    incr nerr1
                    if {$nerr1 > 20} {
                      errorMsg "Processing of $entType entities has stopped" red
                      set nline [expr {$nline + $entCount($thisEntType) - $count($thisEntType)}]
                      break
                    }

                  } elseif {[string first "Insufficient memory to perform operation" $emsg1] != -1} {
                    errorMsg $msg
                    errorMsg "Several options are available to reduce memory usage:\nUse the option to limit the Maximum Rows"
                    if {$opt(INVERSE)} {errorMsg "Turn off Inverse Relationships and process the file again" red}
                    catch {raise .}
                    break
                  }

# error message for trimmed_curve (causes problems for IFCsvr)
                } else {
                  append msg [$objEntity Type]
                }
                errorMsg $msg 
                catch {raise .}
              }
            }
          }
          if {$stat != 1} {
            set nline [expr {$nline + $entCount($thisEntType) - $count($thisEntType)}]
            break
          }
        }
      }

      if {$opt(XLSCSV) == "CSV"} {catch {close $fcsv}}
      
# check for validation properties, PMI presentation and representation
      checkForPMIandValProps $objDesign $entType
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
      set fixfile [open $cfile r]
      while {[gets $fixfile line] >= 0} {
        if {[lsearch $fixtmp $line] == -1 && $line != $lastEnt} {lappend fixtmp $line}
      }
      close $fixfile
    }

    if {[join $fixtmp] == ""} {
      catch {file delete -force $cfile}
    } else {
      set fixfile [open $cfile w]
      foreach item $fixtmp {puts $fixfile $item}
      close $fixfile
    }
  }

# -------------------------------------------------------------------------------------------------
# set viewpoints and close PMI X3DOM file 
  if {$opt(GENX3DOM) && $x3domFileName != ""} {gpmiX3DOMViewpoints}

# -------------------------------------------------------------------------------------------------
# quit IFCsvr, but not sure how to do it properly
  if {[catch {
    outputMsg "\nClosing IFCsvr" green
    $objDesign Delete
    unset objDesign
    unset objIFCsvr
    
# print errors
  } emsg]} {
    errorMsg "ERROR closing IFCsvr: $emsg"
    catch {raise .}
  }

# -------------------------------------------------------------------------------------------------
# add summary worksheet
  if {$opt(XLSCSV) == "Excel"} {
    set tmp [sumAddWorksheet] 
    set sumLinks  [lindex $tmp 0]
    set sheetSort [lindex $tmp 1]
    set sumRow    [lindex $tmp 2]
    set sum "Summary"
  
# add file name and other info to top of Summary
    set sumHeaderRow [sumAddFileName $sum $sumLinks]

# freeze panes (must be before adding color and hyperlinks below)
    [$worksheet($sum) Range "B[expr {$sumHeaderRow+3}]"] Select
    [$excel ActiveWindow] FreezePanes [expr 1]
    [$worksheet($sum) Range "A1"] Select
  
# add Summary color and hyperlinks
    sumAddColorLinks $sum $sumHeaderRow $sumLinks $sheetSort $sumRow
    #getTiming "done generating summary worksheet"

# -------------------------------------------------------------------------------------------------
# format cells on each entity worksheets
    formatWorksheets $sheetSort $sumRow $inverseEnts
    #getTiming "done formatting spreadsheets"
  
# -------------------------------------------------------------------------------------------------
# add PMI Rep. Coverage Analysis worksheet for a single file
    if {$opt(PMISEM)} {
      global spmiTypesPerFile
      if {[info exists spmiTypesPerFile]} {
        set sempmi_coverage "PMI Representation Coverage"
        if {![info exists worksheet($sempmi_coverage)]} {
          outputMsg " Adding PMI Representation Coverage worksheet" green
          spmiCoverageStart 0
          spmiCoverageWrite "" "" 0
          spmiCoverageFormat "" 0
        }
  
# add line to bottow of PMI rep summary worksheet
        global spmiSumName  
        if {$spmiSumRow > 1} {
          [$worksheet($spmiSumName) Columns] AutoFit
          [$worksheet($spmiSumName) Rows] AutoFit
          incr spmiSumRow 2
          $cells($spmiSumName) Item 1 3 "See CAx-IF Recommended Practice for $recPracNames(pmi242)"
          set range [$worksheet($spmiSumName) Range C1:K1]
          $range MergeCells [expr 1]
          set anchor [$worksheet($spmiSumName) Range C1]
          [$worksheet($spmiSumName) Hyperlinks] Add $anchor [join "https://www.cax-if.org/joint_testing_info.html#recpracs"] [join ""] [join "Link to CAx-IF Recommended Practices"]
        }
      }
    }
  
# add PMI Pres. Coverage Analysis worksheet for a single file
    if {$opt(PMIGRF)} {
      global gpmiTypesPerFile
      if {[info exists gpmiTypesPerFile]} {
        set pmi_coverage "PMI Presentation Coverage"
        if {![info exists worksheet($pmi_coverage)]} {
          outputMsg " Adding PMI Presentation Coverage worksheet" green
          gpmiCoverageStart 0
          gpmiCoverageWrite "" "" 0
          gpmiCoverageFormat "" 0
        }
      }
    }
    
# select the first tab
    [$worksheets Item [expr 1]] Select
    [$excel ActiveWindow] ScrollRow [expr 1]
  }

  set cc [clock clicks -milliseconds]
  set proctime [expr {($cc - $lasttime)/1000}]
  if {$proctime <= 60} {set proctime [expr {(($cc - $lasttime)/100)/10.}]}
  outputMsg "Processing time: $proctime seconds" green

# -------------------------------------------------------------------------------------------------
# save spreadsheet

  if {$opt(XLSCSV) == "Excel"} {
    if {[catch {
      #getTiming "save spreadsheet"
      outputMsg " "
      if {$xlsmsg != ""} {errorMsg $xlsmsg}
      if {[string first "\[" $xlFileName] != -1} {
        regsub -all {\[} $xlFileName "(" xlFileName
        regsub -all {\]} $xlFileName ")" xlFileName
        errorMsg "In the spreadsheet file name, the characters \'\[\' and \'\]\' have been\n substituted by \'\(\' and \'\)\'"
      }
      outputMsg "Saving Spreadsheet as:\n [truncFileName $xlFileName 1]" blue
      $workbook SaveAs $xlFileName
      set lastXLS $xlFileName
      lappend xlFileNames $xlFileName
  
      catch {$excel ScreenUpdating 1}

# close Excel
      outputMsg "Closing Excel" green
      $excel Quit
      if {[info exists excel]} {unset excel}
      set openxl 1
      if {[llength $pidExcel] == 1} {
        catch {twapi::end_process $pidExcel -force}
      } else {
        errorMsg " Excel might not have been closed" red
      }
      update idletasks
      #getTiming "save done"

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

# errors
    } emsg]} {
      errorMsg "ERROR saving Spreadsheet: $emsg"
      if {[string first "The file or path name not found" $emsg]} {
        outputMsg " "
        errorMsg "Either copy the STEP file to a different directory and try generating the\n spreadsheet again or use the option to write the spreadsheet to a user-defined\n directory (Spreadsheet tab)."
      }
      catch {raise .}
      set openxl 0
    }
    
# -------------------------------------------------------------------------------------------------
# open spreadsheet
    set ok 0
    if {$openxl && $opt(XL_OPEN)} {
      if {$numFile == 0} {
        set ok 1
      } elseif {[info exists lenfilelist]} {
        if {$lenfilelist == 1} {set ok 1}
      }
    }
    if {$ok} {openXLS $xlFileName}

# open directory of CSV files
  } else {
    unset csvfile
    outputMsg "\nCSV files written to: [file nativename $csvdirnam]\n" blue
    set ok 0
    if {$opt(XL_OPEN)} {
      if {$numFile == 0} {
        set ok 1
      } elseif {[info exists lenfilelist]} {
        if {$lenfilelist == 1} {set ok 1}
      }
    }
    if {$ok} {exec {*}[auto_execok start] [file nativename $csvdirnam]}
  }

# -------------------------------------------------------------------------------------------------
# display X3DOM file of graphical PMI
  displayX3DOM

# -------------------------------------------------------------------------------------------------
# save state
  if {[info exists errmsg]} {unset errmsg}
  saveState
  if {!$multiFile && [info exists buttons]} {$buttons(genExcel) configure -state normal}
  update idletasks

# clean up variables to hopefully release some memory and/or to reset them
  global currX3domPointID numX3domPointID pmiStartCol nrep invGroup dimrep dimrepID pmiColumns
  foreach var {colColor count entsIgnored colinv \
               worksheet worksheets workbook workbooks cells \
               heading entName \
               propDefIDRow propDefID propDefRow propDefOK \
               gpmiIDRow gpmiID gpmiRow gpmiOK \
               currX3domPointID numX3domPointID pmiCol pmivalprop pmiStartCol \
               x3domFileOpen x3domFileName x3domCoord x3domIndex x3domFile x3domMin x3domMax \
               syntaxErr nrep invGroup dimrep dimrepID pmiColumns} {
    if {[info exists $var]} {unset $var}
  }
  if {!$multiFile} {
    global gpmiTypesPerFile spmiTypesPerFile
    foreach var {gpmiTypesPerFile spmiTypesPerFile} {
      if {[info exists $var]} {unset $var}
    }
  }
  update idletasks
  return 1
}

#-------------------------------------------------------------------------------------------------
# add summary worksheet
proc sumAddWorksheet {} {
  global worksheet cells sum sheetSort sheetLast col worksheets row entCategory opt entsIgnored excel
  global x3domFileName spmiEntity entCount gpmiEnts spmiEnts nistVersion

  outputMsg "Adding Summary worksheet"
  set sum "Summary"
  #getTiming "done processing entities"

  set sheetSort {}
  foreach entType [lsort [array names worksheet]] {
    if {$entType != "Summary" && $entType != "Header"} {
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
    set x3domLink 1
    set row($sum) 1
    foreach entType $sheetSort {
      incr row($sum)
      set sumRow [expr {[lsearch $sheetSort $entType]+2}]

# check if entity is compound as opposed to an entity with '_and_'
      set ok 0
      if {[string first "_and_" $entType] == -1} {
        set ok 1
      } else {
        foreach item [array names entCategory] {if {[lsearch $entCategory($item) $entType] != -1} {set ok 1}}
      }
      if {$ok} {
        $cells($sum) Item $sumRow 1 $entType
        
# for STEP add [Validation Properties], [PMI Presentation], [PMI Representation] text string and link to X3DOM file    
        set okao 0
        if {$entType == "property_definition" && $col($entType) > 4} {
          $cells($sum) Item $sumRow 1 "property_definition  \[Validation Properties\]"
        } elseif {$entType == "dimensional_characteristic_representation" && $col($entType) > 3} {
          $cells($sum) Item $sumRow 1 "dimensional_characteristic_representation  \[PMI Representation\]"
        } elseif {[lsearch $spmiEntity $entType] != -1} {
          $cells($sum) Item $sumRow 1 "$entType  \[PMI Representation\]"
        } elseif {[string first "annotation" $entType] != -1} {
          if {$gpmiEnts($entType) && $col($entType) > 5} {set okao 1}
        }
        if {$okao} {
          $cells($sum) Item $sumRow 1 "$entType  \[PMI Presentation\]"
          if {$opt(GENX3DOM) && $x3domFileName != "" && $x3domLink} {
            $cells($sum) Item $sumRow 3 "Graphics"
            set x3domLink 0
            set anchor [$worksheet($sum) Range [cellRange $sumRow 3]]
            $sumLinks Add $anchor [join $x3domFileName] [join ""] [join "Link to Graphics file"]
          }
        }

# for '_and_' (complex entity) split on multiple lines
# '10' is the ascii character for a linefeed          
      } else {
        regsub -all "_and_" $entType ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" entType_multiline
        set entType_multiline "($entType_multiline)"
        $cells($sum) Item $sumRow 1 $entType_multiline

# for STEP add [Validation Properties] or [PMI Presentation] text string and link to X3DOM file    
        set okao 0
        if {[string first "annotation" $entType] != -1} {
          if {$gpmiEnts($entType) && $col($entType) > 7} {set okao 1}
        } elseif {[lsearch $spmiEntity $entType] != -1} {
          $cells($sum) Item $sumRow 1 "$entType_multiline  \[PMI Representation\]"
        }
        if {$okao} {
          $cells($sum) Item $sumRow 1 "$entType_multiline  \[PMI Presentation\]"
          if {$opt(GENX3DOM) && $x3domFileName != "" && $x3domLink} {
            $cells($sum) Item $sumRow 3 "Graphics"
            set x3domLink 0
            set anchor [$worksheet($sum) Range [cellRange $sumRow 3]]
            $sumLinks Add $anchor [join $x3domFileName] [join ""] [join "Link to Graphics file"]
          }
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
      set ok 0
      if {[string first "_and_" $ent] == -1} {
        set ok 1
      } else {
        foreach item [array names entCategory] {if {[lsearch $entCategory($item) $ent] != -1} {set ok 1}}
      }
      if {$ok} {
        $cells($sum) Item [incr rowIgnored] 1 $ent
      } else {
# '10' is the ascii character for a linefeed          
        regsub -all "_and_" $ent ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" ent1
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
    set str ""
    if {$nistVersion} {set str "NIST "}
    $cells($sum) Item [expr {$row($sum)+2}] 1 "Spreadsheet generated by the $str\STEP File Analyzer (v[getVersion])"
    set anchor [$worksheet($sum) Range [cellRange [expr {$row($sum)+2}] 1]]
    [$worksheet($sum) Hyperlinks] Add $anchor [join "http://www.nist.gov/el/msid/infotest/step-file-analyzer.cfm"] [join ""] [join "Link to STEP File Analyzer"]
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
  global worksheet cells timeStamp cadSystem xlFileName localName opt entityCount stepAP schemaLinks
  global tolStandard dim fileSchema1

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
      $cells($sum) Item 1 1 "Tolerance Standard"
      if {$tolStandard(num) != ""} {
        $cells($sum) Item 1 2 "$tolStandard(num)"
      } else {
        $cells($sum) Item 1 2 "$tolStandard(type)"
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
      $cells($sum) Item 1 2 "'$fileSchema1"
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
    if {$opt(XL_LINK1)} {$sumLinks Add $anchor [join $localName] [join ""] [join "Link to STEP file"]}
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
  global worksheet cells row excel entName xlFileName col entsIgnored

  if {[catch {
    outputMsg "Adding links on Summary to STEP documentation"
    set row($sum) [expr {$sumHeaderRow+2}]
    set nline 0

    foreach ent $sheetSort {
      incr nline
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
      $sumLinks Add $anchor $xlFileName "$hlsheet!A4" "Go to $ent"

# color entities on summary
      set cidx [setColorIndex $ent]
      if {$cidx > 0} {
        [$anchor Interior] ColorIndex [expr $cidx]
        catch {
          [[$anchor Borders] Item [expr 8]] Weight [expr 1]
          [[$anchor Borders] Item [expr 9]] Weight [expr 1]
        }
      }

# bold entities for reports
      if {[string first "\[" [$anchor Value]] != -1} {[$anchor Font] Bold [expr 1]}

      set ncol [expr {$col($sum)-1}]
    }

# add links for entsIgnored entities, find row where they start
    set i1 [expr {max([array size worksheet],9)}]
    for {set i $i1} {$i < 1000} {incr i} {
      if {[string first "Entity" [[$cells($sum) Item $i 1] Value]] == 0} {
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
      #entDocLink $sum $ent $rowIgnored $sumDocCol $sumLinks

      set range [$worksheet($sum) Range [cellRange $rowIgnored 1]]
      set cidx [setColorIndex $ent 1]
      if {$cidx > 0} {[$range Interior] ColorIndex [expr $cidx]}      
    }
    [$worksheet($sum) Columns] AutoFit
    [$worksheet($sum) Rows] AutoFit
    [$worksheet($sum) PageSetup] PrintGridlines [expr 1]
    
  } emsg]} {
    errorMsg "ERROR adding Summary colors and links: $emsg"
    catch {raise .}
  }
}

#-------------------------------------------------------------------------------------------------
# format worksheets
proc formatWorksheets {sheetSort sumRow inverseEnts} {
  global buttons worksheet excel cells opt count entCount col row rowmax xlFileName thisEntType schemaLinks stepAP syntaxErr
  global gpmiEnts spmiEnts
  
  outputMsg "Formatting Worksheets"

  if {[info exists buttons]} {$buttons(pgb) configure -maximum [llength $sheetSort]}
  set nline 0
  set nsort 0
  foreach thisEntType $sheetSort {
    #getTiming "START FORMATTING $thisEntType"
    incr nline
    update idletasks
    
    if {[catch {
      $worksheet($thisEntType) Activate
      [$excel ActiveWindow] ScrollRow [expr 1]

# find extent of columns
      set rancol $col($thisEntType)
      for {set i 1} {$i < 10} {incr i} {
        if {[[$cells($thisEntType) Item 3 [expr {$col($thisEntType)+$i}]] Value] != ""} {
          incr rancol
        } else {
          break
        }
      }
      #getTiming " column extent"

# find extent of rows
      set ranrow [expr {$row($thisEntType)+2}]
      if {$ranrow > $rowmax} {set ranrow [expr {$rowmax+2}]}
      set ranrow [expr {$ranrow-2}]
      #getTiming " row extent"
      #outputMsg "$thisEntType  $ranrow  $rancol  $col($thisEntType)"

# autoformat
      set range [$worksheet($thisEntType) Range [cellRange 3 1] [cellRange $ranrow $rancol]]
      $range AutoFormat
      #getTiming " autoformat"

# freeze panes
      [$worksheet($thisEntType) Range "B4"] Select
      [$excel ActiveWindow] FreezePanes [expr 1]
      
# set A4 as default cell
      [$worksheet($thisEntType) Range "A4"] Select

# set column color, border, group for INVERSES and Used In
      if {$opt(INVERSE)} {if {[lsearch $inverseEnts $thisEntType] != -1} {invFormat $rancol}}
      #getTiming " format inverses"

# STEP Property_definition (Validation Properties)
      if {$thisEntType == "property_definition" && $opt(VALPROP)} {
        valPropFormat
        #getTiming " format valprop"

# color STEP annotation occurrence (Graphical PMI)
      } elseif {$gpmiEnts($thisEntType) && $opt(PMIGRF)} {
        pmiFormatColumns "PMI Presentation"
        #getTiming " format gpmi"

# color STEP semantic PMI
      } elseif {$spmiEnts($thisEntType) && $opt(PMISEM)} {
        pmiFormatColumns "PMI Representation"

# add PMI Represenation Summary worksheet
        spmiSummary
        #getTiming " format spmi"
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
      set range [$worksheet($thisEntType) Range "A1:H1"]
      $range MergeCells [expr 1]

# link back to summary
      set anchor [$worksheet($thisEntType) Range "A1"]
      $hlink Add $anchor $xlFileName "Summary!A$sumRow" "Return to Summary"
      #getTiming " insert links in first two rows"

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
      #getTiming " check column width"
      
# color "Bad" (red) for syntax errors
      if {[info exists syntaxErr($thisEntType)]} {colorBadCells $thisEntType}
      #getTiming " color bad syntax"

# landscape page orientation, print gridlines
      #if {$opt(XL_ORIENT)} {
      #  [$worksheet($thisEntType) PageSetup] Orientation [expr 2]
      #  [$worksheet($thisEntType) PageSetup] PrintGridlines [expr 1]
      #}
  
# -------------------------------------------------------------------------------------------------
# add table for sorting and filtering
      if {[expr {int([$excel Version])}] >= 12} {
        if {[catch {
          if {$opt(SORT) && $thisEntType != "property_definition"} {
            if {$ranrow > 8} {
              set range [$worksheet($thisEntType) Range [cellRange 3 1] [cellRange $ranrow $rancol]]
              set tname [string trim "TABLE-$thisEntType"]
              [[$worksheet($thisEntType) ListObjects] Add 1 $range] Name $tname
              [[$worksheet($thisEntType) ListObjects] Item $tname] TableStyle "TableStyleLight1" 
              if {[incr ntable] == 1 && $opt(SORT)} {outputMsg " Generating Tables for Sorting" blue}
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
