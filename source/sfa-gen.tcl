# generate an Excel spreadsheet from a STEP file

proc genExcel {{numFile 0}} {
  global allEntity ap203all ap214all ap242all badAttributes buttons
  global cells cells1 col col1 comma count coverageLegend readPMI noPSA csvdirnam csvfile
  global developer dim entCategories entCategory entColorIndex entCount entityCount entsIgnored env errmsg
  global excel excelVersion excelYear extXLS fcsv feaElemTypes File fileEntity skipEntities skipPerm gpmiTypesPerFile idxColor inverses
  global lastXLS lenfilelist localName localNameList multiFile multiFileDir nistName nistVersion nprogEnts
  global opt p21e3 pmiCol pmiMaster programfiles recPracNames row rowmax sheetLast spmiEntity spmiSumName spmiSumRow spmiTypesPerFile startrow stepAP
  global thisEntType timeStamp tlast tolStandard totalEntity userEntityFile userEntityList userXLSFile
  global workbook workbooks worksheet worksheet1 worksheets writeDir wsCount
  global x3domColor x3domCoord x3domFile x3domFileName x3domFileOpen x3domIndex x3domMax x3domMin
  global xlFileName xlFileNames
  
  if {[info exists errmsg]} {set errmsg ""}

# initialize for PMI OR AP209 X3DOM
  if {$opt(VIZPMI) || $opt(VIZFEA)} {
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
    #outputMsg "\nConnecting to IFCsvr" green
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
    set nprogEnts 0
    outputMsg "\nOpening STEP file"
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
        file delete -force $cfile1
        errorMsg "File of entities to skip '[file tail $cfile1]' renamed to '[file tail $cfile]'."
      }
    }

# check if a file generated from a NIST test case is being processed
    set nistName ""
    set ftail [string tolower [file tail $localName]]
    set ctcftc 0
    set filePrefix [list sp4_ sp5_ sp6_ sp7_ tgp1_ tgp2_ tgp3_ tgp4_ tp3_ tp4_ tp5_ tp6_ lsp_ lpp_ ltg_ ltp_]

    set ok  0
    set ok1 0
    foreach prefix $filePrefix {
      if {[string first $prefix $ftail] == 0 || [string first "nist" $ftail] != -1 || \
          [string first "ctc" $ftail] != -1 || [string first "ftc" $ftail] != -1} {
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
    
# other files
    if {!$ok && [string first "sp3" $ftail] == 0} {
      if {[string first "1101"  $ftail] != -1} {set nistName "sp3-1101"}
      if {[string first "16792" $ftail] != -1} {set nistName "sp3-16792"}
      if {[string first "box"   $ftail] != -1} {set nistName "sp3-box"}
    }
    if {$developer && [string first "step-file-analyzer" $ftail] == 0} {set nistName "nist_ctc_01"}
    
# open expected PMI worksheet (once) if PMI presentation and correct file name
    if {$opt(PMISEM) && $stepAP == "AP242" && $nistName != ""} {
      if {![info exists pmiMaster($nistName)]} {spmiGetPMI}
    }
    
# error opening file, report the schema
  } emsg]} {
    errorMsg "ERROR opening STEP file"
    getSchemaFromFile $fname 1

    if {!$p21e3} {
      errorMsg "Possible causes of the ERROR:\n- Syntax errors in the STEP file\n- STEP schema is not supported, see Help > Supported STEP APs\n- Multiple schemas are used\n- Wrong file extension, should be '.stp'\n- STEP file contains new features from ISO 10303 Part 21 edition 3\n- File is not an ISO 10303 Part 21 STEP file" red
    
# part 21 edition 3
    } else {
      outputMsg " "
      errorMsg "The STEP file uses the new Edition 3 of Part 21 and cannot be processed by the STEP File Analyzer.\n Edit the STEP file to delete the Edition 3 content such as the ANCHOR, REFERENCE, and SIGNATURE sections."
    }
    if {!$nistVersion} {
      outputMsg " "
      errorMsg "You must process at least one STEP file with the NIST version of the STEP File Analyzer\n before using a user-built version."
    }
    #outputMsg "\nClosing IFCsvr" green
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
      set extXLS "xlsx"
      set rowmax [expr {2**20}]
      $excel DefaultSaveFormat [expr 51]
      if {$excelVersion < 12} {
        set extXLS "xls"
        set rowmax [expr {2**16}]
        $excel DefaultSaveFormat [expr 56]
      }
      set excelYear ""
      switch $excelVersion {
        9  {set excelYear 2000}
        10 {set excelYear 2002}
        11 {set excelYear 2003}
        12 {set excelYear 2007}
        14 {set excelYear 2010}
        15 {set excelYear 2013}
        16 {set excelYear 2016}
      }
      if {$excelVersion >= 2000 && $excelVersion < 2100} {set excelYear $excelVersion}
      #outputMsg "Connecting to Excel $excelYear" green
  
      if {$excelVersion < 12} {errorMsg " Some spreadsheet features are not available with this older version of Excel."}
  
# turning off ScreenUpdating saves A LOT of time
      if {$opt(XL_KEEPOPEN) && $numFile == 0} {
        $excel Visible 1
      } else {
        $excel Visible 0
        catch {$excel ScreenUpdating 0}
      }
      
      set rowmax [expr {$rowmax-2}]
      if {$opt(XL_ROWLIM) < $rowmax} {set rowmax $opt(XL_ROWLIM)}
      
# error with Excel, use CSV instead
    } emsg]} {
      errorMsg "ERROR connecting to Excel: $emsg"
      errorMsg "The STEP File will be written to CSV files.  See the setting on the Options tab."
      set opt(XLSCSV) "CSV"
      checkValues
      tk_messageBox -type ok -icon error -title "ERROR connecting to Excel" -message "Cannot connect to Excel or Excel is not installed.\nThe STEP file will be written to CSV files.\nSee the setting on the Options tab."
      catch {raise .}
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
# add header worksheet, for CSV files create directory and header file
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
      foreach ent $line {lappend userEntityList $ent}
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

  set entityTypeNames [$objDesign EntityTypeNames [expr 2]]
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
      set ok1 [setEntsToProcess $entType $objDesign]
      if {$ok == 0} {set ok $ok1}
      
# entities in unsupported APs that are not AP203, AP214, AP242
      if {[string first "AP203" $stepAP] == -1 && $stepAP != "AP214" && $stepAP != "AP242"} {
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
      }
      
# new AP242 entities in a ROSE file, but not yet in ap242all or any entity category, for testing new schemas      
      #if {$stepAP == "AP242" && [lsearch $ap242all $entType] == -1} {set ok 1}

# handle '_and_' due to a complex entity, entType_1 is the first part before the '_and_'
      set entType_1 $entType
      set c1 [string first "_and_" $entType_1]
      if {$c1 != -1} {set entType_1 [string range $entType_1 0 $c1-1]}
      
# check for entities that cause crashes
      set noSkip 1
      if {[info exists skipEntities]} {if {[lsearch $skipEntities $entType] != -1} {set noSkip 0}}
      if {$entType == "presentation_style_assignment" && $noPSA == 1} {set noSkip 0}

# add to list of entities to process (entsToProcess), uses color index to set the order
      if {([lsearch $entCategories $entType_1] != -1 || $ok)} {
        if {$noSkip} {
          lappend entsToProcess "[setColorIndex $entType]$entType"
          incr numEnts $entCount($entType)
        } else {
          if {$entType != "presentation_style_assignment" || $noPSA == 0} {lappend fixlist $entType}
          lappend entsToIgnore $entType
          set entsIgnored($entType) $entCount($entType)
        }
      } elseif {[lsearch $entCategories $entType] != -1} {
        if {$noSkip} {
          lappend entsToProcess "[setColorIndex $entType]$entType"
          incr numEnts $entCount($entType)
        } else {
          if {$entType != "presentation_style_assignment" || $noPSA == 0} {lappend fixlist $entType}
          lappend entsToIgnore $entType
          set entsIgnored($entType) $entCount($entType)
        }
      } else {
        lappend entsToIgnore $entType
        set entsIgnored($entType) $entCount($entType)
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
  
# list entities not processed based on fix file
  if {[llength $fixlist] > 0} {
    outputMsg " "
    if {[file exists $cfile]} {
      set ok 0
      foreach item $fixlist {if {[lsearch $skipPerm $item] == -1} {set ok 1}}
    }
    if {$ok} {
      errorMsg "Worksheets will NOT be generated for entities listed in\n [truncFileName [file nativename $cfile]]:"
      foreach item [lsort $fixlist] {outputMsg "  $item" red}
      errorMsg " See Help > Crash Recovery"
    }
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

# max progress bar - number of entities  
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
            if {[string first "ISO" $val] != -1} {
              set tolStandard(type) "ISO"
              if {[string first "1101" $val] != "" || [string first "16792" $val] != ""} {if {[string first $val $tolStandard(num)] == -1} {append tolStandard(num) "$val    "}}
            }
            if {[string first "ASME" $val] != -1 && [string first "NIST" [string toupper $val]] == -1} {
              set tolStandard(type) "ASME"
              if {[string first "Y14." $val] != ""} {if {[string first $val $tolStandard(num)] == -1} {append tolStandard(num) "$val    "}}
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
    outputMsg "\nGenerating STEP Entity worksheets" blue
  } else {
    outputMsg "\nGenerating STEP Entity CSV files" blue
  }
  if {[catch {

# initialize variables
    set inverseEnts {}
    set lastEnt ""
    set nprogEnts 0
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
          incr nprogEnts
          if {[expr {$nprogEnts%1000}] == 0} {update idletasks}

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
                append msg "\#[$objEntity P21ID]=[$objEntity Type] (row [expr {$row($thisEntType)+2}]): $emsg1"

# handle specific errors
                if {[string first "Unknown error" $emsg1] != -1} {
                  errorMsg $msg
                  catch {raise .}
                  incr nerr1
                  if {$nerr1 > 20} {
                    errorMsg "Processing of $entType entities has stopped" red
                    set nprogEnts [expr {$nprogEnts + $entCount($thisEntType) - $count($thisEntType)}]
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
            set n $nprogEnts
            set ok 1
            if {[string first "element_representation" $thisEntType] != -1 && $opt(VIZFEA)} {set ok 0}
            if {$ok} {set nprogEnts [expr {$nprogEnts + $entCount($thisEntType) - $count($thisEntType)}]}
            break
          }
        }
      }

      if {$opt(XLSCSV) == "CSV"} {catch {close $fcsv}}
      
# check for reports (validation properties, PMI presentation and representation, AP209)
      checkForReports $objDesign $entType
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
      catch {file delete -force $cfile}
    } else {
      set skipFile [open $cfile w]
      foreach item $fixtmp {puts $skipFile $item}
      close $skipFile
    }
  }

# -------------------------------------------------------------------------------------------------
# set viewpoints and close graphic PMI or FEM X3DOM file 
  if {($opt(VIZPMI) || $opt(VIZFEA)) && $x3domFileName != ""} {x3domViewpoints}

# -------------------------------------------------------------------------------------------------
# quit IFCsvr, but not sure how to do it properly
  if {[catch {
    #outputMsg "\nClosing IFCsvr" green
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
    [$worksheet($sum) Range "A[expr {$sumHeaderRow+3}]"] Select
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
      if {[info exists spmiTypesPerFile]} {
        set sempmi_coverage "PMI Representation Coverage"
        if {![info exists worksheet($sempmi_coverage)]} {
          outputMsg " Adding PMI Representation Coverage worksheet" green
          spmiCoverageStart 0
          spmiCoverageWrite "" "" 0
          spmiCoverageFormat "" 0
        }
      }

# format PMI Representation Summary worksheet
      if {[info exists spmiSumName]} {spmiSummaryFormat}
    }
  
# add PMI Pres. Coverage Analysis worksheet for a single file
    if {$opt(PMIGRF)} {
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
    
# -------------------------------------------------------------------------------------------------
# select the first tab
    [$worksheets Item [expr 1]] Select
    [$excel ActiveWindow] ScrollRow [expr 1]
  }

# processing time
  set cc [clock clicks -milliseconds]
  set proctime [expr {($cc - $lasttime)/1000}]
  if {$proctime <= 60} {set proctime [expr {(($cc - $lasttime)/100)/10.}]}
  outputMsg "Processing time: $proctime seconds" blue

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
      #outputMsg "Closing Excel" green
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
      if {[string first "The file or path name not found" $emsg] == -1} {
        #outputMsg " "
        errorMsg "The current version of the Spreadsheet needs to be closed before processing the STEP file again."
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
# open X3DOM file of graphical PMI or analysis model
  openX3DOM

# -------------------------------------------------------------------------------------------------
# save state
  if {[info exists errmsg]} {unset errmsg}
  saveState
  if {!$multiFile && [info exists buttons]} {$buttons(genExcel) configure -state normal}
  update idletasks

# clean up variables to hopefully release some memory and/or to reset them
  global colColor invCol currX3domPointID dimrep dimrepID entName gpmiID gpmiIDRow gpmiOK gpmiRow
  global heading invGroup nrep numX3domPointID pmiColumns pmiStartCol 
  global propDefID propDefIDRow propDefName propDefOK propDefRow syntaxErr 
  foreach var {cells colColor invCol count currX3domPointID dimrep dimrepID entName entsIgnored \
              gpmiID gpmiIDRow gpmiOK gpmiRow heading invGroup nrep numX3domPointID \
              pmiCol pmiColumns pmiStartCol pmivalprop propDefID propDefIDRow propDefName propDefOK propDefRow \
              syntaxErr workbook workbooks worksheet worksheets \
              x3domCoord x3domFile x3domFileName x3domFileOpen x3domIndex x3domMax x3domMin} {
    if {[info exists $var]} {unset $var}
  }
  if {!$multiFile} {
    foreach var {gpmiTypesPerFile spmiTypesPerFile} {if {[info exists $var]} {unset $var}}
  }
  update idletasks
  return 1
}
  
# -------------------------------------------------------------------------------------------------
proc addHeaderWorksheet {objDesign numFile fname} {
  global excel worksheets worksheet cells row timeStamp noPSA fileSchema cadSystem opt localName p21e3
  global excel1 worksheet1 cells1 col1
  global csvdirnam
   
  if {[catch {
    if {$opt(XLSCSV) == "Excel"} {
      outputMsg "Generating Header worksheet" blue
    } else {
      outputMsg "Generating Header CSV file" blue
    }
  
# all app names that might appear in header section
    set cadApps {"3D_Evolution" ACIS "Alias - OpenModel" "Alias AutoStudio" "Alias OpenModel" "Alias Studio" Alibre AutoCAD "Autodesk Inventor" \
      CADDS CADfix CADIF CATIA "CATIA V4" "CATIA V5" "CATIA V6" "CATIA Version 5" CgiStepCamp CoreTechnologie Creo "CV - CADDS 5" \
      DATAKIT Datakit "Datakit CrossCad" DATAVISION Elysium EXPRESSO FEMAP FiberSim HiCAD IDA-STEP "I-DEAS" "Implementor Forum Team" "ITI TranscenData" \
      "jt_step translator" Kubotek "Kubotek KeyCreator" "Mechanical Desktop" "Mentor Graphics" NX "OneSpace Designer" "Open CASCADE" \
      Parasolid Patran PlanetCAD PolyTrans "PRO/ENGINEER" Siemens "SIEMENS PLM Software NX 10.0" "SIEMENS PLM Software NX 11.0" \
      "SIEMENS PLM Software NX 7.0" "SIEMENS PLM Software NX 7.5" "SIEMENS PLM Software NX 8.0" "SIEMENS PLM Software NX 8.5" \
      "SIEMENS PLM Software NX 9.0" "SIEMENS PLM Software NX" "Solid Edge" SolidEdge SolidWorks "ST-ACIS" "STEP Caselib" \
      "STEP-NC Explorer" "STEP-NC Maker" "T3D tool generator" THEOREM Theorem "THEOREM SOLUTIONS" "Theorem Solutions" "T-Systems" \
      "UGS - NX" Unigraphics CoCreate Adobe Elysium ASFALIS CAPVIDIA 3DTransVidia MBDVidia NAFEMS COM209 CADCAM-E 3DEXPERIENCE ECCO SimDM \
      SDS/2 Tekla Revit RISA SAP2000 ETABS SmartPlant CADWorx "Advance Steel" ProSteel STAAD RAM Cype Parabuild RFEM RSTAB BuiltWorks EDMsix \
      "3D Reviewer" "3D Converter" HOOPS}

# sort cadApps by string length
    set cadApps [sortlength2 $cadApps]

    set cadSystem ""
    set timeStamp ""
    set noPSA 0
    set p21e3

    set hdr "Header"
    if {$opt(XLSCSV) == "Excel"} { 
      set worksheet($hdr) [$worksheets Item [expr 1]]
      $worksheet($hdr) Activate
      $worksheet($hdr) Name $hdr
      set cells($hdr) [$worksheet($hdr) Cells]

# create directory for CSV files
    } else {
      foreach var {csvdirnam csvfname fcsv} {catch {unset $var}}
      set csvdirnam "[file join [file dirname $localName] [file rootname [file tail $localName]]]-sfa-csv"
      file mkdir $csvdirnam
      set csvfname [file join $csvdirnam $hdr.csv]
      if {[file exists $csvfname]} {file delete -force $csvfname}
      set fcsv [open $csvfname w]
      #outputMsg $fcsv red
    }

    set row($hdr) 0
    foreach attr {Name FileDirectory FileDescription FileImplementationLevel FileTimeStamp FileAuthor \
                  FileOrganization FilePreprocessorVersion FileOriginatingSystem FileAuthorisation SchemaName} {
      incr row($hdr)
      if {$opt(XLSCSV) == "Excel"} { 
        $cells($hdr) Item $row($hdr) 1 $attr
      } else {
        set csvstr $attr
      }
      set objAttr [string trim [join [$objDesign $attr]]]

# FileDirectory
      if {$attr == "FileDirectory"} {
        if {$opt(XLSCSV) == "Excel"} { 
          $cells($hdr) Item $row($hdr) 2 [$objDesign $attr]
        } else {
          append csvstr ",[$objDesign $attr]"
          puts $fcsv $csvstr
        }
        outputMsg "$attr:  [$objDesign $attr]"

# SchemaName
      } elseif {$attr == "SchemaName"} {
        set sn [getSchemaFromFile $fname]
        if {$opt(XLSCSV) == "Excel"} { 
          $cells($hdr) Item $row($hdr) 2 $sn
        } else {
          append csvstr ",$sn"
          puts $fcsv $csvstr
        }
        outputMsg "$attr:  $sn" blue
        if {[string range $sn end-3 end] == "_MIM"} {
          errorMsg "Syntax Error: Schema name should end with _MIM_LF"
          [$worksheet($hdr) Range B11] Style "Bad"
       }

        set fileSchema  [string toupper [string range $objAttr 0 5]]
        if {[string first "IFC" $fileSchema] == 0} {
          errorMsg "Use the IFC File Analyzer with IFC files."
          after 1000
          openURL https://www.nist.gov/services-resources/software/ifc-file-analyzer
        } elseif {$objAttr == "STRUCTURAL_FRAME_SCHEMA"} {
          errorMsg "This is a CIS/2 file that can be visualized with SteelVis.\n https://www.nist.gov/services-resources/software/steelvis-aka-cis2-viewer"
        }

# other File attributes
      } else {
        if {$attr == "FileDescription" || $attr == "FileAuthor" || $attr == "FileOrganization"} {
          set str1 "$attr:  "
          set str2 ""
          foreach item [$objDesign $attr] {
            append str1 "[string trim $item], "
            if {$opt(XLSCSV) == "Excel"} { 
              append str2 "[string trim $item][format "%c" 10]"
            } else {
              append str2 ",[string trim $item]"
            }
          }
          outputMsg [string range $str1 0 end-2]
          if {$opt(XLSCSV) == "Excel"} { 
            $cells($hdr) Item $row($hdr) 2 "'[string trim $str2]"
            set range [$worksheet($hdr) Range "$row($hdr):$row($hdr)"]
            $range VerticalAlignment [expr -4108]
          } else {
            append csvstr [string trim $str2]
            puts $fcsv $csvstr
          }
        } else {
          outputMsg "$attr:  $objAttr"
          if {$opt(XLSCSV) == "Excel"} { 
            $cells($hdr) Item $row($hdr) 2 "'$objAttr"
            set range [$worksheet($hdr) Range "$row($hdr):$row($hdr)"]
            $range VerticalAlignment [expr -4108]
          } else {
            append csvstr ",$objAttr"
            puts $fcsv $csvstr
          }
        }

# check implementation level        
        if {$attr == "FileImplementationLevel"} {
          if {[string first "\;" $objAttr] == -1} {
            errorMsg "Syntax Error: Implementation Level is usually '2\;1'"
            [$worksheet($hdr) Range B4] Style "Bad"
          } elseif {$objAttr == "4\;1"} {
            set p21e3 1
          }
        }

# check for Creo or Inventor to prevent presentation_style_assignment being processed
        if {$attr == "FileOriginatingSystem"} {
          if {[string first "PRO/ENGINEER" $objAttr] != -1 || [string first "CREO PARAMETRIC" $objAttr] != -1 || \
              [string first "Inventor" $objAttr] != -1} {set noPSA 1}
        }

# check and add time stamp to multi file summary
        if {$attr == "FileTimeStamp"} {
          if {([string first "-" $objAttr] == -1 || [string length $objAttr] < 17) && $objAttr != ""} {
            errorMsg "Syntax Error: Wrong format for FileTimeStamp"            
            if {$opt(XLSCSV) == "Excel"} {[$worksheet($hdr) Range B5] Style "Bad"}
          }
          if {$numFile != 0 && [info exists cells1(Summary)] && $opt(XLSCSV) == "Excel"} {
            set timeStamp $objAttr
            set colsum [expr {$col1(Summary)+1}]
            set range [$worksheet1(Summary) Range [cellRange 5 $colsum]]
            catch {$cells1(Summary) Item 5 $colsum "'[string range $timeStamp 2 9]"}
          }
        }
      }
    }

    if {$opt(XLSCSV) == "Excel"} { 
      [[$worksheet($hdr) Range "A:A"] Font] Bold [expr 1]
      [$worksheet($hdr) Columns] AutoFit
      [$worksheet($hdr) Rows] AutoFit
      catch {[$worksheet($hdr) PageSetup] Orientation [expr 2]}
      catch {[$worksheet($hdr) PageSetup] PrintGridlines [expr 1]}
    }
      
# check for CAx-IF Recommended Practices in the file description
    set caxifrp {}
    foreach fd [$objDesign "FileDescription"] {
      set c1 [string first "CAx-IF Rec." $fd]
      if {$c1 != -1} {lappend caxifrp [string trim [string range $fd $c1+20 end]]}
    }
    if {[llength $caxifrp] > 0} {
      outputMsg "\nCAx-IF Recommended Practices: (www.cax-if.org/joint_testing_info.html#recpracs)"
      foreach item $caxifrp {
        outputMsg " $item" blue
        if {[string first "AP242" $fileSchema] == -1 && [string first "Tessellated" $item] != -1} {
          errorMsg "  Error: Recommended Practices related to 'Tessellated' only apply to AP242 files."
        }
      }
    }


# set the application from various file attributes, cadApps is a list of all application names defined above, take the first one that matches
    set ok 0
    foreach attr {FilePreprocessorVersion FileOriginatingSystem FileDescription FileAuthorisation FileOrganization} {
      foreach app $cadApps {
        set app1 $app
        if {$cadSystem == "" && [string first [string tolower $app] [string tolower [join [$objDesign $attr]]]] != -1} {
          set cadSystem [join [$objDesign $attr]]

# for multiple files, modify the app string to fit in file summary worksheet
          if {$numFile != 0 && [info exists cells1(Summary)]} {
            if {$app == "3D_Evolution"}            {set app1 "CT 3D Evolution"}
            if {$app == "CoreTechnologie"}         {set app1 "CT 3D Evolution"}
            if {$app == "DATAKIT"}                 {set app1 "Datakit"}
            if {$app == "Implementor Forum Team"}  {set app1 "CAx-IF"}
            if {$app == "jt_step translator"}      {set app1 "Siemens NX"}
            if {$app == "PRO/ENGINEER"}            {set app1 "Creo"}
            if {$app == "SIEMENS PLM Software NX"} {set app1 "Siemens NX"}
            if {$app == "UGS - NX"}                {set app1 "UGS-NX"}
            if {$app == "Unigraphics"}             {set app1 "Siemens NX"}
            if {$app == "UNIGRAPHICS"}             {set app1 "Unigraphics"}
            if {$app == "3DEXPERIENCE"}            {set app1 "CATIA"}
            if {[string first "CATIA Version"           $app] == 0} {set app1 "CATIA V[string range $app 14 end]"}
            if {[string first "SIEMENS PLM Software NX" $app] == 0} {set app1 "Siemens NX[string range $app 23 end]"}
            if {[string first "THEOREM"   [$objDesign FilePreprocessorVersion]] != -1} {set app1 "Theorem"}
            if {[string first "T-Systems" [$objDesign FilePreprocessorVersion]] != -1} {set app1 "T-Systems"}

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
    }
    
# add app2 to multiple file summary worksheet    
    if {$numFile != 0 && $opt(XLSCSV) == "Excel"} {
      if {$ok == 0} {set app2 [setCAXIFvendor]}
      set colsum [expr {$col1(Summary)+1}]
      if {$colsum > 16} {[$excel1 ActiveWindow] ScrollColumn [expr {$colsum-16}]}
      regsub -all " " $app2 [format "%c" 10] app2
      $cells1(Summary) Item 6 $colsum [string trim $app2]
    }
    if {$cadSystem == ""} {set cadSystem [setCAXIFvendor]}

# close csv file
    if {$opt(XLSCSV) == "CSV"} {close $fcsv} 

  } emsg]} {
    errorMsg "ERROR adding Header worksheet: $emsg"
    catch {raise .}
  }
}

#-------------------------------------------------------------------------------------------------
# add summary worksheet
proc sumAddWorksheet {} {
  global worksheet cells sum sheetSort sheetLast col worksheets row entCategory opt entsIgnored excel
  global x3domFileName spmiEntity entCount gpmiEnts spmiEnts nistVersion
  global propDefRow stepAP

  outputMsg "\nGenerating Summary worksheet" blue
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
        
# for STEP add [Properties], [PMI Presentation], [PMI Representation] text string and link to X3DOM file    
        set okao 0
        if {$entType == "property_definition" && $col($entType) > 4} {
          $cells($sum) Item $sumRow 1 "property_definition  \[Properties\]"
        } elseif {$entType == "dimensional_characteristic_representation" && $col($entType) > 3} {
          $cells($sum) Item $sumRow 1 "dimensional_characteristic_representation  \[PMI Representation\]"
        } elseif {[lsearch $spmiEntity $entType] != -1} {
          $cells($sum) Item $sumRow 1 "$entType  \[PMI Representation\]"
        } elseif {[string first "annotation" $entType] != -1} {
          if {$gpmiEnts($entType) && $col($entType) > 5} {set okao 1}
        }
        if {$okao} {
          $cells($sum) Item $sumRow 1 "$entType  \[PMI Presentation\]"
          if {$opt(VIZPMI) && $x3domFileName != "" && $x3domLink} {
            $cells($sum) Item $sumRow 3 "Graphic PMI"
            set x3domLink 0
            set anchor [$worksheet($sum) Range [cellRange $sumRow 3]]
            $sumLinks Add $anchor [join $x3domFileName] [join ""] [join "Link to Graphic PMI"]
          }
        }

# for '_and_' (complex entity) split on multiple lines
# '10' is the ascii character for a linefeed          
      } else {
        regsub -all "_and_" $entType ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" entType_multiline
        set entType_multiline "($entType_multiline)"
        $cells($sum) Item $sumRow 1 $entType_multiline

# for STEP add [Properties] or [PMI Presentation] text string and link to X3DOM file    
        set okao 0
        if {[string first "annotation" $entType] != -1} {
          if {$gpmiEnts($entType) && $col($entType) > 7} {set okao 1}
        } elseif {[lsearch $spmiEntity $entType] != -1} {
          $cells($sum) Item $sumRow 1 "$entType_multiline  \[PMI Representation\]"
        }
        if {$okao} {
          $cells($sum) Item $sumRow 1 "$entType_multiline  \[PMI Presentation\]"
          if {$opt(VIZPMI) && $x3domFileName != "" && $x3domLink} {
            $cells($sum) Item $sumRow 3 "Graphic PMI"
            set x3domLink 0
            set anchor [$worksheet($sum) Range [cellRange $sumRow 3]]
            $sumLinks Add $anchor [join $x3domFileName] [join ""] [join "Link to Graphic PMI"]
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
    set str "NIST "
    set url "https://www.nist.gov/services-resources/software/step-file-analyzer"
    if {!$nistVersion} {
      set str ""
      set url "https://github.com/usnistgov/SFA"
    }
    $cells($sum) Item [expr {$row($sum)+2}] 1 "Spreadsheet generated by the $str\STEP File Analyzer (v[getVersion])"
    set anchor [$worksheet($sum) Range [cellRange [expr {$row($sum)+2}] 1]]
    [$worksheet($sum) Hyperlinks] Add $anchor [join $url] [join ""] [join "Link to STEP File Analyzer"]
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
      set cidx [setColorIndex $ent]
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
  global gpmiEnts spmiEnts nprogEnts excelVersion
  
  outputMsg "Formatting Worksheets" blue

  if {[info exists buttons]} {$buttons(pgb) configure -maximum [llength $sheetSort]}
  set nprogEnts 0
  set nsort 0
  foreach thisEntType $sheetSort {
    #getTiming "START FORMATTING $thisEntType"
    incr nprogEnts
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

# add PMI Representation Summary worksheet
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
  
# -------------------------------------------------------------------------------------------------
# add table for sorting and filtering
      if {$excelVersion >= 12} {
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
