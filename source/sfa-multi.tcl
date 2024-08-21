# process multiple files in a directory
proc openMultiFile {{ask 1}} {
  global allEntity andEntAP209 buttons cells1 col1 coverageSTEP entCategory excel1 fileDir fileDir1 fileEntity fileList gen
  global gpmiCoverageWS gpmiRows gpmiTotals gpmiTypes gpmiTypesInvalid lastXLS1 lenfilelist localName localNameList multiFileDir mydocs nfile
  global nistCoverageStyle nprogBarFiles opt pmiElementsMaxRows row1 spmiCoverageWS startrow stepAP totalEntity totalPMI totalPMIrows
  global useXL worksheet1 worksheets1 xlFileNames

  set multiFileDir ""

# select directory of files (default)
  if {$ask == 1} {
    if {![file exists $fileDir1] && [info exists mydocs]} {set fileDir1 $mydocs}
    set multiFileDir [tk_chooseDirectory -title "Select Directory of STEP Files" -mustexist true -initialdir $fileDir1]
    if {[info exists localNameList]} {unset localNameList}

# list of files
  } elseif {$ask == 2} {
    set multiFileDir [file dirname [lindex $localNameList 0]]
    set dlen [expr {[string length [truncFileName $multiFileDir]]+1}]
    set fileList [lsort -dictionary $localNameList]
    set lenfilelist [llength $localNameList]

# don't ask for F4
  } elseif {$ask == 0} {
    if {$lastXLS1 != ""} {set multiFileDir $fileDir1}
    if {[info exists localNameList]} {unset localNameList}
  }

  if {$multiFileDir != "" && [file isdirectory $multiFileDir]} {
    if {$ask != 2} {
      outputMsg "\nSTEP file directory: [truncFileName [file nativename $multiFileDir]]" blue
      set dlen [expr {[string length [truncFileName $multiFileDir]]+1}]
      .tnb select .tnb.status
      set fileDir1 $multiFileDir
      saveState

      set recurse 0
      if {$ask} {
        set choice [tk_messageBox -title "Search Subdirectories?" -type yesno -default yes \
                    -message "Do you want to process STEP files in subdirectories too?" -icon question]
      } else {
        set choice "yes"
      }
      if {$choice == "yes"} {
        set recurse 1
        outputMsg "Searching subdirectories ..."
        update
      }

# find all files in directory and subdirectories
      set fileList {}
      findFile $multiFileDir $recurse
      set lenfilelist [llength $fileList]

# list files and size
      foreach file1 $fileList {
        outputMsg "  [string range [file nativename [truncFileName $file1]] $dlen end]  ([fileSize $file1]  [fileTime $file1])"
      }
    }

    if {$lenfilelist > 0} {
      if {$ask != 2} {outputMsg "($lenfilelist) STEP files found" blue}
      set askstr ""
      if {$opt(xlFormat) == "Excel"} {
        append askstr "Spreadsheets"
      } elseif {$opt(xlFormat) == "CSV"} {
        if {$gen(CSV)} {append askstr "Spreadsheets and "}
        append askstr "CSV Files"
      } elseif {$opt(xlFormat) == "None"} {
        append askstr "Views"
      }
      if {$opt(xlFormat) != "None" && $gen(View) && ($opt(viewPart) || $opt(viewFEA) || $opt(viewPMI) || $opt(viewTessPart))} {
        append askstr " and Views"
      }

      if {$ask != 2} {
        set choice [tk_messageBox -title "Generate $askstr?" -type yesno -default yes -message "Do you want to Generate $askstr for ($lenfilelist) STEP files ?" -icon question]
      } else {
        set choice "yes"
      }

      if {$choice == "yes"} {
        checkForExcel
        set lasttime1 [clock clicks -milliseconds]
        checkValues

# start Excel for summary of all files
        set fileDir $multiFileDir
        if {$lenfilelist > 1 && $useXL && $opt(xlFormat) != "None"} {
          if {[catch {
            set pid2 [twapi::get_process_ids -name "EXCEL.EXE"]
            set excel1 [::tcom::ref createobject Excel.Application]
            set pidExcel1 [lindex [intersect3 $pid2 [twapi::get_process_ids -name "EXCEL.EXE"]] 2]
            [$excel1 ErrorCheckingOptions] TextDate False

            set mf [expr {2**14}]
            set extXLS "xlsx"
            set xlFormat [expr 51]

            set mf [expr {$mf-3}]
            if {$lenfilelist > $mf} {
              errorMsg "Only the first $mf files will be processed due to column limits in Excel."
              set lenfilelist $mf
              set fileList [lrange $fileList 0 [expr {$mf-1}]]
            }

            $excel1 Visible 1

# errors
          } emsg]} {
            set useXL 0
            if {$opt(xlFormat) == "Excel"} {
              errorMsg "Excel is not installed or cannot start Excel: $emsg\n CSV files will be generated instead of a spreadsheet.  See the Generate options.  Some options are disabled."
              set opt(xlFormat) "CSV"
              catch {raise .}
            }
            checkValues
          }

# start summary/analysis spreadsheet
          if {$useXL} {
            if {[catch {
              outputMsg "\nStarting File Summary spreadsheet" blue
              set workbooks1  [$excel1 Workbooks]
              set workbook1   [$workbooks1 Add]
              set worksheets1 [$workbook1 Worksheets]

# determine how many worksheets to add for coverage analysis
              set n1 1
              set coverageSTEP 0
              if {($opt(PMIGRF) && $opt(PMIGRFCOV)) || $opt(PMISEM)} {
                set coverageSTEP 1
                if {$opt(PMIGRF) && $opt(PMIGRFCOV) && $opt(PMISEM)} {
                  set n1 3
                } else {
                  set n1 2
                }
# make sure there are at least 3 worksheets
                if {[$worksheets1 Count] < 3} {$worksheets1 Add; $worksheets1 Add}
              }

# delete 0, 1, or 2 worksheets (0 or 1 for STEP CA)
              catch {$excel1 DisplayAlerts False}
              set sheetCount [$worksheets1 Count]
              for {set n $sheetCount} {$n > $n1} {incr n -1} {[$worksheets1 Item [expr $n]] Delete}
              catch {$excel1 DisplayAlerts True}

# -------------------------------------------------------------------------
# start file summary worksheet
              set sum "Summary"
              set worksheet1($sum) [$worksheets1 Item [expr 1]]
              $worksheet1($sum) Activate
              $worksheet1($sum) Name "File Summary"
              set cells1($sum) [$worksheet1($sum) Cells]
              $cells1($sum) Item 1 1 "STEP Directory"
              $cells1($sum) Item 1 2 "[file nativename $multiFileDir]"
              set range [$worksheet1($sum) Range "B1:K1"]
              $range MergeCells [expr 1]

# set startrow
              set startrow 9
              $cells1($sum) Item $startrow 1 "Entity"
              set col1($sum) 1

# orientation for file info
              set range [$worksheet1($sum) Range "5:$startrow"]
              $range VerticalAlignment [expr -4107]
              $range HorizontalAlignment [expr -4108]

# vertical orientation for file name
              set range [$worksheet1($sum) Range "4:4"]
              $range Orientation [expr 90]
              $range HorizontalAlignment [expr -4108]

              if {!$coverageSTEP} {
                [$excel1 ActiveWindow] TabRatio [expr 0.3]
              } else {
                [$excel1 ActiveWindow] TabRatio [expr 0.6]
              }

# start STEP coverage analysis worksheet
              if {$coverageSTEP} {
                if {$opt(PMISEM)} {spmiCoverageStart}
                if {$opt(PMIGRF) && $opt(PMIGRFCOV)} {gpmiCoverageStart}
              }
              $worksheet1($sum) Activate

# errors
            } emsg]} {
              errorMsg "Error opening Excel workbooks and worksheets for file summary: $emsg"
              catch {raise .}
            }
          } elseif {$useXL && $opt(xlFormat) != "None"} {
            errorMsg "For only one STEP file, no File Summary spreadsheet is generated."
          }
        }

# -------------------------------------------------------------------------------------------------
# loop over all the files and process
        foreach var {fileEntity gpmiRows gpmiTotals totalPMI totalPMIrows totalEntity} {if {[info exists $var]} {unset $var}}
        set xlFileNames {}
        set allEntity {}
        set dirchange {}
        catch {unset nistCoverageStyle}
        set lastdirname ""
        set nfile 0
        set gpmiTypesInvalid {}
        set sum "Summary"

        set nprogBarFiles 0
        pack $buttons(progressBarMulti) -side top -padx 10 -pady {5 0} -expand true -fill x
        $buttons(progressBarMulti) configure -maximum $lenfilelist
        update

# start loop over multiple files
        foreach file1 $fileList {
          incr nfile
          set stat($nfile) 0
          set localName $file1

          outputMsg "\n-------------------------------------------------------------------------------"
          outputMsg "($nfile of $lenfilelist) Ready to process: [file tail $file1]  ([fileSize $file1]  [fileTime $file1])" green

# check for zipped file
          if {[string first ".stpz" [string tolower $localName]] != -1} {unzipFile}

# process the file
          if {[catch {
            set stat($nfile) [genExcel $nfile]

# error processing the file
          } emsg]} {
            errorMsg "Error processing [file tail $file1]: $emsg"
            catch {raise .}
            set stat($nfile) 0
          }

# set fn from file name (file1), change \ to linefeed
          if {$lenfilelist > 1 && $useXL && $opt(xlFormat) != "None"} {
            set fn [string range [file nativename [truncFileName $file1]] $dlen end]
            regsub -all {\\} $fn [format "%c" 10] fn
            incr col1($sum)

# keep track of changes in directory name to have vertical line when directory changes
            set dirname [file dirname $file1]
            if {$lastdirname != "" && $dirname != $lastdirname} {lappend dirchange [expr {$nfile+1}]}
            set lastdirname $dirname

# STEP coverage analysis
            if {$coverageSTEP} {
              if {$opt(PMISEM)} {spmiCoverageWrite $fn $sum}
              if {$opt(PMIGRF) && $opt(PMIGRFCOV)} {gpmiCoverageWrite $fn $sum}
            }
          }

# done adding coverage analysis results
# -------------------------------------------------------------------------
          incr nprogBarFiles
        }

# -------------------------------------------------------------------------------------------------
# time to generate spreadsheets
        set ptime [expr {([clock clicks -milliseconds] - $lasttime1)/1000}]
        if {$ptime < 120} {
          set ptime "$ptime seconds"
        } elseif {$ptime < 3600} {
          set ptime "[trimNum [expr {double($ptime)/60.}] 1] minutes"
        } else {
          set ptime "[trimNum [expr {double($ptime)/3600.}] 1] hours"
        }
        set msg "\n($nfile) "
        if {$opt(xlFormat) == "None"} {
          append msg "Views"
        } elseif {$useXL} {
          append msg "Spreadsheets"
          if {$gen(View)} {append msg " and Views"}
        } elseif {$opt(xlFormat) == "CSV"} {
          append msg "CSV files"
        }
        append msg " generated in $ptime"
        outputMsg $msg green
        outputMsg "-------------------------------------------------------------------------------"

# -------------------------------------------------------------------------------------------------
# file summary ws, entity names
        if {$lenfilelist > 1 && $useXL && $opt(xlFormat) != "None"} {
          catch {$excel1 ScreenUpdating 0}
          outputMsg "\nWriting File Summary information" blue
          if {[catch {
            [$excel1 ActiveWindow] ScrollColumn [expr 1]
            set wid 2
            if {$lenfilelist > 16} {set wid 3}
            if {$lenfilelist > 1}  {incr wid}

            set row1($sum) $startrow
            set allEntity [lsort [lrmdups $allEntity]]
            set links [$worksheet1($sum) Hyperlinks]
            set inc1 0

# entity names, split on _and_
            for {set i 0} {$i < [llength $allEntity]} {incr i} {
              set ent [string range [lindex $allEntity $i] 2 end]
              lset allEntity $i $ent
              set ok 0
              if {[string first "_and_" $ent] == -1} {
                set ok 1
              } else {
                foreach item [array names entCategory] {if {[lsearch $entCategory($item) $ent] != -1} {set ok 1}}
                if {[string first "AP209" $stepAP] != -1} {foreach str $andEntAP209 {if {[string first $str $ent] != -1} {set ok 1}}}
              }
              if {$ok} {
                $cells1($sum) Item [incr row1($sum)] 1 $ent
                set ent2 $ent
              } else {
# '10' is the ascii character for a linefeed
                regsub -all "_and_" $ent ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" ent1
                $cells1($sum) Item [incr row1($sum)] 1 "($ent1)"
                set range [$worksheet1($sum) Range $row1($sum):$row1($sum)]
                $range VerticalAlignment [expr -4108]
                set ent2 "($ent1)"
              }
              set entrow($ent) $row1($sum)

# if more than 16 file, repeat entity names on the right
              if {$lenfilelist > 16} {
                $cells1($sum) Item $row1($sum) [expr {$lenfilelist+$wid+$inc1}] $ent2
              }
            }

#-------------------------------------------------------------------------------
# fix wrap for vertical file names
            set range [$worksheet1($sum) Range "4:4"]
            $range WrapText [expr 0]
            $range WrapText [expr 1]

# format file summary worksheet
            set range [$worksheet1($sum) Range [cellRange 3 1] [cellRange [expr {[llength $allEntity]+$startrow}] [expr {$lenfilelist+$wid+$inc1}]]]
            $range AutoFormat
            set range [$worksheet1($sum) Range "5:$startrow"]
            $range VerticalAlignment [expr -4107]
            $range HorizontalAlignment [expr -4108]

# erase some extra horizontal lines that AutoFormat created, but it doesn't work
            if {$lenfilelist > 2} {
              set range [$worksheet1($sum) Range [cellRange [expr {$startrow-1}] 1] [cellRange [expr {$startrow-1}] [expr {$lenfilelist+$wid+$inc1}]]]
              set borders [$range Borders]
              catch {
                [$borders Item [expr 8]] Weight [expr -4142]
                [$borders Item [expr 9]] Weight [expr -4142]
              }
            }

# generated by
            set c [expr {[llength $allEntity]+$startrow+2}]
            $cells1($sum) Item $c 1 "NIST STEP File Analyzer and Viewer [getVersion]"
            set anchor [$worksheet1($sum) Range [cellRange $c 1]]
            [$worksheet1($sum) Hyperlinks] Add $anchor [join "https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer"] [join ""] \
              [join "Link to NIST STEP File Analyzer and Viewer"]
            $cells1($sum) Item [expr {[llength $allEntity]+$startrow+3}] 1 "[clock format [clock seconds]]"

# entity counts
            if {[info exists infiles]} {unset infiles}
            foreach idx [lsort -integer [array names fileEntity]] {
              set col1($sum) [expr {$idx+1}]
              set scrollcol 16
              if {$col1($sum) > $scrollcol} {[$excel1 ActiveWindow] ScrollColumn [expr {$col1($sum)-$scrollcol}]}
              foreach item $fileEntity($idx) {
                set val [split $item " "]
                $cells1($sum) Item $entrow([lindex $val 0]) $col1($sum) [expr {abs([lindex $val 1])}]

# color entity count gray if there are errors associated with the entity
                if {[lindex $val 1] < 0} {
                  set cr [cellRange $entrow([lindex $val 0]) $col1($sum)]
                  [[$worksheet1($sum) Range $cr] Interior] ColorIndex [expr 15]
                  if {![info exists addcomm]} {
                    set comm [[$worksheet1($sum) Range $cr] AddComment]
                    $comm Text "There are Errors or Warnings for at least one entity of this type."
                    catch {[[$comm Shape] TextFrame] AutoSize [expr 1]}
                    set addcomm 1
                  }
                }
                incr infiles($entrow([lindex $val 0]))
              }
            }
            catch {unset addcomm}

# entity totals
            set col1($sum) [expr {$lenfilelist+2}]
            $cells1($sum) Item $startrow $col1($sum) "Total[format "%c" 10]Entities"
            foreach idx [array names totalEntity] {
              $cells1($sum) Item $entrow($idx) $col1($sum) $totalEntity($idx)
            }

# file occurances
            if {$lenfilelist > 1} {
              $cells1($sum) Item $startrow [incr col1($sum)] "Total[format "%c" 10]Files"
            }
            foreach idx [array names infiles] {
              $cells1($sum) Item $idx $col1($sum) $infiles($idx)
            }
            [$excel1 ActiveWindow] ScrollColumn [expr 1]

# bold text
            set range [$worksheet1($sum) Range [cellRange 5 1] [cellRange $startrow [expr {$col1($sum)}]]]
            [$range Font] Bold [expr 1]
            [[$worksheet1($sum) Range "B1"] Font] Bold [expr 1]

            [$worksheet1($sum) Columns] AutoFit

#-------------------------------------------------------------------------------
# color entity names in first column
            set c1 [expr {[llength $fileList]+$wid+$inc1}]
            for {set i 0} {$i < [llength $allEntity]} {incr i} {
              set ent [lindex $allEntity $i]
              set range [$worksheet1($sum) Range "A$entrow($ent)"]
              set cidx [setColorIndex $ent 1]
              if {$cidx > 0} {
                [$range Interior] ColorIndex [expr $cidx]
                catch {
                  [[$range Borders] Item [expr 8]] Weight [expr 1]
                  [[$range Borders] Item [expr 9]] Weight [expr 1]
                }
                if {$lenfilelist > 16} {
                  set range1 [$worksheet1($sum) Range [cellRange $entrow($ent) $c1]]
                  [$range1 Interior] ColorIndex [expr $cidx]
                  catch {
                    [[$range1 Borders] Item [expr 8]] Weight [expr 1]
                    [[$range1 Borders] Item [expr 9]] Weight [expr 1]
                  }
                }
              }
            }

#-------------------------------------------------------------------------------
# link to STEP file, link to individual spreadsheet
            set nf 1
            set idx -1
            foreach file1 $fileList {
              incr nf
              if {$stat([expr {$nf-1}]) != 0} {

# link to file
                if {!$opt(xlHideLinks) && [string first "#" $file1] == -1} {
                  set range [$worksheet1($sum) Range [cellRange 4 $nf]]
                  $links Add $range [join $file1] [join ""] [join "Link to STEP file"]
                }

# link to spreadsheet
                set range [$worksheet1($sum) Range [cellRange 3 $nf]]
                incr idx
                regsub -all {\\} [lindex $xlFileNames $idx] "/" xls
                if {!$opt(xlHideLinks) && [string first "#" $file1] == -1} {$links Add $range [join $xls] [join ""] [join "Link to Spreadsheet"]}

# add vertical border when directory changes from column to column
                if {[lsearch $dirchange $nf] != -1} {
                  set nf1 [expr {$nf-1}]
                  set range [$worksheet1($sum) Range [cellRange 3 $nf1] [cellRange $row1($sum) $nf1]]
                  set borders [$range Borders]
                  catch {[$borders Item [expr -4152]] Weight [expr 2]}

# also for PMI coverage analysis worksheets
                  catch {
                    set range [$worksheet1($spmiCoverageWS) Range [cellRange 3 $nf1] [cellRange $pmiElementsMaxRows $nf1]]
                    set borders [$range Borders]
                    [$borders Item [expr -4152]] Weight [expr 2]
                  }
                  catch {
                    if {$opt(PMIGRFCOV)} {
                      set range [$worksheet1($gpmiCoverageWS) Range [cellRange 3 $nf1] [cellRange $gpmiRows $nf1]]
                      set borders [$range Borders]
                      [$borders Item [expr -4152]] Weight [expr 2]
                    }
                  }
                }
              }
            }
            set range [$worksheet1($sum) Range [cellRange 3 [expr {$lenfilelist+1}]] [cellRange $row1($sum) [expr {$lenfilelist+1}]]]
            set borders [$range Borders]
            catch {[$borders Item [expr -4152]] Weight [expr 2]}

# fix column widths
            set c1 [[[[$worksheets1 Item [expr 1]] UsedRange] Columns] Count]
            for {set i 2} {$i <= $c1} {incr i} {
              set val [[$cells1($sum) Item 6 $i] Value]
              if {$val != ""} {
                set range [$worksheet1($sum) Range [cellRange -1 $i]]
                $range ColumnWidth [expr 24]
              }
            }
            [$worksheet1($sum) Columns] AutoFit
            [$worksheet1($sum) Rows] AutoFit

# freeze panes
            [$worksheet1($sum) Range "B[expr {$startrow+1}]"] Select
            catch {[$excel1 ActiveWindow] FreezePanes [expr 1]}
            [$worksheet1($sum) Range "A1"] Select

# errors
          } emsg]} {
            errorMsg "Error adding information to File Summary spreadsheet: $emsg"
            catch {raise .}
          }

# -------------------------------------------------------------------------------------------------
# format STEP coverage analysis sheet
          if {$coverageSTEP} {
            if {$opt(PMISEM)} {spmiCoverageFormat $sum}
            if {$opt(PMIGRF) && $opt(PMIGRFCOV)} {gpmiCoverageFormat $sum}
            catch {$worksheet1($sum) Activate}
          }
          catch {$excel1 ScreenUpdating 1}
        }

# -------------------------------------------------------------------------------------------------
# save spreadsheet
        if {$lenfilelist > 1 && $useXL && $opt(xlFormat) != "None"} {
          if {[catch {

# set file name for analysis spreadsheet
            set enddir [lindex [split $multiFileDir "/"] end]
            regsub -all " " $enddir "_" enddir
            set aname [file nativename [file join $multiFileDir SFA-Summary-$enddir-$lenfilelist.$extXLS]]
            if {[string length $aname] > 218} {
              errorMsg "Spreadsheet file name is too long for Excel ([string length $aname])."
              set aname [file nativename [file join $mydocs SFA-Summary-$enddir-$lenfilelist.$extXLS]]
              if {[string length $aname] < 219} {errorMsg " Spreadsheet written to the home directory."}
            }
            catch {file delete -force -- $aname}

# check if file exists and create new name
            if {[file exists $aname]} {set aname [incrFileName $aname]}

# save spreadsheet
            set aname [checkFileName $aname]
            outputMsg "Saving File Summary Spreadsheet to:"
            outputMsg " [truncFileName $aname 1]" blue
            catch {$excel1 DisplayAlerts False}
            if {$xlFormat == 51} {
              $workbook1 -namedarg SaveAs Filename $aname FileFormat $xlFormat
            } else {
              $workbook1 -namedarg SaveAs Filename $aname
            }
            catch {$excel1 DisplayAlerts True}
            set lastXLS1 $aname

# close Excel
            $excel1 Quit
            update idletasks
            catch {unset excel1}
            if {[llength $pidExcel1] == 1} {catch {twapi::end_process $pidExcel1 -force}}

# errors
          } emsg]} {
            errorMsg "Error saving File Summary Spreadsheet: $emsg"
            catch {raise .}
          }

# open spreadsheet
          if {$opt(outputOpen)} {
            openXLS $aname 0 1
            if {!$opt(xlHideLinks)} {outputMsg " Click on the Links in Row 3 to open individual spreadsheets." blue}
          } else {
            outputMsg " Use F7 to open the spreadsheet (see Help > Function Keys)" red
          }

# unset some variables for the multi-file summary
          foreach var {cells1 col1 excel1 lenfilelist row1 worksheet1 worksheets1} {if {[info exists $var]} {unset $var}}
        }
        update idletasks

        saveState
        $buttons(generate) configure -state normal
      }

# no files found
    } elseif {[info exists recurse]} {
      set substr ""
      if {$recurse} {set substr " or subdirectories of"}
      errorMsg "No STEP files were found in the directory$substr: [truncFileName [file nativename $multiFileDir]]"
    }
  }
  update idletasks
}

#-------------------------------------------------------------------------------
proc findFile {startDir {recurse 0}} {
  global fileList

  set pwd [pwd]
  if {[catch {cd $startDir} err]} {
    errorMsg $err
    return
  }

  set exts {".stp" ".step" ".p21" ".stpz" ".ifc"}
  foreach match [glob -nocomplain -- *] {
    foreach ext $exts {
      if {[file extension [string tolower $match]] == $ext} {
        if {$ext != ".stpz" || ![file exists [string range $match 0 end-1]]} {
          lappend fileList [file join $startDir $match]
        }
      }
    }
  }
  if {$recurse} {
    foreach file [glob -nocomplain *] {
      if {[file isdirectory $file]} {findFile [file join $startDir $file] $recurse}
    }
  }
  cd $pwd
}
