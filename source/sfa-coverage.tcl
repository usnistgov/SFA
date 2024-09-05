# Semantic PMI Summary worksheet
proc spmiSummary {} {
  global allPMI cells entName env epmi epmiUD localName mytemp nistName nistPMIexpected opt pmiModifiers recPracNames row sheetLast
  global spmiSumName spmiSumRow spmiSumRowID spmiTypesPerFile thisEntType timeStamp valRounded wdir worksheet worksheets

# first time through, start worksheet
  if {$spmiSumRow == 1} {
    outputMsg " Adding Semantic PMI Summary worksheet" blue

    set spmiSumName "Semantic PMI Summary"
    set worksheet($spmiSumName) [$worksheets Add [::tcom::na] $sheetLast]
    $worksheet($spmiSumName) Activate
    $worksheet($spmiSumName) Name $spmiSumName
    set cells($spmiSumName) [$worksheet($spmiSumName) Cells]
    set wsCount [$worksheets Count]
    [$worksheets Item [expr $wsCount]] -namedarg Move Before [$worksheets Item [expr 3]]

    for {set i 2} {$i <= 3} {incr i} {[$worksheet($spmiSumName) Range [cellRange -1 $i]] ColumnWidth [expr 48]}
    for {set i 1} {$i <= 4} {incr i} {[$worksheet($spmiSumName) Range [cellRange -1 $i]] VerticalAlignment [expr -4160]}

    set txt [file tail $localName]
    if {$timeStamp != ""} {append txt "[format "%c" 10]  ($timeStamp)"}
    $cells($spmiSumName) Item $spmiSumRow 2 $txt
    incr spmiSumRow 2
    $cells($spmiSumName) Item $spmiSumRow 1 "ID"
    $cells($spmiSumName) Item $spmiSumRow 2 "Entity"
    $cells($spmiSumName) Item $spmiSumRow 3 "Semantic PMI"

    set comment "Semantic PMI summarizes results from Datum System, Dimension, Geometric Tolerance, and Datum Target entities.  See Help > User Guide (section 6.1.6)"
    if {$valRounded} {append comment "\n\nSome dimension or tolerance values are rounded."}
    if {$nistName != ""} {
      append comment "\n\nIt is color-coded by the expected PMI in the NIST test case drawing.  The color-coding is explained at the bottom of the column.  Determining if the PMI is Partial and Possible match and corresponding Similar PMI depends on leading and trailing zeros, number precision, associated datum features and dimensions, and repetitive dimensions.\n\nSee Help > Analyzer > NIST CAD Models\nSee Help > User Guide (section 6.6.1)"
    } else {
      append comment "\n\nCompare this PMI to what is in your CAD model or the Viewer"
    }
    append comment "."
    addCellComment $spmiSumName $spmiSumRow 3 $comment

    set range [$worksheet($spmiSumName) Range [cellRange 1 1] [cellRange 3 3]]
    [$range Font] Bold [expr 1]
    set range [$worksheet($spmiSumName) Range [cellRange 3 1] [cellRange 3 3]]
    catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 2]}}
    incr spmiSumRow

    $cells($spmiSumName) Item 1 3 "See CAx-IF Recommended Practice for $recPracNames(pmi242)"
    set range [$worksheet($spmiSumName) Range C1:K1]
    $range MergeCells [expr 1]
    set anchor [$worksheet($spmiSumName) Range C1]
    [$worksheet($spmiSumName) Hyperlinks] Add $anchor [join "https://www.mbx-if.org/home/cax/recpractices/"] [join ""] [join "Link to CAx-IF Recommended Practices"]

    set allPMI ""

# get expected PMI for a NIST model
    set epmi ""
    if {$nistName != ""} {
      set epmi $nistName
    } elseif {$epmiUD != ""} {
      set epmi $epmiUD
    }
    nistGetSummaryPMI $epmi

# check for font file with GD&T symbols (ARIALUNI.TTF), needed only for certain symbols in pmiTypes
    if {![file exists [file nativename C:/Windows/Fonts/ARIALUNI.TTF]] && \
        ![file exists [file join $env(USERPROFILE) AppData Local Microsoft Windows Fonts ARIALUNI.TTF]]} {
      if {![file exists [file join $mytemp ARIALUNI.TTF]]} {catch {[file copy -force -- [file join $wdir images ARIALUNI.TTF] [file join $mytemp ARIALUNI.TTF]]}}
      set pmiTypes [list cylindricity_tolerance symmetry_tolerance all_around projected unequally_disposed free_state tangent_plane least_material_requirement \
        envelope_requirement independency reciprocity_requirement regardless_of_feature_size derived_feature associated_least_square_feature \
        associated_maximum_inscribed_feature associated_minimum_inscribed_feature associated_minmax_feature associated_tangent_feature circle_a]
      set ok 0
      set str ""
      foreach pmi $pmiTypes {if {[lsearch $spmiTypesPerFile $pmi] != -1} {set ok 1; set str " ($pmi)"; break}}
      if {$opt(SHOWALLPMI) || $ok} {
        errorMsg "Some GD&T symbols$str will appear as a question mark on the Semantic PMI\n Summary and Coverage worksheets.  To fix the problem, copy the font file that contains the symbols\n [file join $mytemp ARIALUNI.TTF]  to  C:/Windows/Fonts  to install the fonts.\n You might need administrator privileges."
      }
    }
  }

# add to PMI summary worksheet
  set hlink [$worksheet($spmiSumName) Hyperlinks]
  for {set i 3} {$i <= $row($thisEntType)} {incr i} {

# which entities are processed, check for holes
    set okent 0
    if {$thisEntType != "datum_reference_compartment" && $thisEntType != "datum_reference_element" && \
        $thisEntType != "datum_reference_modifier_with_value" && [string first "datum_feature" $thisEntType] == -1} {
      set okent 1
    }
    set notHole 1
    if {[string first "counter" $thisEntType] != -1 || [string first "spotface" $thisEntType] != -1} {set notHole 0}

# which entities and columns to summarize
    if {$okent} {
      if {$i == 3} {
        set j1 [getNextUnusedColumn $thisEntType]
        for {set j 1} {$j < $j1} {incr j} {
          set val [[$cells($thisEntType) Item 3 $j] Value]
          if {[string first "Datum Reference Frame" $val] != -1 || \
              $val == "GD&T[format "%c" 10]Annotation" || \
              $val == "Dimensional[format "%c" 10]Tolerance" || \
              $val == "Hole[format "%c" 10]Feature" || \
              [string first "Datum Target" $val] == 0 || \
              ($thisEntType == "datum_reference" && [string first "reference" $val] != -1) || \
              ($thisEntType == "referenced_modified_datum" && [string first "datum" $val] != -1)} {set pmiCol $j}
        }

# values
      } else {
        if {[info exists pmiCol]} {
          set id [expr {int([[$cells($thisEntType) Item $i 1] Value])}]
          $cells($spmiSumName) Item $spmiSumRow 1 $id
          set spmiSumRowID($id) $spmiSumRow
          if {[string first "_and_" $thisEntType] == -1} {
            set entstr $thisEntType
          } else {
            regsub -all "_and_" $thisEntType ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" entstr
            set entstr "($entstr)"
          }
          $cells($spmiSumName) Item $spmiSumRow 2 $entstr

          set val [[$cells($thisEntType) Item $i $pmiCol] Value]

# remove (Invalid TZF: ...)
          set c1 [string first "(Invalid TZF:" $val]
          if {$c1 != -1} {
            set c2 [string first ")" $val]
            if {$c2 > $c1} {
              set val [string range $val 0 $c1-2][string range $val $c2+1 end]
            } else {
              set val [string range $val 0 $c1-2]
            }
          }
          $cells($spmiSumName) Item $spmiSumRow 3 "'$val"
          set cellval $val

          if {[string first $pmiModifiers(all_over) $val] == 0} {
            addCellComment $spmiSumName $spmiSumRow 3 "The All Over symbol is approximated with two symbols. ($recPracNames(pmi242), Sec. 6.3)"
          }

# allPMI used to count some modifiers for coverage analysis
          if {[string first "tolerance" $thisEntType] != -1} {append allPMI $val}

# check actual vs. expected PMI for NIST files
          if {[info exists nistPMIexpected($epmi)] && $notHole} {nistCheckExpectedPMI $val $entstr $epmi}

# -------------------------------------------------------------------------------
# link back to worksheets
          set anchor [$worksheet($spmiSumName) Range "B$spmiSumRow"]
          set hlsheet $thisEntType
          if {[string length $thisEntType] > 31} {
            foreach item [array names entName] {if {$entName($item) == $thisEntType} {set hlsheet $item}}
          }
          catch {$hlink Add $anchor [string trim ""] "$hlsheet![cellRange $i $pmiCol]" "Go to $thisEntType"}
          incr spmiSumRow
        } else {
          errorMsg "Missing PMI on [formatComplexEnt $thisEntType]"
        }
      }
    }
  }
}

# -------------------------------------------------------------------------------
# start Semantic PMI Coverage analysis worksheet
proc spmiCoverageStart {{multi 1}} {
  global allPMIelements cells cells1 multiFileDir nistName pmiElements pmiElements1 pmiModifiers pmiModifiersRP
  global pmiUnicode sheetLast spmiCoverageWS spmiTypes worksheet worksheet1 worksheets worksheets1

  if {[catch {
    set spmiCoverageWS "Semantic PMI Coverage"

# single file
    if {!$multi} {
      set worksheet($spmiCoverageWS) [$worksheets Add [::tcom::na] $sheetLast]
      $worksheet($spmiCoverageWS) Name $spmiCoverageWS
      set cells($spmiCoverageWS) [$worksheet($spmiCoverageWS) Cells]
      set wsCount [$worksheets Count]
      set i 4
      if {[[$worksheets Item [expr 3]] Name] != "Semantic PMI Summary"} {set i 3}
      [$worksheets Item [expr $wsCount]] -namedarg Move Before [$worksheets Item [expr $i]]

      $cells($spmiCoverageWS) Item 3 1 "PMI Element[format "%c" 10]  (See Help > Analyzer > PMI Coverage Analysis)"
      $cells($spmiCoverageWS) Item 3 2 "Count"
      if {$nistName == ""} {
        addCellComment $spmiCoverageWS 3 2 "See Help > User Guide (section 6.1.7)"
      } else {
        addCellComment $spmiCoverageWS 3 2 "See Help > User Guide (section 6.6.2)"
      }
      set range [$worksheet($spmiCoverageWS) Range "1:3"]
      [$range Font] Bold [expr 1]
      set range [$worksheet($spmiCoverageWS) Range B3]
      $range HorizontalAlignment [expr -4108]

      [$worksheet($spmiCoverageWS) Range A:A] ColumnWidth [expr 48]
      [$worksheet($spmiCoverageWS) Range B:B] ColumnWidth [expr 6]
      [$worksheet($spmiCoverageWS) Range D:D] ColumnWidth [expr 48]

# multiple files
    } else {
      set worksheet1($spmiCoverageWS) [$worksheets1 Item [expr 2]]
      $worksheet1($spmiCoverageWS) Name $spmiCoverageWS
      set cells1($spmiCoverageWS) [$worksheet1($spmiCoverageWS) Cells]
      $cells1($spmiCoverageWS) Item 1 2 "[file nativename $multiFileDir]"
      $cells1($spmiCoverageWS) Item 3 1 "PMI Element[format "%c" 10]  (See Help > Analyzer > PMI Coverage Analysis)"
      set range [$worksheet1($spmiCoverageWS) Range "B1:K1"]
      $range MergeCells [expr 1]
    }

# rows to start adding pmi types
    set row1($spmiCoverageWS) 3
    set row($spmiCoverageWS) 3
    if {!$multi} {set allPMIelements {}}

# add modifiers
    foreach item $spmiTypes {
      set str0 [join $item]
      set str $str0
      if {$str != "controlled_radius"} {
        if {[info exists pmiModifiers($str0)]}   {append str "  $pmiModifiers($str0)"}
        if {[info exists pmiModifiersRP($str0)]} {append str "  ($pmiModifiersRP($str0))"}

# tolerance
        set str1 $str
        set c1 [string last "_" $str]
        if {$c1 != -1} {set str1 [string range $str 0 $c1-1]}
        if {[info exists pmiUnicode($str1)]} {append str "  $pmiUnicode($str1)"}

        if {!$multi} {
          $cells($spmiCoverageWS) Item [incr row($spmiCoverageWS)] 1 $str
          set pmiElements($row($spmiCoverageWS)) $str
          lappend allPMIelements $str
        } else {
          $cells1($spmiCoverageWS) Item [incr row1($spmiCoverageWS)] 1 $str
          set pmiElements1($row1($spmiCoverageWS)) $str
        }
      }
    }
  } emsg3]} {
    errorMsg "Error starting Semantic PMI Coverage worksheet: $emsg3"
  }
}

# -------------------------------------------------------------------------------
# write Semantic PMI Coverage analysis worksheet
proc spmiCoverageWrite {{fn ""} {sum ""} {multi 1}} {
  global allPMI allPMIelements cells cells1 col1 developer entCount epmi fileList nfile nistCoverageStyle nistName
  global pmiModifiers spmiCoverageWS spmiTypesPerFile totalPMI totalPMIrows usedPMIrows worksheet worksheet1
  global objDesign

  if {![info exists allPMIelements] && ![info exists entCount(datum)]} {return}

  if {[catch {
    if {$multi} {
      set range [$worksheet1($spmiCoverageWS) Range [cellRange 3 $col1($sum)] [cellRange 3 $col1($sum)]]
      $range Orientation [expr 90]
      $range HorizontalAlignment [expr -4108]
      $cells1($spmiCoverageWS) Item 3 $col1($sum) $fn
    }
    if {!$multi} {

# check for 'semantic text'
      ::tcom::foreach thisEnt [$objDesign FindObjects [string trim property_definition]] {
        if {[$thisEnt Type] == "property_definition"} {
          ::tcom::foreach attr [$thisEnt Attributes] {
            if {[$attr Name] == "name"} {
              set val [$attr Value]
              if {$val == "semantic text"} {lappend spmiTypesPerFile "editable text"}
            }
          }
        }
      }

# check for with_datum_feature
      foreach ent {dimensional_size_with_datum_feature dimensional_location_with_datum_feature \
                   composite_unit_shape_aspect_and_dimensional_size_with_datum_feature} {
        ::tcom::foreach thisEnt [$objDesign FindObjects [string trim $ent]] {
          if {[$thisEnt Type] == $ent} {lappend spmiTypesPerFile "datum features"}
        }
      }
    }

# check for some modifiers and count from allPMI
    if {[info exists allPMI]} {
      if {[string length $allPMI] > 0} {
        set mods [list maximum_material_requirement least_material_requirement free_state tangent_plane]
        for {set i 0} {$i < [string length $allPMI]} {incr i} {
          foreach mod $mods {
            if {[string index $allPMI $i] == $pmiModifiers($mod)} {incr numMods($mod)}
          }
        }
      }
    }

# add number of pmi types to rows of coverage analysis worksheet
# count number of spmiTypesPerFile, put in stpf
    set stpf {}
    if {[info exists spmiTypesPerFile]} {
      foreach id $spmiTypesPerFile {if {$id != ""} {incr num($id)}}
      foreach id [array names num] {lappend stpf [list $id $num($id)]}
    }

# search all PMI elements with stpf
    foreach item $stpf {
      set idx [lindex $item 0]
      set r [lsearch -glob $allPMIelements $idx*]

# special handling of line, point
      if {$idx == "point"} {
        set r [lsearch $allPMIelements "point  PT  (6.9.7)"]
      } elseif {$idx == "line"} {
        set r [lsearch $allPMIelements "line  SL  (6.9.7)"]
      }

      if {$r != -1} {
        set r [expr {$r+4}]

# add number of pmi to worksheet
        set npmi [lindex $item 1]

# use other count of some modifiers
        if {[info exists numMods]} {
          foreach mod $mods {if {$idx == $mod && $numMods($mod) > 0} {set npmi $numMods($mod)}}
        }

# write npmi
        if {!$multi} {
          $cells($spmiCoverageWS) Item $r 2 $npmi
          set range [$worksheet($spmiCoverageWS) Range [cellRange $r 2] [cellRange $r 2]]
          lappend usedPMIrows $r
        } else {
          $cells1($spmiCoverageWS) Item $r $col1($sum) $npmi
          set range [$worksheet1($spmiCoverageWS) Range [cellRange $r $col1($sum)] [cellRange $r $col1($sum)]]
          incr totalPMI($r) $npmi
        }
        $range HorizontalAlignment [expr -4108]
        if {$multi} {set totalPMIrows($r) 1}
      } elseif {$developer && !$multi && $idx != "curve length" && $idx != "thickness"} {
        errorMsg "  $idx not found in allPMIelements" red
      }
    }
    catch {if {$multi} {unset spmiTypesPerFile}}

# get spmiCoverages (see sfa-gen.tcl to make sure nistReadExpectedPMI is called)
    if {![info exists nfile]} {
      set nf 0
    } else {
      set nf $nfile
    }

# single file, for NIST file color-code coverage
    if {!$multi} {
      if {![info exists epmi]} {set epmi ""}
      if {$nistName != "" || $epmi != ""} {nistPMICoverage $nf}

# multiple files
    } elseif {$nfile == [llength $fileList]} {
      if {[info exists nistCoverageStyle]} {nistAddCoverageStyle}
    }

  } emsg3]} {
    errorMsg "Error adding to Semantic PMI Coverage worksheet: $emsg3"
  }
}

# -------------------------------------------------------------------------------
# format Semantic PMI Coverage analysis worksheet, also PMI totals
proc spmiCoverageFormat {sum {multi 1}} {
  global cells cells1 col1 epmi epmiMulti excel1 lenfilelist localName nistCoverageLegend nistCoverageStyle nistName opt pmiElementsMaxRows
  global pmiHorizontalLineBreaks recPracNames spmiCoverageWS timeStamp totalPMI totalPMIrows usedPMIrows worksheet worksheet1

# delete worksheet if no semantic PMI
  if {$multi && ![info exists totalPMIrows]} {
    catch {$excel1 DisplayAlerts False}
    $worksheet1($spmiCoverageWS) Delete
    catch {$excel1 DisplayAlerts True}
    return
  }

  if {[catch {
    set i1 1

# total PMI for multiple files, totalPMIrows indicates to PMI totals from other columns, totalPMI is the actual total
    if {$multi} {
      set col1($spmiCoverageWS) [expr {$lenfilelist+2}]
      $cells1($spmiCoverageWS) Item 3 $col1($spmiCoverageWS) "Total PMI"
      set range [$worksheet1($spmiCoverageWS) Range [cellRange 3 $col1($spmiCoverageWS)]]
      [$range Font] Bold [expr 1]
      foreach idx [array names totalPMIrows] {
        if {![info exists totalPMI($idx)]} {set totalPMI($idx) 0}
        $cells1($spmiCoverageWS) Item $idx $col1($spmiCoverageWS) $totalPMI($idx)
      }
      catch {unset totalPMIrows}
      $worksheet1($spmiCoverageWS) Activate
    }

# horizontal break lines, depends on items in semantic PMI coverage worksheet, items defined in sfa-data
    set idx1 $pmiHorizontalLineBreaks
    if {!$multi} {set idx1 [concat [list 3 4] $idx1]}
    for {set r $pmiElementsMaxRows} {$r >= [lindex $idx1 end]} {incr r -1} {
      if {!$multi} {
        set val [[$cells($spmiCoverageWS) Item $r 1] Value]
      } else {
        set val [[$cells1($spmiCoverageWS) Item $r 1] Value]
      }
      if {$val != ""} {
        lappend idx1 [expr {$r+1}]
        break
      }
    }

# horizontal lines
    foreach idx $idx1 {
      if {!$multi} {
        set range [$worksheet($spmiCoverageWS) Range [cellRange $idx 1] [cellRange $idx 2]]
      } else {
        set range [$worksheet1($spmiCoverageWS) Range [cellRange $idx 1] [cellRange $idx [expr {$col1($spmiCoverageWS)+$i1-1}]]]
      }
      catch {[[$range Borders] Item [expr 8]] Weight [expr 2]}
    }

# multi file
    if {$multi} {

# vertical line(s), also in sfa-multi.tcl vertical lines when changing directory
      for {set i 0} {$i < $i1} {incr i} {
        set range [$worksheet1($spmiCoverageWS) Range [cellRange 1 [expr {$col1($spmiCoverageWS)+$i}]] [cellRange [expr {[lindex $idx1 end]-1}] [expr {$col1($spmiCoverageWS)+$i}]]]
        catch {[[$range Borders] Item [expr 7]] Weight [expr 2]}
      }

# fix row 3 height and width
      set range [$worksheet1($spmiCoverageWS) Range 3:3]
      $range RowHeight 300
      [$worksheet1($spmiCoverageWS) Columns] AutoFit

# delete unused rows, check for a value in the Total PMI column (multi file)
      if {!$opt(SHOWALLPMI)} {
        set lineBreak 0
        for {set i $pmiElementsMaxRows} {$i > 3} {incr i -1} {
          if {[lsearch [array names totalPMI] $i] == -1} {
            set range [$worksheet1($spmiCoverageWS) Range A$i]
            [$range EntireRow] Delete [expr -4162]
            if {[lsearch $pmiHorizontalLineBreaks $i] != -1} {set lineBreak 1}
          } elseif {$lineBreak} {
            set range [$worksheet1($spmiCoverageWS) Range [cellRange $i 1] [cellRange $i [expr {$col1($spmiCoverageWS)+$i1-1}]]]
            catch {[[$range Borders] Item [expr 9]] Weight [expr 2]}
            set lineBreak 0
          }
        }
      }

# add color legend for NIST files
      if {[info exists nistCoverageStyle]} {nistAddCoverageLegend $multi}

# final formatting (multi file)
      set range [$worksheet1($spmiCoverageWS) Range A3]
      [$range Font] Bold [expr 1]
      foreach i {0 1} {
        $range WrapText [expr $i]
        [$worksheet1($spmiCoverageWS) Columns] AutoFit
      }
      set range [$worksheet1($spmiCoverageWS) Range B1]
      [$range Font] Bold [expr 1]

      set r2 [expr {[[[$worksheet1($spmiCoverageWS) UsedRange] Rows] Count]+1}]
      if {$nistName != "" || [info exists epmiMulti]} {set r2 [expr {[[[$worksheet1($spmiCoverageWS) UsedRange] Rows] Count]-8}]}
      catch {unset epmiMulti}
      $cells1($spmiCoverageWS) Item $r2 1 "Section numbers above refer to the CAx-IF Recommended Practice for $recPracNames(pmi242)"
      set anchor [$worksheet1($spmiCoverageWS) Range [cellRange $r2 1]]
      [$worksheet1($spmiCoverageWS) Hyperlinks] Add $anchor [join "https://www.mbx-if.org/home/cax/recpractices/"] [join ""] [join "Link to CAx-IF Recommended Practices"]

      [$worksheet1($spmiCoverageWS) Rows] AutoFit
      [$worksheet1($spmiCoverageWS) Range "B4"] Select
      catch {[$excel1 ActiveWindow] FreezePanes [expr 1]}
      [$worksheet1($spmiCoverageWS) Range "A1"] Select

# single file
    } else {
      set i1 3
      for {set i 0} {$i < $i1} {incr i} {
        set range [$worksheet($spmiCoverageWS) Range [cellRange 3 [expr {$i+1}]] [cellRange [expr {[lindex $idx1 end]-1}] [expr {$i+1}]]]
        catch {[[$range Borders] Item [expr 7]] Weight [expr 2]}
      }

# delete unused rows and retain horizontal lines breaks (single file)
      if {!$opt(SHOWALLPMI)} {
        set lineBreak 0
        for {set i $pmiElementsMaxRows} {$i > 3} {incr i -1} {
          if {[lsearch $usedPMIrows $i] == -1} {
            set range [$worksheet($spmiCoverageWS) Range A$i]
            [$range EntireRow] Delete [expr -4162]
            if {[lsearch $pmiHorizontalLineBreaks $i] != -1} {set lineBreak 1}
          } elseif {$lineBreak} {
            set range [$worksheet($spmiCoverageWS) Range [cellRange $i 1] [cellRange $i 2]]
            catch {[[$range Borders] Item [expr 9]] Weight [expr 2]}
            set lineBreak 0
          }
        }
        unset usedPMIrows
      }

# add color legend for NIST files
      if {$nistCoverageLegend} {nistAddCoverageLegend}

# final formatting (single file)
      [$worksheet($spmiCoverageWS) Columns] AutoFit
      [$worksheet($spmiCoverageWS) Rows] AutoFit
      set r2 [expr {[[[$worksheet($spmiCoverageWS) UsedRange] Rows] Count]+1}]
      if {![info exists epmi]} {
        set epmi ""
      } elseif {$epmi != ""} {
        set epmiMulti 1
      }
      if {$nistName != "" || $epmi != ""} {set r2 [expr {[[[$worksheet($spmiCoverageWS) UsedRange] Rows] Count]-8}]}
      $cells($spmiCoverageWS) Item $r2 1 "Section numbers above refer to the CAx-IF Recommended Practice for $recPracNames(pmi242)"
      set anchor [$worksheet($spmiCoverageWS) Range [cellRange $r2 1]]
      [$worksheet($spmiCoverageWS) Hyperlinks] Add $anchor [join "https://www.mbx-if.org/home/cax/recpractices/"] [join ""] [join "Link to CAx-IF Recommended Practices"]

      [$worksheet($spmiCoverageWS) Range "A1"] Select
      set txt [file tail $localName]
      if {$timeStamp != ""} {append txt "[format "%c" 10]  ($timeStamp)"}
      $cells($spmiCoverageWS) Item 1 1 $txt
    }

# errors
  } emsg]} {
    errorMsg "Error formatting Semantic PMI Coverage worksheet: $emsg"
  }
}

# -------------------------------------------------------------------------------
# start Graphic PMI Coverage analysis worksheet
proc gpmiCoverageStart {{multi 1}} {
  global cells cells1 gpmiCoverageWS gpmiTypes multiFileDir opt recPracNames sheetLast worksheet worksheet1 worksheets worksheets1

  if {[catch {
    set gpmiCoverageWS "Graphic PMI Coverage"

# rec prac message
    set msgrp "Recommended Graphic PMI names defined in $recPracNames(pmi242), Sec. 8.4"

# multiple files
    if {$multi} {
      if {$opt(PMISEM)} {
        set worksheet1($gpmiCoverageWS) [$worksheets1 Item [expr 3]]
      } else {
        set worksheet1($gpmiCoverageWS) [$worksheets1 Item [expr 2]]
      }
      $worksheet1($gpmiCoverageWS) Name $gpmiCoverageWS
      set cells1($gpmiCoverageWS) [$worksheet1($gpmiCoverageWS) Cells]
      $cells1($gpmiCoverageWS) Item 1 2 "[file nativename $multiFileDir]"
      $cells1($gpmiCoverageWS) Item 3 1 "Graphic PMI Names"
      set range [$worksheet1($gpmiCoverageWS) Range "B1:K1"]
      $range MergeCells [expr 1]
      set row1($gpmiCoverageWS) 3

# single file
    } else {
      set spmiCoverageWS "Graphic PMI Coverage"
      set n 3
      if {[info exists worksheet($spmiCoverageWS)]} {
        set n 5
      }
      set worksheet($gpmiCoverageWS) [$worksheets Add [::tcom::na] $sheetLast]
      $worksheet($gpmiCoverageWS) Name $gpmiCoverageWS
      set cells($gpmiCoverageWS) [$worksheet($gpmiCoverageWS) Cells]
      set wsCount [$worksheets Count]
      [$worksheets Item [expr $wsCount]] -namedarg Move Before [$worksheets Item [expr $n]]
      $cells($gpmiCoverageWS) Item 3 1 "Graphic PMI Names"
      $cells($gpmiCoverageWS) Item 3 2 "Count"
      addCellComment $gpmiCoverageWS 3 1 $msgrp
      set range [$worksheet($gpmiCoverageWS) Range "1:3"]
      [$range Font] Bold [expr 1]
      set row($gpmiCoverageWS) 3
    }

    foreach item $gpmiTypes {
      set str [join $item]
      if {$multi} {
        $cells1($gpmiCoverageWS) Item [incr row1($gpmiCoverageWS)] 1 $str
      } else {
        $cells($gpmiCoverageWS) Item [incr row($gpmiCoverageWS)] 1 $str
      }
    }
  } emsg3]} {
    errorMsg "Error starting Graphic PMI Coverage worksheet: $emsg3"
  }
}

# -------------------------------------------------------------------------------
# write Graphic PMI Coverage analysis worksheet
proc gpmiCoverageWrite {{fn ""} {sum ""} {multi 1}} {
  global cells cells1 col1 gpmiCoverageWS gpmiRows gpmiTotals gpmiTypes gpmiTypesInvalid gpmiTypesPerFile legendColor worksheet worksheet1

  if {[catch {
    if {$multi} {
      set range [$worksheet1($gpmiCoverageWS) Range [cellRange 3 $col1($sum)] [cellRange 3 $col1($sum)]]
      $range Orientation [expr 90]
      $range HorizontalAlignment [expr -4108]
      $cells1($gpmiCoverageWS) Item 3 $col1($sum) $fn
    }

# add invalid pmi types to column A
# need to fix when there are invalid types, but a subsequent file does not if processing multiple files
    set r1 [expr {[llength $gpmiTypes]+4}]
    if {![info exists gpmiRows]} {set gpmiRows 35}
    set ok 1

    if {[info exists gpmiTypesInvalid]} {
      while {$ok} {
        if {$multi} {
          set val [[$cells1($gpmiCoverageWS) Item $r1 1] Value]
        } else {
          set val [[$cells($gpmiCoverageWS) Item $r1 1] Value]
        }
        if {$val == ""} {
          foreach idx $gpmiTypesInvalid {
            if {$multi} {
              $cells1($gpmiCoverageWS) Item $r1 1 $idx
              [[$worksheet1($gpmiCoverageWS) Range [cellRange $r1 1] [cellRange $r1 1]] Interior] Color $legendColor(red)
            } else {
              $cells($gpmiCoverageWS) Item $r1 1 $idx
              [[$worksheet($gpmiCoverageWS) Range [cellRange $r1 1] [cellRange $r1 1]] Interior] Color $legendColor(red)
            }
            if {$r1 > $gpmiRows} {set gpmiRows $r1}
            incr r1
          }
          set ok 0
        } else {
          foreach idx $gpmiTypesInvalid {
            if {$idx != $val} {
              incr r1
              if {$multi} {
                $cells1($gpmiCoverageWS) Item $r1 1 $idx
                [[$worksheet1($gpmiCoverageWS) Range [cellRange $r1 1] [cellRange $r1 1]] Interior] Color $legendColor(red)
              } else {
                $cells($gpmiCoverageWS) Item $r1 1 $idx
                [[$worksheet($gpmiCoverageWS) Range [cellRange $r1 1] [cellRange $r1 1]] Interior] Color $legendColor(red)
              }
              set val $idx
              if {$r1 > $gpmiRows} {set gpmiRows $r1}
            }
          }
          set ok 0
        }
      }
    }

# add numbers
    if {[info exists gpmiTypesPerFile]} {
      set gpmiTypesPerFile [lrmdups $gpmiTypesPerFile]
      for {set r 4} {$r <= $gpmiRows} {incr r} {
        if {$multi} {
          set val [[$cells1($gpmiCoverageWS) Item $r 1] Value]
        } else {
          set val [[$cells($gpmiCoverageWS) Item $r 1] Value]
        }
        foreach item $gpmiTypesPerFile {
          set idx [lindex [split $item "/"] 0]
          if {$val == $idx} {

# get current value
            if {$multi} {
              set npmi [[$cells1($gpmiCoverageWS) Item $r $col1($sum)] Value]
            } else {
              set npmi [[$cells($gpmiCoverageWS) Item $r 2] Value]
            }

# set or increment npmi
            if {$npmi == ""} {
              set npmi 1
            } else {
              set npmi [expr {int($npmi)+1}]
            }

# write npmi
            if {$multi} {
              $cells1($gpmiCoverageWS) Item $r $col1($sum) $npmi
              set range [$worksheet1($gpmiCoverageWS) Range [cellRange $r $col1($sum)] [cellRange $r $col1($sum)]]
              incr gpmiTotals($r)
            } else {
              $cells($gpmiCoverageWS) Item $r 2 $npmi
              set range [$worksheet($gpmiCoverageWS) Range [cellRange $r 2] [cellRange $r 2]]
            }
            $range HorizontalAlignment [expr -4108]
          }
        }
      }
      catch {if {$multi} {unset gpmiTypesPerFile}}
    }
  } emsg3]} {
    errorMsg "Error adding to Graphic PMI Coverage worksheet: $emsg3"
  }
}

# -------------------------------------------------------------------------------
# format Graphic PMI Coverage analysis worksheet, also PMI totals
proc gpmiCoverageFormat {{sum ""} {multi 1}} {
  global cells cells1 col1 excel excel1 lenfilelist localName nistName
  global gpmiCoverageWS gpmiRows gpmiTotals timeStamp worksheet worksheet1

# delete worksheet if no graphic PMI
  if {$multi && ![info exists gpmiTotals]} {
    catch {$excel1 DisplayAlerts False}
    $worksheet1($gpmiCoverageWS) Delete
    catch {$excel1 DisplayAlerts True}
    return
  }

# total PMI
  if {[catch {
    if {$multi} {
      set col1($gpmiCoverageWS) [expr {$lenfilelist+2}]
      $cells1($gpmiCoverageWS) Item 3 $col1($gpmiCoverageWS) "Total PMI"
      set range [$worksheet1($gpmiCoverageWS) Range [cellRange 3 $col1($gpmiCoverageWS)]]
      [$range Font] Bold [expr 1]
      foreach idx [array names gpmiTotals] {
        $cells1($gpmiCoverageWS) Item $idx $col1($gpmiCoverageWS) $gpmiTotals($idx)
      }
      $worksheet1($gpmiCoverageWS) Activate
    }

# horizontal break lines
    set idx1 [list 21 28 30 35]
    if {!$multi} {set idx1 [list 3 4 21 28 30 35]}
    for {set r 100} {$r >= 35} {incr r -1} {
      if {$multi} {
        set val [[$cells1($gpmiCoverageWS) Item $r 1] Value]
      } else {
        set val [[$cells($gpmiCoverageWS) Item $r 1] Value]
      }
      if {$val != ""} {
        lappend idx1 [expr {$r+1}]
        break
      }
    }

# horizontal lines
    foreach idx $idx1 {
      if {$multi} {
        set range [$worksheet1($gpmiCoverageWS) Range [cellRange $idx 1] [cellRange $idx $col1($gpmiCoverageWS)]]
      } else {
        set range [$worksheet($gpmiCoverageWS) Range [cellRange $idx 1] [cellRange $idx 2]]
      }
      catch {[[$range Borders] Item [expr 8]] Weight [expr 2]}
    }

# vertical line(s)
    if {$multi} {
      set range [$worksheet1($gpmiCoverageWS) Range [cellRange 1 $col1($gpmiCoverageWS)] [cellRange [expr {[lindex $idx1 end]-1}] $col1($gpmiCoverageWS)]]
      catch {[[$range Borders] Item [expr 7]] Weight [expr 2]}

# bold
      set range [$worksheet1($gpmiCoverageWS) Range A3]
      [$range Font] Bold [expr 1]
      set range [$worksheet1($gpmiCoverageWS) Range B1]
      [$range Font] Bold [expr 1]

# fix row 3 height and width
      set range [$worksheet1($gpmiCoverageWS) Range 3:3]
      $range RowHeight 300
      [$worksheet1($gpmiCoverageWS) Columns] AutoFit
      [$worksheet1($gpmiCoverageWS) Rows] AutoFit
      [$worksheet1($gpmiCoverageWS) Range "B4"] Select
      catch {[$excel1 ActiveWindow] FreezePanes [expr 1]}
      [$worksheet1($gpmiCoverageWS) Range "A1"] Select

# single file
    } else {
      set i1 3
      for {set i 0} {$i < $i1} {incr i} {
        set range [$worksheet($gpmiCoverageWS) Range [cellRange 3 [expr {$i+1}]] [cellRange [expr {[lindex $idx1 end]-1}] [expr {$i+1}]]]
        catch {[[$range Borders] Item [expr 7]] Weight [expr 2]}
      }
      [$worksheet($gpmiCoverageWS) Columns] AutoFit

      [$worksheet($gpmiCoverageWS) Range "A1"] Select
      set txt [file tail $localName]
      if {$timeStamp != ""} {append txt "[format "%c" 10]  ($timeStamp)"}
      $cells($gpmiCoverageWS) Item 1 1 $txt
      $cells($gpmiCoverageWS) Item [expr {$gpmiRows+3}] 1 "See Help > Analyzer > PMI Coverage Analysis"

# add images for the CAx-IF and NIST PMI models
      if {$nistName != ""} {nistAddModelPictures $gpmiCoverageWS}
    }

# errors
  } emsg]} {
    errorMsg "Error formatting Graphic PMI Coverage worksheet: $emsg"
  }
}
