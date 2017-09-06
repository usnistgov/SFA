# start semantic PMI coverage analysis worksheet
proc spmiCoverageStart {{multi 1}} {
  global cells cells1 multiFileDir pmiModifiers pmiModifiersRP pmiUnicode
  global sempmi_coverage sheetLast spmiTypes worksheet worksheet1 worksheets worksheets1 
  #outputMsg "spmiCoverageStart $multi" red

  if {[catch {
    set sempmi_coverage "PMI Representation Coverage"

# multiple files
    if {$multi} {
      set worksheet1($sempmi_coverage) [$worksheets1 Item [expr 2]]
      #$worksheet1($sempmi_coverage) Activate
      $worksheet1($sempmi_coverage) Name $sempmi_coverage
      set cells1($sempmi_coverage) [$worksheet1($sempmi_coverage) Cells]
      $cells1($sempmi_coverage) Item 1 1 "STEP Directory"
      $cells1($sempmi_coverage) Item 1 2 "[file nativename $multiFileDir]"
      $cells1($sempmi_coverage) Item 3 1 "PMI Element   (See Help > PMI Coverage Analysis)"
      set range [$worksheet1($sempmi_coverage) Range "B1:K1"]
      [$range Font] Bold [expr 1]
      $range MergeCells [expr 1]

# single file
    } else {
      set worksheet($sempmi_coverage) [$worksheets Add [::tcom::na] $sheetLast]
      #$worksheet($sempmi_coverage) Activate
      $worksheet($sempmi_coverage) Name $sempmi_coverage
      set cells($sempmi_coverage) [$worksheet($sempmi_coverage) Cells]
      set wsCount [$worksheets Count]
      [$worksheets Item [expr $wsCount]] -namedarg Move Before [$worksheets Item [expr 4]]

      $cells($sempmi_coverage) Item 3 1 "PMI Element (See Help > PMI Coverage Analysis)"
      $cells($sempmi_coverage) Item 3 2 "Count"
      set range [$worksheet($sempmi_coverage) Range "1:3"]
      [$range Font] Bold [expr 1]

      [$worksheet($sempmi_coverage) Range A:A] ColumnWidth [expr 48]
      [$worksheet($sempmi_coverage) Range B:B] ColumnWidth [expr 6]
      [$worksheet($sempmi_coverage) Range D:D] ColumnWidth [expr 48]
    }
    
# add pmi types
    set row1($sempmi_coverage) 3
    set row($sempmi_coverage) 3

# add modifiers
    foreach item $spmiTypes {
      set str0 [join $item]
      set str $str0
      if {$str != "square" && $str != "controlled_radius"} {
        if {[info exists pmiModifiers($str0)]}   {append str "  $pmiModifiers($str0)"}
        if {[info exists pmiModifiersRP($str0)]} {append str "  ($pmiModifiersRP($str0))"}

# tolerance
        set str1 $str
        set c1 [string last "_" $str]
        if {$c1 != -1} {set str1 [string range $str 0 $c1-1]}
        if {[info exists pmiUnicode($str1)]} {append str "  $pmiUnicode($str1)"}

        if {!$multi} {
          $cells($sempmi_coverage) Item [incr row($sempmi_coverage)] 1 $str
        } else {
          $cells1($sempmi_coverage) Item [incr row1($sempmi_coverage)] 1 $str
        }
      }
      #outputMsg $str
    }
  } emsg3]} {
    errorMsg "ERROR starting PMI Representation Coverage worksheet: $emsg3"
  }
}

# -------------------------------------------------------------------------------
# write semantic PMI coverage analysis worksheet
proc spmiCoverageWrite {{fn ""} {sum ""} {multi 1}} {
  global cells cells1 col1 coverageLegend coverageStyle entCount fileList legendColor nfile nistName 
  global sempmi_coverage sempmi_totals spmiCoverages spmiTypes spmiTypesPerFile checkPMImods worksheet worksheet1 allPMI pmiModifiers
  #outputMsg "spmiCoverageWrite $multi" red

  if {[catch {
    if {$multi} {
      set range [$worksheet1($sempmi_coverage) Range [cellRange 3 $col1($sum)] [cellRange 3 $col1($sum)]]
      $range Orientation [expr 90]
      $range HorizontalAlignment [expr -4108]
      $cells1($sempmi_coverage) Item 3 $col1($sum) $fn
    }
    
    if {[info exists entCount(datum)]} {
      for {set i 0} {$i < $entCount(datum)} {incr i} {
        lappend spmiTypesPerFile1 "datum (6.5)"
      }
      if {$multi} {unset entCount(datum)}
    }
    
# check for some modifiers and count from allPMI
    if {[info exists allPMI]} {
      if {[string length $allPMI] > 0} {
        set mods [list maximum_material_requirement least_material_requirement free_state tangent_plane]
        foreach mod $mods {set numMods($mod) 0}
        for {set i 0} {$i < [string length $allPMI]} {incr i} {
          foreach mod $mods {
            if {[string index $allPMI $i] == $pmiModifiers($mod)} {incr numMods($mod)}
          }
        }
      }
    }

# add number of pmi types
    if {[info exists spmiTypesPerFile] || [info exists spmiTypesPerFile1]} {
      for {set r 4} {$r <= 130} {incr r} {
        if {$multi} {
          set val [[$cells1($sempmi_coverage) Item $r 1] Value]
        } else {
          set val [[$cells($sempmi_coverage) Item $r 1] Value]
        }

        if {[info exists spmiTypesPerFile]} {
          foreach idx $spmiTypesPerFile {
            set ok 0
            if {$idx != "line" && $idx != "point" && $idx != "free_state"} {
              if {([string first $idx $val] == 0 && [string first "statistical_tolerance" $val] == -1) || \
                  $idx == [lindex [split $val " "] 0]} {set ok 1}
            } else {
              if {[string first "$idx  " $val] == 0} {set ok 1}
            }
            if {$ok} {

# get current value
              if {$multi} {
                set npmi [[$cells1($sempmi_coverage) Item $r $col1($sum)] Value]
              } else {
                set npmi [[$cells($sempmi_coverage) Item $r 2] Value]
              }

# set or increment npmi
              if {$npmi == ""} {
                set npmi 1
              } else {
                set npmi [expr {int($npmi)+1}]
              }

# use other count of some modifiers
              if {[info exists allPMI]} {
                if {[string length $allPMI] > 0} {foreach mod $mods {if {$idx == $mod} {set npmi $numMods($mod)}}}
              }
              
# write npmi
              if {$multi} {
                $cells1($sempmi_coverage) Item $r $col1($sum) $npmi
                set range [$worksheet1($sempmi_coverage) Range [cellRange $r $col1($sum)] [cellRange $r $col1($sum)]]
              } else {
                $cells($sempmi_coverage) Item $r 2 $npmi
                set range [$worksheet($sempmi_coverage) Range [cellRange $r 2] [cellRange $r 2]]
              }
              $range HorizontalAlignment [expr -4108]
              if {$multi} {incr sempmi_totals($r)}
            }
          }
        }

# exact match (only datum)
        if {[info exists spmiTypesPerFile1]} {
          foreach idx $spmiTypesPerFile1 {
            if {$idx == $val} {
              if {$multi} {
                set npmi [[$cells1($sempmi_coverage) Item $r $col1($sum)] Value]
              } else {
                set npmi [[$cells($sempmi_coverage) Item $r 2] Value]
              }
              if {$npmi == ""} {
                set npmi 1
              } else {
                set npmi [expr {int($npmi)+1}]
              }
              if {$multi} {
                $cells1($sempmi_coverage) Item $r $col1($sum) $npmi
                set range [$worksheet1($sempmi_coverage) Range [cellRange $r $col1($sum)] [cellRange $r $col1($sum)]]
              } else {
                $cells($sempmi_coverage) Item $r 2 $npmi
                set range [$worksheet($sempmi_coverage) Range [cellRange $r 2] [cellRange $r 2]]
              }
              $range HorizontalAlignment [expr -4108]
              if {$multi} {incr sempmi_totals($r)}
            }
          }          
        }
      }
      catch {if {$multi} {unset spmiTypesPerFile}}
    }

# get spmiCoverages (see sfa-gen.tcl to make sure spmiGetPMI is called)
    if {![info exists nfile]} {
      set nf 0
    } else {
      set nf $nfile
    }
    if {!$multi} {
      foreach idx [lsort [array names spmiCoverages]] {
        set tval [lindex [split $idx ","] 0]
        set fnam [lindex [split $idx ","] 1]
        if {$fnam == $nistName} {
          set coverage($tval) $spmiCoverages($idx)
        }
      }
      #foreach item [lsort [array names spmiCoverages]] {if {$spmiCoverages($item) != ""} {outputMsg "$item $spmiCoverages($item)" green}}
      #foreach item [lsort [array names coverage]] {if {$coverage($item) != ""} {outputMsg "$item $coverage($item)" red}}
    
# check values for color-coding
      for {set r 4} {$r <= [[[$worksheet($sempmi_coverage) UsedRange] Rows] Count]} {incr r} {
        set ttyp [[$cells($sempmi_coverage) Item $r 1] Value]
        set tval [[$cells($sempmi_coverage) Item $r 2] Value]
        if {$ttyp != ""} {
          if {$tval == ""} {set tval 0}
          set tval [expr {int($tval)}]
          #outputMsg "$r  $tval  $ttyp" red

          foreach item [array names coverage] {
            if {[string first $item $ttyp] == 0} {
              set ok 0

# these words appear in other PMI elements and need to be handled separately, e.g. statistical is also in statistical_tolerance
              if {$item != "datum" && $item != "line" && $item != "spherical" && \
                  $item != "statistical" && $item != "basic" && $item != "point"} {
                set ok 1

# special cases
              } else {
                set str [string range $ttyp 0 [expr {[string last " " $ttyp]-1}]]
                if {$item == $str} {set ok 1}
                if {!$ok} {
                  set str [string range $ttyp 0 [expr {[string last " " $ttyp]-2}]]
                  if {$item == $str} {set ok 1}
                }
                if {!$ok} {
                  set str [string range $ttyp 0 [expr {[string first "<" $ttyp]-3}]]
                  if {$item == $str} {set ok 1}
                }
              }

# need better fix for free_state_condition conflict with free_state
              if {[string first "free_state_condition" $ttyp] != -1} {set ok 0}
              
              if {$ok} {
                set ci $coverage($item)
                catch {set ci [expr {int($ci)}]}
                #outputMsg " $item / $tval / $coverage($item) / $ci" red
# neutral - grey         
                if {$coverage($item) != "" && $ci < 0} {
                  [[$worksheet($sempmi_coverage) Range B$r] Interior] Color $legendColor(gray)
                  set coverageLegend 1
                  lappend coverageStyle "$r $nf gray"

# too few - yellow or red (was red or magenta)
                } elseif {$tval < $ci} {
                  set str "'$tval/$ci"
                  $cells($sempmi_coverage) Item $r 2 $str
                  [$worksheet($sempmi_coverage) Range B$r] HorizontalAlignment [expr -4108]
                  set coverageLegend 1
                  if {$tval == 0} {
                    set clr "red"
                  } else {
                    set clr "yellow"
                  }
                  [[$worksheet($sempmi_coverage) Range B$r] Interior] Color $legendColor($clr)
                  lappend coverageStyle "$r $nf $clr $str"

# too many - cyan or magenta (was yellow)
                } elseif {$tval > $ci && $tval != 0} {
                  set ci1 $coverage($item)
                  set clr "cyan"
                  if {$ci1 == ""} {
                    set ci1 0
                    set clr "magenta"
                  }
                  set str "'$tval/[expr {int($ci1)}]"
                  $cells($sempmi_coverage) Item $r 2 $str
                  [[$worksheet($sempmi_coverage) Range B$r] Interior] Color $legendColor($clr)
                  [$worksheet($sempmi_coverage) Range B$r] NumberFormat "@"
                  set coverageLegend 1
                  lappend coverageStyle "$r $nf $clr $str"

# just right - green
                } elseif {$tval != 0} {
                  [[$worksheet($sempmi_coverage) Range B$r] Interior] Color $legendColor(green)
                  set coverageLegend 1
                  lappend coverageStyle "$r $nf green"
                }
              }
            }
          }
        }
      }

# multiple files
    } elseif {$nfile == [llength $fileList]} {
      if {[info exists coverageStyle]} {
        foreach item $coverageStyle {
          set r [lindex [split $item " "] 0]
          set c [expr {[lindex [split $item " "] 1]+1}]
          set style [lindex [split $item " "] 2]
          if {[llength $item] > 3} {
            set str [lindex [split $item " "] 3]
            $cells1($sempmi_coverage) Item $r $c $str
            [$worksheet1($sempmi_coverage) Range [cellRange $r $c]] HorizontalAlignment [expr -4108]
          }
          #outputMsg "$r $c $style" green
          [[$worksheet1($sempmi_coverage) Range [cellRange $r $c]] Interior] Color $legendColor($style)
        }
      }
    }
  } emsg3]} {
    errorMsg "ERROR adding to PMI Representation Coverage worksheet: $emsg3"
  }
}

# -------------------------------------------------------------------------------
# format semantic PMI coverage analysis worksheet, also PMI totals
proc spmiCoverageFormat {sum {multi 1}} {
  global cells cells1 col1 coverageLegend coverageStyle excel1 lenfilelist localName opt excelVersion
  global pmiModifiers pmiUnicode recPracNames sempmi_coverage sempmi_totals spmiTypes worksheet worksheet1 

  #outputMsg "spmiCoverageFormat $multi" red

# delete worksheet if no semantic PMI
  if {$multi && ![info exists sempmi_totals]} {
    catch {$excel1 DisplayAlerts False}
    $worksheet1($sempmi_coverage) Delete
    catch {$excel1 DisplayAlerts True}
    return
  }

# total PMI, multiple files
  if {[catch {
    set i1 1
    if {$multi} {
      set col1($sempmi_coverage) [expr {$lenfilelist+2}]
      $cells1($sempmi_coverage) Item 3 $col1($sempmi_coverage) "Total PMI"
      foreach idx [array names sempmi_totals] {
        $cells1($sempmi_coverage) Item $idx $col1($sempmi_coverage) $sempmi_totals($idx)
      }
      catch {unset sempmi_totals}
    
# pmi names on right, if necessary
      if {$col1($sempmi_coverage) > 20} {
        set r 3
        foreach item $spmiTypes {
          set str0 [join $item]
          set str $str0
          if {$str != "square" && $str != "controlled_radius"} {
            if {[info exists pmiModifiers($str0)]}   {append str "  $pmiModifiers($str0)"}
            set str1 $str
            set c1 [string last "_" $str]
            if {$c1 != -1} {set str1 [string range $str 0 $c1-1]}
            if {[info exists pmiUnicode($str1)]} {append str "  $pmiUnicode($str1)"}
            $cells1($sempmi_coverage) Item [incr r] [expr {$col1($sempmi_coverage)+1}] $str
          }
        }
        set i1 2
      }
      $worksheet1($sempmi_coverage) Activate
    }
 
# horizontal break lines
    set idx1 [list 20 42 55 61 80]
    if {!$multi} {set idx1 [list 3 4 20 42 55 61 80]}
    for {set r 200} {$r >= [lindex $idx1 end]} {incr r -1} {
      if {$multi} {
        set val [[$cells1($sempmi_coverage) Item $r 1] Value]
      } else {
        set val [[$cells($sempmi_coverage) Item $r 1] Value]
      }
      if {$val != ""} {
        lappend idx1 [expr {$r+1}]
        break
      }
    }    

# horizontal lines
    foreach idx $idx1 {
      if {$multi} {
        set range [$worksheet1($sempmi_coverage) Range [cellRange $idx 1] [cellRange $idx [expr {$col1($sempmi_coverage)+$i1-1}]]]
      } else {
        set range [$worksheet($sempmi_coverage) Range [cellRange $idx 1] [cellRange $idx 2]]
      }
      catch {[[$range Borders] Item [expr 8]] Weight [expr 2]}
    }

# vertical line(s)
    if {$multi} {
      for {set i 0} {$i < $i1} {incr i} {
        set range [$worksheet1($sempmi_coverage) Range [cellRange 1 [expr {$col1($sempmi_coverage)+$i}]] [cellRange [expr {[lindex $idx1 end]-1}] [expr {$col1($sempmi_coverage)+$i}]]]
        catch {[[$range Borders] Item [expr 7]] Weight [expr 2]}
      }
      
# fix row 3 height and width
      set range [$worksheet1($sempmi_coverage) Range 3:3]
      $range RowHeight 300
      [$worksheet1($sempmi_coverage) Columns] AutoFit

      $cells1($sempmi_coverage) Item [expr {[lindex $idx1 end]+1}] 1 "Section Numbers refer to the CAx-IF Recommended Practice for $recPracNames(pmi242)"
      set anchor [$worksheet1($sempmi_coverage) Range [cellRange [expr {[lindex $idx1 end]+1}] 1]]
      [$worksheet1($sempmi_coverage) Hyperlinks] Add $anchor [join "http://www.cax-if.org/joint_testing_info.html#recpracs"] [join ""] [join "Link to CAx-IF Recommended Practices"]
      
      if {[info exists coverageStyle]} {spmiCoverageLegend $multi [expr {[lindex $idx1 end]+3}]}
      
      [$worksheet1($sempmi_coverage) Rows] AutoFit
      [$worksheet1($sempmi_coverage) Range "B4"] Select
      catch {[$excel1 ActiveWindow] FreezePanes [expr 1]}
      [$worksheet1($sempmi_coverage) Range "A1"] Select
      catch {[$worksheet1($sempmi_coverage) PageSetup] PrintGridlines [expr 1]}

# single file
    } else {
      set i1 3
      for {set i 0} {$i < $i1} {incr i} {
        set range [$worksheet($sempmi_coverage) Range [cellRange 3 [expr {$i+1}]] [cellRange [expr {[lindex $idx1 end]-1}] [expr {$i+1}]]]
        catch {[[$range Borders] Item [expr 7]] Weight [expr 2]}
      }
      
      if {$coverageLegend} {spmiCoverageLegend $multi}
      [$worksheet($sempmi_coverage) Columns] AutoFit

      $cells($sempmi_coverage) Item 1 4 "Section Numbers refer to the CAx-IF Recommended Practice for $recPracNames(pmi242)"
      set range [$worksheet($sempmi_coverage) Range D1:N1]
      $range MergeCells [expr 1]
      set anchor [$worksheet($sempmi_coverage) Range D1]
      [$worksheet($sempmi_coverage) Hyperlinks] Add $anchor [join "http://www.cax-if.org/joint_testing_info.html#recpracs"] [join ""] [join "Link to CAx-IF Recommended Practices"]

      [$worksheet($sempmi_coverage) Range "A1"] Select
      catch {[$worksheet($sempmi_coverage) PageSetup] PrintGridlines [expr 1]}
      $cells($sempmi_coverage) Item 1 1 [file tail $localName]
    }

# errors
  } emsg]} {
    errorMsg "ERROR formatting PMI Representation Coverage worksheet: $emsg"
  }
}

# -------------------------------------------------------------------------------
# add coverage legend
proc spmiCoverageLegend {multi {row 3}} {
  global cells cells1 excel excel1 legendColor sempmi_coverage worksheet worksheet1
  
  if {$multi == 0} {
    set cl $cells($sempmi_coverage)
    set ws $worksheet($sempmi_coverage)
    set r $row
    set c D
  } else {
    set cl $cells1($sempmi_coverage)
    set ws $worksheet1($sempmi_coverage)
    set r $row
    set c A
  }

  if {!$multi} {set e $excel}
  if {$multi} {set e $excel1}
  
  set n 0
  set legend {{"Values as Compared to NIST Test Case Drawing" ""} \
              {"See Help > NIST CAD Models" ""} \
              {"Match" "green"} \
              {"More than expected" "cyan"} \
              {"Less than expected" "yellow"} \
              {"None (0/n)" "red"} \
              {"Unexpected (n/0)" "magenta"} \
              {"Not in CAx-IF Recommended Practice" "gray"}}
  foreach item $legend {
    set str [lindex $item 0]
    $cl Item $r $c $str

    set range [$ws Range $c$r]
    [$range Font] Bold [expr 1]

    set color [lindex $item 1]
    if {$color != ""} {[$range Interior] Color $legendColor($color)}

    if {[expr {int([$e Version])}] >= 12} {
      [[$range Borders] Item [expr 10]] Weight [expr 2]
      [[$range Borders] Item [expr 7]] Weight [expr 2]
      incr n
      if {$n == 1} {
        [[$range Borders] Item [expr 8]] Weight [expr 2]
      } elseif {$n == [llength $legend]} {
        [[$range Borders] Item [expr 9]] Weight [expr 2]
      }
    }
    incr r    
  }
}

# -------------------------------------------------------------------------------
# start PMI Presentation coverage analysis worksheet
proc gpmiCoverageStart {{multi 1}} {
  global cells cells1 gpmiTypes multiFileDir opt pmi_coverage recPracNames
  global sheetLast worksheet worksheet1 worksheets worksheets1 
  #outputMsg "gpmiCoverageStart $multi" red
  
  if {[catch {
    set pmi_coverage "PMI Presentation Coverage"

# multiple files
    if {$multi} {
      if {$opt(PMISEM)} {
        set worksheet1($pmi_coverage) [$worksheets1 Item [expr 3]]
      } else {
        set worksheet1($pmi_coverage) [$worksheets1 Item [expr 2]]
      }
      #$worksheet1($pmi_coverage) Activate
      $worksheet1($pmi_coverage) Name $pmi_coverage
      set cells1($pmi_coverage) [$worksheet1($pmi_coverage) Cells]
      $cells1($pmi_coverage) Item 1 1 "STEP Directory"
      $cells1($pmi_coverage) Item 1 2 "[file nativename $multiFileDir]"
      $cells1($pmi_coverage) Item 3 1 "PMI Presentation Names"
      set range [$worksheet1($pmi_coverage) Range "B1:K1"]
      [$range Font] Bold [expr 1]
      $range MergeCells [expr 1]
      set row1($pmi_coverage) 3

# single file
    } else {
      set sempmi_coverage "PMI Representation Coverage"
      set n 3
      if {[info exists worksheet($sempmi_coverage)]} {
        set n 5
      }
      set worksheet($pmi_coverage) [$worksheets Add [::tcom::na] $sheetLast]
      #$worksheet($pmi_coverage) Activate
      $worksheet($pmi_coverage) Name $pmi_coverage
      set cells($pmi_coverage) [$worksheet($pmi_coverage) Cells]
      set wsCount [$worksheets Count]
      [$worksheets Item [expr $wsCount]] -namedarg Move Before [$worksheets Item [expr $n]]
      $cells($pmi_coverage) Item 3 1 "PMI Presentation Names"
      $cells($pmi_coverage) Item 3 2 "Count"
      set range [$worksheet($pmi_coverage) Range "1:3"]
      [$range Font] Bold [expr 1]
      set row($pmi_coverage) 3
    }
      
    foreach item $gpmiTypes {
      set str [join $item]
      if {$multi} {
        $cells1($pmi_coverage) Item [incr row1($pmi_coverage)] 1 $str
      } else {
        $cells($pmi_coverage) Item [incr row($pmi_coverage)] 1 $str
      }
    }
  } emsg3]} {
    errorMsg "ERROR starting PMI Presentation Coverage worksheet: $emsg3"
  }
}

# -------------------------------------------------------------------------------
# write PMI coverage analysis worksheet
proc gpmiCoverageWrite {{fn ""} {sum ""} {multi 1}} {
  global cells cells1 col1 gpmiTypes gpmiTypesInvalid gpmiTypesPerFile pmi_coverage pmi_rows pmi_totals
  global worksheet worksheet1 legendColor
  #outputMsg "gpmiCoverageWrite $multi " red

  if {[catch {
    if {$multi} {
      set range [$worksheet1($pmi_coverage) Range [cellRange 3 $col1($sum)] [cellRange 3 $col1($sum)]]
      $range Orientation [expr 90]
      $range HorizontalAlignment [expr -4108]
      $cells1($pmi_coverage) Item 3 $col1($sum) $fn
    }
  
# add invalid pmi types to column A
# need to fix when there are invalid types, but a subsequent file does not if processing multiple files
    set r1 [expr {[llength $gpmiTypes]+4}]
    if {![info exists pmi_rows]} {set pmi_rows 35}
    set ok 1

    if {[info exists gpmiTypesInvalid]} {
      #outputMsg "gpmiTypesInvalid  $multi  $gpmiTypesInvalid" red
      while {$ok} {
        if {$multi} {
          set val [[$cells1($pmi_coverage) Item $r1 1] Value]
        } else {
          set val [[$cells($pmi_coverage) Item $r1 1] Value]
        }
        if {$val == ""} {
          foreach idx $gpmiTypesInvalid {
            if {$multi} {
              $cells1($pmi_coverage) Item $r1 1 $idx
              [[$worksheet1($pmi_coverage) Range [cellRange $r1 1] [cellRange $r1 1]] Interior] Color $legendColor(red)
            } else {
              $cells($pmi_coverage) Item $r1 1 $idx
              [[$worksheet($pmi_coverage) Range [cellRange $r1 1] [cellRange $r1 1]] Interior] Color $legendColor(red)
            }
            if {$r1 > $pmi_rows} {set pmi_rows $r1}
            incr r1
          }
          set ok 0
        } else {
          foreach idx $gpmiTypesInvalid {
            if {$idx != $val} {
              incr r1
              if {$multi} {
                $cells1($pmi_coverage) Item $r1 1 $idx
                [[$worksheet1($pmi_coverage) Range [cellRange $r1 1] [cellRange $r1 1]] Interior] Color $legendColor(red)
              } else {
                $cells($pmi_coverage) Item $r1 1 $idx
                [[$worksheet($pmi_coverage) Range [cellRange $r1 1] [cellRange $r1 1]] Interior] Color $legendColor(red)
              }
              set val $idx
              if {$r1 > $pmi_rows} {set pmi_rows $r1}
            }
          }
          set ok 0      
        }
      }
    }

# add numbers
    if {[info exists gpmiTypesPerFile]} {
      set gpmiTypesPerFile [lrmdups $gpmiTypesPerFile]
      for {set r 4} {$r <= $pmi_rows} {incr r} {
        if {$multi} {
          set val [[$cells1($pmi_coverage) Item $r 1] Value]
        } else {
          set val [[$cells($pmi_coverage) Item $r 1] Value]
        }
        foreach item $gpmiTypesPerFile {
          set idx [lindex [split $item "/"] 0]
          if {$val == $idx} {

# get current value
            if {$multi} {
              set npmi [[$cells1($pmi_coverage) Item $r $col1($sum)] Value]
            } else {
              set npmi [[$cells($pmi_coverage) Item $r 2] Value]
            }

# set or increment npmi
            if {$npmi == ""} {
              set npmi 1
            } else {
              set npmi [expr {int($npmi)+1}]
            }

# write npmi
            if {$multi} {
              $cells1($pmi_coverage) Item $r $col1($sum) $npmi
              set range [$worksheet1($pmi_coverage) Range [cellRange $r $col1($sum)] [cellRange $r $col1($sum)]]
              incr pmi_totals($r)
            } else {
              $cells($pmi_coverage) Item $r 2 $npmi
              set range [$worksheet($pmi_coverage) Range [cellRange $r 2] [cellRange $r 2]]
            }
            $range HorizontalAlignment [expr -4108]
          }
        }
      }
      catch {if {$multi} {unset gpmiTypesPerFile}}
    }
  } emsg3]} {
    errorMsg "ERROR adding to PMI Presentation Coverage worksheet: $emsg3"
  }
}

# -------------------------------------------------------------------------------
# format PMI coverage analysis worksheet, also PMI totals
proc gpmiCoverageFormat {{sum ""} {multi 1}} {
  global cells cells1 col1 excel excel1 gpmiTypes lenfilelist localName opt
  global pmi_coverage pmi_rows pmi_totals recPracNames stepAP worksheet worksheet1
  #outputMsg "gpmiCoverageFormat $multi" red

# delete worksheet if no graphical PMI
  if {$multi && ![info exists pmi_totals]} {
    catch {$excel1 DisplayAlerts False}
    $worksheet1($pmi_coverage) Delete
    catch {$excel1 DisplayAlerts True}
    return
  }
 
# total PMI
  if {[catch {
    if {$multi} {
      set col1($pmi_coverage) [expr {$lenfilelist+2}]
      $cells1($pmi_coverage) Item 3 $col1($pmi_coverage) "Total PMI"
      foreach idx [array names pmi_totals] {
        $cells1($pmi_coverage) Item $idx $col1($pmi_coverage) $pmi_totals($idx)
      }        
      $worksheet1($pmi_coverage) Activate
    }
 
# horizontal break lines
    set idx1 [list 21 28 30 35 36]
    if {!$multi} {set idx1 [list 3 4 21 28 30 35 36]}
    for {set r 100} {$r >= 35} {incr r -1} {
      if {$multi} {
        set val [[$cells1($pmi_coverage) Item $r 1] Value]
      } else {
        set val [[$cells($pmi_coverage) Item $r 1] Value]
      }
      if {$val != ""} {
        lappend idx1 [expr {$r+1}]
        break
      }
    }    

# horizontal lines
    foreach idx $idx1 {
      if {$multi} {
        set range [$worksheet1($pmi_coverage) Range [cellRange $idx 1] [cellRange $idx $col1($pmi_coverage)]]
      } else {
        set range [$worksheet($pmi_coverage) Range [cellRange $idx 1] [cellRange $idx 2]]        
      }
      catch {[[$range Borders] Item [expr 8]] Weight [expr 2]}
    }

# rec prac
    set rp "$recPracNames(pmi242), Sec. 8.4, Table 14"
    if {$stepAP == "AP203"} {set rp "$recPracNames(pmi203), Sec. 4.3, Table 1"}
    
# vertical line(s)
    if {$multi} {
      set range [$worksheet1($pmi_coverage) Range [cellRange 1 $col1($pmi_coverage)] [cellRange [expr {[lindex $idx1 end]-1}] $col1($pmi_coverage)]]
      catch {[[$range Borders] Item [expr 7]] Weight [expr 2]}
      
# fix row 3 height and width
      set range [$worksheet1($pmi_coverage) Range 3:3]
      $range RowHeight 300
      [$worksheet1($pmi_coverage) Columns] AutoFit
      
      $cells1($pmi_coverage) Item [expr {$pmi_rows+2}] 1 "Presentation Names defined in $rp"
      set anchor [$worksheet1($pmi_coverage) Range [cellRange [expr {$pmi_rows+2}] 1]]
      [$worksheet1($pmi_coverage) Hyperlinks] Add $anchor [join "http://www.cax-if.org/joint_testing_info.html#recpracs"] [join ""] [join "Link to CAx-IF Recommended Practices"]
  
      [$worksheet1($pmi_coverage) Rows] AutoFit
      [$worksheet1($pmi_coverage) Range "B4"] Select
      catch {[$excel1 ActiveWindow] FreezePanes [expr 1]}
      [$worksheet1($pmi_coverage) Range "A1"] Select
      catch {[$worksheet1($pmi_coverage) PageSetup] PrintGridlines [expr 1]}

# single file
    } else {
      set i1 3
      for {set i 0} {$i < $i1} {incr i} {
        set range [$worksheet($pmi_coverage) Range [cellRange 3 [expr {$i+1}]] [cellRange [expr {[lindex $idx1 end]-1}] [expr {$i+1}]]]
        catch {[[$range Borders] Item [expr 7]] Weight [expr 2]}
      }
      [$worksheet($pmi_coverage) Columns] AutoFit        
      
      catch {$cells($pmi_coverage) Item 1 5 "Presentation Names defined in $rp"}
      set range [$worksheet($pmi_coverage) Range E1:O1]
      $range MergeCells [expr 1]
      set anchor [$worksheet($pmi_coverage) Range E1]
      [$worksheet($pmi_coverage) Hyperlinks] Add $anchor [join "http://www.cax-if.org/joint_testing_info.html#recpracs"] [join ""] [join "Link to CAx-IF Recommended Practices"]
      
      [$worksheet($pmi_coverage) Range "A1"] Select
      catch {[$worksheet($pmi_coverage) PageSetup] PrintGridlines [expr 1]}
      $cells($pmi_coverage) Item 1 1 [file tail $localName]
      $cells($pmi_coverage) Item [expr {$pmi_rows+3}] 1 "See Help > PMI Coverage Analysis"

# add images for the CAx-IF and NIST PMI models
      pmiAddModelPictures $pmi_coverage
    }
# errors
  } emsg]} {
    errorMsg "ERROR formatting PMI Presentation Coverage worksheet: $emsg"
  }
}
