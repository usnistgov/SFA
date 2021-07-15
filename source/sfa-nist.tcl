# read expected PMI from spreadsheets, (called from sfa-gen.tcl)
proc nistReadExpectedPMI {} {
  global mytemp nistName nistPMImaster nistVersion spmiCoverages wdir

  if {[catch {
    set lf 1
    if {![info exists spmiCoverages]} {

# check mytemp dir
      checkTempDir

# first mount NIST zip file with images and expected PMI
      if {$nistVersion} {
        set fn [file join $wdir SFA-NIST-files.zip]
        if {[file exists $fn]} {
          file copy -force [file join $wdir $fn] $mytemp
          vfs::zip::Mount [file join $mytemp $fn] NIST
        }
      }

# get PMI coverage
      set fn "SFA-PMI-NIST-coverage.csv"
      if {[file exists NIST/$fn]} {file copy -force NIST/$fn [file join $mytemp $fn]}
      set fname [file nativename [file join $mytemp $fn]]

      if {[file exists $fname]} {
        outputMsg "\nReading Expected PMI Representation Coverage"
        set lf 0
        set f [open $fname r]
        set r 0
        while {[gets $f line] >= 0} {
          set lline [split $line ","]
          set c 0
          if {$r == 0} {
            foreach colName $lline {
              if {$colName != ""} {
                if {[string first "ctc" $colName] == 0 || [string first "ftc" $colName] == 0} {
                  set i2($c) "nist_$colName"
                } else {
                  set i2($c) "$colName"
                }
              }
              incr c
            }
          } else {
            set i1 [lindex $lline 0]
            foreach cval $lline {
              if {[info exists i2($c)] && $c > 0} {set spmiCoverages($i1,$i2($c)) $cval}
              incr c
            }
          }
          incr r
        }
        close $f
      }
    }

# get expected PMI
    if {![info exists nistPMImaster($nistName)]} {
      catch {unset nistPMImaster($nistName)}
      set fn "SFA-PMI-$nistName.xlsx"
      if {[file exists NIST/$fn]} {file copy -force NIST/$fn $mytemp}
      set fname [file nativename [file join $mytemp $fn]]

      if {[file exists $fname]} {
        if {$lf} {outputMsg " "}
        outputMsg "Reading Expected PMI for: $nistName (See Help > Analyze > NIST CAD Models)" blue
        set pid1 [twapi::get_process_ids -name "EXCEL.EXE"]
        set excel2 [::tcom::ref createobject Excel.Application]
        set pid2 [lindex [intersect3 $pid1 [twapi::get_process_ids -name "EXCEL.EXE"]] 2]

        $excel2 Visible 0
        set workbooks2  [$excel2 Workbooks]
        set worksheets2 [[$workbooks2 Open $fname] Worksheets]

        set matrix [GetWorksheetAsMatrix [$worksheets2 Item [expr 1]]]
        set r1 [llength $matrix]
        for {set r 0} {$r < $r1} {incr r} {
          set typ [lindex [lindex $matrix $r] 0]
          set pmi [lindex [lindex $matrix $r] 1]
          if {$typ != "" && $pmi != ""} {lappend nistPMImaster($nistName) "$typ\\$pmi"}
        }

        $workbooks2 Close
        $excel2 Quit
        update idletasks
        catch {unset excel2}
        after 100
        for {set i 0} {$i < 20} {incr i} {catch {twapi::end_process $pid2 -force}}
      }
    }

# errors
  } emsg]} {
    errorMsg "ERROR reading Expected PMI spreadsheet: $emsg"
  }
}

# -------------------------------------------------------------------------------
# get expected PMI for PMI Representation Summary worksheet
proc nistGetSummaryPMI {} {
  global nistName nistPMIactual nistPMIexpected nistPMIexpectedNX nistPMIfound nistPMImaster nsimilar opt pmiType spmiSumName tolNames worksheet

# add pictures
  nistAddModelPictures $spmiSumName
  [$worksheet($spmiSumName) Range "A1"] Select

# get expected PMI values from nistPMImaster
  set nsimilar 0
  if {[info exists nistPMImaster($nistName)]} {
    catch {unset nistPMIexpected($nistName)}
    catch {unset nistPMIexpectedNX($nistName)}

# read master PMI values, remove leading and trailing zeros, other stuff, add to nistPMIexpected
    foreach item $nistPMImaster($nistName) {
      set c1 [string first "\\" $item]
      set typ [string range $item 0 $c1-1]

      if {!$opt(PMISEMDIM) || $typ == "dimensional_characteristic_representation"} {
        set pmi [string range $item $c1+1 end]
        set newpmi [pmiRemoveZeros $pmi]
        lappend nistPMIexpected($nistName) $newpmi

# look for 'nX' in expected
        set c1 [string first "X" $newpmi]
        if {$c1 < 3} {
          set newpminx [string range $newpmi $c1+1 end]
          lappend nistPMIexpectedNX($nistName) [string trim $newpminx]
        } else {
          lappend nistPMIexpectedNX($nistName) $newpmi
        }

        if {[string first "tolerance" $typ] != -1} {
          foreach nam $tolNames {if {[string first $nam $typ] != -1} {set pmiType($newpmi) $nam}}
        } else {
          set pmiType($newpmi) $typ
        }
        set nistPMIactual($newpmi) $pmi
      }
    }
  }
  set nistPMIfound {}
}

# -------------------------------------------------------------------------------
# check actual vs. expected PMI for NIST files
proc nistCheckExpectedPMI {val entstr} {
  global cells legendColor nistExpectedPMI nistName nistPMIactual nistPMIexpected nistPMIexpectedNX nistPMIfound
  global pmiModifiers pmiType pmiUnicode spmiSumName spmiSumRow tolNames worksheet

# modify (composite ..) from value to just (composite)
  set c1 [string first "(composite" $val]
  if {$c1 > 0} {
    set c2 [string first ")" $val]
    set val [string range $val 0 $c1+9][string range $val $c2 end]
  }

# remove (oriented)
  set c1 [string first "(oriented)" $val]
  if {$c1 > 0} {set val [string range $val 0 $c1-2]}

# remove (movable)
  set c1 [string first "(movable)" $val]
  if {$c1 > 0} {set val [string range $val 0 $c1-2]}

# remove between
  set c1 [string first $pmiModifiers(between) $val]
  if {$c1 > 0} {set val [string range $val 0 $c1-2]}

# remove circle (I) (typically found in CATIA files)
  set c1 [string first "\u24BE" $val]
  if {$c1 > 0} {regsub "\u24BE" $val "" val}

# remove <CF> (typically found in CATIA files)
  set c1 [string first "<CF>" $val]
  if {$c1 > 0} {regsub "<CF>" $val "" val}

# remove datum feature on a dimension (typically found in CATIA files)
  if {$entstr == "dimensional_characteristic_representation"} {
    set c1 [string first "\u25BD" $val]
    if {$c1 != -1} {set val [string range $val 0 $c1-5]}
  }

# remove zeros from val
  set val [pmiRemoveZeros $val]
  if {[string first "tolerance" $entstr] != -1} {
    foreach nam $tolNames {if {[string first $nam $entstr] != -1} {set valType($val) $nam}}
  } else {
    set valType($val) $entstr
  }

# -------------------------------------------------------------------------------
# search for PMI in nistPMIexpected list
  set pmiMatch [lsearch $nistPMIexpected($nistName) $val]
  #outputMsg "$val  $pmiMatch $valType($val)" blue
  #outputMsg $nistPMIexpected($nistName)

# found in list, remove from nistPMIexpected
  if {$pmiMatch != -1} {
    [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(green)
    set nistPMIexpected($nistName)   [lreplace $nistPMIexpected($nistName)   $pmiMatch $pmiMatch]
    set nistPMIexpectedNX($nistName) [lreplace $nistPMIexpectedNX($nistName) $pmiMatch $pmiMatch]
    lappend nistPMIfound $val
    incr nistExpectedPMI(exact)
    #outputMsg " exact match  $val $pmiMatch $valType($val)" green

# -------------------------------------------------------------------------------
# not found
  } else {
    set pmiMatch 0
    set pmiMissing ""
    set pmiSimilar ""

# check each value in nistPMIexpected
    foreach pmi $nistPMIexpected($nistName) {

# simple match, remove from nistPMIexpected
      if {$val == $pmi && $pmiMatch != 1} {
        set pmiMatch 1
        set pos [lsearch $nistPMIexpected($nistName) $pmi]
        set nistPMIexpected($nistName)   [lreplace $nistPMIexpected($nistName)   $pos $pos]
        set nistPMIexpectedNX($nistName) [lreplace $nistPMIexpectedNX($nistName) $pos $pos]
        lappend nistPMIfound $pmi
        #outputMsg " simple match  $val $pmiMatch $valType($val)" green
      }
    }

# -------------------------------------------------------------------------------
# try match dimensions to expected without 'nX'
    if {$pmiMatch == 0 && $entstr == "dimensional_characteristic_representation"} {
      set c1 [string first "X" $val]
      if {$c1 < 3} {
        set valnx [string trim [string range $val $c1+1 end]]
        set pmiMatchNX [lsearch $nistPMIexpectedNX($nistName) $valnx]
        if {$pmiMatchNX != -1} {
          set pmiMatch 0.95
          set pmiSim $pmiMatch
          #outputMsg " exact match NX  $val $pmiMatchNX $valType($val)" red
          set pmiSimilar $nistPMIactual([lindex $nistPMIexpected($nistName) $pmiMatchNX])
          set nistPMIexpected($nistName)   [lreplace $nistPMIexpected($nistName)   $pmiMatchNX $pmiMatchNX]
          set nistPMIexpectedNX($nistName) [lreplace $nistPMIexpectedNX($nistName) $pmiMatchNX $pmiMatchNX]
          set pf $val
        }

# try simple match as above
        if {$pmiMatch != 0.95} {
          foreach pmi $nistPMIexpectedNX($nistName) {
            if {$valnx == $pmi && $pmiMatch != 1} {
              set pmiMatch 0.95
              set pmiSim $pmiMatch
              #outputMsg " simple match NX  $valnx $pmiMatch $valType($val)" red
              set pos [lsearch $nistPMIexpected($nistName) $valnx]
              if {$pos != -1} {
                set pmiSimilar $nistPMIactual([lindex $nistPMIexpected($nistName) $pos])
                set nistPMIexpected($nistName)   [lreplace $nistPMIexpected($nistName)   $pos $pos]
                set nistPMIexpectedNX($nistName) [lreplace $nistPMIexpectedNX($nistName) $pos $pos]
              }
              set pf $val
            }
          }
        }
      }
    }

# -------------------------------------------------------------------------------
# no match yet
    if {$pmiMatch == 0} {
      foreach pmi $nistPMIexpected($nistName) {

# look for similar strings
        set look 0
        if {$valType($val) == $pmiType($pmi) && $val != "" && $pmiMatch < 0.9} {set look 1}
        if {$valType($val) == "datum_target" && $pmiType($pmi) == "placed_datum_target_feature" && $val != "" && $pmiMatch < 0.9} {set look 1}

        if {$look} {
          set ok 1

# check for bad dimensions
          if {$valType($val) == "dimensional_characteristic_representation"} {
            #outputMsg "A$val\A [string first "-" $val] [string first "$pmiUnicode(diameter) " $val] [string first "$pmiUnicode(plusminus)" $val]"
            if {[string first "-" $val] == 0 || [string first "$pmiUnicode(diameter) " $val] == 0 || \
                [string first "$pmiUnicode(plusminus)" $val] == 0} {set ok 0}
          }

# -------------------------------------------------------------------------------
# do similarity match
          if {$ok} {
            set pmiSim 0

# datum targets
            if {[string first "datum_target" $valType($val)] != -1} {
              if {[string range $val 0 1] == [string range $pmi 0 1]} {set pmiSim 0.95}

# dimensions
            } elseif {[string first "dimension" $valType($val)] != -1} {
              set diff [expr {[string length $pmi] - [string length $val]}]
              if {$diff <= 2 && $diff >= 0 && [string first $val $pmi] != -1} {
                set pmiSim 0.95
                #outputMsg "$val / $diff / $pmi / $pmiSim / aa"
              } elseif {[string is integer [string index $val 0]] || \
                        [string range $val 0 1] == [string range $pmi 0 1]} {
                set pmiSim [stringSimilarity $val $pmi]
                #outputMsg "$val / $diff / $pmi / $pmiSim / bb"
              }

# tolerances
            } elseif {[string first "tolerance" $valType($val)] != -1} {
              if {[string first $val $pmi] != -1 || [string first $pmi $val] != -1} {
                set pmiSim 0.95
                #outputMsg "$val / $pmi / $pmiSim cc" green
              } else {
                set tol $pmiUnicode([string range $valType($val) 0 [string last "_" $valType($val)]-1])
                set pmiSim [stringSimilarity $val $pmi]

# make sure tolerance datum features are the same
                if {[string index $val end] == "\]" && [string index $pmi end] == "\]"} {
                  if {[string index $val end-1] != [string index $pmi end-1]} {
                    set pmiSim [expr {$pmiSim-0.025}]
                  } else {
                    set pmiSim [expr {$pmiSim+0.025}]
                  }
                }
                #outputMsg "$val / $pmi / $pmiSim dd" green

                if {$pmiSim < 0.9} {
                  set sval [split $val $tol]
                  if {[string length [lindex $sval 0]] > 0} {
                    set spmi [split $pmi $tol]
                    if {[string length [lindex $spmi 0]] > 0} {
                      if {[lindex $sval 1] == [lindex $spmi 1]} {set pmiSim 0.95}
                    }
                  }
                }
              }

# datum features
            } else {
              set pmiSim [stringSimilarity $val $pmi]
              #outputMsg "$val $pmiSim"
            }

            if {$pmiSim < 0.6} {
              if {[string first $val $pmi] != -1 || [string first $pmi $val] != -1 || $valType($val) == "flatness_tolerance"} {set pmiSim 0.6}
            }

# -------------------------------------------------------------------------------
# keep best match
            if {$pmiSim > $pmiMatch} {
              #outputMsg "$pmiSim / $val / $pmi / [string first $val $pmi]" red
              set pmiMatch $pmiSim
              if {[string first "datum_target" $valType($val)] == -1 && [string first "dimension" $valType($val)] == -1} {
                if {$pmiSim >= 0.6} {
                  set pmiSimilar $nistPMIactual($pmi)
                  #append pmiSimilar "[format "%c" 10](Similarity: [string range $pmiMatch 0 4])"
                }

# dimensions
              } elseif {[string first "dimension" $valType($val)] != -1} {
                #outputMsg "$pmiSim / $val / $pmi / [string first $val $pmi]" red
                if {$pmiSim >= 0.6} {set pmiSimilar $nistPMIactual($pmi)}

# dimension rounding issues
                if {$pmiSim > 0.85} {
                  set c1 [string first " " $val]
                  if {$c1 != -1} {
                    set c2 [string first " " $pmi]
                    if {$c2 != -1} {
                      if {[string range $val $c1+1 end] == [string range $pmi $c2+1 end]} {
                        set val1 [string range $val 0 $c1-1]
                        set pmi1 [string range $pmi 0 $c2-1]
                        if {[string index $val1 0] == [string index $pmi1 0]} {
                          set c3 0
                          if {[string index $val 0] == $pmiUnicode(diameter)} {set c3 1}
                          set val1 [string range $val1 $c3 end]
                          set pmi1 [string range $pmi1 $c3 end]
                          set diff 1.
                          catch {set diff [expr {abs($val1-$pmi1)}]}
                          if {$diff < 0.00101} {
                            set pmiSim 1.
                            set pmiMatch $pmiSim
                          }
                        }
                      }
                    }
                  }
                }

# handle datum targets differently
              } elseif {[string first "datum_target" $valType($val)] != -1} {
                set pmiSim 0.95
                set pmiMatch $pmiSim
                set pmiSimilar $pmi
              }
              set pf $pmi
            }
          }
        }
      }
    }

# -------------------------------------------------------------------------------
# perfect match, green
    if {$pmiMatch == 1} {
      [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(green)
      incr nistExpectedPMI(exact)

# partial and possible match, cyan and yellow
    } elseif {$pmiMatch >= 0.6} {
      if {$pmiMatch >= 0.9} {
        [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(cyan)
        incr nistExpectedPMI(partial)
      } else {
        [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(yellow)
        incr nistExpectedPMI(possible)
      }

# add similar pmi
      if {[info exists pmiSimilar] && $pmiSimilar != ""} {
        $cells($spmiSumName) Item $spmiSumRow 4 "'$pmiSimilar"
        [[$worksheet($spmiSumName) Range D$spmiSumRow] Interior] Color $legendColor(gray)
        set range [$worksheet($spmiSumName) Range D$spmiSumRow]
        catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 1]}}
        incr nsimilar

        if {$nsimilar == 1} {
          [$worksheet($spmiSumName) Range [cellRange -1 4]] ColumnWidth [expr 48]
          $cells($spmiSumName) Item 3 4 "Similar PMI"
          addCellComment $spmiSumName 3 4 "Similar PMI is the best match of the PMI Representation in column C, for Partial or Possible matches (blue and yellow), to the expected PMI in the NIST test case drawing to the right."
          set range [$worksheet($spmiSumName) Range D3]
          [$range Font] Bold [expr 1]
          catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 2]}}
        }
      }
      lappend nistPMIfound $pf

# no match red
    } else {
      [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(red)
      incr nistExpectedPMI(no)
    }
  }

# border
  catch {[[[$worksheet($spmiSumName) Range C$spmiSumRow] Borders] Item [expr 9]] Weight [expr 1]}
}

# -------------------------------------------------------------------------------
proc nistPMICoverage {nf} {
  global cells legendColor nistCoverageLegend nistCoverageStyle nistPMIexpected nistName
  global pmiElementsMaxRows spmiCoverages spmiCoverageWS totalPMIrows usedPMIrows worksheet

  foreach idx [lsort [array names spmiCoverages]] {
    set tval [lindex [split $idx ","] 0]
    set fnam [lindex [split $idx ","] 1]
    if {$fnam == $nistName} {set coverage($tval) $spmiCoverages($idx)}
  }

# check values for color-coding
  for {set r 4} {$r <= $pmiElementsMaxRows} {incr r} {
    set ttyp [[$cells($spmiCoverageWS) Item $r 1] Value]
    if {$ttyp != ""} {
      set tval [[$cells($spmiCoverageWS) Item $r 2] Value]
      if {$tval == ""} {set tval 0}
      set tval [expr {int($tval)}]

      foreach item [array names coverage] {
        if {[string first $item $ttyp] == 0} {
          set ok 0

# these words appear in other PMI elements and need to be handled separately
          if {$item != "datum" && $item != "line" && $item != "spherical" && $item != "basic" && $item != "point"} {
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

# check tolerance zone diameter vs. within a cylinder
            set skip 0
            if {$item == "tolerance zone diameter" &&          $tval == 0 && [[$cells($spmiCoverageWS) Item 20 2] Value] != ""} {set skip 1}
            if {$item == "tolerance zone within a cylinder" && $tval == 0 && [[$cells($spmiCoverageWS) Item 19 2] Value] != ""} {set skip 1}

# too few - yellow or red
            if {!$skip} {
              if {$tval < $ci} {
                set str "'$tval/$ci"
                set tc [expr {double($tval)/double($ci)}]
                $cells($spmiCoverageWS) Item $r 2 $str
                [$worksheet($spmiCoverageWS) Range B$r] HorizontalAlignment [expr -4108]
                set nistCoverageLegend 1
                if {$tval == 0} {
                  set clr "red"
                  set totalPMIrows($r) 1
                  lappend usedPMIrows $r
                } elseif {$tc < 0.334} {
                  set clr "orange"
                } elseif {$tc < 0.665} {
                  set clr "yellow"
                } else {
                  set clr "yelgre"
                }
                [[$worksheet($spmiCoverageWS) Range B$r] Interior] Color $legendColor($clr)
                lappend nistCoverageStyle "$r $nf $clr $str"

# too many - cyan or magenta
              } elseif {$tval > $ci && $tval != 0} {
                set ci1 $coverage($item)
                set clr "cyan"
                if {$ci1 == ""} {
                  set ci1 0
                  set clr "magenta"
                }
                set str "'$tval/[expr {int($ci1)}]"
                $cells($spmiCoverageWS) Item $r 2 $str
                [[$worksheet($spmiCoverageWS) Range B$r] Interior] Color $legendColor($clr)
                [$worksheet($spmiCoverageWS) Range B$r] NumberFormat "@"
                set nistCoverageLegend 1
                lappend nistCoverageStyle "$r $nf $clr $str"

# just right - green
              } elseif {$tval != 0} {
                [[$worksheet($spmiCoverageWS) Range B$r] Interior] Color $legendColor(green)
                set nistCoverageLegend 1
                lappend nistCoverageStyle "$r $nf green"
              }
            }
          }
        }
      }
    }
  }

# summarize PMI Representation Summary on PMI Representation Coverage worksheet
  if {[info exists nistPMIexpected($nistName)]} {nistAddExpectedPMIPercent $nf}
}

# -------------------------------------------------------------------------------
# summarize PMI Representation Summary with percentages
proc nistAddExpectedPMIPercent {nf} {
  global cells cells1 legendColor lenfilelist nistCoverageStyle nistName nistExpectedPMI
  global nistPMIexpected nistPMIfound pmiElementsMaxRows spmiCoverageWS worksheet worksheet1

# compute missing and total expected PMI
  set nistExpectedPMI(missing) [llength [lindex [intersect3 $nistPMIexpected($nistName) $nistPMIfound] 0]]
  set nistExpectedPMI(total) 0
  foreach idx {exact partial possible missing} {
    if {[info exists nistExpectedPMI($idx)]} {incr nistExpectedPMI(total) $nistExpectedPMI($idx)}
  }

  set r [expr {$pmiElementsMaxRows+4}]
  $cells($spmiCoverageWS) Item $r 1 "Expected PMI[format "%c" 10]  (See PMI Representation Summary worksheet)"
  $cells($spmiCoverageWS) Item $r 2 "%"
  set range [$worksheet($spmiCoverageWS) Range A$r:B$r]
  [$range Font] Bold [expr 1]
  catch {foreach i {8 10 11} {[[$range Borders] Item $i] Weight [expr 2]}}
  set range [$worksheet($spmiCoverageWS) Range B$r]
  $range HorizontalAlignment [expr -4108]
  addCellComment $spmiCoverageWS $r 1 "This table has color-coded percentages of Exact, Partial, Possible, Missing, and No matches from the PMI Representation Summary worksheet.  The Total PMI on which the percentages are based on is also shown.  The percentages for all matches except 'No match' should total 100.\n\nCoverage Analysis is only based on individual PMI elements.  The PMI Representation Summary is based on the entire PMI annotation and provides a better understanding of the PMI.  The Coverage Analysis might show that there is an Exact match (all green above) for all of the PMI elements, however, the Representation Summary might show less than Exact matches.\n\nSee Help > User Guide (section 6.5.2.1)"

# for multiple files, add more formatting
  if {[info exists lenfilelist]} {
    if {$lenfilelist > 1 && ![info exists nistPMIexpectedFormat]} {
      incr nistPMIexpectedFormat
      set r1 [expr {$pmiElementsMaxRows+4}]
      $cells1($spmiCoverageWS) Item $r1 1 "Expected PMI[format "%c" 10]  (% from PMI Representation Summary worksheets)"
      set range [$worksheet1($spmiCoverageWS) Range A$r]
      [$range Font] Bold [expr 1]
      set range [$worksheet1($spmiCoverageWS) Range [cellRange [expr {$pmiElementsMaxRows+4}] 1] [cellRange [expr {$pmiElementsMaxRows+10}] [expr {$lenfilelist+1}]]]
      catch {foreach i {8 9 10} {[[$range Borders] Item $i] Weight [expr 2]}}
      set r2 [expr {$pmiElementsMaxRows+9}]
      set range [$worksheet1($spmiCoverageWS) Range [cellRange $r2 1] [cellRange $r2 [expr {$lenfilelist+1}]]]
      catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 2]}}
    }
  }

# for each type of expected PMI
  set types [list exact partial possible missing no total]
  foreach idx $types {
    incr r
    set pct 0

# compute percentage
    if {$idx != "total"} {
      if {[info exists nistExpectedPMI($idx)]} {set pct [trimNum [expr {(100.*$nistExpectedPMI($idx))/$nistExpectedPMI(total)}] 0]}
      set str "[string totitle $idx] match"
    } else {
      set pct $nistExpectedPMI(total)
      set str "Total PMI"
    }
    $cells($spmiCoverageWS) Item $r 1 $str
    $cells($spmiCoverageWS) Item $r 2 $pct
    if {$idx == "exact"} {outputMsg " Expected PMI: $pct\%"}

# formatting
    set range [$worksheet($spmiCoverageWS) Range B$r]
    catch {foreach i {7 10} {[[$range Borders] Item $i] Weight [expr 2]}}

# for multiple files, add more formatting
    if {[info exists lenfilelist]} {
      if {$lenfilelist > 1 && $nistPMIexpectedFormat >= 1} {
        incr r1
        if {$idx != "total" } {
          set str "[string totitle $idx] match"
        } else {
          set str "Total PMI"
        }
        $cells1($spmiCoverageWS) Item $r1 1 $str
        incr nistPMIexpectedFormat
      }
    }

# color-code percentages
    set clr "white"
    if {$idx != "total"} {
      if {$idx == "exact"} {
        if {$pct == 100} {
          set clr "green"
        } elseif {$pct >= 85} {
          set clr "yelgre"
        } elseif {$pct >= 70} {
          set clr "yellow"
        } elseif {$pct >= 55} {
          set clr "orange"
        } else {
          set clr "red"
        }
      } elseif {$pct > 0} {
        if {$pct <= 15} {
          set clr "yelgre"
        } elseif {$pct <= 30} {
          set clr "yellow"
        } elseif {$pct <= 45} {
          set clr "orange"
        } else {
          set clr "red"
        }
      }
    }
    if {$clr != "white"} {
      [[$worksheet($spmiCoverageWS) Range B$r] Interior] Color $legendColor($clr)
    }
    lappend nistCoverageStyle "$r $nf $clr $pct"
  }
  unset nistExpectedPMI

# more formatting
  set r2 [expr {$r-1}]
  foreach row [list $r $r2] {
    set range [$worksheet($spmiCoverageWS) Range [cellRange $row 1] [cellRange $row 2]]
    catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 2]}}
  }
}

# -------------------------------------------------------------------------------
proc nistAddCoverageStyle {} {
  global cells1 legendColor nistCoverageStyle spmiCoverageWS worksheet1

  foreach item $nistCoverageStyle {
    set r [lindex [split $item " "] 0]
    set c [expr {[lindex [split $item " "] 1]+1}]
    set style [lindex [split $item " "] 2]
    if {[llength $item] > 3} {
      set str [lindex [split $item " "] 3]
      $cells1($spmiCoverageWS) Item $r $c $str
      [$worksheet1($spmiCoverageWS) Range [cellRange $r $c]] HorizontalAlignment [expr -4108]
    }
    if {$style != "white"} {[[$worksheet1($spmiCoverageWS) Range [cellRange $r $c]] Interior] Color $legendColor($style)}
  }
}

# -------------------------------------------------------------------------------
proc nistPMISummaryFormat {} {
  global cells legendColor nistName nistPMIactual nistPMIexpected nistPMIfound pmiType spmiSumName spmiSumRow worksheet

  set r [incr spmiSumRow]

# legend
  set n 0
  set legend {{"Expected PMI" ""} {"See Help > Analyze > NIST CAD Models" ""} {"Exact match" "green"} {"Partial match" "cyan"} {"Possible match" "yellow"} {"No match" "red"}}
  foreach item $legend {
    set str [lindex $item 0]
    $cells($spmiSumName) Item $r 3 $str

    set range [$worksheet($spmiSumName) Range [cellRange $r 3]]
    [$range Font] Bold [expr 1]

    set color [lindex $item 1]
    if {$color != ""} {[$range Interior] Color $legendColor($color)}

    catch {
      foreach i {7 10} {[[$range Borders] Item $i] Weight [expr 2]}
      incr n
      if {$n == 1} {
        [[$range Borders] Item [expr 8]] Weight [expr 2]
      } elseif {$n == [llength $legend]} {
        [[$range Borders] Item [expr 9]] Weight [expr 2]
      }
    }
    incr r
  }

# add missing pmi
  set pmiMissing [lindex [intersect3 $nistPMIexpected($nistName) $nistPMIfound] 0]
  if {[llength $pmiMissing] > 0} {
    incr r
    $cells($spmiSumName) Item $r 2 "Entity Type"
    $cells($spmiSumName) Item $r 3 "Missing PMI"
    set range [$worksheet($spmiSumName) Range [cellRange $r 2]  [cellRange $r 3]]
    [$range Font] Bold [expr 1]
    catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 2]}}
    foreach item $pmiMissing {
      incr r
      $cells($spmiSumName) Item $r 2 $pmiType($item)
      $cells($spmiSumName) Item $r 3 "'$nistPMIactual($item)"
      [[$worksheet($spmiSumName) Range [cellRange $r 3]] Interior] Color $legendColor(red)
      catch {[[[$worksheet($spmiSumName) Range [cellRange $r 3]] Borders] Item [expr 9]] Weight [expr 1]}
    }
  }
}

# -------------------------------------------------------------------------------
# add coverage legend
proc nistAddCoverageLegend {{multi 0}} {
  global cells cells1 legendColor lenfilelist spmiCoverageWS worksheet worksheet1

  if {$multi == 0} {
    set cl $cells($spmiCoverageWS)
    set ws $worksheet($spmiCoverageWS)
    set r 3
    set c D
  } else {
    set cl $cells1($spmiCoverageWS)
    set ws $worksheet1($spmiCoverageWS)
    set r 4
    set c [expr {$lenfilelist+4}]
  }

  set n 0
  set legend {{"Values as Compared to NIST Test Case Drawing" ""} \
              {"See Help > Analyze > NIST CAD Models" ""} \
              {"More than expected" "cyan"} \
              {"Exact match" "green"} \
              {"Less than expected (upper third)" "yelgre"} \
              {"Less than expected (middle third)" "yellow"} \
              {"Less than expected (lower third)" "orange"} \
              {"None (0/n)" "red"} \
              {"Unexpected (n/0)" "magenta"} \
              {"Not checked" ""}}
  foreach item $legend {
    set str [lindex $item 0]
    $cl Item $r $c $str

    set range [$ws Range [cellRange $r $c]]
    [$range Font] Bold [expr 1]

    set color [lindex $item 1]
    if {$color != ""} {[$range Interior] Color $legendColor($color)}

    catch {
      foreach i {7 10} {[[$range Borders] Item $i] Weight [expr 2]}
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
# add images for the CAx-IF and NIST PMI models
proc nistAddModelPictures {ent} {
  global cells excel localName nistModelPictures nistModelURLs mytemp nistName worksheet

  set ftail [string tolower [file tail $localName]]

  if {[catch {
    set nlink 0
    set fl ""
    foreach pic $nistModelPictures {
      set ok 0
      if {[lindex $pic 0] == $nistName} {
        set ok 1
      } elseif {$nistName == ""} {
        set pic1 [split [lindex $pic 0] "-"]
        if {[string first [lindex $pic1 0] $ftail] != -1 && [string first [lindex $pic1 1] $ftail] != -1} {
          set ok 1
        }
      }
      if {$ok} {
        set fl [lindex $pic 1]
        set fc [lindex $pic 2]

# get image from zip file
        if {[file exists NIST/$fl]} {file copy -force NIST/$fl $mytemp}
        set fn [file nativename [file join $mytemp $fl]]

        if {[file exists $fn]} {
          set cellId [[$worksheet($ent) Cells] Range "$fc:$fc"]
          $cellId Select
          set shapeId [[[$cellId Parent] Shapes] AddPicture $fn [expr 0] [expr 1] [$cellId Left] [$cellId Top] -1 -1]

# freeze panes
          if {$nlink == 0} {
            if {[string first "Representation" $ent] != -1} {
              [$worksheet($ent) Range "E4"] Select
            } else {
              [$worksheet($ent) Range "C4"] Select
            }
          }
          catch {[$excel ActiveWindow] FreezePanes [expr 1]}
          [$worksheet($ent) Range "A1"] Select

# link to test model drawings (doesn't always work)
          if {[string first "nist_" $fl] == 0 && $nlink < 2} {
            set str [string range $fl 0 10]
            foreach item $nistModelURLs {
              if {[string first $str $item] == 0} {
                catch {$cells($ent) Item 3 5 "Test Case Drawing"}
                set range [$worksheet($ent) Range E3:M3]
                $range MergeCells [expr 1]
                set range [$worksheet($ent) Range "E3"]
                [$worksheet($ent) Hyperlinks] Add $range [join "https://s3.amazonaws.com/nist-el/mfg_digitalthread/$item"] [join ""] [join "Link to Test Case Drawing (PDF)"]
                incr nlink
              }
            }
          }
        }
      }
    }
  } emsg]} {
    errorMsg "ERROR adding Picture to PMI Summary or Coverage worksheet.\n  $emsg"
  }
}

# -------------------------------------------------------------------------------------------------
proc nistGetName {} {
  global developer localName opt resetRound

  set nistName ""
  set filePrefix {}
  set prefixes {}
  for {set i 4} {$i < 9} {incr i} {lappend prefixes "sp$i"}
  for {set i 1} {$i < 3} {incr i} {lappend prefixes "tgp$i"}
  for {set i 3} {$i < 4} {incr i} {lappend prefixes "tp$i"}
  lappend prefixes "pmi"
  foreach fp $prefixes {
    lappend filePrefix "$fp\_"
    lappend filePrefix "$fp\-"
  }
  set ftail [string tolower [file tail $localName]]
  set ftail1 $ftail
  set c 3
  if {[string first "tgp" $ftail] == 0} {set c 4}
  foreach str {asme1 ap203 ap214 ap242 242 c3e} {regsub $str $ftail "" ftail}

# first check some specific names, CAx-IF ISO PMI models
  foreach part [list base cheek pole spindle] {
    if {[string first "sp" $ftail1] == 0} {
      if {[string first $part $ftail] != -1} {set nistName "sp6-$part"}
    }
    if {[string first "$part\_r"  $ftail] == 0}      {set nistName "sp6-$part"}
    if {[string first "_$part"    $ftail] != -1}     {set nistName "sp6-$part"}
    if {[string first "$part.stp" $localName] != -1} {set nistName "sp6-$part"}
  }

  if {$developer && [string first "step-file-analyzer" $ftail] == 0} {set nistName "nist_ctc_01"}

# specific name found
  if {$nistName != ""} {return $nistName}

# check for a NIST CTC or FTC
  set ctcftc 0
  set ok  0
  set ok1 0

  if {[lsearch $filePrefix [string range $ftail 0 $c]] != -1 || [string first "nist" $ftail] != -1 || \
      [string first "ctc" $ftail] != -1 || [string first "ftc" $ftail] != -1} {
    if {[lsearch $filePrefix [string range $ftail 0 $c]] != -1} {set ftail [string range $ftail $c+1 end]}

    set tmp "nist_"
    foreach item {ctc ftc} {
      if {[string first $item $ftail] != -1} {
        append tmp "$item\_"
        set ctcftc 1
      }
    }

# find nist_ctc_01 directly
    if {$ctcftc} {
      foreach zero {"0" ""} {
        for {set i 1} {$i <= 11} {incr i} {
          set i1 $i
          if {$i < 10} {set i1 "$zero$i"}
          set tmp1 "$tmp$i1"
          if {[string first $tmp1 $ftail] != -1 && !$ok1} {
            set nistName $tmp1
            set ok1 1
          }
        }
      }
    }

# find the number in the string
    if {!$ok1} {
      foreach zero {"0" ""} {
        for {set i 1} {$i <= 11} {incr i} {
          if {!$ok} {
            set i1 $i
            if {$i < 10} {
              set i1 "$zero$i"
              set k "0$i"
            } else {
              set i1 $i
              set k $i
            }
            set c {""}
            if {[string first $i1 $ftail] != [string last $i1 $ftail]} {set c {"_" "-"}}
            foreach c1 $c {
              for {set j 0} {$j < 2} {incr j} {
                if {$j == 0} {set i2 "$c1$i1"}
                if {$j == 1} {set i2 "$i1$c1"}
                if {[string first $i2 $ftail] != -1 && !$ok} {
                  if {$ctcftc} {
                    append tmp $k
                  } elseif {$i <= 5} {
                    append tmp "ctc_$k"
                  } else {
                    append tmp "ftc_$k"
                  }
                  set nistName $tmp
                  set ok 1
                }
              }
            }
          }
        }
      }
    }
  }

# check required rounding for ftc 6,7,8
  catch {unset resetRound}
  if {$opt(PMISEM)} {
    if {$opt(PMISEMRND) && $nistName == "nist_ftc_06"} {
      set resetRound $opt(PMISEMRND)
      set opt(PMISEMRND) 0
    } elseif {!$opt(PMISEMRND) && ($nistName == "nist_ftc_07" || $nistName == "nist_ftc_08" || $nistName == "nist_ftc_11")} {
      set resetRound $opt(PMISEMRND)
      set opt(PMISEMRND) 1
    }
  }
  return $nistName
}

# -------------------------------------------------------------------------------
proc pmiRemoveZeros {pmi} {
  global pmiUnicode

# line feeds
  regsub -all \n $pmi " " pmi

# |
  regsub -all "\u23B9" $pmi "" pmi
  regsub -all {\|} $pmi "" pmi

# extra spaces
  for {set j 0} {$j < 6} {incr j} {regsub -all "  " $pmi " " pmi}

  if {[string first "." $pmi] != -1} {
    set newpmi ""

# split pmi into individual strings
    foreach spmi [split $pmi " "] {
      set spmi " $spmi "
      if {[string first "." $spmi] != -1} {

# trailing zeros
        if {[string first $pmiUnicode(degree) $spmi] == -1} {
          for {set j 0} {$j < 4} {incr j} {if {[string first "0 " $spmi] != -1} {regsub -all "0 " $spmi " " spmi}}

# leading zero
          regsub -all " 0" $spmi " " spmi
          if {[string first $pmiUnicode(diameter) $spmi] != -1} {regsub -all -- "$pmiUnicode(diameter)0" $spmi $pmiUnicode(diameter) spmi}

# rectangular defined unit area. i.e., 0.50x0.50
          set c1 [string first "x" $spmi]
          if {$c1 != -1} {
            if {[string index $spmi $c1+1] != " "} {
              regsub "0x" $spmi "x" spmi
              regsub "x0" $spmi "x" spmi
            }
          }

# trailing .
          if {[string first ". " $spmi] != -1} {regsub {\. } $spmi " " spmi}

# similar check for degrees
        } else {
          set spmi "[string range $spmi 0 end-2] "
          for {set j 0} {$j < 4} {incr j} {if {[string first "0 " $spmi] != -1} {regsub -all "0 " $spmi " " spmi}}
          regsub -all " 0" $spmi " " spmi
          if {[string first ". " $spmi] != -1} {regsub {\. } $spmi " " spmi}
          set spmi "[string range $spmi 0 end-1]$pmiUnicode(degree) "
        }

# reference dimension
        if {[string first "0\]" $spmi] != -1} {for {set j 0} {$j < 4} {incr j} {regsub {0\]} $spmi "\]" spmi}}
        if {[string first ".\]" $spmi] != -1} {regsub {.\]} $spmi "\]" spmi}
        if {[string first "\[0" $spmi]  == 1} {regsub {\[0} $spmi "\[" spmi}

# basic dimension
        if {[string first "0)" $spmi] != -1} {for {set j 0} {$j < 3} {incr j} {regsub {0\)} $spmi "\)" spmi}}
        if {[string first ".)" $spmi] != -1} {regsub {\.\)} $spmi ")" spmi}
        if {[string first "\(0" $spmi] == 1} {regsub {\(0}  $spmi "(" spmi}

# radius < 1
        if {[string first "R0" $spmi] != -1} {regsub "R0" $spmi "R" spmi}

# add to newpmi
        append newpmi "$spmi "

# not a number
      } else {
        append newpmi "$spmi "
      }
    }
    set pmi [string trim [string range $newpmi 1 end-1]]
  }

# pluses, minuses
  if {[string first "+0" $pmi] != -1} {regsub -all {\+0} $pmi "\+" pmi}
  if {[string first "-0" $pmi] != -1} {regsub -all {\-0} $pmi "\-" pmi}
  for {set j 0} {$j < 6} {incr j} {regsub -all "  " $pmi " " pmi}

  return $pmi
}

#-------------------------------------------------------------------------------
# From http://wiki.tcl.tk/3070
proc stringSimilarity {a b} {
  set totalLength [max [string length $a] [string length $b]]
  return [string range [max [expr {double($totalLength-[levenshteinDistance $a $b])/$totalLength}] 0.0] 0 4]
}

proc levenshteinDistance {s t} {
  if {![set n [string length $t]]} {
    return [string length $s]
  } elseif {![set m [string length $s]]} {
    return $n
  }
  for {set i 0} {$i <= $m} {incr i} {
    lappend d 0
    lappend p $i
  }
  for {set j 0} {$j < $n} {} {
    set tj [string index $t $j]
    lset d 0 [incr j]
    for {set i 0} {$i < $m} {} {
      set a [expr {[lindex $d $i]+1}]
      set b [expr {[lindex $p $i]+([string index $s $i] ne $tj)}]
      set c [expr {[lindex $p [incr i]]+1}]
      lset d $i [expr {$a<$b ? $c<$a ? $c : $a : $c<$b ? $c : $b}]
    }
    set nd $p; set p $d; set d $nd
  }
  return [lindex $p end]
}
