# read expected PMI from spreadsheets, (called from sfa-gen.tcl)
proc nistReadExpectedPMI {{epmiFile ""}} {
  global epmiUD mytemp nistName nistPMImaster nistVersion spmiCoverages wdir

  if {[catch {
    set lf 1
    if {![info exists spmiCoverages] && $epmiFile == ""} {

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
        outputMsg "\nReading Expected Semantic PMI Coverage"
        set lf 0
        set f [open $fname r]
        set r 0
        while {[gets $f line] >= 0} {
          set lline [split $line ","]
          set c 0
          if {$r == 0} {
            foreach colName $lline {
              if {$colName != ""} {
                if {[string first "ctc" $colName] == 0 || [string first "ftc" $colName] == 0 || \
                    [string first "stc" $colName] == 0 || [string first "htc" $colName] == 0 || [string first "pdc" $colName] == 0} {
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
    if {$epmiFile == ""} {
      set name $nistName
      set ud ""
    } else {
      set name $epmiUD
      set ud " User-Defined"
    }
    if {![info exists nistPMImaster($name)]} {
      catch {unset nistPMImaster($name)}
      if {$epmiFile == ""} {
        set fn "SFA-PMI-$nistName.xlsx"
        if {[file exists NIST/$fn]} {file copy -force NIST/$fn $mytemp}
        set fname [file nativename [file join $mytemp $fn]]
      } else {
        set fname $epmiFile
      }

      if {[file exists $fname]} {
        if {$lf} {outputMsg " "}
        outputMsg "Reading$ud Expected PMI for: $name (See Help > Analyzer > NIST CAD Models)" blue
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
          if {$typ != "" && $pmi != ""} {lappend nistPMImaster($name) "$typ\\$pmi"}
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
    errorMsg "Error reading Expected PMI spreadsheet: $emsg"
  }
}

# -------------------------------------------------------------------------------
# get expected PMI for Semantic PMI Summary worksheet
proc nistGetSummaryPMI {name} {
  global nistPMIactual nistPMIexpected nistPMIexpectedGND nistPMIexpectedNX nistPMIfound nistPMImap nistPMImaster
  global nistName nsimilar opt pmiType pmiUnicode spmiSumName tolNames tolSymbols worksheet

# tolerance symbols
  foreach tol $tolNames {
    set tol [string range $tol 0 [string last "_" $tol]-1]
    lappend tolSymbols $pmiUnicode($tol)
  }

# add pictures
  if {$name == $nistName} {
    nistAddModelPictures $spmiSumName
    [$worksheet($spmiSumName) Range "A1"] Select
  }

# get expected PMI values from nistPMImaster
  set nsimilar 0
  if {[info exists nistPMImaster($name)]} {

# read master PMI values, remove leading and trailing zeros, other stuff, add to nistPMIexpected
    foreach item $nistPMImaster($name) {
      set c1 [string first "\\" $item]
      set typ [string range $item 0 $c1-1]

      set ok 0
      if {!$opt(PMISEMDIM) || $typ == "dimensional_characteristic_representation"} {set ok 1}
      if {!$opt(PMISEMDT) || $typ == "placed_datum_target_feature" || $typ == "datum_target"} {set ok 1}

      if {$ok} {
        set pmi [string range $item $c1+1 end]
        set newpmi $pmi
        if {[string first "(point)" $pmi] == -1} {set newpmi [pmiRemoveZeros $pmi]}

# remove datum feature brackets
        if {[string first "\u25BD" $newpmi] != -1} {
          set c2 [string first "\[" $newpmi]
          if {$c2 != -1} {
            if {[string index $newpmi $c2+2] == "\]"} {
              set newpmi [string range $newpmi 0 $c2-1][string index $newpmi $c2+1][string range $newpmi $c2+3 end]
            }
          }
        }

        lappend nistPMIexpected($name) $newpmi
        set nistPMImap($newpmi) $newpmi

# look for 'nX' in expected
        set c1 [string first "X" $newpmi]
        if {$c1 == 1 || $c1 == 2} {
          set newpminx [string trim [string range $newpmi $c1+1 end]]
          lappend nistPMIexpectedNX($name) $newpminx
          set nistPMImap($newpmi) $newpminx
        } else {
          lappend nistPMIexpectedNX($name) $newpmi
        }

# look for geometric tolerance with dimension above, but not with all over and all around
        if {[string first "tolerance" $typ] != -1 && [string first "\u2b69\u25CE" $newpmi] == -1 && [string first "\u232E" $newpmi] == -1} {
          set ndset 0
          foreach sym $tolSymbols {
            set c1 [string first $sym $newpmi]
            if {$c1 > 0} {
              set newpmignd [string trim [string range $newpmi $c1 end]]
              lappend nistPMIexpectedGND($name) $newpmignd
              set nistPMImap($newpmi) $newpmignd
              set ndset 1
              break
            }
          }
        }

# set pmiType
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
proc nistCheckExpectedPMI {val entstr epmiName} {
  global nistExpectedPMI nistPMIactual nistPMIdeduct nistPMIexpected nistPMIexpectedGND nistPMIexpectedNX nistPMIfound nistPMImap
  global cells legendColor pmiModifiers pmiType pmiUnicode spmiSumName spmiSumRow tolNames tolSymbols worksheet
  set debug 0

# modify (composite ..) from value to just (composite)
  set c1 [string first "(composite" $val]
  if {$c1 > 0} {
    set c2 [string first ")" $val]
    set val [string range $val 0 $c1+9][string range $val $c2 end]
  }

# remove (oriented) for NIST test case
  if {[string first "nist" $epmiName] == 0} {
    set c1 [string first "(oriented)" $val]
    if {$c1 > 0} {set val [string range $val 0 $c1-2]}

# remove between for NIST test case
    set c1 [string first $pmiModifiers(between) $val]
    if {$c1 > 0} {set val [string range $val 0 $c1-2]}
  }

# exceptions for datum features on geometric tolerances based on NIST test case
  if {$epmiName == "nist_ctc_02"} {
    set c1 [string first "\[A" $val]
    if {$c1 != -1} {
      set val [string range $val 0 $c1-16][string range $val $c1+3 end]
      set pmiException(dfa) 1
    }
    set c1 [string first "\[C" $val]
    if {$c1 != -1} {
      set val [string range $val 0 $c1-16][string range $val $c1+3 end]
      set pmiException(dfc) 1
    }
  } elseif {$epmiName == "nist_ftc_06" || $epmiName == "nist_stc_06"} {
    set c1 [string first "\[J" $val]
    if {$c1 != -1} {
      set val [string range $val 0 $c1-16][string range $val $c1+3 end]
      set pmiException(dfj) 1
    }
    set c1 [string first "\[K" $val]
    if {$c1 != -1} {
      set val [string range $val 0 $c1-16][string range $val $c1+3 end]
      set pmiException(dfk) 1
    }
  } elseif {$epmiName == "nist_ftc_10"} {
    set c1 [string first "\[K" $val]
    if {$c1 != -1} {
      set val [string range $val 0 $c1-16][string range $val $c1+3 end]
      set pmiException(dfk) 1
    }
    set c1 [string first "\[L" $val]
    if {$c1 != -1} {
      set val [string range $val 0 $c1-16][string range $val $c1+3 end]
      set pmiException(dfl) 1
    }
  }
  if {$epmiName == "nist_ftc_06" || $epmiName == "nist_stc_06" || $epmiName == "nist_ftc_08"} {
    if {[string first "ftc" $epmiName] != -1 || [string first $pmiUnicode(position) $val] == -1} {
      set c1 [string first "\[F" $val]
      if {$c1 != -1} {
        set c2 16
        set char [string index $val $c1-$c2]
        if {$char != " "} {set c2 15}
        set val [string range $val 0 $c1-$c2][string range $val $c1+3 end]
        set pmiException(dff) 1
      }
    }
  }

# exceptions for directed dimensions for NIST test case
  if {$epmiName == "nist_ftc_06" || $epmiName == "nist_ftc_07" || $epmiName == "nist_ftc_10" || $epmiName == "nist_ctc_03" || \
      $epmiName == "nist_stc_06" || $epmiName == "nist_stc_07" || $epmiName == "nist_stc_10" || $epmiName == "nist_htc"} {
    set c1 [string first "(directed)" $val]
    if {$c1 != -1} {
      set val [string trimright [string range $val 0 $c1-1]]
      set idx "Directed dimension"
      set pmiException($idx) 1
    }
  }

# remove some items for NIST test cases typically found in CATIA files
# circle (I), independency
  if {[string first "nist" $epmiName] == 0} {
    set c1 [string first "\u24BE" $val]
    if {$c1 > 0} {
      regsub "\u24BE" $val "" val
      set pmiException(Independency) 1
    }
# <CF>
    set c1 [string first "<CF>" $val]
    if {$c1 > 0} {
      regsub "<CF>" $val "" val
      set idx "Continuous feature"
      set pmiException($idx) 1
    }
# datum feature on a dimension
    if {$entstr == "dimensional_characteristic_representation"} {
      set c1 [string first "\u25BD" $val]
      if {$c1 != -1} {set val [string range $val 0 $c1-5]}
    }
  }

# JPMI test cases
  if {[string first "jpmi" $epmiName] == 0} {
# circle (E)
    set c1 [string first "\u24BA" $val]
    if {$c1 > 0} {
      regsub "\u24BA" $val "" val
      set idx "Envelope requirement"
      set pmiException($idx) 1
    }
# composite
    set c1 [string first "\n(composite" $val]
    if {$c1 > 0} {set val [string range $val 0 $c1-1]}
  }

# remove zeros from val, also removes linefeeds so FCF is one line for comparison to actual FCF
  if {[string first "(point)" $val] == -1} {set val [pmiRemoveZeros $val]}
  if {[string first "tolerance" $entstr] != -1} {
    foreach nam $tolNames {if {[string first $nam $entstr] != -1} {set valType($val) $nam}}
  } else {
    set valType($val) $entstr
  }

# remove brackets for datum features
  if {[string first "\u25BD" $val] != -1} {
    set oldType $valType($val)
    set c1 [string first "\[" $val]
    if {$c1 != -1} {
      if {[string index $val $c1+2] == "\]"} {
        set val [string range $val 0 $c1-1][string index $val $c1+1][string range $val $c1+3 end]
        set valType($val) $oldType
      }
    }
  }

# -------------------------------------------------------------------------------
# search for PMI in nistPMIexpected list
  set pmiMissing ""
  set pmiSimilar ""
  set pmiMatch [lsearch $nistPMIexpected($epmiName) $val]
  if {$debug} {outputMsg "\n$val\n[llength $nistPMIexpected($epmiName)]  $nistPMIexpected($epmiName)" red}

# found in list, remove from nistPMIexpected
  if {$pmiMatch != -1} {
    if {$debug} {outputMsg "remove1 [lindex $nistPMIexpected($epmiName) $pmiMatch]" green}
    set nistPMIexpected($epmiName)   [lreplace $nistPMIexpected($epmiName)   $pmiMatch $pmiMatch]
    set nistPMIexpectedNX($epmiName) [lreplace $nistPMIexpectedNX($epmiName) $pmiMatch $pmiMatch]
    lappend nistPMIfound $val
    set pmiMatch 1

# exceptions
    foreach idx {dfa dfc dff dfj dfk dfl} {
      if {[info exists pmiException($idx)]} {
        set pmiSimilar "Datum feature [string toupper [string index $idx 2]] is ignored"
        set pmiMatch 0.99
      }
    }

# -------------------------------------------------------------------------------
# not found
  } else {
    set pmiMatch 0

# check if val is equal to nistPMIexpected, handles case where -1 is returned above, usually with a [ or ] in the FCF, brackets for datum features are removed above
    set pos -1
    foreach pmi $nistPMIexpected($epmiName) {
      incr pos
      if {$val == $pmi} {
        #outputMsg " simple match  $val $pmiMatch $valType($val)" green
        if {$debug} {outputMsg "remove2 $pos [lindex $nistPMIexpected($epmiName) $pos]" green}
        set pmiMatch 1
        set nistPMIexpected($epmiName)   [lreplace $nistPMIexpected($epmiName)   $pos $pos]
        set nistPMIexpectedNX($epmiName) [lreplace $nistPMIexpectedNX($epmiName) $pos $pos]
        lappend nistPMIfound $pmi
        break
      }
    }

# -------------------------------------------------------------------------------
# try matching dimensions without 'nX'
    if {[catch {
      if {$pmiMatch == 0 && $entstr == "dimensional_characteristic_representation"} {
        set c1 [string first "X" [string range $val 0 3]]
        if {$c1 < 3} {
          set valnx [string trim [string range $val $c1+1 end]]
          set pmiMatchNX [lsearch $nistPMIexpectedNX($epmiName) $valnx]
          if {$pmiMatchNX != -1} {
            set pmiMatch 0.99
            set pmiSim $pmiMatch
            #outputMsg " exact match NX  $val $pmiMatchNX $valType($val)" red
            foreach item $nistPMIexpected($epmiName) {
              if {[info exists nistPMImap($item)]} {
                if {[string equal $valnx $nistPMImap($item)] == 1} {
                  set pmiSimilar "$item  (nX does not match)"
                  lappend nistPMIfound $item
                  break
                }
              }
            }
            if {$debug} {outputMsg "remove3 [lindex $nistPMIexpected($epmiName) $pmiMatchNX]" green}
            catch {
              set nistPMIexpected($epmiName)   [lreplace $nistPMIexpected($epmiName)   $pmiMatchNX $pmiMatchNX]
              set nistPMIexpectedNX($epmiName) [lreplace $nistPMIexpectedNX($epmiName) $pmiMatchNX $pmiMatchNX]
            }
          }

# try simple match as above
          if {$pmiMatch != 0.99} {
            foreach pmi $nistPMIexpectedNX($epmiName) {
              if {$valnx == $pmi && $pmiMatch != 1} {
                set pmiMatch 0.95
                set pmiSim $pmiMatch
                #outputMsg " simple match NX  $valnx $pmiMatch $valType($val)" red
                set pos [lsearch $nistPMIexpected($epmiName) $valnx]
                if {$pos != -1} {
                  if {$debug} {outputMsg "remove4 [lindex $nistPMIexpected($epmiName) $pos]" green}
                  set pmiSimilar $nistPMIactual([lindex $nistPMIexpected($epmiName) $pos])
                  set nistPMIexpected($epmiName)   [lreplace $nistPMIexpected($epmiName)   $pos $pos]
                  set nistPMIexpectedNX($epmiName) [lreplace $nistPMIexpectedNX($epmiName) $pos $pos]
                }
                lappend nistPMIfound $val
              }
            }
          }
        }
      }
    } emsg]} {
      errorMsg "Error matching dimension ($val) without nX: $emsg"
    }

# -------------------------------------------------------------------------------
# try matching geometric tolerance without the dimension (nistPMIexpectedGND)
    if {[catch {
      if {$pmiMatch == 0 && [string first "tolerance" $entstr] != -1} {
        foreach sym $tolSymbols {
          set posSym [string first $sym $val]
          if {$posSym > 0} {break}
        }
        set valgnd [string trim [string range $val $posSym end]]

# check geometric tolerances
        if {[info exists nistPMIexpectedGND($epmiName)]} {
          set ok 0
          set pmiMatchGND [lsearch $nistPMIexpectedGND($epmiName) $valgnd]
          if {$pmiMatchGND != -1} {
            set ok 1

# check a different way with string equal because lsearch above doesn't always work
          } else {
            foreach item $nistPMIexpectedGND($epmiName) {
              incr pmiMatchGND
              if {[string equal $item $valgnd] == 1} {set ok 1; break}
            }
          }
          if {$ok} {
            if {$pmiMatchGND != -1} {
              set pmiMatch 0.99
              set pmiSim $pmiMatch
              #outputMsg " exact match ND  $val $pmiMatchGND $valType($val)" red
              if {$posSym > 0} {
                set pmiSimilar "Dimension does not match"
              } else {
                set pmiSimilar "Missing dimension association"
                if {$epmiName == "jpmi-trim"} {set pmiSimilar "Missing UF tolerance modifier"}
              }
              lappend nistPMIfound $val
              set nistPMIdeduct(tol) 1

# remove from expected list
              if {$debug} {outputMsg "remove5a [lindex $nistPMIexpectedGND($epmiName) $pmiMatchGND]" green}
              set nistPMIexpectedGND($epmiName) [lreplace $nistPMIexpectedGND($epmiName) $pmiMatchGND $pmiMatchGND]
              set pos -1
              foreach item $nistPMIexpected($epmiName) {
                incr pos
                if {[info exists nistPMImap($item)]} {
                  if {$debug} {outputMsg "map $valgnd / $nistPMImap($item)" green}
                  if {[string equal $valgnd $nistPMImap($item)] == 1} {break}
                }
              }
              if {$debug} {outputMsg "[llength $nistPMIexpected($epmiName)] $nistPMIexpected($epmiName)" blue; outputMsg "remove5b [lindex $nistPMIexpected($epmiName) $pos]" green}
              set nistPMIexpected($epmiName)   [lreplace $nistPMIexpected($epmiName)   $pos $pos]
              set nistPMIexpectedNX($epmiName) [lreplace $nistPMIexpectedNX($epmiName) $pos $pos]
              if {$debug} {outputMsg "[llength $nistPMIexpected($epmiName)] $nistPMIexpected($epmiName)" blue}
            }
          }
        }

# unexpected dimension association
        if {$pmiMatch == 0} {
          set pmiMatchGND [lsearch $nistPMIexpected($epmiName) $valgnd]
          if {$pmiMatchGND != -1 && [string first "\u2B69\u25CE" $val] == -1} {
            set pmiMatch 0.99
            set pmiSimilar "Unexpected dimension association"
            set nistPMIexpected($epmiName) [lreplace $nistPMIexpected($epmiName) $pmiMatchGND $pmiMatchGND]
            lappend nistPMIfound $val
            set nistPMIdeduct(tol) 1
          }
        }
      }
    } emsg]} {
      errorMsg "Error matching geometric tolerance without dimension: $emsg"
    }

# -------------------------------------------------------------------------------
# no match yet
    if {$pmiMatch == 0} {
      foreach pmi $nistPMIexpected($epmiName) {

# look for similar strings
        set look 0
        if {$valType($val) == $pmiType($pmi) && $val != "" && $pmiMatch < 0.9} {set look 1}
        if {$valType($val) == "datum_target" && $pmiType($pmi) == "placed_datum_target_feature" && $val != "" && $pmiMatch < 0.9} {set look 1}

        if {$look} {
          set ok 1

# check for bad dimensions
          if {$valType($val) == "dimensional_characteristic_representation"} {
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
              set okdim 1
              if {[string first $pmiUnicode(degree) $val] == -1 && [string first $pmiUnicode(degree) $pmi] != -1} {
                set okdim 0
              } elseif {[string first $pmiUnicode(plusminus) $val] == -1 && [string first $pmiUnicode(plusminus) $pmi] != -1} {
                set okdim 0
              }
              if {$okdim} {
                set diff [expr {[string length $pmi] - [string length $val]}]
                if {$diff <= 2 && $diff >= 0 && [string first $val $pmi] != -1} {
                  set pmiSim 0.95
                } elseif {[string is integer [string index $val 0]] || \
                          [string range $val 0 1] == [string range $pmi 0 1]} {
                  set pmiSim [stringSimilarity $val $pmi]
                }
              }

# tolerances
            } elseif {[string first "tolerance" $valType($val)] != -1} {

# missing all around
              if {[string first "\u232E" $pmi] == 0} {
                if {[string first $val [string range $pmi 2 end]] == 0} {
                  set pmiSimilar "Missing all around"
                  set pmiMatch [lsearch $nistPMIexpected($epmiName) $pmi]
                  set nistPMIexpected($epmiName) [lreplace $nistPMIexpected($epmiName) $pmiMatch $pmiMatch]
                  set pmiSim 0.99
                  set pmiMatch $pmiSim
                  lappend nistPMIfound $pmi
                }

# make sure tolerance datum features are the same
              } elseif {[string first "\u25BD" $val] != -1} {
                set valDF [string index $val [string first "\u25BD" $val]+2]
                set c2 [string first "\u25BD" $pmi]
                if {$c2 != -1} {
                  set pmiDF [string index $pmi $c2+2]
                  if {$valDF == $pmiDF} {set pmiSim [stringSimilarity $val $pmi]}
                }

              } elseif {[string first $val $pmi] != -1 || [string first $pmi $val] != -1} {
                set pmiSim 0.95
              } else {
                set tol $pmiUnicode([string range $valType($val) 0 [string last "_" $valType($val)]-1])
                set pmiSim [stringSimilarity $val $pmi]

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
            }

            if {$pmiSim < 0.6 && [string first "dimensional" $valType($val)] == -1} {
              if {[string first $val $pmi] != -1 || [string first $pmi $val] != -1 || $valType($val) == "flatness_tolerance"} {set pmiSim 0.6}
            }

# -------------------------------------------------------------------------------
# keep best match
            if {$pmiSim > $pmiMatch} {
              set okmatch 0
              set pmiMatch $pmiSim
              if {[string first "datum_target" $valType($val)] == -1 && [string first "dimension" $valType($val)] == -1} {
                if {$pmiSim >= 0.6} {set pmiSimilar $nistPMIactual($pmi); set okmatch 1}

# dimensions
              } elseif {[string first "dimension" $valType($val)] != -1} {
                if {$pmiSim >= 0.6} {set pmiSimilar $nistPMIactual($pmi); set bestPMI $pmi}

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
                            set okmatch 1
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
                set okmatch 1
              }
              if {$okmatch} {lappend nistPMIfound $pmi}
            }
          }
        }
      }
      if {[info exists bestPMI]} {
        lappend nistPMIfound $bestPMI
        unset bestPMI
      }
    }
  }

# more exceptions
  if {$pmiMatch == 1} {
    foreach idx {"Directed dimension" "Continuous feature" Independency} {
      if {[info exists pmiException($idx)]} {
        set pmiSimilar "$idx is ignored"
        set pmiMatch 0.99
      }
    }
  }
  catch {unset pmiException}

# -------------------------------------------------------------------------------
# exact matches, green
  if {$pmiMatch == 1} {
    [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(green)
    incr nistExpectedPMI(exact)
  } elseif {$pmiMatch == 0.99} {
    [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(greyel)
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

# no match red
  } else {
    [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(red)
    incr nistExpectedPMI(no)
  }

# add similar pmi
  if {[info exists pmiSimilar] && $pmiSimilar != ""} {
    $cells($spmiSumName) Item $spmiSumRow 4 "'$pmiSimilar"
    set range [$worksheet($spmiSumName) Range D$spmiSumRow]
    catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 1]}}
    incr nsimilar

# heading
    set heading [[$cells($spmiSumName) Item 3 4] Value]
    set exception 0
    foreach item {"match" "dimension" "ignored"} {if {[string first $item $pmiSimilar] != -1} {set exception 1}}
    if {$exception} {
      [[$worksheet($spmiSumName) Range D$spmiSumRow] Interior] Color $legendColor(litgray)
      if {$heading == ""} {set heading "Exception"}
      if {$heading == "Similar PMI"} {set heading "Similar PMI / Exception"}
    } else {
      [[$worksheet($spmiSumName) Range D$spmiSumRow] Interior] Color $legendColor(gray)
      if {$heading == ""} {set heading "Similar PMI"}
      if {$heading == "Exception"} {set heading "Similar PMI / Exception"}
    }
    if {$heading != ""} {$cells($spmiSumName) Item 3 4 $heading}

    if {$nsimilar == 1} {
      [$worksheet($spmiSumName) Range [cellRange -1 4]] ColumnWidth [expr 48]
      set cmnt ""
      if {[string first "Similar PMI" $heading] != -1} {
        append cmnt "Similar PMI (gray) is the best match of the Semantic PMI in column C, for Partial or Possible matches (blue and yellow), to the expected PMI in the NIST test case drawing."
      }
      if {[string first "Exception" $heading] != -1} {
        if {[string length $cmnt] > 0} {append cmnt "\n\n"}
        append cmnt "Exceptions (light gray) are items in the Semantic PMI in column C that are ignored when matching Expected PMI from the NIST test case drawing."
      }
      addCellComment $spmiSumName 3 4 $cmnt
      set range [$worksheet($spmiSumName) Range D3]
      [$range Font] Bold [expr 1]
      catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 2]}}
    }
  }

# border
  catch {[[[$worksheet($spmiSumName) Range C$spmiSumRow] Borders] Item [expr 9]] Weight [expr 1]}
}

# -------------------------------------------------------------------------------
proc nistPMICoverage {nf} {
  global cells epmi gen legendColor nistCoverageLegend nistCoverageStyle nistPMIexpected nistName
  global opt pmiElementsMaxRows spmiCoverages spmiCoverageWS totalPMIrows usedPMIrows worksheet

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

# need better fix for these that conflict with other PMI elements
          if {[string first "free_state_condition" $ttyp]   != -1} {set ok 0}
          if {[string first "united_feature_of_size" $ttyp] != -1} {set ok 0}

          if {$ok} {
            set ci $coverage($item)
            catch {set ci [expr {int($ci)}]}

# check tolerance zone diameter vs. within a cylinder
            set skip 0
            if {$item == "tolerance zone diameter"          && $tval == 0 && [[$cells($spmiCoverageWS) Item 20 2] Value] != ""} {set skip 1}
            if {$item == "tolerance zone within a cylinder" && $tval == 0 && [[$cells($spmiCoverageWS) Item 19 2] Value] != ""} {set skip 1}

# skip if view not generated
            if {!$gen(View)} {
              if {$item == "section views" && !$opt(viewPart)} {set skip 1}
              if {$item == "saved views"   && !$opt(viewPMI)}  {set skip 1}
            }

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

# some exceptions for JPMI files
                if {[string first "jpmi" $nistName] == -1 || \
                  ([string first "dimension association" $item] == -1 && $item != "repetitive dimensions")} {
                  set str "'$tval/[expr {int($ci1)}]"
                  $cells($spmiCoverageWS) Item $r 2 $str
                  [[$worksheet($spmiCoverageWS) Range B$r] Interior] Color $legendColor($clr)
                  [$worksheet($spmiCoverageWS) Range B$r] NumberFormat "@"
                  set nistCoverageLegend 1
                  lappend nistCoverageStyle "$r $nf $clr $str"
                }

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

# summarize Semantic PMI Summary on Semantic PMI Coverage worksheet
  set name $nistName
  if {[info exists epmi]} {if {$epmi != ""} {set name $epmi}}
  if {[info exists nistPMIexpected($name)]} {nistAddExpectedPMIPercent $nf $name}
}

# -------------------------------------------------------------------------------
# summarize Semantic PMI Summary with percentages
proc nistAddExpectedPMIPercent {nf name} {
  global cells cells1 legendColor lenfilelist nistCoverageStyle pmiElementsMaxRows spmiCoverageWS worksheet worksheet1
  global nistExpectedPMI nistPMIdeduct nistPMIexpected nistPMIfound

# compute missing and total expected PMI
  set nistExpectedPMI(missing) [llength [lindex [intersect3 $nistPMIexpected($name) $nistPMIfound] 0]]
  set nistExpectedPMI(total) 0
  foreach idx {exact partial possible missing} {
    if {[info exists nistExpectedPMI($idx)]} {incr nistExpectedPMI(total) $nistExpectedPMI($idx)}
  }

  set r [expr {$pmiElementsMaxRows+4}]
  $cells($spmiCoverageWS) Item $r 1 "Expected PMI[format "%c" 10]  (See Semantic PMI Summary worksheet)"
  $cells($spmiCoverageWS) Item $r 2 "%"
  set range [$worksheet($spmiCoverageWS) Range A$r:B$r]
  [$range Font] Bold [expr 1]
  catch {foreach i {8 10 11} {[[$range Borders] Item $i] Weight [expr 2]}}
  set range [$worksheet($spmiCoverageWS) Range B$r]
  $range HorizontalAlignment [expr -4108]
  addCellComment $spmiCoverageWS $r 2 "The color-coded PERCENTAGES (not absolute numbers) are based on the Total PMI of the number Exact, Partial, Possible, and Missing matches from column B on the Semantic PMI Summary worksheet.  The percentages for all matches should total 100.\n\n'Missing match' is based on Missing PMI that would appear below the color legend.  'No match' is based on the number of Semantic PMI that appear in red above the legend.\n\nCoverage Analysis is only based on individual PMI elements.  The Semantic PMI Summary is based on the complete Feature Control Frame and provides a better understanding of the PMI.  The Coverage Analysis might show that there is an Exact match (all green above) for all of the PMI elements, however, the PMI Summary might show less than Exact matches.\n\nSee Help > Analyzer > NIST CAD Models\nSee Help > User Guide (section 6.6.2.1)"

# for multiple files, add more formatting
  if {[info exists lenfilelist]} {
    if {$lenfilelist > 1 && ![info exists nistPMIexpectedFormat]} {
      incr nistPMIexpectedFormat
      set r1 [expr {$pmiElementsMaxRows+4}]
      $cells1($spmiCoverageWS) Item $r1 1 "Expected PMI[format "%c" 10]  (% from Semantic PMI Summary worksheets)"
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
      if {[info exists nistExpectedPMI($idx)]} {
        if {$idx != "no"} {
          set pct [trimNum [expr {(100.*$nistExpectedPMI($idx))/$nistExpectedPMI(total)}] 0]
        } else {
          set pct $nistExpectedPMI($idx)
        }
      }
    } else {
      set pct $nistExpectedPMI($idx)
    }
    if {$pct == 100} {
      if {[info exists nistPMIdeduct(dim)]} {set pct [expr {$pct-1}]}
      if {[info exists nistPMIdeduct(tol)]} {set pct [expr {$pct-1}]}
    }
    catch {unset nistPMIdeduct}
    $cells($spmiCoverageWS) Item $r 1 "[string totitle $idx] match"
    $cells($spmiCoverageWS) Item $r 2 $pct
    if {$idx == "exact"} {outputMsg " Expected PMI: $pct\%"}

# formatting
    set range [$worksheet($spmiCoverageWS) Range B$r]
    catch {foreach i {7 10} {[[$range Borders] Item $i] Weight [expr 2]}}

# for multiple files, add more formatting
    if {[info exists lenfilelist]} {
      if {$lenfilelist > 1 && $nistPMIexpectedFormat >= 1} {
        incr r1
        $cells1($spmiCoverageWS) Item $r1 1 "[string totitle $idx] match"
        incr nistPMIexpectedFormat
      }
    }

# color-code percentages
    set clr "white"
    if {$idx != "total" && $idx != "no"} {
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
proc nistPMISummaryFormat {name} {
  global cells legendColor pmiType spmiSumName spmiSumRow worksheet
  global nistPMIactual nistPMIexpected nistPMIexpectedNX nistPMIexpectedGND nistPMIfound nistPMImap

  set r [incr spmiSumRow]

# legend
  set n 0
  set legend {{"Expected PMI" ""} {"See Help > Analyzer > NIST CAD Models" ""} {"Exact match" "green"} {"Exact match with Exceptions" "greyel"} {"Partial match" "cyan"} {"Possible match" "yellow"} {"No match" "red"}}
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
  set pmiMissing [lindex [intersect3 $nistPMIexpected($name) $nistPMIfound] 0]
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

# unset variables
  foreach var {nistPMIactual nistPMIexpected nistPMIexpectedNX nistPMIexpectedGND nistPMIfound nistPMImap} {if {[info exists $var]} {unset -- $var}}
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
  set legend {{"Values as Compared to NIST Test Case Drawing" ""} {"See Help > Analyzer > NIST CAD Models" ""} \
              {"More than expected" "cyan"} {"Exact match" "green"} {"Less than expected (upper third)" "yelgre"} \
              {"Less than expected (middle third)" "yellow"} {"Less than expected (lower third)" "orange"} \
              {"None (0/n)" "red"} {"Unexpected (n/0)" "magenta"} {"Not checked" ""}}
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
  global cells excel localName nistModelPictures mytemp nistName worksheet

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
            if {[string first "Semantic" $ent] != -1} {
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
            if {$str != ""} {
              regsub -all "_" $str "-" name
              if {[string first "tc-" $name] != -1} {
                set name "nist-cad-model-[string range $name 5 end]"
                catch {$cells($ent) Item 3 5 "Test Case Drawing"}
                set range [$worksheet($ent) Range E3:M3]
                $range MergeCells [expr 1]
                set range [$worksheet($ent) Range "E3"]
                [$worksheet($ent) Hyperlinks] Add $range [join "https://www.nist.gov/document/$name"] [join ""] [join "Link to Test Case Drawing (PDF)"]
                incr nlink
              }
            }
          }
        }
      }
    }
  } emsg]} {
    errorMsg "Error adding Picture to PMI Summary or Coverage worksheet.\n  $emsg"
  }
}

# -------------------------------------------------------------------------------------------------
proc nistGetName {} {
  global gen localName opt resetRound

  set nistName ""
  set filePrefix {}
  set prefixes {}
  for {set i 4} {$i < 8} {incr i} {lappend prefixes "sp$i"}
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

# check for JPMI test case
  if {[string first "jpmi" $ftail] == 0} {
    if {[string first "jpmi_" $ftail] == 0} {regsub "jpmi_" $ftail "jpmi-" ftail}
    if {[string first "gears" $ftail] == -1 && [string first "gear" $ftail] != -1} {regsub "gear" $ftail "gears" ftail}
    foreach tc {jpmi-gears jpmi-housing jpmi-knuckle jpmi-trim} {if {[string first $tc $ftail] == 0} {set nistName $tc}}

# check for a NIST CTC, FTC, STC, HTC, PDC
  } else {
    set testCase ""
    set ok  0
    set ok1 0

    if {[lsearch $filePrefix [string range $ftail 0 $c]] != -1 || [string first "htc" $ftail] != -1 || [string first "pdc" $ftail] != -1 || \
        [string first "ctc" $ftail] != -1 || [string first "ftc" $ftail] != -1 || [string first "stc" $ftail] != -1 || \
        ([string first "nist" $ftail] != -1 && [string first "pdi" $ftail] == -1)} {
      if {[lsearch $filePrefix [string range $ftail 0 $c]] != -1} {set ftail [string range $ftail $c+1 end]}

      set tmp "nist_"
      foreach item {ctc ftc stc htc pdc} {
        if {[string first $item $ftail] != -1} {
          append tmp "$item\_"
          set testCase $item
        }
      }
      if {$testCase == "htc"} {set nistName "nist_htc"; return $nistName}
      if {$testCase == "pdc"} {set nistName "nist_pdc"; return $nistName}

# find nist_ctc_01 directly
      if {$testCase != ""} {
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
      if {!$ok1 && $testCase != "htc" && $testCase != "pdc"} {
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
                    if {$testCase != ""} {
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
  }

# check required rounding for some ftc, stc, and jpmi models
  catch {unset resetRound}
  catch {
    if {$opt(PMISEM) && $gen(Excel)} {
      if {$opt(PMISEMRND) && ($nistName == "nist_ftc_06" || $nistName == "nist_stc_06")} {
        set resetRound $opt(PMISEMRND)
        set opt(PMISEMRND) 0
      } elseif {!$opt(PMISEMRND) && ($nistName == "nist_ftc_07" || $nistName == "nist_ftc_08" || \
                $nistName == "nist_ftc_11" || $nistName == "nist_stc_07" || $nistName == "nist_stc_08" || \
                [string first "jpmi" $nistName] == 0)} {
        set resetRound $opt(PMISEMRND)
        set opt(PMISEMRND) 1
      }
    }
  }
  return $nistName
}

# -------------------------------------------------------------------------------
proc pmiRemoveZeros {pmi} {
  global pmiUnicode

# line feeds
  set pmi1 [split $pmi \n]
  if {[string first \n $pmi] != -1} {
    if {![string is double [lindex $pmi1 0]] || ![string is double [lindex $pmi1 1]]} {
      regsub -all \n $pmi " " pmi

# special case for ISO limit dimensions where they are stacked, change to ASME representation
    } else {
      set pmi "[lindex $pmi1 1]-[lindex $pmi1 0]"
      if {[llength $pmi1] > 2} {append pmi " [join [lrange $pmi1 2 end]]"}
    }
  }

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
          if {[string first $pmiUnicode(degree) $spmi] == -1} {set spmi "[string range $spmi 0 end-1]$pmiUnicode(degree) "}
          if {[string first " \[" $spmi] == 0 && [string first "\]" $spmi] == -1} {set spmi "[string range $spmi 0 end-1]\] "}
          if {[string first " \(" $spmi] == 0 && [string first "\)" $spmi] == -1} {set spmi "[string range $spmi 0 end-1]\) "}
        }

# reference dimension
        set deg 0
        if {[string first "$pmiUnicode(degree)\]" $spmi] != -1} {
          set deg 1
          regsub $pmiUnicode(degree) $spmi "" spmi
        }
        if {[string first "0\]" $spmi] != -1} {for {set j 0} {$j < 4} {incr j} {regsub {0\]} $spmi "\]" spmi}}
        if {[string first ".\]" $spmi] != -1} {regsub {.\]} $spmi "\]" spmi}
        if {[string first "\[0" $spmi]  == 1} {regsub {\[0} $spmi "\[" spmi}
        if {$deg} {regsub {\]} $spmi "$pmiUnicode(degree)\]" spmi}

# basic dimension
        set deg 0
        if {[string first "$pmiUnicode(degree)\)" $spmi] != -1} {
          set deg 1
          regsub $pmiUnicode(degree) $spmi "" spmi
        }
        if {[string first "0)" $spmi] != -1} {for {set j 0} {$j < 3} {incr j} {regsub {0\)} $spmi "\)" spmi}}
        if {[string first ".)" $spmi] != -1} {regsub {\.\)} $spmi ")" spmi}
        if {[string first "\(0" $spmi] == 1} {regsub {\(0}  $spmi "(" spmi}
        if {$deg} {regsub {\)} $spmi "$pmiUnicode(degree)\)" spmi}

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

# ctc 1 specific case
  set c1 [string first "80-3" $pmi]
  if {$c1 != -1} {set pmi [string range $pmi 0 $c1][string range $pmi $c1+2 end]}

  return $pmi
}

#-------------------------------------------------------------------------------
# From https://wiki.tcl-lang.org/page/similarity
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
