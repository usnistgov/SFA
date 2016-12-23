#-------------------------------------------------------------------------------
# version number

proc getVersion {} {
  set app_version 1.95
  return $app_version
}

# -------------------------------------------------------------------------------
# dt = 1 for dimtol
proc getAssocGeom {entDef {dt 0}} {
  global assocGeom entCount recPracNames
  
  set entDefType [$entDef Type]
  #outputMsg "getGeom $entDefType [$entDef P21ID]" green

  if {[catch {
    if {$entDefType == "shape_aspect" || \
      ([string first "datum" $entDefType] != -1 && [string first "_and_" $entDefType] == -1)} {

# add shape_aspect to AG for dimtol
      if {$dt && ($entDefType == "shape_aspect" || $entDefType == "datum_feature")} {
        set type [appendAssocGeom $entDef A]
      }

# find datum_feature for datum
      if {$entDefType == "datum"} {
        set e0s [$entDef GetUsedIn [string trim shape_aspect_relationship] [string trim related_shape_aspect]]
        ::tcom::foreach e0 $e0s {
          ::tcom::foreach a0 [$e0 Attributes] {
            if {[$a0 Name] == "relating_shape_aspect"} {
              if {[[$a0 Value] Type] == "datum_feature"} {set entDef [$a0 Value]}
            }
          }
        }
      }
      
# find AF for SA with GISU or IIRU
      foreach usage {geometric_item_specific_usage item_identified_representation_usage} {
        set e1s [$entDef GetUsedIn [string trim $usage] [string trim definition]]
        ::tcom::foreach e1 $e1s {
          ::tcom::foreach a1 [$e1 Attributes] {
            if {[$a1 Name] == "identified_item"} {
              if {[catch {
                set type [appendAssocGeom [$a1 Value] B]
                if {$type == "advanced_face"} {getFaceGeom [$a1 Value] B}
              } emsg1]} {
                ::tcom::foreach e2 [$a1 Value] {
                  set type [appendAssocGeom $e2 C]
                  if {$type == "advanced_face"} {getFaceGeom $e2 C}
                }
              }
            }
          }
        }
      }
      
# look at composite_shape_aspect to find SAs
    } else {
      #outputMsg " $entDefType [$entDef P21ID]" red
      set type [appendAssocGeom $entDef D]
      set e0s [$entDef GetUsedIn [string trim shape_aspect_relationship] [string trim relating_shape_aspect]]
      ::tcom::foreach e0 $e0s {
        ::tcom::foreach a0 [$e0 Attributes] {
          if {[$a0 Name] == "related_shape_aspect"} {
            set type [appendAssocGeom [$a0 Value] E]
            if {$type == "advanced_face"} {getFaceGeom [$a0 Value] E}

            set a0val {}
            if {[[$a0 Value] Type] == "composite_shape_aspect"} {
              set e1s [[$a0 Value] GetUsedIn [string trim shape_aspect_relationship] [string trim relating_shape_aspect]]
              ::tcom::foreach e1 $e1s {
                ::tcom::foreach a1 [$e1 Attributes] {
                  if {[$a1 Name] == "related_shape_aspect"} {
                    lappend a0val [$a1 Value]
                    set type [appendAssocGeom [$a1 Value] F]
                  }
                }
              }
            } else {
              lappend a0val [$a0 Value]
            }

# find AF for SA with GISU
            foreach val $a0val {
              set e1s [$val GetUsedIn [string trim geometric_item_specific_usage] [string trim definition]]
              ::tcom::foreach e1 $e1s {
                ::tcom::foreach a1 [$e1 Attributes] {
                  if {[$a1 Name] == "identified_item"} {
                    set type [appendAssocGeom [$a1 Value] G]
                    if {$type == "advanced_face"} {getFaceGeom [$a1 Value] G}
                  }
                }
              }
            }
          }
        }
      }
      
# check all around
      if {$entDefType == "all_around_shape_aspect"} {
        if {[llength $assocGeom($type)] == 1} {
          #outputMsg " assocGeom $type $assocGeom($type) [llength $assocGeom($type)]" blue
          if {$type == "advanced_face"} {
            errorMsg "Syntax Error: 'shape_aspect relationship' relates '$entDefType' to only one 'shape_aspect'.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 6.4)"
          } elseif {$type == $entDefType} {
            errorMsg "Syntax Error: Missing 'shape_aspect relationship' relating '$entDefType' to 'shape_aspect'.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 6.4)"
            unset assocGeom($type)
          }
        }
      }
    }
  } emsg]} {
    errorMsg "ERROR adding Associated Geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
proc appendAssocGeom {ent {id ""}} {
  global assocGeom
  
  set p21id [$ent P21ID]
  set type  [$ent Type]
  #outputMsg " $type $p21id $id"
  
  if {[string first "annotation" $type] == -1} {
    if {![info exists assocGeom($type)]} {
      lappend assocGeom($type) $p21id
    } elseif {[lsearch $assocGeom($type) $p21id] == -1} {
      lappend assocGeom($type) $p21id
    }
  }
  return $type
}

# -------------------------------------------------------------------------------
proc getFaceGeom {a0 {id ""}} {
  global assocGeom
  
  ::tcom::foreach a1 [$a0 Attributes] {
    if {[$a1 Name] == "face_geometry"} {
      set p21id [[$a1 Value] P21ID]
      set type  [[$a1 Value] Type]
      #outputMsg "  $type $p21id $id"
      
      if {![info exists assocGeom($type)]} {
        lappend assocGeom($type) $p21id
      } elseif {[lsearch $assocGeom($type) $p21id] == -1} {
        lappend assocGeom($type) $p21id
      }
    }
  }
}

# -------------------------------------------------------------------------------
proc reportAssocGeom {{type 1}} {
  global assocGeom
  
  set str ""
  foreach item [array names assocGeom] {
    if {[string first "shape_aspect" $item] == -1 && [string first "centre" $item] == -1 && [string first "datum" $item] == -1 && $item != "advanced_face"} {
      if {[string length $str] > 0} {append str [format "%c" 10]}
      append str "([llength $assocGeom($item)]) $item [lsort -integer $assocGeom($item)]"
    }
  }
  foreach item [array names assocGeom] {
    if {$item == "advanced_face"} {
      if {[string length $str] > 0} {append str [format "%c" 10]}
      append str "([llength $assocGeom($item)]) $item [lsort -integer $assocGeom($item)]"
    }
  }
  if {[string length $str] == 0 && $type} {
    errorMsg "Syntax Error: Missing Associated Geometry for shape_aspect through GISU or IIRU."
  }
  foreach item [array names assocGeom] {
    if {[string first "shape_aspect" $item] != -1 || [string first "centre" $item] != -1 || [string first "datum_feature" $item] != -1} {
      if {[string length $str] > 0} {append str [format "%c" 10]}
      append str "([llength $assocGeom($item)]) $item [lsort -integer $assocGeom($item)]"
    }
  }
  return $str
}

# -------------------------------------------------------------------------------
# Semantic PMI summary worksheet
proc spmiSummary {} {
  global cells entName excelVersion localName row sheetLast spmiSumName spmiSumRow thisEntType worksheet worksheets xlFileName
  global nistName pmiExpected wdir mytemp legendColor pmiUnicode pmiFound pmiModifiers pmiActual recPracNames tolNames pmiType valType
  global nsimilar pmiMaster
  
# first time through, start worksheet
  if {$spmiSumRow == 1} {
    set spmiSumName "PMI Representation Summary"
    set worksheet($spmiSumName) [$worksheets Add [::tcom::na] $sheetLast]
    $worksheet($spmiSumName) Activate
    $worksheet($spmiSumName) Name $spmiSumName
    set cells($spmiSumName) [$worksheet($spmiSumName) Cells]
    set wsCount [$worksheets Count]
    [$worksheets Item [expr $wsCount]] -namedarg Move Before [$worksheets Item [expr 3]]

    for {set i 2} {$i <= 3} {incr i} {[$worksheet($spmiSumName) Range [cellRange -1 $i]] ColumnWidth [expr 48]}
    for {set i 1} {$i <= 4} {incr i} {[$worksheet($spmiSumName) Range [cellRange -1 $i]] VerticalAlignment [expr -4160]}

    $cells($spmiSumName) Item $spmiSumRow 2 [file tail $localName]
    incr spmiSumRow 2
    $cells($spmiSumName) Item $spmiSumRow 1 "ID"
    $cells($spmiSumName) Item $spmiSumRow 2 "Entity"
    $cells($spmiSumName) Item $spmiSumRow 3 "PMI Representation"
    set range [$worksheet($spmiSumName) Range [cellRange 1 1] [cellRange 3 3]]
    [$range Font] Bold [expr 1]
    set range [$worksheet($spmiSumName) Range [cellRange 3 1] [cellRange 3 3]]
    if {$excelVersion >= 12} {
      [[$range Borders] Item [expr 8]] Weight [expr 2]
      [[$range Borders] Item [expr 9]] Weight [expr 2]
    }
    incr spmiSumRow
  
    $cells($spmiSumName) Item 1 3 "See CAx-IF Recommended Practice for $recPracNames(pmi242)"
    set range [$worksheet($spmiSumName) Range C1:K1]
    $range MergeCells [expr 1]
    set anchor [$worksheet($spmiSumName) Range C1]
    [$worksheet($spmiSumName) Hyperlinks] Add $anchor [join "https://www.cax-if.org/joint_testing_info.html#recpracs"] [join ""] [join "Link to CAx-IF Recommended Practices"]
    
    outputMsg " Adding PMI Representation Summary worksheet" green

# add pictures
    pmiAddModelPictures $spmiSumName
    [$worksheet($spmiSumName) Range "A1"] Select
  
# get expected PMI values from pmiMaster
    set nsimilar 0
    if {[info exists pmiMaster($nistName)]} {
      catch {unset pmiExpected($nistName)}
      
# read master PMI values, remove leading and trailing zeros, other stuff, add to pmiExpected
      foreach item $pmiMaster($nistName) {
        set c1 [string first "\\" $item]
        set typ [string range $item 0 $c1-1]
        set pmi [string range $item $c1+1 end]
        set newpmi [pmiRemoveZeros $pmi]
        lappend pmiExpected($nistName) $newpmi
        if {[string first "tolerance" $typ] != -1} {
          foreach nam $tolNames {if {[string first $nam $typ] != -1} {set pmiType($newpmi) $nam}}
        } else {
          set pmiType($newpmi) $typ
        }
        set pmiActual($newpmi) $pmi
      }
    }
    set pmiFound {}
  }
 
# add to PMI summary worksheet
  set hlink [$worksheet($spmiSumName) Hyperlinks]
  for {set i 3} {$i <= $row($thisEntType)} {incr i} {
    if {$thisEntType != "datum_reference_compartment" && $thisEntType != "datum_reference_element" && \
        $thisEntType != "datum_reference_modifier_with_value" && [string first "datum_feature" $thisEntType] == -1} {

# which entities and columns to summarize
      if {$i == 3} {
        for {set j 1} {$j < 30} {incr j} {
          set val [[$cells($thisEntType) Item 3 $j] Value]
          if {[string first "Datum Reference Frame" $val] != -1 || \
              $val == "GD&T[format "%c" 10]Annotation" || \
              $val == "Dimensional[format "%c" 10]Tolerance" || \
              [string first "Datum Target" $val] == 0 || \
              ($thisEntType == "datum_reference" && [string first "reference" $val] != -1) || \
              ($thisEntType == "referenced_modified_datum" && [string first "datum" $val] != -1)} {set pmiCol $j}
        }

# values
      } else {
        if {[info exists pmiCol]} {
          $cells($spmiSumName) Item $spmiSumRow 1 [[$cells($thisEntType) Item $i 1] Value]
          if {[string first "_and_" $thisEntType] == -1} {
            set entstr $thisEntType
          } else {
            regsub -all "_and_" $thisEntType ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" entstr
            set entstr "($entstr)"
          }
          $cells($spmiSumName) Item $spmiSumRow 2 $entstr

          set val [[$cells($thisEntType) Item $i $pmiCol] Value]
          $cells($spmiSumName) Item $spmiSumRow 3 "'$val"
          set cellval $val

# check actual vs. expected PMI for NIST files
          if {[info exists pmiExpected($nistName)]} {

# modify (composite ..) from value to just (composite)
            set c1 [string first "(composite" $val]
            if {$c1 > 0} {
              set val [string range $val 0 $c1+9]
              append val ")"
            }

# remove (oriented)
            set c1 [string first "(oriented)" $val]
            if {$c1 > 0} {set val [string range $val 0 $c1-2]}

# remove zeros from val
            #outputMsg $val red
            set val [pmiRemoveZeros $val]
             #outputMsg $val green
            if {[string first "tolerance" $entstr] != -1} {
              foreach nam $tolNames {if {[string first $nam $entstr] != -1} {set valType($val) $nam}}
            } else {
              set valType($val) $entstr
            }

# search for PMI in pmiExpected list
            set pmiMatch [lsearch $pmiExpected($nistName) $val]
            #outputMsg "$val\n $pmiMatch $valType($val)" green
            #outputMsg $pmiExpected($nistName)

# found in list, remove from pmiExpected
            if {$pmiMatch != -1} {
              #outputMsg "$pmiMatch $val"
              [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(green)
              set pmiExpected($nistName) [lreplace $pmiExpected($nistName) $pmiMatch $pmiMatch]
              lappend pmiFound $val

# not found
            } else {
              set pmiMatch 0
              set pmiMissing ""
              set pmiSimilar ""

# check each value in pmiExpected
              foreach pmi $pmiExpected($nistName) {
                
# simple match, remove from pmiExpected
                if {$val == $pmi && $pmiMatch != 1} {
                  set pmiMatch 1
                  set pos [lsearch $pmiExpected($nistName) $pmi]
                  set pmiExpected($nistName) [lreplace $pmiExpected($nistName) $pos $pos]
                  lappend pmiFound $pmi
                  #outputMsg "$val\n $pmiMatch $valType($val)" blue
                }
              }

# no match yet
              if {$pmiMatch == 0} {
                foreach pmi $pmiExpected($nistName) {

# look for similar strings
                  if {$valType($val) == $pmiType($pmi) && $val != "" && $pmiMatch < 0.9} {
                    set ok 1
 
# check for bad dimensions
                    if {$valType($val) == "dimensional_characteristic_representation"} {
                      #outputMsg "A$val\A [string first "-" $val] [string first "$pmiUnicode(diameter) " $val] [string first "$pmiUnicode(plusminus)" $val]"
                      if {[string first "-" $val] == 0 || [string first "$pmiUnicode(diameter) " $val] == 0 || \
                          [string first "$pmiUnicode(plusminus)" $val] == 0} {set ok 0}
                    }

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
                          #outputMsg "$val $pmiSim aa"
                        } elseif {[string is integer [string index $val 0]] || \
                                  [string range $val 0 1] == [string range $pmi 0 1]} {
                          set pmiSim [stringSimilarity $val $pmi]
                          #outputMsg "$val $pmiSim bb"
                        }

# tolerances
                      } elseif {[string first "tolerance" $valType($val)] != -1} {
                        if {[string first $val $pmi] != -1 || [string first $pmi $val] != -1} {
                          set pmiSim 0.95
                          #outputMsg "$val $pmiSim cc"
                        } else {
                          set tol $pmiUnicode([string range $valType($val) 0 [string last "_" $valType($val)]-1])
                          set pmiSim [stringSimilarity $val $pmi]
                          #outputMsg "$val $pmiSim dd"
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
                        if {[string first $val $pmi] != -1 || [string first $pmi $val] != -1 || \
                            $valType($val) == "flatness_tolerance"} {set pmiSim 0.6}
                      }
                    
# keep best match
                      if {$pmiSim > $pmiMatch} {
                        #outputMsg "$pmiSim / $val / $pmi / [string first $val $pmi]" red
                        set pmiMatch $pmiSim
                        if {[string first "datum_target" $valType($val)] == -1} {
                          if {$pmiSim >= 0.6} {
                            set pmiSimilar $pmiActual($pmi)
                            #append pmiSimilar "[format "%c" 10](Similarity: [string range $pmiMatch 0 4])"
                          }

# handle datum targets differently
                        } else {
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

# perfect match, green
              if {$pmiMatch == 1} {
                [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(green) 

# partial and possible match, cyan and yellow
              } elseif {$pmiMatch >= 0.6} {
                if {$pmiMatch >= 0.9} {
                  [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(cyan)
                } else {
                  [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(yellow)
                }

# add similar pmi
                if {[info exists pmiSimilar] && $pmiSimilar != ""} {
                  $cells($spmiSumName) Item $spmiSumRow 4 "'$pmiSimilar"
                  [[$worksheet($spmiSumName) Range D$spmiSumRow] Interior] Color $legendColor(gray)
                  if {$excelVersion >= 12} {
                    [[[$worksheet($spmiSumName) Range D$spmiSumRow] Borders] Item [expr 8]] Weight [expr 1]
                    [[[$worksheet($spmiSumName) Range D$spmiSumRow] Borders] Item [expr 9]] Weight [expr 1]
                  }
                  incr nsimilar

                  if {$nsimilar == 1} {
                    [$worksheet($spmiSumName) Range [cellRange -1 4]] ColumnWidth [expr 48]
                    $cells($spmiSumName) Item 3 4 "Similar PMI"
                    set range [$worksheet($spmiSumName) Range D3]
                    [$range Font] Bold [expr 1]
                    if {$excelVersion >= 12} {
                      [[$range Borders] Item [expr 8]] Weight [expr 2]
                      [[$range Borders] Item [expr 9]] Weight [expr 2]
                    }
                  }
                }
                lappend pmiFound $pf

# no match red                
              } else {
                [[$worksheet($spmiSumName) Range C$spmiSumRow] Interior] Color $legendColor(red)
              }
            }

# border
            if {$excelVersion >= 12} {
              [[[$worksheet($spmiSumName) Range C$spmiSumRow] Borders] Item [expr 9]] Weight [expr 1]
            }
          }

# link back to worksheets
          set anchor [$worksheet($spmiSumName) Range "B$spmiSumRow"]
          set hlsheet $thisEntType
          if {[string length $thisEntType] > 31} {
            foreach item [array names entName] {
              if {$entName($item) == $thisEntType} {set hlsheet $item}
            }
          }
          $hlink Add $anchor $xlFileName "$hlsheet!A1" "Go to $thisEntType"
          incr spmiSumRow
        } else {
          errorMsg "Cannot find PMI on $thisEntType for PMI Representation Summary worksheet"
        }
      }
    }
  }
}

# -------------------------------------------------------------------------------
proc spmiSummaryFormat {} {
  global cells excelVersion legendColor pmiExpected pmiFound nistName pmiActual spmiSumRow spmiSumName pmiType worksheet

  #[$worksheet($spmiSumName) Columns] AutoFit
  #[$worksheet($spmiSumName) Rows] AutoFit

# for NIST CAD models
  if {$nistName != "" && [info exists pmiExpected($nistName)]} {
      set r [incr spmiSumRow]

# legend
      set n 0
      set legend {{"Expected PMI" ""} {"See Help > NIST CAD Models" ""} {"Exact match" "green"} {"Partial match" "cyan"} {"Possible match" "yellow"} {"No match" "red"}}
      foreach item $legend {
        set str [lindex $item 0]
        $cells($spmiSumName) Item $r 3 $str

        set range [$worksheet($spmiSumName) Range [cellRange $r 3]]
        [$range Font] Bold [expr 1]

        set color [lindex $item 1]
        if {$color != ""} {[$range Interior] Color $legendColor($color)}

        if {$excelVersion >= 12} {
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

# add missing pmi
    set pmiMissing [lindex [intersect3 $pmiExpected($nistName) $pmiFound] 0]
    if {[llength $pmiMissing] > 0} {
      incr r
      $cells($spmiSumName) Item $r 2 "Entity Type"
      $cells($spmiSumName) Item $r 3 "Missing PMI"
      set range [$worksheet($spmiSumName) Range [cellRange $r 2]  [cellRange $r 3]]
      [$range Font] Bold [expr 1]
      if {$excelVersion >= 12} {
        [[$range Borders] Item [expr 8]] Weight [expr 2]
        [[$range Borders] Item [expr 9]] Weight [expr 2]
      }
      foreach item $pmiMissing {
        incr r
        $cells($spmiSumName) Item $r 2 $pmiType($item)
        $cells($spmiSumName) Item $r 3 "'$pmiActual($item)"
        [[$worksheet($spmiSumName) Range [cellRange $r 3]] Interior] Color $legendColor(red)
        if {$excelVersion >= 12} {[[[$worksheet($spmiSumName) Range [cellRange $r 3]] Borders] Item [expr 9]] Weight [expr 1]}
      }
    }
  }

  [$worksheet($spmiSumName) Columns] AutoFit
  [$worksheet($spmiSumName) Rows] AutoFit
  unset spmiSumName
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
      #outputMsg A$spmi\A
      if {[string first "." $spmi] != -1} {

# trailing zeros
        if {[string first $pmiUnicode(degree) $spmi] == -1} {
          for {set j 0} {$j < 4} {incr j} {if {[string first "0 " $spmi] != -1} {regsub -all "0 " $spmi " " spmi}}

# leading zero
          regsub -all " 0" $spmi " " spmi
          if {[string first $pmiUnicode(diameter) $spmi] != -1} {regsub -all "$pmiUnicode(diameter)0" $spmi $pmiUnicode(diameter) spmi}

# trailing .
          if {[string first ". " $spmi] != -1} {regsub {\. } $spmi " " spmi}

# simlar thing for degrees
        } else {
          set spmi "[string range $spmi 0 end-2] "
          for {set j 0} {$j < 4} {incr j} {if {[string first "0 " $spmi] != -1} {regsub -all "0 " $spmi " " spmi}}
          regsub -all " 0" $spmi " " spmi
          if {[string first ". " $spmi] != -1} {regsub {\. } $spmi " " spmi}
          set spmi "[string range $spmi 0 end-1]$pmiUnicode(degree) "
        }
# reference dimension
        if {[string first "0\]" $spmi] != -1} {for {set j 0} {$j < 4} {incr j} {regsub {0\]} $spmi "\]" spmi}}
        #outputMsg $spmi
        if {[string first ".\]" $spmi] != -1} {regsub {.\]} $spmi "\]" spmi}

# basic dimension
        if {[string first "0)" $spmi] != -1} {for {set j 0} {$j < 3} {incr j} {regsub {0\)} $spmi "\)" spmi}}
        if {[string first ".)" $spmi] != -1} {regsub {\.\)} $spmi ")" spmi}

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

  #outputMsg $pmi green
  return $pmi
}

# -------------------------------------------------------------------------------
# get expected PMI from spreadsheets, (called from sfa-gen.tcl)
proc spmiGetPMI {} {
  global wdir mytemp spmiCoverages pmiMaster nistName nistVersion

  if {[catch {
    if {![info exists spmiCoverages]} {
      
# first mount NIST zip file with images and expected PMI
      if {$nistVersion} {
        set fn [file join $wdir SFA-NIST-files.zip]
        if {[file exists $fn]} {
          file copy -force [file join $wdir $fn] $mytemp
          vfs::zip::Mount [file join $mytemp $fn] NIST
        }
      }

# get PMI coverage
      set fn "SFA-PMI-NIST-coverage.xlsx"
      if {[file exists NIST/$fn]} {file copy -force NIST/$fn [file join $mytemp $fn]}
      set fname [file nativename [file join $mytemp $fn]]

      if {[file exists $fname]} {
        set pid1 [twapi::get_process_ids -name "EXCEL.EXE"]
        set excel2 [::tcom::ref createobject Excel.Application]
        set pid2 [lindex [intersect3 $pid1 [twapi::get_process_ids -name "EXCEL.EXE"]] 2]
    
        $excel2 Visible 0
        set workbooks2  [$excel2 Workbooks]
        set workbook2   [$workbooks2 Open $fname]
        set worksheets2 [$workbook2 Worksheets]
        set cells2      [[$worksheets2 Item [expr 1]] Cells]
      
        set c1 [[[[$worksheets2 Item [expr 1]] UsedRange] Columns] Count]
        set r1 [[[[$worksheets2 Item [expr 1]] UsedRange] Rows] Count]
      
        for {set c 2} {$c <= $c1} {incr c} {
          set colName [[$cells2 Item 1 $c] Value]
          if {$colName != ""} {set i2($c) "nist_$colName"}
        }
      
        for {set r 2} {$r <= $r1} {incr r} {
          set i1 [[$cells2 Item $r 1] Value]
          if {$i1 != ""} {
            for {set c 2} {$c <= $c1} {incr c} {
              if {[info exists i2($c)]} {set spmiCoverages($i1,$i2($c)) [[$cells2 Item $r $c] Value]}
            }
          }
        }
        
        $workbooks2 Close
        $excel2 Quit
        catch {twapi::end_process $pid2 -force}
        after 100
      }
    }
          
# get expected PMI
    if {![info exists pmiMaster($nistName)]} {
      catch {unset pmiMaster($nistName)}
      set fn "SFA-PMI-$nistName.xlsx"
      if {[file exists NIST/$fn]} {file copy -force NIST/$fn $mytemp}
      set fname [file nativename [file join $mytemp $fn]]

      if {[file exists $fname]} {
        set pid1 [twapi::get_process_ids -name "EXCEL.EXE"]
        set excel2 [::tcom::ref createobject Excel.Application]
        set pid2 [lindex [intersect3 $pid1 [twapi::get_process_ids -name "EXCEL.EXE"]] 2]
    
        $excel2 Visible 0
        set workbooks2  [$excel2 Workbooks]
        set workbook2   [$workbooks2 Open $fname]
        set worksheets2 [$workbook2 Worksheets]
        set cells2      [[$worksheets2 Item [expr 1]] Cells]
      
        set r1 [[[[$worksheets2 Item [expr 1]] UsedRange] Rows] Count]
  
        for {set r 1} {$r <= $r1} {incr r} {
          set typ [[$cells2 Item $r 1] Value]
          set pmi [[$cells2 Item $r 2] Value]
          if {$typ != "" && $pmi != ""} {lappend pmiMaster($nistName) "$typ\\$pmi"}
        }
    
        $workbooks2 Close
        $excel2 Quit
        catch {twapi::end_process $pid2 -force}
        after 100
      }
    }

# errors
  } emsg]} {
    errorMsg "ERROR reading Expected PMI spreadsheet: $emsg"
  }
}

# -------------------------------------------------------------------------------
# add images for the CAx-IF and NIST PMI models
proc pmiAddModelPictures {ent} {
  global cells excel localName modelPictures modelURLs mytemp nistName wdir worksheet

  set ftail [string tolower [file tail $localName]]
  
  if {[catch {
    set nlink 0
    set fl ""
    foreach pic $modelPictures {
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
          [$excel ActiveWindow] FreezePanes [expr 1]
 
# group columns for image
          #if {[lindex $pic 3] > 0} {
          #  set range [$worksheet($ent) Range [cellRange 1 5] [cellRange 1 [lindex $pic 3]]]
          #  [$range Columns] Group
          #}

# link to test model drawings (doesn't always work)
          if {[string first "nist_" $fl] == 0 && $nlink < 2} {
            set str [string range $fl 0 10]
            foreach item $modelURLs {
              if {[string first $str $item] == 0} {
                catch {$cells($ent) Item 3 5 "Test Case Drawing"}
                set range [$worksheet($ent) Range E3:M3]
                $range MergeCells [expr 1]
                set range [$worksheet($ent) Range "E3"]
                [$worksheet($ent) Hyperlinks] Add $range [join "https://www.nist.gov/sites/default/files/documents/2016/09/08/$item"] [join ""] [join "Link to Test Case Drawing (PDF)"]
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
 
# -------------------------------------------------------------------------------
proc pmiFormatColumns {str} {
		global cells col excelVersion gpmiRow invGroup opt pmiStartCol recPracNames row spmiRow stepAP thisEntType worksheet
		
  if {![info exists pmiStartCol($thisEntType)]} {
    return
  } else {  
    set c1 [expr {$pmiStartCol($thisEntType)-1}]
  }
  
# delete unused columns
  set delcol 0
  set colrange [[[$worksheet($thisEntType) UsedRange] Columns] Count]
  for {set i $colrange} {$i > 3} {incr i -1} {
    set val [[$cells($thisEntType) Item 3 $i] Value]
    if {$val != ""} {set delcol 1}
    if {$val == "" && $delcol} {
      set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
      $range Delete
    }
  }
  set col($thisEntType) [[[$worksheet($thisEntType) UsedRange] Columns] Count]

# format
  if {[info exists cells($thisEntType)] && $col($thisEntType) > $c1} {
    set c2 [expr {$c1+1}]
    set c3 $col($thisEntType)
    
# PMI heading
    outputMsg " Formatting $str on: [formatComplexEnt $thisEntType]" blue
    $cells($thisEntType) Item 2 $c2 $str
    set range [$worksheet($thisEntType) Range [cellRange 2 $c2]]
    $range HorizontalAlignment [expr -4108]
    [$range Font] Bold [expr 1]
    [$range Interior] ColorIndex [expr 36]
    set range [$worksheet($thisEntType) Range [cellRange 2 $c2] [cellRange 2 $c3]]
    $range MergeCells [expr 1]

# set rows for colors and borders
    set r1 1
    set r2 $r1
    set r3 {}
    if {[string first "PMI Presentation" $str] != -1} {
      set rs $gpmiRow($thisEntType)
    } elseif {[string first "PMI Representation" $str] != -1} {
      set rs $spmiRow($thisEntType)
    }
    foreach r $rs {
      set r [expr {$r-2}]
      if {$r != [expr {$r2+1}]} {
        lappend r3 [list [expr {$r1+2}] [expr {$r2+2}]]
        set r1 $r
        set r2 $r1
      } else {
        set r2 $r
      }
    }
    lappend r3 [list [expr {$r1+2}] [expr {$r2+2}]]

# colors and borders
    set j 0
    for {set i $c2} {$i <= $c3} {incr i} {
      foreach r $r3 {
        set r1 [lindex $r 0]
        set r2 [lindex $r 1]

# cell color (yellow or green)        
        set range [$worksheet($thisEntType) Range [cellRange $r1 $i] [cellRange $r2 $i]]
        [$range Interior] ColorIndex [lindex [list 36 35] [expr {$j%2}]]

# dotted line border
        if {$excelVersion >= 12} {
          if {$i == $c2 && $r2 > 3} {
            if {$r1 < 4} {set r1 4}
            set range [$worksheet($thisEntType) Range [cellRange $r1 $c2] [cellRange $r2 $c3]]
            for {set k 7} {$k <= 12} {incr k} {
              catch {if {$k != 9 || [expr {$row($thisEntType)+0}] != $r2} {[[$range Borders] Item [expr $k]] Weight [expr 1]}}
            }
          }
        }
      }
      incr j
    }
    
# left and right borders in header
    catch {
      for {set i $c2} {$i <= $col($thisEntType)} {incr i} {
        set range [$worksheet($thisEntType) Range [cellRange 3 $i] [cellRange 3 $i]]
        [[$range Borders] Item [expr 7]]  Weight [expr 1]
        [[$range Borders] Item [expr 10]] Weight [expr 1]
      }
    }

# group columns
    if {[info exists invGroup($thisEntType)]} {if {$invGroup($thisEntType) < $c2} {set c2 $invGroup($thisEntType)}}
    set range [$worksheet($thisEntType) Range [cellRange 1 $c2] [cellRange [expr {$row($thisEntType)+2}] $c3]]
    [$range Columns] Group
    
# fix column widths
    set colrange [[[$worksheet($thisEntType) UsedRange] Columns] Count]
    for {set i 1} {$i <= $colrange} {incr i} {
      set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
      $range ColumnWidth [expr 96]
    }
    [$worksheet($thisEntType) Columns] AutoFit
    [$worksheet($thisEntType) Rows] AutoFit
    
# link to RP
    set str "pmi242"
    if {$stepAP == "AP203"} {set str "pmi203"}
    $cells($thisEntType) Item 2 1 "See CAx-IF Recommended Practice for $recPracNames($str)"
    if {$thisEntType != "dimensional_characteristic_representation"} {
      set range [$worksheet($thisEntType) Range A2:D2]
    } else {
      set range [$worksheet($thisEntType) Range A2:C2]
    }
    $range MergeCells [expr 1]
    set anchor [$worksheet($thisEntType) Range A2]
    [$worksheet($thisEntType) Hyperlinks] Add $anchor [join "https://www.cax-if.org/joint_testing_info.html#recpracs"] [join ""] [join "Link to CAx-IF Recommended Practices"]
  }
}

# -------------------------------------------------------------------------------
# check for an entity that is checked for semantic PMI
proc spmiCheckEnt {ent} {
  global spmiEntTypes tolNames

  set ok 0
  foreach sp $spmiEntTypes {if {[string first $sp $ent] ==  0} {set ok 1}}
  foreach sp $tolNames     {if {[string first $sp $ent] != -1} {set ok 1}}
  return $ok
}

# -------------------------------------------------------------------------------
# check for a valid form of annotation_occurrence
proc gpmiCheckEnt {ent} {
  global aoEntTypes

  set ok 0
  foreach item $aoEntTypes {
    if {[string first "tessellated" $item] == -1} {
      if {[string first $item $ent] == 0} {
        if {[string first "over_riding_styled_item" $ent] == -1 && \
            [string first "_relationship" $ent] == -1} {
          set ok 1
        }
      }
    } else {
      if {[string first $item $ent] != -1} {set ok 1}
    }
  }
  if {[string first "leader" $ent] != -1} {set ok 0}
  return $ok
}
 
# -------------------------------------------------------------------------------
# which STEP entities are processed depending on options
proc setEntsToProcess {entType objDesign} {
  global opt gpmiEnts spmiEnts
  
  set ok 0
  set gpmiEnts($entType) 0
  set spmiEnts($entType) 0

# for validation properties
  if {$opt(VALPROP)} {
    if {$entType == "boolean_representation_item" || \
        $entType == "conversion_based_unit_and_length_unit" || \
        $entType == "derived_unit_element" || \
        $entType == "descriptive_representation_item" || \
        $entType == "integer_representation_item" || \
        $entType == "length_unit_and_si_unit" || \
        $entType == "mass_unit_and_si_unit" || \
        $entType == "measure_representation_item" || \
        $entType == "property_definition" || \
        $entType == "property_definition_representation" || \
        $entType == "real_representation_item" || \
        $entType == "representation" || \
        $entType == "value_representation_item"} {
      set ok 1
    }
  }

# for PMI (graphical) presentation
  if {$opt(PMIGRF) && $ok == 0} {
    set ok [gpmiCheckEnt $entType]
    set gpmiEnts($entType) $ok
    if {$entType == "advanced_face" || \
        $entType == "characterized_item_within_representation" || \
        $entType == "colour_rgb" || \
        $entType == "curve_style" || \
        $entType == "fill_area_style" || \
        $entType == "fill_area_style_colour" || \
        $entType == "geometric_curve_set" || \
        $entType == "presentation_style_assignment" || \
        $entType == "property_definition" || \
        $entType == "representation_relationship" || \
        $entType == "tessellated_geometric_set" || \
        $entType == "view_volume"} {
      set ok 1
    }
    foreach ent {"annotation" "draughting" "_presentation" "camera" "constructive_geometry"} {
      if {[string first $ent $entType] != -1} {set ok 1}
    }
    if {!$ok} {if {[string first "representation" $entType] == -1 && [string first "presentation_" $entType] != -1} {set ok 1}}
  }

# for PMI (semantic) representation  
  if {$opt(PMISEM) && $ok == 0} {
    set spmiEnts($entType) [spmiCheckEnt $entType]
    if {$entType == "advanced_face" || \
        $entType == "compound_representation_item" || \
        $entType == "descriptive_representation_item" || \
        $entType == "dimensional_characteristic_representation" || \
        $entType == "draughting_model_item_association" || \
        $entType == "geometric_item_specific_usage" || \
        $entType == "id_attribute" || \
        $entType == "product_definition_shape" || \
        $entType == "property_definition" || \
        $entType == "shape_definition_representation" || \
        $entType == "shape_dimension_representation" || \
        $entType == "shape_representation_with_parameters" || \
        $entType == "value_format_type_qualifier" || \
        $entType == "value_range"} {
      set ok 1
    }
    foreach ent {"shape_aspect" "measure_with_unit" "measure_representation_item" "constructive_geometry"} {
      if {[string first $ent $entType] != -1} {set ok 1}
    }
    if {$entType == "axis2_placement_3d" && [$objDesign CountEntities "placed_datum_target_feature"] > 0} {set ok 1}
  }
  #outputMsg "$ok  $entType"
  return $ok
}

# -------------------------------------------------------------------------------
# check validation properties and PMI presentation
proc checkPMIValProps {objDesign entType} {
  global cells fixent gpmiEnts opt pmiColumns savedViewCol spmiEnts
  
# check for validation properties, call valProp
  if {$entType == "property_definition_representation"} {
    if {[catch {
      if {[info exists opt(VALPROP)]} {
        if {$opt(VALPROP)} {
          if {[lsearch $fixent "representation"] == -1} {
            if {[info exists cells(property_definition)]} {
              valPropStart $objDesign
            }
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding Validation Property information to '$entType'\n  $emsg"
    }

# check for PMI Presentation, call pmiProp
  } elseif {$gpmiEnts($entType)} {
    if {[catch {
      if {[info exists opt(PMIGRF)]} {
        if {$opt(PMIGRF)} {
          if {[info exists cells($entType)]} {
            gpmiProp $objDesign $entType
            set ok 0
          }
          catch {unset savedViewCol}
          catch {unset pmiColumns}
        }
      }
    } emsg]} {
      errorMsg "ERROR adding PMI Presentation information to '$entType'\n  $emsg"
    }

# check for Semantic PMI, call spmiDimtolStart or spmiGeotolStart
  } elseif {$spmiEnts($entType)} {
    if {[catch {
      if {[info exists opt(PMISEM)]} {
        if {$opt(PMISEM)} {
          if {[info exists cells($entType)]} {
            if {$entType == "dimensional_characteristic_representation"} {
              spmiDimtolStart $objDesign $entType
            } else {
              spmiGeotolStart $objDesign $entType
            }
            set ok 0
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding PMI Representation information to '[formatComplexEnt $entType]'\n  $emsg"
    }
  }
}

# -------------------------------------------------------------------------------
proc pmiSetEntAttrList {abc} {
  global elevel entAttrList ent opt

  incr elevel
  set ind [string repeat " " [expr {4*($elevel-1)}]]
  if {$opt(DEBUG1)} {outputMsg "$ind PARSE"}

  set ni 0
  foreach item $abc {
    if {[llength $item] > 1} {
      pmiSetEntAttrList $item
    } else {
      if {$ni == 0} {
        set typ "ENT"
        set ent($elevel) $item
      } else {
        set typ "  ATR"
        lappend entAttrList "$ent($elevel) $item"
      }
      if {$opt(DEBUG1)} {
        if {$typ == "ENT"} {
          outputMsg "$ind $typ $elevel $ni $item" blue
        } else {
          outputMsg "$ind $typ $elevel $ni $item"
        }
      }
      incr ni
    }
  }

  incr elevel -1
}  

# -------------------------------------------------------------------------------
# get STEP AP name
proc getStepAP {fname} {
  global fileSchema1
  
  set fs [getSchemaFromFile $fname]
  set fileSchema1 [string toupper $fs]
  #if {$fs != $fileSchema1} {errorMsg "File schema name '$fs' should be uppercase."}
  
  set ap ""
  foreach aps {AP203 AP209 AP210 AP238 AP239 AP242} {if {[string first $aps $fs] != -1} {set ap $aps}}

  if {$ap == ""} {
    if {[string first "CONFIGURATION_CONTROL_3D_DESIGN" $fileSchema1] != -1}  {set ap AP203e1}
    if {[string first "CONFIGURATION_CONTROL_3D_DESIGN_ED2" $fileSchema1] != -1} {set ap AP203}
    if {[string first "CONFIG_CONTROL" $fileSchema1] != -1}                   {set ap AP203e1}
    if {[string first "CCD_CLA_GVP_AST" $fileSchema1] != -1}                  {set ap AP203e1}
    if {[string first "STRUCTURAL_ANALYSIS_DESIGN" $fileSchema1] != -1}       {set ap AP209e1}
    if {[string first "AUTOMOTIVE_DESIGN" $fileSchema1] != -1}                {set ap AP214}
    if {[string first "INTEGRATED_CNC_SCHEMA" $fileSchema1] != -1}            {set ap AP238}
    if {[string first "STRUCTURAL_FRAME_SCHEMA" $fileSchema1] != -1}          {set ap CIS/2}
  }
  set stepAP $ap
  return $stepAP
}

#-------------------------------------------------------------------------------
proc getSchemaFromFile {fname {msg 0}} {
  set schema ""
  set ok 0
  set nline 0
  set stepfile [open $fname r]
  while {[gets $stepfile line] != -1 && $nline < 100} {
    if {$msg} {
      foreach item {"MIME-Version" "Content-Type" "X-MimeOLE" "DOCTYPE HTML" "META content"} {
        if {[string first $item $line] != -1} {
          errorMsg "Syntax Error: The STEP file was probably saved as an EMAIL or HTML file.  The STEP file cannot be translated.\n In the email client, save the STEP file as a TEXT file and try again.\n The first line in the STEP file should be 'ISO-10301-21\;'"
        }
      }
    }

    incr nline
    if {[string first "FILE_SCHEMA" $line] != -1} {
      set ok 1
      set fsline $line
    } elseif {[string first "ENDSEC" $line] != -1} {
      set sline [split $fsline "'"]
      set schema [lindex $sline 1]
      if {$msg} {
        errorMsg "STEP file schema: [lindex [split $schema " "] 0]"
        if {[llength $sline] > 3} {
          set schema1 [lindex $sline 3]
          errorMsg "Second STEP file schema: $schema1\n A second schema is valid but will not work with the STEP File Analyzer.\n Edit the FILE_SCHEMA to delete the second schema."
        } elseif {[string first "_MIM" $fsline] != -1 && [string first "_MIM_LF" $fsline] == -1} {
          errorMsg "The schema name should end with _MIM_LF"
        }

# check for CIS/2 or IFC files
        if {[string first "STRUCTURAL_FRAME_SCHEMA" $fsline] != -1} {
          errorMsg "This is a CIS/2 file that can be visualized with SteelVis.  http://go.usa.gov/s8fm"
        } elseif {[string first "IFC" $fsline] != -1} {
          errorMsg "Use the IFC File Analyzer with IFC files."
          after 1000
          displayURL http://go.usa.gov/xK9gh
        }
      }
      break
    } elseif {$ok} {
      append fsline $line
    }
  }
  close $stepfile
  return $schema
}

# -------------------------------------------------------------------------------
# CAx-IF vendor abbrevitations

proc setCAXIFvendor {} {
  global localName
  
  set vendor(3de) "CT 3D Evolution"
  set vendor(3DE) "CT 3D Evolution"
  set vendor(a3) "Acrobat 3D"
  set vendor(a5) "Acrobat 3D (CATIA V5)"
  set vendor(ac) "AutoCAD"
  set vendor(al) "AliasStudio"
  set vendor(ap) "Acrobat 3D (Pro/E)"
  set vendor(au) "Acrobat 3D (NX)"
  set vendor(c3e) "CATIA"
  set vendor(c4) "CATIA V4"
  set vendor(c5) "CATIA V5"
  set vendor(c6) "CATIA V6"
  set vendor(cg) "CgiStepCamp"
  set vendor(cm) "PTC CoCreate"
  set vendor(cr) "Creo"
  set vendor(dc) "Datakit CrossCad"
  set vendor(d5) "Datakit CrossCad (CATIA V5)"
  set vendor(dp) "Datakit CrossCad (Creo)"
  set vendor(du) "Datakit CrossCad (NX)"
  set vendor(dw) "Datakit CrossCad (SolidWorks)"
  set vendor(fs) "Vistagy FiberSim"
  set vendor(h3) "HOOPS 3D Exchange"
  set vendor(h5) "HOOPS 3D (CATIA V5)"
  set vendor(hp) "HOOPS 3D (Creo)"
  set vendor(hu) "HOOPS 3D (NX)"
  set vendor(i4) "ITI CADifx (CATIA V4)"
  set vendor(i5) "ITI CADfix (CATIA V5)"
  set vendor(id) "NX I-DEAS"
  set vendor(if) "ITI CADfix"
  set vendor(ii) "ITI CADfix (Inventor)"
  set vendor(in) "Autodesk Inventor"
  set vendor(ip) "ITI CADfix (Creo)"
  set vendor(iu) "ITI CADfix (NX)"
  set vendor(iw) "ITI CADfix (SolidWorks)"
  set vendor(kc) "Kubotek KeyCreator"
  set vendor(kr) "Kubotek REALyze"
  set vendor(lk) "LKSoft IDA-STEP"
  set vendor(nx) "Siemens NX"
  set vendor(oc) "Datakit CrossCad (OpenCascade)"
  set vendor(pc) "PTC CADDS"
  set vendor(pe) "Pro/E"
  set vendor(s4) "T-Systems COM/STEP (CATIA V4)"
  set vendor(s5) "T-Systems COM/FOX (CATIA V5)"
  set vendor(se) "SolidEdge"
  set vendor(sw) "SolidWorks"
  set vendor(t4) "Theorem Cadverter (CATIA V4)"
  set vendor(t5) "Theorem Cadverter (CATIA V5)"
  set vendor(tc) "Theorem Cadverter (CADDS)"
  set vendor(tp) "Theorem Cadverter (Creo)"
  set vendor(ts) "Theorem Cadverter (I-DEAS)"
  set vendor(tu) "Theorem Cadverter (NX)"
  set vendor(tx) "Theorem Cadverter"
  set vendor(ug) "Unigraphics"
  
  set fn [file tail $localName]
  set chars [list "-" "_" "."]
  foreach char $chars {
    set c1 [string first $char $fn]
    if {$c1 != -1} {
      set c2 [expr {$c1+1}]
      foreach idx [lsort [array names vendor]] {
        if {[string first $idx [string range $fn $c2 end]] == 0} {
          set c3 [expr {$c1+3}]
          set c4 [expr {$c1+4}]
          foreach char1 $chars {
            if {[string index $fn $c3] == $char1 || [string index $fn $c4] == $char1} {
              return $vendor($idx)
            }
          }
        }
      }
    }
  }
}
