#-------------------------------------------------------------------------------
# version numbers, software and user's guide
proc getVersion {}   {return 2.62}
proc getVersionUG {} {return 2.34}

# -------------------------------------------------------------------------------
# dt = 1 for dimtol
proc getAssocGeom {entDef {dt 0}} {
  global assocGeom entCount gtEntity recPracNames syntaxErr
  
  set entDefType [$entDef Type]
  #outputMsg "getGeom $dt $entDefType [$entDef P21ID]" blue

  if {[catch {
    if {$entDefType == "shape_aspect" || $entDefType == "centre_of_symmetry" || \
      ([string first "datum" $entDefType] != -1 && [string first "_and_" $entDefType] == -1)} {

# add shape_aspect to AG for dimtol
      if {$dt && ($entDefType == "shape_aspect" || $entDefType == "centre_of_symmetry" || $entDefType == "datum_feature" || \
                  [string first "datum_target" $entDefType] != -1)} {
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
      getAssocGeomFace $entDef
    
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
            set e1 [$a0 Value]
            if {[[$a0 Value] Type] == "composite_shape_aspect" || [[$a0 Value] Type] == "composite_group_shape_aspect"} {
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

# find AF for SA with GISU or IIRU
            foreach val $a0val {getAssocGeomFace $val}
          }
        }
      }
      
# check all around
      if {$entDefType == "all_around_shape_aspect"} {
        if {[llength $assocGeom($type)] == 1} {
          #outputMsg " assocGeom $type $assocGeom($type) [llength $assocGeom($type)]" green
          if {$type == "advanced_face"} {
            set msg "Syntax Error: For All Around tolerance, 'shape_aspect relationship' entity relates '$entDefType' to only one 'shape_aspect'.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 6.4, Fig. 31)"
            errorMsg $msg
            if {[info exists gtEntity]} {lappend syntaxErr([$gtEntity Type]) [list [$gtEntity P21ID] "toleranced_shape_aspect" $msg]}
          } elseif {$type == $entDefType} {
            set msg "Syntax Error: For All Around tolerance, missing 'shape_aspect relationship' entity relating '$entDefType' to 'shape_aspect'.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 6.4, Fig. 31)"
            errorMsg $msg
            if {[info exists gtEntity]} {lappend syntaxErr([$gtEntity Type]) [list [$gtEntity P21ID] "toleranced_shape_aspect" $msg]}
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
proc getAssocGeomFace {entDef} {

# look at GISU and IIRU for geometry associated with shape_aspect
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
}

# -------------------------------------------------------------------------------
proc appendAssocGeom {ent {id ""}} {
  global assocGeom
  
  set p21id [$ent P21ID]
  set type  [$ent Type]
  #outputMsg " appendAssocGeom $type $p21id $id" red
  
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
proc reportAssocGeom {entType {row ""}} {
  global assocGeom recPracNames dimRepeat dimRepeatDiv syntaxErr
  #outputMsg "reportAssocGeom $entType" red
  
  set str ""
  set dimRepeat 0
  set dimtol 0
  if {[string first "dimensional_" $entType] != -1 || [string first "angular_" $entType] != -1} {set dimtol 1}
  
# geometric entities
  foreach item [array names assocGeom] {
    if {[string first "shape_aspect" $item] == -1 && [string first "centre" $item] == -1 && \
        [string first "datum" $item] == -1 && [string first "draughting_callout" $item] == -1 && $item != "advanced_face"} {
      if {[string length $str] > 0} {append str [format "%c" 10]}
      append str "([llength $assocGeom($item)]) $item [lsort -integer $assocGeom($item)]"

# dimension count, e.g. 4X
      if {[string first "_size" $entType] != -1 || [string first "angular_location" $entType] != -1} {
        if {$item == "cylindrical_surface" || $item == "spherical_surface" || $item == "toroidal_surface" || $item == "conical_surface"} {

# set divider based on cylinders, assume two half cylinders, but if odd number of cylinders, then one complete cylinder
          set dc [llength $assocGeom($item)]
          if {$dc == 1} {set dimRepeatDiv 1}

          if {$dimRepeatDiv == 1} {
            if {$dc > 1} {incr dimRepeat $dc}
          } else {
            if {[expr {$dc%2}] == 0} {
              if {$dc > 3} {incr dimRepeat [expr {$dc/2}]}
            } else {
              if {$dc > 1} {incr dimRepeat $dc}
            }
          }
        }
      }
    }
  }

# advanced face
  foreach item [array names assocGeom] {
    if {$item == "advanced_face"} {
      if {[string length $str] > 0} {append str [format "%c" 10]}
      append str "([llength $assocGeom(advanced_face)]) $item [lsort -integer $assocGeom(advanced_face)]"
    }
  }
  if {[string length $str] == 0 && $dimtol} {
    set msg "Syntax Error: Associated Geometry not found for a '[formatComplexEnt $entType]'.\n[string repeat " " 14]Check GISU or IIRU 'definition' attribute or shape_aspect_relationship 'relating_shape_aspect' attribute.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.1, Figs. 5, 6, 12)"
    errorMsg $msg
    if {$row != ""} {lappend syntaxErr(dimensional_characteristic_representation) [list "-$row" "Associated Geometry" $msg]}
  }

# shape aspect
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
  global nistName pmiExpected pmiExpectedNX wdir mytemp legendColor pmiUnicode pmiFound pmiModifiers pmiActual recPracNames tolNames pmiType valType
  global nsimilar pmiMaster allPMI
  
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
    
    set comment "PMI Representation is collected from the datum systems, dimensions, tolerances, and datum target entities in column B"
    if {$nistName != ""} {append comment " and is color-coded by the expected PMI in the NIST test case drawing to the right.  The color-coding is explained at the bottom of the column.  Determining if the PMI is Partial and Possible match and corresponding Similar PMI depends on leading and trailing zeros, number precision, associated datum features and dimensions, and repetitive dimensions."}
    append comment "."
    addCellComment $spmiSumName $spmiSumRow 3 $comment 300 100
    
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
    
    outputMsg " Adding PMI Representation Summary worksheet" blue

# add pictures
    pmiAddModelPictures $spmiSumName
    [$worksheet($spmiSumName) Range "A1"] Select
  
# get expected PMI values from pmiMaster
    set nsimilar 0
    if {[info exists pmiMaster($nistName)]} {
      catch {unset pmiExpected($nistName)}
      catch {unset pmiExpectedNX($nistName)}
      
# read master PMI values, remove leading and trailing zeros, other stuff, add to pmiExpected
      foreach item $pmiMaster($nistName) {
        set c1 [string first "\\" $item]
        set typ [string range $item 0 $c1-1]
        set pmi [string range $item $c1+1 end]
        set newpmi [pmiRemoveZeros $pmi]
        lappend pmiExpected($nistName) $newpmi
        
# look for 'nX' in expected
        set c1 [string first "X" $newpmi]
        if {$c1 < 3} {
          set newpminx [string range $newpmi $c1+1 end]
          lappend pmiExpectedNX($nistName) [string trim $newpminx]
        } else {
          lappend pmiExpectedNX($nistName) $newpmi
        }
        
        if {[string first "tolerance" $typ] != -1} {
          foreach nam $tolNames {if {[string first $nam $typ] != -1} {set pmiType($newpmi) $nam}}
        } else {
          set pmiType($newpmi) $typ
        }
        set pmiActual($newpmi) $pmi
      }
      #for {set i 0} {$i < [llength $pmiExpected($nistName)]} {incr i} {outputMsg "$i / [lindex $pmiExpected($nistName) $i] / [lindex $pmiExpectedNX($nistName) $i]"}
    }
    set pmiFound {}
    set allPMI ""
    set checkPMImods 0
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
          set c1 [string first "(TZF:" $val]
          if {$c1 != -1} {set val [string range $val 0 $c1-2]}
          $cells($spmiSumName) Item $spmiSumRow 3 "'$val"
          set cellval $val

# allPMI used to count some modifiers for coverage analysis          
          if {[string first "tolerance" $thisEntType] != -1} {append allPMI $val}

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

# remove between
            set c1 [string first $pmiModifiers(between) $val]
            if {$c1 > 0} {set val [string range $val 0 $c1-2]}

# remove zeros from val
            set val [pmiRemoveZeros $val]
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
              set pmiExpected($nistName)   [lreplace $pmiExpected($nistName)   $pmiMatch $pmiMatch]
              set pmiExpectedNX($nistName) [lreplace $pmiExpectedNX($nistName) $pmiMatch $pmiMatch]
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
                  set pmiExpected($nistName)   [lreplace $pmiExpected($nistName)   $pos $pos]
                  set pmiExpectedNX($nistName) [lreplace $pmiExpectedNX($nistName) $pos $pos]
                  lappend pmiFound $pmi
                  #outputMsg "$val\n $pmiMatch $valType($val)" blue
                }
              }
              
# try match to expected without 'nX'              
              if {$pmiMatch == 0} {
                set pmiMatchNX [lsearch $pmiExpectedNX($nistName) $val]
                if {$pmiMatchNX != -1} {
                  #outputMsg "$val\n $pmiMatchNX $valType($val)" green
                  set pmiMatch 0.95
                  set pmiSim $pmiMatch
                  set pmiSimilar $pmiActual([lindex $pmiExpected($nistName) $pmiMatchNX])
                  set pmiExpected($nistName)   [lreplace $pmiExpected($nistName)   $pmiMatchNX $pmiMatchNX]
                  set pmiExpectedNX($nistName) [lreplace $pmiExpectedNX($nistName) $pmiMatchNX $pmiMatchNX]
                  set pf $val
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
                        if {[string first "datum_target" $valType($val)] == -1 && [string first "dimension" $valType($val)] == -1} {
                          if {$pmiSim >= 0.6} {
                            set pmiSimilar $pmiActual($pmi)
                            #append pmiSimilar "[format "%c" 10](Similarity: [string range $pmiMatch 0 4])"
                          }

# dimensions
                        } elseif {[string first "dimension" $valType($val)] != -1} {
                          #outputMsg "$pmiSim / $val / $pmi / [string first $val $pmi]" red
                          if {$pmiSim >= 0.6} {set pmiSimilar $pmiActual($pmi)}

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
                                    #outputMsg $diff green
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
                    addCellComment $spmiSumName 3 4 "Similar PMI is the best match of the PMI Representation in column C, for Partial or Possible matches (blue and yellow), to the expected PMI in the NIST test case drawing to the right."
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
          errorMsg "Missing PMI on [formatComplexEnt $thisEntType]"
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
    set lf 1
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
              if {$colName != ""} {set i2($c) "nist_$colName"}
              incr c
            }
          } else {
            set i1 [lindex $lline 0]
            foreach cval $lline {
              if {[info exists i2($c)]} {set spmiCoverages($i1,$i2($c)) $cval}
              incr c
            }
          }
          incr r
        }
        close $f
      }
    }
          
# get expected PMI
    if {![info exists pmiMaster($nistName)]} {
      catch {unset pmiMaster($nistName)}
      set fn "SFA-PMI-$nistName.xlsx"
      if {[file exists NIST/$fn]} {file copy -force NIST/$fn $mytemp}
      set fname [file nativename [file join $mytemp $fn]]

      if {[file exists $fname]} {
        if {$lf} {outputMsg " "}
        outputMsg "Reading Expected PMI for: $nistName (See Help > NIST CAD Models)" blue
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
          if {$typ != "" && $pmi != ""} {lappend pmiMaster($nistName) "$typ\\$pmi"}
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
          catch {[$excel ActiveWindow] FreezePanes [expr 1]}

# link to test model drawings (doesn't always work)
          if {[string first "nist_" $fl] == 0 && $nlink < 2} {
            set str [string range $fl 0 10]
            foreach item $modelURLs {
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
 
# -------------------------------------------------------------------------------
proc pmiFormatColumns {str} {
  global cells col excelVersion gpmiRow invGroup opt pmiStartCol recPracNames row spmiRow stepAP thisEntType worksheet tcl_platform excelYear
		
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
    outputMsg " [formatComplexEnt $thisEntType]"
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
      if {$opt(XLSBUG1) > 0 && ![file exists [file nativename C:/Windows/Fonts/ARIALUNI.TTF]]} {
        errorMsg "Excel $excelYear might not show some GD&T symbols correctly in PMI Representation reports.  The missing\n symbols will appear as question mark inside a square.  The likely cause is a missing font\n 'Arial Unicode MS' from the font file 'ARIALUNI.TTF'."
        incr opt(XLSBUG1) -1
      } elseif {$opt(XLSBUG1) < 30 && [file exists [file nativename C:/Windows/Fonts/ARIALUNI.TTF]]} {
        set opt(XLSBUG1) 30
      }
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

# group columns for inverses
    if {$c1 > 2} {
      set range [$worksheet($thisEntType) Range [cellRange 1 2] [cellRange [expr {$row($thisEntType)+2}] $c1]]
      [$range Columns] Group
    }
    
# fix column widths
    set colrange [[[$worksheet($thisEntType) UsedRange] Columns] Count]
    for {set i 1} {$i <= $colrange} {incr i} {
      set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
      $range ColumnWidth [expr 96]
    }
    [$worksheet($thisEntType) Columns] AutoFit
    [$worksheet($thisEntType) Rows] AutoFit
    
# link to RP
    set str1 "pmi242"
    if {[string first "AP203" $stepAP] == 0} {set str1 "pmi203"}
    $cells($thisEntType) Item 2 1 "See CAx-IF Rec. Prac. for $recPracNames($str1)"
    if {$thisEntType != "dimensional_characteristic_representation" && $thisEntType != "datum_reference"} {
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
proc setEntsToProcess {entType} {
  global objDesign
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

# for PMI (graphical) presentation report and viz
  if {($opt(PMIGRF) || $opt(VIZPMI)) && $ok == 0} {
    set ok [gpmiCheckEnt $entType]
    set gpmiEnts($entType) $ok
    if {$entType == "advanced_face" || \
        $entType == "characterized_item_within_representation" || \
        $entType == "colour_rgb" || \
        $entType == "curve_style" || \
        $entType == "fill_area_style" || \
        $entType == "fill_area_style_colour" || \
        $entType == "geometric_curve_set" || \
        $entType == "geometric_set" || \
        $entType == "presentation_style_assignment" || \
        $entType == "property_definition" || \
        $entType == "representation_relationship" || \
        $entType == "view_volume"} {
      set ok 1
    }
    foreach ent {"annotation" "draughting" "_presentation" "camera" "constructive_geometry" "tessellated_geometric_set"} {
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

# for tessellated geometry
  if {$opt(VIZTES) && $ok == 0} {
    if {[string first "tessellated" $entType] != -1} {set ok 1}
  }

  #outputMsg "$ok  $entType"
  return $ok
}

# -------------------------------------------------------------------------------
# check for all types of reports
proc checkForReports {entType} {
  global objDesign
  global cells skipEntities gpmiEnts opt pmiColumns savedViewCol spmiEnts
  
# check for validation properties report, call valPropStart
  if {$entType == "property_definition_representation"} {
    if {[catch {
      if {[info exists opt(VALPROP)]} {
        if {$opt(VALPROP)} {
          if {[lsearch $skipEntities "representation"] == -1} {
            if {[info exists cells(property_definition)]} {valPropStart}
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding Validation Properties to '$entType'\n  $emsg"
    }

# check for PMI Presentation report or viz graphical PMI, call gpmiAnnotation
  } elseif {$gpmiEnts($entType)} {
    if {[catch {
      set ok 0
      if {[info exists opt(PMIGRF)]} {if {$opt(PMIGRF)} {set ok 1}}
      if {[info exists opt(VIZPMI)]} {if {$opt(VIZPMI)} {set ok 1}}
      if {$ok} {
        if {[info exists cells($entType)] || $opt(VIZPMI)} {gpmiAnnotation $entType}
        catch {unset savedViewCol}
        catch {unset pmiColumns}
      }
    } emsg]} {
      errorMsg "ERROR adding PMI Presentation to '$entType'\n  $emsg"
    }
  
# viz tessellated part geometry, call tessPart
  } elseif {$entType == "tessellated_solid" || $entType == "tessellated_shell"} {
    if {[catch {
      if {[info exists opt(VIZTES)]} {if {$opt(VIZTES)} {tessPart $entType}}
    } emsg]} {
      errorMsg "ERROR adding Tessellated Part Geometry\n  $emsg"
    }

# check for Semantic PMI report, call spmiDimtolStart or spmiGeotolStart
  } elseif {$spmiEnts($entType)} {
    if {[catch {
      if {[info exists opt(PMISEM)]} {
        if {$opt(PMISEM)} {
          if {[info exists cells($entType)]} {
            if {$entType == "dimensional_characteristic_representation"} {
              spmiDimtolStart $entType
            } else {
              spmiGeotolStart $entType
            }
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding PMI Representation to '[formatComplexEnt $entType]'\n  $emsg"
    }

# check for AP209 analysis entities
  } elseif {$entType == "curve_3d_element_representation"   || \
            $entType == "surface_3d_element_representation" || \
            $entType == "volume_3d_element_representation"  || \
            $entType == "nodal_freedom_action_definition"   || \
            $entType == "single_point_constraint_element_values"} {
    if {[catch {
      if {[info exists opt(VIZFEA)]} {if {$opt(VIZFEA)} {feaModel $entType}}
    } emsg]} {
      errorMsg "ERROR adding Analysis Model to '$entType'\n  $emsg"
    }
  }
}

# -------------------------------------------------------------------------------
proc setEntAttrList {abc} {
  global entLevel entAttrList ent opt

  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]
  if {$opt(DEBUG1)} {outputMsg "$ind PARSE"}

  set ni 0
  foreach item $abc {
    if {[llength $item] > 1} {
      setEntAttrList $item
    } else {
      if {$ni == 0} {
        set typ "ENT"
        set ent($entLevel) $item
      } else {
        set typ "  ATR"
        lappend entAttrList "$ent($entLevel) $item"
      }
      if {$opt(DEBUG1)} {
        if {$typ == "ENT"} {
          outputMsg "$ind $typ $entLevel $ni $item" blue
        } else {
          outputMsg "$ind $typ $entLevel $ni $item"
        }
      }
      incr ni
    }
  }
  incr entLevel -1
}  

# -------------------------------------------------------------------------------
# get STEP AP name
proc getStepAP {fname} {
  global fileSchema1
  
  set fs [getSchemaFromFile $fname]
  set fileSchema1 [string toupper $fs]
  
  set ap ""
  foreach aps {AP203 AP209 AP210 AP238 AP239 AP242} {if {[string first $aps $fs] != -1} {set ap $aps}}

  if {$ap == ""} {
    if {[string first "CONFIGURATION_CONTROL_3D_DESIGN" $fileSchema1] != -1}     {set ap AP203e1}
    if {[string first "CONFIGURATION_CONTROL_3D_DESIGN_ED2" $fileSchema1] != -1} {set ap AP203}
    if {[string first "CONFIG_CONTROL" $fileSchema1] != -1}                      {set ap AP203e1}
    if {[string first "CCD_CLA_GVP_AST" $fileSchema1] != -1}                     {set ap AP203e1}

    if {[string first "STRUCTURAL_ANALYSIS_DESIGN" $fileSchema1] != -1} {set ap AP209e1}
    if {[string first "INTEGRATED_CNC_SCHEMA" $fileSchema1] != -1}      {set ap AP238}

    if {[string first "AUTOMOTIVE_DESIGN_CC2" $fileSchema1] != -1} {
      set ap AP214
    } elseif {[string first "AUTOMOTIVE_DESIGN" $fileSchema1] != -1} {
      set ap AP214e3
    }

    if {[string first "STRUCTURAL_FRAME_SCHEMA" $fileSchema1] != -1} {set ap CIS/2}
    if {[string first "IFC" $fileSchema1] != -1} {set ap $fileSchema1}
  }
  return $ap
}

#-------------------------------------------------------------------------------
proc getSchemaFromFile {fname {msg 0}} {
  global p21e3
  
  set p21e3 0
  set schema ""
  set ok 0
  set nline 0
  set niderr 0
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
          errorMsg "Second STEP file schema: $schema1\n A second schema is valid but will not work with the STEP File Analyzer.\n Export the STEP file with a single schema or edit the FILE_SCHEMA to delete the second schema."
        } elseif {[string first "_MIM" $fsline] != -1 && [string first "_MIM_LF" $fsline] == -1} {
          errorMsg "The schema name should end with _MIM_LF"
        }

# check for CIS/2 or IFC files
        if {[string first "STRUCTURAL_FRAME_SCHEMA" $fsline] != -1} {
          errorMsg "Use SteelVis to visualize the CIS/2 file.\n https://www.nist.gov/services-resources/software/steelvis-aka-cis2-viewer"
        } elseif {[string first "IFC" $fsline] != -1} {
          errorMsg "Use the IFC File Analyzer with IFC files."
          after 1000
          openURL https://www.nist.gov/services-resources/software/ifc-file-analyzer
        }
      }
      if {$p21e3} {break}

# check for IDs >= 2^31, valid but will be a different number in the spreadsheet
    } elseif {[string first "#" $line] == 0} {
      set id [string range $line 1 [string first "=" $line]-1]
      if {$id > 2147483647 && $niderr == 0} {
        errorMsg "An entity ID (#$id) >= 2147483648 (2^31)\n Very large IDs are valid but will appear as different numbers in the spreadsheet."
        incr niderr
      }
      
# check for part 21 edition 3 files
    } elseif {[string first "4\;1" $line] != -1 || [string first "ANCHOR\;" $line] != -1 || \
              [string first "REFERENCE\;" $line] != -1 || [string first "SIGNATURE\;" $line] != -1} {
      set p21e3 1
      if {[string first "4\;1" $line] == -1} {break}

    } elseif {$ok} {
      append fsline $line
    }
  }
  close $stepfile
  return $schema
}

#-------------------------------------------------------------------------------
proc checkP21e3 {fname} {
  global p21e3Section
  
  set p21e3Section {}
  set p21e3 0
  set nline 0
  set f1 [open $fname r]
      
# check for part 21 edition 3 file
  while {[gets $f1 line] != -1} {
    if {[string first "DATA\;" $line] == 0} {
      set nname $fname
      break
    } elseif {[string first "4\;1" $line] != -1 || \
              [string first "ANCHOR\;" $line] == 0 || \
              [string first "REFERENCE\;" $line] == 0 || \
              [string first "SIGNATURE\;" $line] == 0} {
      set p21e3 1
      break
    }
  }
  close $f1
  
# part 21 edition 3 file
  if {$p21e3} {

# new file name
    set nname "[file rootname $fname]-NOE3[file extension $fname]"
    catch {file delete -force $nname}
    set f2 [open $nname w]

# read file
    set write 1
    set data 0
    set sects {}
    
    set f1 [open $fname r]
    while {[gets $f1 line] != -1} {
      if {!$data} {
        if {[string first "DATA\;" $line] == 0} {
          set write 1
          set data 1
          regsub -all " " [join $sects] " and " sects
          errorMsg "The STEP file uses $sects section(s) from Edition 3 of Part 21.\n A new file ([file tail $nname]) without the those sections\n will be written and processed instead of ([file tail $fname])."
          
# check for part 21 edition 3 content
        } elseif {[string first "ANCHOR\;" $line] == 0 || \
                  [string first "REFERENCE\;" $line] == 0 || \
                  [string first "SIGNATURE\;" $line] == 0} {
          set write 0
          lappend sects [string range $line 0 end-1]
        }
  
# write new file w/o part 21 edition 3 content
        if {$write} {
          puts $f2 $line
        } else {
          lappend p21e3Section [string range $line 0 end-1]
        }

# in DATA section
      } else {
        puts $f2 $line
      }
    }
    close $f1
    close $f2
  }
  return $nname
}

# -------------------------------------------------------------------------------
# CAx-IF vendor abbrevitations, allVendor defined in sfa-data.tcl

proc setCAXIFvendor {} {
  global localName allVendor
  
  set fn [file tail $localName]
  set chars [list "-" "_" "."]
  foreach char $chars {
    set c1 [string first $char $fn]
    if {$c1 != -1} {
      set c2 [expr {$c1+1}]
      foreach idx [lsort [array names allVendor]] {
        if {[string first $idx [string range $fn $c2 end]] == 0} {
          set c3 [expr {$c1+3}]
          set c4 [expr {$c1+4}]
          foreach char1 $chars {
            if {[string index $fn $c3] == $char1 || [string index $fn $c4] == $char1} {
              return $allVendor($idx)
            }
          }
        }
      }
    }
  }
}
