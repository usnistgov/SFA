proc gpmiAnnotation {objDesign entType} {
  global ao aoEntTypes cells col entLevel ent entAttrList gpmiRow nindex opt pmiCol pmiHeading pmiStartCol
  global recPracNames stepAP syntaxErr x3domColor x3domCoord x3domFile x3domIndex x3domShape x3domMsg

  if {$opt(DEBUG1)} {outputMsg "START gpmiAnnotation $entType" red}

# basic geometry
  if {$opt(VIZPMI)} {
    set direction       [list direction name direction_ratios]
    set cartesian_point [list cartesian_point coordinates]
    set polyline        [list polyline name points $cartesian_point]
    set a2p3d           [list axis2_placement_3d location $cartesian_point axis $direction]
    #set a2p3d           [list axis2_placement_3d location $cartesian_point axis $direction ref_direction $direction]
    set circle          [list circle name position $a2p3d radius]
    set trimmed_curve   [list trimmed_curve name basis_curve $circle]
  } else {
    set circle          [list circle]
    set polyline        [list polyline]
    set trimmed_curve   [list trimmed_curve basis_curve]
  }
  set composite_curve [list composite_curve segments [list composite_curve_segment parent_curve $trimmed_curve]]                        

# tessellated geometry
  set complex_triangulated_surface_set [list complex_triangulated_surface_set name]
  set tessellated_curve_set            [list tessellated_curve_set name]
  set tessellated_geometric_set        [list tessellated_geometric_set name children]
  set repo_tessellated_geometric_set   [list repositioned_tessellated_item_and_tessellated_geometric_set name children]
  
# curve and fill style
  set colour      [list colour_rgb name red green blue]
  set curve_style [list presentation_style_assignment styles [list curve_style name curve_colour $colour [list draughting_pre_defined_colour name]]]
  #set curve_style [list presentation_style_assignment styles [list curve_style name curve_font curve_width curve_colour $colour [list draughting_pre_defined_colour name]]]
  set fill_style  [list presentation_style_assignment styles [list surface_style_usage style [list surface_side_style styles [list surface_style_fill_area fill_area [list fill_area_style name fill_styles [list fill_area_style_colour fill_colour $colour [list draughting_pre_defined_colour name]]]]]]]

  set geometric_curve_set  [list geometric_curve_set name elements $polyline $circle $trimmed_curve $composite_curve]
  set annotation_fill_area [list annotation_fill_area name boundaries $polyline $circle $trimmed_curve]

# annotation occurrence (clean up)
  set PMIP(annotation_occurrence)             [list annotation_occurrence name styles $curve_style item $geometric_curve_set]
  set PMIP(annotation_curve_occurrence)       [list annotation_curve_occurrence name styles $curve_style item $geometric_curve_set]
  set PMIP(annotation_curve_occurrence_and_geometric_representation_item) [list annotation_curve_occurrence_and_geometric_representation_item name styles $curve_style item $geometric_curve_set]
  set PMIP(annotation_fill_area_occurrence)   [list annotation_fill_area_occurrence name styles $fill_style item $annotation_fill_area]
  set PMIP(annotation_placeholder_occurrence) [list annotation_placeholder_occurrence name item $geometric_curve_set role]
  set PMIP(draughting_annotation_occurrence)  [list draughting_annotation_occurrence name styles $curve_style item $geometric_curve_set]
  set PMIP(draughting_annotation_occurrence_and_geometric_representation_item) [list draughting_annotation_occurrence_and_geometric_representation_item name styles $curve_style item $geometric_curve_set]
  set PMIP(tessellated_annotation_occurrence) [list tessellated_annotation_occurrence name styles $curve_style item $tessellated_geometric_set $repo_tessellated_geometric_set]

  #set PMIP(over_riding_styled_item_and_tessellated_annotation_occurrence) [list over_riding_styled_item_and_tessellated_annotation_occurrence name styles $curve_style item $tessellated_geometric_set $repo_tessellated_geometric_set]
    
# generate correct PMIP variable accounting for variations like characterized_object
  if {![info exists PMIP($entType)]} {
    foreach item $aoEntTypes {
      if {[string first $item $entType] != -1} {
        set PMIP($entType) $PMIP($item)
        lset PMIP($entType) 0 $entType
        break
      }
    }
  }
   
  if {![info exists PMIP($entType)]} {return}
  set ao $entType

  set entAttrList {}
  set pmiCol 0
  set nindex 0
  set x3domShape 0
  set gpmiRow($ao) {}
  if {![info exists x3domMsg]} {set x3domMsg {}}

  if {[info exists pmiHeading]} {unset pmiHeading}
  if {[info exists ent]}        {unset ent}

  outputMsg " Adding PMI Presentation" green
  
# look for syntax errors with entity usage
  if {$stepAP == "AP242"} {
    set c1 [string first "_and_characterized_object" $ao]
    set c2 [string first "characterized_object_and_" $ao]
    if {$c1 != -1} {
      errorMsg "Syntax Error: Using 'characterized_object' with '[string range $ao 0 $c1-1]' is not valid for PMI Presentation.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 10.2, 10.3)"
      lappend syntaxErr($ao) [list 1 1]
    } elseif {$c2 != -1} {
      errorMsg "Syntax Error: Using 'characterized_object' with '[string range $ao 25 end]' is not valid for PMI Presentation.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 10.2, 10.3)"
      lappend syntaxErr($ao) [list 1 1]
    }
  
    if {[string first "annotation_occurrence" $ao] != -1 && [string first "tessellated" $ao] == -1 && [string first "draughting_annotation_occurrence" $ao] == -1} {
      errorMsg "Syntax Error: Using 'annotation_occurrence' with $stepAP is not valid for PMI Presentation.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 8.1.1)"
      lappend syntaxErr($ao) [list 1 1]
    }
  }

  if {[string first "AP203" $stepAP] == 0 || $stepAP == "AP214"} {
    if {[string first "annotation_curve_occurrence" $ao] != -1} {
      errorMsg "Syntax Error: Using 'annotation_curve_occurrence' with $stepAP is not valid for PMI Presentation.\n[string repeat " " 14]\($recPracNames(pmi203), Sec. 4.1.1)"
      lappend syntaxErr($ao) [list 1 1]
    }
  }
  
  if {[string first "draughting" $ao] != -1} {
    set msg "Syntax Error: Using 'draughting_annotation_*_occurrence' is not valid for PMI Presentation.\n[string repeat " " 14]"
    if {$stepAP == "AP242"} {
      append msg "($recPracNames(pmi242), Sec. 8.1)"
    } else {
      append msg "($recPracNames(pmi203), Sec. 4.1)"
    }
    errorMsg $msg
    lappend syntaxErr($ao) [list 1 1]
  }

  if {$opt(DEBUG1)} {outputMsg \n}
  set entLevel 0
  setEntAttrList $PMIP($ao)
  if {$opt(DEBUG1)} {outputMsg "entattrlist $entAttrList"}
  if {$opt(DEBUG1)} {outputMsg \n}
    
  set startent [lindex $PMIP($ao) 0]
  set n 0
  set entLevel 0
  
# get next unused column by checking if there is a colName
  set pmiStartCol($ao) [getNextUnusedColumn $startent 3]

# process all annotation_occurrence entities, call gpmiAnnotationReport
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    if {[$objEntity Type] == $startent} {
      if {$n < 10000000} {
        if {[expr {$n%2000}] == 0} {
          if {$n > 0} {outputMsg "  $n"}
          update idletasks
        }
        gpmiAnnotationReport $objEntity
        if {$opt(DEBUG1)} {outputMsg \n}
      }
      incr n
    }
  }
  set col($ao) $pmiCol
  
# write any remaining X3DOM
  if {[info exists x3domCoord] || $x3domShape} {
    if {[string length $x3domCoord] > 0} {
      puts $x3domFile " <IndexedLineSet coordIndex='[string trim $x3domIndex]'>\n  <Coordinate point='[string trim $x3domCoord]'></Coordinate>\n </IndexedLineSet>\n</Shape>"
      set x3domCoord ""
      set x3domIndex ""
    } elseif {$x3domShape} {
      puts $x3domFile "</IndexedLineSet></Shape>"
    }
    set x3domShape 0
    set x3domColor ""
  }
}

# -------------------------------------------------------------------------------

proc gpmiAnnotationReport {objEntity} {
  global ao aoname assocGeom avgX3domColor badAttributes cells circleCenter col currX3domPointID curveTrim dirRatio dirType draftModelCameras
  global entLevel ent entAttrList entCount geomType gpmiEnts gpmiID gpmiIDRow gpmiOK gpmiRow gpmiTypes gpmiTypesInvalid gpmiTypesPerFile gpmiValProp
  global iCompCurve iCompCurveSeg incrcol iPolyline localName nindex numCompCurve numCompCurveSeg numPolyline numX3domPointID
  global objEntity1 opt pmiCol pmiColumns pmiHeading pmiStartCol pointLimit prefix propDefIDS recPracNames savedViewCol stepAP syntaxErr 
  global x3domColor x3domCoord x3domFile x3domFileName x3domFileOpen x3domIndex x3domMax x3domMin x3domPoint x3domPointID x3domShape x3domMsg
  global nistVersion

  #outputMsg "gpmiAnnotationReport" red
  #if {[info exists gpmiOK]} {if {$gpmiOK == 0} {return}}

# entLevel is very important, keeps track level of entity in hierarchy
  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]

  set maxcp $pointLimit

  if {[string first "handle" $objEntity] == -1} {
    #if {$objEntity != ""} {outputMsg "$ind $objEntity"}
    #outputMsg "  $objEntity" red
  } else {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set objAttributes [$objEntity Attributes]
    set ent($entLevel) $objType
    if {$entLevel == 1} {set objEntity1 $objEntity}

    if {$opt(DEBUG1) && $objType != "cartesian_point"} {outputMsg "$ind ENT $entLevel #$objID=$objType (ATR=[$objAttributes Count])" blue}
    #if {$entLevel == 1} {outputMsg "#$objID=$objType" blue}

# check if there are rows with ao
    if {$gpmiEnts($objType)} {
      set gpmiID $objID
      if {![info exists gpmiIDRow($ao,$gpmiID)]} {
        incr entLevel -1
        set gpmiOK 0
        return
      } else {
        set gpmiOK 1
      }

# write any leftover X3DOM from previous Shape
      if {[info exists x3domCoord] || $x3domShape} {
        if {[string length $x3domCoord] > 0} {
          puts $x3domFile " <IndexedLineSet coordIndex='[string trim $x3domIndex]'>\n  <Coordinate point='[string trim $x3domCoord]'></Coordinate>\n </IndexedLineSet>\n</Shape>"
          set x3domCoord ""
          set x3domIndex ""
        } elseif {$x3domShape} {
          puts $x3domFile "</IndexedLineSet></Shape>"
        }
        set x3domShape 0
        set x3domColor ""
      }
    }
    
# keep track of the number of c_c or c_c_s, if not polyline
    if {$objType == "composite_curve"} {
      incr iCompCurve
    } elseif {$objType == "composite_curve_segment"} {
      incr iCompCurveSeg
    }
    
    if {$entLevel == 2 && \
        $objType != "geometric_curve_set" && $objType != "annotation_fill_area" && $objType != "presentation_style_assignment" && \
        [string first "tessellated_geometric_set" $objType] == -1} {
      set msg "Syntax Error: '$objType' is not allowed as an 'item' attribute of: $ao\n[string repeat " " 14]"
      if {$stepAP == "AP242"} {
        append msg "($recPracNames(pmi242), Sec. 8.1.1, 8.1.2, 8.2)"
      } else {
        append msg "($recPracNames(pmi203), Sec. 4.1.1, 4.1.2)"
      }
      errorMsg $msg
      lappend syntaxErr($ao) [list $gpmiID item]
    }

    ::tcom::foreach objAttribute $objAttributes {
      set objName  [$objAttribute Name]
      set ent1 "$ent($entLevel) $objName"
      set ent2 "$ent($entLevel).$objName"

# look for entities with bad attributes that cause a crash
      set okattr 1
      if {[info exists badAttributes($objType)]} {foreach ba $badAttributes($objType) {if {$ba == $objName} {set okattr 0}}}
      #outputMsg "$ent1 $okattr" blue

      if {$okattr} {
        set objValue    [$objAttribute Value]
        set objNodeType [$objAttribute NodeType]
        set objSize     [$objAttribute Size]
        set objAttrType [$objAttribute Type]
  
        set idx [lsearch $entAttrList $ent1]

# -----------------
# nodeType = 18,19
        if {$objNodeType == 18 || $objNodeType == 19} {
          if {[catch {
            if {$idx != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}
    
              if {[info exists cells($ao)]} {
                set ok 0

# get values for these entity and attribute pairs
                switch -glob $ent1 {
                  "trimmed_curve basis_curve" {
                    if {[$objValue Type] != "circle"} {
                      set msg "Syntax Error: '[$objValue Type]' is not allowed as a 'basis_curve' for trimmed_curve\n[string repeat " " 14]"
                      if {$stepAP == "AP242"} {
                        append msg "($recPracNames(pmi242), Sec. 8.1.1, 8.1.2)"
                      } else {
                        append msg "($recPracNames(pmi203), Sec. 4.1.1, 4.1.2)"
                      }
                      errorMsg $msg
                      set msg "Some annotation line segments may be missing."
                      if {[lsearch $x3domMsg $msg] == -1} {lappend x3domMsg $msg}
                    }
                  }
                  "axis2_placement_3d axis" {set dirType "axis"}
                }
  
                set colName "value"
    
                if {$ok && [info exists gpmiID]} {
                  set c [string index [cellRange 1 $col($ao)] 0]
                  set r $gpmiIDRow($ao,$gpmiID)

# column name
                  if {![info exists pmiHeading($col($ao))]} {
                    #$cells($ao) Item 1 $c $colName
                    $cells($ao) Item 3 $c $colName
                    set pmiHeading($col($ao)) 1
                  }

# keep track of rows with validation properties
                  if {[lsearch $gpmiRow($ao) $r] == -1} {lappend gpmiRow($ao) $r}

# value in spreadsheet
                  set val [[$cells($ao) Item $r $c] Value]
                  if {$val == ""} {
                    $cells($ao) Item $r $c $objValue
                  } else {
                    $cells($ao) Item $r $c "$val[format "%c" 10]$objValue"
                  }

# entity reference in spreadsheet
                  incr col($ao)
                  set c [string index [cellRange 1 $col($ao)] 0]
                  set val [[$cells($ao) Item $r $c] Value]
                  if {$val == ""} {
                    $cells($ao) Item $r $c "#$objID $ent2"
                  } else {
                    $cells($ao) Item $r $c "$val[format "%c" 10]#$objID $ent2"
                  }

# keep track of max column
                  set pmiCol [expr {max($col($ao),$pmiCol)}]
                }
              }

# if referred to another, get the entity
              if {[string first "handle" $objEntity] != -1} {gpmiAnnotationReport $objValue}
            }
          } emsg3]} {
            errorMsg "ERROR processing PMI Presentation ($objNodeType $ent2): $emsg3"
          }

# --------------
# nodeType = 20
        } elseif {$objNodeType == 20} {
          if {[catch {
            if {$idx != -1} {
              if {$opt(DEBUG1) && $objName != "coordinates"} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}
          
# start of a list of cartesian points, assuming it is for a polyline, entLevel = 3
              if {$objAttrType == "ListOfcartesian_point" && $entLevel == 3} {
                #outputMsg 1entLevel$entLevel red
                if {$maxcp <= 10 && $maxcp < $objSize} {
                  append x3domPointID "($maxcp of $objSize) cartesian_point "
                } else {
                  append x3domPointID "($objSize) cartesian_point "
                }
                set numX3domPointID $objSize
                set currX3domPointID 0
                incr iPolyline
    
                set str ""
                for {set i 0} {$i < $objSize} {incr i} {
                  append x3domIndex "[expr {$i+$nindex}] "
                }
                append x3domIndex "-1 "
                incr nindex $objSize
              }
    
              if {[info exists cells($ao)]} {
                set ok 0

# get values for these entity and attribute pairs
# g_c_s and a_f_a both start keeping track of their polylines
# cartesian_point is need to generated X3DOM
                switch -glob $ent1 {
                  "cartesian_point coordinates" {
                    if {$opt(VIZPMI) && $x3domFileName != ""} {
                      #outputMsg "$entLevel $geomType $ent1" red

# entLevel = 4 for polyline
                      if {$entLevel == 4 && $geomType == "polyline"} {
                        append x3domCoord "[format "%.4f" [lindex $objValue 0]] [format "%.4f" [lindex $objValue 1]] [format "%.4f" [lindex $objValue 2]] " 
    
                        set x3domPoint(x) [lindex $objValue 0]
                        set x3domPoint(y) [lindex $objValue 1]
                        set x3domPoint(z) [lindex $objValue 2]

# min,max of points
                        foreach idx {x y z} {
                          if {$x3domPoint($idx) > $x3domMax($idx)} {set x3domMax($idx) $x3domPoint($idx)}
                          if {$x3domPoint($idx) < $x3domMin($idx)} {set x3domMin($idx) $x3domPoint($idx)}
                        }

# write coord and index to X3DOM file for polyline
                        if {$entLevel == 4} {
                          if {$iPolyline == $numPolyline && $currX3domPointID == $numX3domPointID} {
                            outputMsg "polyline" blue
                            puts $x3domFile " <IndexedLineSet coordIndex='[string trim $x3domIndex]'>\n  <Coordinate point='[string trim $x3domCoord]'></Coordinate>\n </IndexedLineSet>\n</Shape>"
                            set x3domCoord ""
                            set x3domIndex ""
                            set x3domShape 0
                            set x3domColor ""
                          }
                        }             

# circle center
                      } elseif {$geomType == "circle"} {
                        set circleCenter $objValue
                      }
                    }
                  }
                  "geometric_curve_set elements" {
                    set ok 1
                    set col($ao) [expr {$pmiStartCol($ao)+1}]
                    if {$stepAP == "AP242"} {
                      set colName "elements[format "%c" 10](Sec. 8.1.1)"
                    } else {
                      set colName "elements[format "%c" 10](Sec. 4.1.1)"
                    }
# keep track of polyline items
                    set numPolyline $objSize
                    set x3domPointID ""
                    set iPolyline 0
                    set x3domIndex ""
                    set x3domCoord ""
                    set nindex 0
# keep track of composite curve items
                    set numCompCurve $objSize
                    set iCompCurve 0
                  }
                  "annotation_fill_area boundaries" {
                    set ok 1
                    set col($ao) [expr {$pmiStartCol($ao)+1}]
                    if {$stepAP == "AP242"} {
                      set colName "boundaries[format "%c" 10](Sec. 8.1.2)"
                    } else {
                      set colName "boundaries[format "%c" 10](Sec. 4.1.2)"
                    }
# keep track of polyline items
                    set numPolyline $objSize
                    set x3domPointID ""
                    set iPolyline 0
                    set x3domIndex ""
                    set x3domCoord ""
                    set nindex 0
                  }
                  "*tessellated_geometric_set children" {
                    set ok 1
                    set col($ao) [expr {$pmiStartCol($ao)+1}]
                    set colName "children[format "%c" 10](Sec. 8.2)"
                  }
                  "direction direction_ratios" {
                    set dirRatio(x) [format "%.4f" [lindex $objValue 0]]
                    set dirRatio(y) [format "%.4f" [lindex $objValue 1]]
                    set dirRatio(z) [format "%.4f" [lindex $objValue 2]]
                  }
                  "composite_curve segments" {set numCompCurveSeg $objSize}
                }

# value in spreadsheet
                if {$ok && [info exists gpmiID]} {
                  set c [string index [cellRange 1 $col($ao)] 0]
                  set r $gpmiIDRow($ao,$gpmiID)

# column name
                  if {![info exists pmiHeading($col($ao))]} {
                    $cells($ao) Item 3 $c $colName
                    set pmiHeading($col($ao)) 1
                    set pmiCol [expr {max($col($ao),$pmiCol)}]
                  }

# keep track of rows with PMI properties
                  if {[lsearch $gpmiRow($ao) $r] == -1} {lappend gpmiRow($ao) $r}

# format cellval into str
                  if {[catch {
                    ::tcom::foreach val [$objAttribute Value] {
                      append cellval([$val Type]) "[$val P21ID] "
                      if {$ent1 == "tessellated_geometric_set children" && [$val Type] != "tessellated_curve_set" && [$val Type] != "complex_triangulated_surface_set"} {
                        errorMsg "Syntax Error: Invalid '[$val Type]' attribute for tessellated_geometric_set.children\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 8.2)"
                      }
                    }
                  } emsg]} {
                    foreach val [$objAttribute Value] {
                      append cellval([$val Type]) "[$val P21ID] "
                      if {$ent1 == "geometric_curve_set items" && [$val Type] != "polyline" && [$val Type] != "trimmed_curve" && [$val Type] != "circle"} {
                        set msg "Syntax Error: Invalid '[$val Type]' attribute for geometric_curve_set.items\n[string repeat " " 14]"
                        if {$stepAP == "AP242"} {
                          append msg "($recPracNames(pmi242), Sec. 8.1.1, 8.1.2)"
                        } else {
                          append msg "($recPracNames(pmi203), Sec. 4.1.1, 4.1.2)"
                        }
                        errorMsg $msg
                      }
                    }
                  }
  
                  set str ""
                  set size 0
                  catch {set size [array size cellval]}
                  if {$size > 0} {
                    foreach idx [lsort [array names cellval]] {
                      set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
                      if {$ncell > 1 || $size > 1} {
                        if {$ncell < 25} {
                          append str "($ncell) $idx $cellval($idx)  "
                        } else {
                          append str "($ncell) $idx  "
                        }
                      } else {
                        append str "(1) $idx $cellval($idx)  "
                      }
                    }
                  }
                  set ov [string trim $str]
    
                  set val [[$cells($ao) Item $r $c] Value]
                  if {$val == ""} {
                    $cells($ao) Item $r $c $ov
                  } else {
                    if {[catch {
                      $cells($ao) Item $r $c "$val[format "%c" 10]$ov"
                    } emsg]} {
                      errorMsg "  ERROR: Too much data to show in a cell" red
                    }
                  }
              
# keep track of max column
                  set pmiCol [expr {max($col($ao),$pmiCol)}]
                }
              }

# -------------------------------------------------
# recursively get the entities that are referred to
              if {[catch {
                ::tcom::foreach val1 $objValue {gpmiAnnotationReport $val1}
              } emsg]} {
                foreach val2 $objValue {gpmiAnnotationReport $val2}
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing PMI Presentation ($objNodeType $ent2): $emsg3"
          }

# ---------------------
# nodeType = 5 (!= 18,19,20)
        } else {
          if {[catch {
            if {$idx != -1} {
              if {$opt(DEBUG1) && $ent1 != "cartesian_point name"} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objAttrType)  ($ent1)"}
    
              if {[info exists cells($ao)]} {
                set ok 0
                set colName ""

# get values for these entity and attribute pairs
                switch -glob $ent1 {
                  "circle radius" {
                    if {$opt(VIZPMI) && $x3domFileName != ""} {
# write circle to X3DOM                    
                      #set ns 8
                      set ns 24
                      set angle 0
                      set dlt [expr {6.28319/$ns}]
                      set trimmed 0
                      if {[info exists curveTrim(trim_1)]} {
                        set angle $curveTrim(trim_1)
                        set dlt [expr {($curveTrim(trim_2)-$curveTrim(trim_1))/$ns}]
                        set trimmed 1
                        incr ns
                        unset curveTrim
                      }
                      for {set i 0} {$i < $ns} {incr i} {append x3domIndex "[expr {$i+$nindex}] "}
                      if {!$trimmed} {
                        append x3domIndex "$nindex -1 "
                      } else {
                        append x3domIndex "-1 "
                      }
                      incr nindex $ns
  
                      for {set i 0} {$i < $ns} {incr i} {
                        if {[expr {abs($dirRatio(z))}] > 0.99} {
                          set x3domPoint(x) [expr {$objValue*cos($angle)+[lindex $circleCenter 0]}]
                          set x3domPoint(y) [expr {-1.*$objValue*sin($angle)+[lindex $circleCenter 1]}]
                          set x3domPoint(z) [lindex $circleCenter 2]
                        } elseif {[expr {abs($dirRatio(y))}] > 0.99} {
                          set x3domPoint(x) [expr {$objValue*cos($angle)+[lindex $circleCenter 0]}]
                          set x3domPoint(z) [expr {-1.*$objValue*sin($angle)+[lindex $circleCenter 2]}]
                          set x3domPoint(y) [lindex $circleCenter 1]
                        } elseif {[expr {abs($dirRatio(x))}] > 0.99} {
                          set x3domPoint(z) [expr {$objValue*cos($angle)+[lindex $circleCenter 2]}]
                          set x3domPoint(y) [expr {-1.*$objValue*sin($angle)+[lindex $circleCenter 1]}]
                          set x3domPoint(x) [lindex $circleCenter 0]
                        } else {
                          errorMsg "PMI annotation circle orientation ($dirRatio(x), $dirRatio(y), $dirRatio(z)) is ignored."
                          set msg "Complex circle orientation is ignored."
                          if {[lsearch $x3domMsg $msg] == -1} {lappend x3domMsg $msg}
                          set x3domPoint(x) [expr {$objValue*cos($angle)+[lindex $circleCenter 0]}]
                          set x3domPoint(y) [expr {-1.*$objValue*sin($angle)+[lindex $circleCenter 1]}]
                          set x3domPoint(z) [lindex $circleCenter 2]
                        }
                        foreach idx {x y z} {
                          if {$x3domPoint($idx) > $x3domMax($idx)} {set x3domMax($idx) $x3domPoint($idx)}
                          if {$x3domPoint($idx) < $x3domMin($idx)} {set x3domMin($idx) $x3domPoint($idx)}
                        }
                        append x3domCoord "[format "%.4f" $x3domPoint(x)] [format "%.4f" $x3domPoint(y)] [format "%.4f" $x3domPoint(z)] "
                        set angle [expr {$angle+$dlt}]
                      }
                    }
                  }
                  "trimmed_curve name" {
# get trim values here
                    ::tcom::foreach a0 $objAttributes {
                      if {[string first "trim" [$a0 Name]] != -1} {
                        set val [$a0 Value]
                        if {[string first "handle" $val] != -1} {set val [lindex $val 1]}
                        set curveTrim([$a0 Name]) $val
                      }
                    }
                    errorMsg "Trimmed circles in PMI annotations might have the wrong orientation."
                    set msg "Trimmed circles might have the wrong orientation."
                    if {[lsearch $x3domMsg $msg] == -1} {lappend x3domMsg $msg}
                  }
                  "cartesian_point name" {
                    if {$entLevel == 4} {
                      set ok 1
                      set col($ao) [expr {$pmiStartCol($ao)+2}]
                    }
                  }
                  "geometric_curve_set name" -
                  "annotation_fill_area name" -
                  "*tessellated_geometric_set name" {
                    set ok 1
                    set col($ao) $pmiStartCol($ao)
                    if {$stepAP == "AP242"} {
                      set colName "name[format "%c" 10](Sec. 8.4)"
                    } else {
                      set colName "name[format "%c" 10](Sec. 4.3)"
                    }
                  }
                  "annotation_curve_occurrence* name" -
                  "annotation_fill_area_occurrence* name" -
                  "annotation_placeholder_occurrence* name" -
                  "*annotation_occurrence* name" -
                  "*tessellated_annotation_occurrence* name" {
                    set aoname $objValue
                    if {[string first "fill" $ent1] != -1} {
                      set msg "Filled characters are not filled."
                      if {[lsearch $x3domMsg $msg] == -1} {lappend x3domMsg $msg}
                    }
                  }
                  "curve_style name" -
                  "fill_area_style name" {
                    set ok 1
                    set col($ao) [expr {$pmiStartCol($ao)+2}]
                    if {$stepAP == "AP242"} {
                      set colName "presentation style[format "%c" 10](Sec. 8.5)"
                    } else {
                      set colName "presentation style[format "%c" 10](Sec. 4.4)"
                    }
                    update idletasks
                  }
                  "colour_rgb red" {
                    if {$entLevel == 4 || $entLevel == 8} {
                      set x3domColor $objValue
                      if {$opt(gpmiColor) > 0} {set x3domColor [gpmiSetColor $opt(gpmiColor)]}
                    }
                  }
                  "colour_rgb green" {
                    if {$entLevel == 4 || $entLevel == 8} {
                      append x3domColor " $objValue"
                      if {$opt(gpmiColor) > 0} {set x3domColor [gpmiSetColor $opt(gpmiColor)]}
                    }
                  }
                  "colour_rgb blue" {
                    if {$entLevel == 4 || $entLevel == 8} {
                      append x3domColor " $objValue"
                      set ok 1
                      set col($ao) [expr {$pmiStartCol($ao)+3}]
                      if {$stepAP == "AP242"} {
                        set colName "color[format "%c" 10](Sec. 8.5)"
                      } else {
                        set colName "color[format "%c" 10](Sec. 4.4)"
                      }
                      set pmiCol [expr {max($col($ao),$pmiCol)}]
                      if {$opt(gpmiColor) > 0} {set x3domColor [gpmiSetColor $opt(gpmiColor)]}
                    }
                  }
                  "draughting_pre_defined_colour name" {
                    if {$entLevel == 4 || $entLevel == 8} {
                      if {$objValue == "white"} {
                        set x3domColor "1 1 1"
                      } elseif {$objValue == "black"} {
                        set x3domColor "0 0 0"
                      } elseif {$objValue == "red"} {
                        set x3domColor "1 0 0"
                      } elseif {$objValue == "yellow"} {
                        set x3domColor "1 1 0"
                      } elseif {$objValue == "green"} {
                        set x3domColor "0 1 0"
                      } elseif {$objValue == "cyan"} {
                        set x3domColor "0 1 1"
                      } elseif {$objValue == "blue"} {
                        set x3domColor "0 0 1"
                      } elseif {$objValue == "magenta"} {
                        set x3domColor "1 0 1"
                      } else {
                        errorMsg "Syntax Error: Unknown draughting_pre_defined_colour name '$objValue'"
                      }
                      if {$opt(gpmiColor) > 0} {set x3domColor [gpmiSetColor $opt(gpmiColor)]}
                      set ok 1
                      set col($ao) [expr {$pmiStartCol($ao)+3}]
                      if {$stepAP == "AP242"} {
                        set colName "color[format "%c" 10](Sec. 8.5)"
                      } else {
                        set colName "color[format "%c" 10](Sec. 4.4)"
                      }
                      set pmiCol [expr {max($col($ao),$pmiCol)}]
                    }
                  }
                  "polyline name" {set geomType "polyline"}
                  "circle name"   {set geomType "circle"}
                  "composite_curve name" {set iCompCurveSeg 0}
                }

# value in spreadsheet
                if {$ok && [info exists gpmiID]} {
                  set c [string index [cellRange 1 $col($ao)] 0]
                  set r $gpmiIDRow($ao,$gpmiID)

# column name
                  if {$colName != ""} {
                    if {![info exists pmiHeading($col($ao))]} {
                      $cells($ao) Item 3 $c $colName
                      set pmiHeading($col($ao)) 1
                      set pmiCol [expr {max($col($ao),$pmiCol)}]
                    }
                  }

# keep track of rows with validation properties
                  if {[lsearch $gpmiRow($ao) $r] == -1} {lappend gpmiRow($ao) $r}

# look for correct PMI name on 
# geometric_curve_set  annotation_fill_area  tessellated_geometric_set  composite_curve
                  if {$ent1 == "geometric_curve_set name"  || \
                      $ent1 == "annotation_fill_area name" || \
                      $ent1 == "tessellated_geometric_set name" || \
                      $ent1 == "repositioned_tessellated_item_and_tessellated_geometric_set name" || \
                      $ent1 == "composite_curve name"} {
                    set ov $objValue


# look for invalid 'name' values                  
                    set invalid 0
                    if {$ov == ""} {
                      set msg "Syntax Error: Missing 'name' attribute on [lindex $ent1 0].\n[string repeat " " 14]"
                      if {$stepAP == "AP242"} {
                        append msg "($recPracNames(pmi242), Sec. 8.4, Table 14)"
                      } else {
                        append msg "($recPracNames(pmi203), Sec. 4.3, Table 1)"
                      }
                      errorMsg $msg
                      set ov "(blank)"
                      if {[info exists gpmiTypesInvalid]} {
                        if {[lsearch $gpmiTypesInvalid $ov] == -1} {lappend gpmiTypesInvalid $ov}
                      }
                      set invalid 1
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1]]
                    } elseif {[lsearch $gpmiTypes $ov] == -1} {
                      set msg "Syntax Error: Invalid 'name' attribute ($ov) on [lindex $ent1 0].\n[string repeat " " 14]"
                      if {$stepAP == "AP242"} {
                        append msg "($recPracNames(pmi242), Sec. 8.4, Table 14)"
                      } else {
                        append msg "($recPracNames(pmi203), Sec. 4.3, Table 1)"
                      }
                      errorMsg $msg
                      if {[info exists gpmiTypesInvalid]} {
                        if {[lsearch $gpmiTypesInvalid $ov] == -1} {lappend gpmiTypesInvalid $ov}
                      }
                      set invalid 1
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1]]
                    }
                    
# count number of gpmi types
                    #outputMsg "$ov  $aoname" green
                    if {$ov != $aoname} {
                      lappend gpmiTypesPerFile "$ov/$aoname"
                    } else {
                      set n 0
                      set objGuiEntities [$objEntity1 GetUsedIn [string trim draughting_model_item_association] [string trim identified_item]]
                      ::tcom::foreach objGuiEntity $objGuiEntities {
                        incr n
                        if {$n == 1} {lappend gpmiTypesPerFile "$ov/$aoname[$objEntity1 P21ID]"}
                        #lappend gpmiTypesPerFile "$ov/$aoname[$objGuiEntity P21ID]"
                      }
                      if {$n == 0} {
                        set objGuiEntities [$objEntity1 GetUsedIn [string trim draughting_callout] [string trim contents]]
                        ::tcom::foreach objGuiEntity $objGuiEntities {
                          incr n
                          if {$n == 1} {lappend gpmiTypesPerFile "$ov/$aoname[$objGuiEntity P21ID]"}
                        }
                      }
                    }
                    #catch {outputMsg $gpmiTypesPerFile red}
                
# start X3DOM file
                    if {$opt(VIZPMI)} {
                      if  {[string first "tessellated" $ao] == -1} {
                        if {$x3domFileOpen} {
                          set x3domFileOpen 0
                          set x3domFileName [file rootname $localName]_x3dom.html
                          catch {file delete -force $x3domFileName}
                          set x3domFile [open $x3domFileName w]
                          outputMsg " Writing PMI Annotations to: [truncFileName [file nativename $x3domFileName]]" green
                          
                          set str "NIST "
                          set url "https://go.usa.gov/yccx"
                          if {!$nistVersion} {
                            set str ""
                            set url "https://github.com/usnistgov/SFA"
                          }
                          
                          puts $x3domFile "<!DOCTYPE html>\n<html>\n<head>\n<title>[file tail $localName] | PMI Annotations</title>\n<base target=\"_blank\">\n<meta http-equiv='Content-Type' content='text/html;charset=utf-8'/>\n<link rel='stylesheet' type='text/css' href='https://www.x3dom.org/x3dom/release/x3dom.css'/>\n<script type='text/javascript' src='https://www.x3dom.org/x3dom/release/x3dom.js'></script>\n</head>"
                          puts $x3domFile "\n<body><font face=\"arial\">\n<h3>PMI Annotations for:  [file tail $localName]</h3>"
                          puts $x3domFile "<ul><li>Visualization of PMI annotations generated by the <a href=\"$url\">$str\STEP File Analyzer (v[getVersion])</a> and rendered with <a href=\"https://www.x3dom.org/\">X3DOM</a>."
                          puts $x3domFile "<li>Only the PMI annotations are shown.  Part geometry can be viewed with <a href=\"https://www.cax-if.org/step_viewers.html\">STEP file viewers</a>."
                          puts $x3domFile "<li><a href=\"https://www.x3dom.org/documentation/interaction/\">Use the mouse</a> to rotate, pan, and zoom.  Use Page Down to switch between perspective and orthographic views.  Saved Views are ignored."
                          puts $x3domFile "</ul><table><tr><td>"

# x3d window size
                          set height 800
                          set width [expr {int($height*1.5)}]
                          catch {
                            set height [expr {int([winfo screenheight .]*0.7)}]
                            set width [expr {int($height*[winfo screenwidth .]/[winfo screenheight .])}]
                          }
                          puts $x3domFile "\n<X3D id='someUniqueId' showStat='false' showLog='false' x='0px' y='0px' width='$width\px' height='$height\px'>\n<Scene DEF='scene'>"

                          for {set i 0} {$i < 4} {incr i} {set avgX3domColor($i) 0}
                        }

# start X3DOM Shape node                    
                        if {$ao == "annotation_fill_area_occurrence"} {errorMsg "PMI annotations with filled characters are not filled."}
                        if {$x3domColor != ""} {
                          puts $x3domFile "<Shape id='ao$objID'>\n <Appearance><Material emissiveColor='$x3domColor'></Material></Appearance>"
                          set colors [split $x3domColor " "]
                          for {set i 0} {$i < 3} {incr i} {set avgX3domColor($i) [expr {$avgX3domColor($i)+[lindex $colors $i]}]}
                          incr avgX3domColor(3)
                        } elseif {[string first "annotation_occurrence" $ao] == 0} {
                          puts $x3domFile "<Shape id='ao$objID'>\n <Appearance><Material emissiveColor='1 0.5 0'></Material></Appearance>"
                          errorMsg "Syntax Error: Color not specified for PMI Presentation (using orange)"
                        } elseif {[string first "annotation_fill_area_occurrence" $ao] == 0} {
                          puts $x3domFile "<Shape id='ao$objID'>\n <Appearance><Material emissiveColor='1 0.5 0'></Material></Appearance>"
                          errorMsg "Syntax Error: Color not specified for PMI Presentation (using orange)"
                        }
                        set x3domShape 1
                        update idletasks
                      } else {
                        errorMsg " Visualization of Tessellated PMI Annotations is not supported." red
                      }
                    }               

# value in spreadsheet  
                    set val [[$cells($ao) Item $r $c] Value]
                    if {$invalid} {lappend syntaxErr($ao) [list $r $col($ao)]}
  
                    if {$val == ""} {
                      $cells($ao) Item $r $c $ov
                    } else {
                      $cells($ao) Item $r $c "$val[format "%c" 10]$ov"
                    }

# keep track of max column
                    set pmiCol [expr {max($col($ao),$pmiCol)}]

# keep track of cartesian point ids (x3domPointID)
                  } elseif {[info exists currX3domPointID] && $ent1 == "cartesian_point name"} {
                    if {$currX3domPointID < $maxcp} {append x3domPointID "$objID "}
                    incr currX3domPointID

# cell value for presentation style or color
                  } else {
                    if {$colName != "colour"} {
                      $cells($ao) Item $r $c "$ent($entLevel) $objID"
                    } else {
                      if {$ent($entLevel) == "colour_rgb"} {
                        $cells($ao) Item $r $c "$ent($entLevel) $objID  ($x3domColor)"
                      } else {
                        $cells($ao) Item $r $c "$ent($entLevel) $objID  ($objValue)"
                      }
                    }
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing PMI Presentation ($objNodeType $ent2): $emsg3"
            set entLevel 1
          }
        }
      }
    }
  }
  incr entLevel -1
  
# write a few more things at the end of processing an annotation_occurrence entity
  if {$entLevel == 0} {

# associated geometry, (1) find link between annotation_occurrence and a geometric item through
# draughting_model_item_association or draughting_callout and geometric_item_specific_usage
# (2) find link to PMI representation
    if {[catch {
      catch {unset assocGeom}
      catch {unset assocSPMI}
      
      if {[string first "placeholder" $ao] == -1} {
        set ok 0
        set objDC [$objEntity GetUsedIn [string trim draughting_callout] [string trim contents]]
        ::tcom::foreach objGuiEntity $objDC {
          set objGuiEntities [$objGuiEntity GetUsedIn [string trim draughting_model_item_association] [string trim identified_item]]
          set ok 1
        }
        if {!$ok} {set objGuiEntities [$objEntity GetUsedIn [string trim draughting_model_item_association] [string trim identified_item]]}
      } else {
        set ok 0
        set objDC [$objEntity GetUsedIn [string trim draughting_callout] [string trim contents]]
        ::tcom::foreach objGuiEntity $objDC {
          set objGuiEntities [$objGuiEntity GetUsedIn [string trim draughting_model_item_association_with_placeholder [string trim identified_item]]
          set ok 1
        }
        if {!$ok} {set objGuiEntities [$objEntity GetUsedIn [string trim draughting_model_item_association_with_placeholder] [string trim identified_item]]}
      }
      
      ::tcom::foreach objGuiEntity $objGuiEntities {
        #outputMsg " [$objGuiEntity Type] [$objGuiEntity P21ID]  (ao [$objEntity P21ID])" red
        ::tcom::foreach attrDMIA [$objGuiEntity Attributes] {
          if {[$attrDMIA Name] == "name"} {set attrName [$attrDMIA Value]}

# get shape_aspect (dmiaDef)
          if {[$attrDMIA Name] == "definition"} {
            set dmiaDef [$attrDMIA Value]
            if {[string first "handle" $dmiaDef] != -1} {
              set dmiaDefType [$dmiaDef Type]
              #outputMsg "  $dmiaDefType [$dmiaDef P21ID]  $attrName" green

# look for link to pmi representation
              if {$attrName == "PMI representation to presentation link"} {
                if {[string first "shape_aspect" $dmiaDefType] == -1} {
                  set spmi_p21id [$dmiaDef P21ID]
                  if {![info exists assocSPMI($dmiaDefType)]} {
                    lappend assocSPMI($dmiaDefType) $spmi_p21id
                  } elseif {[lsearch $assocSPMI($dmiaDefType) $spmi_p21id] == -1} {
                    lappend assocSPMI($dmiaDefType) $spmi_p21id
                  }
                }

# look at shape_aspect or datums to find associated geometry
              } elseif {[string first "shape_aspect" $dmiaDefType] != -1 || [string first "datum" $dmiaDefType] != -1} {
                #outputMsg "   $dmiaDefType [$dmiaDef P21ID]  $attrName" blue
                getAssocGeom $dmiaDef 1
              } else {
                #outputMsg "   $dmiaDefType [$dmiaDef P21ID]  $attrName" red
              }
            } else {
              set msg "Syntax Error: Missing 'definition' attribute on \#[$objGuiEntity P21ID]=draughting_model_item_association\n[string repeat " " 14]"
              if {$stepAP == "AP242"} {
                append msg "($recPracNames(pmi242), Sec. 9.3.1, Fig. 76)"
              } else {
                append msg "($recPracNames(pmi203), Sec. 5.3.1, Fig. 12)"
              }
              errorMsg $msg
            }
          } elseif {[$attrDMIA Name] == "used_representation"} {
            set dmiaDef [$attrDMIA Value]
            if {[string first "handle" $dmiaDef] != -1} {
              set dmiaDefType [$dmiaDef Type]
              if {[string first "draughting_model" $dmiaDefType] == -1} {
                errorMsg "Syntax Error: Invalid 'used_representation' attribute ($dmiaDefType) on draughting_model_item_association"
              }
            }
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding Associated Geometry: $emsg"
    }

# report annotation plane
    if {[catch {
      #outputMsg "Annotation Plane" red
      set aps {}
      set ents [$objEntity GetUsedIn [string trim annotation_plane] [string trim elements]]
      ::tcom::foreach ap $ents {lappend aps $ap}
      if {[llength $aps] == 0} {
        set ents [$objEntity GetUsedIn [string trim draughting_callout] [string trim contents]]
        ::tcom::foreach dc $ents {
          set ents1 [$dc GetUsedIn [string trim annotation_plane] [string trim elements]]
        }
        if {[info exists ents1]} {::tcom::foreach ap $ents1 {lappend aps $ap}}
      }
      foreach ap $aps {
        ::tcom::foreach attrAP [$ap Attributes] {
          if {[$attrAP Name] == "name"} {
            set str "[$ap Type] [$ap P21ID]"
            set nam [$attrAP Value]
            if {$nam != ""} {append str "[format "%c" 10]($nam)"}
            if {![info exists pmiColumns(aplane)]} {set pmiColumns(aplane) [getNextUnusedColumn $ao 3]}
            if {$stepAP == "AP242"} {
              set colName "plane[format "%c" 10](Sec. 9.1)"
            } else {
              set colName "plane[format "%c" 10](Sec. 5.1)"
            }
            set c [string index [cellRange 1 $pmiColumns(aplane)] 0]
            set r $gpmiIDRow($ao,$gpmiID)
            if {![info exists pmiHeading($pmiColumns(aplane))]} {
              $cells($ao) Item 3 $c $colName
              set pmiHeading($pmiColumns(aplane)) 1
              set pmiCol [expr {max($pmiColumns(aplane),$pmiCol)}]
            }
            $cells($ao) Item $r $pmiColumns(aplane) [string trim $str]
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR reporting Annotation Plane: $emsg"
    }
      
# report associated geometry
    if {[catch {
      if {[info exists assocGeom]} {
        set str [reportAssocGeom $ao 0]
        if {$str != ""  } {
          #outputMsg "  Adding Associated Geometry" green
          if {![info exists pmiColumns(ageom)]} {set pmiColumns(ageom) [getNextUnusedColumn $ao 3]}
          if {$stepAP == "AP242"} {
            set colName "Associated Geometry[format "%c" 10](Sec. 9.3.1)"
          } else {
            set colName "Associated Geometry[format "%c" 10](Sec. 5.3.1)"
          }
          set c [string index [cellRange 1 $pmiColumns(ageom)] 0]
          set r $gpmiIDRow($ao,$gpmiID)
          if {![info exists pmiHeading($pmiColumns(ageom))]} {
            $cells($ao) Item 3 $c $colName
            set pmiHeading($pmiColumns(ageom)) 1
            set pmiCol [expr {max($pmiColumns(ageom),$pmiCol)}]
          }
          $cells($ao) Item $r $pmiColumns(ageom) [string trim $str]
        }
      }

# report associated semantic PMI
      if {[info exists assocSPMI]} {
        set str ""
        set nspmi 0
        foreach item [array names assocSPMI] {
          if {[string length $str] > 0} {append str [format "%c" 10]}
          append str "([llength $assocSPMI($item)]) [formatComplexEnt $item 1] [lsort -integer $assocSPMI($item)]"
          incr nspmi [llength $assocSPMI($item)]
        }
        if {$nspmi == 1} {set str [string range $str 4 end]}
        if {$str != ""} {
          #errorMsg "  Adding Associated Representation" green
          if {![info exists pmiColumns(spmi)]} {set pmiColumns(spmi) [getNextUnusedColumn $ao 3]}
          set colName "Associated Representation[format "%c" 10](Sec. 7.3)"
          set c [string index [cellRange 1 $pmiColumns(spmi)] 0]
          set r $gpmiIDRow($ao,$gpmiID)
          if {![info exists pmiHeading($pmiColumns(spmi))]} {
            $cells($ao) Item 3 $c $colName
            set pmiHeading($pmiColumns(spmi)) 1
            set pmiCol [expr {max($pmiColumns(spmi),$pmiCol)}]
          }
          $cells($ao) Item $r $pmiColumns(spmi) [string trim $str]
        }
      }
    } emsg]} {
      errorMsg "ERROR reporting Associated Geometry and Representation: $emsg"
    }

# report camera models associated with the annotation_occurrence through draughting_model
    if {[catch {
      set savedViews ""    
      set nsv 0
      if {[info exists draftModelCameras]} {
        set dmlist {}
        foreach dms [list draughting_model characterized_object_and_draughting_model characterized_representation_and_draughting_model characterized_representation_and_draughting_model_and_representation] {
          if {[info exists entCount($dms)]} {if {$entCount($dms) > 0} {lappend dmlist $dms}}
        }
        foreach dm $dmlist {
          set entDraughtingModels [$objEntity GetUsedIn [string trim $dm] [string trim items]]
          
# check for draughting_callout.contents -> ao (PMI RP, section 9.4.4, figure 93)
          set entDraughtingCallouts [$objEntity GetUsedIn [string trim draughting_callout] [string trim contents]]
          ::tcom::foreach entDraughtingCallout $entDraughtingCallouts {
            set entDraughtingModels [$entDraughtingCallout GetUsedIn [string trim $dm] [string trim items]]
            #outputMsg [$entDraughtingCallout P21ID][$entDraughtingCallout Type] green
          }
      
          ::tcom::foreach entDraughtingModel $entDraughtingModels {
            if {[info exists draftModelCameras([$entDraughtingModel P21ID])]} {
              #outputMsg "[$entDraughtingModel P21ID] [$entDraughtingModel Type]" red
              set str $draftModelCameras([$entDraughtingModel P21ID])
              if {[string first $str $savedViews] == -1} {
                append savedViews $draftModelCameras([$entDraughtingModel P21ID])
                incr nsv
              }
              #errorMsg "  Adding Saved Views" green
              if {$stepAP == "AP242"} {
                set colName "Saved Views[format "%c" 10](Sec. 9.4)"
              } else {
                set colName "Saved Views[format "%c" 10](Sec. 5.4)"
              }
              if {![info exists savedViewCol]} {set savedViewCol [getNextUnusedColumn $ao 3]}
              set c [string index [cellRange 1 $savedViewCol] 0]
              set r $gpmiIDRow($ao,$gpmiID)
              if {![info exists pmiHeading($savedViewCol)]} {
                $cells($ao) Item 3 $c $colName
                set pmiHeading($savedViewCol) 1
                set pmiCol [expr {max($savedViewCol,$pmiCol)}]
              }
              set str "($nsv) camera_model_d3 [string trim $savedViews]"
              if {$nsv == 1} {set str "camera_model_d3 [string trim $savedViews]"}
              $cells($ao) Item $r $c $str
              if {[string first "()" $savedViews] != -1} {lappend syntaxErr($ao) [list $r $savedViewCol]}
              
# check for a mapped_item in draughting_model.items
              set attrsDraughtingModel [$entDraughtingModel Attributes]
              ::tcom::foreach attrDraughtingModel $attrsDraughtingModel {
                if {[$attrDraughtingModel Name] == "name"} {
                  set nameDraughtingModel [$attrDraughtingModel Value]
                }
                if {[$attrDraughtingModel Name] == "items" && $nameDraughtingModel != ""} {
                  set okcm 0
                  set okmi 0
                  ::tcom::foreach item [$attrDraughtingModel Value] {
                    if {[$item Type] == "mapped_item"} {set okmi 1}
                    if {[string first "camera_model_d3" [$item Type]] == 0} {set okcm 1}
                  }
                  if {$okcm && $okmi == 0} {
                    set msg "Syntax Error: For Saved Views, [formatComplexEnt [$entDraughtingModel Type]].items missing required reference to 'mapped_item'\n[string repeat " " 14]"
                    if {$stepAP == "AP242"} {
                      append msg "($recPracNames(pmi242), Sec. 9.4.2.1, Fig. 86)"
                    } else {
                      append msg "($recPracNames(pmi203), Sec. 5.4.2, Fig. 14)"
                    }
                    errorMsg $msg
                    lappend syntaxErr([$entDraughtingModel Type]) [list [$entDraughtingModel P21ID] items]
                  }
                }
              }
          
# check MDADR (or RR) rep_1 vs. rep_2
              if {$stepAP == "AP242"} {
                set relType ""
                foreach item [list mechanical_design_and_draughting_relationship representation_relationship] {
                  if {[info exists entCount($item)]} {if {$entCount($item) > 0} {set relType $item}}
                }
                if {$relType != ""} {
                  set ok 1
                  set rep1Ents [$entDraughtingModel GetUsedIn [string trim $relType] [string trim rep_1]]
                  ::tcom::foreach rep1Ent $rep1Ents {set ok 0}
                  if {$ok} {
                    set msg "Syntax Error: For Saved Views, '$relType' reference to '[formatComplexEnt [$entDraughtingModel Type]]' uses rep_2 instead of rep_1\n[string repeat " " 14]"
                    append msg "($recPracNames(pmi242), Sec. 9.4.4 Note 1, Fig. 93, Table 16)"
                    errorMsg $msg
                    set rep2Ents [$entDraughtingModel GetUsedIn [string trim $relType] [string trim rep_2]]
                    ::tcom::foreach rep2Ent $rep2Ents {set mdadrID [$rep2Ent P21ID]}
                    lappend syntaxErr($relType) [list $mdadrID rep_2]
                  }
                  if {$relType == "representation_relationship"} {
                    set msg "For Saved Views, recommend using 'mechanical_design_and_draughting_relationship' instead of 'representation_relationship'\n  to relate draughting models "
                    if {$stepAP == "AP242"} {
                      append msg "($recPracNames(pmi242), Sec. 9.4.4 Note 2)"
                    } else {
                      append msg "($recPracNames(pmi203), Sec. 5.4.4 Note 2)"
                    }
                    errorMsg $msg
                  }
                }
              }
            }
          }
        }
        catch {unset entDraughtingModel}
        catch {unset entDraughtingModels}

# missing MDADR or RR
        set relTypes [list representation_relationship]
        if {$stepAP != "AP214"} {lappend relTypes mechanical_design_and_draughting_relationship}
        set relType ""
        foreach item $relTypes {if {[info exists entCount($item)]} {if {$entCount($item) > 0} {set relType $item}}}
        if {$relType == ""} {
          set str "mechanical_design_and_draughting_relationship"
          if {$stepAP == "AP214"} {set str "representation_relationship"}
          set msg "Syntax Error: For Saved Views, missing '$str' to relate 'draughting_model'\n[string repeat " " 14]"
          if {$stepAP == "AP242"} {
            append msg "($recPracNames(pmi242), Sec. 9.4.4, Fig. 93)"
          } else {
            append msg "($recPracNames(pmi203), Sec. 5.4.4, Fig. 20)"
          }
          errorMsg $msg
        }
      }
    } emsg]} {
      errorMsg "ERROR adding Saved Views: $emsg"
    }

# check if there are PMI validation properties (propDefIDS) associated with the annotation_occurrence
    if {[catch {
      if {[info exists propDefIDS]} {

# look for annotation_occurrence used in property_definition.definition
        set objGuiEntities [$objEntity GetUsedIn [string trim property_definition] [string trim definition]]
        ::tcom::foreach objGuiEntity $objGuiEntities {
          set GuiID [$objGuiEntity P21ID]
          foreach item [lsort [array names propDefIDS]] {
            if {$item == $GuiID && $propDefIDS($item) == $objID} {append gpmiValProp($objID) "$GuiID "}
          }
        }

# look for annotation_occurrence used in characterized_item_within_representation.item
        set objGuiEntities [$objEntity GetUsedIn [string trim characterized_item_within_representation] [string trim item]]
        ::tcom::foreach objGuiEntity $objGuiEntities {
          set GuiID [$objGuiEntity P21ID]
          foreach item [lsort [array names propDefIDS]] {
            if {$propDefIDS($item) == $GuiID} {append gpmiValProp($objID) "$item "}
          }
        }

# look for annotation_occurrence used in draughting_callout.contents
        set objGuiEntities [$objEntity GetUsedIn [string trim draughting_callout] [string trim contents]]
        ::tcom::foreach objGuiEntity $objGuiEntities {
          set refGuiEntities [$objGuiEntity GetUsedIn [string trim characterized_item_within_representation] [string trim item]]
          ::tcom::foreach refGuiEntity $refGuiEntities {
            foreach item [lsort [array names propDefIDS]] {
              if {$propDefIDS($item) == [$refGuiEntity P21ID]} {append gpmiValProp($objID) "$item "}
            }
          }
        }
      }

# add valprop info to spreadsheet
      if {[info exists gpmiValProp($objID)]} {
        if {![info exists pmiColumns(vp)]} {set pmiColumns(vp) [getNextUnusedColumn $ao 3]}
        #errorMsg "  Adding Validation Properties" green
        if {$stepAP == "AP242"} {
          set colName "Validation Properties[format "%c" 10](Sec. 10.3)"
        } else {
          set colName "Validation Properties[format "%c" 10](Sec. 6.3)"
        }
        set c [string index [cellRange 1 $pmiColumns(vp)] 0]
        set r $gpmiIDRow($ao,$gpmiID)
        if {![info exists pmiHeading($pmiColumns(vp))]} {
          #$cells($ao) Item 1 $c $colName
          $cells($ao) Item 3 $c $colName
          set pmiHeading($pmiColumns(vp)) 1
          set pmiCol [expr {max($pmiColumns(vp),$pmiCol)}]
        }
        set str "([llength $gpmiValProp($objID)]) property_definition [string trim $gpmiValProp($objID)]"
        if {[llength $gpmiValProp($objID)] == 1} {set str "property_definition [string trim $gpmiValProp($objID)]"}
        $cells($ao) Item $r $pmiColumns(vp) $str
      }
    } emsg]} {
      errorMsg "ERROR adding PMI Presentation Validation Properties: $emsg"
    }
  }
}

# -------------------------------------------------------------------------------
# set X3DOM color
proc gpmiSetColor {type} {
  global idxColor

  if {$type == 1} {return "0 0 0"}

  if {$type == 2} {
    incr idxColor
    switch $idxColor {
      1 {set color "0 0 0"}
      2 {set color "1 0 0"}
      3 {set color "1 1 0"}
      4 {set color "0 1 0"}
      5 {set color "0 1 1"}
      6 {set color "0 0 1"}
      7 {set color "1 0 1"}
      8 {set color "1 1 1"}
    }
    if {$idxColor == 8} {set idxColor 0}
  }
  return $color
} 

# -------------------------------------------------------------------------------
# get camera models and validation properties
proc pmiGetCamerasAndProperties {objDesign} {
  global draftModelCameras gpmiValProp syntaxErr propDefIDS stepAP recPracNames entCount

  #outputMsg getCameras blue
  catch {unset draftModelCameras}
  if {[catch {
    set cmlist {}
    foreach cms [list camera_model_d3 camera_model_d3_multi_clipping] {
      if {[info exists entCount($cms)]} {if {$entCount($cms) > 0} {lappend cmlist $cms}}
    }
    set dmlist {}
    foreach dms [list draughting_model characterized_object_and_draughting_model characterized_representation_and_draughting_model characterized_representation_and_draughting_model_and_representation] {
      if {[info exists entCount($dms)]} {if {$entCount($dms) > 0} {lappend dmlist $dms}}
    }

# loop over camera model entities
    foreach cm $cmlist {
      ::tcom::foreach entCameraModel [$objDesign FindObjects [string trim $cm]] {
        set attrCameraModels [$entCameraModel Attributes]
  
# loop over draughting model entities
        foreach dm $dmlist {
          set entDraughtingModels [$entCameraModel GetUsedIn [string trim $dm] [string trim items]]
          ::tcom::foreach entDraughtingModel $entDraughtingModels {
            ::tcom::foreach attrCameraModel $attrCameraModels {
              set nameCameraModel [$attrCameraModel Name]
              if {$nameCameraModel == "name"} {set name [$attrCameraModel Value]}
              if {$name == ""} {
                set msg "Syntax Error: For Saved Views, required 'name' attribute on camera_model_d3 is blank.\n[string repeat " " 14]"
                if {$stepAP == "AP242"} {
                  append msg "($recPracNames(pmi242), Sec. 9.4.2, Fig. 76)"
                } else {
                  append msg "($recPracNames(pmi203), Sec. 5.4.2, Fig. 14)"
                }
                errorMsg $msg
                lappend syntaxErr($cm) [list [$entCameraModel P21ID] name]
              }
            }
  
            set str "[$entCameraModel P21ID] ($name)  "
            set id [$entDraughtingModel P21ID]
            if {![info exists draftModelCameras($id)]} {
              set draftModelCameras($id) $str
            } elseif {[string first $str $draftModelCameras($id)] == -1} {
              append draftModelCameras($id) "[$entCameraModel P21ID] ($name)  "
            }
            #outputMsg "$id  $draftModelCameras($id)  $dm  $cm"
          }
        }
      }
    }
  } emsg]} {
    errorMsg "ERROR getting Camera Models: $emsg"
    catch {raise .}
  }
  
# get pmi validation properties so that they can be annotation_occurrence
  catch {unset gpmiValProp}
  catch {unset propDefIDS}
  
  if {[catch {
    ::tcom::foreach objPDEntity [$objDesign FindObjects [string trim property_definition]] {
      set objPDAttributes [$objPDEntity Attributes]
      set idx ""
      ::tcom::foreach objPDAttribute $objPDAttributes {
        set objPDName [$objPDAttribute Name]
        if {$objPDName == "name" && [string first "pmi validation property" [$objPDAttribute Value]] != -1} {
          set idx [$objPDEntity P21ID]
        } elseif {$objPDName == "definition" && $idx != "" && [string first "handle" [$objPDAttribute Value]] != -1} {
          set propDefIDS($idx) [[$objPDAttribute Value] P21ID]
        }
      }
    }
  } emsg]} {
    errorMsg "ERROR getting PMI validation properities: $emsg"
    catch {raise .}
  }
}  

# -------------------------------------------------------------------------------
# set viewpoints, add navigation and background color, and close X3DOM file
proc x3domViewpoints {} {
  global avgX3domColor opt stepAP x3domMax x3domMin x3domFile x3domMsg stepAP entCount
  
# viewpoints
  foreach idx {x y z} {
    set delt($idx) [expr {$x3domMax($idx)-$x3domMin($idx)}]
    set xyzcen($idx) [format "%.4f" [expr {0.5*$delt($idx) + $x3domMin($idx)}]]
  }
  set maxxyz $delt(x)
  if {$delt(y) > $maxxyz} {set maxxyz $delt(y)}
  if {$delt(z) > $maxxyz} {set maxxyz $delt(z)}
  set cor "centerOfRotation='$xyzcen(x) $xyzcen(y) $xyzcen(z)'"
  puts $x3domFile "\n<Viewpoint $cor position='$xyzcen(x) [trimNum [expr {0. - ($xyzcen(y) + 1.4*$maxxyz)}]] $xyzcen(z)' orientation='1 0 0 1.5708' description='Front'></Viewpoint>"
  set fov [trimNum [expr {$delt(z)*0.5 + $delt(y)*0.5}]]
  puts $x3domFile "<OrthoViewpoint fieldOfView='\[-$fov,-$fov,$fov,$fov\]' $cor position='$xyzcen(x) [trimNum [expr {0. - ($xyzcen(y) + 1.4*$maxxyz)}]] $xyzcen(z)' orientation='1 0 0 1.5708' description='Ortho'></OrthoViewpoint>"

# old viewpoints
  #puts $x3domFile "<Viewpoint $cor position='$xyzcen(x) $xyzcen(y) [trimNum [expr {$xyzcen(z) + 1.4*$maxxyz}]]' description='Top'></Viewpoint>"
  #puts $x3domFile "<Viewpoint $cor position='[trimNum [expr {$xyzcen(x) + 1.4*$maxxyz}]] $xyzcen(y) $xyzcen(z)' orientation='0 1 0 1.5708' description='Side 1'></Viewpoint>"
  #puts $x3domFile "<Viewpoint $cor position='[trimNum [expr {$xyzcen(x) + 1.4*$maxxyz}]] $xyzcen(y) $xyzcen(z)' orientation='1 1 1 2.0944' description='Side 2'></Viewpoint>"

  #puts $x3domFile "<viewpoint position='$xyzcen(x) $xyzcen(y) [trimNum [expr {$xyzcen(z) + 1.4*$maxxyz}]]' description='Front'></viewpoint>"
  #puts $x3domFile "<viewpoint position='[trimNum [expr {0. - ($xyzcen(x) + 1.4*$maxxyz)}]] $xyzcen(y) $xyzcen(z)' orientation='0 1 0 -1.5708' description='Left'></viewpoint>"
  #puts $x3domFile "<viewpoint position='$xyzcen(x) $xyzcen(y) [trimNum [expr {0. - ($xyzcen(z) + 1.4*$maxxyz)}]]' orientation='0 1 0 3.1416' description='Back'></viewpoint>"
  #puts $x3domFile "<viewpoint position='[trimNum [expr {$xyzcen(x) + 1.4*$maxxyz}]] $xyzcen(y) $xyzcen(z)' orientation='0 1 0 1.5708' description='Right'></viewpoint>"
  #puts $x3domFile "<viewpoint position='$xyzcen(x) [trimNum [expr {$xyzcen(y) + 1.4*$maxxyz}]] $xyzcen(z)' orientation='-1 0 0 1.5708' description='Top'></viewpoint>"
  #puts $x3domFile "<viewpoint position='$xyzcen(x) [trimNum [expr {0. - ($xyzcen(y) + 1.4*$maxxyz)}]] $xyzcen(z)' orientation='-1 0 0 -1.5708' description='Bottom'></viewpoint>"
  
  puts $x3domFile "<NavigationInfo type='\"EXAMINE\" \"ANY\"'></NavigationInfo>"

# find average color
  set ok 0
  if {[string first "AP209" $stepAP] == -1} {
    for {set i 0} {$i < 3} {incr i} {set avgX3domColor($i) [expr {$avgX3domColor($i)/$avgX3domColor(3)}]}
    set acolor [expr {$avgX3domColor(0)+$avgX3domColor(1)+$avgX3domColor(2)}]
    if {($acolor < 0.2 || $opt(gpmiColor) == 1) && $opt(gpmiColor) != 2} {set ok 1}

    if {[info exists entCount(tessellated_annotation_occurrence)]} {lappend x3domMsg "Segments of the PMI annotations modeled with tessellated geometry are not displayed."}

# AP209 color
  } else {
    set ok 1
  }

# background color
  if {$ok} {
    puts $x3domFile "<Background skyColor='1. 1. 1.'></Background>"
  } elseif {$acolor < 2.0} {
    puts $x3domFile "<Background skyColor='.8 .8 .8'></Background>"
  } else {
    puts $x3domFile "<Background skyColor='.4 .4 .4'></Background>"
  }

  puts $x3domFile "</Scene></X3D>\n\n</td></tr></table>\n<p>"
  
# extra messages
  if {[info exists x3domMsg]} {
    if {[llength $x3domMsg] > 0} {
      puts $x3domFile "<ul>"
      foreach item $x3domMsg {puts $x3domFile "<li>$item"}
      puts $x3domFile "</ul>"
      unset x3domMsg
    }
  }
  
  puts $x3domFile "[clock format [clock seconds]]"

  puts $x3domFile "</font></body></html>"
  close $x3domFile
  update idletasks
  
  unset x3domMax
  unset x3domMin
  catch {unset avgX3domColor}
}

# -------------------------------------------------------------------------------------------------
# open X3DOM file 
proc openX3DOM {} {
  global opt x3domFileName multiFile stepAP
  
  if {($opt(VIZPMI) || $opt(VIZFEA)) && $x3domFileName != "" && $multiFile == 0} {
    set str "PMI Presentation Annotations"
    if {$stepAP == "AP209"} {set str "Analysis Model"}
    outputMsg "Opening $str in the default Web Browser" blue
    if {[catch {
      exec {*}[auto_execok start] "" $x3domFileName
    } emsg]} {
      errorMsg "No application is associated with HTML files.  Open the file in a web browser that supports X3DOM.  https://www.x3dom.org/check/\n $emsg"
    }
    update idletasks
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
  global worksheet worksheet1
  #outputMsg "gpmiCoverageWrite $multi" red

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
    if {![info exists pmi_rows]} {set pmi_rows 33}
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
              [$worksheet1($pmi_coverage) Range [cellRange $r1 1] [cellRange $r1 1]] Style "Bad"
            } else {
              $cells($pmi_coverage) Item $r1 1 $idx
              [$worksheet($pmi_coverage) Range [cellRange $r1 1] [cellRange $r1 1]] Style "Bad"
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
                [$worksheet1($pmi_coverage) Range [cellRange $r1 1] [cellRange $r1 1]] Style "Bad"
              } else {
                $cells($pmi_coverage) Item $r1 1 $idx
                [$worksheet($pmi_coverage) Range [cellRange $r1 1] [cellRange $r1 1]] Style "Bad"
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
      for {set r 4} {$r <= 100} {incr r} {
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
  global pmi_coverage pmi_rows pmi_totals recPracNames worksheet worksheet1
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
    set idx1 [list 21 28 30 34]
    if {!$multi} {set idx1 [list 3 4 21 28 30 34]}
    for {set r 100} {$r >= 34} {incr r -1} {
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

# vertical line(s)
    if {$multi} {
      set range [$worksheet1($pmi_coverage) Range [cellRange 1 $col1($pmi_coverage)] [cellRange [expr {[lindex $idx1 end]-1}] $col1($pmi_coverage)]]
      catch {[[$range Borders] Item [expr 7]] Weight [expr 2]}
      
# fix row 3 height and width
      set range [$worksheet1($pmi_coverage) Range 3:3]
      $range RowHeight 300
      [$worksheet1($pmi_coverage) Columns] AutoFit
      
      $cells1($pmi_coverage) Item [expr {$pmi_rows+2}] 1 "Presentation Names defined in $recPracNames(pmi242), Sec. 8.4, Table 14"
      set anchor [$worksheet1($pmi_coverage) Range [cellRange [expr {$pmi_rows+2}] 1]]
      [$worksheet1($pmi_coverage) Hyperlinks] Add $anchor [join "https://www.cax-if.org/joint_testing_info.html#recpracs"] [join ""] [join "Link to CAx-IF Recommended Practices"]
  
      [$worksheet1($pmi_coverage) Rows] AutoFit
      [$worksheet1($pmi_coverage) Range "B4"] Select
      [$excel1 ActiveWindow] FreezePanes [expr 1]
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
      
      catch {$cells($pmi_coverage) Item 1 5 "Presentation Names defined in $recPracNames(pmi242), Sec. 8.4, Table 14"}
      set range [$worksheet($pmi_coverage) Range E1:O1]
      $range MergeCells [expr 1]
      set anchor [$worksheet($pmi_coverage) Range E1]
      [$worksheet($pmi_coverage) Hyperlinks] Add $anchor [join "https://www.cax-if.org/joint_testing_info.html#recpracs"] [join ""] [join "Link to CAx-IF Recommended Practices"]
      
      [$worksheet($pmi_coverage) Range "A1"] Select
      catch {[$worksheet($pmi_coverage) PageSetup] PrintGridlines [expr 1]}
      $cells($pmi_coverage) Item 1 1 [file tail $localName]
      $cells($pmi_coverage) Item 35 1 "See Help > PMI Coverage Analysis"

# add images for the CAx-IF and NIST PMI models
      pmiAddModelPictures $pmi_coverage
    }
# errors
  } emsg]} {
    errorMsg "ERROR formatting PMI Presentation Coverage worksheet: $emsg"
  }
}
