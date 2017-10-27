proc gpmiAnnotation {entType} {
  global objDesign
  global ao aoEntTypes cells col entLevel ent entAttrList gpmiRow gtEntity nindex opt pmiCol pmiHeading pmiStartCol
  global recPracNames stepAP syntaxErr x3dShape x3dMsg useXL
  global geomType tessCoordID

  if {$opt(DEBUG1)} {outputMsg "START gpmiAnnotation $entType" red}

# basic geometry
  set cartesian_point [list cartesian_point coordinates]
  set direction       [list direction name direction_ratios]
  set a2p3d           [list axis2_placement_3d location $cartesian_point axis $direction ref_direction $direction]

  if {$opt(VIZPMI)} {
    set polyline        [list polyline name points $cartesian_point]
    set circle          [list circle name position $a2p3d radius]
    set trimmed_curve   [list trimmed_curve name basis_curve $circle]
  } else {
    set circle          [list circle]
    set polyline        [list polyline]
    set trimmed_curve   [list trimmed_curve basis_curve]
  }
  set composite_curve [list composite_curve segments [list composite_curve_segment parent_curve $trimmed_curve]]                        

# tessellated geometry
  set triangulated_face                [list triangulated_face name]
  set complex_triangulated_face        [list complex_triangulated_face name]
  set triangulated_surface_set         [list triangulated_surface_set name]
  set complex_triangulated_surface_set [list complex_triangulated_surface_set name]
  set tessellated_curve_set            [list tessellated_curve_set name]
  set tessellated_geometric_set        [list tessellated_geometric_set name children $tessellated_curve_set $complex_triangulated_surface_set $triangulated_surface_set]
  set repo_tessellated_geometric_set   [list repositioned_tessellated_item_and_tessellated_geometric_set name location $a2p3d children $tessellated_curve_set $complex_triangulated_surface_set $triangulated_surface_set]

# curve and fill style
  set colour      [list colour_rgb name red green blue]
  set curve_style [list presentation_style_assignment styles [list curve_style name curve_colour $colour [list draughting_pre_defined_colour name]]]
  set fill_style  [list presentation_style_assignment styles [list surface_style_usage style [list surface_side_style styles [list surface_style_fill_area fill_area [list fill_area_style name fill_styles [list fill_area_style_colour fill_colour $colour [list draughting_pre_defined_colour name]]]]]]]

  set geometric_curve_set  [list geometric_curve_set name elements $polyline $circle $trimmed_curve $composite_curve]
  set annotation_fill_area [list annotation_fill_area name boundaries $polyline $circle $trimmed_curve]

# annotation occurrences
  set PMIP(annotation_occurrence)             [list annotation_occurrence name styles $curve_style item $geometric_curve_set]
  set PMIP(annotation_curve_occurrence)       [list annotation_curve_occurrence name styles $curve_style item $geometric_curve_set]
  set PMIP(annotation_curve_occurrence_and_geometric_representation_item) [list annotation_curve_occurrence_and_geometric_representation_item name styles $curve_style item $geometric_curve_set]
  set PMIP(annotation_fill_area_occurrence)   [list annotation_fill_area_occurrence name styles $fill_style item $annotation_fill_area]
  set PMIP(draughting_annotation_occurrence)  [list draughting_annotation_occurrence name styles $curve_style item $geometric_curve_set]
  set PMIP(draughting_annotation_occurrence_and_geometric_representation_item) [list draughting_annotation_occurrence_and_geometric_representation_item name styles $curve_style item $geometric_curve_set]
  set PMIP(tessellated_annotation_occurrence) [list tessellated_annotation_occurrence name styles $curve_style item $tessellated_geometric_set $repo_tessellated_geometric_set]

# annotation placeholder
  set planar_box    [list planar_box size_in_x size_in_y placement $a2p3d]
  set geometric_set [list geometric_set name elements $cartesian_point $a2p3d $planar_box]
  set PMIP(annotation_placeholder_occurrence) [list annotation_placeholder_occurrence name styles $curve_style item $geometric_set]
    
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
  set x3dShape 0
  set gpmiRow($ao) {}
  set geomType ""
  set tessCoordID {}
  catch {unset gtEntity}
  if {![info exists x3dMsg]} {set x3dMsg {}}

  if {[info exist pmiHeading]} {unset pmiHeading}
  if {[info exists ent]} {unset ent}

  if {$opt(PMIGRF)} {outputMsg " Adding PMI Presentation Report" blue}
  if {$opt(VIZPMI)} {
    set msg " Adding PMI Presentation Visualization"
    if {$opt(XLSCSV) == "None"} {append msg " ([formatComplexEnt $entType])"}
    outputMsg $msg green
  }

# look for syntax errors with entity usage
  if {$stepAP == "AP242"} {
    set c1 [string first "_and_characterized_object" $ao]
    set c2 [string first "characterized_object_and_" $ao]
    if {$c1 != -1} {
      set msg "Syntax Error: Using 'characterized_object' with '[string range $ao 0 $c1-1]' is not valid for PMI Presentation.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 10.2, 10.3)"
      errorMsg $msg
      lappend syntaxErr($ao) [list 1 1 $msg]
    } elseif {$c2 != -1} {
      set msg "Syntax Error: Using 'characterized_object' with '[string range $ao 25 end]' is not valid for PMI Presentation.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 10.2, 10.3)"
      errorMsg $msg
      lappend syntaxErr($ao) [list 1 1 $msg]
    }
  
    if {[string first "annotation_occurrence" $ao] != -1 && [string first "tessellated" $ao] == -1 && [string first "draughting_annotation_occurrence" $ao] == -1} {
      set msg "Syntax Error: Using 'annotation_occurrence' with $stepAP is not valid for PMI Presentation.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 8.1.1)"
      errorMsg $msg
      lappend syntaxErr($ao) [list 1 1 $msg]
    }
  }

  if {[string first "AP203" $stepAP] == 0 || [string first "AP214" $stepAP] == 0} {
    if {[string first "annotation_curve_occurrence" $ao] != -1} {
      set msg "Syntax Error: Using 'annotation_curve_occurrence' with $stepAP is not valid for PMI Presentation.\n[string repeat " " 14]\($recPracNames(pmi203), Sec. 4.1.1)"
      errorMsg $msg
      lappend syntaxErr($ao) [list 1 1 $msg]
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
    lappend syntaxErr($ao) [list 1 1 $msg]
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
  if {$useXL} {set pmiStartCol($ao) [getNextUnusedColumn $startent 3]}

# process all annotation_occurrence entities, call gpmiAnnotationReport
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    if {[$objEntity Type] == $startent} {
      if {$n < 10000000} {
        if {[expr {$n%2000}] == 0} {
          if {$n > 0} {outputMsg "  $n"}
          update
        }
        gpmiAnnotationReport $objEntity
        if {$opt(DEBUG1)} {outputMsg \n}
      }
      incr n
    }
  }
  set col($ao) $pmiCol
  
# write any remaining geometry for polyline annotations
  if {$opt(VIZPMI)} {x3dPolylinePMI}
}

# -------------------------------------------------------------------------------
proc gpmiAnnotationReport {objEntity} {
  global objDesign
  global ao aoname assocGeom badAttributes boxSize cells circleCenter col currx3dPID curveTrim developer dirRatio dirType draftModelCameras draftModelCameraNames
  global entCount entLevel ent entAttrList entCount entsWithErrors geomType gpmiEnts gpmiID gpmiIDRow gpmiRow gpmiTypes gpmiTypesInvalid gpmiTypesPerFile gpmiValProp
  global iCompCurve iCompCurveSeg incrcol iPolyline localName nindex nistVersion nshape numCompCurve numCompCurveSeg numPolyline numx3dPID
  global objEntity1 opt pmiCol pmiColumns pmiHeading pmiStartCol pointLimit prefix propDefIDS recPracNames savedViewCol stepAP syntaxErr 
  global x3dColor x3dCoord x3dFile x3dFileName x3dStartFile x3dIndex x3dPoint x3dPID x3dShape x3dMsg x3dIndexType x3dMax x3dMin
  global tessCoord tessIndex tessIndexCoord tessRepo tessPlacement gpmiPlacement placeNCP placeOrigin placeAnchor useXL
  global savedViewName
  #outputMsg "gpmiAnnotationReport" red

# entLevel is very important, keeps track level of entity in hierarchy
  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]

  set maxcp $pointLimit

  if {[string first "handle" $objEntity] != -1} {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set objAttributes [$objEntity Attributes]
    set ent($entLevel) $objType
    if {$entLevel == 1} {set objEntity1 $objEntity}

    if {$opt(DEBUG1) && $geomType != "polyline"} {outputMsg "$ind ENT $entLevel #$objID=$objType (ATR=[$objAttributes Count])" blue}

# check if there are rows with ao for a report and not visualize
    if {$gpmiEnts($objType)} {
      set gpmiID $objID
      if {![info exists gpmiIDRow($ao,$gpmiID)] && $opt(PMIGRF) && !$opt(VIZPMI)} {
        incr entLevel -1
        return
      }

# write geometry polyline annotations
      if {$opt(VIZPMI)} {x3dPolylinePMI}
    }
    
# keep track of the number of c_c or c_c_s, if not polyline
    if {$objType == "composite_curve"} {
      incr iCompCurve
    } elseif {$objType == "composite_curve_segment"} {
      incr iCompCurveSeg
    }
    
    if {[string first "occurrence" $ao] != -1 && $objType != $ao}  {
      if {$entLevel == 2 && \
          $objType != "geometric_curve_set" && $objType != "annotation_fill_area" && $objType != "presentation_style_assignment" && \
          $objType != "geometric_set" && [string first "tessellated_geometric_set" $objType] == -1} {
        set msg "Syntax Error: '[formatComplexEnt $objType]' is not allowed as an 'item' attribute of: [formatComplexEnt $ao]\n[string repeat " " 14]"
        if {$stepAP == "AP242"} {
          append msg "($recPracNames(pmi242), Sec. 8.1.1, 8.1.2, 8.2)"
        } else {
          append msg "($recPracNames(pmi203), Sec. 4.1.1, 4.1.2)"
        }
        errorMsg $msg
        lappend syntaxErr($ao) [list $gpmiID item $msg]
      }
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

# get values for these entity and attribute pairs
              set ok 0
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
                    if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
                  }
                }
                "axis2_placement_3d axis" {
                  set dirType "axis"
                }
                "axis2_placement_3d ref_direction" {
                  set dirType "refdir"
                }
              }

              set colName "value"
  
              #if {$ok && [info exists gpmiID] && $opt(PMIGRF)} 
              if {$ok && [info exists gpmiIDRow($ao,$gpmiID)] && $opt(PMIGRF)} {
                set c [string index [cellRange 1 $col($ao)] 0]
                set r $gpmiIDRow($ao,$gpmiID)

# column name
                if {![info exists pmiHeading($col($ao))]} {
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
              if {$opt(DEBUG1) && $geomType != "polyline"} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}
          
# start of a list of cartesian points, assuming it is for a polyline, entLevel = 3
              if {$objAttrType == "ListOfcartesian_point" && $entLevel == 3} {
                #outputMsg 1entLevel$entLevel red
                if {$maxcp <= 10 && $maxcp < $objSize} {
                  append x3dPID "($maxcp of $objSize) cartesian_point "
                } else {
                  append x3dPID "($objSize) cartesian_point "
                }
                set numx3dPID $objSize
                set currx3dPID 0
                incr iPolyline
    
                set str ""
                for {set i 0} {$i < $objSize} {incr i} {append x3dIndex "[expr {$i+$nindex}] "}
                append x3dIndex "-1 "
                incr nindex $objSize
                set x3dIndexType "Line"
              }

# get values for these entity and attribute pairs
# g_c_s and a_f_a both start keeping track of their polylines
# cartesian_point is need to generated X3DOM
              set ok 0
              switch -glob $ent1 {
                "cartesian_point coordinates" {
                  set coord [vectrim $objValue]
                  #set coord "[trimNum [lindex $objValue 0]] [trimNum [lindex $objValue 1]] [trimNum [lindex $objValue 2]]"

# save origin for tessellated placement, convert Z = -Y, Y = Z
                  if {[info exists tessRepo]} {
                    if {$tessRepo} {lappend tessPlacement(origin) $coord}
                  }

# placeholder origin and anchor, convert as above?
                  if {[string first "placeholder" $ao] != -1} {
                    incr placeNCP
                    if {$placeNCP == 1} {
                      set gpmiPlacement(origin) $coord
                      set placeOrigin $coord
                    } elseif {$placeNCP == 3} {
                      set placeAnchor $coord
                    }
                  }

                  if {$opt(VIZPMI) && $x3dFileName != ""} {

# entLevel = 4 for polyline
                    if {$entLevel == 4 && $geomType == "polyline"} {
                      append x3dCoord "$coord "
                      setCoordMinMax $objValue

# circle center
                    } elseif {$geomType == "circle"} {
                      set circleCenter $objValue

# planar_box corner
                    } elseif {$geomType == "planar_box"} {
                      setCoordMinMax $objValue

# write placeholder box after getting origin and anchor
                      if {$placeNCP == 3} {
                        append x3dCoord "0 0 0 "
                        append x3dCoord "$boxSize(x) 0 0 " 
                        append x3dCoord "0 $boxSize(y) 0 "
                        append x3dCoord "$boxSize(x) $boxSize(y) 0 "
                        append x3dIndex "0 1 3 2 0 -1"
                        set x3dIndexType "Line"
                      }
                    }
                  }
                }
                "geometric_curve_set elements" {
                  set ok 1
                  if {$opt(PMIGRF)} {
                    set col($ao) [expr {$pmiStartCol($ao)+1}]
                    if {$stepAP == "AP242"} {
                      set colName "elements[format "%c" 10](Sec. 8.1.1)"
                    } else {
                      set colName "elements[format "%c" 10](Sec. 4.1.1)"
                    }
                  }
# keep track of polyline items
                  set numPolyline $objSize
                  set x3dPointID ""
                  set iPolyline 0
                  set x3dIndex ""
                  set x3dCoord ""
                  set nindex 0
# keep track of composite curve items
                  set numCompCurve $objSize
                  set iCompCurve 0
                }
                "geometric_set elements" {
                  set ok 1
                  if {$opt(PMIGRF)} {
                    set col($ao) [expr {$pmiStartCol($ao)+1}]
                    if {$stepAP == "AP242"} {
                      set colName "elements[format "%c" 10](Sec. 8.1.1)"
                    } else {
                      set colName "elements[format "%c" 10](Sec. 4.1.1)"
                    }
                  }
                }
                "annotation_fill_area boundaries" {
                  set ok 1
                  if {$opt(PMIGRF)} {
                    set col($ao) [expr {$pmiStartCol($ao)+1}]
                    if {$stepAP == "AP242"} {
                      set colName "boundaries[format "%c" 10](Sec. 8.1.2)"
                    } else {
                      set colName "boundaries[format "%c" 10](Sec. 4.1.2)"
                    }
                  }
# keep track of polyline items
                  set numPolyline $objSize
                  set x3dPointID ""
                  set iPolyline 0
                  set x3dIndex ""
                  set x3dCoord ""
                  set nindex 0
                }
                "*tessellated_geometric_set children" {
                  set ok 1
                  if {$opt(PMIGRF)} {
                    set col($ao) [expr {$pmiStartCol($ao)+1}]
                    set colName "children[format "%c" 10](Sec. 8.2)"
                  }
                }
                "direction direction_ratios" {
                  set dirRatio(x) [format "%.4f" [lindex $objValue 0]]
                  set dirRatio(y) [format "%.4f" [lindex $objValue 1]]
                  set dirRatio(z) [format "%.4f" [lindex $objValue 2]]
                  if {[info exists tessRepo]} {if {$tessRepo} {lappend tessPlacement($dirType) $objValue}}
                  if {[string first "placeholder" $ao] != -1} {set gpmiPlacement($dirType) $objValue}
                }
                "composite_curve segments" {set numCompCurveSeg $objSize}
              }

# value in spreadsheet
              if {$ok && $useXL && [info exists gpmiIDRow($ao,$gpmiID)] && $opt(PMIGRF)} {
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
                      set msg "Syntax Error: Invalid '[$val Type]' attribute for tessellated_geometric_set.children\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 8.2)"
                      errorMsg $msg
                      lappend syntaxErr(tessellated_geometric_set) [list "-$r" children $msg]
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
                      lappend syntaxErr(tessellated_geometric_set) [list "-$r" children $msg]
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
    
# get values for these entity and attribute pairs
              set ok 0
              set colName ""
              switch -glob $ent1 {
                "circle radius" {
                  if {$opt(VIZPMI) && $x3dFileName != ""} {
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
                    for {set i 0} {$i < $ns} {incr i} {append x3dIndex "[expr {$i+$nindex}] "}
                    if {!$trimmed} {
                      append x3dIndex "$nindex -1 "
                    } else {
                      append x3dIndex "-1 "
                    }
                    incr nindex $ns
                    set x3dIndexType "Line"

                    for {set i 0} {$i < $ns} {incr i} {
                      if {[expr {abs($dirRatio(z))}] > 0.99} {
                        set x3dPoint(x) [expr {$objValue*cos($angle)+[lindex $circleCenter 0]}]
                        set x3dPoint(y) [expr {-1.*$objValue*sin($angle)+[lindex $circleCenter 1]}]
                        set x3dPoint(z) [lindex $circleCenter 2]
                      } elseif {[expr {abs($dirRatio(y))}] > 0.99} {
                        set x3dPoint(x) [expr {$objValue*cos($angle)+[lindex $circleCenter 0]}]
                        set x3dPoint(z) [expr {-1.*$objValue*sin($angle)+[lindex $circleCenter 2]}]
                        set x3dPoint(y) [lindex $circleCenter 1]
                      } elseif {[expr {abs($dirRatio(x))}] > 0.99} {
                        set x3dPoint(z) [expr {$objValue*cos($angle)+[lindex $circleCenter 2]}]
                        set x3dPoint(y) [expr {-1.*$objValue*sin($angle)+[lindex $circleCenter 1]}]
                        set x3dPoint(x) [lindex $circleCenter 0]
                      } else {
                        errorMsg " PMI annotation circle orientation ($dirRatio(x), $dirRatio(y), $dirRatio(z)) is ignored."
                        set msg "Complex circle orientation is ignored."
                        if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
                        set x3dPoint(x) [expr {$objValue*cos($angle)+[lindex $circleCenter 0]}]
                        set x3dPoint(y) [expr {-1.*$objValue*sin($angle)+[lindex $circleCenter 1]}]
                        set x3dPoint(z) [lindex $circleCenter 2]
                      }

                      foreach idx {x y z} {
                        if {$x3dPoint($idx) > $x3dMax($idx)} {set x3dMax($idx) $x3dPoint($idx)}
                        if {$x3dPoint($idx) < $x3dMin($idx)} {set x3dMin($idx) $x3dPoint($idx)}
                      }
                      append x3dCoord "[format "%.4f" $x3dPoint(x)] [format "%.4f" $x3dPoint(y)] [format "%.4f" $x3dPoint(z)] "
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
                  errorMsg " Trimmed circles in PMI annotations might have the wrong orientation."
                  set msg "Trimmed circles might have the wrong orientation."
                  if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
                }
                "cartesian_point name" {
                  if {$entLevel == 4} {
                    set ok 1
                    if {$opt(PMIGRF)} {set col($ao) [expr {$pmiStartCol($ao)+2}]}
                  }
                }
                "geometric_set name" -
                "geometric_curve_set name" -
                "annotation_fill_area name" -
                "*tessellated_geometric_set name" {
                  set ok 1
                  if {$opt(PMIGRF)} {
                    set col($ao) $pmiStartCol($ao)
                    if {$stepAP == "AP242"} {
                      set colName "name[format "%c" 10](Sec. 8.4)"
                    } else {
                      set colName "name[format "%c" 10](Sec. 4.3)"
                    }
                  }
                  set tessRepo 0
                  if {[string first "repositioned" $ent1] != -1} {
                    set tessRepo 1
                    catch {unset tessPlacement}
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
                    if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
                  }
                  if {[string first "placeholder" $ent1] != -1} {
                    set placeNCP 0
                    set msg "Annotation placeholder leaders lines might not have the correct anchor points."
                    if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
                  }
                  if {[string first "tessellated" $ent1] != -1} {
                    set ok 1
                    foreach ann [list annotation_curve_occurrence_and_geometric_representation_item annotation_curve_occurrence] {
                      if {[info exist entCount($ann)] && $ok} {
                        set msg "Syntax Error: Using both '[formatComplexEnt $ann]' and 'tessellated_annotation_occurrence' is not recommended.\n[string repeat " " 14]($recPracNames(pmi242), Sec. 8.3, Important Note)"
                        errorMsg $msg
                        addCellComment $ann 1 1 $msg
                        lappend syntaxErr($ann) [list 1 1 $msg]
                        set ok 0
                      }
                    }
                  }
                } 
                "*triangulated_face name" -
                "*triangulated_surface_set name" -
                "tessellated_curve_set name" {
# write tessellated coords and index for pmi and part geometry
                  if {$opt(VIZPMI) && $ao == "tessellated_annotation_occurrence"} {
                    if {[info exists tessIndex($objID)] && [info exists tessCoord($tessIndexCoord($objID))]} {
                      x3dTessGeom $objID $objEntity1 $ent1
                    } else {
                      errorMsg "Missing tessellated coordinates and index for $objID"
                    }
                  }
                }
                "curve_style name" -
                "fill_area_style name" {
                  set ok 1
                  if {$opt(PMIGRF)} {
                    set col($ao) [expr {$pmiStartCol($ao)+2}]
                    if {$stepAP == "AP242"} {
                      set colName "presentation style[format "%c" 10](Sec. 8.5)"
                    } else {
                      set colName "presentation style[format "%c" 10](Sec. 4.4)"
                    }
                  }
                }
                "colour_rgb red" {
                  if {$entLevel == 4 || $entLevel == 8} {
                    if {$opt(gpmiColor) > 0} {
                      set x3dColor [x3dSetColor $opt(gpmiColor)]
                    } else {
                      set x3dColor $objValue
                    }
                  }
                }
                "colour_rgb green" {
                  if {$entLevel == 4 || $entLevel == 8} {
                    if {$opt(gpmiColor) == 0} {append x3dColor " $objValue"}
                  }
                }
                "colour_rgb blue" {
                  if {$entLevel == 4 || $entLevel == 8} {
                    if {$opt(gpmiColor) == 0} {append x3dColor " $objValue"}
                    if {$opt(PMIGRF)} {
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
                }
                "draughting_pre_defined_colour name" {
                  if {$entLevel == 4 || $entLevel == 8} {
                    switch $objValue {
                      black   {set x3dColor "0 0 0"}
                      white   {set x3dColor "1 1 1"}
                      red     {set x3dColor "1 0 0"}
                      yellow  {set x3dColor "1 1 0"}
                      green   {set x3dColor "0 1 0"}
                      cyan    {set x3dColor "0 1 1"}
                      blue    {set x3dColor "0 0 1"}
                      magenta {set x3dColor "1 0 1"}
                      default {
                        set x3dColor ".7 .7 .7"
                        errorMsg "Syntax Error: Unknown draughting_pre_defined_colour name '$objValue' (using gray)\n[string repeat " " 14]($recPracNames(model), Sec. 4.2.3, Table 2)"
                      }
                    }

                    if {$opt(gpmiColor) > 0} {set x3dColor [x3dSetColor $opt(gpmiColor)]}
                    if {$opt(PMIGRF)} {
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
                }
                "polyline name" {set geomType "polyline"}
                "circle name"   {set geomType "circle"}
                "composite_curve name" {set iCompCurveSeg 0}
                "planar_box size_in_*" {
                  set geomType "planar_box"
                  if {[string first "y" $ent1] != -1} {
                    set boxSize(y) [trimNum $objValue]
                  } else {
                    set boxSize(x) [trimNum $objValue]
                  }
                }
              }

# value in spreadsheet
              if {$ok} {
                if {[info exists gpmiIDRow($ao,$gpmiID)] && [string first "occurrence" $ao] != -1 && $opt(PMIGRF)} {
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
                }

# look for correct PMI name on 
# geometric_curve_set  annotation_fill_area  tessellated_geometric_set  composite_curve
                if {$ent1 == "geometric_curve_set name" || \
                    $ent1 == "geometric_set name" || \
                    $ent1 == "annotation_fill_area name" || \
                    $ent1 == "tessellated_geometric_set name" || \
                    $ent1 == "repositioned_tessellated_item_and_tessellated_geometric_set name" || \
                    $ent1 == "composite_curve name"} {
                  set ov $objValue

# look for invalid 'name' values                  
                  set invalid ""
                  if {[string first "occurrence" $ao] != -1} {
                    if {$ov == "" || [lsearch $gpmiTypes $ov] == -1} {
                      if {$ov == ""} {
                        set msg "Syntax Error: Missing 'name' attribute on [formatComplexEnt [lindex $ent1 0]]."
                        set ov "(blank)"
                      } else {
                        set msg "Syntax Error: 'name' attribute ($ov) on [formatComplexEnt [lindex $ent1 0]] is not recommended."
                      }
                      append msg "\n[string repeat " " 14]"
                      if {$stepAP == "AP242"} {
                        append msg "($recPracNames(pmi242), Sec. 8.4, Table 14)"
                      } else {
                        append msg "($recPracNames(pmi203), Sec. 4.3, Table 1)"
                      }
                      errorMsg $msg
                      if {[info exists gpmiTypesInvalid]} {if {[lsearch $gpmiTypesInvalid $ov] == -1} {lappend gpmiTypesInvalid $ov}}
                      set invalid $msg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                    }
                  }
                  
# count number of gpmi types
                  if {[info exists aoname]} {
                    if {$ov != $aoname} {
                      lappend gpmiTypesPerFile "$ov/$aoname"
                    } else {
                      set n 0
                      set objGuiEntities [$objEntity1 GetUsedIn [string trim draughting_model_item_association] [string trim identified_item]]
                      ::tcom::foreach objGuiEntity $objGuiEntities {
                        incr n
                        if {$n == 1} {lappend gpmiTypesPerFile "$ov/$aoname[$objEntity1 P21ID]"}
                      }
                      if {$n == 0} {
                        set objGuiEntities [$objEntity1 GetUsedIn [string trim draughting_callout] [string trim contents]]
                        ::tcom::foreach objGuiEntity $objGuiEntities {
                          incr n
                          if {$n == 1} {lappend gpmiTypesPerFile "$ov/$aoname[$objGuiEntity P21ID]"}
                        }
                      }
                    }
                  }
              
# start X3DOM file, read tessellated geometry
                  if {$opt(VIZPMI) && [string first "occurrence" $ao] != -1} {
                    if {$x3dStartFile} {x3dFileStart}

# moved (start shape node if not tessellated)
                    if {$ao == "annotation_fill_area_occurrence"} {errorMsg " PMI annotations with filled characters are not filled."}
                    if {[string first "tessellated" $ao] == -1} {set x3dShape 1}
                    update
                  }               

# value in spreadsheet  
                  if {[info exists gpmiIDRow($ao,$gpmiID)] && [string first "occurrence" $ao] != -1 && $opt(PMIGRF)} {
                    set val [[$cells($ao) Item $r $c] Value]
                    if {$invalid != ""} {lappend syntaxErr($ao) [list $r $col($ao) $invalid]}
  
                    if {$val == ""} {
                      $cells($ao) Item $r $c $ov
                    } else {
                      $cells($ao) Item $r $c "$val[format "%c" 10]$ov"
                    }

# keep track of max column
                    set pmiCol [expr {max($col($ao),$pmiCol)}]
                  }

# keep track of cartesian point ids (x3dPID)
                } elseif {[info exists currx3dPID] && $ent1 == "cartesian_point name"} {
                  if {$currx3dPID < $maxcp} {append x3dPID "$objID "}
                  incr currx3dPID

# cell value for presentation style or color
                } elseif {[info exists gpmiIDRow($ao,$gpmiID)] && $opt(PMIGRF)} {
                  if {$colName != "colour"} {
                    $cells($ao) Item $r $c "$ent($entLevel) $objID"
                  } else {
                    if {$ent($entLevel) == "colour_rgb"} {
                      $cells($ao) Item $r $c "$ent($entLevel) $objID  ($x3dColor)"
                    } else {
                      $cells($ao) Item $r $c "$ent($entLevel) $objID  ($objValue)"
                    }
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing PMI Presentation ($objNodeType $ent2): $emsg3"
            set entLevel 2
          }
        }
      }
    }
  }
  incr entLevel -1
  
# write a few more things at the end of processing an annotation_occurrence entity
  if {$entLevel == 0 && $opt(PMIGRF) && [info exists gpmiIDRow($ao,$gpmiID)]} {

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
          set objGuiEntities [$objGuiEntity GetUsedIn [string trim draughting_model_item_association_with_placeholder] [string trim identified_item]]
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
              set msg "Syntax Error: Missing 'definition' attribute on draughting_model_item_association\n[string repeat " " 14]"
              if {$stepAP == "AP242"} {
                append msg "($recPracNames(pmi242), Sec. 9.3.1, Fig. 76)"
              } else {
                append msg "($recPracNames(pmi203), Sec. 5.3.1, Fig. 12)"
              }
              errorMsg $msg
              lappend syntaxErr([$objGuiEntity Type]) [list [$objGuiEntity P21ID] "definition" $msg]
            }
          } elseif {[$attrDMIA Name] == "used_representation"} {
            set dmiaDef [$attrDMIA Value]
            if {[string first "handle" $dmiaDef] != -1} {
              set dmiaDefType [$dmiaDef Type]
              if {[string first "draughting_model" $dmiaDefType] == -1} {
                set msg "Syntax Error: Invalid 'used_representation' attribute ($dmiaDefType) on draughting_model_item_association.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 7.3)"
                errorMsg $msg
                lappend syntaxErr([$objGuiEntity Type]) [list [$objGuiEntity P21ID] "used_representation" $msg]
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
        set str [reportAssocGeom $ao]
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
          if {[lsearch $gpmiRow($ao) $r] == -1} {lappend gpmiRow($ao) $r}
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
          if {[lsearch $gpmiRow($ao) $r] == -1} {lappend gpmiRow($ao) $r}
        }
      }
    } emsg]} {
      errorMsg "ERROR reporting Associated Geometry and Representation: $emsg"
    }
  }
    
# report camera models associated with the annotation_occurrence through draughting_model
  if {$entLevel == 0 && (($opt(PMIGRF) && [info exists gpmiIDRow($ao,$gpmiID)]) || ($opt(VIZPMI) && !$opt(PMIGRF)))} {
    if {[catch {
      set savedViews ""
      set savedViewName {}
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
                append savedViews $str
                incr nsv
              }
              lappend savedViewName $draftModelCameraNames([$entDraughtingModel P21ID])
              #errorMsg "  Adding Saved Views" green

              if {$opt(PMIGRF)} {
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
                if {[string first "()" $savedViews] != -1} {
                  set msg "Syntax Error: For Saved Views, missing required 'name' attribute on camera_model_d3\n[string repeat " " 14]"
                  if {$stepAP == "AP242"} {
                    append msg "($recPracNames(pmi242), Sec. 9.4.2.1, Fig. 86)"
                  } else {
                    append msg "($recPracNames(pmi203), Sec. 5.4.2.1, Fig. 14)"
                  }
                  lappend syntaxErr($ao) [list $r $savedViewCol $msg]
                  errorMsg $msg
                }
                if {[lsearch $gpmiRow($ao) $r] == -1} {lappend gpmiRow($ao) $r}
              }
              
# check for a mapped_item in draughting_model.items, do not check style_item (see old code)
              set attrsDraughtingModel [$entDraughtingModel Attributes]
              ::tcom::foreach attrDraughtingModel $attrsDraughtingModel {
                if {[$attrDraughtingModel Name] == "name"} {
                  set nameDraughtingModel [$attrDraughtingModel Value]
                }
                if {[$attrDraughtingModel Name] == "items" && $nameDraughtingModel != ""} {
                  set okcm 0
                  set okmi 0
                  set oksi 0
                  ::tcom::foreach item [$attrDraughtingModel Value] {
                    if {[$item Type] == "mapped_item"} {set okmi 1}
                    if {[string first "camera_model_d3" [$item Type]] == 0} {set okcm 1}
                  }
                  if {$okcm} {
                    if {$okmi == 0} {
                      set msg "Syntax Error: For Saved Views, missing required reference to 'mapped_item' on [formatComplexEnt [$entDraughtingModel Type]].items\n[string repeat " " 14]"
                      if {$stepAP == "AP242"} {
                        append msg "($recPracNames(pmi242), Sec. 9.4.2.1, Fig. 86)"
                      } else {
                        append msg "($recPracNames(pmi203), Sec. 5.4.2, Fig. 14)"
                      }
                      errorMsg $msg
                      lappend syntaxErr([$entDraughtingModel Type]) [list [$entDraughtingModel P21ID] items $msg]
                    }
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
                    lappend syntaxErr($relType) [list $mdadrID rep_2 $msg]
                  }
                  if {$relType == "representation_relationship"} {
                    set msg "For Saved Views, recommend using 'mechanical_design_and_draughting_relationship' instead of 'representation_relationship'\n  to relate draughting models "
                    if {$stepAP == "AP242"} {
                      append msg "($recPracNames(pmi242), Sec. 9.4.4 Note 2)"
                    } else {
                      append msg "($recPracNames(pmi203), Sec. 5.4.4 Note 2)"
                    }
                    errorMsg $msg
                    lappend syntaxErr(representation_relationship) [list 1 1 $msg]
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
        if {[string first "AP214" $stepAP] == -1} {lappend relTypes mechanical_design_and_draughting_relationship}
        set relType ""
        foreach item $relTypes {if {[info exists entCount($item)]} {if {$entCount($item) > 0} {set relType $item}}}
        if {$relType == ""} {
          set str "mechanical_design_and_draughting_relationship"
          if {[string first "AP214" $stepAP] == 0} {set str "representation_relationship"}
          set msg "Syntax Error: For Saved Views, missing '$str' to relate 'draughting_model'\n[string repeat " " 14]"
          if {$stepAP == "AP242"} {
            append msg "($recPracNames(pmi242), Sec. 9.4.4 Note 1, Fig. 93)"
          } else {
            append msg "($recPracNames(pmi203), Sec. 5.4.4 Note 1, Fig. 20)"
          }
          errorMsg $msg
        }
      }
    } emsg]} {
      errorMsg "ERROR adding Saved Views: $emsg"
    }
  }
  
# check if there are PMI validation properties (propDefIDS) associated with the annotation_occurrence
  if {$entLevel == 0 && $opt(PMIGRF) && [info exists gpmiIDRow($ao,$gpmiID)]} {
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
        if {$stepAP == "AP242"} {
          set colName "Validation Properties[format "%c" 10](Sec. 10.3)"
        } else {
          set colName "Validation Properties[format "%c" 10](Sec. 6.3)"
        }
        set c [string index [cellRange 1 $pmiColumns(vp)] 0]
        set r $gpmiIDRow($ao,$gpmiID)
        if {![info exists pmiHeading($pmiColumns(vp))]} {
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
# get camera models and validation properties
proc pmiGetCamerasAndProperties {} {
  global objDesign
  global draftModelCameras draftModelCameraNames gpmiValProp syntaxErr propDefIDS stepAP recPracNames entCount
  global opt savedViewNames savedViewFile savedViewFileName mytemp savedViewName savedViewpoint

  #outputMsg getCameras blue
  catch {unset draftModelCameras}
  catch {unset draftModelCameraNames}
  if {[catch {

# camera list
    set cmlist {}
    foreach cms [list camera_model_d3 camera_model_d3_multi_clipping] {
      if {[info exists entCount($cms)]} {if {$entCount($cms) > 0} {lappend cmlist $cms}}
    }

# draughting model list
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
            set attrDraughtingModels [$entDraughtingModel Attributes]

# DM name attribute
            set ok 0
            if {[llength $dmlist] == 1 || [string first "characterized" $dm] != -1} {set ok 1}
            if {$ok} {
              set nattr 0
              set iattr 1
              if {[string first "object" $dm] != -1} {set iattr 3}
              ::tcom::foreach attrDraughtingModel $attrDraughtingModels {
                incr nattr
                set nameDraughtingModel [$attrDraughtingModel Name]
                if {$nameDraughtingModel == "name" && $nattr == $iattr} {
                  set name [$attrDraughtingModel Value]
                  if {$name == ""} {
                    set msg "Syntax Error: For Saved Views, missing required 'name' attribute on [formatComplexEnt $dm]\n[string repeat " " 14]"
                    if {$stepAP == "AP242"} {
                      append msg "($recPracNames(pmi242), Sec. 9.4.2)"
                    } else {
                      append msg "($recPracNames(pmi203), Sec. 5.4.2)"
                    }
                    errorMsg $msg
                    lappend syntaxErr($dm) [list [$entDraughtingModel P21ID] name $msg]
                  }
                }
              }
            }

# CM name attribute
            ::tcom::foreach attrCameraModel $attrCameraModels {
              set nameCameraModel [$attrCameraModel Name]
              if {$nameCameraModel == "name"} {
                set name [$attrCameraModel Value]

# clean up the camera name
                regsub -all " " [string trim $name] "_" name1  
                regsub -all {\(} [string trim $name1] "_" name1 
                regsub -all {\)} [string trim $name1] "" name1  
                regsub -all {:~$%&*<>?/+\|\"\#\\\{\}} [string trim $name1] "_" name1
                
                if {$name == ""} {
                  set msg "Syntax Error: For Saved Views, missing required 'name' attribute on $cm\n[string repeat " " 14]"
                  if {$stepAP == "AP242"} {
                    append msg "($recPracNames(pmi242), Sec. 9.4.2.1, Fig. 86)"
                  } else {
                    append msg "($recPracNames(pmi203), Sec. 5.4.2.1, Fig. 14)"
                  }
                  errorMsg $msg
                  lappend syntaxErr($cm) [list [$entCameraModel P21ID] name $msg]
                }

# get axis2_placement_3d for camera viewpoint
              } elseif {$nameCameraModel == "view_reference_system"} {
                catch {unset savedViewpoint($name1)}
                if {[catch {
                  set a2p3d [[$attrCameraModel Value] Attributes]
                  set origin [[[[[$a2p3d Item 2] Value] Attributes] Item 2] Value]
                  set axis   [[[[[$a2p3d Item 3] Value] Attributes] Item 2] Value]
                  ::tcom::foreach attr $a2p3d {
                    if {[$attr Name] == "ref_direction"} {
                      set refdir [[[[$attr Value] Attributes] Item 2] Value]
                    }
                  }
                  lappend savedViewpoint($name1) [vectrim $origin]
                  lappend savedViewpoint($name1) [x3dRotation $axis $refdir]
                } emsg]} {
                  errorMsg "ERROR getting Saved View position and orientation: $emsg"
                  catch {raise .}
                }
              }
            }

# cameras associated with draughting models
            set str "[$entCameraModel P21ID] ($name)  "
            set id [$entDraughtingModel P21ID]
            if {![info exists draftModelCameras($id)]} {
              set draftModelCameras($id) $str
            } elseif {[string first $str $draftModelCameras($id)] == -1} {
              append draftModelCameras($id) "[$entCameraModel P21ID] ($name)  "
            }
            if {![info exists draftModelCameraNames($id)]} {
              set draftModelCameraNames($id) $name1
            } elseif {[string first $name $draftModelCameraNames($id)] == -1 && [string first $name1 $draftModelCameraNames($id)] == -1} {
              append draftModelCameraNames($id) " $name1"
            }

# keep track of saved views for graphic PMI
            if {$opt(VIZPMI)} {
              set dmcn $draftModelCameraNames([$entDraughtingModel P21ID])
              if {[lsearch $savedViewName $dmcn] == -1} {lappend savedViewName $dmcn}
              if {[lsearch $savedViewNames $name1] == -1} {
                lappend savedViewNames $name1
                set savedViewFileName($name1) [file join $mytemp $name1.txt]
                catch {file delete -force $savedViewFileName($name1)}
                set savedViewFile($name1) [open $savedViewFileName($name1) w]
                #outputMsg "camera name $name1" green
              }
            }
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
