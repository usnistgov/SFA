proc gpmiAnnotation {entType} {
  global objDesign
  global ao aoEntTypes col ent entAttrList entLevel gen geomType gpmiRow gtEntity nindex opt pmiCol
  global pmiHeading pmiStartCol recPracNames spaces stepAP syntaxErr tessCoordID useXL x3dShape

  if {$opt(DEBUG1)} {outputMsg "START gpmiAnnotation $entType" red}

# basic geometry
  set cartesian_point [list cartesian_point coordinates]
  set direction       [list direction name direction_ratios]
  set a2p3d           [list axis2_placement_3d location $cartesian_point axis $direction ref_direction $direction]

  if {$gen(View) && $opt(viewPMI)} {
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
  set fill_style  [list presentation_style_assignment styles [list surface_style_usage style \
                                                               [list surface_side_style styles \
                                                                 [list surface_style_fill_area fill_area \
                                                                   [list fill_area_style name fill_styles \
                                                                     [list fill_area_style_colour fill_colour $colour [list draughting_pre_defined_colour name]]]]]]]

  set geometric_curve_set  [list geometric_curve_set name elements $polyline $circle $trimmed_curve $composite_curve]
  set annotation_fill_area [list annotation_fill_area name boundaries $polyline $circle $trimmed_curve]

# annotation occurrences
  set PMIP(annotation_occurrence)             [list annotation_occurrence name styles $curve_style item $geometric_curve_set]
  set PMIP(annotation_curve_occurrence)       [list annotation_curve_occurrence name styles $curve_style item $geometric_curve_set]
  set PMIP(annotation_curve_occurrence_and_geometric_representation_item) [list annotation_curve_occurrence_and_geometric_representation_item name styles $curve_style item $geometric_curve_set]
  set PMIP(annotation_fill_area_occurrence)   [list annotation_fill_area_occurrence name styles $fill_style item $annotation_fill_area]
  set PMIP(tessellated_annotation_occurrence) [list tessellated_annotation_occurrence name styles $curve_style $fill_style item $tessellated_geometric_set $repo_tessellated_geometric_set]

# annotation placeholder
  set planar_box      [list planar_box size_in_x size_in_y placement $a2p3d]
  set geometric_set   [list geometric_set name elements $cartesian_point $a2p3d $planar_box]
  set PMIP(annotation_placeholder_occurrence) [list annotation_placeholder_occurrence name styles $curve_style item $geometric_set line_spacing]
  set apowll "annotation_placeholder_occurrence_with_leader_line"
  set PMIP($apowll) [lreplace $PMIP(annotation_placeholder_occurrence) 0 0 $apowll]

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

  if {[info exists pmiHeading]} {unset pmiHeading}
  if {[info exists ent]} {unset ent}

  if {$opt(PMIGRF) && $opt(xlFormat) != "None"} {outputMsg " Adding PMI Presentation Analyzer report" blue}
  if {$gen(View) && $opt(viewPMI)} {
    set msg " Adding Graphical PMI for the Viewer"
    if {$opt(xlFormat) == "None"} {append msg " ([formatComplexEnt $entType])"}
    outputMsg $msg green
  }

# look for syntax errors with entity usage
  if {[string first "AP242" $stepAP] == 0} {
    set c1 [string first "_and_characterized_object" $ao]
    set c2 [string first "characterized_object_and_" $ao]
    if {$c1 != -1} {
      set msg "Syntax Error: Using 'characterized_object' with '[string range $ao 0 $c1-1]' is not valid for PMI Presentation.$spaces\($recPracNames(pmi242), Sec. 10.2, 10.3)"
      errorMsg $msg
      lappend syntaxErr($ao) [list 1 1 $msg]
    } elseif {$c2 != -1} {
      set msg "Syntax Error: Using 'characterized_object' with '[string range $ao 25 end]' is not valid for PMI Presentation.$spaces\($recPracNames(pmi242), Sec. 10.2, 10.3)"
      errorMsg $msg
      lappend syntaxErr($ao) [list 1 1 $msg]
    }

    if {[string first "annotation_occurrence" $ao] != -1 && [string first "tessellated" $ao] == -1 && [string first "draughting_annotation_occurrence" $ao] == -1} {
      set msg "Syntax Error: Using 'annotation_occurrence' with $stepAP is not valid for PMI Presentation.$spaces\($recPracNames(pmi242), Sec. 8.1.1)"
      errorMsg $msg
      lappend syntaxErr($ao) [list 1 1 $msg]
    }
  }

  if {[string first "AP203" $stepAP] == 0 || [string first "AP214" $stepAP] == 0} {
    if {[string first "annotation_curve_occurrence" $ao] != -1} {
      set msg "Syntax Error: Using 'annotation_curve_occurrence' with $stepAP is not valid for PMI Presentation.$spaces\($recPracNames(pmi203), Sec. 4.1.1)"
      errorMsg $msg
      lappend syntaxErr($ao) [list 1 1 $msg]
    }
  }

  if {[string first "draughting" $ao] != -1} {
    set msg "Syntax Error: Using 'draughting_annotation_*_occurrence' is not valid for PMI Presentation.$spaces"
    if {[string first "AP242" $stepAP] == 0} {
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
  if {$useXL} {set pmiStartCol($ao) [getNextUnusedColumn $startent]}

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
  if {$gen(View) && $opt(viewPMI)} {x3dPolylinePMI}
}

# -------------------------------------------------------------------------------
proc gpmiAnnotationReport {objEntity} {
  global ao aoname assocGeom badAttributes boxSize cells circleCenter col currx3dPID curveTrim dirRatio dirType
  global draughtingModels draftModelCameraNames draftModelCameras ent entAttrList entCount entLevel gen geomType gpmiEnts gpmiID gpmiIDRow
  global gpmiName gpmiPlacement gpmiRow gpmiTypes gpmiTypesInvalid gpmiTypesPerFile gpmiValProp grayBackground iCompCurve iCompCurveSeg iPolyline
  global nindex numCompCurve numCompCurveSeg numPolyline numx3dPID objEntity1 opt placeAnchor placeNCP placeOrigin
  global pmiCol pmiColumns pmiHeading pmiStartCol propDefIDs recPracNames savedViewCol savedViewName spaces stepAP syntaxErr
  global tessCoord tessIndex tessIndexCoord tessPlacement tessPlacementID tessRepo useXL
  global x3dColor x3dCoord x3dFile x3dFileName x3dIndex x3dIndexType x3dMax x3dMin x3dPID x3dPoint x3dShape x3dStartFile

# entLevel is very important, keeps track level of entity in hierarchy
  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]

  set maxcp 2

  if {[string first "handle" $objEntity] != -1} {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set objAttributes [$objEntity Attributes]
    set ent($entLevel) $objType
    if {$entLevel == 1} {set objEntity1 $objEntity}

    if {$opt(DEBUG1) && $geomType != "polyline"} {outputMsg "$ind ENT $entLevel #$objID=$objType (ATR=[$objAttributes Count])" blue}

# check if there are rows with ao for a report and not view
    if {$gpmiEnts($objType)} {
      set gpmiID $objID
      if {![info exists gpmiIDRow($ao,$gpmiID)] && $opt(PMIGRF) && $opt(xlFormat) != "None" && !$opt(viewPMI)} {
        incr entLevel -1
        return
      }

# write geometry polyline annotations
      if {$gen(View) && $opt(viewPMI) && [string first "tessellated" $objType] == -1} {x3dPolylinePMI $objEntity1}
    }

# keep track of the number of c_c or c_c_s, if not polyline
    if {$objType == "composite_curve"} {
      incr iCompCurve
    } elseif {$objType == "composite_curve_segment"} {
      incr iCompCurveSeg
    }

    if {[string first "occurrence" $ao] != -1 && $objType != $ao && $opt(xlFormat) != "None"}  {
      if {$entLevel == 2 && \
          $objType != "geometric_curve_set" && $objType != "annotation_fill_area" && $objType != "presentation_style_assignment" && \
          $objType != "geometric_set" && [string first "tessellated_geometric_set" $objType] == -1} {
        set msg "Syntax Error: '[formatComplexEnt $objType]' is not allowed as an 'item' attribute of: [formatComplexEnt $ao]$spaces"
        if {[string first "AP242" $stepAP] == 0} {
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
                    set msg "Syntax Error: '[$objValue Type]' is not allowed as a 'basis_curve' for trimmed_curve$spaces"
                    if {[string first "AP242" $stepAP] == 0} {
                      append msg "($recPracNames(pmi242), Sec. 8.1.1, 8.1.2)"
                    } else {
                      append msg "($recPracNames(pmi203), Sec. 4.1.1, 4.1.2)"
                    }
                    errorMsg $msg
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

              if {$ok && [info exists gpmiIDRow($ao,$gpmiID)] && $opt(PMIGRF) && $opt(xlFormat) != "None"} {
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
            set msg "Error processing Graphical PMI ($objNodeType $ent2): $emsg3"
            errorMsg $msg
            lappend syntaxErr([lindex $ent1 0]) [list $objID [lindex $ent1 1] $msg]
          }

# --------------
# nodeType = 20
        } elseif {$objNodeType == 20} {
          if {[catch {
            if {$idx != -1} {
              if {$opt(DEBUG1) && $geomType != "polyline"} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}

# start of a list of cartesian points, assuming it is for a polyline, entLevel = 3
              if {$objAttrType == "ListOfcartesian_point" && $entLevel == 3} {
                if {$maxcp < $objSize} {
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
# save cartesian_point to x3d coordinates
              set ok 0
              switch -glob $ent1 {
                "composite_curve segments" {set numCompCurveSeg $objSize}

                "cartesian_point coordinates" {
                  set coord [vectrim $objValue]

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
                      catch {unset placeAnchor}
                    } elseif {$placeNCP == 3} {
                      set placeAnchor $coord
                    }
                  }

                  if {$gen(View) && $opt(viewPMI) && $x3dFileName != ""} {

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
                  if {$opt(PMIGRF) && $opt(xlFormat) != "None"} {
                    set col($ao) [expr {$pmiStartCol($ao)+1}]
                    if {[string first "AP242" $stepAP] == 0} {
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
                  if {$opt(PMIGRF) && $opt(xlFormat) != "None"} {
                    set col($ao) [expr {$pmiStartCol($ao)+1}]
                    set colName "elements[format "%c" 10](Sec. "
                    if {[string first "placeholder" $ao] != -1} {
                      append colName "7.2.2"
                    } elseif {[string first "AP242" $stepAP] == 0} {
                      append colName "8.1.1"
                    } else {
                      append colName "4.1.1"
                    }
                    append colName ")"
                  }
# for placeholder, check for required a2p3d, cartesian_point, and optional planar_box
                  if {[string first "placeholder" $ao] != -1} {
                    set elements {}
                    set msg "Syntax Error: "
                    foreach elem $objValue {
                      set etype [$elem Type]
                      if {[string first "handle" $etype] != -1} {set etype [[$etype Value] Type]}
                      lappend elements $etype
                      if {$etype != "axis2_placement_3d" && $etype != "cartesian_point" && $etype != "planar_box" && \
                          [string first $etype $msg] == -1} {append msg "'$etype' is not supported"}
                    }
                    if {[string first "not supported" $msg] == -1} {
                      if {[lsearch $elements "axis2_placement_3d"] == -1} {
                        append msg "Missing required 'axis2_placement_3d'"
                      } elseif {[lsearch $elements "cartesian_point"] == -1} {
                        append msg "Missing required 'cartesian_point'"
                      }
                    }
                    if {[string length $msg] < 15 && [lsearch $elements "planar_box"] == -1} {set msg "$ao 'planar_box' is not supported."}
                    if {[string length $msg] > 14} {
                      if {[string first "Syntax" $msg] == 0} {
                        append msg " in 'geometric_set.elements'.$spaces\($recPracNames(pmi242), Sec. 7.2.2)"
                      }
                      errorMsg $msg
                      lappend syntaxErr($ao) [list $gpmiID "elements" $msg]
                    }
                  }
                }
                "annotation_fill_area boundaries" {
                  set ok 1
                  if {$opt(PMIGRF) && $opt(xlFormat) != "None"} {
                    set col($ao) [expr {$pmiStartCol($ao)+1}]
                    if {[string first "AP242" $stepAP] == 0} {
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
                  if {$opt(PMIGRF) && $opt(xlFormat) != "None"} {
                    set col($ao) [expr {$pmiStartCol($ao)+1}]
                    set colName "children[format "%c" 10](Sec. 8.2)"
                  }
                }
                "direction direction_ratios" {
                  set dir [vecnorm $objValue]
                  set dirRatio(x,$dirType) [format "%.3f" [lindex $dir 0]]
                  set dirRatio(y,$dirType) [format "%.3f" [lindex $dir 1]]
                  set dirRatio(z,$dirType) [format "%.3f" [lindex $dir 2]]
                  if {[info exists tessRepo]} {
                    if {$tessRepo} {
                      lappend tessPlacement($dirType) $dir

# check for bad directions
                      set msg ""
                      if {[veclen $dir] == 0} {
                        set msg "Syntax Error: The axis2_placement_3d axis or ref_direction vector is '0 0 0' for a repositioned_tessellated_item."
                      } elseif {$dirType == "refdir" && [veclen [veccross [join $tessPlacement(axis)] [join $tessPlacement(refdir)]]] == 0} {
                        set msg "Syntax Error: The axis2_placement_3d axis and ref_direction vectors '[join $tessPlacement(refdir)]' are parallel for a repositioned_tessellated_item."
                      }
                      if {$msg != ""} {
                        errorMsg $msg
                        lappend syntaxErr(direction) [list $objID direction_ratios $msg]
                        lappend syntaxErr(repositioned_tessellated_item_and_tessellated_geometric_set) [list $tessPlacementID location $msg]
                      }
                    }
                  }
                  if {[string first "placeholder" $ao] != -1} {set gpmiPlacement($dirType) $dir}
                }
              }

# value in spreadsheet
              if {$ok && $useXL && [info exists gpmiIDRow($ao,$gpmiID)] && $opt(PMIGRF) && $opt(xlFormat) != "None"} {
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
                  set nval 0
                  ::tcom::foreach val [$objAttribute Value] {
                    append cellval([$val Type]) "[$val P21ID] "
                    incr nval
                    if {[string first "tessellated_geometric_set" $ent1] != -1 && [$val Type] != "tessellated_curve_set" && [$val Type] != "complex_triangulated_surface_set"} {
                      set msg "Syntax Error: Bad '[$val Type]' attribute for tessellated_geometric_set.children$spaces\($recPracNames(pmi242), Sec. 8.2)"
                      errorMsg $msg
                      lappend syntaxErr($objType) [list $objID children $msg]
                    }
                  }
                  if {$nval == 0 && [string first "tessellated_geometric_set" $ent1] != -1} {
                    set msg "Syntax Error: Missing 'children' attribute on [formatComplexEnt $objType].$spaces\($recPracNames(pmi242), Sec. 8.2)"
                    errorMsg $msg
                    lappend syntaxErr($objType) [list $objID children $msg]
                    lappend syntaxErr(tessellated_annotation_occurrence) [list $gpmiID children $msg]
                  }
                } emsg]} {
                  foreach val [$objAttribute Value] {
                    append cellval([$val Type]) "[$val P21ID] "
                    if {$ent1 == "geometric_curve_set items" && [$val Type] != "polyline" && [$val Type] != "trimmed_curve" && [$val Type] != "circle"} {
                      set msg "Syntax Error: Bad '[$val Type]' attribute for geometric_curve_set 'items'$spaces"
                      if {[string first "AP242" $stepAP] == 0} {
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
                    errorMsg "  Too much data to show in a cell: $emsg" red
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
            set msg "Error processing Graphical PMI ($objNodeType $ent2): $emsg3"
            errorMsg $msg
            lappend syntaxErr([lindex $ent1 0]) [list $objID [lindex $ent1 1] $msg]
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
                  if {$gen(View) && $opt(viewPMI) && $x3dFileName != ""} {
# write circle to x3d points and index
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
                      if {[expr {abs($dirRatio(z,axis))}] > 0.99} {
                        set x3dPoint(x) [expr {$objValue*cos($angle)+[lindex $circleCenter 0]}]
                        set x3dPoint(y) [expr {$objValue*sin($angle)+[lindex $circleCenter 1]}]
                        set x3dPoint(z) [lindex $circleCenter 2]
                      } elseif {[expr {abs($dirRatio(y,axis))}] > 0.99} {
                        set x3dPoint(x) [expr {$objValue*cos($angle)+[lindex $circleCenter 0]}]
                        set x3dPoint(z) [expr {$objValue*sin($angle)+[lindex $circleCenter 2]}]
                        set x3dPoint(y) [lindex $circleCenter 1]
                      } elseif {[expr {abs($dirRatio(x,axis))}] > 0.99} {
                        set x3dPoint(z) [expr {$objValue*cos($angle)+[lindex $circleCenter 2]}]
                        set x3dPoint(y) [expr {$objValue*sin($angle)+[lindex $circleCenter 1]}]
                        set x3dPoint(x) [lindex $circleCenter 0]
                      } else {
                        errorMsg " Circles in PMI annotations might be shown with the wrong orientation."
                        set x3dPoint(x) [expr {$objValue*cos($angle)+[lindex $circleCenter 0]}]
                        set x3dPoint(y) [expr {$objValue*sin($angle)+[lindex $circleCenter 1]}]
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
                  errorMsg " Trimmed circles in PMI annotations might be incorrectly trimmed."
                }
                "cartesian_point name" {
                  if {$entLevel == 4} {
                    set ok 1
                    if {$opt(PMIGRF) && $opt(xlFormat) != "None"} {set col($ao) [expr {$pmiStartCol($ao)+2}]}
                  }
                }
                "geometric_set name" -
                "geometric_curve_set name" -
                "annotation_fill_area name" -
                "*tessellated_geometric_set name" {
# do not delete this comment
                  set ok 1
                  if {$opt(PMIGRF) && $opt(xlFormat) != "None"} {
                    set col($ao) $pmiStartCol($ao)
                    if {[string first "AP242" $stepAP] == 0} {
                      set colName "name[format "%c" 10](Sec. 8.4)"
                    } else {
                      set colName "name[format "%c" 10](Sec. 4.3)"
                    }
                  }
                  set tessRepo 0
                  if {[string first "repositioned" $ent1] != -1} {
                    set tessRepo 1
                    set tessPlacementID $objID
                    catch {unset tessPlacement}
                  }
                }
                "annotation_curve_occurrence* name" -
                "annotation_fill_area_occurrence* name" -
                "annotation_placeholder_occurrence* name" -
                "annotation_occurrence* name" -
                "*tessellated_annotation_occurrence* name" {
                  set aoname $objValue
                  if {[string first "fill" $ent1] != -1 && $gen(View) && $opt(viewPMI)} {errorMsg " Annotations with filled characters are not filled."}
                  if {[string first "placeholder" $ent1] != -1} {set placeNCP 0}
                  if {[string first "tessellated" $ent1] != -1 && $opt(xlFormat) != "None"} {
                    set ok 1
                    foreach ann [list annotation_curve_occurrence_and_geometric_representation_item annotation_curve_occurrence] {
                      if {[info exists entCount($ann)] && $ok} {
                        set msg "Syntax Error: Using both '[formatComplexEnt $ann]' and 'tessellated_annotation_occurrence' is not recommended.$spaces\($recPracNames(pmi242), Sec. 8.3, Important Note)"
                        errorMsg $msg
                        addCellComment $ann 1 1 $msg
                        lappend syntaxErr($ann) [list 1 1 $msg]
                        set ok 0
                      }
                    }

# check new rule with AP242 edition 2
                    if {$stepAP == "AP242e2" && \
                        ![info exists entCount(draughting_model_and_tessellated_shape_representation)] && \
                        ![info exists entCount(characterized_representation_and_draughting_model_and_tessellated_shape_representation)]} {
                      errorMsg "Syntax Error: Missing (draughting_model)(tessellated_shape_representation) entity$spaces\($recPracNames(pmi242), Sec. 8.2, note for AP242 Edition 2)"
                    }
                  }
                }
                "annotation_placeholder_occurrence* line_spacing" {
                  if {$objValue <= 1.E-6} {
                    set msg "Syntax Error: [lindex $ent1 0] 'line_spacing' attribute must be greater that zero.$spaces\($recPracNames(pmi242), Sec. 7.2.2)"
                    errorMsg $msg
                    lappend syntaxErr([lindex $ent1 0]) [list $objID "line_spacing" $msg]
                  }
                }
                "*triangulated_face name" -
                "*triangulated_surface_set name" -
                "tessellated_curve_set name" {
# write tessellated coords and index for pmi and part geometry
                  if {$gen(View) && $opt(viewPMI) && $ao == "tessellated_annotation_occurrence"} {
                    if {[info exists tessIndex($objID)] && [info exists tessCoord($tessIndexCoord($objID))]} {x3dTessGeom $objID $objEntity1 $ent1}
                  }
                }
                "curve_style name" -
                "fill_area_style name" {
                  set ok 1
                  if {$opt(PMIGRF) && $opt(xlFormat) != "None"} {
                    set col($ao) [expr {$pmiStartCol($ao)+2}]
                    if {[string first "AP242" $stepAP] == 0} {
                      set colName "presentation style[format "%c" 10](Sec. 8.5)"
                    } else {
                      set colName "presentation style[format "%c" 10](Sec. 4.4)"
                    }
                  }
                }
                "colour_rgb red" {
                  if {$entLevel == 4 || $entLevel == 8} {
                    set objValue [trimNum $objValue]
                    set colorRGB $objValue
                    if {$opt(gpmiColor) > 0} {
                      set x3dColor [x3dSetPMIColor $opt(gpmiColor)]
                    } else {
                      set x3dColor $objValue
                    }
                  }
                }
                "colour_rgb green" {
                  if {$entLevel == 4 || $entLevel == 8} {
                    set objValue [trimNum $objValue]
                    append colorRGB " $objValue"
                    if {$opt(gpmiColor) == 0} {append x3dColor " $objValue"}
                  }
                }
                "colour_rgb blue" {
                  if {$entLevel == 4 || $entLevel == 8} {
                    set objValue [trimNum $objValue]
                    append colorRGB " $objValue"
                    if {$opt(gpmiColor) == 0} {
                      append x3dColor " $objValue"
                      if {[expr {([lindex $x3dColor 0]+[lindex $x3dColor 1]+[lindex $x3dColor 2])/3.}] > 0.93} {set grayBackground 1}
                    }
                    if {$opt(PMIGRF) && $opt(xlFormat) != "None"} {
                      set ok 1
                      set col($ao) [expr {$pmiStartCol($ao)+3}]
                      if {[string first "AP242" $stepAP] == 0} {
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
                    set x3dColor [x3dPreDefinedColor $objValue]
                    if {$opt(gpmiColor) > 0} {
                      set x3dColor [x3dSetPMIColor $opt(gpmiColor)]
                    } elseif {$objValue == "white"} {
                      set grayBackground 1
                    }
                    if {$opt(PMIGRF) && $opt(xlFormat) != "None"} {
                      set ok 1
                      set col($ao) [expr {$pmiStartCol($ao)+3}]
                      if {[string first "AP242" $stepAP] == 0} {
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
                if {[info exists gpmiIDRow($ao,$gpmiID)] && [string first "occurrence" $ao] != -1 && $opt(PMIGRF) && $opt(xlFormat) != "None"} {
                  set c [string index [cellRange 1 $col($ao)] 0]
                  set r $gpmiIDRow($ao,$gpmiID)

# column name
                  if {$colName != ""} {
                    if {![info exists pmiHeading($col($ao))]} {
                      $cells($ao) Item 3 $c $colName
                      set pmiHeading($col($ao)) 1
                      set pmiCol [expr {max($col($ao),$pmiCol)}]
                      if {[string first "name" $colName] == 0} {
                        set comment "Section numbers refer to the CAx-IF Recommended Practice for Representation and Presentation of PMI (AP242)."
                        addCellComment $ao 3 $c $comment
                      }
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
                  set ov [string tolower $objValue]
                  set gpmiName $ov

# look for invalid 'name' values
                  set invalid ""
                  if {[string first "occurrence" $ao] != -1 && $opt(xlFormat) != "None"} {
                    if {$ov == "" || [lsearch $gpmiTypes $ov] == -1} {
                      if {$ov == ""} {
                        set msg "Missing 'name' attribute on [formatComplexEnt [lindex $ent1 0]]"
                      } else {
                        set msg "The [formatComplexEnt [lindex $ent1 0]] 'name' attribute is not a recommended name for presented PMI type."
                      }
                      if {[string first "AP242" $stepAP] == 0} {
                        append msg " ($recPracNames(pmi242), Sec. 8.4)"
                      } else {
                        append msg " ($recPracNames(pmi203), Sec. 4.3)"
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
                  set ov $objValue

# start x3dom file, read tessellated geometry
                  if {$gen(View) && $opt(viewPMI) && [string first "occurrence" $ao] != -1} {
                    if {$x3dStartFile} {x3dFileStart}

# moved (start shape node if not tessellated)
                    if {$ao == "annotation_fill_area_occurrence"} {errorMsg " PMI annotations with filled characters are not filled."}
                    if {[string first "tessellated" $ao] == -1 && [string first "placeholder" $ao] == -1} {set x3dShape 1}
                    update
                  }

# value in spreadsheet
                  if {[info exists gpmiIDRow($ao,$gpmiID)] && [string first "occurrence" $ao] != -1 && $opt(PMIGRF) && $opt(xlFormat) != "None"} {
                    set val [[$cells($ao) Item $r $c] Value]
                    if {$invalid != ""} {lappend syntaxErr($ao) [list "-$r" $col($ao) $invalid]}

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
                } elseif {[info exists gpmiIDRow($ao,$gpmiID)] && $opt(PMIGRF) && $opt(xlFormat) != "None"} {
                  if {$entLevel > 1} {
                    if {[string first "color" $colName] == -1} {
                      $cells($ao) Item $r $c "$ent($entLevel) $objID"
                    } else {
                      if {$ent($entLevel) == "colour_rgb"} {
                        $cells($ao) Item $r $c "$ent($entLevel) $objID  ($colorRGB)"
                      } else {
                        $cells($ao) Item $r $c "$ent($entLevel) $objID  ($objValue)"
                      }
                    }
                  }
                }
              }
            }
          } emsg3]} {
            set msg "Error processing Graphical PMI ($objNodeType $ent2): $emsg3"
            errorMsg $msg
            lappend syntaxErr([lindex $ent1 0]) [list $objID [lindex $ent1 1] $msg]
            set entLevel 2
          }
        }
      }
    }
  }
  incr entLevel -1

# write a few more things at the end of processing an annotation_occurrence entity
  if {$entLevel == 0 && $opt(PMIGRF) && $opt(xlFormat) != "None" && [info exists gpmiIDRow($ao,$gpmiID)]} {

# associated geometry, (1) find link between annotation_occurrence and a geometric item through
# draughting_model_item_association or draughting_callout and geometric_item_specific_usage
# (2) find link to PMI representation
    if {[catch {
      foreach var {assocSPMI assocGeom} {if {[info exists $var]} {unset $var}}

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
        ::tcom::foreach attrDMIA [$objGuiEntity Attributes] {
          if {[$attrDMIA Name] == "name"} {set attrName [$attrDMIA Value]}

# get shape_aspect (dmiaDef)
          if {[$attrDMIA Name] == "definition"} {
            set dmiaDef [$attrDMIA Value]
            if {[string first "handle" $dmiaDef] != -1} {
              set dmiaDefType [$dmiaDef Type]

# look for link to pmi representation
              if {$attrName == "PMI representation to presentation link"} {
                if {([string first "shape_aspect" $dmiaDefType] == -1 && [string first "property_definition" $dmiaDefType] == -1) || \
                     [string first "_datum_feature" $dmiaDefType] != -1} {
                  set spmi_p21id [$dmiaDef P21ID]
                  if {![info exists assocSPMI($dmiaDefType)]} {
                    lappend assocSPMI($dmiaDefType) $spmi_p21id
                  } elseif {[lsearch $assocSPMI($dmiaDefType) $spmi_p21id] == -1} {
                    lappend assocSPMI($dmiaDefType) $spmi_p21id
                  }
                } elseif {[string first "property_definition" $dmiaDefType] == -1} {
                  set msg "Syntax Error: Bad 'definition' attribute on draughting_model_item_association when 'name' attribute is 'PMI representation to presentation link'.$spaces\($recPracNames(pmi242), Sec. 7.3)"
                  errorMsg $msg
                  lappend syntaxErr(draughting_model_item_association) [list [$objGuiEntity P21ID] definition $msg]
                }

# look at shape_aspect or datums to find associated geometry
              } elseif {[string first "shape_aspect" $dmiaDefType] != -1 || [string first "datum" $dmiaDefType] != -1} {
                getAssocGeom $dmiaDef 1 $ao
              }
            } elseif {$opt(xlFormat) != "None"} {
              set msg "Syntax Error: Missing 'definition' attribute on draughting_model_item_association$spaces"
              if {[string first "AP242" $stepAP] == 0} {
                append msg "($recPracNames(pmi242), Sec. 9.3.1, Fig. 89)"
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
                set msg "Syntax Error: Bad 'used_representation' attribute ($dmiaDefType) on draughting_model_item_association.$spaces\($recPracNames(pmi242), Sec. 7.3)"
                errorMsg $msg
                lappend syntaxErr([$objGuiEntity Type]) [list [$objGuiEntity P21ID] "used_representation" $msg]
              }
            } else {
              set msg "Syntax Error: Missing 'used_representation' attribute on draughting_model_item_association.$spaces\($recPracNames(pmi242), Sec. 7.3)"
              errorMsg $msg
              lappend syntaxErr([$objGuiEntity Type]) [list [$objGuiEntity P21ID] "used_representation" $msg]
            }
          }
        }
      }
    } emsg]} {
      errorMsg "Error adding Associated Geometry: $emsg"
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
      if {[llength $aps] == 0 && $opt(xlFormat) != "None"} {
        set msg "Syntax Error: Annotation missing a required 'annotation_plane'.$spaces\($recPracNames(pmi242), Sec. 9.1, Fig. 86)"
        errorMsg $msg
        lappend syntaxErr($ao) [list $objID "plane" $msg]
      }

      foreach ap $aps {
        ::tcom::foreach attrAP [$ap Attributes] {
          if {[$attrAP Name] == "name"} {
            set str "[$ap Type] [$ap P21ID]"
            set nam [$attrAP Value]
            if {$nam != ""} {append str "[format "%c" 10]($nam)"}
            if {![info exists pmiColumns(aplane)]} {set pmiColumns(aplane) [getNextUnusedColumn $ao]}
            if {[string first "AP242" $stepAP] == 0} {
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

# check plane for the annotation plane
            set pl [[[$ap Attributes] Item [expr 3]] Value]
            set a2p3d [[[$pl Attributes] Item [expr 2]] Value]
            set axis [[[[[[$a2p3d Attributes] Item [expr 3]] Value] Attributes] Item [expr 2]] Value]
            set refdir [[[[[[$a2p3d Attributes] Item [expr 4]] Value] Attributes] Item [expr 2]] Value]
            if {[veclen [veccross $axis $refdir]] == 0} {
              set msg "Syntax Error: The axis2_placement_3d axis and ref_direction vectors '$refdir' are parallel for the 'annotation_plane' plane."
              errorMsg $msg
              lappend syntaxErr([$objEntity Type]) [list [$objEntity P21ID] "plane" $msg]
            }
          }
        }
      }
    } emsg]} {
      errorMsg "Error reporting Annotation Plane: $emsg"
    }

# report associated geometry
    if {[catch {
      if {[info exists assocGeom]} {
        set str [reportAssocGeom $ao $gpmiIDRow($ao,$gpmiID)]
        if {$str != ""  } {
          if {![info exists pmiColumns(ageom)]} {set pmiColumns(ageom) [getNextUnusedColumn $ao]}
          if {[string first "AP242" $stepAP] == 0} {
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
            set comment "See Help > User Guide (section 6.1.5) for an explanation of Associated Geometry."
            addCellComment $ao 3 $c $comment
          }
          $cells($ao) Item $r $pmiColumns(ageom) [string trim $str]

# supplemental geometry comment
          if {[string first "*" $str] != -1} {
            set comment "See Help > User Guide (section 6.1.5) for an explanation of Associated Geometry.  IDs marked with an asterisk (*) are also Supplemental Geometry."
            addCellComment $ao 3 $pmiColumns(ageom) $comment
          }
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
          if {![info exists pmiColumns(spmi)]} {set pmiColumns(spmi) [getNextUnusedColumn $ao]}
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
      errorMsg "Error reporting Associated Geometry and Representation: $emsg"
    }
  }

# report camera models associated with the annotation occurrence (not placeholder) through draughting_model
  if {$entLevel == 0 && (($opt(PMIGRF) && $opt(xlFormat) != "None") || ($gen(View) && $opt(viewPMI)))} {
    if {[catch {
      set savedViews ""
      set savedViewName {}
      set nsv 0

# get used draughting_model entities
      if {[info exists draftModelCameras] && [string first "placeholder" $ao] == -1} {
        set okdm 0
        set entDraughtingModels {}
        foreach dm $draughtingModels {
          ::tcom::foreach e0 [$objEntity GetUsedIn [string trim $dm] [string trim items]] {lappend entDraughtingModels $e0}

# check for draughting_callout.contents -> ao (PMI RP, section 9.4.4, figure 102)
          ::tcom::foreach entDraughtingCallout [$objEntity GetUsedIn [string trim draughting_callout] [string trim contents]] {
            ::tcom::foreach e0 [$entDraughtingCallout GetUsedIn [string trim $dm] [string trim items]] {lappend entDraughtingModels $e0}
          }

# check if there are any entDraughtingModel, if none then there are no camera models for the annotation
          foreach entDraughtingModel $entDraughtingModels {incr okdm}
        }

# get saved view names
        if {$okdm > 0} {
          foreach entDraughtingModel $entDraughtingModels {
            if {[info exists draftModelCameras([$entDraughtingModel P21ID])]} {
              set str $draftModelCameras([$entDraughtingModel P21ID])
              if {[string first $str $savedViews] == -1} {
                append savedViews $str
                incr nsv
              }
              lappend savedViewName $draftModelCameraNames([$entDraughtingModel P21ID])

              if {$opt(PMIGRF) && $opt(xlFormat) != "None" && [info exists gpmiIDRow($ao,$gpmiID)]} {
                if {[string first "AP242" $stepAP] == 0} {
                  set colName "Saved Views[format "%c" 10](Sec. 9.4)"
                } else {
                  set colName "Saved Views[format "%c" 10](Sec. 5.4)"
                }
                if {![info exists savedViewCol]} {set savedViewCol [getNextUnusedColumn $ao]}
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
                if {[string first "()" $savedViews] != -1 && $opt(xlFormat) != "None"} {
                  set msg "Syntax Error: For Saved Views, missing required 'name' attribute on camera_model_d3$spaces"
                  if {[string first "AP242" $stepAP] == 0} {
                    append msg "($recPracNames(pmi242), Sec. 9.4.2.1, Fig. 95)"
                  } else {
                    append msg "($recPracNames(pmi203), Sec. 5.4.2.1, Fig. 14)"
                  }
                  lappend syntaxErr($ao) [list "-$r" $savedViewCol $msg]
                  errorMsg $msg
                }
                if {[lsearch $gpmiRow($ao) $r] == -1} {lappend gpmiRow($ao) $r}
              }

# check for a mapped_item in draughting_model 'items', do not check style_item (see old code)
              set attrsDraughtingModel [$entDraughtingModel Attributes]
              ::tcom::foreach attrDraughtingModel $attrsDraughtingModel {
                if {[$attrDraughtingModel Name] == "name"} {set nameDraughtingModel [$attrDraughtingModel Value]}
                if {[$attrDraughtingModel Name] == "items" && $nameDraughtingModel != ""} {
                  set okcm 0
                  set okmi 0
                  set oksi 0
                  ::tcom::foreach item [$attrDraughtingModel Value] {
                    set itype [$item Type]
                    if {$itype == "mapped_item"} {set okmi 1}
                    if {[string first "camera_model_d3" $itype] == 0} {set okcm 1}
                  }

                  if {$okcm} {
                    if {$okmi == 0 && $opt(xlFormat) != "None"} {
                      set msg "Syntax Error: For Saved Views, missing required reference to 'mapped_item' on [formatComplexEnt [$entDraughtingModel Type]] 'items'$spaces"
                      if {[string first "AP242" $stepAP] == 0} {
                        append msg "($recPracNames(pmi242), Sec. 9.4.2.1, Fig. 95)"
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
              if {[string first "AP242" $stepAP] == 0} {
                set relType ""
                foreach item [list mechanical_design_and_draughting_relationship representation_relationship] {
                  if {[info exists entCount($item)]} {if {$entCount($item) > 0} {set relType $item}}
                }
                if {$relType != ""} {
                  set ok 1
                  set rep1Ents [$entDraughtingModel GetUsedIn [string trim $relType] [string trim rep_1]]
                  ::tcom::foreach rep1Ent $rep1Ents {set ok 0}
                  if {$ok && $opt(xlFormat) != "None"} {
                    set msg "Syntax Error: For Saved Views, '$relType' reference to '[formatComplexEnt [$entDraughtingModel Type]]' uses rep_2 instead of rep_1$spaces"
                    append msg "($recPracNames(pmi242), Sec. 9.4.4 Note 1, Fig. 104, Table 18)"
                    errorMsg $msg
                    set rep2Ents [$entDraughtingModel GetUsedIn [string trim $relType] [string trim rep_2]]
                    ::tcom::foreach rep2Ent $rep2Ents {set mdadrID [$rep2Ent P21ID]}
                    lappend syntaxErr($relType) [list $mdadrID rep_2 $msg]
                  }
                  if {$relType == "representation_relationship" && $opt(xlFormat) != "None"} {
                    set msg "Syntax Error: For Saved Views, use 'mechanical_design_and_draughting_relationship' instead of 'representation_relationship' to relate draughting models$spaces"
                    if {[string first "AP242" $stepAP] == 0} {
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

# not in a saved view
        } elseif {$opt(PMIGRF) && $opt(xlFormat) != "None"} {

# report missing saved view only if not text, etc.
          set oknm 1
          foreach str {note title block label text} {if {[string first $str $gpmiName] != -1} {set oknm 0}}
          if {$oknm} {
            set msg "An [$objEntity Type] is not in a Saved View.  If the annotation should be in a Saved View, then check draughting_model 'items' for a missing draughting_callout related to the annotation.  Also check the View for Graphical PMI to see if the annotations are not in a Saved View.\n  "
            if {[string first "AP242" $stepAP] == 0} {
              append msg "($recPracNames(pmi242), Sec. 9.4.2.1, Fig. 95)"
            } else {
              append msg "($recPracNames(pmi203), Sec. 5.4.2, Fig. 14)"
            }
            errorMsg $msg
            lappend syntaxErr($ao) [list $objID "Saved Views" $msg]
          }
        }
        catch {unset entDraughtingModel}
        catch {unset entDraughtingModels}

# missing MDADR or RR
        set relTypes [list representation_relationship]
        if {[string first "AP214" $stepAP] == -1} {lappend relTypes mechanical_design_and_draughting_relationship}
        set relType ""
        foreach item $relTypes {if {[info exists entCount($item)]} {if {$entCount($item) > 0} {set relType $item}}}
        if {$relType == "" && $opt(xlFormat) != "None"} {
          set str "mechanical_design_and_draughting_relationship"
          if {[string first "AP214" $stepAP] == 0} {set str "representation_relationship"}
          set msg "Syntax Error: For Saved Views, missing '$str' to relate 'draughting_model'$spaces"
          if {[string first "AP242" $stepAP] == 0} {
            append msg "($recPracNames(pmi242), Sec. 9.4.4 Note 1, Fig. 104)"
          } else {
            append msg "($recPracNames(pmi203), Sec. 5.4.4 Note 1, Fig. 20)"
          }
          errorMsg $msg
        }
      }
    } emsg]} {
      errorMsg "Error adding Saved Views: $emsg"
    }
  }

# check if there are PMI validation properties (propDefIDs) associated with the annotation_occurrence
  if {$entLevel == 0 && $opt(PMIGRF) && $opt(xlFormat) != "None" && [info exists gpmiIDRow($ao,$gpmiID)]} {
    if {[info exists propDefIDs]} {

# look for annotation_occurrence used in property_definition.definition
      if {[info exists propDefIDs($objID)]} {append gpmiValProp($objID) "$propDefIDs($objID) "}

# look for annotation_occurrence used in characterized_item_within_representation.item
      set e0s [$objEntity GetUsedIn [string trim characterized_item_within_representation] [string trim item]]
      ::tcom::foreach e0 $e0s {
        set e0id [$e0 P21ID]
        if {[info exists propDefIDs($e0id)]} {append gpmiValProp($objID) "$propDefIDs($e0id) "}
      }

# look for annotation_occurrence used in draughting_callout.contents
      set e0s [$objEntity GetUsedIn [string trim draughting_callout] [string trim contents]]
      ::tcom::foreach e0 $e0s {
        set e1s [$e0 GetUsedIn [string trim characterized_item_within_representation] [string trim item]]
        ::tcom::foreach e1 $e1s {
          set e1id [$e1 P21ID]
          if {[info exists propDefIDs($e1id)]} {append gpmiValProp($objID) "$propDefIDs($e1id) "}
        }
      }
    }

# add valprop info to spreadsheet
    if {[info exists gpmiValProp($objID)]} {
      if {![info exists pmiColumns(vp)]} {set pmiColumns(vp) [getNextUnusedColumn $ao]}
      set c $pmiColumns(vp)
      set r $gpmiIDRow($ao,$gpmiID)
      valPropColumn $ao $r $c $gpmiValProp($objID)
    }
  }
}

# -------------------------------------------------------------------------------
# get camera models
proc pmiGetCameras {} {
  global objDesign
  global draughtingModels draftModelCameraNames draftModelCameras entCount mytemp opt recPracNames savedViewFile
  global savedViewFileName savedViewItems savedViewName savedViewNames savedViewpoint spaces spmiTypesPerFile stepAP syntaxErr

  catch {unset draftModelCameras}
  catch {unset draftModelCameraNames}
  checkTempDir

# camera list
  set cmlist {}
  foreach cms [list camera_model_d3 camera_model_d3_multi_clipping camera_model_d3_multi_clipping_intersection \
                    camera_model_d3_multi_clipping_union camera_model_d3_with_hlhsr] {
    if {[info exists entCount($cms)]} {if {$entCount($cms) > 0} {lappend cmlist $cms}}
  }

  if {[llength $cmlist] > 0} {
    if {[catch {

# loop over camera model entities
      foreach cm $cmlist {
        ::tcom::foreach entCameraModel [$objDesign FindObjects [string trim $cm]] {
          if {[$entCameraModel Type] == $cm} {
            set attrCameraModels [$entCameraModel Attributes]

# loop over draughting model entities
            foreach dm $draughtingModels {
              set entDraughtingModels [$entCameraModel GetUsedIn [string trim $dm] [string trim items]]
              ::tcom::foreach entDraughtingModel $entDraughtingModels {
                set attrDraughtingModels [$entDraughtingModel Attributes]
                set dmitems([$entDraughtingModel P21ID]) ""
                set annForDM([$entDraughtingModel P21ID]) 0

# DM name attribute
                set nattr 0
                set iattr 1
                if {[string first "object" $dm] != -1} {set iattr 3}
                ::tcom::foreach attrDraughtingModel $attrDraughtingModels {
                  incr nattr
                  set nameDraughtingModel [$attrDraughtingModel Name]
                  if {$nameDraughtingModel == "name" && $nattr == $iattr} {
                    set name [$attrDraughtingModel Value]
                    if {$name == ""} {
                      set msg "Syntax Error: For viewpoints, missing required 'name' attribute on [formatComplexEnt $dm]$spaces"
                      if {[string first "AP242" $stepAP] == 0} {
                        append msg "($recPracNames(pmi242), Sec. 9.4.2)"
                      } else {
                        append msg "($recPracNames(pmi203), Sec. 5.4.2)"
                      }
                      errorMsg $msg
                      lappend syntaxErr($dm) [list [$entDraughtingModel P21ID] name $msg]
                    }
                  }
                  if {$nameDraughtingModel == "items"} {
                    ::tcom::foreach item [$attrDraughtingModel Value] {
                      set itype [$item Type]
                      if {$itype != "mapped_item" && [string first "camera_model_d3" $itype] == -1} {
                        append dmitems([$entDraughtingModel P21ID]) "[$item P21ID] "
                      }
                      if {[string first "annotation" $itype] != -1 || [string first "_callout" $itype] != -1} {
                        set annForDM([$entDraughtingModel P21ID]) 1
                      }
                    }
                  }
                }

# CM name attribute
                ::tcom::foreach attrCameraModel $attrCameraModels {
                  set nameCameraModel [$attrCameraModel Name]
                  if {$nameCameraModel == "name"} {
                    set name [$attrCameraModel Value]
                    set name1 [string trim $name]
                    if {$name1 == ""} {set name1 "Missing name"}

                    if {$name == ""} {
                      set msg "Syntax Error: For viewpoints, missing required 'name' attribute on $cm$spaces"
                      if {[string first "AP242" $stepAP] == 0} {
                        append msg "($recPracNames(pmi242), Sec. 9.4.2.1, Fig. 95)"
                      } else {
                        append msg "($recPracNames(pmi203), Sec. 5.4.2.1, Fig. 14)"
                      }
                      errorMsg $msg
                      lappend syntaxErr($cm) [list [$entCameraModel P21ID] name $msg]
                    }

# check for default saved view
                    ::tcom::foreach e0 [$entCameraModel GetUsedIn [string trim default_model_geometric_view] [string trim item]] {
                      if {[$e0 Type] == "default_model_geometric_view"} {append name " (default)"; append name1 " (default)"}
                    }

# get axis2_placement_3d for camera viewpoint
                  } elseif {$nameCameraModel == "view_reference_system"} {
                    catch {unset savedViewpoint($name1)}
                    if {[catch {
                      set a2p3d [[$attrCameraModel Value] Attributes]
                      set origin [[[[[$a2p3d Item [expr 2]] Value] Attributes] Item [expr 2]] Value]
                      set axis   [[[[[$a2p3d Item [expr 3]] Value] Attributes] Item [expr 2]] Value]
                      set refdir [[[[[$a2p3d Item [expr 4]] Value] Attributes] Item [expr 2]] Value]
                      lappend savedViewpoint($name1) [vectrim $origin]
                      lappend savedViewpoint($name1) [x3dGetRotation $axis $refdir]
                    } emsg]} {
                      errorMsg "Error getting viewpoint position and orientation: $emsg"
                      catch {raise .}
                    }

# view_volume > view_plane_distance, projection_point, planar_box > x, y, a2p3d
                  } elseif {$nameCameraModel == "perspective_of_volume" && [info exists savedViewpoint($name1)]} {

# view volume
                    set vv [[$attrCameraModel Value] Attributes]
                    set vpd [[$vv Item [expr 3]] Value]
                    lappend savedViewpoint($name1) $vpd

# projection point, should be 0 0 0
                    set pp [vectrim [[[[[$vv Item [expr 2]] Value] Attributes] Item [expr 2]] Value]]
                    lappend savedViewpoint($name1) $pp

# planar box dimensions
                    set pb [[$vv Item [expr 9]] Value]
                    set pbx [trimNum [[[$pb Attributes] Item [expr 2]] Value]]
                    set pby [trimNum [[[$pb Attributes] Item [expr 3]] Value]]
                    lappend savedViewpoint($name1) $pbx
                    lappend savedViewpoint($name1) $pby

# planar box a2p3d
                    set a2p3d [[[$pb Attributes] Item [expr 4]] Value]
                    lappend savedViewpoint($name1) [vectrim [[[[[[$a2p3d Attributes] Item [expr 2]] Value] Attributes] Item [expr 2]] Value]]
                    set axis   [[[[[[$a2p3d Attributes] Item [expr 3]] Value] Attributes] Item [expr 2]] Value]
                    set refdir [[[[[[$a2p3d Attributes] Item [expr 4]] Value] Attributes] Item [expr 2]] Value]
                    lappend savedViewpoint($name1) [x3dGetRotation $axis $refdir]
                    lappend savedViewpoint($name1) $axis
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

# keep track of saved views for graphical PMI
                set dmcn $draftModelCameraNames([$entDraughtingModel P21ID])
                if {[lsearch $savedViewName $dmcn] == -1} {lappend savedViewName $dmcn}
                if {[lsearch $savedViewNames $name1] == -1} {
                  lappend savedViewNames $name1
                  if {($opt(viewPMI) || $opt(PMISEM)) && $annForDM([$entDraughtingModel P21ID])} {
                    if {$opt(PMISEM)} {lappend spmiTypesPerFile "saved views"}

# create temp file ViewN.txt for saved view graphical PMI x3d, where 'N' is an integer
                    if {$opt(viewPMI)} {
                      set name2 "View[lsearch $savedViewNames $name1]"
                      set fn [file join $mytemp $name2.txt]
                      catch {file delete -force -- $fn}
                      set savedViewFile($name2) [open $fn w]
                      set savedViewFileName($name2) $fn
                      if {[string length $dmitems([$entDraughtingModel P21ID])] > 0} {set savedViewItems($dmcn) $dmitems([$entDraughtingModel P21ID])}
                    }
                  }
                }
              }
            }
          }
        }
      }
    } emsg]} {
      errorMsg "Error getting Camera Models: $emsg"
      catch {raise .}
    }
  }
}
