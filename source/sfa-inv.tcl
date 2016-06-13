# check 'inverses' with GetUsedIn

proc invFind {objEntity} {
  global inverses invs invmsg opt developer

  set DEBUGINV $opt(DEBUGINV)
  if {$DEBUGINV} {outputMsg \ninvFind red}

  set objType [$objEntity Type]
  set objP21ID [$objEntity P21ID]
  set stat ""
  
  if {[catch {
    foreach inverse $inverses {
      set invEntities [$objEntity GetUsedIn [lindex $inverse 0] [lindex $inverse 1]]
      ::tcom::foreach invEntity $invEntities {
        set invType [$invEntity Type]
        set invRule [lindex $inverse 2]
        #outputMsg here2
        set msg " Inverse ($invRule) [formatComplexEnt $invType]:"
        set msgok 0
        set nattr 0
  
        ::tcom::foreach invAttribute [$invEntity Attributes] {
          set attrName  [string tolower [$invAttribute Name]]
          set attrValue [$invAttribute Value]
          if {$DEBUGINV} {
            incr nattr
            if {$nattr == 1} {outputMsg "Inverse [lindex $inverse 0]  [lindex $inverse 1]  [lindex $inverse 2]" blue}
            outputMsg "  attrName $attrName"
          }

# look for 'relating'
          if {[string first "relating" $attrName] != -1} {
            set stat "relating [lindex $inverse 0] [lindex $inverse 1]"
            if {$DEBUGINV} {outputMsg "relating $attrName [$invAttribute NodeType] [$invAttribute Value]" red}
            if {[string first "handle" $attrValue] != -1} {
              set subType [$attrValue Type]
              if {$DEBUGINV} {outputMsg " $subType [$attrValue P21ID]  $objType $objP21ID" red}
              if {[$attrValue P21ID] != $objP21ID} {
                if {$DEBUGINV} {outputMsg "  OK" red}
                lappend invs($invRule) "$subType [$attrValue P21ID]"
                if {[string first "DELETETHIS" $msg] == -1} {append msg " $subType DELETETHIS$objType"}
                set msgok 1
              }
            }

# look for 'related'
          } elseif {[string first "related" $attrName] != -1} {
            if {[string first "handle" $attrValue] != -1} {
              set stat "related [$invAttribute NodeType] [lindex $inverse 0] [lindex $inverse 1]"
              if {[$invAttribute NodeType] == 18 || [$invAttribute NodeType] == 19} {
                if {$DEBUGINV} {outputMsg "related $attrName [$invAttribute NodeType] [$invAttribute Value]" green}
                set subType [$attrValue Type]
                if {$DEBUGINV} {outputMsg " $subType [$attrValue P21ID] $objType $objP21ID" green}
                if {[$attrValue P21ID] != $objP21ID} {
                  if {$DEBUGINV} {outputMsg "  OK" green}
                  lappend invs($invRule) "$subType [$attrValue P21ID]"
                  if {[string first "DELETETHIS" $msg] == -1} {append msg " $subType DELETETHIS$objType"}
                  set msgok 1
                }

# still needs some work to remove the catch below           
# nodetype = 20
              } elseif {[$invAttribute NodeType] == 20 && $attrName != [string tolower [lindex $inverse 1]]} {
                set stat "related [$invAttribute NodeType] A [lindex $inverse 0] [lindex $inverse 1]"
                if {$DEBUGINV} {outputMsg "related $attrName [$invAttribute NodeType] [$invAttribute Value]" magenta}
                if {$DEBUGINV} {outputMsg " $objType $objP21ID / [$invEntity Type] [$invEntity P21ID]" magenta}
                if {[catch {
                  ::tcom::foreach aval [$invAttribute Value] {
                    if {$DEBUGINV} {outputMsg "  $aval / [$aval Type] [$aval P21ID] / $objType $objP21ID / [$invEntity Type] [$invEntity P21ID]" magenta}
                    if {[$aval P21ID] != $objP21ID && [$aval Type] != $objType} {
                      if {$DEBUGINV} {outputMsg "  OK A  [$aval Type] [$aval P21ID]" magenta}
                      lappend invs($invRule) "[$aval Type] [$aval P21ID]"
                      if {[string first "DELETETHIS" $msg] == -1} {append msg " [$aval Type] DELETETHIS$objType"}
                      set msgok 1
                    }
                  }
                } emsg1]} {

# still needs some work to remove the catch                
                  set stat "related [$invAttribute NodeType] B [lindex $inverse 0] [lindex $inverse 1]"
                  foreach aval [$invAttribute Value] {
                    catch {
                      if {$DEBUGINV} {outputMsg "  [$aval Type] [$aval P21ID] / $objType $objP21ID / [$invEntity Type] [$invEntity P21ID]" magenta}
                      if {[$aval P21ID] != $objP21ID && [$aval Type] != $objType} {
                        if {$DEBUGINV} {outputMsg "  OK B" magenta}
                        lappend invs($invRule) "[$aval Type] [$aval P21ID]"
                        if {[string first "DELETETHIS" $msg] == -1} {append msg " [$aval Type] DELETETHIS$objType"}
                        set msgok 1
                      }
                    }
                  }
                }
              }
            }

# look for other GetUsedIn relationships that are not relating/related Inverses 
          } else {
            set attrValue [$invAttribute Value]
            if {$DEBUGINV} {outputMsg "  attrValue(L=[llength $attrValue])  ($attrValue)  NodeType=[$invAttribute NodeType]"}
            if {[string first "handle" $attrValue] != -1} {

# nodetype = 20
              if {[$invAttribute NodeType] == 20} {
                if {[catch {
                  set stat "used in [$invAttribute NodeType] A [lindex $inverse 0] [lindex $inverse 1]"
                  ::tcom::foreach val $attrValue {
                    if {$DEBUGINV} {outputMsg "   $val [$val Type] [$val P21ID] / $objType $objP21ID / [$invEntity Type] [$invEntity P21ID]"}
                    if {[$val P21ID] == $objP21ID} {
                      #lappend invs($invRule) "[$invEntity Type].[lindex $inverse 1] [$invEntity P21ID]"
                      set str "[formatComplexEnt [$invEntity Type]].[lindex $inverse 1] [$invEntity P21ID]"
                      if {![info exists invs($invRule)]} {
                        lappend invs($invRule) $str
                      } elseif {[lsearch $invs($invRule) $str] == -1} {
                        lappend invs($invRule) $str
                      }
                      if {[string first "DELETETHIS" $msg] == -1} {append msg " [lindex $inverse 1] DELETETHIS$objType"}
                      set msgok 1
                    }
                  }
                } emsg]} {
                  set stat "used in [$invAttribute NodeType] B [lindex $inverse 0] [lindex $inverse 1]"
                  if {[catch {
                    if {$DEBUGINV} {outputMsg "   $val / $objType $objP21ID / [$invEntity Type] [$invEntity P21ID]"}
                    #lappend invs($invRule) "[$invEntity Type].[lindex $inverse 1] [$invEntity P21ID]"
                    set str "[$invEntity Type].[lindex $inverse 1] [$invEntity P21ID]"
                    if {![info exists invs($invRule)]} {
                      lappend invs($invRule) $str
                    } elseif {[lsearch $invs($invRule) $str] == -1} {
                      lappend invs($invRule) $str
                    }
                    if {[string first "DELETETHIS" $msg] == -1} {append msg " [lindex $inverse 1] DELETETHIS$objType"}
                    set msgok 1
                  } emsg1]} {
                    errorMsg "Error checking GetUsedIn relationship: $inverse"
                  }
                }

# STEP Inverses non-related/relating get processed here, nodetype = 18, 19
              } else {
                if {[catch {
                  set stat "used in [$invAttribute NodeType] C [lindex $inverse 0] [lindex $inverse 1]"
                  set subType [$attrValue Type]
                  if {$DEBUGINV} {outputMsg "     $subType [$attrValue P21ID] / $objType $objP21ID / [$invEntity Type] [$invEntity P21ID]"}
  
                  if {[$attrValue P21ID] == $objP21ID} {
                    if {$DEBUGINV} {outputMsg "     $invType  $attrValue  $subType"}
                    set str "[$invEntity Type].[lindex $inverse 1] [$invEntity P21ID]"
                    if {![info exists invs($invRule)]} {
                      lappend invs($invRule) $str
                    } elseif {[lsearch $invs($invRule) $str] == -1} {
                      lappend invs($invRule) $str
                    }
                    if {[string first "DELETETHIS" $msg] == -1} {append msg " [lindex $inverse 1] DELETETHIS$objType"}
                    set msgok 1
                  }

# special case for a nodetype 18 or 19 that is more like 20 (a list)
                } emsg2]} {
                  set stat "used in [$invAttribute NodeType] D [lindex $inverse 0] [lindex $inverse 1]"
                  ::tcom::foreach aval [$invAttribute Value] {
                    #if {$DEBUGINV} {outputMsg "  $aval / [$aval Type] [$aval P21ID] / $objType $objP21ID / [$invEntity Type] [$invEntity P21ID]" magenta}
                    if {[$aval P21ID] != $objP21ID && [$aval Type] != $objType} {
                      #if {$DEBUGINV} {outputMsg "  OK A  [$aval Type] [$aval P21ID]" magenta}
                      lappend invs($invRule) "[$aval Type] [$aval P21ID]"
                      if {[string first "DELETETHIS" $msg] == -1} {append msg " [$aval Type] DELETETHIS$objType"}
                      set msgok 1
                    }
                  }
                }
              }
            } else {
              if {$DEBUGINV} {outputMsg "   Non-handle value (plain text)"}
            }
          }
        }
    
        if {$msgok} {
          if {[string first "used_in" $msg] != -1} {
            set msg [string range $msg 19 end]
            regsub ":" $msg "." msg
            regsub " " $msg ""  msg
            set msg " Used In: $msg"
          } else {
            set lmsg [split $msg " "]
            set newmsg " Inverse: [string range [lindex $lmsg 3] 0 end-1].[string range [lindex $lmsg 2] 1 end-1] > [formatComplexEnt [lindex $lmsg 4]] [lindex $lmsg 5]"
            set msg $newmsg
          }
        }     
  
        if {$msgok && [string first $msg $invmsg] == -1} {
          append invmsg $msg
          errorMsg $msg blue
        }
      }
    }
  } emsg]} {
    if {!$developer} {
      errorMsg "ERROR processing Inverse for '[$objEntity Type]': $emsg"
    } else {
      errorMsg "ERROR processing Inverse for '[$objEntity Type]': $emsg\n ($stat)"
    }
  }
}

# -------------------------------------------------------------------------------
# report inverses 
proc invReport {} {
  global invs cellval cells thisEntType row col colinv
  #outputMsg invReport red
  
# inverse values and heading
  foreach item [array names invs] {
    catch {foreach idx [array names cellval] {unset cellval($idx)}}

    foreach val $invs($item) {
      set val [split $val " "]
      set val0 [lindex $val 0]
      set val1 "[lindex $val 1] "
      if {[info exists cellval($val0)]} {
        if {[string first $val1 $cellval($val0)] == -1} {
          append cellval($val0) $val1
        }
      } else {
        append cellval($val0) $val1        
      }
    }

    set str ""
    set size 0
    catch {set size [array size cellval]}

    if {$size > 0} {
      foreach idx [lsort [array names cellval]] {
        set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
        if {$ncell > 1 || $size > 1} {
          if {$ncell < 30} {
            if {[string length $str] > 0} {append str [format "%c" 10]}
            append str "($ncell) [invFormatComplexEnt $idx 1] $cellval($idx)"
          } else {
            if {[string length $str] > 0} {append str [format "%c" 10]}
            append str "($ncell) [invFormatComplexEnt $idx 1]"
          }
        } else {
          if {[string length $str] > 0} {append str [format "%c" 10]}
          append str "(1) [invFormatComplexEnt $idx 1] $cellval($idx)"
        }
      }
    }
    
    set idx "$thisEntType $item"
    if {[info exists colinv($idx)]} {
      $cells($thisEntType) Item $row($thisEntType) $colinv($idx) [string trim $str]
    } else {
      while {[[$cells($thisEntType) Item 3 $col($thisEntType)] Value] != ""} {incr col($thisEntType)}
      $cells($thisEntType) Item $row($thisEntType) $col($thisEntType) [string trim $str]
    }

# heading
    if {[[$cells($thisEntType) Item 3 $col($thisEntType)] Value] == ""} {
      if {$item != "used_in"} {
        $cells($thisEntType) Item 3 $col($thisEntType) "INV-$item"
        set idx "$thisEntType $item"
      } else {
        $cells($thisEntType) Item 3 $col($thisEntType) "Used In"
        set idx "$thisEntType $item"
      }
      set colinv($idx) $col($thisEntType)
    }
  }
}

# -------------------------------------------------------------------------------
# set column color, border, group for INVERSES and Used In
proc invFormat {rancol} {
  global thisEntType col cells row rowmax worksheet invGroup excel
  
  set igrp1 100
  set igrp2 0
  set i1 [expr {$rancol+1}]
  #set i1 [expr {$col($thisEntType)+20}]
    
# fix column widths
  for {set i 1} {$i <= $i1} {incr i} {
    set val [[$cells($thisEntType) Item 3 $i] Value]
    if {$val == "Used In" || [string first "INV-" $val] != -1} {
      set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
      $range ColumnWidth [expr 255]
    }
  }
  [$worksheet($thisEntType) Columns] AutoFit
  [$worksheet($thisEntType) Rows] AutoFit

# set colors, borders
  for {set i 1} {$i <= $i1} {incr i} {
    set val [[$cells($thisEntType) Item 3 $i] Value]
    if {[string first "INV-" $val] != -1 || [string first "Used In" $val] != -1} {
      #set r1 [expr {$row($thisEntType)+2}]
      set r1 $row($thisEntType)
      if {$r1 > $rowmax} {set r1 [expr {$r1-1}]}
      set range [$worksheet($thisEntType) Range [cellRange 3 $i] [cellRange $r1 $i]]
      if {[string first "INV-" $val] != -1} {
        [$range Interior] ColorIndex [expr 20]
      } else {
        [$range Interior] Color [expr 16768477]
      }
      if {$i < $igrp1} {set igrp1 $i}
      if {$i > $igrp2} {set igrp2 $i}
      if {[expr {int([$excel Version])}] >= 12} {
        set range [$worksheet($thisEntType) Range [cellRange 4 $i] [cellRange $r1 $i]]
        for {set k 7} {$k <= 12} {incr k} {
          if {$k != 9} {
            catch {[[$range Borders] Item [expr $k]] Weight [expr 1]}
          }
        }
        set range [$worksheet($thisEntType) Range [cellRange 3 $i] [cellRange 3 $i]]
        catch {
          [[$range Borders] Item [expr 7]]  Weight [expr 1]
          [[$range Borders] Item [expr 10]] Weight [expr 1]
        }
      }
    }
  }

# group
  if {$igrp2 > 0} {
    set grange [$worksheet($thisEntType) Range [cellRange 1 $igrp1] [cellRange [expr {$row($thisEntType)+2}] $igrp2]]
    [$grange Columns] Group
    set invGroup($thisEntType) $igrp1
  }
}

# -------------------------------------------------------------------------------
# decide if inverses should be checked for this entity type
# as more inverse relationships are checked, it is important that only entities with inverses are checked
# checking an entity without inverses is costly, particularly if there are a lot of them
proc invSetCheck {entType} {
  global opt entCategory
  
  set checkInv 0

# tolerance and AP209 entities
  if {($opt(PR_STEP_TOLR)  && [lsearch $entCategory(PR_STEP_TOLR)  $entType] != -1) || \
      ($opt(PR_STEP_AP209) && [lsearch $entCategory(PR_STEP_AP209) $entType] != -1)} {
    set checkInv 1

# entities related to shape_aspect, PMI, product
  } elseif { \
      [string first "action"        $entType] != -1 || \
      [string first "angular"       $entType] != -1 || \
      [string first "annotation"    $entType] != -1 || \
      [string first "application"   $entType] != -1 || \
      [string first "camera"        $entType] != -1 || \
      [string first "composite"     $entType] != -1 || \
      [string first "constructive"  $entType] != -1 || \
      [string first "datum"         $entType] != -1 || \
      [string first "dimension"     $entType] != -1 || \
      [string first "document"      $entType] != -1 || \
      [string first "item_"         $entType] != -1 || \
      [string first "machining"     $entType] != -1 || \
      [string first "product"       $entType] != -1 || \
      [string first "resource"      $entType] != -1 || \
      [string first "shape"         $entType] != -1 || \
      [string first "symmetry"      $entType] != -1 || \
      [string first "draughting_model"    $entType] != -1 || \
      [string first "draughting_callout"  $entType] != -1 || \
      [string first "mechanical_design"   $entType] != -1 || \
      [string first "next_assembly"       $entType] != -1 || \
      [string first "property_definition" $entType] != -1} {

# not these specific entities
        #$entType != "draughting_model_item_association" &&
    if {$entType != "annotation_fill_area" && \
        $entType != "dimensional_characteristic_representation" && \
        $entType != "draughting_pre_defined_colour" && \
        $entType != "draughting_pre_defined_curve_font" && \
        $entType != "geometric_item_specific_usage" && \
        $entType != "property_definition_representation" && \
        $entType != "shape_aspect_deriving_relationship" && \
        $entType != "shape_aspect_relationship"} {
      set checkInv 1
    }
  }
  return $checkInv
}

# -------------------------------------------------------------------------------
# set inverse relationships
proc initDataInverses {} {
  global inverses tolNames stepAP opt

  set inverses {}
  lappend inverses [list annotation_curve_occurrence item used_in]
  lappend inverses [list annotation_occurrence item used_in]
  lappend inverses [list annotation_occurrence_relationship related_annotation_occurrence relating_annotation_occurrence]
  lappend inverses [list annotation_occurrence_relationship relating_annotation_occurrence related_annotation_occurrence]
  lappend inverses [list annotation_plane elements used_in]
  lappend inverses [list applied_presented_item items used_in]

  lappend inverses [list characterized_item_within_representation item used_in]
  lappend inverses [list characterized_item_within_representation rep used_in]
  lappend inverses [list component_path_shape_aspect location used_in]
  lappend inverses [list component_path_shape_aspect component_shape_aspect used_in]
  lappend inverses [list context_dependent_shape_representation representation_relation used_in]
  lappend inverses [list context_dependent_shape_representation represented_product_relation used_in]

  lappend inverses [list dimension_callout_relationship related_draughting_callout relating_draughting_callout]
  lappend inverses [list dimension_callout_relationship relating_draughting_callout related_draughting_callout]

  lappend inverses [list document_product_equivalence related_product relating_document]
  lappend inverses [list document_product_equivalence relating_document related_product]

  lappend inverses [list draughting_callout contents used_in]
  lappend inverses [list draughting_model items used_in]
  lappend inverses [list draughting_model_item_association definition used_in]
  lappend inverses [list draughting_model_item_association identified_item used_in]
  lappend inverses [list draughting_model_item_association used_representation used_in]

  lappend inverses [list geometric_item_specific_usage definition used_in]
  lappend inverses [list geometric_item_specific_usage identified_item used_in]
  lappend inverses [list geometric_item_specific_usage used_representation used_in]
  lappend inverses [list geometric_tolerance_relationship related_geometric_tolerance relating_geometric_tolerance]
  lappend inverses [list geometric_tolerance_relationship relating_geometric_tolerance related_geometric_tolerance]

  lappend inverses [list id_attribute identified_item used_in]
  lappend inverses [list kinematic_joint edge_end used_in]
  lappend inverses [list kinematic_joint edge_start used_in]
  lappend inverses [list leader_directed_callout contents used_in]

  lappend inverses [list make_from_usage_option related_product_definition relating_product_definition]
  lappend inverses [list make_from_usage_option relating_product_definition related_product_definition]
  lappend inverses [list multi_level_reference_designator location used_in]
  lappend inverses [list multi_level_reference_designator related_product_definition relating_product_definition]
  lappend inverses [list multi_level_reference_designator relating_product_definition related_product_definition]
  lappend inverses [list multi_level_reference_designator location used_in]

  lappend inverses [list next_assembly_usage_occurrence related_product_definition relating_product_definition]
  lappend inverses [list next_assembly_usage_occurrence relating_product_definition related_product_definition]

  lappend inverses [list pair_representation_relationship rep_1 used_in]
  lappend inverses [list pair_representation_relationship rep_2 used_in]

  lappend inverses [list product frame_of_reference used_in]
  lappend inverses [list product_context frame_of_reference used_in]
  lappend inverses [list product_definition formation used_in]
  lappend inverses [list product_definition frame_of_reference used_in]
  lappend inverses [list product_definition_context frame_of_reference used_in]
  lappend inverses [list product_definition_formation of_product used_in]
  lappend inverses [list product_definition_shape definition used_in]
  lappend inverses [list product_related_product_category products used_in]

  lappend inverses [list property_definition definition used_in]
  lappend inverses [list referenced_modified_datum referenced_datum used_in]
  lappend inverses [list representation_relationship_with_transformation transformation_operator used_in]
  lappend inverses [list representation_relationship_with_transformation_and_shape_representation_relationship transformation_operator used_in]
  lappend inverses [list rigid_link_representation represented_link used_in]
  lappend inverses [list roundness_definition projection_end used_in]

  lappend inverses [list shape_aspect of_shape used_in]
  lappend inverses [list shape_aspect_relationship related_shape_aspect relating_shape_aspect]
  lappend inverses [list shape_aspect_relationship relating_shape_aspect related_shape_aspect]
  lappend inverses [list shape_definition_representation definition used_in]
  lappend inverses [list shape_definition_representation used_representation used_in]
  lappend inverses [list shape_representation_relationship rep1 used_in]
  lappend inverses [list shape_representation_relationship rep2 used_in]

  lappend inverses [list tessellated_annotation_occurrence item used_in]

# tolerance related
  lappend inverses [list datum_reference referenced_datum used_in]
  lappend inverses [list datum_reference_compartment base used_in]
  lappend inverses [list datum_reference_element base used_in]
  lappend inverses [list datum_system constituents used_in]
  lappend inverses [list dimensional_characteristic_representation dimension used_in]
  lappend inverses [list dimensional_characteristic_representation representation used_in]
  lappend inverses [list dimensional_size applies_to used_in]
  lappend inverses [list non_uniform_zone_definition zone used_in]
  lappend inverses [list plus_minus_tolerance toleranced_dimension used_in]
  lappend inverses [list projected_zone_definition projection_end used_in]
  lappend inverses [list projected_zone_definition zone used_in]
  lappend inverses [list tolerance_zone defining_tolerance used_in]
  
  set modlist [list geometric_tolerance_with_datum_reference \
                    geometric_tolerance_with_datum_reference_and_geometric_tolerance_with_modifiers \
                    geometric_tolerance_with_datum_reference_and_modified_geometric_tolerance]

  foreach tol $tolNames {
    lappend inverses [list $tol toleranced_shape_aspect used_in]
    lappend inverses [list $tol datum_system used_in]
  
    foreach mod $modlist {
      if {[string compare [string index $mod 0] [string index $tol 0]] == -1} {
        lappend inverses [list $mod\_and_$tol toleranced_shape_aspect used_in]
        lappend inverses [list $mod\_and_$tol datum_system used_in]
        lappend inverses [list $mod\_and_$tol\_and_unequally_disposed_geometric_tolerance toleranced_shape_aspect used_in]
        lappend inverses [list $mod\_and_$tol\_and_unequally_disposed_geometric_tolerance datum_system used_in]
      } else {
        lappend inverses [list $tol\_and_$mod toleranced_shape_aspect used_in]
        lappend inverses [list $tol\_and_$mod datum_system used_in]
        lappend inverses [list $tol\_and_$mod\_and_unequally_disposed_geometric_tolerance toleranced_shape_aspect used_in]
        lappend inverses [list $tol\_and_$mod\_and_unequally_disposed_geometric_tolerance datum_system used_in]
      }
    }
  }
  
# AP209 related
  lappend inverses [list curve_3d_element_representation node_list used_in]
  lappend inverses [list nodal_freedom_action_definition degrees_of_freedom used_in]
  lappend inverses [list nodal_freedom_action_definition node used_in]
  lappend inverses [list nodal_freedom_values degrees_of_freedom used_in]
  lappend inverses [list nodal_freedom_values node used_in]
  lappend inverses [list node_geometric_relationship node_ref used_in]
  lappend inverses [list node_group nodes used_in]
  lappend inverses [list single_point_constraint_element freedoms_and_values used_in]
  lappend inverses [list single_point_constraint_element required_node used_in]
  lappend inverses [list surface_3d_element_representation node_list used_in]
  lappend inverses [list volume_3d_element_representation node_list used_in]
  #lappend inverses [list surface_3d_element_location_point_volume_variable_values values_and_locations used_in]
  #lappend inverses [list volume_3d_element_location_point_volume_variable_values values_and_locations used_in]
  
  lappend inverses [list action_relationship related_action relating_action]
  lappend inverses [list action_relationship relating_action related_action]
  lappend inverses [list action_resource usage used_in]
  lappend inverses [list action_resource kind used_in]
  lappend inverses [list process_product_association defined_product used_in]
  lappend inverses [list process_product_association process used_in]
  lappend inverses [list requirement_for_action_resource kind used_in]
  lappend inverses [list requirement_for_action_resource operations used_in]
  lappend inverses [list requirement_for_action_resource resources used_in]
  lappend inverses [list resource_property_representation property used_in]
  lappend inverses [list resource_property_representation representation used_in]
}

# -------------------------------------------------------------------------------
proc invFormatComplexEnt {str {space 0}} {

  set str1 $str
  set lstr [split $str1 "."]
  if {[string first "_and_" [lindex $lstr 0]] != -1} {
    set str1 "[string trimright [formatComplexEnt [lindex $lstr 0] $space].[lindex $lstr 1] "."]"
  }
  return $str1
}
