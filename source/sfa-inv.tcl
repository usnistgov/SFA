# check 'inverses' with GetUsedIn

proc invFind {objEntity} {
  global inverses invVals invMsg opt developer

  set objType [$objEntity Type]
  set objP21ID [$objEntity P21ID]
  set stat ""

  set DEBUGINV $opt(DEBUGINV)
  if {$DEBUGINV} {outputMsg "invFind  $objP21ID=$objType  [llength $inverses]" red}
  #errorMsg "invFind  $objType  [llength $inverses]" red
  
  if {[catch {
    foreach inverse $inverses {
      catch {unset invEnts}
      set invEnts [$objEntity GetUsedIn [join [lindex $inverse 0]] [join [lindex $inverse 1]]]
      ::tcom::foreach invEnt $invEnts {
        set invType [$invEnt Type]
        set invRule [lindex $inverse 2]
        set msg " Inverse ($invRule) [formatComplexEnt $invType]:"
        set msgok 0
        set nattr 0
  
        ::tcom::foreach invAttr [$invEnt Attributes] {
          set attrName  [string tolower [$invAttr Name]]
          set attrValue [$invAttr Value]
          if {$DEBUGINV} {
            incr nattr
            if {$nattr == 1} {outputMsg "Inverse [lindex $inverse 0]  [lindex $inverse 1]  [lindex $inverse 2]" blue}
            outputMsg "  attrName $attrName"
          }

# look for 'relating' or 'rep_1'
          if {[string first "relating" $attrName] != -1 || [string first "rep_1" $attrName] != -1} {
            set stat "relating [lindex $inverse 0] [lindex $inverse 1]"
            if {$DEBUGINV} {outputMsg "relating $attrName [$invAttr NodeType] [$invAttr Value]" red}
            if {[string first "handle" $attrValue] != -1} {
              set subType [$attrValue Type]
              if {$DEBUGINV} {outputMsg " $subType [$attrValue P21ID]  $objType $objP21ID" red}
              if {[$attrValue P21ID] != $objP21ID} {
                if {$DEBUGINV} {outputMsg "  OK" red}
                lappend invVals($invRule) "$subType [$attrValue P21ID]"
                if {[string first "DELETETHIS" $msg] == -1} {append msg " $subType DELETETHIS$objType"}
                set msgok 1
              }
            }

# look for 'related' or 'rep_2'
          } elseif {[string first "related" $attrName] != -1 || [string first "rep_2" $attrName] != -1} {
            if {[string first "handle" $attrValue] != -1} {
              set stat "related [$invAttr NodeType] [lindex $inverse 0] [lindex $inverse 1]"
              if {[$invAttr NodeType] == 18 || [$invAttr NodeType] == 19} {
                if {$DEBUGINV} {outputMsg "related $attrName [$invAttr NodeType] [$invAttr Value]" green}
                set subType [$attrValue Type]
                if {$DEBUGINV} {outputMsg " $subType [$attrValue P21ID] $objType $objP21ID" green}
                if {[$attrValue P21ID] != $objP21ID} {
                  if {$DEBUGINV} {outputMsg "  OK" green}
                  lappend invVals($invRule) "$subType [$attrValue P21ID]"
                  if {[string first "DELETETHIS" $msg] == -1} {append msg " $subType DELETETHIS$objType"}
                  set msgok 1
                }

# still needs some work to remove the catch below           
# nodetype = 20
              } elseif {[$invAttr NodeType] == 20 && $attrName != [string tolower [lindex $inverse 1]]} {
                set stat "related [$invAttr NodeType] A [lindex $inverse 0] [lindex $inverse 1]"
                if {$DEBUGINV} {outputMsg "related $attrName [$invAttr NodeType] [$invAttr Value]" magenta}
                if {$DEBUGINV} {outputMsg " $objType $objP21ID / [$invEnt Type] [$invEnt P21ID]" magenta}
                if {[catch {
                  ::tcom::foreach aval [$invAttr Value] {
                    if {$DEBUGINV} {outputMsg "  $aval / [$aval Type] [$aval P21ID] / $objType $objP21ID / [$invEnt Type] [$invEnt P21ID]" magenta}
                    if {[$aval P21ID] != $objP21ID && [$aval Type] != $objType} {
                      if {$DEBUGINV} {outputMsg "  OK A  [$aval Type] [$aval P21ID]" magenta}
                      lappend invVals($invRule) "[$aval Type] [$aval P21ID]"
                      if {[string first "DELETETHIS" $msg] == -1} {append msg " [$aval Type] DELETETHIS$objType"}
                      set msgok 1
                    }
                  }
                } emsg1]} {

# still needs some work to remove the catch                
                  set stat "related [$invAttr NodeType] B [lindex $inverse 0] [lindex $inverse 1]"
                  foreach aval [$invAttr Value] {
                    catch {
                      if {$DEBUGINV} {outputMsg "  [$aval Type] [$aval P21ID] / $objType $objP21ID / [$invEnt Type] [$invEnt P21ID]" magenta}
                      if {[$aval P21ID] != $objP21ID && [$aval Type] != $objType} {
                        if {$DEBUGINV} {outputMsg "  OK B" magenta}
                        lappend invVals($invRule) "[$aval Type] [$aval P21ID]"
                        if {[string first "DELETETHIS" $msg] == -1} {append msg " [$aval Type] DELETETHIS$objType"}
                        set msgok 1
                      }
                    }
                  }
                }
              }
            }

# look for other GetUsedIn relationships that are not relating/related Inverses 
          } elseif {$attrName == [lindex $inverse 1]} {
            set attrValue [$invAttr Value]
            if {$DEBUGINV} {outputMsg "    attrValue(L=[llength $attrValue])  ($attrValue)  NodeType=[$invAttr NodeType]"}
            if {[string first "handle" $attrValue] != -1} {

# nodetype = 20
              if {[$invAttr NodeType] == 20} {
                if {[catch {
                  set stat "used in [$invAttr NodeType] A [lindex $inverse 0] [lindex $inverse 1]"
                  ::tcom::foreach val $attrValue {
                    if {$DEBUGINV} {outputMsg "       $val [$val Type] [$val P21ID] / $objType $objP21ID / [$invEnt Type] [$invEnt P21ID]"}
                    if {[$val P21ID] == $objP21ID} {
                      #lappend invVals($invRule) "[$invEnt Type].[lindex $inverse 1] [$invEnt P21ID]"
                      set str "[formatComplexEnt [$invEnt Type]].[lindex $inverse 1] [$invEnt P21ID]"
                      if {![info exists invVals($invRule)]} {
                        lappend invVals($invRule) $str
                      } elseif {[lsearch $invVals($invRule) $str] == -1} {
                        lappend invVals($invRule) $str
                      }
                      if {[string first "DELETETHIS" $msg] == -1} {append msg " [lindex $inverse 1] DELETETHIS$objType"}
                      set msgok 1
                    }
                  }
                } emsg]} {
                  set stat "used in [$invAttr NodeType] B [lindex $inverse 0] [lindex $inverse 1]"
                  if {[catch {
                    if {$DEBUGINV} {outputMsg "       $val / $objType $objP21ID / [$invEnt Type] [$invEnt P21ID]"}
                    #lappend invVals($invRule) "[$invEnt Type].[lindex $inverse 1] [$invEnt P21ID]"
                    set str "[$invEnt Type].[lindex $inverse 1] [$invEnt P21ID]"
                    if {![info exists invVals($invRule)]} {
                      lappend invVals($invRule) $str
                    } elseif {[lsearch $invVals($invRule) $str] == -1} {
                      lappend invVals($invRule) $str
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
                  set stat "used in [$invAttr NodeType] C [lindex $inverse 0] [lindex $inverse 1]"
                  set subType [$attrValue Type]
                  if {$DEBUGINV} {outputMsg "     $subType [$attrValue P21ID] / $objType $objP21ID / [$invEnt Type] [$invEnt P21ID]"}
  
                  if {[$attrValue P21ID] == $objP21ID} {
                    if {$DEBUGINV} {outputMsg "     $invType  $attrValue  $subType"}
                    set str "[$invEnt Type].[lindex $inverse 1] [$invEnt P21ID]"
                    if {![info exists invVals($invRule)]} {
                      lappend invVals($invRule) $str
                    } elseif {[lsearch $invVals($invRule) $str] == -1} {
                      lappend invVals($invRule) $str
                    }
                    if {[string first "DELETETHIS" $msg] == -1} {append msg " [lindex $inverse 1] DELETETHIS$objType"}
                    set msgok 1
                  }

# special case for a nodetype 18 or 19 that is more like 20 (a list)
                } emsg2]} {
                  set stat "used in [$invAttr NodeType] D [lindex $inverse 0] [lindex $inverse 1]"
                  ::tcom::foreach aval [$invAttr Value] {
                    #if {$DEBUGINV} {outputMsg "  $aval / [$aval Type] [$aval P21ID] / $objType $objP21ID / [$invEnt Type] [$invEnt P21ID]" magenta}
                    if {[$aval P21ID] != $objP21ID && [$aval Type] != $objType} {
                      #if {$DEBUGINV} {outputMsg "  OK A  [$aval Type] [$aval P21ID]" magenta}
                      lappend invVals($invRule) "[$aval Type] [$aval P21ID]"
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
  
        if {$msgok && [string first $msg $invMsg] == -1} {
          append invMsg $msg
          errorMsg $msg blue
        }
      }
    }
  } emsg]} {
    if {!$developer || $stat == ""} {
      errorMsg "ERROR processing Inverse for '[$objEntity Type]': $emsg"
    } else {
      errorMsg "ERROR processing Inverse for '[$objEntity Type]': $emsg\n ($stat)"
    }
  }
}

# -------------------------------------------------------------------------------
# report inverses 
proc invReport {} {
  global invVals cellval cells thisEntType row col invCol
  #outputMsg invReport red
  
# inverse values and heading
  foreach item [array names invVals] {
    catch {foreach idx [array names cellval] {unset cellval($idx)}}

    foreach val $invVals($item) {
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
    if {[info exists invCol($idx)]} {
      $cells($thisEntType) Item $row($thisEntType) $invCol($idx) [string trim $str]
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
      set invCol($idx) $col($thisEntType)
    }
  }
}

# -------------------------------------------------------------------------------
# set column color, border, group for INVERSES and Used In
proc invFormat {rancol} {
  global thisEntType col cells row rowmax worksheet invGroup excelVersion
  
  set igrp1 100
  set igrp2 0
  set i1 [expr {$rancol+1}]
    
# fix column widths
  for {set i 1} {$i <= $i1} {incr i} {
    set val [[$cells($thisEntType) Item 3 $i] Value]
    if {$val == "Used In" || [string first "INV-" $val] != -1} {
      set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
      $range ColumnWidth [expr 80]
    }
  }
  [$worksheet($thisEntType) Columns] AutoFit
  [$worksheet($thisEntType) Rows] AutoFit

# set colors, borders
  for {set i 1} {$i <= $i1} {incr i} {
    set val [[$cells($thisEntType) Item 3 $i] Value]
    if {[string first "INV-" $val] != -1 || [string first "Used In" $val] != -1} {
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
      if {$excelVersion >= 12} {
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
  global opt entCategory userEntityList stepAP ap203all ap214all ap242all
  
  set checkInv 0
  
# other APs  
  if {[string first "AP203" $stepAP] == -1 && $stepAP != "AP214" && $stepAP != "AP242"} {
    if {[lsearch $ap203all $entType] == -1 && \
        [lsearch $ap214all $entType] == -1 && \
        [lsearch $ap242all $entType] == -1} {
      set checkInv 1
    }

# tolerance, user-defined entities
  } elseif {($opt(PR_STEP_TOLR)  && [lsearch $entCategory(PR_STEP_TOLR)  $entType] != -1) || \
      ($opt(PR_USER) && [lsearch $userEntityList $entType] != -1)} {
    set checkInv 1
    #outputMsg "checkInv $entType" blue

# other types of entities (should be more selective to make it faster)
  } elseif { \
      [string first "additive"      $entType] != -1 || \
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
      [string first "kinematic"     $entType] != -1 || \
      [string first "machining"     $entType] != -1 || \
      [string first "milling"       $entType] != -1 || \
      [string first "process"       $entType] != -1 || \
      [string first "product"       $entType] != -1 || \
      [string first "representation" $entType] != -1 || \
      [string first "resource"      $entType] != -1 || \
      [string first "shape"         $entType] != -1 || \
      [string first "symmetry"      $entType] != -1 || \
      [string first "draughting_model"    $entType] != -1 || \
      [string first "draughting_callout"  $entType] != -1 || \
      [string first "instanced_feature"   $entType] != -1 || \
      [string first "mechanical_design"   $entType] != -1 || \
      [string first "next_assembly"       $entType] != -1 || \
      [string first "property_definition" $entType] != -1} {

# not these specific entities
    if {$entType != "annotation_fill_area" && \
        $entType != "dimensional_characteristic_representation" && \
        $entType != "draughting_model_item_association" && \
        $entType != "draughting_pre_defined_colour" && \
        $entType != "draughting_pre_defined_curve_font" && \
        $entType != "geometric_item_specific_usage" && \
        $entType != "representation_item" && \
        $entType != "product_definition_relationship" && \
        $entType != "property_definition_representation" && \
        $entType != "shape_aspect_deriving_relationship" && \
        $entType != "shape_aspect_relationship"} {
      set checkInv 1
      #outputMsg "checkInv $entType" green
    }
  }
  return $checkInv
}

# -------------------------------------------------------------------------------
# set inverse relationships
proc initDataInverses {} {
  global tolNames inverses
  
# when adding used_in or inverse (related, relating) relationships also add entities to check in invSetCheck above  
  set inverses {}
  lappend inverses [list annotation_curve_occurrence item used_in]
  lappend inverses [list annotation_occurrence item used_in]
  lappend inverses [list annotation_occurrence_relationship related_annotation_occurrence relating_annotation_occurrence]
  lappend inverses [list annotation_occurrence_relationship relating_annotation_occurrence related_annotation_occurrence]
  lappend inverses [list annotation_placeholder_occurrence item used_in]
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
  lappend inverses [list draughting_model_item_association_with_placeholder definition used_in]
  lappend inverses [list draughting_model_item_association_with_placeholder identified_item used_in]
  lappend inverses [list draughting_model_item_association_with_placeholder used_representation used_in]

  lappend inverses [list geometric_item_specific_usage definition used_in]
  lappend inverses [list geometric_item_specific_usage identified_item used_in]
  lappend inverses [list geometric_item_specific_usage used_representation used_in]
  lappend inverses [list geometric_tolerance_relationship related_geometric_tolerance relating_geometric_tolerance]
  lappend inverses [list geometric_tolerance_relationship relating_geometric_tolerance related_geometric_tolerance]

  lappend inverses [list id_attribute identified_item used_in]
  lappend inverses [list leader_directed_callout contents used_in]

  lappend inverses [list make_from_usage_option related_product_definition relating_product_definition]
  lappend inverses [list make_from_usage_option relating_product_definition related_product_definition]
  lappend inverses [list multi_level_reference_designator location used_in]
  lappend inverses [list multi_level_reference_designator related_product_definition relating_product_definition]
  lappend inverses [list multi_level_reference_designator relating_product_definition related_product_definition]

  lappend inverses [list name_attribute named_item used_in]
  lappend inverses [list next_assembly_usage_occurrence related_product_definition relating_product_definition]
  lappend inverses [list next_assembly_usage_occurrence relating_product_definition related_product_definition]

  lappend inverses [list process_plan chosen_method used_in]
  lappend inverses [list process_product_association defined_product used_in]
  lappend inverses [list process_product_association process used_in]

  lappend inverses [list product frame_of_reference used_in]
  lappend inverses [list product_context frame_of_reference used_in]
  lappend inverses [list product_category_relationship category used_in]
  lappend inverses [list product_category_relationship sub_category used_in]
  lappend inverses [list product_definition formation used_in]
  lappend inverses [list product_definition frame_of_reference used_in]
  lappend inverses [list product_definition_context frame_of_reference used_in]
  lappend inverses [list product_definition_formation of_product used_in]
  lappend inverses [list product_definition_process chosen_method used_in]
  lappend inverses [list product_definition_shape definition used_in]
  lappend inverses [list product_related_product_category products used_in]

  lappend inverses [list property_definition definition used_in]
  lappend inverses [list referenced_modified_datum referenced_datum used_in]
  lappend inverses [list representation_relationship rep_1 rep_2]
  lappend inverses [list representation_relationship rep_2 rep_1]
  lappend inverses [list representation_relationship_with_transformation rep_1 rep_2]
  lappend inverses [list representation_relationship_with_transformation rep_2 rep_1]
  lappend inverses [list representation_relationship_with_transformation_and_shape_representation_relationship rep_1 rep_2]
  lappend inverses [list representation_relationship_with_transformation_and_shape_representation_relationship rep_2 rep_1]

  lappend inverses [list shape_aspect of_shape used_in]
  lappend inverses [list shape_aspect_relationship related_shape_aspect relating_shape_aspect]
  lappend inverses [list shape_aspect_relationship relating_shape_aspect related_shape_aspect]
  lappend inverses [list shape_definition_representation definition used_in]
  lappend inverses [list shape_definition_representation used_representation used_in]
  lappend inverses [list shape_representation context_of_items used_in]

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

  foreach tol $tolNames {
    lappend inverses [list $tol toleranced_shape_aspect used_in]
    lappend inverses [list $tol datum_system used_in]
  }
  lappend inverses [list geometric_tolerance_with_datum_reference datum_system used_in]
  
# AP209
  lappend inverses [list control_linear_static_analysis_step analysis_control used_in]
  lappend inverses [list control_linear_static_analysis_step initial_state used_in]
  lappend inverses [list control_linear_static_analysis_step process used_in]
  lappend inverses [list control_linear_static_load_increment_process final_input_state used_in]
  lappend inverses [list control_result_relationship control result]
  lappend inverses [list control_result_relationship result control]

  lappend inverses [list curve_3d_element_property interval_definitions used_in]
  lappend inverses [list curve_3d_element_property end_offsets used_in]
  lappend inverses [list curve_3d_element_property end_releases used_in]
  #lappend inverses [list curve_3d_element_representation items used_in]
  #lappend inverses [list curve_3d_element_representation context_of_items used_in]
  lappend inverses [list curve_3d_element_representation node_list used_in]
  lappend inverses [list curve_element_interval_constant finish_position used_in]
  lappend inverses [list curve_element_interval_constant eu_angles used_in]
  lappend inverses [list curve_element_interval_constant section used_in]
  lappend inverses [list curve_element_location coordinate used_in]

  lappend inverses [list element_geometric_relationship element_ref item]
  lappend inverses [list element_geometric_relationship item element_ref]
  lappend inverses [list element_group model_ref used_in]
  lappend inverses [list element_group elements used_in]

  lappend inverses [list fea_material_property_representation definition used_in]
  lappend inverses [list fea_material_property_representation used_representation used_in]
  #lappend inverses [list fea_material_property_representation dependent_environment used_in]
  lappend inverses [list fea_model_3d items used_in]
  lappend inverses [list fea_model_definition of_shape used_in]

  lappend inverses [list nodal_freedom_action_definition defined_state used_in]
  lappend inverses [list nodal_freedom_action_definition node used_in]
  lappend inverses [list nodal_freedom_action_definition coordinate_system used_in]
  lappend inverses [list nodal_freedom_action_definition degrees_of_freedom used_in]
  lappend inverses [list nodal_freedom_values defined_state used_in]
  lappend inverses [list nodal_freedom_values degrees_of_freedom used_in]
  lappend inverses [list nodal_freedom_values coordinate_system used_in]
  lappend inverses [list nodal_freedom_values node used_in]
  lappend inverses [list node_geometric_relationship node_ref used_in]
  lappend inverses [list node_group nodes used_in]
  lappend inverses [list node_group model_ref used_in]
  lappend inverses [list node_set nodes used_in]

  lappend inverses [list output_request_state steps used_in]
  lappend inverses [list point_representation items used_in]
  lappend inverses [list product_definition_formation_relationship relating_product_definition_formation related_product_definition_formation]
  lappend inverses [list product_definition_formation_relationship related_product_definition_formation relating_product_definition_formation]
  lappend inverses [list result_linear_modes_and_frequencies_analysis_sub_step control used_in]
  lappend inverses [list result_linear_modes_and_frequencies_analysis_sub_step result used_in]
  lappend inverses [list result_linear_modes_and_frequencies_analysis_sub_step states used_in]
  lappend inverses [list result_linear_static_analysis_sub_step analysis_control used_in]
  lappend inverses [list result_linear_static_analysis_sub_step analysis_result used_in]
  lappend inverses [list result_linear_static_analysis_sub_step state used_in]

  lappend inverses [list single_point_constraint_element steps used_in]
  lappend inverses [list single_point_constraint_element required_node used_in]
  lappend inverses [list single_point_constraint_element freedoms_and_values used_in]
  lappend inverses [list single_point_constraint_element_values defined_state used_in]
  lappend inverses [list single_point_constraint_element_values element used_in]
  lappend inverses [list single_point_constraint_element_values degrees_of_freedom used_in]

  lappend inverses [list state_component state used_in]
  lappend inverses [list state_relationship related_state relating_state]
  lappend inverses [list state_relationship relating_state related_state]

  lappend inverses [list surface_3d_element_location_point_volume_variable_values defined_state used_in]
  lappend inverses [list surface_3d_element_location_point_volume_variable_values element used_in]
  lappend inverses [list surface_3d_element_location_point_volume_variable_values values_and_locations used_in]
  lappend inverses [list surface_3d_element_representation node_list used_in]
  lappend inverses [list surface_3d_element_location_point_volume_variable_values values_and_locations used_in]
  lappend inverses [list surface_element_location coordinates used_in]

  lappend inverses [list volume_3d_element_representation node_list used_in]
  lappend inverses [list volume_3d_element_location_point_volume_variable_values values_and_locations used_in]
  lappend inverses [list whole_model_modes_and_frequencies_analysis_message defined_state used_in]

# kinematics  
  lappend inverses [list kinematic_joint edge_end used_in]
  lappend inverses [list kinematic_joint edge_start used_in]
  lappend inverses [list rigid_link_representation represented_link used_in]
  
# AP214 (AP238, AP242)
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

  lappend inverses [list roundness_definition projection_end used_in]
  
# AP238
  lappend inverses [list feature_component_relationship relating_shape_aspect related_shape_aspect]
  lappend inverses [list feature_component_relationship related_shape_aspect relating_shape_aspect]
  lappend inverses [list shape_defining_relationship relating_shape_aspect related_shape_aspect]
  lappend inverses [list shape_defining_relationship related_shape_aspect relating_shape_aspect]
  lappend inverses [list machining_feature_relationship relating_method related_method]
  lappend inverses [list machining_feature_relationship related_method relating_method]
  lappend inverses [list machining_functions_relationship relating_method related_method]
  lappend inverses [list machining_functions_relationship related_method relating_method]
  lappend inverses [list machining_operation_relationship relating_method related_method]
  lappend inverses [list machining_operation_relationship related_method relating_method]
  lappend inverses [list machining_process_branch_relationship relating_method related_method]
  lappend inverses [list machining_process_branch_relationship related_method relating_method]
  lappend inverses [list machining_process_sequence_relationship relating_method related_method]
  lappend inverses [list machining_process_sequence_relationship related_method relating_method]
  lappend inverses [list machining_technology_relationship relating_method related_method]
  lappend inverses [list machining_technology_relationship related_method relating_method]
  lappend inverses [list machining_toolpath_sequence_relationship relating_method related_method]
  lappend inverses [list machining_toolpath_sequence_relationship related_method relating_method]

# new AM entities
  lappend inverses [list additive_manufacturing_build_plate_relationship relating_product_definition related_product_definition]
  lappend inverses [list additive_manufacturing_build_plate_relationship related_product_definition relating_product_definition]
  lappend inverses [list additive_manufacturing_setup_support_relationship relating_product_definition related_product_definition]
  lappend inverses [list additive_manufacturing_setup_support_relationship related_product_definition relating_product_definition]
  lappend inverses [list additive_manufacturing_setup_workpiece_relationship relating_product_definition related_product_definition]
  lappend inverses [list additive_manufacturing_setup_workpiece_relationship related_product_definition relating_product_definition]
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
