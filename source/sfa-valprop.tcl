proc valPropStart {} {
  global objDesign
  global cells col entLevel ent entAttrList ncartpt opt pd pdcol pdheading propDefRow valPropLink valPropNames rowmax
  
# CAx-IF RP Geometric and Assembly Validation Properties, section 8
  set valPropNames(geometric_validation_property) [list \
    [list "bounding box" [list "bounding box corner point"]] \
    [list centroid [list "centre point"]] \
    [list "independent curve centroid" [list "curve centre point"]] \
    [list "independent curve length" [list "curve length measure"]] \
    [list "independent points centroid" [list "independent points centre point"]] \
    [list "independent surface area" [list "independent surface area measure"]] \
    [list "independent surface centroid" [list "surface centre point"]] \
    [list "number of independent points" [list "number of independent points"]] \
    [list "sharp sampling points" [list "sampling point"]] \
    [list "smooth sampling points" [list "sampling point"]] \
    [list "surface area" [list "surface area measure" "wetted area measure"]] \
    [list volume [list "volume measure"]]]

# CAx-IF RP Geometric and Assembly Validation Properties, section 8
  set valPropNames(assembly_validation_property) [list \
    [list "notional solids centroid" [list "centre point"]] \
    [list "number of children" [list "number of children"]]]

# includes tessellated pmi and semantic pmi valprops, section 10
  set valPropNames(pmi_validation_property) [list \
    [list "" [list "affected area" "affected curve length" "datum references" "equivalent unicode string" "font name" \
      "number of annotations" "number of composite tolerances" "number of datum features" "number of datum references" \
      "number of datum targets" "number of dimensional locations" "number of dimensional sizes" "number of facets" \
      "number of geometric tolerances" "number of PMI presentation elements" "number of segments" "number of semantic pmi elements" \
      "number of views" "polyline centre point" "polyline curve length" "tessellated curve centre point" \
      "tessellated curve length" "tessellated surface area" "tessellated surface centre point"]]]

# CAx-IF RP User Defined Attributes, section 8
  set valPropNames(attribute_validation_property) [list \
    [list "" [list "boolean user attributes" "edge user attributes" "face user attributes" "group user attributes" \
      "instance user attributes" "integer user attributes" "measure value user attributes" "part user attributes" \
      "real user attributes" "solid user attributes" "text user attributes" "user attribute groups" "vertex user attributes"]]]

# tessellated geometry recommended practice
  set valPropNames(tessellated_validation_property) [list \
    [list "bounding box" [list "bounding box corner point"]] \
    [list centroid [list "tessellated surface centre point" "tessellated curve centre point" "tessellated point set centre point"]] \
    [list "curve length" [list "tessellated curve length"]] \
    [list "number of facets" [list "number of facets"]] \
    [list "number of segments" [list "number of segments"]] \
    [list "surface area" [list "tessellated surface area"]]]

# composite recommended practice (new vp in the last line)
  set valPropNames(composite_validation_property) [list \
    [list "" [list "number of composite tables" "number of composite materials per part" "number of composite materials per part" \
      "number of orientations per part" "number of plies per part" "number of plies per laminate table" \
      "number of composite sequences per laminate table" "number of composite materials per laminate table" \
      "number of composite orientations per laminate table" "ordered sequences per laminate table" \
      "notational centroid" "number of ply pieces per ply" \
      "number of tables" "number of sequences" "number of plies" "number of materials" "number of orientations" "sum of all ply surfaces areas" "centre point of all plies"]]]

  set derived_unit_element [list derived_unit_element unit \
    [list conversion_based_unit_and_length_unit dimensions name conversion_factor] \
    [list conversion_based_unit_and_mass_unit dimensions name conversion_factor] \
    [list length_unit_and_si_unit prefix name] exponent]
  set cartesian_point [list cartesian_point name coordinates]
  set a2p3d [list axis2_placement_3d name location $cartesian_point axis [list direction name direction_ratios] ref_direction [list direction name direction_ratios]]

  set drep [list descriptive_representation_item name description]
  set vrep [list value_representation_item name value_component]
  set brep [list boolean_representation_item name the_value]
  set irep [list integer_representation_item name the_value]
  set rrep [list real_representation_item name the_value]
  set mrep [list measure_representation_item name value_component unit_component \
    [list derived_unit elements $derived_unit_element] \
    [list area_unit elements $derived_unit_element] \
    [list volume_unit elements $derived_unit_element] \
    [list mass_unit_and_si_unit prefix name] \
    [list si_unit_and_thermodynamic_temperature_unit dimensions prefix name]]

  set len1 [list length_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set len2 [list length_measure_with_unit_and_measure_representation_item_and_qualified_representation_item value_component unit_component name qualifiers]
  set mass1 [list mass_measure_with_unit_and_measure_representation_item value_component unit_component name]

  set rep1 [list representation name items $a2p3d $drep $vrep $brep $irep $rrep $mrep $len1 $len2 $mass1 $cartesian_point]
  set rep2 [list shape_representation_with_parameters name items $a2p3d $drep $vrep $brep $irep $rrep $mrep $len1 $len2 $mass1 $cartesian_point]

  set gvp [list property_definition_representation \
    definition [list property_definition name description definition] \
    used_representation $rep1 $rep2]

  set valPropLink 0
  set entAttrList {}
  set pdcol 0
  set propDefRow {}
  set pd "property_definition"

  if {[info exists pdheading]} {unset pdheading}
  if {[info exists ent]}       {unset ent}

  outputMsg " Adding Properties to property_definition worksheet" blue

  if {$opt(DEBUG1)} {outputMsg \n}
  set entLevel 0
  setEntAttrList $gvp
  if {$opt(DEBUG1)} {outputMsg "entAttrList $entAttrList"}
  if {$opt(DEBUG1)} {outputMsg \n}
  unset ent
    
  set startent [lindex $gvp 0]
  set n 0
  set entLevel 0
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    set ncartpt 0
    if {[$objEntity Type] == $startent} {
      if {$n < $rowmax} {
        if {[expr {$n%2000}] == 0} {
          if {$n > 0} {outputMsg "  $n"}
          update
        }
        valPropReport $objEntity
        if {$opt(DEBUG1)} {outputMsg \n}
      }
      incr n
    }
  }
  set col($pd) $pdcol
}

# -------------------------------------------------------------------------------
proc valPropReport {objEntity} {
  global cells col entLevel ent entAttrList maxrep ncartpt nrep opt pd pdcol pdheading pmivalprop pointLimit prefix 
  global propDefID propDefIDRow propDefName propDefOK propDefRow recPracNames repName stepAP syntaxErr valName valPropLink valPropNames

  if {[info exists propDefOK]} {if {$propDefOK == 0} {return}}

  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]

  if {[string first "handle" $objEntity] == -1} {
    #if {$objEntity != ""} {outputMsg "$ind $objEntity"}
  } else {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set ent($entLevel) [$objEntity Type]
    set objAttributes [$objEntity Attributes]

    if {$opt(DEBUG1)} {outputMsg "$ind ENT $entLevel #$objID=$objType (ATR=[$objAttributes Count])" blue}

# limit sampling points to pointLimit (2)
    if {[info exists repName]} {
      if {$objType == "cartesian_point" && [string first "sampling points" $repName] != -1} {
        incr ncartpt
        if {$ncartpt > $pointLimit} {
          errorMsg " Only the first $pointLimit sampling points are reported" red
          incr entLevel -1
          return  
        }
      }
    }

    if {$objType == "property_definition"} {
      set propDefID $objID
      if {![info exists propDefIDRow($propDefID)]} {
        incr entLevel -1
        set propDefOK 0
        return
      } else {
        set propDefOK 1
      }
    }
    
    if {$entLevel == 1} {set pmivalprop 0}

    ::tcom::foreach objAttribute $objAttributes {
      set objName  [$objAttribute Name]
      set objValue [$objAttribute Value]
      set objNodeType [$objAttribute NodeType]
      set objSize [$objAttribute Size]
      set objAttrType [$objAttribute Type]

      set ent1 "$ent($entLevel) $objName"
      set ent2 "$ent($entLevel).$objName"
      set idx [lsearch $entAttrList $ent1]
      set invalid ""

# -----------------
# nodeType = 18,19
      if {$objNodeType == 18 || $objNodeType == 19} {
        if {$idx != -1} {
          if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}
          
          if {[string length $objValue] == 0 && $objName == "unit_component" && \
              ([string first "volume" $valName] == -1 || [string first "area" $valName] == -1 || [string first "length" $valName] == -1)} {
            set msg "Syntax Error: Missing or invalid '$objName' attribute on $ent($entLevel).\n[string repeat " " 14]Units will not be reported for a length, area, or volume validation property.\n[string repeat " " 14]\($recPracNames(valprop))"
            errorMsg $msg
            lappend syntaxErr($ent($entLevel)) [list $objID unit_component $msg]
          }

          if {[info exists cells($pd)]} {
            set ok 0

# get values for these entity and attribute pairs
            switch -glob $ent1 {
              "*measure_representation_item* value_component" {
                set ok 1
                set col($pd) 9
                if {$objValue <= 0.} {
                  if {[string first "length" $valName] != -1 || [string first "area" $valName] != -1 || \
                      [string first "volume" $valName] != -1 || [string first "number of" $valName] != -1} {
                    errorMsg " Validation property '$valName' = $objValue"
                  }
                }
              }
              "value_representation_item value_component"     -
              "descriptive_representation_item description"   {set ok 1; set col($pd) 9}
              "property_definition definition" {
                if {[string first "validation_property" $propDefName] != -1} {
                  if {[string length $objValue] == 0} {
                    set msg "Syntax Error: Missing 'definition' attribute on property_definition.\n[string repeat " " 14]\($recPracNames(valprop), Sec. 4)"
                    errorMsg $msg
                    lappend syntaxErr(property_definition) [list $propDefIDRow($propDefID) 4 $msg]
                  }
                }
              }
            }
            set colName "value"

            if {$ok && [info exists propDefID]} {
              set c [string index [cellRange 1 $col($pd)] 0]
              set r $propDefIDRow($propDefID)

# colName
              if {![info exists pdheading($col($pd))]} {
                $cells($pd) Item 3 $c $colName
                $cells($pd) Item 3 [string index [cellRange 1 [expr {$col($pd)+1}]] 0] "attribute"
                set pdheading($col($pd)) 1
              }

# keep track of rows with validation properties
              if {[lsearch $propDefRow $r] == -1 && [string first "validation_property" $propDefName] != -1} {lappend propDefRow $r}

# value in spreadsheet
              set val [[$cells($pd) Item $r $c] Value]
              if {$val == ""} {
                $cells($pd) Item $r $c $objValue
              } else {
                $cells($pd) Item $r $c "$val[format "%c" 10]$objValue"
              }

# entity reference in spreadsheet
              incr col($pd)
              set c [string index [cellRange 1 $col($pd)] 0]
              set val [[$cells($pd) Item $r $c] Value]
              if {$val == ""} {
                $cells($pd) Item $r $c "#$objID $ent2"
              } else {
                $cells($pd) Item $r $c "$val[format "%c" 10]#$objID $ent2"
              }

# keep track of max column
              set pdcol [expr {max($col($pd),$pdcol)}]
            }
          }

# if referred to another, get the entity
          if {[string first "handle" $objEntity] != -1} {valPropReport $objValue}
        }

# --------------
# nodeType = 20
      } elseif {$objNodeType == 20} {
        if {$idx != -1} {
          if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}

          if {[info exists cells($pd)]} {
            set ok 0

# get values for these entity and attribute pairs
# nrep keeps track of multiple representation items
            switch -glob $ent1 {
              "cartesian_point coordinates" -
              "direction direction_ratios"  {set ok 1; set col($pd) 9; set colName "value"}
              
              "representation items" -
              "shape_representation_with_parameters items" {set nrep 0; set maxrep $objSize}
            }

# value in spreadsheet
            if {$ok && [info exists propDefID]} {
              set c [string index [cellRange 1 $col($pd)] 0]
              set r $propDefIDRow($propDefID)

# colName
              if {![info exists pdheading($col($pd))]} {
                $cells($pd) Item 3 $c $colName
                $cells($pd) Item 3 [string index [cellRange 1 [expr {$col($pd)+1}]] 0] "attribute"
                set pdheading($col($pd)) 1
              }

# keep track of rows with validation properties
              if {[lsearch $propDefRow $r] == -1 && [string first "validation_property" $propDefName] != -1} {lappend propDefRow $r}

              set ov $objValue
              if {$ent1 == "cartesian_point coordinates" || $ent1 == "direction direction_ratios"} {
                regsub -all " " $ov "    " ov
              }

              set val [[$cells($pd) Item $r $c] Value]
              if {$val == ""} {
                $cells($pd) Item $r $c $ov
              } else {
                if {[catch {
                  $cells($pd) Item $r $c "$val[format "%c" 10]$ov"
                } emsg]} {
                  errorMsg "  ERROR: Too much data to show in a cell" red
                }
              }

# entity reference in spreadsheet (cartesian point or direction)
              incr col($pd)
              set c [string index [cellRange 1 $col($pd)] 0]
              set val [[$cells($pd) Item $r $c] Value]
              if {$val == ""} {
                $cells($pd) Item $r $c "#$objID $ent2"
              } else {
                $cells($pd) Item $r $c "$val[format "%c" 10]#$objID $ent2"
              }
              set pdcol [expr {max($col($pd),$pdcol)}]
              
# add blank columns for units and exponent, if more than one representation
              if {[info exists maxrep]} {
                if {$maxrep > 1} {
                  for {set i 0} {$i < 4} {incr i} {
                    incr col($pd)
                    set c [string index [cellRange 1 $col($pd)] 0]
                    set val [[$cells($pd) Item $r $c] Value]
                    if {$val == ""} {
                      $cells($pd) Item $r $c " "
                    } else {
                      $cells($pd) Item $r $c "$val[format "%c" 10] "
                    }
                    set pdcol [expr {max($col($pd),$pdcol)}]
                  }             
                }
              } else {
                errorMsg "maxrep does not exist"
              }
            }
          }

# write sampling points
          if {$objName == "items" && [info exists r]} {
            set val [[$cells($pd) Item $r 5] Value]
            if {[string first " sampling points" $val] != -1} {
              $cells($pd) Item $r 5 "$val ([expr {min($pointLimit,$objSize)}] of $objSize)"
            }
          }

# get the entities that are referred to, but only up to pointLimit cartesian points for sampling points
          if {$ncartpt < $pointLimit} {
            if {[catch {
              ::tcom::foreach val1 $objValue {valPropReport $val1}
            } emsg]} {
              foreach val2 $objValue {valPropReport $val2}
            }
          }
        }

# ---------------------
# nodeType != 18,19,20
      } else {
        if {$idx != -1} {
          if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objAttrType)"}

          if {[info exists cells($pd)]} {
            set ok 0
            set colName ""
            set invalid ""

# get values for these entity and attribute pairs
            switch -glob $ent1 {
              "*_representation_item the_value" {
                set ok 1
                set col($pd) 9
                set colName "value"
                if {$objValue <= 0.} {
                  if {[string first "length" $valName] != -1 || [string first "area" $valName] != -1 || \
                      [string first "volume" $valName] != -1 || [string first "number of" $valName] != -1} {
                    errorMsg " Validation property '$valName' = $objValue"
                  }
                }
              }
              
              "descriptive_representation_item description" {set ok 1; set col($pd) 9; set colName "value"}

              "conversion_based_unit_and_*_unit name" {set ok 1; set col($pd) 11; set colName "units"}

              "*_unit_and_si_unit name" -
              "si_unit_and_*_unit name" {set ok 1; set col($pd) 11; set colName "units"; set objValue "$prefix$objValue"}

              "derived_unit_element exponent" {set ok 1; set col($pd) 13; set colName "exponent"}

              "*_unit_and_si_unit prefix" -
              "si_unit_and_*_unit prefix" {set ok 0; set prefix $objValue}

              "property_definition name" {
                set ok 0
                set pmivalprop 1
                regsub -all " " $objValue "_" propDefName
                
                if {[string first "validation property" $objValue] != -1} {
                  if {[string first "geometric" $objValue] != 0 && [string first "assembly" $objValue] != 0 && \
                      [string first "pmi" $objValue] != 0 && [string first "tessellated" $objValue] != 0 && \
                      [string first "attribute" $objValue] != 0} {
                    set msg "Possible Syntax Error: Unexpected Validation Property '$objValue'"
                    errorMsg $msg
                    set invalid $msg
                    lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                  }
                  set valPropLink 1
                }
              }

              "representation name" {
                set ok 1
                set col($pd) 5
                set colName "representation name"

                set repName $objValue
                if {[string first "sampling points" $repName] != -1} {set ncardpt 0}

                if {[info exists propDefName]} {
                  if {$entLevel == 2 && [info exists valPropNames($propDefName)]} {
                    set ok1 0

# look for valid representation.name in valPropNames
# new RP allows for blank representation.name (repName) except for sampling points
                    if {[string trim $repName] != ""} {
                      if {$repName != ""} {
                        foreach idx $valPropNames($propDefName) {
                          if {[lindex $idx 0] == $repName || [lindex $idx 0] == ""} {
                            set ok1 1
                            break
                          }
                        }
                      }
                    } else {
                      set ok1 1
                    }
  
                    if {!$ok1} {
                      set emsg "Syntax Error: Invalid '$ent2' attribute ($repName) for '$propDefName'.\n              "
                      if {$propDefName == "geometric_validation_property" || $propDefName == "assembly_validation_property"} {
                        append emsg "($recPracNames(valprop), Sec. 8, Table 1)"
                      } elseif {$propDefName == "pmi_validation_property"} {
                        if {$stepAP == "AP242"} {
                          append emsg "($recPracNames(pmi242), Sec. 10)"
                        } else {
                          append emsg "($recPracNames(pmi203), Sec. 6)"
                        }
                      } elseif {$propDefName == "tessellated_validation_property"} {
                        append emsg "($recPracNames(tessgeom), Sec. 8.4)"
                      } elseif {$propDefName == "attribute_validation_property"} {
                        append emsg "($recPracNames(uda), Sec. 8)"
                      }
                      errorMsg $emsg
                      set invalid $emsg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $emsg]
                    }
                  }
                }
              }

              "*_representation_item name"       -
              "*_representation_item_and_* name" -
              "cartesian_point name"             -
              "direction name" {
                set ok 1
                set col($pd) 7
                set colName "value name"
                set valName $objValue
                if {[info exists nrep]} {incr nrep}

# new RP allows for blank representation.name (repName) except for sampling points
                if {[info exists propDefName]} {
                  if {$entLevel == 3 && [info exists valPropNames($propDefName)]} {
                    set ok1 0
                    foreach idx $valPropNames($propDefName) {
                      if {[lindex $idx 0] == $repName || [lindex $idx 0] == "" || [string trim $repName] == ""} {
                        foreach item [lindex $idx 1] {
                          set repItemName $item
                          if {$objValue == $repItemName} {
                            set ok1 1
                            if {$objValue == "sampling point" && [string trim $repName] == ""} {
                              set emsg "Syntax Error: Invalid blank representation.name for '$objValue'.\n              "
                              append emsg "($recPracNames(valprop), Sec. 4.11)"
                              errorMsg $emsg
                            }
                            break

# check if wrong case used
                          } elseif {[string tolower $objValue] == $repItemName} {
                            set ok1 2
                            break
                          }
                        }
                      }
                    }

# do not flag cartesian_point.name errors with entity ids           
                    if {!$ok1 && $ent2 == "cartesian_point.name"} {
                      if {[string first "\#" $objValue] != -1} {set ok1 1}
                    }
                    
                    if {$ok1 != 1} {
                      if {$ok1 == 0} {
                        set emsg "Syntax Error: Invalid "
                      } elseif {$ok1 == 2} {
                        set emsg "Syntax Error: Use lower case for "
                      }
                      append emsg "'$ent2' attribute ($objValue) for '$propDefName'.\n              "
                      if {$propDefName == "geometric_validation_property" || $propDefName == "assembly_validation_property"} {
                        append emsg "($recPracNames(valprop), Sec. 8)"
                      } elseif {$propDefName == "pmi_validation_property"} {
                        if {$stepAP == "AP242"} {
                          append emsg "($recPracNames(pmi242), Sec. 10)"
                        } else {
                          append emsg "($recPracNames(pmi203), Sec. 6)"
                        }
                      } elseif {$propDefName == "tessellated_validation_property"} {
                        append emsg "($recPracNames(tessgeom), Sec. 8)"
                      } elseif {$propDefName == "attribute_validation_property"} {
                        append emsg "($recPracNames(uda), Sec. 8)"
                      }
                      errorMsg $emsg
                      set invalid $emsg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $emsg]
                    }
                  }
                }
              }
            }

# value in spreadsheet
            if {$ok && [info exists propDefID]} {
              set c [string index [cellRange 1 $col($pd)] 0]
              set r $propDefIDRow($propDefID)

# colName
              if {![info exists pdheading($col($pd))]} {
                $cells($pd) Item 3 $c $colName
                $cells($pd) Item 3 [string index [cellRange 1 [expr {$col($pd)+1}]] 0] "attribute"
                set pdheading($col($pd)) 1
              }

# keep track of rows with validation properties
              if {[lsearch $propDefRow $r] == -1 && [string first "validation_property" $propDefName] != -1} {lappend propDefRow $r}

              set val [[$cells($pd) Item $r $c] Value]
              if {$val == " "} {set val ""}
              if {$invalid != ""} {lappend syntaxErr($pd) [list $r $col($pd) $invalid]}

              if {$val == ""} {
                $cells($pd) Item $r $c $objValue
              } else {

# adjust val for length with unnecessary exponent of 1
                if {[string first "exponent" $ent2] != -1 && [string first "length" $valName] != -1} {
                  set val [string range $val 0 end-3]
                }

                if {![info exists nrep] || $c != "G"} {
                  $cells($pd) Item $r $c "$val[format "%c" 10]$objValue"

# add nrep count
                } elseif {$maxrep > 1} {
                  if {$nrep != 2} {
                    $cells($pd) Item $r $c "$val[format "%c" 10]($nrep) $objValue"
                  } else {
                    if {[string range $val 0 2] != "(1)"} {
                      $cells($pd) Item $r $c "(1) $val[format "%c" 10]($nrep) $objValue"
                    } else {
                      $cells($pd) Item $r $c "$val[format "%c" 10]($nrep) $objValue"
                    }
                  }
                } else {
                  $cells($pd) Item $r $c "$val[format "%c" 10]$objValue"
                }
              }

# entity reference in spreadsheet
              if {[info exists r]} {
                incr col($pd)
                set c [string index [cellRange 1 $col($pd)] 0]
                set val [[$cells($pd) Item $r $c] Value]
                if {$val == ""} {
                  $cells($pd) Item $r $c "#$objID $ent2"
                } else {
                  $cells($pd) Item $r $c "$val[format "%c" 10]#$objID $ent2"
                }
                set pdcol [expr {max($col($pd),$pdcol)}]

# add blank columns for units and exponent
                if {[info exists valName]} {
                  if {[string first "area" $valName] == -1 && [string first "volume" $valName] == -1 && [string first "density" $valName] == -1 && \
                      ([string first "length_unit" $ent2] != -1 || [string first "description" $ent2] != -1 || [string first "value" $ent2] > 0)} {
                    set i1 4
                    if {[string first "length" $valName] != -1 || [string first "thickness" $valName] != -1 || $valName == ""} {set i1 2}
                    #outputMsg "$ent2\n $valName\n  here $i1"
                    for {set i 0} {$i < $i1} {incr i} {
                      incr col($pd)
                      set c [string index [cellRange 1 $col($pd)] 0]
                      set val [[$cells($pd) Item $r $c] Value]
                      if {$val == ""} {
                        $cells($pd) Item $r $c " "
                      } else {
                        $cells($pd) Item $r $c "$val[format "%c" 10] "
                      }
                      set pdcol [expr {max($col($pd),$pdcol)}]
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }
  incr entLevel -1
}

# -------------------------------------------------------------------------------
proc valPropFormat {} {
  global cells col excelVersion propDefRow recPracNames row stepAP thisEntType worksheet valPropLink

  if {[info exists cells($thisEntType)] && $col($thisEntType) > 4} {
    outputMsg " property_definition"
  
# delete unused columns
    set delcol 0
    set ndelcol 0
    for {set i [expr {$col($thisEntType)-0}]} {$i > 3} {incr i -1} {
      set val [[$cells($thisEntType) Item 3 $i] Value]
      if {$val == ""} {
        set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
        $range Delete
        incr ndelcol 
      }
    }
    set col($thisEntType) [expr {$col($thisEntType)-$ndelcol}]
  
# sort
    if {$excelVersion >= 12} {
      set ranrow $row($thisEntType)
      if {$ranrow > 8} {
        set range [$worksheet($thisEntType) Range [cellRange 3 1] [cellRange $ranrow $col($thisEntType)]]
        set tname [string trim "TABLE-$thisEntType"]
        [[$worksheet($thisEntType) ListObjects] Add 1 $range] Name $tname
        [[$worksheet($thisEntType) ListObjects] Item $tname] TableStyle "TableStyleLight1" 
      }
    }

# header
    catch {$cells($thisEntType) Item 2 5 "Properties"}
    set range [$worksheet($thisEntType) Range "E2"]
    $range HorizontalAlignment [expr -4108]
    [$range Font] Bold [expr 1]
    [$range Interior] ColorIndex [expr 36]
    set range [$worksheet($thisEntType) Range [cellRange 2 5] [cellRange 2 $col($thisEntType)]]
    $range MergeCells [expr 1]

# set rows for colors and borders
    if {[catch {
      set r1 1
      set r2 $r1
      set r3 {}
      foreach r $propDefRow {
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
    } emsg]} {
      errorMsg "ERROR formatting Validation Properties 1: $emsg"
    }

# colors and borders
    if {[catch {
      set j 0
      for {set i 5} {$i <= $col($thisEntType)} {incr i 2} {
        foreach r $r3 {
          set r1 [lindex $r 0]
          set r2 [lindex $r 1]
          set range [$worksheet($thisEntType) Range [cellRange $r1 $i] [cellRange $r2 [expr {$i+1}]]]
          [$range Interior] ColorIndex [lindex [list 36 35] [expr {$j%2}]]
  
          if {$i == 5 && $r2 > 3} {
            if {$r1 < 4} {set r1 4}
            set range [$worksheet($thisEntType) Range [cellRange $r1 5] [cellRange $r2 $col($thisEntType)]]
            for {set k 7} {$k <= 12} {incr k} {
              if {$k != 9 || [expr {$row($thisEntType)+2}] != $r2} {
                catch {[[$range Borders] Item [expr $k]] Weight [expr 1]}
              }
            }
          }
        }
        incr j
      }
    } emsg]} {
      errorMsg "ERROR formatting Validation Properties 2: $emsg"
    }
    
# left and right borders in header
    for {set i 5} {$i <= $col($thisEntType)} {incr i} {
      set range [$worksheet($thisEntType) Range [cellRange 3 $i] [cellRange 3 $i]]
      catch {[[$range Borders] Item [expr 7]]  Weight [expr 1]}
      catch {[[$range Borders] Item [expr 10]] Weight [expr 1]}
    }
    
# bottom bold line
    #set range [$worksheet($thisEntType) Range [cellRange $row($thisEntType) 5] [cellRange $row($thisEntType) $col($thisEntType)]]
    #[[$range Borders] Item [expr 9]] Weight [expr -4138]
    
# fix column widths
    set colrange [[[$worksheet($thisEntType) UsedRange] Columns] Count]
    for {set i 1} {$i <= $colrange} {incr i} {
      set val [[$cells($thisEntType) Item 3 $i] Value]
      if {$val == "value name"} {
        for {set i1 $i} {$i1 <= $colrange} {incr i1} {
          set range [$worksheet($thisEntType) Range [cellRange -1 $i1]]
          $range ColumnWidth [expr 96]
        }
        break
      }
    }
    [$worksheet($thisEntType) Columns] AutoFit
    [$worksheet($thisEntType) Rows] AutoFit

# group columns
    set ni 0
    for {set i 6} {$i <= $col($thisEntType)} {incr i 2} {
      set let "[string index [cellRange 1 $i] 0]"
      set range [$worksheet($thisEntType) Range "$let:$let"]
      [$range Columns] Group
    }
    [$worksheet($thisEntType) Outline] ShowLevels [expr 0] [expr 1]
    
# link to RP
    if {$valPropLink} {
      $cells($thisEntType) Item 2 1 "See CAx-IF Recommended Practices for Validation Property Definitions"
      set range [$worksheet($thisEntType) Range A2:D2]
      $range MergeCells [expr 1]
      set anchor [$worksheet($thisEntType) Range A2]
      [$worksheet($thisEntType) Hyperlinks] Add $anchor [join "https://www.cax-if.org/joint_testing_info.html#recpracs"] [join ""] [join "Link to CAx-IF Recommended Practices"]    
    }
  }
}