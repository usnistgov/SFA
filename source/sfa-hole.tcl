proc spmiHoleStart {entType} {
  global objDesign
  global col entLevel ent entAttrList holeEntity ht lastEnt opt pmiCol pmiHeading pmiStartCol spmiEntity spmiRow

  if {$opt(DEBUG1)} {outputMsg "START spmiHoleStart $entType" red}

  set dir   [list direction direction_ratios]
  set a2p3d [list axis2_placement_3d name axis $dir ref_direction $dir]

# measures
  set qualifier1 [list type_qualifier name]
  set qualifier2 [list value_format_type_qualifier format_type]

  set length_measure1 [list length_measure_with_unit value_component unit_component]
  set length_measure2 [list length_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set length_measure3 [list length_measure_with_unit_and_measure_representation_item_and_qualified_representation_item value_component unit_component name qualifiers $qualifier1 $qualifier2]
  set angle_measure1  [list plane_angle_measure_with_unit value_component unit_component]
  set plength_measure [list positive_length_measure_with_unit value_component unit_component]
  set pangle_measure  [list positive_plane_angle_measure_with_unit value_component unit_component]

# counter and spotface holes
  set tol_val [list tolerance_value lower_bound $length_measure2 $length_measure1 $angle_measure1 upper_bound $length_measure2 $length_measure1 $angle_measure1]
  set lim_fit [list limits_and_fits form_variance zone_variance grade source]

  set basic_round_hole [list basic_round_hole name depth $plength_measure depth_tolerance $tol_val diameter $plength_measure diameter_tolerance $tol_val $lim_fit placement $a2p3d through_hole]
  set explicit_round_hole [list explicit_round_hole depth $plength_measure depth_tolerance $tol_val diameter $plength_measure diameter_tolerance $tol_val $lim_fit placement $a2p3d]

  set spotface_definition [lreplace $explicit_round_hole 0 0 "spotface_definition"]
  set spotface_definition [concat $spotface_definition [list spotface_radius $plength_measure spotface_radius_tolerance $tol_val]]

  set PMIP(basic_round_hole) $basic_round_hole
  set PMIP(simplified_counterbore_hole_definition)  [list simplified_counterbore_hole_definition name placement $a2p3d counterbore $explicit_round_hole]
  set PMIP(simplified_counterdrill_hole_definition) [list simplified_counterdrill_hole_definition name placement $a2p3d counterbore $explicit_round_hole \
      counterdrill_angle $pangle_measure counterdrill_angle_tolerance $tol_val]
  set PMIP(simplified_countersink_hole_definition)  [list simplified_countersink_hole_definition name placement $a2p3d \
      countersink_angle $pangle_measure countersink_angle_tolerance $tol_val countersink_diameter $plength_measure countersink_diameter_tolerance $tol_val $lim_fit]
  set PMIP(simplified_spotface_hole_definition)     [list simplified_spotface_hole_definition name placement $a2p3d counterbore $spotface_definition]

# add drilled hole attributes
  set drilled_hole [list drilled_hole_depth $plength_measure drilled_hole_depth_tolerance $tol_val drilled_hole_diameter $plength_measure drilled_hole_diameter_tolerance $tol_val $lim_fit through_hole]
  foreach idx [array names PMIP] {set PMIP($idx) [concat $PMIP($idx) $drilled_hole]}

# generate non-simplified versions
  foreach idx [array names PMIP] {
    if {[string first "simplified" $idx] == 0} {
      set nidx [string range $idx 11 end]
      set PMIP($nidx) [lreplace $PMIP($idx) 0 0 $nidx]
    }
  }
  if {![info exists PMIP($entType)]} {return}

  set ht $entType
  set lastEnt {}
  set entAttrList {}
  set pmiCol 0
  set spmiRow($ht) {}

  catch {unset pmiHeading}
  catch {unset ent}
  catch {unset holeEntity}

  outputMsg " Adding PMI Representation Analysis" blue
  lappend spmiEntity $entType

  if {$opt(DEBUG1)} {outputMsg \n}
  set entLevel 0
  setEntAttrList $PMIP($ht)
  if {$opt(DEBUG1)} {outputMsg "entAttrList $entAttrList"}
  if {$opt(DEBUG1)} {outputMsg \n}

  set startent [lindex $PMIP($ht) 0]
  set n 0
  set entLevel 0

# get next unused column by checking if there is a colName
  set pmiStartCol($ht) [expr {[getNextUnusedColumn $startent]+1}]

# process all, call spmiHoleReport
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    if {[$objEntity Type] == $startent} {
      if {$n < 1048576} {
        if {[expr {$n%2000}] == 0} {
          if {$n > 0} {outputMsg "  $n"}
          update
        }
        spmiHoleReport $objEntity
        if {$opt(DEBUG1)} {outputMsg \n}
      }
      incr n
    }
  }
  set col($ht) $pmiCol
  set pmiStartCol($ht) [expr {$pmiStartCol($ht)-1}]
}

# -------------------------------------------------------------------------------
proc spmiHoleReport {objEntity} {
  global badAttributes cells col dim DTR hole holerep holeDim holeDimType holeDefinitions holeEntity holeType holeUnit ht
  global entLevel ent entAttrList lastEnt numBore opt pmiCol pmiColumns pmiHeading pmiModifiers pmiUnicode recPracNames spaces
  global spmiEnts spmiID spmiIDRow spmiRow spmiTypesPerFile syntaxErr thruHole

  if {$opt(DEBUG1)} {outputMsg "spmiHoleReport" red}

# entLevel is very important, keeps track level of entity in hierarchy
  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]

  if {[string first "handle" $objEntity] != -1} {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set objAttributes [$objEntity Attributes]
    set ent($entLevel) $objType

    if {$opt(DEBUG1) && $objType != "cartesian_point"} {outputMsg "$ind ENT $entLevel #$objID=$objType (ATR=[$objAttributes Count])" blue}

    set follow 1
    set lastEnt "$objID $objType"
    if {$entLevel == 1} {
      set holeEntity $objEntity
      set holeType [$objEntity Type]
      set numBore 0
    }

# check if there are rows with hole features
    if {$spmiEnts($objType)} {
      set spmiID $objID
      if {![info exists spmiIDRow($ht,$spmiID)]} {
        incr entLevel -1
        return
      }
    }

# loop on all attributes
    ::tcom::foreach objAttribute $objAttributes {
      set objName  [$objAttribute Name]
      if {$entLevel < 1} {set entLevel 1}
      set ent1 "$ent($entLevel) $objName"
      set ent2 "$ent($entLevel).$objName"

# look for entities with bad attributes that cause a crash
      set okattr 1
      if {[info exists badAttributes($objType)]} {foreach ba $badAttributes($objType) {if {$ba == $objName} {set okattr 0}}}

# get attribute value
      if {[catch {
        set objValue [$objAttribute Value]
      } emsg]} {
        set okattr 0
      }

      if {$okattr} {
        set objNodeType [$objAttribute NodeType]
        set objSize [$objAttribute Size]
        set objAttrType [$objAttribute Type]
        set idx [lsearch $entAttrList $ent1]

# -----------------
# nodeType = 18, 19
        if {$objNodeType == 18 || $objNodeType == 19} {
          if {[catch {
            if {$idx != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType-[$objAttribute AggNodeType], $objSize, $objAttrType)"}

              if {[info exists cells($ht)]} {
                set ok 0

# get values for these entity and attribute pairs
                switch -glob $ent1 {
                  "*counter* counter*" -
                  "*counter* drilled*" -
                  "*spotface* drilled*" -
                  "basic_round_hole d*" -
                  "explicit_round_hole d*" -
                  "spotface_definition d*" -
                  "spotface_definition s*" {

# get type of hole dimension or tolerance
                    set holeDimType $objName

# check for multiple counterbore
                    if {$ent1 == "explicit_round_hole depth"} {
                      incr numBore
                      if {$numBore > 1} {errorMsg "Multiple counterbores for a '$holeType' ($recPracNames(holes), Sec. 5.1.3, Fig. 8)"}
                    }
                  }

                  "*length_measure_with_unit* value_component" -
                  "*plane_angle_measure_with_unit* value_component" {
# lengths and angles, save based on the type, tolerances have two values
                    if {[info exists holeDimType]} {
                      set ok 1
                      set invalid ""
                      if {[string first "angle" $ent1] != -1} {
                        set aunit [[[[$objEntity Attributes] Item [expr 2]] Value] Type]
                        if {[string first "length_unit" $aunit] != -1} {
                          set msg "Syntax Error: Bad units [formatComplexEnt $aunit] for a '$objType'."
                          errorMsg $msg
                          lappend syntaxErr([lindex $ent1 0]) [list [$objEntity P21ID] "unit_component" $msg]
                        }
                        if {[string first "conversion" $aunit] == -1} {set objValue [trimNum [expr {$objValue/$DTR}]]}
                      }
# first value
                      if {![info exists holeDim($holeDimType)]} {
                        set holeDim($holeDimType) $objValue
# second value
                      } else {
# format for bilateral tolerance
                        if {[string first "tolerance" $holeDimType] != -1} {
                          if {$objValue == [expr {abs($holeDim($holeDimType))}] && $objValue > $holeDim($holeDimType)} {
                            set holeDim($holeDimType) "$pmiUnicode(plusminus) $objValue"
                            lappend spmiTypesPerFile "bilateral tolerance"
                          } elseif {$objValue != [expr {abs($holeDim($holeDimType))}]} {
                            append holeDim($holeDimType) " $objValue"
                            lappend spmiTypesPerFile "non-bilateral tolerance"
                          } else {
                            set msg "Syntax Error: Tolerance lower and upper bounds ($objValue) are equal."
                            errorMsg $msg
                            lappend syntaxErr(tolerance_value) [list "-$spmiIDRow($ht,$spmiID)" "lower_bound" $msg]
                            lappend syntaxErr(tolerance_value) [list "-$spmiIDRow($ht,$spmiID)" "upper_bound" $msg]
                          }
                        } else {
                          append holeDim($holeDimType) " $objValue"
                        }
                      }
                      incr hole(idx)
                    }
                  }

                  "*length_measure_with_unit* unit_component" {
# get units if not already found processing dimensional tolerances
                    if {[string first "si_unit" [$objValue Type]] != -1} {
                      set holeUnit "MM"
                    } else {
                      set holeUnit "INCH"
                    }
                    if {![info exists dim(unit)]} {set dim(unit) $holeUnit}
                  }
                }
              }

# if referred to another, get the entity
              if {[string first "handle" $objValue] != -1} {
                if {[catch {
                  [$objValue Type]
                  set errstat [spmiHoleReport $objValue]
                  if {$errstat} {break}
                } emsg1]} {

# referred entity is actually a list of entities
                  if {[catch {
                    ::tcom::foreach val1 $objValue {spmiHoleReport $val1}
                  } emsg2]} {
                    foreach val2 $objValue {spmiHoleReport $val2}
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "Error processing Holtol ($objNodeType $ent2)\n $emsg3"
            set entLevel 1
          }

# --------------
# nodeType = 20
        } elseif {$objNodeType == 20} {
          if {$idx != -1} {
            if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType-[$objAttribute AggNodeType], $objSize, $objAttrType)"}

# no values to get from this nodetype, but get the entities that are referred to
            if {[catch {
              ::tcom::foreach val1 $objValue {spmiHoleReport $val1}
            } emsg]} {
              foreach val2 $objValue {spmiHoleReport $val2}
            }
          }

# ---------------------
# nodeType = 5 (!= 18,19,20)
        } else {
          if {[catch {
            if {$idx != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objAttrType)  ($ent1)"}

              if {[info exists cells($ht)]} {
                set ok 0
                set colName ""
                set ov $objValue
                set invalid ""

# get values for these entity and attribute pairs
                switch -glob $ent1 {
                  "limits_and_fits form_variance" -
                  "limits_and_fits grade" {
                    append holeDim($holeDimType) "$objValue "
                    if {[string first "grade" $ent1] != -1} {
                      set holeDim($holeDimType) "([string trim $holeDim($holeDimType)])"
                      lappend spmiTypesPerFile "limits and fits"
                    }
                  }
                  "basic_round_hole through_hole" -
                  "*hole_definition through_hole" {
                    set thruHole $objValue
                  }
                  "basic_round_hole name" -
                  "*hole_definition name" {
                    set holeName $objValue
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "Error processing Holtol ($objNodeType $ent2)\n $emsg3"
            set entLevel 1
          }
        }
      }
    }
  }
  incr entLevel -1

# -------------------------------------------------------------------------------
# done processing entities, construct hole dimension
  if {$entLevel == 0} {
    set holerep ""

# check for repetitive dimensions
    set nhole 0
    foreach occ [list basic_round_hole_occurrence counterbore_hole_occurrence counterdrill_hole_occurrence \
              countersink_hole_occurrence spotface_occurrence] {
      set e0s [$holeEntity GetUsedIn [string trim $occ] [string trim definition]]
      ::tcom::foreach e0 $e0s {incr nhole}
    }
    if {$nhole > 1} {
      append holerep "$nhole\X "
      lappend spmiTypesPerFile "repetitive dimensions"
    } elseif {$nhole == 0} {
      errorMsg "No hole occurrences refer to '[$holeEntity Type]'."
    }

# main hole diameter, depth, and tolerances
    catch {unset hd}
    if {[info exists holeDim(drilled_hole_diameter)]} {
      append holerep "$pmiUnicode(diameter)[trimNum $holeDim(drilled_hole_diameter)]"
      lappend spmiTypesPerFile "diameter"
      set hd "drill $holeDim(drilled_hole_diameter)"
      if {[info exists holeDim(drilled_hole_diameter_tolerance)]} {append holerep " $holeDim(drilled_hole_diameter_tolerance)"}
    }

# drill depth specified, correct only if thru hole is false
    if {[info exists holeDim(drilled_hole_depth)]} {
      append holerep "  $pmiModifiers(depth)[trimNum $holeDim(drilled_hole_depth)]"
      if {[info exists holeDim(drilled_hole_depth_tolerance)]} {append holerep " $holeDim(drilled_hole_depth_tolerance)"}
      lappend spmiTypesPerFile "depth"
      append hd " $holeDim(drilled_hole_depth)"
      if {$thruHole} {
        set msg "Syntax Error: through_hole should be FALSE if drilled_hole_depth is specified OR if though_hole is TRUE then the drilled_hole_depth should not be specified.$spaces\($recPracNames(holes), Sec. 5.1.1.1)"
        errorMsg $msg
        lappend syntaxErr([$holeEntity Type]) [list [$holeEntity P21ID] "through_hole" $msg]
        lappend syntaxErr([$holeEntity Type]) [list [$holeEntity P21ID] "drilled_hole_depth" $msg]
      }

# no depth and not a thru hole
    } elseif {!$thruHole} {
      set msg ""
      if {[string first "basic_round" [$holeEntity Type]] == -1} {
        set msg "Syntax Error: through_hole should be TRUE if drilled_hole_depth is not specified.$spaces\($recPracNames(holes), Sec. 5.1.1.1)"
      } elseif {![info exists holeDim(depth)]} {
        set msg "Syntax Error: through_hole should be TRUE if depth is not specified.$spaces\($recPracNames(holes), Sec. 5.1.1.1)"
      }
      if {$msg != ""} {
        errorMsg $msg
        lappend syntaxErr([$holeEntity Type]) [list [$holeEntity P21ID] "through_hole" $msg]
      }
    }
    if {[info exists hd]} {lappend holeDefinitions([$holeEntity P21ID]) $hd}

# countersink diameter, angle, and tolerances
    if {[info exists holeDim(countersink_diameter)]} {
      append holerep "[format "%c" 10]$pmiModifiers(countersink)$pmiUnicode(diameter)[trimNum $holeDim(countersink_diameter)]"
      if {[info exists holeDim(countersink_diameter_tolerance)]} {append holerep " $holeDim(countersink_diameter_tolerance)"}
      append holerep " X $holeDim(countersink_angle)$pmiUnicode(degree)"
      lappend spmiTypesPerFile "diameter"
      lappend spmiTypesPerFile "countersink"
      if {[info exists holeDim(countersink_angle_tolerance)]} {append holerep " $holeDim(countersink_angle_tolerance)$pmiUnicode(degree)"}
      lappend holeDefinitions([$holeEntity P21ID]) "countersink $holeDim(countersink_diameter) $holeDim(countersink_angle)"
    }

# basic, (multiple) counterbore, or spotface diameter, depth, and tolerances
    if {[info exists holeDim(diameter)]} {
      set nhdim 0
      foreach hdim $holeDim(diameter) {
        if {$holerep != ""} {append holerep [format "%c" 10]}
        if {[string first "counterbore" [$holeEntity Type]] != -1} {
          append holerep $pmiModifiers(counterbore)
          lappend spmiTypesPerFile "counterbore"
          set type "counterbore"
        } elseif {[string first "spotface" [$holeEntity Type]] != -1} {
          append holerep "$pmiModifiers(spotface) "
          lappend spmiTypesPerFile "spotface"
          set type "spotface"
        } elseif {[string first "basic_round" [$holeEntity Type]] != -1} {
          lappend spmiTypesPerFile "round_hole"
          set type "round_hole"
        }

        append holerep "$pmiUnicode(diameter)$hdim"
        if {[info exists holeDim(diameter_tolerance)]} {append holerep " $holeDim(diameter_tolerance)"}
        lappend spmiTypesPerFile "diameter"
        if {[info exists holeDim(depth)]} {
          append holerep "  $pmiModifiers(depth)[lindex $holeDim(depth) $nhdim]"
          if {[info exists holeDim(depth_tolerance)]} {append holerep " $holeDim(depth_tolerance)"}
          lappend spmiTypesPerFile "depth"
          lappend holeDefinitions([$holeEntity P21ID]) "$type $hdim [lindex $holeDim(depth) $nhdim]"
        } else {
          lappend holeDefinitions([$holeEntity P21ID]) "$type $hdim"
        }
        incr nhdim
      }
    }

# thru hole and name
    lappend holeDefinitions([$holeEntity P21ID]) $thruHole
    lappend holeDefinitions([$holeEntity P21ID]) $holeName

# -------------------------------------------------------------------------------
# report complete hole representation (holerep)
    if {[catch {
      set cellComment ""
      if {[info exists holerep] && [info exists spmiIDRow($ht,$spmiID)]} {
        if {![info exists pmiColumns([$holeEntity Type])]} {set pmiColumns([$holeEntity Type]) [getNextUnusedColumn $ht]}
        set c [string index [cellRange 1 $pmiColumns([$holeEntity Type])] 0]
        set r $spmiIDRow($ht,$spmiID)
        if {![info exists pmiHeading($pmiColumns([$holeEntity Type]))]} {
          set colName "Hole[format "%c" 10]Feature"
          $cells($ht) Item 3 $c $colName
          set pmiHeading($pmiColumns([$holeEntity Type])) 1
          set pmiCol [expr {max($pmiColumns([$holeEntity Type]),$pmiCol)}]
          set comment "See Help > User Guide (section 6.1.3) for an explanation of how the hole dimensions below are constructed."
          if {[info exists hole(unit)]} {append comment "\n\nDimension units: $hole(unit)"}
          append comment "\n\nRepetitive dimensions (e.g., 4X) might be shown for holes.  They are computed based on the number of counterbore, sink, drill, and spotface occurrence entities that reference the hole definition."
          addCellComment $ht 3 $c $comment
        }

# write hole to spreadsheet
        $cells($ht) Item $r $pmiColumns([$holeEntity Type]) $holerep
        catch {unset holerep}
        catch {unset holeDim}

# keep track of rows with semantic PMI
        if {[lsearch $spmiRow($ht) $r] == -1} {lappend spmiRow($ht) $r}
      }

    } emsg]} {
      errorMsg "Error adding Hole Tolerance: $emsg"
    }
    set hole(name) ""
    catch {unset holeDim}
  }

  return 0
}
