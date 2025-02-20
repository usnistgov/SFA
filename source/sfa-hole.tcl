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

# missing values
              if {$objName == "countersink_angle" || $objName == "countersink_diameter" || $objName == "drilled_hole_diameter"} {
                if {$objValue == ""} {
                  set msg "Syntax Error: Missing required '$objName' attribute.$spaces\($recPracNames(holes), Sec. 5.1.1.1)"
                  errorMsg $msg
                  lappend syntaxErr([$holeEntity Type]) [list [$holeEntity P21ID] $objName $msg]
                }
              }

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
    set htype [$holeEntity Type]
    set hid [$holeEntity P21ID]

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
      errorMsg "No hole occurrences refer to '$htype'."
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
        set msg "Syntax Error: drilled_hole_depth is not required if though_hole is TRUE.$spaces\($recPracNames(holes), Sec. 5.1.1.1)"
        errorMsg $msg
        lappend syntaxErr($htype) [list $hid "through_hole" $msg]
        lappend syntaxErr($htype) [list $hid "drilled_hole_depth" $msg]
      }

# no depth and not a thru hole
    } elseif {!$thruHole} {
      if {[string first "basic_round" $htype] == -1 || ![info exists holeDim(depth)]} {
        set aname "drilled_hole_depth"
        if {[string first "basic_round" $htype] != -1} {set aname "depth"}
        set msg "Syntax Error: $aname is required if through_hole is FALSE.$spaces\($recPracNames(holes), Sec. 5.1.1.1)"
        errorMsg $msg
        lappend syntaxErr($htype) [list $hid "through_hole" $msg]
        lappend syntaxErr($htype) [list $hid $aname $msg]
      }
    }
    if {[info exists hd]} {lappend holeDefinitions($hid) $hd}

# countersink angle, diameter, and tolerances
    if {[info exists holeDim(countersink_angle)]} {
      append holerep "[format "%c" 10]$pmiModifiers(countersink)"
      if {[info exists holeDim(countersink_diameter)]} {
        append holerep "$pmiUnicode(diameter)[trimNum $holeDim(countersink_diameter)]"
        if {[info exists holeDim(countersink_diameter_tolerance)]} {append holerep " $holeDim(countersink_diameter_tolerance)"}
        lappend spmiTypesPerFile "diameter"
        append holerep " X "
      }
      append holerep "$holeDim(countersink_angle)$pmiUnicode(degree)"
      lappend spmiTypesPerFile "countersink"
      if {[info exists holeDim(countersink_angle_tolerance)]} {append holerep " $holeDim(countersink_angle_tolerance)$pmiUnicode(degree)"}
      if {[info exists holeDim(countersink_diameter)]} {
        lappend holeDefinitions($hid) "countersink $holeDim(countersink_diameter) $holeDim(countersink_angle)"
      }
    }

# basic, (multiple) counterbore, or spotface diameter, depth, and tolerances
    if {[info exists holeDim(diameter)]} {
      set nhdim 0
      foreach hdim $holeDim(diameter) {
        if {$holerep != ""} {append holerep [format "%c" 10]}
        if {[string first "counterbore" $htype] != -1} {
          append holerep $pmiModifiers(counterbore)
          lappend spmiTypesPerFile "counterbore"
          set type "counterbore"
        } elseif {[string first "spotface" $htype] != -1} {
          append holerep "$pmiModifiers(spotface) "
          lappend spmiTypesPerFile "spotface"
          set type "spotface"
        } elseif {[string first "basic_round" $htype] != -1} {
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
          lappend holeDefinitions($hid) "$type $hdim [lindex $holeDim(depth) $nhdim]"
          if {$thruHole && [string first "counter" $htype] == -1} {
            set msg "Syntax Error: depth is not required if though_hole is TRUE.$spaces\($recPracNames(holes), Sec. 5.1.1.1)"
            errorMsg $msg
            lappend syntaxErr($htype) [list $hid "through_hole" $msg]
            lappend syntaxErr($htype) [list $hid "depth" $msg]
          }
        } else {
          lappend holeDefinitions($hid) "$type $hdim"
        }
        incr nhdim
      }
    }

# thru hole and name
    lappend holeDefinitions($hid) $thruHole
    lappend holeDefinitions($hid) $holeName

# -------------------------------------------------------------------------------
# report complete hole representation (holerep)
    if {[catch {
      set cellComment ""
      if {[info exists holerep] && [info exists spmiIDRow($ht,$spmiID)]} {
        if {![info exists pmiColumns($htype)]} {set pmiColumns($htype) [getNextUnusedColumn $ht]}
        set c [string index [cellRange 1 $pmiColumns($htype)] 0]
        set r $spmiIDRow($ht,$spmiID)
        if {![info exists pmiHeading($pmiColumns($htype))]} {
          set colName "Hole[format "%c" 10]Feature"
          $cells($ht) Item 3 $c $colName
          set pmiHeading($pmiColumns($htype)) 1
          set pmiCol [expr {max($pmiColumns($htype),$pmiCol)}]
          set comment "See Help > User Guide (section 6.1.3) for an explanation of how the hole dimensions below are constructed."
          if {[info exists hole(unit)]} {append comment "\n\nDimension units: $hole(unit)"}
          append comment "\n\nRepetitive dimensions (e.g., 4X) might be shown for holes.  They are computed based on the number of counterbore, sink, drill, and spotface occurrence entities that reference the hole definition."
          addCellComment $ht 3 $c $comment
        }

# write hole to spreadsheet
        $cells($ht) Item $r $pmiColumns($htype) $holerep
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

# -------------------------------------------------------------------------------
# holes counter and spotface
proc x3dHoles {} {
  global brepScale DTR entCount gen holeDefinitions holeUnit maxxyz opt recPracNames spaces syntaxErr viz x3dFile
  global objDesign

  set drillPoint [trimNum [expr {$maxxyz*0.02}]]
  set head 1
  set holeDEF {}

  set scale 1.
  if {$brepScale == 1. && $holeUnit == "INCH"} {set scale 25.4}

  ::tcom::foreach e0 [$objDesign FindObjects [string trim item_identified_representation_usage]] {
    if {[catch {
      set e1 [[[$e0 Attributes] Item [expr 3]] Value]
      set e2 [[[$e0 Attributes] Item [expr 5]] Value]
      if {[catch {
        set e2type [$e2 Type]
      } emsg1]} {
        ::tcom::foreach e2a $e2 {set e2 $e2a; break}
      }
      if {[string first "occurrence" [$e1 Type]] != -1 && [$e2 Type] == "mapped_item"} {
        set defID   [[[[$e1 Attributes] Item [expr 5]] Value] P21ID]
        set defType [[[[$e1 Attributes] Item [expr 5]] Value] Type]
        set holeOccName [[[$e1 Attributes] Item [expr 1]] Value]

# hole name
        set holeName [split $defType "_"]
        foreach idx {0 1} {
          if {[string first "counter" [lindex $holeName $idx]] != -1 || [string first "spotface" [lindex $holeName $idx]] != -1} {set holeName [lindex $holeName $idx]}
        }
        if {$defType == "basic_round_hole"} {set holeName $defType}

# check if there is an a2p3d associated with a hole occurrence
        set e3 [[[$e2 Attributes] Item [expr 3]] Value]
        if {[$e3 Type] == "axis2_placement_3d"} {
          if {$head} {
            outputMsg " Processing hole geometry" green
            puts $x3dFile "\n<!-- HOLES -->\n<Switch whichChoice='0' id='swHole'><Group>"
            set head 0
            set viz(HOLE) 1
          }
          if {[lsearch $holeDEF $defID] == -1} {puts $x3dFile "<!-- $defType $defID -->"}

# hole geometry
          if {[info exists holeDefinitions($defID)]} {

# hole origin and axis transform
            set a2p3d [x3dGetA2P3D $e3]
            set transform [x3dTransform [lindex $a2p3d 0] [lindex $a2p3d 1] [lindex $a2p3d 2] $holeName]

# drilled hole dimensions
            set drill [lindex $holeDefinitions($defID) 0]
            set drillRad [trimNum [expr {[lindex $drill 1]*0.5*$scale}] 5]
            set drillPoint $drillRad
            catch {unset drillDep}
            if {[llength $drill] > 2} {set drillDep [expr {[lindex $drill 2]*$scale}]}

# through hole
            set holeTop "true"
            set thruHole [lindex $holeDefinitions($defID) end-1]
            if {$thruHole == 1} {set holeTop "false"}

# bottom condition
            catch {unset tipDepth}
            set e4s [[[[$e1 Attributes] Item [expr 5]] Value] GetUsedIn [string trim "round_hole_bottom_condition"] [string trim target]]
            ::tcom::foreach e4 $e4s {
              if {$e4 != ""} {
                set rhbc [[[$e4 Attributes] Item [expr 2]] Value]
                if {$rhbc == "conical"} {
                  set e5 [[[$e4 Attributes] Item [expr 6]] Value]
                  if {$e5 != ""} {
                    set tipAngle [[[$e5 Attributes] Item [expr 1]] Value]
                    set tipDepth [expr {$drillRad/tan($tipAngle*0.5*$DTR)}]
                  } else {
                    set msg "Syntax Error: Missing 'tip_angle' for 'conical' round hole bottom condition"
                    errorMsg $msg
                    lappend syntaxErr(round_hole_bottom_condition) [list [$e4 P21ID] "tip_angle" $msg]
                  }
                } elseif {$rhbc != "flat"} {
                  errorMsg " Round hole bottom condition '$rhbc' is not supported"
                }
              }
            }

            catch {unset sink}
            catch {unset bore}
            set lhd [llength $holeDefinitions($defID)]
            if {$lhd > 1} {
              set holeType [lindex [lindex $holeDefinitions($defID) [expr {$lhd-3}]] 0]

# countersink hole (cylinder, cone)
              if {$holeType == "countersink"} {
                set sink [lindex $holeDefinitions($defID) 1]

# compute length of countersink from angle and radius
                set sinkRad [trimNum [expr {[lindex $sink 1]*0.5*$scale}] 5]
                set sinkAng [expr {[lindex $sink 2]*0.5}]
                set sinkDep [expr {($sinkRad-$drillRad)/tan($sinkAng*$DTR)}]

# check for bad radius and depth
                if {$sinkRad <= $drillRad} {
                  set msg "Syntax Error: $holeType diameter <= drill diameter"
                  errorMsg $msg
                  foreach ent [list $holeType\_hole_definition simplified_$holeType\_hole_definition] {
                    if {[info exists entCount($ent)]} {
                      lappend syntaxErr($ent) [list $defID "countersink_diameter" $msg]
                      lappend syntaxErr($ent) [list $defID "drilled_hole_diameter" $msg]
                    }
                  }
                }
                if {[info exists drillDep]} {
                  if {$sinkDep >= $drillDep} {
                    set msg "Syntax Error: $holeType computed 'depth' >= drill depth"
                    errorMsg $msg
                    foreach ent [list $holeType\_hole_definition simplified_$holeType\_hole_definition] {
                      if {[info exists entCount($ent)]} {lappend syntaxErr($ent) [list $defID "drilled_hole_depth" $msg]}
                    }
                  }
                }

                if {[lsearch $holeDEF $defID] == -1} {
                  puts $x3dFile "$transform<Group DEF='$holeName$defID'>"
                  if {[info exists drillDep]} {
                    puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {($drillDep+$sinkDep)*0.5}] 5]'>"
                    if {$holeTop == "false" || ![info exists tipDepth]} {
                      puts $x3dFile "  <Shape><Cylinder radius='$drillRad' height='[trimNum [expr {$drillDep-$sinkDep}] 5]' top='$holeTop' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 1'/></Appearance></Shape></Transform>"
                    } else {
                      puts $x3dFile "  <Shape><Cylinder radius='$drillRad' height='[trimNum [expr {$drillDep-$sinkDep}] 5]' top='false' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 1'/></Appearance></Shape></Transform>"
                      puts $x3dFile "  <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {$drillDep+($tipDepth*0.5)}] 5]'>"
                      puts $x3dFile "   <Shape><Cone bottomRadius='$drillRad' topRadius='0' height='[trimNum $tipDepth 5]' top='false' bottom='false' solid='false'></Cone><Appearance><Material diffuseColor='0 1 1'/></Appearance></Shape></Transform>"
                    }
                  }
                  puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {$sinkDep*0.5}] 5]'>"
                  puts $x3dFile "  <Shape><Cone bottomRadius='$sinkRad' topRadius='$drillRad' height='[trimNum $sinkDep 5]' top='false' bottom='false' solid='false'></Cone><Appearance><Material diffuseColor='0 1 1'/></Appearance></Shape></Transform>"
                  puts $x3dFile "</Group></Transform>"
                  lappend holeDEF $defID
                } else {
                  puts $x3dFile "$transform<Group USE='$holeName$defID'></Group></Transform>"
                }

# counterbore or spotface hole (2 cylinders, flat cone)
              } elseif {$holeType == "counterbore" || $holeType == "spotface"} {
                set bore [lindex $holeDefinitions($defID) 1]
                set boreRad [expr {[lindex $bore 1]*0.5*$scale}]
                set boreDep [expr {[lindex $bore 2]*$scale}]

# check for bad radius and depth
                if {$boreRad <= $drillRad} {
                  set msg "Syntax Error: $holeType diameter <= drill diameter"
                  errorMsg $msg
                  foreach ent [list $holeType\_hole_definition simplified_$holeType\_hole_definition] {
                    if {[info exists entCount($ent)]} {
                      lappend syntaxErr($ent) [list $defID "counterbore" $msg]
                      lappend syntaxErr($ent) [list $defID "drilled_hole_diameter" $msg]
                    }
                  }
                }
                if {[info exists drillDep]} {
                  if {$boreDep >= $drillDep} {
                    set msg "Syntax Error: $holeType depth >= drill depth"
                    errorMsg $msg
                    foreach ent [list $holeType\_hole_definition simplified_$holeType\_hole_definition] {
                      if {[info exists entCount($ent)]} {
                        lappend syntaxErr($ent) [list $defID "counterbore" $msg]
                        lappend syntaxErr($ent) [list $defID "drilled_hole_depth" $msg]
                      }
                    }
                  }
                }

                if {[lsearch $holeDEF $defID] == -1} {
                  puts $x3dFile "$transform<Group DEF='$holeName$defID'>"
                  if {[info exists drillDep]} {
                    puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {($drillDep+$boreDep)*0.5}] 5]'>"
                    if {$holeTop == "false" || ![info exists tipDepth]} {
                      puts $x3dFile "  <Shape><Cylinder radius='$drillRad' height='[trimNum [expr {$drillDep-$boreDep}] 5]' top='$holeTop' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                    } else {
                      puts $x3dFile "  <Shape><Cylinder radius='$drillRad' height='[trimNum [expr {$drillDep-$boreDep}] 5]' top='false' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                      puts $x3dFile "  <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {$drillDep+($tipDepth*0.5)}] 5]'>"
                      puts $x3dFile "   <Shape><Cone bottomRadius='$drillRad' topRadius='0' height='[trimNum $tipDepth 5]' top='false' bottom='false' solid='false'></Cone><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                    }
                  }
                  puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum $boreDep 5]'>"
                  puts $x3dFile "  <Shape><Cone bottomRadius='$boreRad' topRadius='$drillRad' height='0.001' top='false' bottom='false' solid='false'></Cone><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                  puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {$boreDep*0.5}] 5]'>"
                  puts $x3dFile "  <Shape><Cylinder radius='$boreRad' height='[trimNum $boreDep 5]' top='false' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                  puts $x3dFile "</Group></Transform>"
                  lappend holeDEF $defID
                } else {
                  puts $x3dFile "$transform<Group USE='$holeName$defID'></Group></Transform>"
                }

# basic round hole
              } elseif {$holeType == "round_hole"} {
                set hole [lindex $holeDefinitions($defID) 0]
                set holeRad [expr {[lindex $hole 1]*0.5*$scale}]
                if {[lindex $hole 2] != ""} {
                  set holeDep [expr {[lindex $hole 2]*$scale}]
                } else {
                  set holeDep [expr {[lindex $hole 1]*0.15*$scale}]
                }
                if {[lsearch $holeDEF $defID] == -1} {
                  puts $x3dFile "$transform<Group DEF='$holeName$defID'>"
                  puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {$holeDep*0.5}] 5]'>"
                  if {$holeTop == "false" || ![info exists tipDepth]} {
                    puts $x3dFile "  <Shape><Cylinder radius='$holeRad' height='[trimNum $holeDep 5]' top='$holeTop' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                  } else {
                    puts $x3dFile "  <Shape><Cylinder radius='$holeRad' height='[trimNum $holeDep 5]' top='false' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                    puts $x3dFile "  <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {$holeDep+($tipDepth*0.5)}] 5]'>"
                    puts $x3dFile "   <Shape><Cone bottomRadius='$holeRad' topRadius='0' height='[trimNum $tipDepth 5]' top='false' bottom='false' solid='false'></Cone><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                  }
                  puts $x3dFile "</Group></Transform>"
                  lappend holeDEF $defID
                } else {
                  puts $x3dFile "$transform<Group USE='$holeName$defID'></Group></Transform>"
                }
              }
            }
          } elseif {!$opt(PMISEM) || $gen(None)} {
            errorMsg " Only hole drill entry points are shown when the Analyzer report for Semantic PMI is not selected."
            if {[lsearch $holeDEF $defID] == -1} {lappend holeDEF $defID}
          }

# point and occurrence name at origin of hole
          set e4 [[[$e3 Attributes] Item [expr 2]] Value]
          if {![info exists thruHole]} {set thruHole 0}
          set hname $holeOccName
          if {$hname == ""} {set hname [lindex $holeDefinitions($defID) end]}
          x3dSuppGeomPoint $e4 $drillPoint $hname
        }
      }
    } emsg]} {
      errorMsg "Error adding 'hole' geometry: $emsg"
    }
  }
  if {$viz(HOLE)} {puts $x3dFile "</Group></Switch>\n"}
  catch {unset holeDefinitions}

  set ok 0
  if {![info exists entCount(item_identified_representation_usage)]} {set ok 1} elseif {$entCount(item_identified_representation_usage) == 0} {set ok 1}
  if {$ok} {errorMsg "Syntax Error: Missing IIRU to link hole with explicit geometry.$spaces\($recPracNames(holes), Sec. 5.1.1.2)"}
}
