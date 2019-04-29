proc spmiCounterStart {entType} {
  global objDesign
  global cells col ht entLevel ent entAttrList gtEntity lastEnt opt pmiCol pmiHeading pmiStartCol spmiRow stepAP

  if {$opt(DEBUG1)} {outputMsg "START spmiCounterStart $entType" red}

  set dir           [list direction direction_ratios]
  set a2p3d         [list axis2_placement_3d name axis $dir ref_direction $dir]

# measures
  set qualifier1      [list type_qualifier name]
  set qualifier2      [list value_format_type_qualifier format_type]

  set length_measure1 [list length_measure_with_unit value_component]
  set length_measure2 [list length_measure_with_unit_and_measure_representation_item value_component name]
  set length_measure3 [list length_measure_with_unit_and_measure_representation_item_and_qualified_representation_item value_component name qualifiers $qualifier1 $qualifier2]
  set angle_measure1  [list plane_angle_measure_with_unit value_component]
  set plength_measure [list positive_length_measure_with_unit value_component]
  set pangle_measure  [list positive_plane_angle_measure_with_unit value_component]

# counter and spotface holes
  set tol_val [list tolerance_value lower_bound $length_measure2 $length_measure1 $angle_measure1 upper_bound $length_measure2 $length_measure1 $angle_measure1]
  set lim_fit [list limits_and_fits form_variance zone_variance grade source]
  set explicit_round_hole [list explicit_round_hole depth $plength_measure depth_tolerance $tol_val diameter $plength_measure diameter_tolerance $tol_val $lim_fit placement $a2p3d]

  set PMIP(simplified_counterbore_hole_definition) \
    [list simplified_counterbore_hole_definition placement $a2p3d counterbore $explicit_round_hole \
      drilled_hole_depth $plength_measure drilled_hole_depth_tolerance $tol_val drilled_hole_diameter $plength_measure drilled_hole_diameter_tolerance $tol_val $lim_fit \
      through_hole \
    ]
  set PMIP(simplified_counterdrill_hole_definition) \
    [list simplified_counterdrill_hole_definition placement $a2p3d counterbore $explicit_round_hole \
      counterdrill_angle $pangle_measure counterdrill_angle_tolerance $tol_val \
      drilled_hole_depth $plength_measure drilled_hole_depth_tolerance $tol_val drilled_hole_diameter $plength_measure drilled_hole_diameter_tolerance $tol_val $lim_fit \
      through_hole \
    ]
  set PMIP(simplified_countersink_hole_definition) \
    [list simplified_countersink_hole_definition placement $a2p3d \
      countersink_angle $pangle_measure countersink_angle_tolerance $tol_val countersink_diameter $plength_measure countersink_diameter_tolerance $tol_val $lim_fit \
      drilled_hole_depth $plength_measure drilled_hole_depth_tolerance $tol_val drilled_hole_diameter $plength_measure drilled_hole_diameter_tolerance $tol_val $lim_fit \
      through_hole \
    ]
  set PMIP(simplified_spotface_hole_definition) $PMIP(simplified_counterbore_hole_definition)

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
  catch {unset gtEntity}

  outputMsg " Adding PMI Representation Analysis" blue

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

# process all, call spmiCounterReport
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    if {[$objEntity Type] == $startent} {
      if {$n < 1048576} {
        if {[expr {$n%2000}] == 0} {
          if {$n > 0} {outputMsg "  $n"}
          update
        }
        spmiCounterReport $objEntity
        if {$opt(DEBUG1)} {outputMsg \n}
      }
      incr n
    }
  }
  set col($ht) $pmiCol
  set pmiStartCol($ht) [expr {$pmiStartCol($ht)-1}]
}

# -------------------------------------------------------------------------------
proc spmiCounterReport {objEntity} {
  global assocGeom badAttributes cells col DTR hole holerep holerepID holeDim holeDimType
  global counterEnt counterEntType holtolGeom holeval draftModelCameras ht entLevel ent entAttrList entCount entlevel2 entsWithErrors
  global incrcol lastAttr lastEnt nistName opt pmiCol pmiColumns pmiHeading pmiModifiers pmiStartCol
  global pmiUnicode prefix angDegree recPracNames savedModifier spmiEnts spmiID spmiIDRow spmiRow spmiTypesPerFile syntaxErr tolStandard worksheet

  if {$opt(DEBUG1)} {outputMsg "spmiCounterReport" red}

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
      catch {unset counterEnt}
      catch {unset entlevel2}
      catch {unset assocGeom}
    } elseif {$entLevel == 2} {
      if {![info exists entlevel2]} {set entlevel2 [list $objID $objType]}
    }

# check if there are rows with dt
    if {$spmiEnts($objType) && [string first "datum_feature" $objType] == -1 && [string first "datum_target" $objType] == -1} {
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
      #outputMsg "$ind  $ent1"

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
              set lastAttr $objName

              if {[info exists cells($ht)]} {
                set ok 0

# get values for these entity and attribute pairs
                switch -glob $ent1 {
                  "*counter* counter*" -
                  "*counter* drilled*" {
# get type of hole dimension or tolerance
                    set holeDimType $objName
                  }
                  
# lengths and angles, save based on the type
                  "*length_measure_with_unit* value_component" -
                  "*plane_angle_measure_with_unit* value_component" {
                    set ok 1
                    set invalid ""
                    set holeval $objValue
                    if {[info exists holeDim($holeDimType)]} {
                      append holeDim($holeDimType) " $objValue"
                    } else {
                      set holeDim($holeDimType) $objValue
                    }
                    incr hole(idx)
                  }
                }
              }

# if referred to another, get the entity
              if {[string first "handle" $objValue] != -1} {
                if {[catch {
                  [$objValue Type]
                  set errstat [spmiCounterReport $objValue]
                  if {$errstat} {break}
                } emsg1]} {

# referred entity is actually a list of entities
                  if {[catch {
                    ::tcom::foreach val1 $objValue {spmiCounterReport $val1}
                  } emsg2]} {
                    foreach val2 $objValue {spmiCounterReport $val2}
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing Holtol ($objNodeType $ent2)\n $emsg3"
            set entLevel 1
          }

# --------------
# nodeType = 20
        } elseif {$objNodeType == 20} {
          if {$idx != -1} {
            if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType-[$objAttribute AggNodeType], $objSize, $objAttrType)"}

# no values to get from this nodetype, but get the entities that are referred to
            if {[catch {
              ::tcom::foreach val1 $objValue {spmiCounterReport $val1}
            } emsg]} {
              foreach val2 $objValue {spmiCounterReport $val2}
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
                  "*counter* through_hole" {

# through_hole is the last counter attribute, construct counter FCF                    
                    #foreach item [array names holeDim] {
                    #  outputMsg $item$holeDim($item) green
                    #}
                    set holerep "$pmiUnicode(diameter)$holeDim(drilled_hole_diameter)"
                    lappend spmiTypesPerFile "diameter"
                    if {[info exists holeDim(drilled_hole_depth)]} {
                      append holerep "  $pmiModifiers(depth)$holeDim(drilled_hole_depth)"
                      lappend spmiTypesPerFile "depth"
                    }
                    if {[info exists holeDim(countersink_diameter)]} {
                      append holerep "[format "%c" 10]$pmiModifiers(countersink)$pmiUnicode(diameter)$holeDim(countersink_diameter) X [trimNum [expr {$holeDim(countersink_angle)/$DTR}]]$pmiUnicode(degree)"
                      lappend spmiTypesPerFile "diameter"
                      lappend spmiTypesPerFile "countersink"
                    }
                    #outputMsg $holerep green
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing Holtol ($objNodeType $ent2)\n $emsg3"
            set entLevel 1
          }
        }
      }
    }
  }
  incr entLevel -1

# -------------------------------------------------------------------------------
# write a few more things at the end of processing a semantic PMI entity
  if {$entLevel == 0} {

# associated geometry (5.1.5), find link between dimtol and geometry through geometric_item_specific_usage (gisu)
# counterEnt is either dimensional_location, angular_location, dimensional_size, or angular_size
    if {[catch {
      if {[info exists counterEnt]} {
        set dimtolType [$counterEnt Type]
        set dimtolID   [$counterEnt P21ID]

        if {[string first "_location" $dimtolType] != -1} {
          ::tcom::foreach dimtolAtt [$counterEnt Attributes] {
            if {[string first "relat" [$dimtolAtt Name]] != -1} {
              set val [$dimtolAtt Value]
              #outputMsg " [$dimtolAtt Name] / [$dimtolAtt Value]  [$val Type] [$val P21ID]" green
              if {$val != ""} {getAssocGeom $val 1}
            }
          }
        }
      }

# report associated geometry
      if {[info exists assocGeom]} {
        if {[info exists spmiIDRow($ht,$spmiID)]} {
          set str [reportAssocGeom $dimtolType $spmiIDRow($ht,$spmiID)]
        } else {
          set str [reportAssocGeom $dimtolType]
        }
        set holtolGeomEnts ""

        if {$str != "" && [info exists spmiIDRow($ht,$spmiID)]} {
          if {![info exists pmiColumns(ch)]} {set pmiColumns(ch) [expr {$pmiStartCol($ht)+12}]}
          set colName "Associated Geometry[format "%c" 10](Sec. 5.1.1, 5.1.5)"
          set c [string index [cellRange 1 $pmiColumns(ch)] 0]
          set r $spmiIDRow($ht,$spmiID)
          if {![info exists pmiHeading($pmiColumns(ch))]} {
            $cells($ht) Item 3 $c $colName
            set pmiHeading($pmiColumns(ch)) 1
            set pmiCol [expr {max($pmiColumns(ch),$pmiCol)}]
            set comment "See Help > User Guide (section 5.1.5) for an explanation of Associated Geometry."
            addCellComment $ht 3 $c $comment
          }
          $cells($ht) Item $r $pmiColumns(ch) [string trim $str]

# supplemental geometry comment
          if {[string first "*" $str] != -1} {
            set comment "Geometry IDs marked with an asterisk (*) are also Supplemental Geometry.  ($recPracNames(suppgeom), Sec. 4.3, Fig. 4)"
            addCellComment $ht $r $pmiColumns(ch) $comment
            errorMsg "Some Associated Geometry associated with a Dimension is also Supplemental Geometry."
          }

# check for unexpected associated geometry for diameters and radius
          if {[info exists dimName]} {
            if {[string first "diameter" $dimName] != -1 || [string first "radius" $dimName] != -1} {
              set badGeom {}
              set okSurf 0
              foreach item {"plane" "edge_curve" "manifold_solid_brep"} {
                if {[string first $item $str] != -1} {lappend badGeom [list $dimName $item]}
                if {[string first "surface" $str] != -1} {set okSurf 1}
              }
              foreach item $badGeom {
                if {$okSurf} {
                  errorMsg "Associated Geometry for a '[lindex $item 0]' dimension also refers to '[lindex $item 1]'.  Check that this is the intended association."
                  addCellComment $ht $r $pmiColumns(ch) "[string totitle $dimName] dimension (column E) also refers to '[lindex $item 1]'.  Check that this is the intended association."
                } else {
                  errorMsg "Associated Geometry for a '[lindex $item 0]' dimension is only a '[lindex $item 1]'.  Check that this is the intended association."
                  addCellComment $ht $r $pmiColumns(ch) "[string totitle $dimName] dimension (column E) is not associated with curved surfaces.  Check that this is the intended association."
                }
                lappend entsWithErrors "dimensional_characteristic_representation"
              }
            }
          }

          foreach item [split $str "\n"] {
            if {[string first "shape_aspect" $item] == -1 && \
                [string first "advanced_face" $item] == -1 && \
                [string first "centre_of_symmetry" $item] == -1 && \
                [string first "datum_feature" $item] == -1} {lappend nstr $item}
          }
          if {[info exists nstr]} {
            set holtolGeomEnts [join [lsort $nstr]]
            set counterEntType($holtolGeomEnts) "$dimtolType $dimtolID"
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding Associated Geometry: $emsg"
    }

# -------------------------------------------------------------------------------
# report complete dimension representation (holerep)
    if {[catch {
      set cellComment ""
      if {[info exists holerep] && [info exists spmiIDRow($ht,$spmiID)]} {
        if {![info exists pmiColumns(hlrp)]} {set pmiColumns(hlrp) [expr {[[[$worksheet($ht) UsedRange] Columns] Count]+1}]}
        set c [string index [cellRange 1 $pmiColumns(hlrp)] 0]
        set r $spmiIDRow($ht,$spmiID)
        if {![info exists pmiHeading($pmiColumns(hlrp))]} {
          set colName "Hole[format "%c" 10]Tolerance"
          $cells($ht) Item 3 $c $colName
          set pmiHeading($pmiColumns(hlrp)) 1
          set pmiCol [expr {max($pmiColumns(hlrp),$pmiCol)}]
          set comment "See Help > User Guide (section 5.1.3) for an explanation of how the dimensions below are constructed."
          if {[info exists hole(unit)]} {append comment "\n\nDimension units: $hole(unit)"}
          append comment "\n\nRepetitive dimensions (e.g., 4X) might be shown for diameters and radii.  They are computed based on the number of cylindrical, spherical, and toroidal surfaces associated with a dimension (see Associated Geometry column to the right) and, depending on the CAD system, might be off by a factor of two, have the wrong value, or be missing."
          if {$nistName != ""} {
            append comment "\n\nSee the PMI Representation Summary worksheet to see how the Dimensional Tolerance below compares to the expected PMI."
          }
          addCellComment $ht 3 $c $comment
        }

# write dimension to spreadsheet
        $cells($ht) Item $r $pmiColumns(hlrp) $holerep
        if {$cellComment != ""} {
          addCellComment $ht $r $pmiColumns(hlrp) $cellComment
          if {[string first "'directed'" $cellComment] == -1 && [string first "'oriented'" $cellComment] == -1} {
            lappend entsWithErrors "dimensional_characteristic_representation"
          }
        }
        unset holeDim

# -------------------------------------------------------------------------------
# save dimension with associated geometry
        if {[info exists holtolGeomEnts]} {
          if {$holtolGeomEnts != ""} {
            if {[string first "'" $dr] == 0} {set dr [string range $dr 1 end]}
            if {[info exists holtolGeom($holtolGeomEnts)]} {
              if {[lsearch $holtolGeom($holtolGeomEnts) $dr] == -1} {lappend holtolGeom($holtolGeomEnts) $dr}
            } else {
              lappend holtolGeom($holtolGeomEnts) $dr
            }
          }
        }
      }

    } emsg]} {
      errorMsg "ERROR adding Hole Tolerance: $emsg"
    }
    set hole(name) ""
  }

  return 0
}
