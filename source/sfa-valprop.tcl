proc valPropStart {} {
  global objDesign
  global cells col entLevel ent entAttrList letters ncartpt opt pd pdcol pdheading propDefID
  global propDefIDRow propDefRow rowmax valPropEnts valPropLink valPropNames valProps

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
    [list "" [list "affected area" "affected curve length" "datum references" "equivalent unicode string" "font name" "number of annotations" \
      "number of composite tolerances" "number of datum annotations" "number of datum features" "number of datum references" "number of datum targets" \
      "number of datums" "number of dimension annotations" "number of dimensional locations" "number of dimensional sizes" "number of facets" \
      "number of geometric tolerances" "number of other annotations" "number of PMI presentation elements" "number of segments" \
      "number of semantic pmi elements" "number of semantic text notes on geometry" "number of semantic text notes on part" \
      "number of semantic text notes on PMI" "number of semantic text notes" "number of tolerance annotations" "number of views" "polyline centre point" \
      "polyline curve length" "tessellated curve centre point" "tessellated curve length" "tessellated surface area" "tessellated surface centre point"]]]

# CAx-IF RP User Defined Attributes, section 8
  set valPropNames(attribute_validation_property) [list \
    [list "" [list "boolean user attributes" "edge user attributes" "face user attributes" "geometric set user attributes" "group user attributes" \
      "instance user attributes" "integer user attributes" "measure value user attributes" "part user attributes" "pmi user attributes" \
      "real user attributes" "solid user attributes" "text user attributes" "user attribute groups" "vertex user attributes"]]]

# tessellated geometry recommended practice
  set valPropNames(tessellated_validation_property) [list \
    [list "bounding box" [list "bounding box corner point"]] \
    [list centroid [list "tessellated surface centre point" "tessellated curve centre point" "tessellated point set centre point"]] \
    [list "curve length" [list "tessellated curve length"]] \
    [list "number of facets" [list "number of facets"]] \
    [list "number of segments" [list "number of segments"]] \
    [list "surface area" [list "tessellated surface area"]]]

# composite recommended practice (new vp in the last 2 lines)
  set valPropNames(composite_validation_property) [list \
    [list "" [list "number of composite tables" "number of composite materials per part" \
      "number of orientations per part" "number of plies per part" "number of plies per laminate table" \
      "number of composite sequences per laminate table" "number of composite materials per laminate table" \
      "number of composite orientations per laminate table" "ordered sequences per laminate table" \
      "notational centroid" "number of ply pieces per ply" \
      "number of tables" "number of sequences" "number of plies" "number of materials" "number of orientations" \
      "sum of all ply surfaces areas" "centre point of all plies" "number of cores"]]]

# new composite validation properties
#"boundary length" "bounded area" "centroid for all the ply shapes in the part" "centroids for all the ply shapes in each laminate table"
#"curve centroid for rosette type guide by a curve" "curve length for rosette type guide by a curve"
#"geometric boundary (inner + outer) curve centroid (for implicit plies)" "geometric boundary (inner + outer) curve length (for implicit plies)"
#"geometric centroid" "geometric centroid (for cores and explicit plies)" "geometric surface area (for cores and explicit plies)"
#"geometric volume (for cores and explicit plies)" "notional rosette centroid (association ply/rosette)" "number of composite materials for each laminate table"
#"number of composite materials in the part" "number of composite sequences for each laminate table" "number of composite sequences in the part"
#"number of cores for each laminate table" "number of cores in the part" "number of laminates tables in the part" "number of orientations for each laminate table"
#"number of orientations in the part" "number of plies" "number of plies for each laminate table" "number of plies in each sequence" "number of plies in the part"
#"number of plies using the rosette" "number of ply pieces per ply" "number of rosette used in the part"
#"ordered (alphanumeric ascending) list of orientation names used in each laminate table" "ordered (alphanumeric ascending) list of orientation names used in part"
#"ordered (numerically ascending) list of orientation values used in each laminate table" "ordered (numerically ascending) list of orientation values used in part"
#"ordered sequences name for each laminate table" "sum of area (for exact implicit ply representation) in the part"
#"sum of area (for exact implicit ply representation) of each laminate table" "sum of ply area by material" "sum of ply surface areas for each laminate table"
#"sum of ply volume by material" "sum of the geometric boundary length of the plies using the rosette"
#"sum of volume (for core and explicit plies) of each laminate table" "sum of volume explicit plies in the part" "sum of volume for core in the part"

# FEA validation properties
  set valPropNames(FEA_validation_property) [list \
    [list "number of nodes" [list "number of nodes"]] \
    [list "number of elements" [list "number of elements"]] \
    [list "FEA model bounding box" [list "FEA model bounding box corner point"]] \
    [list "3D elements volume" [list "3D elements volume measure"]] \
    [list "2D elements area" [list "2D elements area measure"]] \
    [list "1D elements length" [list "1D elements length measure"]] \
    [list "physical model volume" [list "physical model volume measure"]] \
    [list "3D elements centroid" [list "3D elements centre point"]] \
    [list "2D elements centroid" [list "2D elements centre point"]] \
    [list "1D elements centroid" [list "1D elements centre point"]] \
    [list "total model mass" [list "total model mass measure"]] \
    [list "centre of gravity" [list "centre of gravity"]] \
    [list "number of load cases" [list "number of load cases"]] \
    [list "number of fixed DOF" [list "number of fixed DOF"]] \
    [list "FEA model resultant force of applied forces" [list "Fx resultant force measure" "Fy resultant force measure" "Fz resultant force measure"]] \
    [list "FEA model resultant moment of applied moments" [list "Mxx resultant moment of moments measure" "Myy resultant moment of moments measure" "Mzz resultant moment of moments measure"]] \
    [list "FEA model resultant moment of applied forces" [list "reference point for resultant moment of forces" "Mxx resultant moment of forces measure" "Myy resultant moment of forces measure" "Mzz resultant moment of forces measure"]] \
    [list "max nodal displacement" [list "max nodal displacement measure"]] \
    [list "max nodal rotation" [list "max nodal rotation measure"]] \
    [list "min-max volume change ratio" [list "min volume change ratio measure" "max volume change ratio measure"]] \
    [list "max Von-Mises stress" [list "max Von-Mises stress measure"]] \
    [list "min-max 2D elements membrane_X force" [list "min 2D elements membrane_X force measure" "max 2D elements membrane_X force measure"]] \
    [list "min-max 2D elements membrane_Y force" [list "min 2D elements membrane_Y force measure" "max 2D elements membrane_Y force measure"]] \
    [list "min-max 2D elements membrane_XY force" [list "min 2D elements membrane_XY force measure" "max 2D elements membrane_XY force measure"]] \
    [list "min-max 2D elements shear_X force" [list "min 2D elements shear_X force measure" "max 2D elements shear_X force measure"]] \
    [list "min-max 2D elements shear_Y force" [list "min 2D elements shear_Y force measure" "max 2D elements shear_Y force measure"]] \
    [list "min-max 2D elements bending_X moment" [list "min 2D elements bending_X moment measure" "max 2D elements bending_X moment measure"]] \
    [list "min-max 2D elements bending_Y moment" [list "min 2D elements bending_Y moment measure" "max 2D elements bending_Y moment measure"]] \
    [list "min-max 2D elements bending_XY moment" [list "min 2D elements bending_XY moment measure" "max 2D elements bending_XY moment measure"]] \
    [list "min-max 1D elements axial force" [list "min 1D elements axial force measure" "max 1D elements axial force measure"]] \
    [list "min-max 1D elements shear1 force" [list "min 1D elements shear1 force measure" "max 1D elements shear1 force measure"]] \
    [list "min-max 1D elements shear2 force" [list "min 1D elements shear2 force measure" "max 1D elements shear2 force measure"]] \
    [list "min-max 1D elements torsion moment" [list "min 1D elements torsion moment measure" "max 1D elements torsion moment measure"]] \
    [list "min-max 1D elements bending1 moment" [list "min 1D elements bending1 moment measure" "max 1D elements bending1 moment measure"]] \
    [list "min-max 1D elements bending2 moment" [list "min 1D elements bending2 moment measure" "max 1D elements bending2 moment measure"]] \
  ]

# semantic text, section 7.4.2 (not a traditional valprop)
  set valPropNames(semantic_text) [list [list "" [list ""]]]

  set derived_unit_element [list derived_unit_element unit \
    [list conversion_based_unit_and_length_unit dimensions name conversion_factor] \
    [list conversion_based_unit_and_mass_unit dimensions name conversion_factor] \
    [list conversion_based_unit_and_plane_angle_unit dimensions name conversion_factor] \
    [list length_unit_and_si_unit prefix name] exponent \
    [list mass_unit_and_si_unit prefix name] exponent \
    [list plane_angle_unit_and_si_unit prefix name] exponent \
    [list si_unit_and_time_unit prefix name] exponent]

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
    [list si_unit_and_thermodynamic_temperature_unit dimensions prefix name] \
    [list force_unit elements $derived_unit_element] \
    [list moment_unit elements $derived_unit_element]]

  set ang  [list plane_angle_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set len1 [list length_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set len2 [list length_measure_with_unit_and_measure_representation_item_and_qualified_representation_item value_component unit_component name]
  set area [list area_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set vol  [list volume_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set forc [list force_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set pres [list pressure_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set mass [list mass_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set rat  [list ratio_measure_with_unit_and_measure_representation_item value_component unit_component name]

  set rep1 [list representation name items $a2p3d $drep $vrep $brep $irep $rrep $mrep $len1 $len2 $mass $cartesian_point $ang $area $vol $forc $pres $rat]
  set rep2 [list shape_representation_with_parameters name items $a2p3d $drep $vrep $brep $irep $rrep $mrep $len1 $len2 $mass $cartesian_point $ang $area $vol $forc $pres $rat]

  set gvp [list property_definition_representation \
    definition [list property_definition name description definition] \
    used_representation $rep1 $rep2]

  set entAttrList {}
  set pd "property_definition"
  set pdcol 0
  set propDefRow {}
  set valPropLink 0
  catch {unset ent}
  catch {unset pdheading}

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

# get all property_definition_representation
  set pdr {}
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    set objType [$objEntity Type]
    if {$objType == $startent} {lappend pdr $objEntity}
  }

# process all property_definition_representation
  foreach objEntity $pdr {
    set objType [$objEntity Type]
    set ncartpt 0
    if {$n < $rowmax} {
      if {[expr {$n%2000}] == 0} {
        if {$n > 0} {outputMsg "  $n"}
        update
      }
      valPropReport $objEntity
      if {$opt(DEBUG1)} {outputMsg \n}

# write valProps and valPropEnts to worksheet cells
      if {[info exists propDefIDRow($propDefID)] && [info exists valProps]} {
        set vprow $propDefIDRow($propDefID)
        set vpadd 0
        if {[string length [[$cells($pd) Item $vprow E] Value]] > 0} {set vpadd 1}

        foreach {id c1} {0 E 1 G 2 I 3 K 4 M} {
          set c2 [string index $letters [expr {[string first $c1 $letters]+1}]]
          set vp1 [lindex $valProps $id]
          set vp2 [lindex $valPropEnts $id]
          set lvp [llength $vp1]
          set str1 ""
          set str2 ""
          set okstr 0

# add linefeeds for multiline vp
          for {set i 0} {$i < $lvp} {incr i} {
            set val1 [lindex $vp1 $i]
            set val2 [lindex $vp2 $i]
            append str1 $val1[format "%c" 10]
            append str2 $val2[format "%c" 10]
            if {[string length $val2] > 0} {set okstr 1}
          }

# write cells
          if {$okstr} {

# remove last linefeed
            set str1 [string range $str1 0 end-1]
            set str2 [string range $str2 0 end-1]

# append new str to existing cells, usually for sampling points
            if {$vpadd} {
              set val1 [[$cells($pd) Item $vprow $c1] Value]
              set str1 $val1[format "%c" 10]$str1
              set val2 [[$cells($pd) Item $vprow $c2] Value]
              set str2 $val2[format "%c" 10]$str2
            }
            if {[catch {
              $cells($pd) Item $vprow $c1 $str1
              $cells($pd) Item $vprow $c2 $str2
            } emsg]} {
              errorMsg "  Error writing validation property to cell $c1$vprow" red
            }
          }
        }
        unset valProps
        unset valPropEnts
      }
    }
    incr n
  }
  set col($pd) $pdcol
}

# -------------------------------------------------------------------------------
proc valPropReport {objEntity} {
  global cells col entLevel ent entAttrList maxelem maxrep ncartpt nelem nrep opt pd pdcol pdheading pmivalprop prefix propDefID propDefIDRow
  global propDefName propDefOK propDefRow recPracNames repName stepAP syntaxErr valName valPropEnts valPropLink valPropNames valProps

  if {$opt(DEBUG1)} {outputMsg "valPropReport" red}
  if {[info exists propDefOK]} {if {$propDefOK == 0} {return}}
  set lf "\n[string repeat " " 14]"

  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]
  set pointLimit 2

  if {[string first "handle" $objEntity] != -1} {
    if {[catch {
      set objType [$objEntity Type]
      set objID   [$objEntity P21ID]
      set ent($entLevel) [$objEntity Type]
      set objAttributes [$objEntity Attributes]

      if {$opt(DEBUG1)} {outputMsg "$ind ENT $entLevel #$objID=$objType (ATR=[$objAttributes Count])" blue}

# limit sampling points to pointLimit
      if {[info exists repName]} {
        if {$objType == "cartesian_point" && [string first "sampling points" $repName] != -1} {
          incr ncartpt
          if {$ncartpt > $pointLimit} {
            errorMsg " Only the first $pointLimit sampling points are reported."
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

      if {$entLevel == 1} {
        set pmivalprop 0
        catch {unset nelem}
        catch {unset maxelem}
        catch {unset repName}
        set valProps [list {} {} {} {} {}]
        set valPropEnts $valProps
      }

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

# missing units
            set nounits 0
            if {$objName == "unit_component"} {
              if {[string length $objValue] == 0 && \
                  ([string first "volume" $valName] == -1 || [string first "area" $valName] == -1 || [string first "length" $valName] == -1)} {
                set msg "Syntax Error: Missing 'unit_component' attribute on $ent($entLevel).  No units assigned to validation property values."
                if {$propDefName == "geometric_validation_property"} {append msg "$lf\($recPracNames(valprop))"}
                errorMsg $msg
                lappend syntaxErr($ent($entLevel)) [list $objID unit_component $msg]
                lappend syntaxErr(property_definition) [list $propDefID 9 $msg]
                set nounits 1
              }

# wrong units
              catch {
                if {[string first "length" [$objValue Type]] != -1 && \
                   ([string first "force" $valName] != -1 || [string first "moment" $valName] != -1 || [string first "mass" $valName] != -1)} {
                  set msg "Syntax Error: Bad units for the validation property value."
                  errorMsg $msg
                  lappend syntaxErr($ent($entLevel)) [list $objID unit_component $msg]
                  lappend syntaxErr(property_definition) [list $propDefID 11 $msg]
                }
              }
            }

            if {[info exists cells($pd)]} {
              set ok 0

# get values for these entity and attribute pairs
              switch -glob $ent1 {
                "*measure_representation_item* value_component" -
                "value_representation_item value_component" -
                "descriptive_representation_item description" {
                  set ok 1
                  set col($pd) 9
                  addValProps 2 $objValue "#$objID $ent2"
                  #outputMsg "  VALUE    [llength [lindex $valProps 2]]  $valProps" red
                }

                "*measure_representation_item* unit_component" {
# check for exponent (derived_unit) for area and volume
                  if {!$nounits} {
                    foreach mtype [list area volume] {
                      if {[string first $mtype $valName] != -1} {
                        set munit $mtype
                        append munit "_unit"
                        set typ [$objValue Type]
                        if {$typ != "derived_unit" && $typ != $munit} {
                          set msg "Syntax Error: Missing units exponent for a '$mtype' validation property.  '$ent2' must refer to '$munit' or 'derived_unit'."
                          errorMsg $msg
                          set vpcol 11
                          catch {if {[[$cells($pd) Item 3 13] Value] != ""} {set vpcol 13}}
                          lappend syntaxErr(property_definition) [list $propDefID $vpcol $msg]
                        }
                      }
                    }
                  }
                }

                "property_definition definition" {
                  if {[string first "validation_property" $propDefName] != -1} {
                    if {[string length $objValue] == 0} {
                      set msg "Syntax Error: Missing property_definition 'definition' attribute."
                      if {$propDefName == "geometric_validation_property"} {append msg "\n[string repeat " " 14]\($recPracNames(valprop), Sec. 4)"}
                      errorMsg $msg
                      lappend syntaxErr(property_definition) [list $propDefID 4 $msg]
                    }
                  }
                }
              }
              set colName "value"

# colName
              if {$ok && [info exists propDefID]} {
                set c [string index [cellRange 1 $col($pd)] 0]
                set r $propDefIDRow($propDefID)
                if {![info exists pdheading($col($pd))]} {
                  $cells($pd) Item 3 $c $colName
                  $cells($pd) Item 3 [string index [cellRange 1 [expr {$col($pd)+1}]] 0] "attribute"
                  set pdheading($col($pd)) 1
                }

# keep track of rows with validation properties
                if {[lsearch $propDefRow $r] == -1 && [string first "validation_property" $propDefName] != -1} {lappend propDefRow $r}

# keep track of max column
                incr col($pd)
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

# get values for these entity and attribute pairs, nrep keeps track of multiple representation items
              switch -glob $ent1 {
                "cartesian_point coordinates" -
                "direction direction_ratios"  {
                  set ok 1
                  set col($pd) 9
                  set colName "value"
                  addValProps 2 $objValue "#$objID $ent2"
                  #outputMsg "  VALUE    [llength [lindex $valProps 2]]  $valProps" red
                }

                "representation items" -
                "shape_representation_with_parameters items" {
                  set nrep 0
                  set maxrep $objSize

# add number of sampling points to representation name
                  if {[string first "sampling points" $repName] != -1} {set valProps [lreplace $valProps 0 0 [list "$repName ($maxrep)"]]}
                }

                "* elements" {set maxelem $objSize}
              }

# colName
              if {$ok && [info exists propDefID]} {
                set c [string index [cellRange 1 $col($pd)] 0]
                set r $propDefIDRow($propDefID)
                if {![info exists pdheading($col($pd))]} {
                  $cells($pd) Item 3 $c $colName
                  $cells($pd) Item 3 [string index [cellRange 1 [expr {$col($pd)+1}]] 0] "attribute"
                  set pdheading($col($pd)) 1
                }

# keep track of rows with validation properties
                if {[lsearch $propDefRow $r] == -1 && [string first "validation_property" $propDefName] != -1} {lappend propDefRow $r}

                set ov $objValue
                if {$ent1 == "cartesian_point coordinates" || $ent1 == "direction direction_ratios"} {regsub -all " " $ov "    " ov}
                incr col($pd)
                set pdcol [expr {max($col($pd),$pdcol)}]
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
                "*_representation_item the_value" -
                "descriptive_representation_item description" {
                  set ok 1
                  set col($pd) 9
                  set colName "value"
                  addValProps 2 $objValue "#$objID $ent2"
                  #outputMsg "  VALUE    [llength [lindex $valProps 2]]  $valProps" red
                }

                "*_unit_and_si_unit prefix" -
                "si_unit_and_*_unit prefix" {set ok 0; set prefix $objValue}

                "*_unit_and_si_unit name" -
                "si_unit_and_*_unit name" {
                  set ok 1
                  set col($pd) 11
                  set colName "units"
                  set objValue "$prefix$objValue"
                  addValProps 3 $objValue "#$objID $ent2"
                  #outputMsg "   UNITS    [llength [lindex $valProps 3]]  $valProps" red
                }
                "conversion_based_unit_and_*_unit name" {
                  set ok 1
                  set col($pd) 11
                  set colName "units"
                  addValProps 3 $objValue "#$objID $ent2"
                  #outputMsg "   UNITS    [llength [lindex $valProps 3]]  $valProps" red
                }

                "derived_unit_element exponent" {
                  set ok 1
                  set col($pd) 13
                  set colName "exponent"
                  addValProps 4 $objValue "#$objID $ent2"
                  #outputMsg "    EXP      [llength [lindex $valProps 4]]  $valProps" red

# wrong exponent
                  catch {
                    if {([string first "length" $valName] != -1 && $objValue != 1) || \
                        ([string first "area" $valName]   != -1 && $objValue != 2) || \
                        ([string first "volume" $valName] != -1 && $objValue != 3)} {
                      set msg "Syntax Error: Bad exponent for the validation property units"
                      errorMsg $msg
                      lappend syntaxErr($ent($entLevel)) [list $objID exponent $msg]
                      lappend syntaxErr(property_definition) [list $propDefID 13 $msg]
                    }
                  }
                }

                "property_definition name" {
                  set ok 0
                  set pmivalprop 1
                  regsub -all " " $objValue "_" propDefName

# check for valid validation property names and 'semantic text'
                  if {[string first "validation property" $objValue] != -1 || $objValue == "semantic text"} {
                    set okvp 0
                    set vps [list "geometric" "assembly" "pmi" "tessellated" "attribute" "FEA" "composite"]
                    foreach vp $vps {if {[string first $vp $objValue] == 0} {set okvp 1}}
                    if {$objValue == "semantic text"} {set okvp 1}

# bad valprop name
                    if {!$okvp} {
                      set okvp 0
                      foreach vp $vps {
                        if {[string first $vp [string tolower $objValue]] == 0 && $objValue != "FEA validation property"} {
                          set okvp 1
                          set emsg "Syntax Error: Use lower case 'property_definition' attribute 'name' ($objValue)."
                          regsub -all " " [string tolower $objValue] "_" propDefName
                          errorMsg $emsg
                          set invalid $emsg
                          lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $emsg]
                        }
                      }
                    }
                    if {!$okvp} {
                      set msg "Syntax Error: Unexpected Validation Property '$objValue'"
                      errorMsg $msg
                      set invalid $msg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                    }
                    set valPropLink 1
                  }
                }

                "representation name" -
                "shape_representation_with_parameters name" {
                  set ok 1
                  set col($pd) 5
                  set colName "representation name"
                  set repName $objValue
                  if {[string first "sampling points" $repName] != -1} {set ncardpt 0}

# add representation name to valProps
                  addValProps 0 $objValue "#$objID $ent2"
                  #outputMsg "REPNAME  [llength [lindex $valProps 0]]  $valProps" red

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
                        set emsg "Syntax Error: Bad '$ent2' attribute for '$propDefName'."
                        switch $propDefName {
                          geometric_validation_property -
                          assembly_validation_property {append emsg "$lf\($recPracNames(valprop), Sec. 8)"}
                          pmi_validation_property {
                            if {$stepAP == "AP242"} {
                              append emsg "$lf\($recPracNames(pmi242), Sec. 10)"
                            } else {
                              append emsg "$lf\($recPracNames(pmi203), Sec. 6)"
                            }
                          }
                          tessellated_validation_property {append emsg "$lf\($recPracNames(tessgeom), Sec. 8.4)"}
                          attribute_validation_property   {append emsg "$lf\($recPracNames(uda), Sec. 8)"}
                          composite_validation_property   {append emsg "$lf\($recPracNames(comp), Sec. 3)"}
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
                  addValProps 1 $objValue "#$objID $ent2"
                  #outputMsg " VALNAME  [llength [lindex $valProps 1]]  $valProps" red

# RP allows for blank representation.name (repName) except for sampling points
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
                                set emsg "Syntax Error: Bad representation 'name' attribute for '$objValue'.\n[string repeat " " 14]($recPracNames(valprop), Sec. 4.11)"
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

# do not flag cartesian_point.name errors with entity ids, or semantic text property name
                      if {!$ok1 && $ent2 == "cartesian_point.name"} {
                        if {[string first "\#" $objValue] != -1} {set ok1 1}
                      }
                      if {$propDefName == "semantic_text"} {set ok1 1}

                      if {$ok1 != 1} {
                        if {$ok1 == 0} {
                          set emsg "Syntax Error: Bad "
                        } elseif {$ok1 == 2} {
                          set emsg "Syntax Error: Use lower case for "
                        }
                        append emsg "'$ent2' attribute for '$propDefName'."
                        switch $propDefName {
                          geometric_validation_property -
                          assembly_validation_property {append emsg "$lf\($recPracNames(valprop), Sec. 8)"}
                          pmi_validation_property {
                            if {$stepAP == "AP242"} {
                              append emsg "$lf\($recPracNames(pmi242), Sec. 10)"
                            } else {
                              append emsg "$lf\($recPracNames(pmi203), Sec. 6)"
                            }
                          }
                          tessellated_validation_property {append emsg "$lf\($recPracNames(tessgeom), Sec. 8.4)"}
                          attribute_validation_property   {append emsg "$lf\($recPracNames(uda), Sec. 8)"}
                          composite_validation_property   {append emsg "$lf\($recPracNames(comp), Sec. 3)"}
                        }
                        errorMsg $emsg
                        set invalid $emsg
                        lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $emsg]
                      }
                    }
                  }
                }
              }

# colName
              if {$ok && [info exists propDefID]} {
                set c [string index [cellRange 1 $col($pd)] 0]
                set r $propDefIDRow($propDefID)
                if {![info exists pdheading($col($pd))]} {
                  $cells($pd) Item 3 $c $colName
                  $cells($pd) Item 3 [string index [cellRange 1 [expr {$col($pd)+1}]] 0] "attribute"
                  set pdheading($col($pd)) 1
                }

# keep track of rows with validation properties
                if {[lsearch $propDefRow $r] == -1 && \
                   ([string first "validation_property" $propDefName] != -1 || $propDefName == "semantic_text")} {lappend propDefRow $r}
                if {$invalid != ""} {lappend syntaxErr($pd) [list "-$r" $col($pd) $invalid]}
                incr col($pd)
                set pdcol [expr {max($col($pd),$pdcol)}]
              }
            }
          }
        }
      }

# error reading valprop
    } emsg1]} {
      set msg ""
      if {[info exists objName]} {
        if {$objName == "unit_component" && $objValue == ""} {
          set msg "Syntax Error: Missing 'unit_component' attribute on $ent($entLevel).  No units assigned to validation property values."
          errorMsg $msg
          lappend syntaxErr($ent($entLevel)) [list $objID unit_component $msg]
        }
      }
      if {$msg == ""} {
        set emsg1 [string trim $emsg1]
        if {[string length $emsg1] > 0 && [string first "can't read \"ent(" $emsg1] == -1} {errorMsg "ERROR adding Validation Properties: $emsg1"}
      }
    }
  }
  incr entLevel -1
}

# -------------------------------------------------------------------------------
proc valPropFormat {} {
  global cells col entCount excelVersion opt propDefRow row thisEntType worksheet valPropLink

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
    if {$excelVersion > 11} {
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

# colors
    if {[catch {
      set j 0
      for {set i 5} {$i <= $col($thisEntType)} {incr i 2} {
        foreach r $r3 {
          set r1 [lindex $r 0]
          set r2 [lindex $r 1]
          set range [$worksheet($thisEntType) Range [cellRange $r1 $i] [cellRange $r2 [expr {$i+1}]]]
          [$range Interior] ColorIndex [lindex [list 36 35] [expr {$j%2}]]
        }
        incr j
      }
    } emsg]} {
      errorMsg "ERROR formatting Validation Properties (colors): $emsg"
    }

# borders (inside horizontal and vertical)
    if {[catch {
      set range [$worksheet($thisEntType) Range [cellRange 4 5] \
        [cellRange [expr {[[[$worksheet($thisEntType) UsedRange] Rows] Count]+1}] [[[$worksheet($thisEntType) UsedRange] Columns] Count]]]
      [[$range Borders] Item [expr 11]] Weight [expr 1]
      [[$range Borders] Item [expr 12]] Weight [expr 1]
    } emsg]} {
      errorMsg "ERROR formatting Validation Properties (borders): $emsg"
    }

# left and right borders in header
    for {set i 5} {$i <= $col($thisEntType)} {incr i} {
      set range [$worksheet($thisEntType) Range [cellRange 3 $i] [cellRange 3 $i]]
      catch {
        [[$range Borders] Item [expr 7]]  Weight [expr 1]
        [[$range Borders] Item [expr 10]] Weight [expr 1]
      }
    }

# bold lines top and bottom
    set colrange [[[$worksheet($thisEntType) UsedRange] Columns] Count]
    set r $row($thisEntType)
    if {$opt(XL_ROWLIM) < $entCount($thisEntType) && $r > $opt(XL_ROWLIM)} {set r $opt(XL_ROWLIM)}
    set range [$worksheet($thisEntType) Range [cellRange $r 5] [cellRange $r $colrange]]
    catch {[[$range Borders] Item [expr 9]] Weight [expr -4138]}
    set range [$worksheet($thisEntType) Range [cellRange 2 5] [cellRange 2 $colrange]]
    catch {[[$range Borders] Item [expr 9]] Weight [expr -4138]}

# fix column widths
    for {set i 1} {$i <= $colrange} {incr i} {
      set val [[$cells($thisEntType) Item 3 $i] Value]
      if {$val == "value name"} {
        for {set i1 $i} {$i1 <= $colrange} {incr i1} {
          set range [$worksheet($thisEntType) Range [cellRange -1 $i1]]
          $range ColumnWidth [expr 196]
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
      [$worksheet($thisEntType) Hyperlinks] Add $anchor [join "https://www.cax-if.org/cax/cax_recommPractice.php"] [join ""] [join "Link to CAx-IF Recommended Practices"]
    }
  }
}

# -------------------------------------------------------------------------------
# add (idx=0) representation name, (1) value name, (2) value, (3) units, or (4) exponent to valProps and ents to valPropEnts
proc addValProps {idx val ent} {
  global maxelem nvp valName valstr valent

# accumulate val in valstr($idx) if maxelem > 1 for units and exponents
  if {[info exists maxelem] && $idx >= 3} {
    if {$maxelem > 1} {
      incr nvp($idx)
      append valstr($idx) "$val  "
      append valent($idx) "$ent   "
      if {$nvp($idx) == $maxelem} {
        set val [string trim $valstr($idx)]
        set ent [string trim $valent($idx)]
        unset valstr($idx)
        unset valent($idx)
        set nvp($idx) 0
      } else {
        return
      }
    }
  }

# set or append val in valProps
  setValProps $idx $val $ent

# add blank depending on the value name, units and exponents
  if {$idx == 2} {
    foreach str [list name "number of" point string] {if {[string first $str $valName] != -1} {foreach id {3 4} {setValProps $id}; break}}

# only exponent
  } elseif {$idx == 3} {
    foreach str [list degree length mass] {if {[string first $str $valName] != -1} {setValProps 4; break}}
  }
}

# -------------------------------------------------------------------------------
proc setValProps {idx {val ""} {ent ""}} {
  global valPropEnts valProps

# set
  if {[lindex $valProps $idx] == ""} {
    lset valProps $idx    [list $val]
    lset valPropEnts $idx [list $ent]

# append
  } else {
    set vn [lindex $valProps $idx]
    lappend vn $val
    lset valProps $idx $vn

    set vn [lindex $valPropEnts $idx]
    lappend vn $ent
    lset valPropEnts $idx $vn
  }
}
