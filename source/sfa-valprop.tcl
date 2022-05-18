proc valPropStart {defRep} {
  global objDesign
  global cells col entLevel ent entAttrList letters ncartpt opt pd pdcol pdheading propDefID
  global propDefIDRow propDefRow samplingPoints valPropEnts valPropLink valPropNames valProps

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
      "number of geometric tolerances" "number of linked annotations" "number of other annotations" "number of PMI presentation elements" \
      "number of segments" "number of semantic pmi elements" "number of semantic text notes on geometry" "number of semantic text notes on part" \
      "number of semantic text notes on PMI" "number of semantic text notes" "number of tolerance annotations" "number of views" "polyline centre point" \
      "polyline curve length" "saved view camera coordinates" "saved view scale" "saved view world coordinates" "tessellated curve centre point" \
      "tessellated curve length" "tessellated surface area" "tessellated surface centre point" "visible geometry curve length" \
      "visible geometry surface area"]]]

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

# composite structures validation properties recommended practice
  set valPropNames(composite_validation_property) [list \
    [list "" [list "number of cores" "number of materials" "number of orientations" "number of plies" "number of ply pieces per ply" \
      "number of rosettes" "number of sequences" "number of tables" "ordered list of orientation names" "ordered list of orientation values" \
      "ordered sequences per laminate table"]] \
    [list "centroid" [list "centre point"]] \
    [list "curve centroid" [list "curve centre point"]] \
    [list "curve length" [list "curve length measure"]] \
    [list "guide curve length" [list "curve length measure"]] \
    [list "notional rosette centroid" [list "notional centre point"]] \
    [list "number of facets" [list "number of facets"]] \
    [list "ply centroid" [list "centre point of all plies"]] \
    [list "sum of all core volumes" [list "volume measure"]] \
    [list "sum of all geometric boundary curve length" [list "curve length measure"]] \
    [list "sum of all ply surface areas" [list "surface area measure"]] \
    [list "sum of all ply volumes" [list "volume measure"]] \
    [list "surface area" [list "surface area measure"]] \
    [list "volume" [list "volume measure"]]]

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

# -------------------------------------------------------------------------------------------------
  set derived_unit_element [list derived_unit_element unit \
    [list conversion_based_unit_and_length_unit dimensions [list dimensional_exponents length_exponent] conversion_factor \
      [list length_measure_with_unit value_component unit_component [list length_unit_and_si_unit prefix name]]] \
    [list conversion_based_unit_and_mass_unit dimensions [list dimensional_exponents mass_exponent] conversion_factor \
      [list mass_measure_with_unit value_component unit_component [list mass_unit_and_si_unit prefix name]]] \
    [list conversion_based_unit_and_plane_angle_unit name conversion_factor] \
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
    [list si_unit_and_thermodynamic_temperature_unit prefix name] \
    [list force_unit elements $derived_unit_element] \
    [list moment_unit elements $derived_unit_element]]
  set rowrep [list row_representation_item name item_element $drep]

  set ang  [list plane_angle_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set len1 [list length_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set len2 [list length_measure_with_unit_and_measure_representation_item_and_qualified_representation_item value_component unit_component name]
  set area [list area_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set vol  [list volume_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set forc [list force_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set pres [list pressure_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set mass [list mass_measure_with_unit_and_measure_representation_item value_component unit_component name]
  set rat  [list ratio_measure_with_unit_and_measure_representation_item value_component unit_component name]

  set cartesian11 [lreplace $a2p3d 0 0 axis2_placement_3d_and_cartesian_11]
  set curve11 [list composite_curve_and_curve_11_and_measure_representation_item value_component unit_component name]

  set def1 [list characterized_representation_and_draughting_model name]
  set def2 [list model_geometric_view item [list camera_model_d3 name]]
  set def3 [list default_model_geometric_view item [list camera_model_d3 name]]

  set rep1 [list representation name items $a2p3d $drep $vrep $brep $irep $rrep $mrep $rowrep $len1 $len2 $mass $cartesian_point $ang $area $vol $forc $pres $rat]
  set rep2 [lreplace $rep1 0 0 shape_representation]
  set rep3 [lreplace $rep1 0 0 shape_representation_with_parameters]
  set rep4 [lreplace $rep1 0 0 tessellated_shape_representation]
  set rep5 [list ply_angle_representation name items $ang]
  set rep6 [list reinforcement_orientation_basis name items $cartesian11 $curve11]

# defRep is either property_definition_representation or shape_definition_representation
  set gvp [list $defRep definition [list property_definition name description definition $def1 $def2 $def3] used_representation $rep1 $rep2 $rep3 $rep4 $rep5 $rep6]

  set entAttrList {}
  set pd "property_definition"
  if {![info exists propDefRow]} {
    set propDefRow {}
    set pdcol 0
  }
  set valPropLink 0
  foreach var {ent pdheading samplingPoints} {if {[info exists $var]} {unset $var}}

  outputMsg " Adding Properties to property_definition worksheet" blue

  if {$opt(DEBUG1)} {outputMsg \n}
  set entLevel 0
  setEntAttrList $gvp
  if {$opt(DEBUG1)} {outputMsg "entAttrList $entAttrList"}
  if {$opt(DEBUG1)} {outputMsg \n}

  set startent [lindex $gvp 0]
  set n 0
  set entLevel 0

# get {property,shape}_definition_representation for property_definition in spreadsheet
  set pdr {}
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    set objType [$objEntity Type]
    if {$objType == $startent} {
      set p21id [[[[$objEntity Attributes] Item [expr 1]] Value] P21ID]
      if {[info exists propDefIDRow($p21id)]} {lappend pdr $objEntity}
    }
  }

# process {property,shape}_definition_representation
  foreach objEntity $pdr {
    set objType [$objEntity Type]
    set ncartpt 0
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
  set col($pd) [expr {max($col($pd),$pdcol)}]
}

# -------------------------------------------------------------------------------
proc valPropReport {objEntity} {
  global cells col convUnit defComment entLevel ent entAttrList gen maxelem maxrep ncartpt nelem nrep opt pd pdclass pdcol pdheading
  global pmivalprop prefix propDefID propDefIDRow propDefName propDefOK propDefRow recPracNames repName repNameOK samplingPoints spaces
  global stepAP syntaxErr tessCoord tessCoordName unicodeEnts unicodeString valName valPropEnts valPropLink valPropNames valProps

  if {$opt(DEBUG1)} {outputMsg "valPropReport" red}
  if {[info exists propDefOK]} {if {$propDefOK == 0} {return}}

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
          if {$ncartpt == 1 && $gen(View) && $opt(viewPart)} {errorMsg "  Processing cloud of points" green}
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
        foreach var {maxelem nelem pdclass repName} {if {[info exists $var]} {unset $var}}
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
              if {[info exists valName]} {
                if {[string length $objValue] == 0 && \
                    ([string first "volume" $valName] == -1 || [string first "area" $valName] == -1 || [string first "length" $valName] == -1)} {
                  set msg "Syntax Error: Missing 'unit_component' attribute on $ent($entLevel).  No units set for validation property values."
                  if {$propDefName == "geometric_validation_property"} {append msg "$spaces\($recPracNames(valprop))"}
                  errorMsg $msg
                  lappend syntaxErr($ent($entLevel)) [list $objID unit_component $msg]
                  lappend syntaxErr(property_definition) [list $propDefID 9 $msg]
                  set nounits 1
                }
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
# descriptive_representation_item is generally handled below with nodeType=5
                  set ok 1
                  set col($pd) 9
                  addValProps 2 $objValue "#$objID [formatComplexEnt $ent2]"
                }

                "*measure_representation_item* unit_component" {
# check for exponent (derived_unit) for area and volume
                  if {!$nounits} {

# get valName for complex entity with measure after _and_
                    if {[string first "_and_measure" $ent2] != -1} {set valName [[[$objEntity Attributes] Item [expr 3]] Value]}

                    foreach mtype [list area volume] {
                      if {[string first $mtype $valName] != -1} {
                        set munit $mtype
                        append munit "_unit"
                        set typ [$objValue Type]
                        if {$typ != "derived_unit" && $typ != $munit} {
                          set msg "Syntax Error: Missing units exponent for a '$mtype' validation property.  '[formatComplexEnt $ent2]' must refer to '$munit' or 'derived_unit'."
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
# check for missing definition
                  if {[string first "validation_property" $propDefName] != -1 || $propDefName == "semantic_text"} {
                    if {[string length $objValue] == 0} {
                      set msg "Syntax Error: Missing property_definition 'definition' attribute."
                      switch $propDefName {
                        "geometric_validation_property" {append msg "$spaces\($recPracNames(valprop), Sec. 4)"}
                        "pmi_validation_property" {append msg "$spaces\($recPracNames(pmi242), Sec. 10)"}
                        "semantic_text" {append msg "$spaces\($recPracNames(pmi242), Sec. 7.4.2)"}
                      }
                      errorMsg $msg
                      lappend syntaxErr(property_definition) [list $propDefID 4 $msg]
                    }
                  }

# add name or description attribute of entity referred to by the definition, check for unicode version
                  if {[string first "handle" $objValue] != -1} {
                    set n 0
                    ::tcom::foreach attr [$objValue Attributes] {
                      incr n
                      if {$n <= 2} {
                        set prodDefName [string trim [$attr Value]]
                        set idx "[$objValue Type],[$attr Name],[$objValue P21ID]"
                        if {[info exists unicodeString($idx)]} {set prodDefName $unicodeString($idx)}
                        if {$prodDefName != "" && [string first "handle" $prodDefName] == -1} {break}
                      }
                      if {$n == 2} {break}
                    }

                    if {$prodDefName != "" && [string first "handle" $prodDefName] == -1} {
                      set r $propDefIDRow($propDefID)
                      set val [[$cells($pd) Item $r D] Value]
                      if {[string first "$prodDefName" $val] == -1} {
                        append val "  \[$prodDefName\]"
                        $cells($pd) Item $r D $val
                        if {![info exists defComment]} {
                          addCellComment "property_definition" 3 D "Text in brackets is the 'name' or 'description' attribute of the definition entity."
                          set defComment 1
                        }
                      }
                    }
                  }
                }

                "conversion_based_unit_and_*_unit dimensions" {set convUnit [string range $objType 26 end]}
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
            if {[string first "handle" $objEntity] != -1} {
              if {$ent1 == "row_representation_item item_element"} {addValProps 2 "" ""}
              if {[catch {
                ::tcom::foreach val1 $objValue {valPropReport $val1}
              } emsg]} {
                foreach val2 $objValue {valPropReport $val2}
              }
            }
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
                  if {[string first "sampling points" $repName] == -1} {
                    addValProps 2 $objValue "#$objID $ent2"
                    if {$ent1 == "direction direction_ratios"} {
                      if {[veclen $objValue] == 0} {
                        set msg "Syntax Error: The validation property direction vector is '0 0 0'."
                        errorMsg $msg
                        lappend syntaxErr(direction) [list $objID direction_ratios $msg]
                        lappend syntaxErr(property_definition) [list $propDefID 9 $msg]
                      }
                    }
                  } else {
                    if {$ncartpt <= $pointLimit} {addValProps 2 $objValue "#$objID $ent2"}
                    if {$gen(View) && $opt(viewPart)} {append samplingPoints "[vectrim $objValue 5] "}
                  }
                }

                "representation items" -
                "shape_representation items" -
                "shape_representation_with_parameters items" -
                "tessellated_shape_representation items" -
                "reinforcement_orientation_basis items" -
                "ply_angle_representation items" {
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
            if {$ent1 != "tessellated_shape_representation items"} {
              if {$ncartpt < $pointLimit} {
                if {[catch {
                  ::tcom::foreach val1 $objValue {valPropReport $val1}
                } emsg]} {
                  foreach val2 $objValue {valPropReport $val2}
                }
              }

# handle tessellated_shape_representation.items for coordinates_list
            } else {
              ::tcom::foreach val1 $objValue {
                set id [$val1 P21ID]
                if {[$val1 Type] == "coordinates_list" && ![info exists tessCoord($id)]} {tessReadGeometry 1}
                if {[llength $tessCoord($id)] != 24} {
                  set msg "Syntax Error: Bad number of points ([expr {[llength $tessCoord($id)]/3}]) in coordinates_list for saved view validation property.$spaces\($recPracNames(pmi242), Sec. 10.2.2)"
                  errorMsg $msg
                  lappend syntaxErr(property_definition) [list $propDefID 9 $msg]
                }
                addValProps 1 $tessCoordName($id) "#$id coordinates_list.name"
                addValProps 2 $tessCoord($id) "#$id coordinates_list.position_coords"
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
                  if {[lsearch $unicodeEnts "DESCRIPTIVE_REPRESENTATION_ITEM"] != -1} {
                    set idx "descriptive_representation_item,description,$objID"
                    if {[info exists unicodeString($idx)]} {set objValue $unicodeString($idx)}
                  }
                  addValProps 2 $objValue "#$objID [formatComplexEnt $ent2]"
                }

                "*_unit_and_si_unit prefix" -
                "si_unit_and_*_unit prefix" {set ok 0; set prefix $objValue}

                "*_unit_and_si_unit name" -
                "si_unit_and_*_unit name" {
                  set ok 1
                  set col($pd) 11
                  set colName "units"
                  set objValue "$prefix$objValue"
                  addValProps 3 $objValue "#$objID [formatComplexEnt $ent2]"

# check mass recommended practice
                  if {[string first "mass" $ent1] != -1 && $objValue != "kilogram"} {
                    set msg "Syntax Error: For mass units, '$prefix' is not a valid prefix for 'gram', only 'kilo' is allowed.$spaces\($recPracNames(uda), Annex C.1)"
                    errorMsg $msg
                    lappend syntaxErr($ent($entLevel)) [list $objID prefix $msg]
                    lappend syntaxErr(property_definition) [list $propDefID 11 $msg]
                  }
                }

                "conversion_based_unit_and_*_unit name" {
                  set ok 1
                  set col($pd) 11
                  set colName "units"
                  addValProps 3 $objValue "#$objID [formatComplexEnt $ent2]"
                }

                "dimensional_exponents *_exponent" {

# check exponents
                  if {[string first "mass" $ent1] != -1 && $convUnit == "mass_unit" && $objValue != 1.} {
                    set msg "Syntax Error: For conversion based mass unit, wrong 'mass_exponent' on dimensional_exponents.$spaces\($recPracNames(uda), Annex C.4.2)"
                    errorMsg $msg
                    lappend syntaxErr($ent($entLevel)) [list $objID mass_exponent $msg]
                    lappend syntaxErr(property_definition) [list $propDefID 11 $msg]
                  } elseif {[string first "length" $ent1] != -1 && $convUnit == "length_unit" && $objValue != 1.} {
                    set msg "Syntax Error: For conversion based length unit, wrong 'length_exponent' on dimensional_exponents.$spaces\($recPracNames(uda), Annex C.4.2)"
                    errorMsg $msg
                    lappend syntaxErr($ent($entLevel)) [list $objID length_exponent $msg]
                    lappend syntaxErr(property_definition) [list $propDefID 11 $msg]
                  }
                }

                "derived_unit_element exponent" {
                  set ok 1
                  set col($pd) 13
                  set colName "exponent"
                  addValProps 4 $objValue "#$objID $ent2"

# wrong exponent
                  catch {
                    if {([string first "length" $valName] != -1 && $objValue != 1) || \
                        ([string first "area" $valName]   != -1 && $objValue != 2) || \
                        ([string first "volume" $valName] != -1 && $objValue != 3)} {
                      set msg "Syntax Error: Bad exponent for the value name and units"
                      errorMsg $msg
                      lappend syntaxErr($ent($entLevel)) [list $objID exponent $msg]
                      lappend syntaxErr(property_definition) [list $propDefID 13 $msg]
                    }
                  }
                  if {([string tolower $propDefName] == "density" || [string tolower $repName] == "density") && $objValue == 3} {
                    set msg "Syntax Error: For density, the length unit exponent should be -3"
                    errorMsg $msg
                    lappend syntaxErr($ent($entLevel)) [list $objID exponent $msg]
                    lappend syntaxErr(property_definition) [list $propDefID 13 $msg]
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
                          errorMsg "Use lower case property_definition 'name' attribute for '$objValue'."
                          regsub -all " " [string tolower $objValue] "_" propDefName
                        }
                      }
                    }
                    if {!$okvp} {
                      set msg "Syntax Error: Validation property '$objValue' is not valid."
                      errorMsg $msg
                      set invalid $msg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                    }
                    set valPropLink 1
                  }

# check general property association
                  if {[string first "validation property" $objValue] == -1 && [info exists entCount(general_property_association)]} {
                    set e0s [$objEntity GetUsedIn [string trim general_property_association] [string trim derived_definition]]

# check for gpa entity
                    set ngpa 0
                    ::tcom::foreach e0 $e0s {incr ngpa}
                    if {$ngpa == 0} {
                      set msg "Syntax Error: Missing corresponding general_property_association entity for the property_definition.$spaces\($recPracNames(uda), Sec. 5)"
                      errorMsg $msg
                      set invalid $msg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                    }
                    ::tcom::foreach e0 $e0s {
                      set e1 [[[$e0 Attributes] Item [expr 3]] Value]

# check gpa.name vs pd.name
                      if {[string first "handle" $e1] != -1} {
                        set gpname [[[$e1 Attributes] Item [expr 2]] Value]
                        if {$gpname != $objValue} {
                          set msg "Syntax Error: property_definition 'name' attribute is not the same as the associated general_property 'name' attribute.$spaces\($recPracNames(uda), Sec. 5)"
                          errorMsg $msg
                          set invalid $msg
                          lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                          lappend syntaxErr(general_property) [list [$e1 P21ID] name $msg]
                        }

# missing gp entity
                      } else {
                        set msg "Syntax Error: Missing associated general_property entity for the property_definition.$spaces\($recPracNames(uda), Sec. 5)"
                        errorMsg $msg
                        set invalid $msg
                        lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                      }
                    }
                  }

# check for classification
                  ::tcom::foreach e0 [$objEntity GetUsedIn [string trim id_attribute] [string trim identified_item]] {
                    set pdclass [list [[[$e0 Attributes] Item [expr 1]] Value] "#[$e0 P21ID] [$e0 Type].attribute_value"]
                  }
                }

                "representation name" -
                "shape_representation name" -
                "shape_representation_with_parameters name" -
                "tessellated_shape_representation name" -
                "reinforcement_orientation_basis name" -
                "ply_angle_representation name" {
                  set ok 1
                  set col($pd) 5
                  set colName "representation name"
                  set repName [string trim $objValue]
                  if {[string first "sampling points" $repName] != -1} {set ncardpt 0}

# add representation name to valProps
                  if {$objValue ==  ""} {set objValue " "}
                  addValProps 0 $objValue "#$objID $ent2"

                  if {[info exists propDefName]} {
                    if {$entLevel == 2 && [info exists valPropNames($propDefName)]} {

# look for valid representation.name in valPropNames
                      set ok1 0
                      set repNameOK 1
                      if {$repName != ""} {
                        foreach idx $valPropNames($propDefName) {
                          if {[lindex $idx 0] == $repName && [lindex $idx 0] != ""} {set ok1 1; break}
                        }
                      } else {
                        set ok1 1
                      }

                      if {!$ok1} {
                        set repNameOK 0
                        if {$propDefName != "pmi_validation_property" && $propDefName != "attribute_validation_property"} {
                          set emsg "Syntax Error: Bad '$ent2' attribute for '$propDefName'."
                        } else {
                          set emsg "Syntax Error: The [lindex $ent1 0] 'name' attribute must be empty."
                        }
                        switch $propDefName {
                          geometric_validation_property -
                          assembly_validation_property {append emsg "$spaces\($recPracNames(valprop), Sec. 8)"}
                          pmi_validation_property {
                            if {[string first "AP242" $stepAP] == 0} {
                              append emsg "$spaces\($recPracNames(pmi242), Sec. 10)"
                            } else {
                              append emsg "$spaces\($recPracNames(pmi203), Sec. 6)"
                            }
                          }
                          tessellated_validation_property {append emsg "$spaces\($recPracNames(tessgeom), Sec. 8.4)"}
                          attribute_validation_property   {append emsg "$spaces\($recPracNames(uda), Sec. 8)"}
                          composite_validation_property   {append emsg "$spaces\($recPracNames(comp), Sec. 3)"}
                        }
                        errorMsg $emsg
                        set invalid $emsg
                        lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $emsg]
                      }
                    }

# missing composite validation property name
                    if {[string first "ply" $ent1] == 0 || [string first "reinforcement" $ent1] == 0} {
                      if {[string tolower $propDefName] != "composite validation property"} {
                        set emsg "Syntax Error: property_definition 'name' attribute should be 'composite validation property'.$spaces\($recPracNames(comp), Sec. 4)"
                        errorMsg $emsg
                        lappend syntaxErr(property_definition) [list $propDefID name $emsg]
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
                  if {$objValue != "sampling point" || $ncartpt < 3} {addValProps 1 $objValue "#$objID [formatComplexEnt $ent2]"}

# RP allows for blank representation.name (repName) except for sampling points
                  if {[info exists propDefName]} {
                    if {$entLevel == 3 && [info exists valPropNames($propDefName)]} {
                      set ok1 0
                      if {$repNameOK} {
                        foreach idx $valPropNames($propDefName) {
                          if {[lindex $idx 0] == $repName || $repName == ""} {
                            foreach item [lindex $idx 1] {
                              if {$objValue == $item} {
                                set ok1 1
                                if {$objValue == "sampling point" && $repName == ""} {
                                  set emsg "Syntax Error: Bad representation 'name' attribute for '$objValue'.$spaces\($recPracNames(valprop), Sec. 4.11)"
                                  errorMsg $emsg
                                }
                                break
                              } elseif {[string tolower $objValue] == $item} {
                                errorMsg "Use lower case for [lindex $ent1 0] name attribute '$objValue'."
                                set ok1 1
                                break
                              }
                            }
                          }
                        }
                      } else {
                        set ok1 1
                      }

                      if {!$ok1 && $propDefName != "semantic_text"} {
                        set emsg "Syntax Error: Bad '[formatComplexEnt $ent2]' attribute for '$propDefName'."
                        switch $propDefName {
                          geometric_validation_property -
                          assembly_validation_property {append emsg "$spaces\($recPracNames(valprop), Sec. 8)"}
                          pmi_validation_property {
                            if {[string first "AP242" $stepAP] == 0} {
                              append emsg "$spaces\($recPracNames(pmi242), Sec. 10)"
                            } else {
                              append emsg "$spaces\($recPracNames(pmi203), Sec. 6)"
                            }
                          }
                          tessellated_validation_property {append emsg "$spaces\($recPracNames(tessgeom), Sec. 8.4)"}
                          attribute_validation_property   {append emsg "$spaces\($recPracNames(uda), Sec. 8)"}
                          composite_validation_property   {append emsg "$spaces\($recPracNames(comp), Sec. 4)"}
                        }
                        errorMsg $emsg
                        set invalid $emsg
                        lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $emsg]
                      }
                    }
                  }
                }

                "camera_model_d3 name" -
                "characterized_representation_and_draughting_model name" {
# saved view name
                  set ok 1
                  set col($pd) 15
                  set colName "saved view"
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
                  if {$colName == "saved view"} {addCellComment "property_definition" 3 $c "Saved View validation properties are defined in $recPracNames(pmi242), Sec. 10.2.2"}
                }

# saved view
                if {$colName == "saved view"} {
                  if {$objValue != ""} {
                    $cells($pd) Item $r $c $objValue
                    $cells($pd) Item $r [expr {$col($pd)+1}] "#$objID [formatComplexEnt $ent2]"
                  }
                }

# classification
                if {[info exists pdclass]} {
                  set col($pd) 17
                  set colName "classification"
                  set c [string index [cellRange 1 $col($pd)] 0]
                  if {![info exists pdheading($col($pd))]} {
                    $cells($pd) Item 3 $c $colName
                    $cells($pd) Item 3 [string index [cellRange 1 [expr {$col($pd)+1}]] 0] "attribute"
                    set pdheading($col($pd)) 1
                  }
                  $cells($pd) Item $r $c [lindex $pdclass 0]
                  $cells($pd) Item $r [expr {$col($pd)+1}] [lindex $pdclass 1]
                }

# keep track of rows with validation properties
                if {[lsearch $propDefRow $r] == -1 && \
                   ([string first "validation_property" $propDefName] != -1 || $propDefName == "semantic_text")} {lappend propDefRow $r}
                if {[string first "ply" $ent1] == 0 || [string first "reinforcement" $ent1] == 0} {
                  if {[lsearch $propDefRow $r] == -1} {lappend propDefRow $r}
                }

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
          set msg "Syntax Error: Missing 'unit_component' attribute on [formatComplexEnt $ent($entLevel)].  No units set for validation property values."
          errorMsg $msg
          lappend syntaxErr($ent($entLevel)) [list $objID unit_component $msg]
        }
      }
      if {$msg == ""} {
        set emsg1 [string trim $emsg1]
        if {[string length $emsg1] > 0 && [string first "can't read \"ent(" $emsg1] == -1} {errorMsg "Error adding Validation Properties: $emsg1"}
      }
    }
  }
  incr entLevel -1
}

# -------------------------------------------------------------------------------
proc valPropFormat {} {
  global cells col entRows propDefRow row thisEntType worksheet valPropLink

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
    catch {
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
      errorMsg "Error formatting Validation Properties 1: $emsg"
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
      errorMsg "Error formatting Validation Properties (colors): $emsg"
    }

# borders (inside horizontal and vertical)
    if {[catch {
      set range [$worksheet($thisEntType) Range [cellRange 4 5] \
        [cellRange [expr {[[[$worksheet($thisEntType) UsedRange] Rows] Count]+1}] [[[$worksheet($thisEntType) UsedRange] Columns] Count]]]
      [[$range Borders] Item [expr 11]] Weight [expr 1]
      [[$range Borders] Item [expr 12]] Weight [expr 1]
    } emsg]} {
      errorMsg "Error formatting Validation Properties (borders): $emsg"
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
    set r $entRows($thisEntType)
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
  if {[info exists valName]} {
    if {$idx == 2} {
      foreach str [list name "number of" "datum references" point string] {if {[string first $str $valName] != -1} {foreach id {3 4} {setValProps $id}; break}}

# only exponent
    } elseif {$idx == 3} {
      foreach str [list degree length mass] {if {[string first $str $valName] != -1} {setValProps 4; break}}
    }
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

# -------------------------------------------------------------------------------
# get validation properties
proc getValProps {} {
  global gpmiValProp propDefIDs spmiValProp unicodeString
  global objDesign

# get validation properties association to PMI, etc.
  foreach var {gpmiValProp propDefIDs spmiValProp} {if {[info exists $var]} {unset $var}}

  if {[catch {

# get all property_definitions attributes
    ::tcom::foreach e0 [$objDesign FindObjects [string trim property_definition]] {
      if {[$e0 Type] == "property_definition"} {
        set a0s [$e0 Attributes]
        set pid ""

# check name and definition attributes
        ::tcom::foreach a0 $a0s {
          set a0val  [$a0 Value]
          switch -- [$a0 Name] {
            name {
# property_definition name attribute
              set vpname $a0val
              if {([string first "validation property" $vpname] != -1 && [string first "geometric" $vpname] == -1 && [string first "tessellated" $vpname] == -1) || \
                $vpname == "semantic text"} {set pid [$e0 P21ID]}
            }

            definition {
# property_definition definition attribute
              if {$pid != "" && [string first "handle" $a0val] != -1} {

# get names of validation properties, add to vpname
                set names ""
                foreach defRep [list property_definition_representation shape_definition_representation] {
                  set e1s [$e0 GetUsedIn [string trim $defRep] [string trim definition]]
                  ::tcom::foreach e1 $e1s {
                    set e2  [[[$e1 Attributes] Item [expr 2]] Value]
                    set e3s [[[$e2 Attributes] Item [expr 2]] Value]
                    ::tcom::foreach e3 $e3s {
                      set a3s [$e3 Attributes]
                      ::tcom::foreach a3 $a3s {
                        if {[$a3 Name] == "name" && $vpname != "semantic text"} {
                          set name [$a3 Value]
                          if {$name != "" && [string first $name $names] == -1} {append names "$name, "}

# semantic text
                        } elseif {[$a3 Name] == "description" && $vpname == "semantic text" && $defRep == "property_definition_representation"} {
                          set name [$a3 Value]
                          if {$name != ""} {
                            set idx "descriptive_representation_item,description,[$e3 P21ID]"
                            if {[info exists unicodeString($idx)]} {set name $unicodeString($idx)}
                            append names $name
                          }
                        }
                      }
                    }
                  }
                }

# set propDefIDs
                if {$names != ""} {
                  if {$vpname != "semantic text"} {
                    append vpname " - [string range $names 0 end-2]"
                  } else {
                    append vpname "[format "%c" 10]$names"
                  }
                }
                set a0id [$a0val P21ID]
                if {![info exists propDefIDs($a0id)]} {
                  set propDefIDs($a0id) [list $pid $vpname]
                } else {
                  set propDefIDs($a0id) [concat $propDefIDs($a0id) [list $pid $vpname]]
                }
              }
            }
          }
        }
      }
    }
  } emsg]} {
    errorMsg "Error getting PMI validation properities: $emsg"
    catch {raise .}
  }
}

# -------------------------------------------------------------------------------
# report validation properties to some worksheets that are not associated with any PMI analysis
proc reportValProps {} {
  global cells col entCount opt pmiCol pmiHeading pmiStartCol propDefIDs vpEnts vpmiRow worksheet

# entities to check
  set vpEnts {}
  set vpCheck [list characterized_representation_and_draughting_model characterized_representation_and_draughting_model_and_tessellated_shape_representation \
                composite_group_shape_aspect datum model_geometric_view default_model_geometric_view shape_aspect composite_assembly_sequence_definition \
                composite_assembly_table ply_laminate_table ply_laminate_sequence_definition reinforcement_orientation_basis]

  foreach ent $vpCheck {
    if {[info exists entCount($ent)] && [info exists worksheet($ent)]} {
      if {$entCount($ent) > 0} {
        set c [getNextUnusedColumn $ent]
        set pmiCol $c
        set pmiStartCol($ent) $c
        set col($ent) $c

        set vpmiRow($ent) {}
        set r1 [expr {[[[$worksheet($ent) UsedRange] Rows] Count]+2}]

# read ids in column 1
        for {set r 4} {$r <= $r1} {incr r} {
          set id [string range [[$cells($ent) Item $r 1] Value] 0 end-2]

# check for a vp
          if {[info exists propDefIDs($id)]} {
            if {![info exists pmiHeading($ent$c)]} {
              set heading "Validation Properties[format "%c" 10](Sec. 10.3)"
              $cells($ent) Item 3 $c $heading
              set pmiHeading($ent$c) 1
              if {$opt(valProp)} {
                set comment "See the property_definition worksheet for validation property values."
              } else {
                set comment "Select 'Validation Properties' on the Options tab to see the validation property values."
              }
              addCellComment $ent 3 $c $comment
            }

# add to new column
            set str ""
            set llen [llength $propDefIDs($id)]
            if {$llen > 2} {append str ([expr {$llen/2}])}
            append str "property definition"
            for {set i 0} {$i < $llen} {incr i 2} {append str " [lindex $propDefIDs($id) $i]"}
            for {set i 1} {$i < $llen} {incr i 2} {append str "[format "%c" 10]([lindex $propDefIDs($id) $i])"}
            $cells($ent) Item $r $c [string trim $str]

            if {[lsearch $vpmiRow($ent) $r] == -1} {lappend vpmiRow($ent) $r}
            if {[lsearch $vpEnts $ent] == -1} {lappend vpEnts $ent}
          }
        }
      }
    }
  }

  if {[llength $vpEnts] > 0} {
    set str ""
    foreach ent $vpEnts {append str " [formatComplexEnt $ent]"}
    outputMsg "\nAdding Validation Properties to:$str" blue
  }
}

# -------------------------------------------------------------------------------
# add column of validation properties to worksheet
proc valPropColumn {ent r c propID} {
  global cells opt pmiCol pmiHeading

  if {[catch {
    if {![info exists pmiHeading($c)]} {
      set heading "Validation Properties[format "%c" 10](Sec. 10.3)"
      $cells($ent) Item 3 $c $heading
      set pmiHeading($c) 1
      set pmiCol [expr {max($c,$pmiCol)}]
      if {$opt(valProp)} {
        set comment "See the property_definition worksheet for validation property values."
      } else {
        set comment "Select 'Validation Properties' on the Options tab to see the validation property values."
      }
      addCellComment $ent 3 $c $comment
    }
    set str ""
    set llen [llength $propID]
    if {$llen > 2} {append str "([expr {$llen/2}]) "}
    append str "property definition"
    for {set i 0} {$i < $llen} {incr i 2} {append str " [lindex $propID $i]"}
    for {set i 1} {$i < $llen} {incr i 2} {append str "[format "%c" 10]([lindex $propID $i])"}
    $cells($ent) Item $r $c [string trim $str]
  } emsg]} {
    errorMsg "Error adding PMI representation validation properties: $emsg"
  }
}
