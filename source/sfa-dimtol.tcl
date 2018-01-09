proc spmiDimtolStart {entType} {
  global objDesign
  global cells col dt entLevel ent entAttrList gtEntity lastEnt opt pmiCol pmiHeading pmiStartCol spmiRow stepAP

  if {$opt(DEBUG1)} {outputMsg "START spmiDimtolStart $entType" red}

# dimensional size, location
  set dim_size      [list dimensional_size applies_to name]
  set dim_size_wdf  [list dimensional_size_with_datum_feature applies_to name]
  set dim_size_wdf1 [list composite_unit_shape_aspect_and_dimensional_size_with_datum_feature applies_to name]
  set dim_size_wdf2 [list composite_unit_shape_aspect_and_dimensional_size_with_datum_feature_and_symmetric_shape_aspect applies_to name]
  set dim_loc       [list dimensional_location name]
  set dim_loc_dir   [list directed_dimensional_location name]
  set dim_loc_wdf   [list dimensional_location_with_datum_feature name]
  set ang_loc       [list angular_location name angle_selection]
  set ang_loc1      [list angular_location_and_directed_dimensional_location name angle_selection]
  set ang_size      [list angular_size applies_to name angle_selection]
  set ang_size1     [list angular_size_and_dimensional_size_with_datum_feature applies_to name angle_selection]

  set dir           [list direction direction_ratios]
  set a2p3d         [list axis2_placement_3d name axis $dir ref_direction $dir]

# measures
  set qualifier1      [list type_qualifier name]
  set qualifier2      [list value_format_type_qualifier format_type]
  
  set length_measure1 [list length_measure_with_unit value_component]
  set length_measure2 [list length_measure_with_unit_and_measure_representation_item value_component name]
  set length_measure3 [list length_measure_with_unit_and_measure_representation_item_and_qualified_representation_item value_component name qualifiers $qualifier1 $qualifier2]

  set angle_measure1  [list plane_angle_measure_with_unit value_component]
  set angle_measure2  [list measure_representation_item_and_plane_angle_measure_with_unit name value_component]
  set angle_measure3  [list measure_representation_item_and_plane_angle_measure_with_unit_and_qualified_representation_item name value_component qualifiers $qualifier1 $qualifier2]

  set value_range     [list value_range name item_element $length_measure3 $length_measure2 $length_measure1]
  set measure_rep     [list measure_representation_item_and_qualified_representation_item name value_component qualifiers $qualifier1 $qualifier2]
  set descript_rep    [list descriptive_representation_item name description]
  set compound_rep    [list compound_representation_item item_element $descript_rep]
  
  set PMIP(dimensional_characteristic_representation) \
    [list dimensional_characteristic_representation \
      dimension $dim_size $dim_size_wdf $dim_size_wdf1 $dim_size_wdf2 $dim_loc $dim_loc_wdf $dim_loc_dir $ang_loc $ang_loc1 $ang_size $ang_size1 \
      representation [list shape_dimension_representation name \
        items $length_measure1 $length_measure2 $length_measure3 \
        $angle_measure1 $angle_measure2 $angle_measure3 \
        $value_range $measure_rep $descript_rep $compound_rep \
        $a2p3d \
      ]
    ]  
   
  if {![info exists PMIP($entType)]} {return}

  set dt $entType
  set lastEnt {}
  set entAttrList {}
  set pmiCol 0
  set spmiRow($dt) {}

  catch {unset pmiHeading}
  catch {unset ent}
  catch {unset gtEntity}

  outputMsg " Adding PMI Representation Report" blue
  
  if {[string first "AP203" $stepAP] == 0 || [string first "AP214" $stepAP] == 0} {
    errorMsg "Syntax Error: There is no Recommended Practice for PMI Representation in $stepAP files.  Use AP242 for PMI Representation."
  }

  if {$opt(DEBUG1)} {outputMsg \n}
  set entLevel 0
  setEntAttrList $PMIP($dt)
  if {$opt(DEBUG1)} {outputMsg "entAttrList $entAttrList"}
  if {$opt(DEBUG1)} {outputMsg \n}
    
  set startent [lindex $PMIP($dt) 0]
  set n 0
  set entLevel 0
  
# get next unused column by checking if there is a colName
  set pmiStartCol($dt) [expr {[getNextUnusedColumn $startent]+1}]
  #outputMsg pmiStartCol$pmiStartCol($dt)

# process all, call spmiDimtolReport
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    if {[$objEntity Type] == $startent} {
      if {$n < 1048576} {
        if {[expr {$n%2000}] == 0} {
          if {$n > 0} {outputMsg "  $n"}
          update
        }
        spmiDimtolReport $objEntity
        if {$opt(DEBUG1)} {outputMsg \n}
      }
      incr n
    }
  }
  set col($dt) $pmiCol
  set pmiStartCol($dt) [expr {$pmiStartCol($dt)-1}]
}

# -------------------------------------------------------------------------------

proc spmiDimtolReport {objEntity} {
  global assocGeom badAttributes cells col developer dim dimBasic dimRepeat dimDirected dimName dimModNames dimOrient dimReference dimrep dimrepID
  global dimSizeNames dimtolEnt dimtolEntType dimtolGeom dimval draftModelCameras dt dtpmivalprop entLevel ent entAttrList entCount entlevel2 entsWithErrors
  global incrcol lastAttr lastEnt nistName opt pmiCol pmiColumns pmiHeading pmiModifiers pmiStartCol
  global pmiUnicode prefix angDegree recPracNames savedModifier spmiEnts spmiID spmiIDRow spmiRow spmiTypesPerFile syntaxErr tolStandard
  global numDSnames dimtolComment

  if {$opt(DEBUG1)} {outputMsg "spmiDimtolReport" red}

# entLevel is very important, keeps track level of entity in hierarchy
  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]
  
  if {[string first "handle" $objEntity] != -1} {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set objAttributes [$objEntity Attributes]
    set ent($entLevel) $objType

    #outputMsg "  [$objEntity Type]" red
    #if {$entLevel == 1} {outputMsg "#$objID=$objType" blue}
    if {$opt(DEBUG1) && $objType != "cartesian_point"} {outputMsg "$ind ENT $entLevel #$objID=$objType (ATR=[$objAttributes Count])" blue}

# do not follow referred entity when the entity refers to itself, results in infinite loop, unusual case with dimensional_*_and_datum_feature
    set follow 1
    if {[info exists lastEnt]} {
      if {$lastEnt == "$objID $objType"} {set follow 0}
    }
    set lastEnt "$objID $objType"
    
    if {$entLevel == 1} {
      catch {unset dimtolEnt}
      catch {unset entlevel2}
      catch {unset assocGeom}
      set numDSnames 0
    } elseif {$entLevel == 2} {
      if {![info exists entlevel2]} {set entlevel2 [list $objID $objType]}
    }

# check if there are rows with dt
    if {$spmiEnts($objType) && [string first "datum_feature" $objType] == -1 && [string first "datum_target" $objType] == -1} {
      set spmiID $objID
      if {![info exists spmiIDRow($dt,$spmiID)]} {
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

      if {$okattr} {
        set objValue [$objAttribute Value]
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
    
              if {[info exists cells($dt)]} {
                set ok 0

# get values for these entity and attribute pairs
                switch -glob $ent1 {
# length/angle value, add to dimrep
                  "*length_measure_with_unit* value_component" -
                  "*plane_angle_measure_with_unit* value_component" {
                    set ok 1
                    set invalid ""
                    set col($dt) [expr {$pmiStartCol($dt)+2}]
                    set colName "length/angle[format "%c" 10](Sec. 5.2.1)"
                    set dimval $objValue
                    incr dim(idx)
                    catch {unset dim(qual)}
# get units
                    ::tcom::foreach attr $objAttributes {
                      if {[$attr Name] == "unit_component"} {
                        ::tcom::foreach attr1 [[$attr Value] Attributes] {
                          set val [string toupper [$attr1 Value]]
                          if {[string first "INCH" $val] != -1 } {
                            if {$val != "INCH"} {errorMsg "Syntax Error: Use 'INCH' instead of '[$attr1 Value]' to specify inches on conversion_based_unit."}
                            set dim(unit) "INCH"
                            errorMsg " Dimension units: $dim(unit)" red
                            break
                          } elseif {$val == "MILLI"} {
                            set dim(unit) "MM"
                            errorMsg " Dimension units: $dim(unit)" red
                            break
                          }
                        }
# get name
                      } elseif {[$attr Name] == "name"} {
                        set dim(name) [$attr Value]

# get qualifier (in the form of NR2 x.y from ASN.1, ISO 6093), format dimension
                      } elseif {[$attr Name] == "qualifiers"} {
                        foreach ent2 [$attr Value] {
                          if {[$ent2 Type] == "value_format_type_qualifier"} {
                            ::tcom::foreach attr1 [$ent2 Attributes] {
                              set tmp [split [lindex [split [$attr1 Value] " "] 1] "."]
                              set prec1 [expr {abs([lindex $tmp 0])}]
                              set dim(qual) [lindex $tmp 1]
                              
                              if {$prec1 != 0 || $dim(qual) != 0} {
                                if {$dim(qual) > 5} {
                                  set msg "Syntax Error: value_format_type_qualifier ([$attr1 Value]) might have too many decimal places\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.4)"
                                  errorMsg $msg
                                  lappend syntaxErr([$ent2 Type]) [list [$ent2 P21ID] "format_type" $msg]
                                  lappend syntaxErr(dimensional_characteristic_representation) [list "-$spmiIDRow($dt,$spmiID)" "decimal places" $msg]
                                }
                                set objValue [string trimright [format "%.4f" $objValue] "0"]
                                set val1 [lindex [split $objValue "."] 0]
                                set val2 [lindex [split $objValue "."] 1]
                                append val2 "0000"
                                
# format for precision
                                if {$prec1 != 0} {
                                  set dec [string range $val2 0 $dim(qual)-1]
                                  if {$dec != ""} {
                                    set dimtmp "$val1.$dec"
                                  } else {
                                    set dimtmp $val1
                                  }

# problems with NR2 values relative to dimension
                                  if {[string length $val1] > $prec1} {
                                    set msg "Syntax Error: value_format_type_qualifier ([$attr1 Value]) too small for the length/angle: $objValue\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.4)"
                                    errorMsg $msg
                                    lappend syntaxErr([$ent2 Type]) [list [$ent2 P21ID] "format_type" $msg]
                                    lappend syntaxErr(dimensional_characteristic_representation) [list "-$spmiIDRow($dt,$spmiID)" "decimal places" $msg]
                                  }
                                  if {[info exists dim(unit)]} {
                                    if {$dim(unit) == "INCH"} {
                                      if {$objValue < 1. && $prec1 > 0} {
                                        set msg "Syntax Error: For INCH units and Dimensions < 1 (no leading zero), value_format_type_qualifier 'NR2 1.n' should be 'NR2 0.n'\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.4)"
                                        errorMsg $msg
                                        lappend syntaxErr([$ent2 Type]) [list [$ent2 P21ID] "format_type" $msg]
                                        lappend syntaxErr(dimensional_characteristic_representation) [list "-$spmiIDRow($dt,$spmiID)" "decimal places" $msg]
                                      }
                                    }
                                  }
                                } else {
                                  set dimtmp ".[string range $val2 0 $dim(qual)-1]"
                                }

# bad NR2 value
                              } else {
                                set msg "Syntax Error: Invalid value_format_type_qualifier ([$attr1 Value])\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.4)"
                                errorMsg $msg
                                lappend syntaxErr([$ent2 Type]) [list [$ent2 P21ID] "format_type" $msg]
                                lappend syntaxErr(dimensional_characteristic_representation) [list "-$spmiIDRow($dt,$spmiID)" "decimal places" $msg]
                                set dimtmp $objValue
                              }
                              
# problems with NR2 values relative to dimension
                              if {$dimtmp == 0 && $objValue != 0} {
                                set msg "Syntax Error: value_format_type_qualifier ([$attr1 Value]) too small for the length/angle: $objValue\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.4)"
                                errorMsg $msg
                                lappend syntaxErr([$ent2 Type]) [list [$ent2 P21ID] "format_type" $msg]
                                lappend syntaxErr(dimensional_characteristic_representation) [list "-$spmiIDRow($dt,$spmiID)" "decimal places" $msg]
                              }
                            }
                          }
                        }
                      }
                    }
                    
                    if {[info exists dim(unit)]} {

# fix leading and trailing zeros depending on units
                      if {![info exists dim(qual)]} {
                        if {$dim(unit) == "INCH" && $objValue < 1.} {set objValue [string range $objValue 1 end]}
                        set objValue [removeTrailingZero $objValue]
                      }
                      
# check units of NIST models
                      if {$nistName != ""} {
                        set ln $nistName
                        if {$dim(unit) == "MM" && ([string first "ctc_03" $ln] != -1 || [string first "ctc_05" $ln] != -1 || \
                                                   [string first "ftc_06" $ln] != -1 || [string first "ftc_07" $ln] != -1 || [string first "ftc_08" $ln] != -1 || [string first "ftc_09" $ln] != -1)} {
                          errorMsg " INCH dimensions are used in the NIST [string toupper [string range $ln 5 end]] test case."
                        }
                        if {$dim(unit) == "INCH" && ([string first "ctc_01" $ln] != -1 || [string first "ctc_02" $ln] != -1 || [string first "ctc_04" $ln] != -1 || \
                                                     [string first "ftc_10" $ln] != -1 || [string first "ftc_11" $ln] != -1)} {
                          errorMsg " MM dimensions are used in the NIST [string toupper [string range $ln 5 end]] test case."
                        }
                      }
                    }

# get dimension precision, possibly truncate
                    if {![info exists dim(qual)]} {
                      set dim(prec) [getPrecision $objValue]
                      if {$dim(prec) > 4} {
                        set objValue [string trimright [format "%.4f" $objValue] "0"]
                        set dim(prec) [getPrecision $objValue]
                      } elseif {$dim(prec) > $dim(prec,max) && $dim(prec) < 6} {
                        set dim(prec,max) $dim(prec)
                      }
                    } else {
                      set dim(prec) $dim(qual)
                      set dim(prec,max) $dim(prec)
                    }

                    if {[info exists dimrepID]} {
                      set dim(prec,$dimrepID) $dim(prec)
                      set tmp [lindex [split $dim(name) " "] 0]
                      if {![info exists dim(qual)]} {
                        set dim($tmp) $objValue
                      } else {
                        set dim($tmp) $dimtmp
                      }

# nominal value or something other than limit
                      if {[string first "limit" $dim(name)] == -1} {
                        append dimrep($dimrepID) $dim($tmp)

# limit dimensions
                      } else {
                        if {$dim(num) > 1 && [info exists dim(upper)] && [info exist dim(lower)]} {
                          if {$tolStandard(type) == "ISO"} {
                            set dimrep($dimrepID) "$dim(symbol)$dim(upper)[format "%c" 10]$dim(symbol)$dim(lower)"
                          } else {
                            set dimrep($dimrepID) "$dim(symbol)$dim(lower)-$dim(upper)"
                          }
                          if {([info exists dim(nominal)] && $dim(upper) < $dim(nominal)) || $dim(upper) < $dim(lower)} {
                            set msg "Syntax Error: Upper limit ($dim(upper)) < 'nominal value' or 'lower limit'\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.2.4)"
                            errorMsg $msg
                            set invalid $msg
                          }
                          if {$dim(upper) == $dim(lower)} {
                            set msg "Syntax Error: Upper limit ($dim(upper)) = 'lower limit'\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.2.4)"
                            errorMsg $msg
                            set invalid $msg
                         }
                          catch {unset dim(nominal)}
                          unset dim(lower)
                          unset dim(upper)
                        }
                      }

# add degree symbol for an angle
                      set dim(angle) 0
                      if {[string first "angular_" [lindex $entlevel2 1]] != -1} {
                        if {[string index $dimrep($dimrepID) end] != $pmiUnicode(degree) && $angDegree} {
                          append dimrep($dimrepID) $pmiUnicode(degree)
                        }
                        set dim(angle) 1
                      } 
                    }
                  }
                }
    
                if {$ok && [info exists spmiID] && [info exists spmiIDRow($dt,$spmiID)]} {
                  set c [string index [cellRange 1 $col($dt)] 0]
                  set r $spmiIDRow($dt,$spmiID)

# column name (length/angle)
                  if {![info exists pmiHeading($col($dt))]} {
                    $cells($dt) Item 3 $c $colName
                    set pmiHeading($col($dt)) 1
                  }

# value in spreadsheet, original dimval
                  set val [[$cells($dt) Item $r $c] Value]
                  if {$val == ""} {
                    $cells($dt) Item $r $c "'$dimval"
                  } else {
                    $cells($dt) Item $r $c "$val[format "%c" 10]$dimval"
                  }
                  if {$invalid != ""} {
                    if {$colName != ""} {
                      lappend syntaxErr($dt) [list "-$r" $colName $invalid]
                    } else {
                      lappend syntaxErr($dt) [list $r $col($dt) $invalid]
                    }
                  }

# keep track of max column
                  set pmiCol [expr {max($col($dt),$pmiCol)}]
                  
# check that length or plane_angle are used
                } elseif {[string first "value_component" $ent1] != -1} {
                  set dte [string range [$dimtolEnt Type] 0 2]
                  if {$dte == "dim" && [string first "length" $ent1] == -1} {
                    set msg "Syntax Error: Dimension value incorrectly specified with '[formatComplexEnt [lindex [split $ent1 " "] 0]]' instead of 'length_measure_with_unit'."
                    append msg "\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.2.1)"
                    errorMsg $msg
                    lappend syntaxErr(dimensional_characteristic_representation) [list "-$r" "length/angle" $msg]
                  } elseif {$dte == "ang" && [string first "plane_angle" $ent1] == -1} {
                    set msg "Syntax Error: Angle value incorrectly specified with '[formatComplexEnt [lindex [split $ent1 " "] 0]]' instead of 'plane_angle_measure_with_unit'."
                    lappend syntaxErr(dimensional_characteristic_representation) [list "-$r" "length/angle" $msg]
                  }
                }
              }

# if referred to another, get the entity
              if {$follow} {
                if {[string first "handle" $objValue] != -1} {
                  if {[catch {
                    [$objValue Type]
                    set errstat [spmiDimtolReport $objValue]
                    if {$errstat} {break}
                  } emsg1]} {

# referred entity is actually a list of entities
                    if {[catch {
                      ::tcom::foreach val1 $objValue {spmiDimtolReport $val1}
                    } emsg2]} {
                      foreach val2 $objValue {spmiDimtolReport $val2}
                    }
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing Dimtol ($objNodeType $ent2)\n $emsg3"
            set entLevel 1
          }

# --------------
# nodeType = 20
        } elseif {$objNodeType == 20} {
          if {$idx != -1} {
            if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType-[$objAttribute AggNodeType], $objSize, $objAttrType)"}

# get number of dimensions
            if {$objType == "shape_dimension_representation" && $objName == "items"} {
              set dim(num) $objSize
              set dim(idx) 0
              if {$objSize == 0} {
                set msg "Syntax Error: Missing reference to dimension for shape_dimension_representation.items\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.2.1, Figure 15)"
                errorMsg $msg
                lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                lappend syntaxErr(dimensional_characteristic_representation) [list "-$spmiIDRow($dt,$spmiID)" "length/angle" $msg]
              }
            }
            
# no values to get from this nodetype, but get the entities that are referred to
            if {$follow} {
              if {[catch {
                set ok 0
                ::tcom::foreach val1 $objValue {
                  if {[string first "length" [$val1 Type]] != -1 || [string first "angle" [$val1 Type]] != -1} {set ok 1}
                }
                if {!$ok} {
                  set msg "Syntax Error: Missing reference to dimension length or angle measure for shape_dimension_representation.items\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.2.1, Figure 15)"
                  errorMsg $msg
                  lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                  lappend syntaxErr(dimensional_characteristic_representation) [list "-$spmiIDRow($dt,$spmiID)" "length/angle" $msg]
                }
                ::tcom::foreach val1 $objValue {spmiDimtolReport $val1}
              } emsg]} {
                foreach val2 $objValue {spmiDimtolReport $val2}
              }
            }
          }

# ---------------------
# nodeType = 5 (!= 18,19,20)
        } else {
          if {[catch {
            if {$idx != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objAttrType)  ($ent1)"}
    
              if {[info exists cells($dt)]} {
                set ok 0
                set colName ""
                set ov $objValue
                set invalid ""

# get values for these entity and attribute pairs
                switch -glob $ent1 {
                  "angular_location* name" -
                  "angular_size* name" {
# angular_location/size.name, add nothing to dimrep as there is no symbol associated with the location                    
                    set ok 1
                    set col($dt) $pmiStartCol($dt)
                    set colName "dimension name[format "%c" 10](Sec. 5.1.1, 5.1.5)"
                    set dimtolEnt $objEntity
  
                    set dimrepID $objID
                    set dimrep($dimrepID) ""
                    set dim(symbol) ""
                    set dimName $ov
                    if {[string first "directed" $ent1] != -1} {set dimDirected 1}

                    if {[string first "angular_location" $ent1] != -1} {
                      set item "angular location"
                    } else {
                      set item "angular size"
                    }
                    lappend spmiTypesPerFile $item
                    
                    set angDegree 1
                    if {![info exists entCount(conversion_based_unit_and_plane_angle_unit)]} {set angDegree 0}
                  }
                  "angular_location* angle_selection" -
                  "angular_size* angle_selection" {
# check angle selection values                    
                    if {$ov != "equal" && $ov != "large" && $ov != "small"} {
                      set msg "Syntax Error: Invalid 'angle_selection' attribute ($ov) on [formatComplexEnt [lindex $ent1 0]].\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.1.2, 5.1.6)"
                      errorMsg $msg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                    }
                  }
                  "*dimensional_size* name" {
# dimensional_size.name, from the name add symbol to dimrep for spherical, radius, diameter or thickness  (Section 5.1.5, Table 4)
                    set okname 0
                    set ok 1
                    set col($dt) $pmiStartCol($dt)
                    set colName "dimension name[format "%c" 10](Sec. 5.1.1, 5.1.5)"
                    set dimtolEnt $objEntity
                    
                    set dimrepID $objID
                    if {![info exists dimrep($dimrepID)]} {set dimrep($dimrepID) ""}
                    set item ""
                    set dim(symbol) ""
                    set dimName $ov
                    
                    set d1 ""
                    if {[info exists dimrep($dimrepID)]} {set d1 [string index $dimrep($dimrepID) 0]}
                      
                    if {[string first "spherical" $ov] != -1} {
                      if {[string index $dimrep($dimrepID) 0] != "S"} {
                        append dimrep($dimrepID) "S"
                        append item "spherical "
                      }
                    }
                    if {[string first "diameter" $ov] != -1 && $d1 != $pmiUnicode(diameter)} {
                      if {[string first $pmiUnicode(diameter) $dimrep($dimrepID)] == -1} {
                        append dimrep($dimrepID) $pmiUnicode(diameter)
                        append item "diameter"
                      }
                    }
                    if {[string first "radius" $ov] != -1} {
                      if {[string first "R" $dimrep($dimrepID)] == -1} {
                        append dimrep($dimrepID) "R"
                        append item "radius"
                      }
                    }
                    if {$ov == "thickness"} {
                      # no symbol (leader lines) for thickness
                      #append dimrep($dimrepID) $pmiUnicode(thickness)
                      append item "thickness"
                    }

                    set dim(symbol) $dimrep($dimrepID)
                    if {$entLevel == 2} {
                      incr numDSnames
                      set okname 1
                      if {[string first "dimensional_size_with" $ent1] != -1 && $numDSnames == 1} {
                        set okname 0
                        set ok 0
                      }
                      if {$okname} {
                        lappend spmiTypesPerFile "dimensional size"
                        lappend spmiTypesPerFile $ov
                      }
                    }

# syntax check for correct dimensional_size.name attribute (dimSizeNames) from the RP                  
                    if {$okname} {
                      if {$ov == ""} {
                        set ov "(blank)"
                        set msg "Syntax Error: Missing 'name' attribute on [formatComplexEnt [lindex $ent1 0]].\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.1.5, Table 4)"
                        errorMsg $msg
                        lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                        set invalid $msg
                      } elseif {[lsearch $dimSizeNames $ov] == -1} {
                        set msg "Syntax Error: Invalid 'name' attribute ($ov) on [formatComplexEnt [lindex $ent1 0]].\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.1.5, Table 4)"
                        errorMsg $msg
                        lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                        set invalid $msg
                      }
                    }
                  }
                  "directed_dimensional_location* name" -
                  "dimensional_location* name" {
# dimensional_location.name, add nothing to dimrep as there is no symbol associated with the location                    
                    set ok 1
                    set col($dt) $pmiStartCol($dt)
                    set colName "dimension name[format "%c" 10](Sec. 5.1.1, 5.1.5)"
                    set dimtolEnt $objEntity
  
                    set dimrepID $objID
                    set dim(symbol) ""
                    set dimrep($dimrepID) ""
                    set dimName $ov
                    if {[string first "directed" $ent1] != -1} {
                      set dimDirected 1
                      lappend spmiTypesPerFile "directed dimension"
                    }
                    lappend spmiTypesPerFile "dimensional location"

# syntax check for correct dimensional_location.name attribute from the RP                  
                    if {$ent1 == "dimensional_location name"} {
                      if {$ov == ""} {
                        set ov "(blank)"
                        set msg "Syntax Error: Missing 'name' attribute on [lindex $ent1 0].\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.1.1, Tables 1 and 2)"
                        errorMsg $msg
                        set invalid $msg
                        lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                      } elseif {$ov != "curved distance" && [string first "linear distance" $ov] == -1} {
                        set msg "Syntax Error: Invalid 'name' attribute ($ov) on [lindex $ent1 0].\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.1.1, Tables 1 and 2)"
                        errorMsg $msg
                        lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                        set invalid $msg
                      }
                    }
                  }
                  "shape_dimension_representation name" {
# shape_dimension.name, look for independency or envelope requirement per the RP, Sec 5.2.1                    
                    set invalid ""
                    if {$ov != ""} {
                      set ok 1
                      set col($dt) [expr {$pmiStartCol($dt)+1}]
                      set colName "representation name[format "%c" 10](Sec. 5.2.1)"
                      if {$ov == "independency" || $ov == "envelope requirement"} {
                        regsub -all " " $ov "_" ov1
                        append savedModifier $pmiModifiers($ov1)
                        lappend spmiTypesPerFile $ov
                      } else {
                        set msg "Syntax Error: Invalid 'name' attribute ($ov) on [lindex $ent1 0].\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.2.1, Table 5)"
                        errorMsg $msg
                        lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                        set invalid $msg
                      }
                    }
                  }
                  "length_measure_with_unit_and_measure_representation_item name" -
                  "measure_representation_item_and_plane_angle_measure_with_unit* name" -
                  "*measure_representation_item_and_qualified_representation_item* name" {
# measure_representation_item.name, get the name of the length/angle
                    set ok 1
                    set col($dt) [expr {$pmiStartCol($dt)+3}]
                    set colName "length/angle name"
                  }
                  "descriptive_representation_item name" {
# dimension modifiers, Sec 5.3, descriptive_representation_item.name must be 'dimensional note'
                    set ok 1
                    set col($dt) [expr {$pmiStartCol($dt)+6}]
                    set colName "modifier type 1[format "%c" 10](Sec. 5.3)"
                    if {$ov == ""} {
                      set msg "Syntax Error: Missing 'name' attribute on [lindex $ent1 0] must be 'dimensional note'.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.3)"
                      errorMsg $msg
                      set invalid $msg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                    } elseif {$ov != "dimensional note"} {
                      set msg "Syntax Error: Invalid 'name' attribute ($ov) on [lindex $ent1 0] must be 'dimensional note'.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.3)"
                      errorMsg $msg
                      set invalid $msg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                    }
                  }
                  "descriptive_representation_item description" {
# dimension modifiers, Sec 5.3, Tables 7 and 8
                    set ok 1
                    set col($dt) [expr {$pmiStartCol($dt)+7}]
                    set colName "modifier type 2[format "%c" 10](Sec. 5.3)"
                    if {$ov == ""} {
                      set msg "Syntax Error: Missing 'description' attribute on [lindex $ent1 0].\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.3, Table 7)"
                      errorMsg $msg
                      set invalid $msg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
# theoretical, auxiliary, i.e.  basic [50], reference (50) (Section 5.3)
                    } elseif {$entLevel == 3} {
                      if {$ov == "theoretical"} {
                        set dimBasic 1
                        lappend spmiTypesPerFile "basic dimension"
                      } elseif {$ov == "auxiliary"} {
                        set dimReference 1
                        lappend spmiTypesPerFile "reference dimension"
                      } else {
                        set msg "Syntax Error: Invalid 'description' attribute ($ov) on [lindex $ent1 0] must be 'theoretical' or 'auxiliary'.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.3, Table 7)"
                        errorMsg $msg
                        set invalid $msg
                        lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                      }
# dimension modifier - statistical tolerance, continuous feature, controlled radius, square, etc. (Section 5.3, Table 8)
                    } elseif {$entLevel == 4} {
                      if {[lsearch $dimModNames $ov] != -1} {
                        regsub -all " " $ov "_" ov
# controlled radius and square are prefixes, instead of the default suffix
                        if {$ov == "controlled_radius"} {
                          if {[string index $dimrep($dimrepID) 0] == "R"} {
                            set dimrep($dimrepID) "C$dimrep($dimrepID)"
                          } else {
                            set dimrep($dimrepID) "$pmiModifiers($ov)$dimrep($dimrepID)"
                          }
                          set ov "controlled radius"
                          set pos [lsearch $spmiTypesPerFile "radius"]
                          set spmiTypesPerFile [lreplace $spmiTypesPerFile $pos $pos]
                        } elseif {$ov == "square"} {
                          set dimrep($dimrepID) "$pmiModifiers($ov)$dimrep($dimrepID)"
# suffix, append to savedModifier
                        } else {
                          append savedModifier $pmiModifiers($ov)
                        }
                        lappend spmiTypesPerFile $ov
# bad dimension modifier
                      } else {
                        append dimrep($dimrepID) " ($ov)"
                        set msg "Syntax Error: Invalid 'description' attribute ($ov) on [lindex $ent1 0].\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.3, Table 8)"
                        errorMsg $msg
                        set invalid $msg
                        lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                      }
                    }
                  }
                  "type_qualifier name" {
# type qualifier (Section 5.2.2, Table 6)
                    set ok 1
                    set col($dt) [expr {$pmiStartCol($dt)+8}]
                    set colName "qualifier[format "%c" 10](Sec. 5.2.2)"
                    if {$ov == ""} {
                      set msg "Syntax Error: Missing 'name' attribute on [lindex $ent1 0].\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.2.2, Table 6)"
                      errorMsg $msg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                      set invalid $msg
                    } elseif {$ov != "maximum" && $ov != "minimum" && $ov != "average"} {
                      set msg "Syntax Error: Invalid 'name' attribute ($ov) on [lindex $ent1 0].\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.2.2, Table 6)"
                      errorMsg $msg
                      lappend syntaxErr([lindex [split $ent1 " "] 0]) [list $objID [lindex [split $ent1 " "] 1] $msg]
                      set invalid $msg
                    } else {
                      append dimrep($dimrepID) " [string toupper [string range $ov 0 2]]"
                      lappend spmiTypesPerFile "type qualifier"
                    }
                  }
                  "value_format_type_qualifier format_type" {
# decimal places, Sec 5.4, in the form of NR x.y from ASN.1, ISO 6093
                    set ok 1
                    set col($dt) [expr {$pmiStartCol($dt)+5}]
                    set colName "decimal places[format "%c" 10](Sec. 5.4)"
                    lappend spmiTypesPerFile "decimal places"
                  }
                  "axis2_placement_3d name" {
# oriented dimension location
                    if {$objValue == "orientation"} {
                      set dimOrient 1
                      set dimOrientVal "($objType $objID)"
                      set ok 1
                      set col($dt) [expr {$pmiStartCol($dt)+10}]
                      set colName "oriented dimension[format "%c" 10](Sec. 5.1.3)"
                      lappend spmiTypesPerFile "oriented dimensional location"
                      if {[string first "dimensional_location" [$dimtolEnt Type]] != 0} {
                        set msg "Syntax Error: Oriented Dimension Location cannot be used with [$dimtolEnt Type].\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.1.3)"
                        errorMsg $msg
                        lappend syntaxErr(dimensional_characteristic_representation) [list "-$spmiIDRow($dt,$spmiID)" "oriented dimension" $msg]
                        #set invalid 1
                      }
                    }
                  }
                }

# value in spreadsheet
                if {$ok && [info exists spmiID]} {
                  set c [string index [cellRange 1 $col($dt)] 0]
                  if {[info exists spmiIDRow($dt,$spmiID)]} {
                    set r $spmiIDRow($dt,$spmiID)
                    #outputMsg "$dt $spmiID $r" blue

# column name
                    if {$colName != ""} {
                      if {![info exists pmiHeading($col($dt))]} {
                        $cells($dt) Item 3 $c $colName
                        set pmiHeading($col($dt)) 1
                        set pmiCol [expr {max($col($dt),$pmiCol)}]
                      }
                    }

# keep track of rows with semantic PMI
                    if {[lsearch $spmiRow($dt) $r] == -1} {lappend spmiRow($dt) $r}
    
                    set ov $objValue 
                    set val [[$cells($dt) Item $r $c] Value]

                    if {$invalid != ""} {
                      if {$colName != ""} {
                        lappend syntaxErr($dt) [list "-$r" $colName $invalid]
                      } else {
                        lappend syntaxErr($dt) [list $r $col($dt) $invalid]
                      }
                    }

# append a2p3d orientation
                    if {[info exists dimOrientVal]} {append ov "[format "%c" 10]$dimOrientVal"}
                    
                    if {$val == ""} {
                      $cells($dt) Item $r $c $ov
                    } else {

# value range (limit dimension), Sec. 5.2.4, usually 'nominal value' is missing
                      if {[string first "limit" $val] != -1 && [string first "limit" $ov] != -1 && $dim(num) == 2} {
                        set msg "Syntax Error: Missing 'nominal value' for value range.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.2.4)"
                        errorMsg $msg
                        lappend syntaxErr($dt) [list -$r "length/angle name" $msg]
                      }

                      if {$ov == "upper limit"} {
                        set item "value range"
                        lappend spmiTypesPerFile $item
                      }
                      if {[string first $ov $val] == -1} {$cells($dt) Item $r $c "$val[format "%c" 10]$ov"}
                    }

# keep track of max column
                    set pmiCol [expr {max($col($dt),$pmiCol)}]
                  } else {
                    errorMsg "ERROR processing Dimensional Tolerance"
                    #outputMsg "$dt $spmiID [info exists spmiIDRow($dt,$spmiID)]" red
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing Dimtol ($objNodeType $ent2)\n $emsg3"
            set entLevel 1
          }
        }
      }
    }
  }
  incr entLevel -1
  
# write a few more things at the end of processing a semantic PMI entity
  if {$entLevel == 0} {
    
# associated geometry (5.1.5), find link between dimtol and geometry through geometric_item_specific_usage (gisu)
# dimtolEnt is either dimensional_location, angular_location, dimensional_size, or angular_size
    if {[catch {
      if {[info exists dimtolEnt]} {
        set dimtolType [$dimtolEnt Type]
        set dimtolID   [$dimtolEnt P21ID]
        
# dimensional_size
        if {[string first "_size" $dimtolType] != -1} {
          ::tcom::foreach dimtolAtt [$dimtolEnt Attributes] {
            if {[$dimtolAtt Name] == "applies_to"} {
              set val [$dimtolAtt Value]
              #outputMsg " [$dimtolAtt Name] / [$dimtolAtt Value]  [$val Type] [$val P21ID]" green

# directly to GISA
              if {$val != ""} {
                getAssocGeom $val 1

# through SAR(s) to GISU
                if {[llength [array names assocGeom]] == 0} {
                  set sars [$val GetUsedIn [string trim shape_aspect_relationship] [string trim relating_shape_aspect]]
                  ::tcom::foreach sar $sars {
                    ::tcom::foreach asar [$sar Attributes] {
                      if {[$asar Name] == "related_shape_aspect"} {
                        set val1 [$asar Value]
                        getAssocGeom $val1 1
                      }
                    }
                  }
                }
              }
            }
          }

# dimensional_location
        } elseif {[string first "_location" [$dimtolEnt Type]] != -1} {
          ::tcom::foreach dimtolAtt [$dimtolEnt Attributes] {
            if {[string first "relat" [$dimtolAtt Name]] != -1} {
              set val [$dimtolAtt Value]
              if {$val != ""} {getAssocGeom $val 1}
            }
          }
        }
      }

# report associated geometry
      if {[info exists assocGeom]} {
        if {[info exists spmiIDRow($dt,$spmiID)]} {
          set str [reportAssocGeom $dimtolType $spmiIDRow($dt,$spmiID)]
        } else {
          set str [reportAssocGeom $dimtolType]
        }
        set dimtolGeomEnts ""
        
        if {$str != "" && [info exists spmiIDRow($dt,$spmiID)]} {
          if {![info exists pmiColumns(ch)]} {set pmiColumns(ch) [expr {$pmiStartCol($dt)+12}]}
          set colName "Associated Geometry[format "%c" 10](Sec. 5.1.1, 5.1.5)"
          set c [string index [cellRange 1 $pmiColumns(ch)] 0]
          set r $spmiIDRow($dt,$spmiID)
          if {![info exists pmiHeading($pmiColumns(ch))]} {
            $cells($dt) Item 3 $c $colName
            set pmiHeading($pmiColumns(ch)) 1
            set pmiCol [expr {max($pmiColumns(ch),$pmiCol)}]
          }
          $cells($dt) Item $r $pmiColumns(ch) [string trim $str]
          
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
                  errorMsg "Associated Geometry for a '[lindex $item 0]' dimension also refers to a '[lindex $item 1]'.\n [string totitle $dimName] dimensions are assumed to be associated only with curved surfaces."
                  addCellComment $dt $r $pmiColumns(ch) "[string totitle $dimName] dimension (column E) also refers to a '[lindex $item 1]'." 200 50
                } else {
                  errorMsg "Associated Geometry for a '[lindex $item 0]' dimension is only a '[lindex $item 1]'.\n [string totitle $dimName] dimensions are assumed to be associated with curved surfaces."
                  addCellComment $dt $r $pmiColumns(ch) "[string totitle $dimName] dimension (column E) is assumed to be associated with curved surfaces and not a '[lindex $item 1]'." 200 50
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
            set dimtolGeomEnts [join [lsort $nstr]]
            set dimtolEntType($dimtolGeomEnts) "$dimtolType $dimtolID"
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding Associated Geometry: $emsg"
    }
    
# plus minus tolerance on dimtolEnt
    if {[catch {
      if {[info exists dimtolEnt]} {
        #outputMsg [$dimtolEnt Type] blue
        set objGuiEntities [$dimtolEnt GetUsedIn [string trim plus_minus_tolerance] [string trim toleranced_dimension]]
        set plusminus ""
        set colName ""
        set npm 0
        ::tcom::foreach objGuiEntity $objGuiEntities {
          incr npm
          #outputMsg "[$objGuiEntity Type] [$objGuiEntity P21ID]"
          ::tcom::foreach attrPMT [$objGuiEntity Attributes] {

# +- range
            if {[$attrPMT Name] == "range"} {
              set subEntity [$attrPMT Value]
              set subType [$subEntity Type]
              
              if {$subType == "tolerance_value"} {
                set colName "plus minus bounds[format "%c" 10](Sec. 5.2.3)"

# +- range lower/upper bound on tolerance_value
                ::tcom::foreach subAttr [$subEntity Attributes] {
                  if {[string first "bound" [$subAttr Name]] != -1} {
                    #outputMsg " [$subAttr Name] [$subAttr Value]"
                    ::tcom::foreach measureAttr [[$subAttr Value] Attributes] {
                      if {[$measureAttr Name] == "value_component"} {
                        #outputMsg "  [$measureAttr Name] [$measureAttr Type] [$measureAttr Value]"
                        append plusminus "[$measureAttr Value] "
                      }
                    } 
                  }
                }
                if {$npm > 1} {errorMsg "Syntax Error: Multiple plus-minus tolerance values ($plusminus) for a single dimension."}

# tolerance class on limits_and_fits
              } elseif {$subType == "limits_and_fits"} {
                set colName "tolerance class[format "%c" 10](Sec. 5.2.5)"
                ::tcom::foreach subAttr [$subEntity Attributes] {
                  append plusminus "[$subAttr Name] - [$subAttr Value][format "%c" 10]"
                  if {[$subAttr Name] == "form_variance"} {set form_variance [$subAttr Value]}
                  if {[$subAttr Name] == "grade"}         {set form_grade    [$subAttr Value]}
                }                
              }
            }
          }
        }

# construct correct plus-minus
        if {$plusminus != "" && [info exists spmiIDRow($dt,$spmiID)]} {

# tolerance class
          if {[info exists form_variance]} {
            if {![info exists pmiColumns(tolclass)]} {set pmiColumns(tolclass) [expr {$pmiStartCol($dt)+9}]}
            set c [string index [cellRange 1 $pmiColumns(tolclass)] 0]
            set r $spmiIDRow($dt,$spmiID)
            if {![info exists pmiHeading($pmiColumns(tolclass))]} {
              $cells($dt) Item 3 $c $colName
              set pmiHeading($pmiColumns(tolclass)) 1
              set pmiCol [expr {max($pmiColumns(tolclass),$pmiCol)}]
            }
            $cells($dt) Item $r $pmiColumns(tolclass) [string trim $plusminus]
            append dimrep($dimrepID) " $form_variance"
            if {[info exists form_grade]} {append dimrep($dimrepID) $form_grade}
            catch {unset form_variance}
            catch {unset form_grade}
            lappend spmiTypesPerFile "tolerance class"

# add +- values to cell
          } else {
            if {![info exists pmiColumns(pmt)]} {set pmiColumns(pmt) [expr {$pmiStartCol($dt)+4}]}
            set c [string index [cellRange 1 $pmiColumns(pmt)] 0]
            set r $spmiIDRow($dt,$spmiID)
            if {![info exists pmiHeading($pmiColumns(pmt))]} {
              $cells($dt) Item 3 $c $colName
              set pmiHeading($pmiColumns(pmt)) 1
              set pmiCol [expr {max($pmiColumns(pmt),$pmiCol)}]
            }
            $cells($dt) Item $r $pmiColumns(pmt) [string trim $plusminus]

# reconstruct +-            
            set pm [split [string trim $plusminus] " "]
            if {[llength $pm] == 2} {

# truncate
              if {[getPrecision [lindex $pm 0]] > 4 || [getPrecision [lindex $pm 1]] > 4} {
                set pm1 [string trimright [format "%.4f" [lindex $pm 0]] "0"]
                set pm2 [string trimright [format "%.4f" [lindex $pm 1]] "0"]
                set pm [list $pm1 $pm2]
              }

              set sdimrep [split $dimrep($dimrepID) " "]
              set dmval [lindex $sdimrep 0]
              
# (0) < (1)              
              for {set i 0} {$i < 2} {incr i} {
                set pmval($i) [lindex $pm $i]

# fix trailing . and 0
                set pmval($i) [removeTrailingZero $pmval($i)]
              }
              
# errors with +/- values
              if {$pmval(0) > $pmval(1)} {
                set msg "Syntax Error: Plus-minus tolerance values for the lower ($pmval(0)) and upper ($pmval(1)) limits are reversed.  See tolerance_value entity.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.2.3)"
                errorMsg $msg
                lappend syntaxErr($dt) [list -$r "plus minus bounds" $msg]
              } elseif {$pmval(0) == $pmval(1)} {
                set msg "Syntax Error: Plus-minus tolerance values are equal ($pmval(0)).  See tolerance_value entity.\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 5.2.3)"
                errorMsg $msg
                lappend syntaxErr($dt) [list -$r "plus minus bounds" $msg]
              }
              
# EQUAL values              
              if {[string range $pmval(0) 1 end] == $pmval(1)} {

# get precision of +- and compare to dimension
                if {[info exists dim(unit)]} {
                  if {$dim(unit) == "INCH"} {
                    if {$pmval(1) < 1} {set pmval(1) [string range $pmval(1) 1 end]}

# add trailing zeros to dimension
                    set pmprec [getPrecision $pmval(1)]
                    set n0 [expr {$pmprec-$dim(prec,$dimrepID)}]
                    if {$n0 > 0 && ![info exists dim(qual)]} {
                      if {[string first "." $dmval] == -1} {append dmval "."}
                      append dmval [string repeat "0" $n0]
                      set dim(prec,$dimrepID) $pmprec
                      if {$dim(prec,$dimrepID) > $dim(prec,max) && $dim(prec,$dimrepID) < 6} {set dim(prec,max) $dim(prec,$dimrepID)}
                    }
                  }
                }

# rewrite dimrep
                set dimrep($dimrepID) "$dmval $pmiUnicode(plusminus) $pmval(1)"
                if {[info exists dim(angle)] && [info exists angDegree]} {
                  if {$dim(angle) && $angDegree} {append dimrep($dimrepID) $pmiUnicode(degree)}
                }
                if {[llength $sdimrep] > 1} {append dimrep($dimrepID) "  [lrange $sdimrep 1 end]"}
                lappend spmiTypesPerFile "bilateral tolerance"

# NON EQUAL values
              } elseif {$pmval(0) != $pmval(1)} {

# inch units
                if {[info exists dim(unit)]} {
                  if {$dim(unit) == "INCH"} {

# get precision of +- and compare to dimension                
                    set pmprec [expr {max([getPrecision $pmval(0)],[getPrecision $pmval(1)])}]
                    if {$pmprec > $dim(prec,$dimrepID)} {set dim(prec,$dimrepID) $pmprec}
                    if {$dim(prec,$dimrepID) > $dim(prec,max) && $dim(prec,$dimrepID) < 6} {set dim(prec,max) $dim(prec,$dimrepID)}

# add trailing zeros to dimension
                    set n0 [expr {$pmprec-$dim(prec,$dimrepID)}]
                    if {$n0 > 0} {
                      if {[string first "." $dmval] == -1} {append dmval "."}
                      append dmval [string repeat "0" $n0]
                      set dimrep($dimrepID) $dmval
                    }

# remove leading zeros
                    #outputMsg "\n$pmval(0) $pmval(1)"
                    for {set i 0} {$i < 2} {incr i} {
                      if {$pmval($i) == "-0."} {set pmval($i) "0."}
                      if {$pmval($i) < 0} {
                        set pmval($i) [string range $pmval($i) 1 end]
                        if {$pmval($i) < 1} {set pmval($i) [string range $pmval($i) 1 end]}
                        set pmval($i) "-$pmval($i)"
                      } else {
                        if {$pmval($i) < 1} {set pmval($i) [string range $pmval($i) 1 end]}
                      }
                    }

# fix -.000
                    #outputMsg "$pmval(0) $pmval(1)"
                    if {$pmval(1) > 0 && $pmval(0) == "."} {
                      set pmval(1) "+$pmval(1)"
                      set pmval(0) "-.0"
                    } elseif {$pmval(1) > 0 && $pmval(0) < 0} {
                      set pmval(1) "+$pmval(1)"
                    } elseif {$pmval(1) == "." && $pmval(0) < 0} {
                      set pmval(1) "-.0"
                    }

# add trailing zeros to precision
                    for {set i 0} {$i < 2} {incr i} {
                      if {$pmprec > [getPrecision $pmval($i)]} {
                        set n0 [expr {$pmprec-[getPrecision $pmval($i)]}]
                        append pmval($i) [string repeat "0" $n0]
                      }
                    }
                    #outputMsg "$pmval(0) $pmval(1)"

# mm units
                  } elseif {$dim(unit) == "MM"} {

# get precision of +                
                    set pmprec [expr {max([getPrecision $pmval(0)],[getPrecision $pmval(1)])}]

# fix 0.
                    if {$pmval(1) >= 0 && $pmval(0) <= 0} {
                      if {$pmval(1) > 0} {
                        set pmval(1) "+$pmval(1)"
                      } else {
                        set pmval(1) "0"
                      }
                      if {$pmval(0) == 0} {set pmval(0) "0"}
                    }

# add trailing zeros to precision
                    for {set i 0} {$i < 2} {incr i} {
                      if {$pmprec > [getPrecision $pmval($i)] && $pmval($i) != 0} {
                        set n0 [expr {$pmprec-[getPrecision $pmval($i)]}]
                        append pmval($i) [string repeat "0" $n0]
                      }
                    }
                  }
                }
                
# show reconstructed +- tolerance
                set deg ""
                if {$dim(angle)} {if {$angDegree} {set deg $pmiUnicode(degree)}}
                set indent [string repeat " " [expr {3*[string length $dimrep($dimrepID)]}]]
                append dimrep($dimrepID) "  $pmval(1)$deg[format "%c" 10]$indent$pmval(0)$deg"
                if {[llength $sdimrep] > 1} {append dimrep($dimrepID) "  [lrange $sdimrep 1 end]"}
                lappend spmiTypesPerFile "non-bilateral tolerance"
              }
            }
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding Plus Minus Tolerance: $emsg"
    }
    
# report complete dimension representation (dimrep)
    if {[catch {
      set cellComment ""
      if {[info exists dimrep] && [info exists spmiIDRow($dt,$spmiID)]} {
        if {![info exists pmiColumns(dmrp)]} {set pmiColumns(dmrp) 4}
        set c [string index [cellRange 1 $pmiColumns(dmrp)] 0]
        set r $spmiIDRow($dt,$spmiID)
        if {![info exists pmiHeading($pmiColumns(dmrp))]} {
          set colName "Dimensional[format "%c" 10]Tolerance"           
          $cells($dt) Item 3 $c $colName
          set pmiHeading($pmiColumns(dmrp)) 1
          set pmiCol [expr {max($pmiColumns(dmrp),$pmiCol)}]
          set dimtolComment "Repetitive Dimensions (e.g., 4X) are shown for diameters and radii.  They are computed based on the number of cylindrical, spherical, and toroidal surfaces associated with a dimension (see Associated Geometry column to the right) and, depending on the CAD system, might be off by a factor of two, have the wrong value, or be missing."
        }
        
# add brackets or parentheses for basic or reference dimensions
        if {[info exists dimBasic]} {
          if {$dim(unit) == "INCH" && $dim(angle) == 0} {
            set n0 [expr {$dim(prec,max)-$dim(prec,$dimrepID)}]
            if {$n0 > 0} {
              if {[string first "." $dimrep($dimrepID)] == -1} {append dimrep($dimrepID) "."}
              append dimrep($dimrepID) [string repeat "0" $n0]
            }
          }
          set dimrep($dimrepID) "\[$dimrep($dimrepID)\]"
          unset dimBasic

        } elseif {[info exists dimReference]} {
          if {[string is double $dimrep($dimrepID)]} {
            set dimrep($dimrepID) "'($dimrep($dimrepID))"
          } else {
            set dimrep($dimrepID) "($dimrep($dimrepID))"
          }
          unset dimReference
        }

        set dr [string trim $dimrep($dimrepID)]
        if {[string is integer $dimrep($dimrepID)] || [string is double $dimrep($dimrepID)]} {
          if {$dim(unit) == "INCH"} {
            set n0 [expr {$dim(prec,max)-$dim(prec,$dimrepID)}]
            if {$n0 > 0} {
              if {[string first "." $dr] == -1} {append dr "."}
              append dr [string repeat "0" $n0]
            }
          }
          set dr "'$dr"
        }
        
# saved modifiers to append to dimension
        if {[info exist savedModifier]} {
          append dr " $savedModifier"
          unset savedModifier
        }
        
# directed
        if {[info exist dimDirected]} {
          append dr "[format "%c" 10](directed)"
          unset dimDirected
          set cellComment "For the definition of a 'directed' dimension, see the CAx-IF Recommended Practice for $recPracNames(pmi242), Sec. 5.1.1, 5.1.7"
        }
        
# oriented
        if {[info exist dimOrient]} {
          append dr "[format "%c" 10](oriented)"
          unset dimOrient
          #set cellComment "For the definition of an 'oriented' dimension, see the CAx-IF Recommended Practice for $recPracNames(pmi242), Sec. 5.1.3"
        }
        
# dimension count
        if {[info exists dimRepeat]} {
          if {$dimRepeat > 1} {
            if {[string index $dr 0] == "'"} {set dr [string range $dr 1 end]}
            set dr "$dimRepeat\X $dr"
            set dimRepeat ""
            if {$dimtolComment != ""} {
              addCellComment $dt 3 $c $dimtolComment
              set dimtolComment ""
            }
          }
        }

# write dimension to spreadsheet
        $cells($dt) Item $r $pmiColumns(dmrp) $dr
        if {$cellComment != ""} {
          addCellComment $dt $r $pmiColumns(dmrp) $cellComment
          lappend entsWithErrors "dimensional_characteristic_representation"
        }
        
# save dimension with associated geometry
        if {$dimtolGeomEnts != ""} {
          if {[string first "'" $dr] == 0} {set dr [string range $dr 1 end]}
          if {[info exists dimtolGeom($dimtolGeomEnts)]} {
            if {[lsearch $dimtolGeom($dimtolGeomEnts) $dr] == -1} {lappend dimtolGeom($dimtolGeomEnts) $dr}
          } else {
            lappend dimtolGeom($dimtolGeomEnts) $dr
          }

# multiple dimensions for same geometry
          if {[llength $dimtolGeom($dimtolGeomEnts)] > 1} {
            set dtg {}
            foreach item $dimtolGeom($dimtolGeomEnts) {
              regsub -all [format "%c" 10] $item " " tmp
              for {set i 0} {$i < 10} {incr i} {regsub -all "  " $tmp " " tmp}
              lappend dtg $tmp
            }
            errorMsg "Multiple dimensions $dtg associated with the same geometry\n $dimtolGeomEnts"
            addCellComment $dt $r $pmiColumns(ch) "Multiple dimensions are associated with the same geometry.  The identical information in this cell should appear in another Associated Geometry cell above."
            lappend entsWithErrors "dimensional_characteristic_representation"
          }
        }
      }

# unset variables
      foreach var {dimBasic dimReference savedModifier dimDirected dimOrient} {if {[info exists $var]} {unset $var}}

    } emsg]} {
      errorMsg "ERROR adding Dimensional Tolerance: $emsg"
    }
    set dim(name) ""
  }

  return 0
}

#-------------------------------------------------------------------------------
proc getPrecision {val} {
  set c1 [string first "." $val]
  set prec 0
  if {$c1 != -1} {set prec [string length [lindex [split $val "."] 1]]}
  return $prec
}

#-------------------------------------------------------------------------------
proc removeTrailingZero {val} {
  global dim

# remove trailing zero for cases like 1.0, inch > 1., mm > 1  
  if {[info exists dim(unit)]} {
    if {$dim(unit) == "INCH"} {
      if {[string index $val end] == "0" && [string index $val end-1] == "."} {set val [string range $val 0 end-1]}
    } elseif {$dim(unit) == "MM"} {
      if {[string index $val end] == "0" && [string index $val end-1] == "."} {set val [string range $val 0 end-2]}
    }
  }
  return $val
}
