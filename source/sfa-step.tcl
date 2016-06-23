#-------------------------------------------------------------------------------
proc getVersion {{pt ""}} {
  set app_version 1.72
  return $app_version
}

# -------------------------------------------------------------------------------
# check for an entity that is checked for semantic PMI
proc spmiCheckEnt {ent} {
  global spmiEntTypes tolNames

  set ok 0
  foreach sp $spmiEntTypes {if {[string first $sp $ent] ==  0} {set ok 1}}
  foreach sp $tolNames     {if {[string first $sp $ent] != -1} {set ok 1}}
  return $ok
}

# -------------------------------------------------------------------------------
# check for a valid form of annotation_occurrence
proc gpmiCheckEnt {ent} {
  global aoEntTypes

  set ok 0
  foreach item $aoEntTypes {
    if {[string first "tessellated" $item] == -1} {
      if {[string first $item $ent] == 0} {
        if {[string first "over_riding_styled_item" $ent] == -1 && \
            [string first "_relationship" $ent] == -1} {
          set ok 1
        }
      }
    } else {
      if {[string first $item $ent] != -1} {set ok 1}
    }
  }
  if {[string first "leader" $ent] != -1} {set ok 0}
  return $ok
}
 
# -------------------------------------------------------------------------------
# which STEP entities are processed depending on options
proc setEntsToProcess {entType objDesign} {
  global opt gpmiEnts spmiEnts
  
  set ok 0
  set gpmiEnts($entType) 0
  set spmiEnts($entType) 0

# for validation properties
  if {$opt(VALPROP)} {
    if {$entType == "boolean_representation_item" || \
        $entType == "conversion_based_unit_and_length_unit" || \
        $entType == "derived_unit_element" || \
        $entType == "descriptive_representation_item" || \
        $entType == "integer_representation_item" || \
        $entType == "length_unit_and_si_unit" || \
        $entType == "mass_unit_and_si_unit" || \
        $entType == "measure_representation_item" || \
        $entType == "property_definition" || \
        $entType == "property_definition_representation" || \
        $entType == "real_representation_item" || \
        $entType == "representation" || \
        $entType == "value_representation_item"} {
      set ok 1
    }
  }

# for PMI (graphical) presentation
  if {$opt(PMIGRF) && $ok == 0} {
    set ok [gpmiCheckEnt $entType]
    set gpmiEnts($entType) $ok
    if {$entType == "advanced_face" || \
        $entType == "characterized_item_within_representation" || \
        $entType == "colour_rgb" || \
        $entType == "curve_style" || \
        $entType == "fill_area_style" || \
        $entType == "fill_area_style_colour" || \
        $entType == "geometric_curve_set" || \
        $entType == "presentation_style_assignment" || \
        $entType == "property_definition" || \
        $entType == "representation_relationship" || \
        $entType == "tessellated_geometric_set" || \
        $entType == "view_volume"} {
      set ok 1
    }
    #foreach ent {"annotation" "style" "draughting" "colour" "tessellated" "_presentation" "camera" "shape_aspect" "pre_defined" "constructive_geometry"} 
    foreach ent {"annotation" "draughting" "_presentation" "camera" "constructive_geometry"} {
      if {[string first $ent $entType] != -1} {set ok 1}
    }
    if {!$ok} {if {[string first "representation" $entType] == -1 && [string first "presentation_" $entType] != -1} {set ok 1}}
  }

# for PMI (semantic) representation  
  if {$opt(PMISEM) && $ok == 0} {
    #set ok [spmiCheckEnt $entType]
    set spmiEnts($entType) [spmiCheckEnt $entType]
    if {$entType == "advanced_face" || \
        $entType == "compound_representation_item" || \
        $entType == "descriptive_representation_item" || \
        $entType == "dimensional_characteristic_representation" || \
        $entType == "draughting_model_item_association" || \
        $entType == "geometric_item_specific_usage" || \
        $entType == "id_attribute" || \
        $entType == "product_definition_shape" || \
        $entType == "property_definition" || \
        $entType == "shape_definition_representation" || \
        $entType == "shape_dimension_representation" || \
        $entType == "shape_representation_with_parameters" || \
        $entType == "value_format_type_qualifier" || \
        $entType == "value_range"} {
      set ok 1
    }
    foreach ent {"shape_aspect" "measure_with_unit" "measure_representation_item" "constructive_geometry"} {
      if {[string first $ent $entType] != -1} {set ok 1}
    }
    if {$entType == "axis2_placement_3d" && [$objDesign CountEntities "placed_datum_target_feature"] > 0} {set ok 1}
  }
  return $ok
}

# -------------------------------------------------------------------------------
# keep track of property_defintion, annotation occurrence, or semantic PMI rows in
# propDefIDRow, gpmiIDRow, spmiIDRow

proc setIDRow {entType p21id} {
  global gpmiEnts gpmiIDRow propDefIDRow row spmiEnts spmiIDRow
  
  if {$entType == "property_definition"} {
    set propDefIDRow($p21id) $row($entType)
  } elseif {$gpmiEnts($entType)} {
    set gpmiIDRow($entType,$p21id) $row($entType)
  } elseif {$spmiEnts($entType)} {
    set spmiIDRow($entType,$p21id) $row($entType)
  }
}

# -------------------------------------------------------------------------------
# check validation properties and PMI presentation
proc checkForPMIandValProps {objDesign entType} {
  global cells fixent gpmiEnts opt pmiColumns savedViewCol spmiEnts
  
# check for validation properties, call valProp
  if {$entType == "property_definition_representation"} {
    if {[catch {
      if {[info exists opt(VALPROP)]} {
        if {$opt(VALPROP)} {
          if {[lsearch $fixent "representation"] == -1} {
            if {[info exists cells(property_definition)]} {
              valPropStart $objDesign
            }
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding Validation Property information to '$entType'\n  $emsg"
    }

# check for PMI Presentation, call pmiProp
  } elseif {$gpmiEnts($entType)} {
    if {[catch {
      if {[info exists opt(PMIGRF)]} {
        if {$opt(PMIGRF)} {
          if {[info exists cells($entType)]} {
            gpmiProp $objDesign $entType
            set ok 0
          }
          catch {unset savedViewCol}
          catch {unset pmiColumns}
        }
      }
    } emsg]} {
      errorMsg "ERROR adding PMI Presentation information to '$entType'\n  $emsg"
    }

# check for Semantic PMI, call spmiDimtolStart or spmiGeotolStart
  } elseif {$spmiEnts($entType)} {
    if {[catch {
      if {[info exists opt(PMISEM)]} {
        if {$opt(PMISEM)} {
          if {[info exists cells($entType)]} {
            if {$entType == "dimensional_characteristic_representation"} {
              spmiDimtolStart $objDesign $entType
            } else {
              spmiGeotolStart $objDesign $entType
            }
            set ok 0
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding PMI Representation information to '[formatComplexEnt $entType]'\n  $emsg"
    }
  }
}

# -------------------------------------------------------------------------------
# get STEP AP name
proc getStepAP {fname} {
  global fileSchema1
  
  set fs [getSchemaFromFile $fname]
  set fileSchema1 $fs
  
  set ap ""
  foreach aps {AP203 AP209 AP210 AP214 AP242 AP238} {if {[string first $aps $fs] != -1} {set ap $aps}}
  if {$ap == ""} {
    if {[string first "CONFIGURATION_CONTROL_3D_DESIGN" $fs] != -1}  {set ap AP203}
    if {[string first "CONFIGURATION-CONTROL-3D-DESIGN" $fs] != -1}  {set ap AP203}
    if {[string first "CONFIG_CONTROL_DESIGN" $fs] != -1}            {set ap AP203}
    if {[string first "CONFIG_CONTROL_3D_DESIGN" $fs] != -1}         {set ap AP203}
    if {[string first "ccd_cla_gvp_ast" [string tolower $fs]] != -1} {set ap AP203}
    if {[string first "AUTOMOTIVE_DESIGN" $fs] != -1}                {set ap AP214}
    if {[string first "STRUCTURAL_ANALYSIS_DESIGN" $fs] != -1}       {set ap AP209}
    if {[string first "structural_analysis_design" $fs] != -1}       {set ap AP209}
    if {[string first "INTEGRATED_CNC_SCHEMA" $fs] != -1}            {set ap AP238}
  }
  set stepAP $ap
  return $stepAP
}

# -------------------------------------------------------------------------------
proc pmiFormatColumns {str} {
		global cells col excel gpmiRow invGroup opt pmiStartCol recPracNames row spmiRow thisEntType worksheet
		
  if {![info exists pmiStartCol($thisEntType)]} {
    return
  } else {  
    set c1 [expr {$pmiStartCol($thisEntType)-1}]
  }
  
# delete unused columns
  set delcol 0
  set ndelcol 0
  set colrange [[[$worksheet($thisEntType) UsedRange] Columns] Count]
  for {set i $colrange} {$i > 3} {incr i -1} {
    set val [[$cells($thisEntType) Item 3 $i] Value]
    if {$val != ""} {set delcol 1}
    if {$val == "" && $delcol} {
      set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
      $range Delete
      incr ndelcol
    }
  }
  set col($thisEntType) [expr {$col($thisEntType)-$ndelcol}]

# format
  if {[info exists cells($thisEntType)] && $col($thisEntType) > $c1} {
    set c2 [expr {$c1+1}]
    set c3 [expr {[getNextUnusedColumn $thisEntType 3]-1}]
    #outputMsg "$c1  $c2  $c3"
    
    outputMsg " Formatting $str on: [formatComplexEnt $thisEntType]" blue
    $cells($thisEntType) Item 2 $c2 $str
    set range [$worksheet($thisEntType) Range [cellRange 2 $c2]]
    $range HorizontalAlignment [expr -4108]
    [$range Font] Bold [expr 1]
    [$range Interior] ColorIndex [expr 36]
    set range [$worksheet($thisEntType) Range [cellRange 2 $c2] [cellRange 2 $c3]]
    $range MergeCells [expr 1]

# set rows for colors and borders
    set r1 1
    set r2 $r1
    set r3 {}
    if {[string first "PMI Presentation" $str] != -1} {
      set rs $gpmiRow($thisEntType)
    } elseif {[string first "PMI Representation" $str] != -1} {
      set rs $spmiRow($thisEntType)      
    }
    foreach r $rs {
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

# colors and borders
    set j 0
    for {set i $c2} {$i <= $c3} {incr i} {
      foreach r $r3 {
        set r1 [lindex $r 0]
        set r2 [lindex $r 1]
        set range [$worksheet($thisEntType) Range [cellRange $r1 $i] [cellRange $r2 $i]]
        [$range Interior] ColorIndex [lindex [list 36 35] [expr {$j%2}]]

        if {$i == $c2 && $r2 > 3} {
          if {$r1 < 4} {set r1 4}
          set range [$worksheet($thisEntType) Range [cellRange $r1 $c2] [cellRange $r2 $c3]]
          if {[expr {int([$excel Version])}] >= 12} {
            for {set k 7} {$k <= 12} {incr k} {
              catch {if {$k != 9 || [expr {$row($thisEntType)+0}] != $r2} {[[$range Borders] Item [expr $k]] Weight [expr 1]}}
            }
          }
        }
      }
      incr j
    }
    
# left and right borders in header
    catch {
      for {set i $c2} {$i <= $col($thisEntType)} {incr i} {
        set range [$worksheet($thisEntType) Range [cellRange 3 $i] [cellRange 3 $i]]
        [[$range Borders] Item [expr 7]]  Weight [expr 1]
        [[$range Borders] Item [expr 10]] Weight [expr 1]
      }
    }

# group columns
    if {[info exists invGroup($thisEntType)]} {if {$invGroup($thisEntType) < $c2} {set c2 $invGroup($thisEntType)}}
    set range [$worksheet($thisEntType) Range [cellRange 1 $c2] [cellRange [expr {$row($thisEntType)+2}] $c3]]
    [$range Columns] Group
    
# fix column widths depending on the name of the heading
    set colrange [[[$worksheet($thisEntType) UsedRange] Columns] Count]
    for {set i 1} {$i <= $colrange} {incr i} {
      set val [[$cells($thisEntType) Item 3 $i] Value]
      if {$val != ""} {
        set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
        $range ColumnWidth [expr 255]
      }
    }
    [$worksheet($thisEntType) Columns] AutoFit
    [$worksheet($thisEntType) Rows] AutoFit
  }
}
  
# -------------------------------------------------------------------------------
# Semantic PMI summary worksheet
proc spmiSummary {} {
  global cells entName excel localName row sheetLast spmiSumName spmiSumRow thisEntType worksheet worksheets xlFileName
  
  if {$spmiSumRow == 1} {
    set spmiSumName "PMI Representation Summary"
    set worksheet($spmiSumName) [$worksheets Add [::tcom::na] $sheetLast]
    $worksheet($spmiSumName) Activate
    $worksheet($spmiSumName) Name $spmiSumName
    set cells($spmiSumName) [$worksheet($spmiSumName) Cells]
    set wsCount [$worksheets Count]
    [$worksheets Item [expr $wsCount]] -namedarg Move Before [$worksheets Item [expr 3]]

    $cells($spmiSumName) Item $spmiSumRow 2 [file tail $localName]
    incr spmiSumRow 2
    $cells($spmiSumName) Item $spmiSumRow 1 "ID"
    $cells($spmiSumName) Item $spmiSumRow 2 "Entity"
    $cells($spmiSumName) Item $spmiSumRow 3 "PMI Representation"
    set range [$worksheet($spmiSumName) Range [cellRange 1 1] [cellRange 3 3]]
    [$range Font] Bold [expr 1]
    set range [$worksheet($spmiSumName) Range [cellRange 3 1] [cellRange 3 3]]
    if {[expr {int([$excel Version])}] >= 12} {
      [[$range Borders] Item [expr 8]] Weight [expr 2]
      [[$range Borders] Item [expr 9]] Weight [expr 2]
    }
    incr spmiSumRow

    [$worksheet($spmiSumName) PageSetup] Orientation [expr 2]
    [$worksheet($spmiSumName) PageSetup] PrintGridlines [expr 1]
    
    for {set i 1} {$i <= 3} {incr i} {
      [$worksheet($spmiSumName) Range [cellRange -1 $i]] ColumnWidth [expr 255]
      [$worksheet($spmiSumName) Range [cellRange -1 $i]] VerticalAlignment [expr -4160]
    }
    outputMsg " Adding PMI Representation Summary worksheet" green
    pmiAddModelPictures $spmiSumName
    [$worksheet($spmiSumName) Range "A1"] Select
  }
  
# add to PMI summary worksheet
  set hlink [$worksheet($spmiSumName) Hyperlinks]
  for {set i 3} {$i <= $row($thisEntType)} {incr i} {
    if {$thisEntType != "datum_reference_compartment" && $thisEntType != "datum_reference_element" && $thisEntType != "datum_reference_modifier_with_value"} {

# which entities and columns to summarize
      if {$i == 3} {
        for {set j 1} {$j < 30} {incr j} {
          set val [[$cells($thisEntType) Item 3 $j] Value]
          if {[string first "Datum Reference Frame" $val] != -1 || $val == "GD&T[format "%c" 10]Annotation" || $val == "Dimensional[format "%c" 10]Tolerance" || [string first "Datum Target" $val] == 0 || \
             ($thisEntType == "datum_reference" && [string first "reference" $val] != -1) || ($thisEntType == "referenced_modified_datum" && [string first "datum" $val] != -1)} {set pmiCol $j}
        }
      } else {
        if {[info exists pmiCol]} {

# values
          $cells($spmiSumName) Item $spmiSumRow 1 [[$cells($thisEntType) Item $i 1] Value]
          if {[string first "_and_" $thisEntType] == -1} {
            set entstr $thisEntType
          } else {
            regsub -all "_and_" $thisEntType ")[format "%c" 10][format "%c" 32][format "%c" 32][format "%c" 32](" entstr
            set entstr "($entstr)"
          }
          $cells($spmiSumName) Item $spmiSumRow 2 $entstr
          $cells($spmiSumName) Item $spmiSumRow 3 "'[[$cells($thisEntType) Item $i $pmiCol] Value]"

# link back to worksheets
          set anchor [$worksheet($spmiSumName) Range "B$spmiSumRow"]
          set hlsheet $thisEntType
          if {[string length $thisEntType] > 31} {
            foreach item [array names entName] {
              if {$entName($item) == $thisEntType} {set hlsheet $item}
            }
          }
          $hlink Add $anchor $xlFileName "$hlsheet!A1" "Go to $thisEntType"
          incr spmiSumRow
        } else {
          errorMsg "Cannot find PMI on $thisEntType for PMI Representation Summary worksheet"
        }
      }
    }
  }
}

# -------------------------------------------------------------------------------
# add images for the CAx-IF and NIST PMI models
proc pmiAddModelPictures {ent} {
  global cells localName modelPictures modelURLs mytemp nistName wdir worksheet

  set ftail [string tolower [file tail $localName]]
  
  if {[catch {
    set nlink 0
    set fl ""
    foreach pic $modelPictures {
      set ok 0
      if {[lindex $pic 0] == $nistName} {
        set ok 1
      } elseif {$nistName == ""} {
        set pic1 [split [lindex $pic 0] "-"]
        if {[string first [lindex $pic1 0] $ftail] != -1 && [string first [lindex $pic1 1] $ftail] != -1} {
          set ok 1
        }
      }
      if {$ok} {
        set fl [lindex $pic 1]
        set fc [lindex $pic 2]

        if {[file exists [file join $wdir images $fl]]} {file copy -force [file join $wdir images $fl] $mytemp}
        set fn [file nativename [file join $mytemp $fl]]
        #set fn [file nativename [file join $wdir images $fl]]

        if {[file exists $fn]} {
          set cellId [[$worksheet($ent) Cells] Range "$fc:$fc"]
          $cellId Select
          set shapeId [[[$cellId Parent] Shapes] AddPicture $fn [expr 0] [expr 1] [$cellId Left] [$cellId Top] -1 -1]

# group columns for image
          if {[lindex $pic 3] > 0} {
            set range [$worksheet($ent) Range [cellRange 1 5] [cellRange 1 [lindex $pic 3]]]
            [$range Columns] Group
          }

# link to test model drawings (doesn't always work)
          if {[string first "nist_" $fl] == 0 && $nlink < 2} {
            set str [string range $fl 0 10]
            foreach item $modelURLs {
              if {[string first $str $item] == 0} {
                catch {$cells($ent) Item 3 5 "Test Case Drawing"}
                set range [$worksheet($ent) Range E3:M3]
                $range MergeCells [expr 1]
                set range [$worksheet($ent) Range "E3"]
                [$worksheet($ent) Hyperlinks] Add $range [join "http://www.nist.gov/el/msid/infotest/upload/$item"] [join ""] [join "Link to Test Case Drawing (PDF)"]
                incr nlink
              }
            }
          }
        }
      }
    }
  } emsg]} {
    errorMsg "ERROR adding Picture to PMI Summary or Coverage worksheet.\n  $emsg"
  }
}

# -------------------------------------------------------------------------------
proc pmiSetEntAttrList {abc} {
  global elevel entAttrList ent opt

  incr elevel
  set ind [string repeat " " [expr {4*($elevel-1)}]]
  if {$opt(DEBUG1)} {outputMsg "$ind PARSE"}

  set ni 0
  foreach item $abc {
    if {[llength $item] > 1} {
      pmiSetEntAttrList $item
    } else {
      if {$ni == 0} {
        set typ "ENT"
        set ent($elevel) $item
      } else {
        set typ "  ATR"
        lappend entAttrList "$ent($elevel) $item"
      }
      if {$opt(DEBUG1)} {
        if {$typ == "ENT"} {
          outputMsg "$ind $typ $elevel $ni $item" blue
        } else {
          outputMsg "$ind $typ $elevel $ni $item"
        }
      }
      incr ni
    }
  }

  incr elevel -1
}  

# -------------------------------------------------------------------------------
# CAx-IF vendor abbrevitations

proc setCAXIFvendor {} {
  global localName
  
  set vendor(3de) "CT 3D Evolution"
  set vendor(3DE) "CT 3D Evolution"
  set vendor(a3) "Acrobat 3D"
  set vendor(a5) "Acrobat 3D (CATIA V5)"
  set vendor(ac) "AutoCAD"
  set vendor(al) "AliasStudio"
  set vendor(ap) "Acrobat 3D (Pro/E)"
  set vendor(au) "Acrobat 3D (NX)"
  set vendor(c4) "CATIA V4"
  set vendor(c5) "CATIA V5"
  set vendor(c6) "CATIA V6"
  set vendor(cg) "CgiStepCamp"
  set vendor(cm) "PTC CoCreate"
  set vendor(cr) "Creo"
  set vendor(dc) "Datakit CrossCad"
  set vendor(d5) "Datakit CrossCad (CATIA V5)"
  set vendor(dp) "Datakit CrossCad (Creo)"
  set vendor(du) "Datakit CrossCad (NX)"
  set vendor(dw) "Datakit CrossCad (SolidWorks)"
  set vendor(fs) "Vistagy FiberSim"
  set vendor(h3) "HOOPS 3D Exchange"
  set vendor(h5) "HOOPS 3D (CATIA V5)"
  set vendor(hp) "HOOPS 3D (Creo)"
  set vendor(hu) "HOOPS 3D (NX)"
  set vendor(i4) "ITI CADifx (CATIA V4)"
  set vendor(i5) "ITI CADfix (CATIA V5)"
  set vendor(id) "NX I-DEAS"
  set vendor(if) "ITI CADfix"
  set vendor(ii) "ITI CADfix (Inventor)"
  set vendor(in) "Autodesk Inventor"
  set vendor(ip) "ITI CADfix (Creo)"
  set vendor(iu) "ITI CADfix (NX)"
  set vendor(iw) "ITI CADfix (SolidWorks)"
  set vendor(kc) "Kubotek KeyCreator"
  set vendor(kr) "Kubotek REALyze"
  set vendor(lk) "LKSoft IDA-STEP"
  set vendor(nx) "Siemens NX"
  set vendor(oc) "Datakit CrossCad (OpenCascade)"
  set vendor(pc) "PTC CADDS"
  set vendor(pe) "Pro/E"
  set vendor(s4) "T-Systems COM/STEP (CATIA V4)"
  set vendor(s5) "T-Systems COM/FOX (CATIA V5)"
  set vendor(se) "SolidEdge"
  set vendor(sw) "SolidWorks"
  set vendor(t4) "Theorem Cadverter (CATIA V4)"
  set vendor(t5) "Theorem Cadverter (CATIA V5)"
  set vendor(tc) "Theorem Cadverter (CADDS)"
  set vendor(tp) "Theorem Cadverter (Creo)"
  set vendor(ts) "Theorem Cadverter (I-DEAS)"
  set vendor(tu) "Theorem Cadverter (NX)"
  set vendor(tx) "Theorem Cadverter"
  set vendor(ug) "Unigraphics"
  
  set fn [file tail $localName]
  set chars [list "-" "_" "."]
  foreach char $chars {
    set c1 [string first $char $fn]
    if {$c1 != -1} {
      set c2 [expr {$c1+1}]
      foreach idx [lsort [array names vendor]] {
        if {[string first $idx [string range $fn $c2 end]] == 0} {
          set c3 [expr {$c1+3}]
          set c4 [expr {$c1+4}]
          foreach char1 $chars {
            if {[string index $fn $c3] == $char1 || [string index $fn $c4] == $char1} {
              return $vendor($idx)
            }
          }
        }
      }
    }
  }
}
