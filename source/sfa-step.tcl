proc generateBOM {} {
  global bomAssembly bomAssemblyID bomIndent bomItems bomNAUO bomNAUOID cells entCount
  global localName opt prodName sheetLast timeStamp unicodeString worksheet worksheets wsCount
  global objDesign

  if {[catch {

# NAUO
    ::tcom::foreach nauo [$objDesign FindObjects [string trim next_assembly_usage_occurrence]] {
      foreach idx {4 5} {
        set pd($idx) [[[$nauo Attributes] Item [expr $idx]] Value]
        set p21id [$pd($idx) P21ID]

# product_definition name
        if {![info exists prodName($p21id)]} {
          set name [string trim [[[$pd($idx) Attributes] Item [expr 1]] Value]]
          set id1 "product_definition,id,$p21id"
          if {[info exists unicodeString($id1)]} {set name $unicodeString($id1)}

# or product name
          if {$name == "" || $name == "design" || $name == "part definition" || $name == "None" || $name == "UNKNOWN" || $name == "BLNMEYEN" || $name == "NON CONOSCIUTO" || $name == "NEZNM"} {
            set pdf [[[$pd($idx) Attributes] Item [expr 3]] Value]
            set pro [[[$pdf Attributes] Item [expr 3]] Value]
            set name [string trim [[[$pro Attributes] Item [expr 1]] Value]]
            set id1 "product,id,[$pro P21ID]"
            if {[info exists unicodeString($id1)]} {set name $unicodeString($id1)}
            if {$name == ""} {
              set name "NoName$p21id"
              errorMsg " Product name attribute is blank or contains non-English characters, using NoName...  See Help > Text Strings and Numbers" red
            }
          }
          set prodName($p21id) $name
        }

# relating
        set pname($idx) $prodName($p21id)
        if {$idx == 4} {
          set relating $pname($idx)
          set relatingID "$pname($idx) $p21id"
          lappend bomNAUO(relating) $relating
          lappend bomNAUOID(relating) $relatingID

# related
        } else {
          lappend bomAssembly($relating) $pname($idx)
          lappend bomAssemblyID($relatingID) "$pname($idx) $p21id"
          lappend bomNAUO(related) $pname($idx)
          lappend bomNAUOID(related) "$pname($idx) $p21id"
        }
      }
    }

    if {[info exists bomNAUO(relating)]} {
      outputMsg "Generating BOM worksheet" blue
      #outputMsg "\nbomNAUO relating $bomNAUO(relating)\nbomNAUO related $bomNAUO(related)" green
      #foreach idx [lsort [array names bomAssembly]] {outputMsg "assembly $idx / $bomAssembly($idx)"}
      #compareLists "relating related" [lrmdups $bomNAUO(relating)] [lrmdups $bomNAUO(related)]

# assembly structure (uses ID to account for assemblies and parts with the same names)
      set bomIndent 0
      set bom [intersect3 [lrmdups $bomNAUOID(relating)] [lrmdups $bomNAUOID(related)]]
      set root [lindex $bom 0]

      foreach asm $root {
        if {[info exists bomAssemblyID($asm)]} {
          set bomIndent 0
          #outputMsg [string range $asm 0 [string last " " $asm]-1] green
          foreach assem $bomAssemblyID($asm) {bomAssembly $assem}
        }
      }

# list assemblies
      set bom [intersect3 [lrmdups $bomNAUOID(relating)] [lrmdups $bomNAUOID(related)]]
      set assemblies [lindex $bom 1]
      set lassem {}
      foreach item [lsort -nocase $assemblies] {
        if {[info exists bomItems($item)] && [info exists bomAssemblyID($item)]} {
          foreach one $bomAssemblyID($item) {incr assemParts($one)}
          set str ""
          set n 0
          foreach one [lsort [array names assemParts]] {
            incr n
            append str "($assemParts($one)) [string range $one 0 [string last " " $one]-1]  "
            if {$n == 4} {append str [format "%c" 10]; set n 0}
          }
          lappend lassem [list $bomItems($item) [string range $item 0 [string last " " $item]-1] [string trim $str]]
          unset assemParts
        }
      }

# start BOM worksheet
      set parts [lindex $bom 2]
      if {[llength $parts] > 0} {
        set worksheet(BOM) [$worksheets Add [::tcom::na] $sheetLast]
        $worksheet(BOM) Activate
        $worksheet(BOM) Name BOM
        set cells(BOM) [$worksheet(BOM) Cells]
        set wsCount [$worksheets Count]
        [$worksheets Item [expr $wsCount]] -namedarg Move Before [$worksheets Item [expr 3]]

        set txt [file tail $localName]
        if {$timeStamp != ""} {append txt "  ($timeStamp)"}
        $cells(BOM) Item 1 1 $txt
        set range [$worksheet(BOM) Range A1 C1]
        $range MergeCells [expr 1]
        [$range Rows] AutoFit

# list parts
        set r 3
        $cells(BOM) Item $r 1 "Qty"
        foreach item [lsort -nocase $parts] {
          if {[info exists bomItems($item)]} {
            incr r
            $cells(BOM) Item $r 1 $bomItems($item)
            $cells(BOM) Item $r 2 "'[string range $item 0 [string last " " $item]-1]"
          }
        }
        $cells(BOM) Item 3 2 "Parts ([expr {$r-3}])"
        if {[info exists entCount(property_definition)] && [info exists entCount(property_definition_representation)]} {
          if {$entCount(property_definition) > 0 && $entCount(property_definition_representation) > 0} {
            if {$opt(valProp)} {
              set msg "See the property_definition worksheet"
            } else {
              set msg "Generate the Analyzer report for Validation Properties"
            }
            append msg " for possible properties associated with Parts."
            addCellComment "BOM" 3 2 $msg
          }
        }

# format table
        set range [$worksheet(BOM) Range A3 B$r]
        [$range Columns] AutoFit
        set tname "BOM-parts"
        [[$worksheet(BOM) ListObjects] Add 1 $range] Name $tname
        [[$worksheet(BOM) ListObjects] Item $tname] TableStyle "TableStyleLight8"
        set partWidth [[$worksheet(BOM) Range B3 B$r] ColumnWidth]

# list assemblies
        if {[llength $lassem] > 0} {
          incr r 2
          set rassem $r
          $cells(BOM) Item $r 1 "Qty"
          $cells(BOM) Item $r 2 "Assemblies ([llength $lassem])"
          $cells(BOM) Item $r 3 "Components"
          addCellComment "BOM" $r 2 "Assemblies do not necessarily contain all Parts."
          foreach item $lassem {
            incr r
            $cells(BOM) Item $r 1 [lindex $item 0]
            $cells(BOM) Item $r 2 "'[lindex $item 1]"
            $cells(BOM) Item $r 3 [lindex $item 2]
          }

# format table
          set tname "BOM-assemblies"
          set range [$worksheet(BOM) Range A$rassem C$r]
          [[$worksheet(BOM) ListObjects] Add 1 $range] Name $tname
          [[$worksheet(BOM) ListObjects] Item $tname] TableStyle "TableStyleLight9"

# column widths
          [$worksheet(BOM) Range C$rassem C$r] ColumnWidth [expr 150]
          [$range Columns] AutoFit
          [$range Rows] AutoFit
          $range VerticalAlignment [expr -4160]
          set assemWidth [[$worksheet(BOM) Range B$rassem B$r] ColumnWidth]
          if {$assemWidth < $partWidth} {[$worksheet(BOM) Range B1 B$r] ColumnWidth [expr $partWidth]}

# group parts list
          if {$rassem > 40} {[[$worksheet(BOM) Range A4 A[expr {$rassem-2}]] Rows] Group}
        }
      }
    }

# error
  } emsg]} {
    errorMsg "Error generating BOM worksheet: $emsg"
  }
  foreach var {bomAssembly bomAssemblyID bomIndent bomItems bomNAUO bomNAUOID prodName} {if {[info exists $var]} {unset -- $var}}
}

# -------------------------------------------------------------------------------
proc bomAssembly {assem} {
  global bomAssembly bomAssemblyID bomIndent bomItems lastAssem

  set assem  [join $assem]
  set assem1 [string range $assem 0 [string last " " $assem]-1]
  incr bomItems($assem)

  incr bomIndent 2
  if {$bomIndent > 100} {
    errorMsg " Problem with nesting components in an assembly"
    foreach var {bomAssembly bomIndent} {if {[info exists $var]} {unset -- $var}}
    return
  }

  if {![info exists lastAssem]} {set lastAssem ""}
  set str [string repeat " " $bomIndent]$assem1
  #outputMsg $str blue
  #if {$str != $lastAssem} {outputMsg $str blue}
  set lastAssem $str

  if {[info exists bomAssemblyID($assem)]} {
    foreach subassem [lsort -nocase $bomAssemblyID($assem)] {bomAssembly [join $subassem]}
  }
  incr bomIndent -2
  if {$bomIndent < 0} {set bomIndent 0}
}

# -------------------------------------------------------------------------------
proc pmiFormatColumns {str} {
  global cells col gpmiRow pmiStartCol recPracNames row spmiRow stepAP thisEntType vpmiRow worksheet

  if {![info exists pmiStartCol($thisEntType)]} {
    return
  } else {
    set c1 [expr {$pmiStartCol($thisEntType)-1}]
  }

# delete unused columns
  set delcol 0
  set colrange [[[$worksheet($thisEntType) UsedRange] Columns] Count]
  for {set i $colrange} {$i > 3} {incr i -1} {
    set val [[$cells($thisEntType) Item 3 $i] Value]
    if {$val != ""} {set delcol 1}
    if {$val == "" && $delcol} {
      set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
      $range Delete
    }
  }
  set col($thisEntType) [[[$worksheet($thisEntType) UsedRange] Columns] Count]

# format
  if {[info exists cells($thisEntType)] && $col($thisEntType) > $c1} {
    set c2 [expr {$c1+1}]
    set c3 $col($thisEntType)

# PMI heading
    outputMsg " [formatComplexEnt $thisEntType]"
    if {$str != "Validation Properties" || $thisEntType == "property_definition"} {
      $cells($thisEntType) Item 2 $c2 $str
      set range [$worksheet($thisEntType) Range [cellRange 2 $c2]]
      $range HorizontalAlignment [expr -4108]
      [$range Font] Bold [expr 1]
      [$range Interior] ColorIndex [expr 36]
      set range [$worksheet($thisEntType) Range [cellRange 2 $c2] [cellRange 2 $c3]]
      $range MergeCells [expr 1]
    }

# set rows for colors and borders
    set r1 1
    set r2 $r1
    set r3 {}
    if {[string first "PMI Presentation" $str] != -1} {
      set rs $gpmiRow($thisEntType)
    } elseif {[string first "Validation Properties" $str] != -1} {
      set rs $vpmiRow($thisEntType)
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

# cell color (yellow or green)
        set range [$worksheet($thisEntType) Range [cellRange $r1 $i] [cellRange $r2 $i]]
        [$range Interior] ColorIndex [lindex [list 36 35] [expr {$j%2}]]

# dotted line border
        if {$i == $c2 && $r2 > 3} {
          if {$r1 < 4} {set r1 4}
          set range [$worksheet($thisEntType) Range [cellRange $r1 $c2] [cellRange $r2 $c3]]
          for {set k 7} {$k <= 12} {incr k} {
            catch {if {$k != 9 || [expr {$row($thisEntType)+0}] != $r2} {[[$range Borders] Item [expr $k]] Weight [expr 1]}}
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
    if {$c1 > 2} {
      set range [$worksheet($thisEntType) Range [cellRange 1 2] [cellRange [expr {$row($thisEntType)+2}] $c1]]
      [$range Columns] Group
      set colrange [[[$worksheet($thisEntType) UsedRange] Columns] Count]

# entities with PMI
      if {$thisEntType == "dimensional_characteristic_representation"} {
        for {set i 1} {$i <= $colrange} {incr i} {
          set val [[$cells($thisEntType) Item 3 $i] Value]
          if {[string first "Associated Geometry" $val] != -1} {
            [[$worksheet($thisEntType) Range [cellRange 1 5] [cellRange [expr {$row($thisEntType)+2}] [expr {$i-1}]]] Columns] Group
          }
        }
      } elseif {[string first "tolerance" $thisEntType] != -1} {
        for {set i 1} {$i <= $colrange} {incr i} {
          set val [[$cells($thisEntType) Item 3 $i] Value]
          if {[string first "GD&T" $val] != -1} {
            set cgdt [expr {$i+1}]
          } elseif {[string first "Equivalent Unicode String" $val] != -1} {
            [[$worksheet($thisEntType) Range [cellRange 1 $cgdt] [cellRange [expr {$row($thisEntType)+2}] [expr {$i-1}]]] Columns] Group
          }
        }
      } elseif {[string first "annotation" $thisEntType] != -1} {
        for {set i 1} {$i <= $colrange} {incr i} {
          set val [[$cells($thisEntType) Item 3 $i] Value]
          if {[string first "Associated Geometry" $val] != -1} {
            [[$worksheet($thisEntType) Range [cellRange 1 6] [cellRange [expr {$row($thisEntType)+2}] [expr {$i-1}]]] Columns] Group
            set cag [expr {$i+1}]
          } elseif {[string first "Equivalent Unicode String" $val] != -1 && [info exists cag]} {
            [[$worksheet($thisEntType) Range [cellRange 1 $cag] [cellRange [expr {$row($thisEntType)+2}] [expr {$i-1}]]] Columns] Group
          }
        }
      }
    }

# fix column widths
    set colrange [[[$worksheet($thisEntType) UsedRange] Columns] Count]
    for {set i 1} {$i <= $colrange} {incr i} {
      set range [$worksheet($thisEntType) Range [cellRange -1 $i]]
      $range ColumnWidth [expr 96]
    }
    [$worksheet($thisEntType) Columns] AutoFit
    [$worksheet($thisEntType) Rows] AutoFit

# link to RP
    set str1 "pmi242"
    if {[string first "AP203" $stepAP] == 0} {set str1 "pmi203"}
    $cells($thisEntType) Item 2 1 "See CAx-IF Rec. Prac. for $recPracNames($str1)"
    if {$thisEntType != "dimensional_characteristic_representation" && $thisEntType != "datum_reference"} {
      set range [$worksheet($thisEntType) Range A2:D2]
    } else {
      set range [$worksheet($thisEntType) Range A2:C2]
    }
    $range MergeCells [expr 1]
    set anchor [$worksheet($thisEntType) Range A2]
    [$worksheet($thisEntType) Hyperlinks] Add $anchor [join "https://www.mbx-if.org/home/cax/recpractices/"] [join ""] [join "Link to CAx-IF Recommended Practices"]
  }
}

# -------------------------------------------------------------------------------
# check for an entity that is checked for semantic PMI
proc spmiCheckEnt {ent} {
  global opt spmiEntTypes tolNames
  set ok 0

# all tolerances, dimensions, datums, etc. (defined in sfa-data.tcl)
  if {!$opt(PMISEMDIM) && !$opt(PMISEMDT)} {
    foreach sp $spmiEntTypes {if {[string first $sp $ent] ==  0} {set ok 1}}
    foreach sp $tolNames     {if {[string first $sp $ent] != -1} {set ok 1}}
  }

# only dimensions or datum targets
  if {$opt(PMISEMDIM) && $ent == "dimensional_characteristic_representation"} {set ok 1}
  if {$opt(PMISEMDT) && ($ent == "placed_datum_target_feature" || $ent == "datum_target")} {set ok 1}

# counter holes
  if {([string first "counter" $ent] != -1 || [string first "spotface" $ent] != -1 || [string first "basic_round" $ent] != -1) && [string first "occurrence" $ent] == -1} {
    if {$ent != "spotface_definition"} {set ok 1}
  }
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

  if {[string first "leader" $ent] != -1 && $ent != "annotation_placeholder_occurrence_with_leader_line"} {set ok 0}
  if {[string first "over_riding_styled_item" $ent] != -1} {set ok 0}
  if {[string first "annotation_occurrence_associativity" $ent] != -1} {set ok 0}
  if {[string first "annotation_occurrence_relationship"  $ent] != -1} {set ok 0}

  return $ok
}

# -------------------------------------------------------------------------------
# which STEP entities are processed depending on options (most are already selected through Process categories)
proc setEntsToProcess {entType} {
  global objDesign
  global gen gpmiEnts spmiEnts opt

  set ok 0
  set gpmiEnts($entType) 0
  set spmiEnts($entType) 0

# for PMI (graphic) presentation report and view
  if {($opt(PMIGRF) || ($gen(View) && $opt(viewPMI))) && $ok == 0} {
    set ok [gpmiCheckEnt $entType]
    set gpmiEnts($entType) $ok
    if {$entType == "geometric_curve_set" || $entType == "planar_box" || [string first "geometric_set" $entType] != -1} {set ok 1}
  }

# for PMI (semantic) representation
  if {$opt(PMISEM) && $ok == 0} {
    set spmiEnts($entType) [spmiCheckEnt $entType]
    if {$entType == "axis2_placement_3d" && [$objDesign CountEntities "placed_datum_target_feature"] > 0} {set ok 1}
  }

# for tessellated geometry
  if {$gen(View) && $opt(viewTessPart) && $ok == 0} {if {[string first "tessellated" $entType] != -1} {set ok 1}}

  return $ok
}

# -------------------------------------------------------------------------------
# check for all types of reports
proc checkForReports {entType} {
  global cells gen gpmiEnts opt pmiColumns savedViewCol skipEntities spmiEnts stepAP stepAPreport

# check for validation properties report, call valPropStart
  if {$entType == "property_definition_representation" || $entType == "shape_definition_representation"} {
    if {[catch {
      if {[info exists opt(valProp)]} {
        if {$opt(valProp)} {
          if {[lsearch $skipEntities "representation"] == -1} {
            if {[info exists cells(property_definition)]} {valPropStart $entType}
          }
        }
      }
    } emsg]} {
      errorMsg "Error adding Validation Properties for $entType: $emsg"
    }

# check for PMI Presentation report or view graphic PMI, call gpmiAnnotation
  } elseif {$gpmiEnts($entType)} {
    if {[catch {
      set ok 0
      if {[info exists opt(PMIGRF)]} {if {$opt(PMIGRF)} {set ok 1}}
      if {[info exists opt(viewPMI)]} {if {$opt(viewPMI)} {set ok 1}}
      if {$ok} {
        if {[info exists cells($entType)] || $opt(viewPMI)} {gpmiAnnotation $entType}
        catch {unset savedViewCol}
        catch {unset pmiColumns}
      }
    } emsg]} {
      errorMsg "Error adding PMI Presentation for [formatComplexEnt $entType]: $emsg"
    }

# check for Semantic PMI reports
  } elseif {$spmiEnts($entType)} {
    if {[catch {
      if {[info exists opt(PMISEM)]} {
        if {$opt(PMISEM)} {
          if {[info exists cells($entType)]} {
            if {$stepAPreport} {

# dimensions
              if {$entType == "dimensional_characteristic_representation"} {
                spmiDimtolStart $entType

# hole occurrences
              } elseif {([string first "counter" $entType] != -1 || [string first "spotface" $entType] != -1 || [string first "basic_round" $entType] != -1) && [string first "occurrence" $entType] == -1} {
                if {$entType != "spotface_definition"} {spmiHoleStart $entType}

# geometric tolerances
              } else {
                spmiGeotolStart $entType
              }

# AP not supported
            } else {
              errorMsg " Analyzer reports for Semantic and Graphic PMI are not supported in $stepAP files." red
            }
          }
        }
      }
    } emsg]} {
      errorMsg "Error adding PMI Representation for [formatComplexEnt $entType]: $emsg"
    }

# check for AP209 entities that contain information to be processed for the viewer
  } elseif {$entType == "curve_3d_element_representation"   || \
            $entType == "surface_3d_element_representation" || \
            $entType == "volume_3d_element_representation"  || \
            $entType == "nodal_freedom_action_definition"   || \
            $entType == "nodal_freedom_values"              || \
            $entType == "surface_3d_element_boundary_constant_specified_surface_variable_value" || \
            $entType == "volume_3d_element_boundary_constant_specified_variable_value" || \
            $entType == "single_point_constraint_element_values"} {
    if {[catch {
      if {[info exists opt(viewFEA)]} {
        if {$gen(View) && $opt(viewFEA)} {
          set opt(x3dSave) 0
          if {[string first "element_representation" $entType] != -1 || \
              ($opt(feaBounds) && $entType == "single_point_constraint_element_values") || \
              ($opt(feaLoads) && \
                ($entType == "nodal_freedom_action_definition" || \
                 $entType == "surface_3d_element_boundary_constant_specified_surface_variable_value" || \
                 $entType == "volume_3d_element_boundary_constant_specified_variable_value")) || \
              ($opt(feaDisp) && $entType == "nodal_freedom_values")
          } {
            feaModel $entType
          }
        }
      }
    } emsg]} {
      errorMsg "Error adding FEM for [formatComplexEnt $entType]: $emsg"
    }
  }
}

# -------------------------------------------------------------------------------
proc setEntAttrList {abc} {
  global ent entAttrList entLevel opt

  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]
  if {$opt(DEBUG1)} {outputMsg "$ind PARSE"}

  set ni 0
  foreach item $abc {
    if {[llength $item] > 1} {
      setEntAttrList $item
    } else {
      if {$ni == 0} {
        set typ "ENT"
        set ent($entLevel) $item
      } else {
        set typ "  ATR"
        lappend entAttrList "$ent($entLevel) $item"
      }
      if {$opt(DEBUG1)} {
        if {$typ == "ENT"} {
          outputMsg "$ind $typ $entLevel $ni $item" blue
        } else {
          outputMsg "$ind $typ $entLevel $ni $item"
        }
      }
      incr ni
    }
  }
  incr entLevel -1
}

#-------------------------------------------------------------------------------
# run syntax checker with the command-line version (sfa-cl.exe) and output filtered result
proc syntaxChecker {fileName} {
  global buttons env ifcsvrDir opt wdir writeDir

  if {[file size $fileName] > 429000000} {outputMsg " The file is too large to run the Syntax Checker.  The limit is about 430 MB." red; return}

  set roseSchemas ""
  if {[info exists env(ROSE_SCHEMAS)]} {set roseSchemas $env(ROSE_SCHEMAS)}
  set env(ROSE_SCHEMAS) [file nativename $ifcsvrDir]

  if {[info exists buttons]} {
    outputMsg "\n[string repeat "-" 29]\nRunning Syntax Checker"
    set exe "STEP-File-Analyzer.exe"
  } else {
    outputMsg "\nRunning Syntax Checker"
    set exe "sfa-cl.exe"
  }

# check header
  set schema [getSchemaFromFile $fileName]
  if {$schema == ""} {return}

# get path for command-line version to run syntax checker
  set path [split $wdir "/"]
  set sfacl {}
  foreach item $path {
    if {$item != $exe} {
      lappend sfacl $item
    } else {
      break
    }
  }
  lappend sfacl "sfa-cl.exe"
  set sfacl [join $sfacl "/"]

# get syntax errors and warnings by running command-line version with stats option
  if {[file exists $sfacl]} {
    if {[info exists buttons]} {.tnb select .tnb.status}
    outputMsg "Syntax Checker results for: [file tail $fileName]"
    if {[catch {
      set sfaout [exec $sfacl [file nativename $fileName] stats nolog]
      set sfaout [split $sfaout "\n"]
      catch {unset sfaerr}
      set lineLast ""
      set paren 0
      foreach line $sfaout {

# get lines with errors and warnings
        if {[string first "error:" $line] != -1 || [string first "warning:" $line] != -1} {

# but not with these messages
          if {[string first "Converting 'integer' value" $line] == -1 && [string first "<Done>" $line] == -1} {
            if {$line != $lineLast} {append sfaerr " $line\n"}
            set lineLast $line
          }
          if {[string first "warning: No schemas" $line] != -1} {break}
          if {[string first "warning: Couldn't find schema" $line] != -1} {errorMsg "See Help > Supported STEP APs"}
          if {[string first "(" $line] != -1 && [string first ")" $line] != -1} {set paren 1}
        } elseif {[string first "Error opening" $line] != -1} {
          append sfaerr "$line "
        }
      }

# done
      if {[info exists sfaerr]} {
        outputMsg [string range $sfaerr 0 end-1] red
        if {$paren} {outputMsg "The number in parentheses is the line number in the file where the error or warning was detected."}

# output to log file
        if {$opt(logFile)} {
          set lfile [file rootname $fileName]
          if {$opt(writeDirType) == 2} {set lfile [file join $writeDir [file rootname [file tail $fileName]]]}
          append lfile "-sfa-err.log"
          set lf [open $lfile w]
          puts $lf "Syntax Checker results for: $fileName\nGenerated by the NIST STEP File Analyzer and Viewer [getVersion] ([string trim [clock format [clock seconds]]])\nThe number in parentheses is the line number in the STEP file where the error or warning was detected.\n"
          puts $lf [string range $sfaerr 0 end-1]
          close $lf
          outputMsg "Syntax Checker results saved to: [truncFileName [file nativename $lfile]]" blue
        }

# no errors
      } else {
        outputMsg " No syntax errors or warnings" green
      }
      if {[info exists buttons]} {outputMsg "See Help > Syntax Checker"}

# error running syntax checker
    } emsg]} {
      errorMsg " Syntax Checker failed: $emsg" red
    }
  } else {
    outputMsg " Syntax Checker cannot be run.  Make sure the command-line version 'sfa-cl.exe' is in the same directory as 'STEP-File-Analyzer.exe" red
  }

  unset env(ROSE_SCHEMAS)
  if {$roseSchemas != ""} {set env(ROSE_SCHEMAS) $roseSchemas}

  if {[info exists buttons]} {outputMsg "[string repeat "-" 29]"}
}

# -------------------------------------------------------------------------------
# get STEP AP name
proc getStepAP {fname} {
  global fileSchema opt stepAPs useXL

  set ap ""
  set limit 0
  if {![info exists useXL]} {set useXL 1}
  if {$opt(xlFormat) != "None" && $useXL} {set limit 1}
  set fs [string toupper [getSchemaFromFile $fname $limit]]
  set fileSchema $fs

  set c1 [string first " " $fs]
  if {$c1 != -1} {set fs [string range $fs 0 $c1-1]}
  if {[string first "AP2" $fs] == 0} {
    set ap [string range $fs 0 4]
  } elseif {[info exists stepAPs($fs)]} {
    set ap $stepAPs($fs)
  } else {
    set ap $fileSchema
  }

# check AP242 edition
  if {$ap == "AP242"} {
    if {[string first "442 1 1 4" $fileSchema] != -1} {
      append ap "e1"
    } elseif {[string first "442 2 1 4" $fileSchema] != -1 || [string first "442 3 1 4" $fileSchema] != -1} {
      append ap "e2"
    } elseif {[string first "442 4 1 4" $fileSchema] != -1} {
      append ap "e3"
    }
  }

# check AP214 edition
  if {$ap == "AP214"} {
    if {[string first "214 1 1 1" $fileSchema] != -1} {
      append ap "e1"
    } elseif {[string first "214 3 1 1" $fileSchema] != -1} {
      append ap "e3"
    }
  }
  return $ap
}

#-------------------------------------------------------------------------------
proc getSchemaFromFile {fname {limit 0}} {
  global cadApps cadSystem developer opt p21e3 rawBytes timeStamp unicodeInFile useXL

  set p21e3 0
  set schema ""
  set fsline ""
  set ok 0
  set ok1 0
  set nline 0
  set nendsec 0
  set filename 0
  set unicodeInFile 0
  set stepfile [open $fname r]
  catch {unset rawBytes}

  set ulimit 100
  if {$limit} {set ulimit 500000}

# read first N lines, all HEADER section information should be in the first 100 lines, reading more detects \X2\ Unicode characters
  while {[gets $stepfile line] != -1 && $nline < $ulimit} {
    incr nline

# check file
    if {[string first "ISO-10303-21;" $line] != -1} {set ok1 1}

# check for filename
    if {[string first "FILE_NAME" $line] != -1} {set filename 1}

# check for CAD apps
    if {$filename && $nendsec == 0} {
      if {![info exists cadSystem]} {
        foreach app $cadApps {
          if {[string first $app $line] != -1} {
            set cadSystem $app
            if {$app == "SolidWorks"} {
              set c1 [string first "SolidWorks 20" $line]
              if {$c1 != -1} {set cadSystem [string range $line $c1 $c1+14]}
            } elseif {$app == "Autodesk Inventor"} {
              set c1 [string first "Autodesk Inventor 20" $line]
              if {$c1 != -1} {set cadSystem [string range $line $c1 $c1+21]}
            }
            break
          }
        }
      }

# check for time stamp
      if {![info exists timeStamp]} {
        foreach year {199 200 201 202 203} {
          set c1 [string first "'$year" $line]
          if {$c1 != -1} {
            set c2 [string first "'" [string range $line $c1+1 end]]
            set timeStamp [string range $line $c1+1 $c1+$c2]
            if {[string index $timeStamp 4] != "-"} {unset timeStamp}
          }
        }
      }
    }

# check for X2 Unicode control directives when generating a spreadsheet, set xlUnicode if possible, \X\ does not have to be handled separately for a spreadsheet
    if {[string first "\\X2\\" $line] != -1 && $opt(xlFormat) != "None" && $useXL} {
      if {[info exists schema]} {if {$schema == "CUTTING_TOOL_SCHEMA_ARM" || [string first "ISO13" $schema] == 0} {set opt(xlUnicode) 1}}
      if {[file size $fname] <= 10000000} {set opt(xlUnicode) 1}
      set unicodeInFile 1
      if {!$opt(xlUnicode) && $limit} {errorMsg "Symbols or non-English text found for some entity attributes.  See the More tab to process those symbols and characters.  Also see Help > Text Strings and Numbers." red}
    }

# check for OPTIONS comment
    if {[string first "/* OPTION:" $line] == 0} {
      if {[string first "raw bytes" $line] != -1 || ($developer && [string first "custom schema-name" $line] == -1)} {
        set emsg "HEADER section comment: [string range $line 11 end-3]"
        if {[string first "raw bytes" $emsg] != -1} {
          append emsg " (See Help > Text Strings and Numbers)\nThis might affect Parts displayed in the Viewer."
          set rawBytes 1
        }
        errorMsg $emsg red
      }
    }

# look for FILE_SCHEMA
    if {[string first "FILE_SCHEMA" $line] != -1} {
      set ok 1
      set fsline $line

# done reading header section
    } elseif {[string first "ENDSEC" $line] != -1 && $nendsec == 0} {

# missing file schema
      if {$fsline == ""} {
        errorMsg "Missing FILE_SCHEMA in HEADER section"
        return
      }

# check for double parentheses
      if {[string first "(" $fsline] == [string last "(" $fsline] || [string first ")" $fsline] == [string last ")" $fsline]} {
        errorMsg "FILE_SCHEMA must use a double set of parentheses."
      }

# get schema
      set fsline [string range $fsline 0 [string first ";" $fsline]]
      set sline [split $fsline "'"]
      set schema [lindex $sline 1]
      incr nendsec
      if {$schema == ""} {errorMsg "FILE_SCHEMA schema name is blank.  See Help > Supported STEP APs for supported schema names."}

# multiple schemas
      if {[string first "," $fsline] != -1} {
        regsub -all " " $fsline "" fsline
        set schema [string range $fsline [string first "'" $fsline] [string last "'" $fsline]]
      }
      if {$p21e3} {break}

# check for Part 21 edition 3 files
    } elseif {[string first "4\;1" $line] != -1 || [string first "ANCHOR\;" $line] != -1 || \
              [string first "REFERENCE\;" $line] != -1 || [string first "SIGNATURE\;" $line] != -1} {
      set p21e3 1
      if {[string first "4\;1" $line] == -1} {break}

    } elseif {$ok} {
      append fsline $line
    }
  }
  close $stepfile

# not a STEP file
  if {!$ok1} {outputMsg "The file does not start with 'ISO-10303-21;' and is probably not a STEP file.  See Websites > STEP" red}
  return $schema
}

#-------------------------------------------------------------------------------
# convert \X2\ in strings, see sfa-data.tcl for unicodeAttributes
# \X\ does not have to be handled separately for a spreadsheet, complex entities need to be handled explicitly
proc unicodeStrings {unicodeEnts} {
  global driUnicode localName unicodeActual unicodeAttributes unicodeNumEnts unicodeString

  if {[catch {
    set nent 0
    set okread 1
    set unicodeActual {}
    set sf [open $localName r]

    while {[gets $sf line] >= 0 && $okread} {
      while {[string first ";" $line] == -1} {gets $sf line1; append line $line1}

      set ok 0
      set c1 [string first "=" $line]
      set c2 [string first "(" $line]
      set str [string trim [string range $line $c1+1 $c2-1]]
      foreach ent $unicodeEnts {
        if {$str == $ent} {
          set ok 1
        } elseif {[string first "COMPOSITE_SHAPE_ASPECT()DATUM_FEATURE" $line] != -1} {
          set ok 1
          set ent "COMPOSITE_SHAPE_ASPECT_AND_DATUM_FEATURE"
        }
        if {$ok} {
          set ent1 [string tolower $ent]
          break
        }
      }

# process Unicode
      if {$ok} {
        set lattr [llength $unicodeAttributes($ent1)]
        incr nent

# check for Unicode X2
        set oku 0
        set cx2 [string first "\\X2\\" $line]
        if {$cx2 != -1} {
          set oku 1
        } elseif {[string first "equivalent" $line] != -1} {
          if {[string first "\\w" $line] != -1 || [string first "\\n" $line] != -1 || [string first "\\u" $line] != -1 || [string first "\\x" $line] != -1} {set oku 1}
        }
        if {$oku} {
          if {$cx2 != -1} {errorMsg "\nProcessing Unicode characters (See Help > Text Strings and Numbers)" blue}

          set id [string trim [string range $line 1 [string first "=" $line]-1]]
          set idx "$ent1,[lindex $unicodeAttributes($ent1) 0],$id"

# get strings that have Unicode and convert
          switch -- $ent1 {
            item_names {
# entity attributes from iso13...
              if {[string first "LABEL" $line] != -1} {
                set str [string range $line [string first "'" $line]+1 [string first ")" $line]-2]
              } else {
                set str [string range $line [string first "'" $line]+1 [string first "," $line]-2]
              }
              set unicodeString($idx) [getUnicode $str "attr"]
              if {[lsearch $unicodeActual $ent1] == -1} {lappend unicodeActual $ent1}
            }
            string_with_language {
              set str [string range $line [string first "'" $line]+1 [string last "'" $line]-1]
              set unicodeString($idx) [getUnicode $str "attr"]
              if {[lsearch $unicodeActual $ent1] == -1} {lappend unicodeActual $ent1}
            }
            translated_label -
            translated_text {
              set str1 [string range $line [string first "('" $line]+2 [string first "')" $line]-1]
              set str1 [getUnicode $str1 "attr"]
              set str {}
              set len [string length $str1]
              for {set i 0} {$i < $len} {incr i} {
                if {[string index $str1 $i] == "'" && [string index $str1 $i+1] == ","} {
                  lappend str [string range $str1 0 $i-1]
                  lappend str [string range $str1 $i+3 end]
                  break
                }
              }
              regsub -all "''" $str "'" str
              set unicodeString($idx) $str
              if {[lsearch $unicodeActual $ent1] == -1} {lappend unicodeActual $ent1}
            }

            default {
# entity attributes from ap2..
              catch {unset attr}
              set na 0
              set line1 [getUnicode $line "attr"]
              set i1 [expr {[string first "(" $line1]+1}]
              set i2 [string last  ")" $line1]
              for {set i $i1} {$i < $i2} {incr i} {
                append attr($na) [string index $line1 $i]
                if {([string index $line1 $i] == "'" && [string index $line1 $i+1] == ",") || \
                    ($i == $i1 && [string index $line1 $i] == "$")} {
                  incr i
                  incr na
                  if {$na > $lattr} {break}
                }
              }
              foreach ia [array names attr] {
                set str [string trim $attr($ia)]
                if {[lindex $unicodeAttributes($ent1) $ia] == "reference_designator"} {set str [string range $str [string first "'" $str] [string last "'" $str]]}
                set c0 [string index $str 0]
                if {$c0 != "$" && $c0 != "#"} {
                  set c1 [string first "'" $str]
                  set c2 [string last  "'" $str]
                  if {$c1 != -1 && $c2 != -2} {set str [string range $str $c1+1 $c2-1]}
                  set idx "$ent1,[lindex $unicodeAttributes($ent1) $ia],$id"
                  if {[string index $str 0] == "="} {set str " $str"}
                  set unicodeString($idx) $str
                }
              }
              if {[lsearch $unicodeActual $ent1] == -1} {lappend unicodeActual $ent1}
            }
          }
        }

# stop reading
        if {$nent == $unicodeNumEnts} {set okread 0}
      }
    }

# report entity types with Unicode
    if {[llength $unicodeActual] > 0} {
      regsub -all " " [join [lsort $unicodeActual]] ", " str
      outputMsg " [llength $unicodeActual] entity type(s): $str"
      if {[lsearch $unicodeActual "descriptive_representation_item"] != -1} {set driUnicode 1}
    }

  } emsg]} {
    errorMsg "Error processing Unicode string attribute: $emsg"
  }
}

# -------------------------------------------------------------------------------
# process X and X2 control directives with Unicode characters
proc getUnicode {ent {type "view"}} {
  global fontErr mytemp noFontFile

  set x "&#x"
  set u "\\u"
  set z "00"
  set msgChar5 0
  set ent1 $ent

  foreach xl [list X X2] {
    set cx [string first "\\$xl\\" $ent]
    if {$cx != -1} {
      switch -- $xl {
        X {
          while {$cx != -1} {
            set xu ""
            set uc [string range $ent $cx+3 $cx+4]
            switch -- $type {
              view {append xu "$x$uc;"}
              attr {append xu [join [eval list $u$z$uc]]}
            }
            set ent [string range $ent 0 $cx-1]$xu[string range $ent $cx+5 end]
            set cx [string first "\\$xl\\" $ent]
          }
        }
        X2 {
          while {$cx != -1} {
            set xu ""
            for {set i 4} {$i < 200} {incr i 4} {
              set uc [string range $ent $cx+$i [expr {$cx+$i+3}]]

# change flatness and straightness Unicode so that they look OK in the spreadsheet
              if {$uc == "23E5"} {set uc "25B1"}
              if {$uc == "23E4"} {set uc "2212"}

# check for font file with GD&T symbols (ARIALUNI.TTF), needed only for certain Unicode characters, noFontFile set in genExcel
              if {$noFontFile && ![info exists fontErr]} {
                set ok 0
                foreach char [list 232D 232F 232E 24C4 24CA 24BB 24C9 24C1 24BA 24BE 24C7 24C8 24B6] {if {[string first $char $uc] != -1} {set ok 1; break}}
                if {$ok} {
                  errorMsg "Some GD&T symbols will appear as a question mark on the descriptive_representation_item worksheet.\n To fix the problem, copy the font file that contains the symbols\n [file join $mytemp ARIALUNI.TTF]  to  C:/Windows/Fonts\n You might need administrator privileges."
                  set fontErr 1
                }
              }

# check for double backslash instead of single backslash
              set cx2 [string first "\\\\X2\\\\" $ent]
              if {$cx2 != -1} {
                errorMsg " Double backslash '\\\\X2\\\\' is not valid for delimiting Unicode characters with \\X2\\ and \\X0\\."
                set uc [string range $ent [expr {$cx+$i+1}] [expr {$cx+$i+4}]]
              }

# check for Unicode using five characters
              set xcheck [string range $ent [expr {$cx+$i+5}] [expr {$cx+$i+8}]]
              if {$xcheck == "\\X0\\"} {
                set msgChar5 1
                set uc "[string range $ent $cx+$i [expr {$cx+$i+4}]]"
                incr i
              }

              if {$type == "attr" && $uc == "000A"} {
                set xu [format "%c" 10]
              } elseif {[string first "\\" $uc] == -1} {

# add $x or $u delimiter depending if for a view or spreadsheet
                switch -- $type {
                  view {append xu "$x$uc;"}
                  attr {
                    if {[string length $uc] == 4} {
                      append xu [join [eval list $u$uc]]
                    } else {
                      append xu [join [eval list $uc]]
                      append xu " "
                    }
                  }
                }
              } else {
                set cx0 [string first "\\X0\\" $ent]
                if {$cx0 != -1} {
                  if {$cx2 == -1} {
                    set ent "[string range $ent 0 $cx-1]$xu[string range $ent $cx0+4 end]"
                  } else {
                    set ent "[string range $ent 0 $cx-2]$xu[string range $ent $cx0+5 end]"
                  }
                  break
                } else {
                  errorMsg " For encoding a Unicode character, \\X0\\ is missing to close an \\X2\\"
                  return $ent
                }
              }
            }
            set cx [string first "\\$xl\\" $ent]
          }
        }
      }
    }
  }

# check equivalent Unicode strings
  if {[string first "DESCRIPTIVE_REPRESENTATION_ITEM" $ent] != -1} {
    set c1 [string first "equivalent" $ent]
    if {$c1 != -1} {getEquivUnicodeString $c1 $ent $ent1 $msgChar5}
  }
  return $ent
}

#-------------------------------------------------------------------------------
# substitute \w and \n, and save the equivalent Unicode string
proc getEquivUnicodeString {c1 ent ent1 msgChar5} {
  global equivUnicodeString equivUnicodeStringErr syntaxErr

# format string
  if {[catch {
    regsub -all {\\\\} $ent {\\} ent
    set eus [string range $ent $c1+28 end-3]
    regsub -all {\\w} $eus " | " eus
    regsub -all {\\n} $eus [format "%c" 10] eus
    foreach tag {"DIM |" "FCF |" "TXT |"} {
      set c2 [string first $tag $eus]
      if {$c2 > 0} {
        if {[string index $eus $c2-1] != [format "%c" 10]} {
          set eus [string range $eus 0 $c2-1][format "%c" 10][string range $eus $c2 end]
        }
      }
    }
    set num [string trim [string range $ent 1 [string first "=" $ent]-1]]
    set equivUnicodeString($num) $eus

# check for errors
    set msg ""
    set nu 0
    set upos 0
    while {[string first "\\u" $ent $upos] != -1} {
      incr nu
      set upos [expr {[string first "\\u" $ent $upos]+2}]
    }
    if {[expr {$nu%2}] != 0} {set msg "There should be an even number of '\\u' delimiters"}

    if {$msg == ""} {foreach tag {"<COUNTERBORE" "DIM\\\\wHOLE"} {if {[string first $tag $ent1] != -1} {set msg "Unexpected keyword"}}}
    if {$msg == ""} {
      foreach tag {DIMENSION DATUM POS TARGET SL_PROF TEXT FLAG_NOTE APPCSR} {
        append tag "_V1"
        if {[string first $tag $ent1] != -1} {set msg "Unexpected keyword"}
      }
    }
    if {$msg == ""} {
      foreach char {2300 24C4 24CA 24C1 24C2 24BB 24C9} {
        if {$msg == "" && [string first "\\w\\X2\\$char\\X0\\\\\\w" $ent1] != -1} {set msg "Symbol is in its own compartment"}
      }
    }
    if {$msg == ""} {foreach char {F055 F056} {if {[string first "\\X2\\$char\\X0\\" $ent1] != -1} {set msg "Unicode character $char is not valid"}}}
    if {$msg == "" && [string first "\\X2\\2501\\X0\\" $ent1] != -1} {set msg "Unexpected Unicode character 2501"}
    if {$msg == "" && [string first "\\X2\\2313\\X0\\AAS" $ent1] != -1} {set msg "AAS is in the wrong position"}
    if {$msg == "" && [string first "\\\\s" $ent1] != -1} {set msg "Unexpected delimiter \\s"}
    if {$msgChar5} {set msg "Unicode using five characters is not supported"}

# report errors
    if {$msg != ""} {
      if {![info exists equivUnicodeStringErr]} {set equivUnicodeStringErr {}}
      if {[lsearch $equivUnicodeStringErr $msg] == -1} {lappend equivUnicodeStringErr $msg}
      lappend syntaxErr(descriptive_representation_item) [list $num "Equivalent Unicode String" $msg]
    }
  } emsg]} {
    errorMsg "Error getting equivalent Unicode string: $emsg"
  }
}

#-------------------------------------------------------------------------------
proc checkP21e3 {fname} {
  global p21e3Section

  set p21e3Section {}
  set p21e3 0
  set nline 0
  set f1 [open $fname r]

# check for Part 21 edition 3 file
  while {[gets $f1 line] != -1} {
    if {[string first "DATA\;" $line] == 0} {
      set nname $fname
      break
    } elseif {[string first "ANCHOR\;" $line] == 0 || \
              [string first "REFERENCE\;" $line] == 0 || \
              [string first "SIGNATURE\;" $line] == 0} {
      set p21e3 1
      break
    }
  }
  close $f1

# Part 21 edition 3 file
  if {$p21e3} {

# new file name (now -mod, previously -p21e2)
    set oname "[file rootname $fname]-p21e2[file extension $fname]"
    catch {file delete -force -- $oname}
    set nname "[file rootname $fname]-mod[file extension $fname]"
    catch {file delete -force -- $nname}
    set f2 [open $nname w]

# read file
    set write 1
    set data 0
    set sects {}

    set f1 [open $fname r]
    while {[gets $f1 line] != -1} {
      if {!$data} {
        if {[string first "DATA\;" $line] == 0} {
          set write 1
          set data 1
          regsub -all " " [join $sects] " and " sects
          outputMsg " "
          errorMsg "The STEP file uses ISO 10303 Part 21 Edition *3* '$sects' section(s)."
          outputMsg " A modified Part 21 Edition *2* file will be written and processed.\n  [truncFileName [file nativename $nname]]"
          outputMsg " See Help > User Guide (section 5.6)"

# check for Part 21 edition 3 content
        } elseif {[string first "ANCHOR\;" $line] == 0 || \
                  [string first "REFERENCE\;" $line] == 0 || \
                  [string first "SIGNATURE\;" $line] == 0} {
          set write 0
          lappend sects [string range $line 0 end-1]
        }

# write new file w/o Part 21 edition 3 content, change 4;1 to 2;1
        if {$write} {
          set c1 [string first "4\;1" $line]
          if {$c1 != -1} {set line [string replace $line $c1 $c1 2]}
          puts $f2 $line
        } else {
          lappend p21e3Section [string range $line 0 end-1]
        }

# in DATA section
      } else {
        puts $f2 $line
      }
    }
    close $f1
    close $f2
  }
  return $nname
}

# -------------------------------------------------------------------------------
# CAx-IF vendor abbrevitations, allVendor defined in sfa-data.tcl
proc setCAXIFvendor {} {
  global allVendor localName

  set fn [file tail $localName]
  set chars [list "-" "_" "."]

  foreach idx [lsort [array names allVendor]] {
    if {[string first $idx $fn] != -1} {
      foreach c1 $chars {
        if {[string first $c1$idx $fn] != -1} {
          foreach c2 $chars {
            if {[string first $c1$idx$c2 $fn] != -1} {
              return $allVendor($idx)
            }
          }
        }
      }
    }
  }
}
