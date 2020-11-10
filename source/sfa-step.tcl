proc pmiFormatColumns {str} {
  global cells col gpmiRow pmiStartCol recPracNames row spmiRow stepAP thisEntType worksheet

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
      if {![file exists [file nativename C:/Windows/Fonts/ARIALUNI.TTF]]} {
        errorMsg "Excel might not show some GD&T symbols correctly in PMI Representation analysis.  The missing\nsymbols will appear as question mark inside a square.  The likely cause is a missing font\n'Arial Unicode MS' from the font file 'ARIALUNI.TTF'.  Find a copy of this font file and install it."
      }
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
    [$worksheet($thisEntType) Hyperlinks] Add $anchor [join "https://www.cax-if.org/cax/cax_recommPractice.php"] [join ""] [join "Link to CAx-IF Recommended Practices"]
  }
}

# -------------------------------------------------------------------------------
# check for an entity that is checked for semantic PMI
proc spmiCheckEnt {ent} {
  global opt spmiEntTypes tolNames
  set ok 0

# all tolerances, dimensions, datums, etc. (defined in sfa-data.tcl)
  if {!$opt(PMISEMDIM)} {
    foreach sp $spmiEntTypes {if {[string first $sp $ent] ==  0} {set ok 1}}
    foreach sp $tolNames     {if {[string first $sp $ent] != -1} {set ok 1}}

# only dimensions
  } elseif {$ent == "dimensional_characteristic_representation"} {
    set ok 1
  }

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

  if {[string first "leader" $ent] != -1} {set ok 0}
  if {[string first "over_riding_styled_item" $ent] != -1} {set ok 0}
  if {[string first "annotation_occurrence_associativity" $ent] != -1} {set ok 0}
  if {[string first "annotation_occurrence_relationship"  $ent] != -1} {set ok 0}

  return $ok
}

# -------------------------------------------------------------------------------
# which STEP entities are processed depending on options
proc setEntsToProcess {entType} {
  global objDesign
  global gpmiEnts spmiEnts opt

  set ok 0
  set gpmiEnts($entType) 0
  set spmiEnts($entType) 0

# for validation properties
  if {$opt(valProp)} {
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

# for PMI (graphical) presentation report and viz
  if {($opt(PMIGRF) || $opt(viewPMI)) && $ok == 0} {
    set ok [gpmiCheckEnt $entType]
    set gpmiEnts($entType) $ok
    if {$entType == "advanced_face" || \
        $entType == "characterized_item_within_representation" || \
        $entType == "colour_rgb" || \
        $entType == "curve_style" || \
        $entType == "fill_area_style" || \
        $entType == "fill_area_style_colour" || \
        $entType == "geometric_curve_set" || \
        $entType == "geometric_set" || \
        $entType == "presentation_style_assignment" || \
        $entType == "property_definition" || \
        $entType == "representation_relationship" || \
        $entType == "view_volume"} {
      set ok 1
    }
    foreach ent {"annotation" "draughting" "_presentation" "camera" "constructive_geometry" "tessellated_geometric_set"} {
      if {[string first $ent $entType] != -1} {set ok 1}
    }
    if {!$ok} {if {[string first "representation" $entType] == -1 && [string first "presentation_" $entType] != -1} {set ok 1}}
  }

# for PMI (semantic) representation
  if {$opt(PMISEM) && $ok == 0} {
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

# for tessellated geometry
  if {$opt(viewTessPart) && $ok == 0} {if {[string first "tessellated" $entType] != -1} {set ok 1}}

  return $ok
}

# -------------------------------------------------------------------------------
# check for all types of reports
proc checkForReports {entType} {
  global cells gpmiEnts opt pmiColumns savedViewCol skipEntities spmiEnts

# check for validation properties report, call valPropStart
  if {$entType == "property_definition_representation"} {
    if {[catch {
      if {[info exists opt(valProp)]} {
        if {$opt(valProp)} {
          if {[lsearch $skipEntities "representation"] == -1} {
            if {[info exists cells(property_definition)]} {valPropStart}
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding Validation Properties for $entType: $emsg"
    }

# check for PMI Presentation report or viz graphical PMI, call gpmiAnnotation
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
      errorMsg "ERROR adding PMI Presentation for [formatComplexEnt $entType]: $emsg"
    }

# viz tessellated part geometry, call tessPart
  } elseif {$entType == "tessellated_solid" || $entType == "tessellated_shell" || $entType == "tessellated_wire"} {
    if {[catch {
      if {[info exists opt(viewTessPart)]} {if {$opt(viewTessPart)} {tessPart $entType}}
    } emsg]} {
      errorMsg "ERROR adding Tessellated Part Geometry: $emsg"
    }

# check for Semantic PMI reports
  } elseif {$spmiEnts($entType)} {
    if {[catch {
      if {[info exists opt(PMISEM)]} {
        if {$opt(PMISEM)} {
          if {[info exists cells($entType)]} {

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
          }
        }
      }
    } emsg]} {
      errorMsg "ERROR adding PMI Representation for [formatComplexEnt $entType]: $emsg"
    }

# check for AP209 analysis entities that contain information to be processed for view
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
        if {$opt(viewFEA)} {
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

# for results at element nodes
      #$entType == "element_nodal_freedom_actions"
      #($opt(feaLoads) && ($entType == "nodal_freedom_action_definition" || $entType == "element_nodal_freedom_actions"))
    } emsg]} {
      errorMsg "ERROR adding FEM for [formatComplexEnt $entType]: $emsg"
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
  global sfacl opt writeDir

  outputMsg "\n[string repeat "-" 29]\nRunning Syntax Checker"

# check header
  getSchemaFromFile $fileName

# get syntax errors and warnings by running command-line version with stats option
  if {[file exists $sfacl]} {
    .tnb select .tnb.status
    outputMsg "Syntax Checker results for: [file tail $fileName]"
    if {[catch {
      set sfaout [exec $sfacl [file nativename $fileName] stats nolog]
      set sfaout [split $sfaout "\n"]
      catch {unset sfaerr}
      set lineLast ""
      foreach line $sfaout {

# get lines with errors and warnings
        if {[string first "error:" $line] != -1 || [string first "warning:" $line] != -1} {

# but not with these messages
          if {[string first "Converting 'integer' value" $line] == -1 && \
              [string first "ROSE_RUNTIME" $line] == -1 && \
              [string first "End ST-Developer" $line] == -1} {
            if {$line != $lineLast} {append sfaerr " $line\n"}
            set lineLast $line
          }
          if {[string first "warning: No schemas" $line] != -1} {break}
          if {[string first "warning: Couldn't find schema" $line] != -1} {
            errorMsg "See Help > Supported STEP APs"
          }
        } elseif {[string first "ERROR opening" $line] != -1} {
          append sfaerr "$line "
        }
      }

# done
      if {[info exists sfaerr]} {
        outputMsg [string range $sfaerr 0 end-1] red

# output to log file
        if {$opt(logFile)} {
          set lfile [file rootname $fileName]
          if {$opt(writeDirType) == 2} {set lfile [file join $writeDir [file rootname [file tail $fileName]]]}
          append lfile "-sfa-err.log"
          set lf [open $lfile w]
          puts $lf "Syntax Checker results for: $fileName\nGenerated by the NIST STEP File Analyzer and Viewer [getVersion] ([string trim [clock format [clock seconds]]])\n"
          puts $lf $sfaerr
          close $lf
          outputMsg "Syntax Checker results saved to: [truncFileName [file nativename $lfile]]" blue
        }

# no errors
      } else {
        outputMsg " No syntax errors or warnings" green
      }
      outputMsg "See Help > Syntax Checker"

# error running syntax checker
    } emsg]} {
      errorMsg " Syntax Checker failed: $emsg" red
    }
  } else {
    outputMsg " Syntax Checker cannot be run.  Make sure the command-line version 'sfa-cl.exe' is in the same directory as 'STEP-File-Analyzer.exe" red
  }

  outputMsg "[string repeat "-" 29]"
}

# -------------------------------------------------------------------------------
# get STEP AP name
proc getStepAP {fname} {
  global fileSchema stepAPs

  set ap ""
  set fs [string toupper [getSchemaFromFile $fname]]
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
    if {[string first "442 2 1 4" $fileSchema] != -1 || [string first "442 3 1 4" $fileSchema] != -1} {
      append ap "e2"
    } else {
      append ap "e1"
    }
  }
  return $ap
}

#-------------------------------------------------------------------------------
proc getSchemaFromFile {fname {msg 0}} {
  global cadApps cadSystem developer p21e3 timeStamp unicode

  set p21e3 0
  set schema ""
  set ok 0
  set ok1 0
  set nline 0
  set niderr 0
  set nendsec 0
  set filename 0
  set stepfile [open $fname r]

# read first 100 lines
  while {[gets $stepfile line] != -1 && $nline < 100} {
    if {$msg} {
      foreach item {"MIME-Version" "Content-Type" "X-MimeOLE" "DOCTYPE HTML" "META content"} {
        if {[string first $item $line] != -1} {
          errorMsg "Syntax Error: The STEP file was probably saved as an EMAIL or HTML file.  The STEP file cannot be translated.\n In the email client, save the STEP file as a TEXT file and try again.\n The first line in the STEP file should be 'ISO-10301-21\;'"
        }
      }
    }
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

# check for X and X2 control directives
    if {[string first "\\X\\" $line] != -1 || [string first "\\X2\\" $line] != -1} {
      errorMsg "\\X2\\ or \\X\\ control directives are used in some text strings.  See Help > Text Strings" red
      set unicode 1
    }

# check for OPTIONS from ST-Developer toolkit
    if {[string first "/* OPTION:" $line] == 0} {
      set emsg "HEADER section comment: "
      if {[string first "raw bytes" $line] != -1 || ($developer && [string first "custom schema-name" $line] == -1)} {
        set emsg "HEADER section comment: [string range $line 11 end-3]"
        if {[string first "raw bytes" $emsg] != -1} {append emsg " (See Help > Text Strings)"}
        errorMsg $emsg red
      }
    }

# look for FILE_SCHEMA
    if {[string first "FILE_SCHEMA" $line] != -1} {
      set ok 1
      set fsline $line

# done reading header section
    } elseif {[string first "ENDSEC" $line] != -1 && $nendsec == 0} {

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

# check for IDs >= 2^31, valid but will be a different number in the spreadsheet
    } elseif {[string first "#" $line] == 0} {
      set id [string range $line 1 [string first "=" $line]-1]
      if {$id > 2147483647 && $niderr == 0} {
        errorMsg "An entity ID (#$id) >= 2147483648 (2^31)\n Very large IDs are valid but will appear as different numbers in the spreadsheet."
        incr niderr
      }

# check for part 21 edition 3 files
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
  if {!$ok1} {errorMsg "ERROR: The STEP file does not start with ISO-10303-21;"}
  return $schema
}

#-------------------------------------------------------------------------------
proc checkP21e3 {fname} {
  global p21e3Section

  set p21e3Section {}
  set p21e3 0
  set nline 0
  set f1 [open $fname r]

# check for part 21 edition 3 file
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

# part 21 edition 3 file
  if {$p21e3} {

# new file name
    set nname "[file rootname $fname]-p21e2[file extension $fname]"
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
          outputMsg " This software cannot directly process Edition 3 files.  A new Part 21 Edition *2* file:" red
          outputMsg "   [truncFileName $nname]"
          outputMsg " without those sections will be written and processed." red
          outputMsg " The '$sects' section(s) from the original file will be processed separately for the spreadsheet.\n See Help > User Guide (section 5.7)\n See Websites > STEP Format and Schemas > ISO 10303 Part 21 Edition 3"

# check for part 21 edition 3 content
        } elseif {[string first "ANCHOR\;" $line] == 0 || \
                  [string first "REFERENCE\;" $line] == 0 || \
                  [string first "SIGNATURE\;" $line] == 0} {
          set write 0
          lappend sects [string range $line 0 end-1]
        }

# write new file w/o part 21 edition 3 content, change 4;1 to 2;1
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
