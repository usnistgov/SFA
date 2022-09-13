# read entity and write to spreadsheet
proc getEntity {objEntity rmax checkInverse checkBadAttributes unicodeCheck} {
  global attrType badAttributes cells col count entComment entCount entName entRows heading invMsg invVals matrixList opt roseLogical row
  global sheetLast skipEntities skipFileName skipPerm syntaxErr thisEntType unicodeAttributes unicodeString worksheet worksheets wsCount wsNames

# get entity type
  set thisEntType [$objEntity Type]

  if {[info exists invVals]} {unset invVals}
  set cellLimit1 500
  set cellLimit2 3000

# -------------------------------------------------------------------------------------------------
# open worksheet for each entity if it does not already exist
  if {![info exists worksheet($thisEntType)]} {
    set msg "[formatComplexEnt $thisEntType] ("
    set rm [expr {$rmax-3}]
    if {$entCount($thisEntType) > $rm} {append msg "$rm of "}
    append msg "$entCount($thisEntType))"
    outputMsg $msg

    if {$entCount($thisEntType) > $rm} {errorMsg " Maximum Rows ($rm) exceeded (see Spreadsheet tab)" red}
    if {$entCount($thisEntType) > 10000 && $rm > 10000} {errorMsg " Number of entities > 10000.  Consider using the Maximum Rows option." red}

    set wsCount [$worksheets Count]
    if {$wsCount < 1} {
      set worksheet($thisEntType) [$worksheets Item [expr [incr wsCount]]]
    } else {
      set worksheet($thisEntType) [$worksheets Add [::tcom::na] $sheetLast]
    }
    $worksheet($thisEntType) Activate

    set sheetLast $worksheet($thisEntType)

    set name $thisEntType
    if {[string length $name] > 31} {
      set name [string range $name 0 30]
      for {set i 1} {$i < 10} {incr i} {
        if {[info exists entName($name)]} {set name "[string range $name 0 29]$i"}
      }
    }
    set wsNames($name) $thisEntType
    set ws_name($thisEntType) [$worksheet($thisEntType) Name $name]
    set cells($thisEntType)   [$worksheet($thisEntType) Cells]
    set heading($thisEntType) 1

    set row($thisEntType) 4
    $cells($thisEntType) Item 3 1 "ID"

# set vertical alignment
    $cells($thisEntType) VerticalAlignment [expr -4160]

    set entName($name) $thisEntType
    set count($thisEntType) 0
    set invMsg ""

# color tab, not available in very old versions of Excel
    catch {
      set cidx [setColorIndex $thisEntType]
      if {$cidx > 0} {[$worksheet($thisEntType) Tab] ColorIndex [expr $cidx]}
    }

    set wsCount [$worksheets Count]
    set sheetLast $worksheet($thisEntType)

# file of entities not to process
    if {[catch {
      set skipFile [open $skipFileName w]
      foreach item $skipEntities {if {[lsearch $skipPerm $item] == -1} {puts $skipFile $item}}
      if {[lsearch $skipEntities $thisEntType] == -1 && [lsearch $skipPerm $thisEntType] == -1} {puts $skipFile $thisEntType}
      close $skipFile
    } emsg]} {
      errorMsg "Error processing 'skip' file: $emsg"
    }
    update idletasks

# -------------------------------------------------------------------------------------------------
# entity worksheet already open
  } else {
    incr row($thisEntType)
    set heading($thisEntType) 0
  }

# -------------------------------------------------------------------------------------------------
# if less than max allowed rows, append attribute values to rowList, append rowList to matrixList
  if {$row($thisEntType) <= $rmax} {
    set entRows($thisEntType) $row($thisEntType)
    set col($thisEntType) 1
    incr count($thisEntType)

# entity ID
    set p21id [$objEntity P21ID]
    lappend rowList $p21id
    [$worksheet($thisEntType) Range A$row($thisEntType)] NumberFormat "0"

# keep track of the entity ID for a row
    setIDRow $thisEntType $p21id

# -------------------------------------------------------------------------------------------------
# find inverse relationships for specific entities
    if {$checkInverse} {invFind $objEntity}
    set invLen 0
    if {[info exists invVals]} {set invLen [array size invVals]}

# -------------------------------------------------------------------------------------------------
# for all attributes of the entity
    set nattr 0
    set objAttributes [$objEntity Attributes]

    ::tcom::foreach objAttribute $objAttributes {
      set attrName [$objAttribute Name]

      if {[catch {
        if {!$checkBadAttributes} {
          set objValue [$objAttribute Value]

# substitute correct unicode
          if {$unicodeCheck} {
            foreach attr $unicodeAttributes($thisEntType) {
              if {$attr == $attrName} {
                set idx "$thisEntType,$attrName,$p21id"
                if {[info exists unicodeString($idx)]} {set objValue $unicodeString($idx)}
              }
            }
          }

# look for bad attributes that cause a crash
        } else {
          set ok 1
          if {[lsearch $badAttributes($thisEntType) $attrName] != -1} {set ok 0}
          if {$ok} {
            set objValue [$objAttribute Value]
          } else {
            set objValue "???"
            if {[llength $badAttributes($thisEntType)] == 1} {
              errorMsg " Reporting [formatComplexEnt $thisEntType] '$attrName' attribute is not supported." red
              errorMsg " '???' will appear in spreadsheet for this attribute.  See User Guide section 5.4" red
            } else {
              set str $badAttributes($thisEntType)
              regsub -all " " $str "' '" str
              errorMsg " Reporting [formatComplexEnt $thisEntType] '$str' attribute is not supported." red
              errorMsg " '???' will appear in spreadsheet for these attributes.  See User Guide section 5.4" red
            }
          }
        }

# error getting attribute value
      } emsgv]} {
        set msg "Error processing '$attrName' attribute on '[formatComplexEnt [$objEntity Type]]': $emsgv"
        errorMsg $msg
        lappend syntaxErr([$objEntity Type]) [list -$row($thisEntType) $attrName $msg]
        set objValue ""
        catch {raise .}
      }

      incr nattr

# -------------------------------------------------------------------------------------------------
# headings in first row only for first instance of an entity
      if {$heading($thisEntType) != 0} {
        incr heading($thisEntType)
        $cells($thisEntType) Item 3 $heading($thisEntType) $attrName
        if {[info exists badAttributes($thisEntType)]} {
          if {[lsearch $badAttributes($thisEntType) $attrName] != -1} {addCellComment $thisEntType 3 $heading($thisEntType) "Reporting this attribute is not supported.  Check the STEP file for the actual values."}
        }
        set attrType($heading($thisEntType)) [$objAttribute Type]
        set entComment($attrName) 1
      }

# -------------------------------------------------------------------------------------------------
# values in rows
      incr col($thisEntType)

# not a handle, just a single value
      if {[string first "handle" $objValue] == -1} {
        set ov $objValue

# if value is a boolean, substitute string roseLogical
        if {([$objAttribute Type] == "RoseBoolean" || [$objAttribute Type] == "RoseLogical") && [info exists roseLogical($ov)]} {set ov $roseLogical($ov)}

# check if showing numbers without rounding
        catch {
          if {!$opt(xlNoRound)} {
            lappend rowList $ov
          } elseif {$attrType($col($thisEntType)) != "double" && $attrType($col($thisEntType)) != "measure_value"} {
            lappend rowList $ov
          } elseif {[string length $ov] < 12} {
            lappend rowList $ov

# no rounding, show as text '
          } else {
            lappend rowList "'$ov"
          }
        }

# -------------------------------------------------------------------------------------------------
# node type 18=ENTITY, 19=SELECT TYPE  (node type is 20 for SET or LIST is processed below)
      } elseif {[$objAttribute NodeType] == 18 || [$objAttribute NodeType] == 19} {
        set refEntity [$objAttribute Value]

# get refType, however, sometimes this is not a single reference, but rather a list
        if {[catch {
          set refType [$refEntity Type]
          set valnotlist 1
        } emsg2]} {

# process like a list
          catch {foreach idx [array names cellval] {unset cellval($idx)}}
          ::tcom::foreach val $refEntity {append cellval([$val Type]) "[$val P21ID] "}
          set str ""
          set size 0
          catch {set size [array size cellval]}

          if {$size > 0} {
            foreach idx [lsort [array names cellval]] {
              set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
              if {$ncell > 1 || $size > 1} {
                set ok 1
                if {$ncell > $cellLimit1 && ([string first "styled_item" $idx] != -1 || [string first "triangulated" $idx] != -1 || \
                    [string first "connecting_edge" $idx] != -1 || [string first "3d_element_representation" $idx] != -1 || \
                    $idx == "node" || $idx == "cartesian_point" || $idx == "advanced_face")} {
                  set ok 0
                } elseif {$ncell > $cellLimit2} {
                  set ok 0
                }
                if {$ok} {
                  append str "($ncell) [formatComplexEnt $idx 1] $cellval($idx)  "
                } else {
                  append str "($ncell) [formatComplexEnt $idx 1]  "
                }
              } else {
                append str "(1) [formatComplexEnt $idx 1] $cellval($idx)  "
              }
            }
          }

          lappend rowList [string trim $str]
          set valnotlist 0
        }

# value is not a list which is the most common
        if {$valnotlist} {
          set str "[formatComplexEnt $refType 1] [$refEntity P21ID]"

# for length measure (and other measures), add the actual measure value
          set cellComment 0
          if {[string first "measure_with_unit" $refType] != -1} {
            ::tcom::foreach refAttribute [$refEntity Attributes] {
              if {[$refAttribute Name] == "value_component"} {
                set str "[$refAttribute Value] ($str)"
                set cellComment 1
              }
            }
          }
          if {$cellComment && $entComment($attrName)} {
            addCellComment $thisEntType 3 $col($thisEntType) "The values of *_measure_with_unit are also shown."
            set entComment($attrName) 0
          }

# next_assembly_usage_occurrence relating/related names
          if {$thisEntType == "next_assembly_usage_occurrence" && $opt(BOM)} {
            set pname ""
            if {[string first "relat" [$objAttribute Name]] == 0} {

# look for name of product_definition
              if {[$refEntity Type] == "product_definition" || [string first "composite_assembly" [$refEntity Type]] == 0} {
                set pname [string trim [[[$refEntity Attributes] Item [expr 1]] Value]]
                set idx "$refType,id,[$refEntity P21ID]"
                if {[info exists unicodeString($idx)]} {set pname $unicodeString($idx)}

# look for name on product
                if {$pname == "" || $pname == "design" || $pname == "part definition" || $pname == "None" || $pname == "UNKNOWN" || $pname == "BLNMEYEN"} {
                  set pdf [[[$refEntity Attributes] Item [expr 3]] Value]
                  set pro [[[$pdf Attributes] Item [expr 3]] Value]
                  set pname [string trim [[[$pro Attributes] Item [expr 1]] Value]]
                  set idx "$refType,id,[$refEntity P21ID]"
                  if {[info exists unicodeString($idx)]} {set pname $unicodeString($idx)}
                }
              }
            }

# add name to string
            if {$pname != ""} {
              set str "$str  ($pname)"
              if {$entComment($attrName)} {
                if {[string first "relating" [$objAttribute Name]] == 0} {
                  outputMsg " Adding Assembly/Component names" blue
                  addCellComment $thisEntType 3 5 "Name of the 'relating' Assembly is also shown."
                  addCellComment $thisEntType 3 6 "Name of the 'related' Component, which can be a part or subassembly, is also shown."
                  set entComment($attrName) 0
                }
              }
            }
          }

# add to row list
          lappend rowList $str
        }

# -------------------------------------------------------------------------------------------------
# node type 20=AGGREGATE (ENTITIES), usually SET or LIST, try as a tcom list or regular list (SELECT type)
      } elseif {[$objAttribute NodeType] == 20} {
        catch {foreach idx [array names cellval] {unset cellval($idx)}}
        catch {unset cellparam}
        set valMeasure {}

# collect the reference id's (P21ID) for the Type of entity in the SET or LIST
        if {[catch {
          ::tcom::foreach val [$objAttribute Value] {
            set valType [$val Type]
            append cellval($valType) "[$val P21ID] "

# check for length or plane measures
            if {[string first "measure_with_unit" $valType] != -1} {
              if {[string first "length" $valType] != -1 || [string first "plane" $valType] != -1} {
                ::tcom::foreach refAttribute [$val Attributes] {
                  if {[$refAttribute Name] == "value_component"} {lappend valMeasure [$refAttribute Value]}
                }
              }
            }
          }
        } emsg]} {
          foreach val [$objAttribute Value] {
            if {[string first "handle" $val] != -1} {
              set valType [$val Type]
              append cellval($valType) "[$val P21ID] "
            } else {
              append cellparam "$val "
            }
          }
        }

# -------------------------------------------------------------------------------------------------
# format cell values for the SET or LIST
        set str ""
        set size 0
        catch {set size [array size cellval]}

        set strMeasure ""
        if {[llength $valMeasure] > 0 && [llength $valMeasure] < 5} {set strMeasure "[join $valMeasure]  "}

        if {[info exists cellparam]} {append str "$cellparam "}
        if {$size > 0} {
          foreach idx [lsort [array names cellval]] {
            set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
            if {$ncell > 1 || $size > 1} {
              set ok 1
              if {$ncell > $cellLimit1 && ([string first "styled_item" $idx] != -1 || [string first "triangulated" $idx] != -1 || \
                  [string first "connecting_edge" $idx] != -1 || [string first "3d_element_representation" $idx] != -1 || \
                  $idx == "node" || $idx == "cartesian_point" || $idx == "advanced_face")} {
                set ok 0
              } elseif {$ncell > $cellLimit2} {
                set ok 0
              }
              if {$ok} {
                if {[string first "measure_with_unit" $idx] != -1} {
                  append str "$strMeasure\($ncell) [formatComplexEnt $idx 1] $cellval($idx)  "
                } else {
                  append str "($ncell) [formatComplexEnt $idx 1] $cellval($idx)  "
                }
              } else {
                append str "($ncell) [formatComplexEnt $idx 1]  "
              }
            } elseif {[string first "measure_with_unit" $idx] != -1} {
              append str "$strMeasure\(1) [formatComplexEnt $idx 1] $cellval($idx)  "
            } else {
              append str "(1) [formatComplexEnt $idx 1] $cellval($idx)  "
            }
          }
        }

        lappend rowList [string trim $str]
        if {$strMeasure != "" && $entComment($attrName)} {
          addCellComment $thisEntType 3 $col($thisEntType) "Values of *_measure_with_unit are also shown."
          set entComment($attrName) 0
        }
      }
    }

# append rowList to matrixList which will be written to spreadsheet after all entities have been processed in genExcel
    lappend matrixList $rowList

# -------------------------------------------------------------------------------------------------
# report inverses
    if {$invLen > 0} {invReport}

# rows exceeded, return of 0 will break the loop to process an entity type
  } else {
    return 0
  }

# clean up variables to hopefully release some memory
  foreach var {objAttributes attrName refEntity refType} {if {[info exists $var]} {unset $var}}
  update idletasks
  return 1
}

# -------------------------------------------------------------------------------
# keep track of the entity ID for a row
proc setIDRow {entType p21id} {
  global gpmiEnts gpmiIDRow idRow propDefIDRow row spmiEnts spmiIDRow

# row id for an entity id
  set idRow($entType,$p21id) $row($entType)

# specific arrays for properties and PMI
  if {$entType == "property_definition"} {
    set propDefIDRow($p21id) $row($entType)
  } elseif {$gpmiEnts($entType)} {
    set gpmiIDRow($entType,$p21id) $row($entType)
  } elseif {$spmiEnts($entType)} {
    set spmiIDRow($entType,$p21id) $row($entType)
  }
}

# -------------------------------------------------------------------------------------------------
# read entity and write to CSV file
proc getEntityCSV {objEntity checkBadAttributes} {
  global badAttributes count csvdirnam csvfile csvinhome csvstr entCount fcsv mydocs roseLogical row rowmax skipEntities skipFileName skipPerm thisEntType

# get entity type
  set thisEntType [$objEntity Type]
  set cellLimit1 500
  set cellLimit2 3000

# -------------------------------------------------------------------------------------------------
# csv file for each entity if it does not already exist
  if {![info exists csvfile($thisEntType)]} {
    set countMsg "[formatComplexEnt $thisEntType] ("
    set rm [expr {$rowmax-3}]
    if {$entCount($thisEntType) > $rm} {append countMsg "$rm of "}
    append countMsg "$entCount($thisEntType))"
    outputMsg $countMsg

    if {$entCount($thisEntType) > $rm} {errorMsg " Maximum Rows ($rm) exceeded (see Spreadsheet tab)" red}
    if {$entCount($thisEntType) > 10000 && $rm > 10000} {errorMsg " Number of entities > 10000.  Consider using the Maximum Rows option." red}

# open csv file
    set csvfile($thisEntType) 1
    set csvfname [file join $csvdirnam $thisEntType.csv]
    if {[string length $csvfname] > 218} {
      set csvfname [file nativename [file join $mydocs $thisEntType.csv]]
      errorMsg " Some CSV files are saved in the home directory." red
      set csvinhome 1
    }
    set fcsv [open $csvfname w]
    puts $fcsv $countMsg

# headings in first row
    set csvstr "ID"
    ::tcom::foreach objAttribute [$objEntity Attributes] {append csvstr ",[$objAttribute Name]"}
    puts $fcsv $csvstr
    unset csvstr

    set count($thisEntType) 0
    set row($thisEntType) 4

# file of entities not to process
    if {[catch {
      set skipFile [open $skipFileName w]
      foreach item $skipEntities {if {[lsearch $skipPerm $item] == -1} {puts $skipFile $item}}
      if {[lsearch $skipEntities $thisEntType] == -1 && [lsearch $skipPerm $thisEntType] == -1} {puts $skipFile $thisEntType}
      close $skipFile
    } emsg]} {
      errorMsg "Error processing 'skip' file: $emsg"
    }
    update idletasks

# CSV file already open
  } else {
    incr row($thisEntType)
  }

# -------------------------------------------------------------------------------------------------
# start appending to csvstr, if less than max allowed rows
  update idletasks
  if {$row($thisEntType) <= $rowmax} {
    incr count($thisEntType)

# entity ID
    set p21id [$objEntity P21ID]

# -------------------------------------------------------------------------------------------------
# for all attributes of the entity
    set nattr 0
    set csvstr $p21id
    set objAttributes [$objEntity Attributes]
    ::tcom::foreach objAttribute $objAttributes {
      set attrName [$objAttribute Name]

      if {[catch {
        if {!$checkBadAttributes} {
          set objValue [$objAttribute Value]

# look for bad attributes that cause a crash
        } else {
          set ok 1
          foreach ba $badAttributes($thisEntType) {if {$ba == $attrName} {set ok 0}}
          if {$ok} {
            set objValue [$objAttribute Value]
          } else {
            set objValue "???"
            errorMsg " Reporting [formatComplexEnt $thisEntType] '$attrName' attribute is not supported." red
            errorMsg " '???' will appear in CSV file for this attribute.  See User Guide section 5.4" red
          }
        }

# error getting attribute value
      } emsgv]} {
        errorMsg "Error processing '$attrName' attribute on '[formatComplexEnt [$objEntity Type]]': $emsgv"
        set objValue ""
        catch {raise .}
      }
      incr nattr

# -------------------------------------------------------------------------------------------------
# not a handle, just a single value
      if {[string first "handle" $objValue] == -1} {
        set ov $objValue

# if value is a boolean, substitute string roseLogical
        if {([$objAttribute Type] == "RoseBoolean" || [$objAttribute Type] == "RoseLogical") && [info exists roseLogical($ov)]} {set ov $roseLogical($ov)}

# check for commas and double quotes
        if {[string first "," $ov]  != -1} {
          if {[string first "\"" $ov] != -1} {regsub -all "\"" $ov "\"\"" ov}
          set ov "\"$ov\""
        }

        append csvstr ",$ov"

# -------------------------------------------------------------------------------------------------
# node type 18=ENTITY, 19=SELECT TYPE  (node type is 20 for SET or LIST is processed below)
      } elseif {[$objAttribute NodeType] == 18 || [$objAttribute NodeType] == 19} {
        set refEntity [$objAttribute Value]

# get refType, however, sometimes this is not a single reference, but rather a list
#  which causes an error and it has to be processed like a list below
        if {[catch {
          set refType [$refEntity Type]
          set valnotlist 1
        } emsg2]} {

# process like a list which is very unusual
          catch {foreach idx [array names cellval] {unset cellval($idx)}}
          ::tcom::foreach val $refEntity {append cellval([$val Type]) "[$val P21ID] "}
          set str ""
          set size 0
          catch {set size [array size cellval]}

          if {$size > 0} {
            foreach idx [lsort [array names cellval]] {
              set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
              if {$ncell > 1 || $size > 1} {
                set ok 1
                if {$ncell > $cellLimit1 && ([string first "styled_item" $idx] != -1 || [string first "triangulated" $idx] != -1 || \
                    [string first "connecting_edge" $idx] != -1 || [string first "3d_element_representation" $idx] != -1 || \
                    $idx == "node" || $idx == "cartesian_point" || $idx == "advanced_face")} {
                  set ok 0
                } elseif {$ncell > $cellLimit2} {
                  set ok 0
                }
                if {$ok} {
                  append str "($ncell) [formatComplexEnt $idx 1] $cellval($idx)  "
                } else {
                  append str "($ncell) [formatComplexEnt $idx 1]  "
                }
              } else {
                append str "(1) [formatComplexEnt $idx 1] $cellval($idx)  "
              }
            }
          }
          append csvstr ",$str"
          set valnotlist 0
        }

# value is not a list which is the most common
        if {$valnotlist} {
          set str "[formatComplexEnt $refType 1] [$refEntity P21ID]"

# for length measure (and other measures), add the actual measure value
          if {[string first "measure_with_unit" $refType] != -1} {
            ::tcom::foreach refAttribute [$refEntity Attributes] {
              if {[$refAttribute Name] == "value_component"} {set str "[$refAttribute Value] ($str)"}
            }
          }
          append csvstr ",$str"
        }

# -------------------------------------------------------------------------------------------------
# node type 20=AGGREGATE (ENTITIES), usually SET or LIST, try as a tcom list or regular list (SELECT type)
      } elseif {[$objAttribute NodeType] == 20} {
        catch {foreach idx [array names cellval] {unset cellval($idx)}}
        catch {unset cellparam}

# collect the reference id's (P21ID) for the Type of entity in the SET or LIST
        if {[catch {
          ::tcom::foreach val [$objAttribute Value] {
            append cellval([$val Type]) "[$val P21ID] "
          }
        } emsg]} {
          foreach val [$objAttribute Value] {
            if {[string first "handle" $val] != -1} {
              append cellval([$val Type]) "[$val P21ID] "
            } else {
              append cellparam "$val "
            }
          }
        }

# -------------------------------------------------------------------------------------------------
# format cell values for the SET or LIST
        set str ""
        set size 0
        catch {set size [array size cellval]}

        if {[info exists cellparam]} {append str "$cellparam "}
        if {$size > 0} {
          foreach idx [lsort [array names cellval]] {
            set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
            if {$ncell > 1 || $size > 1} {
              set ok 1
              if {$ncell > $cellLimit1 && ([string first "styled_item" $idx] != -1 || [string first "triangulated" $idx] != -1 || \
                                   [string first "connecting_edge" $idx] != -1 || [string first "3d_element_representation" $idx] != -1 || \
                                   $idx == "node" || $idx == "cartesian_point" || $idx == "advanced_face")} {
                set ok 0
              } elseif {$ncell > $cellLimit2} {
                set ok 0
              }
              if {$ok} {
                append str "($ncell) [formatComplexEnt $idx 1] $cellval($idx)  "
              } else {
                append str "($ncell) [formatComplexEnt $idx 1]  "
              }
            } else {
              append str "(1) [formatComplexEnt $idx 1] $cellval($idx)  "
            }
          }
        }
        append csvstr ",[string trim $str]"
      }
    }

# write to CSV file
    if {[catch {
      puts $fcsv $csvstr
    } emsg]} {
      errorMsg "Error writing to CSV file for: $thisEntType"
    }

# rows exceeded, return of 0 will break the loop to process an entity type
  } else {
    return 0
  }

# -------------------------------------------------------------------------------------------------
# clean up variables to hopefully release some memory
  foreach var {objAttributes attrName refEntity refType} {if {[info exists $var]} {unset $var}}
  update idletasks
  return 1
}
