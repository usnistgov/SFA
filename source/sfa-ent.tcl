# read entity and write to spreadsheet

proc getEntity {objEntity checkInverse} {
  global attrType badAttributes cells col count developer entCount entName excelVersion
  global skipEntities skipPerm heading invMsg invVals localName opt roseLogical row rowmax sheetLast
  global thisEntType worksheet worksheets wsCount
  
# get entity type
  set thisEntType [$objEntity Type]
  #if {$developer} {if {$thisEntType != $expectedEnt} {errorMsg "Mismatch: $thisEntType  $expectedEnt"}}

  if {[info exists invVals]} {unset invVals}

# -------------------------------------------------------------------------------------------------
# open worksheet for each entity if it does not already exist
  if {![info exists worksheet($thisEntType)]} {
    set msg "[formatComplexEnt $thisEntType] ($entCount($thisEntType))"
    outputMsg $msg
    
    if {$entCount($thisEntType) > 50000 && $rowmax > 50010} {errorMsg " Number of entities > 50000.  Consider using the Maximum Rows option." red}

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
      errorMsg " Worksheet names are truncated to the first 31 characters" red
    }
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

# color tab
    if {$excelVersion >= 12} {
      set cidx [setColorIndex $thisEntType]
      if {$cidx > 0} {[$worksheet($thisEntType) Tab] ColorIndex [expr $cidx]}      
    }

    set wsCount [$worksheets Count]
    set sheetLast $worksheet($thisEntType)

# file of entities not to process
    set cfile [file rootname $localName]
    append cfile "-skip.dat"
    if {[catch {
      set skipFile [open $cfile w]
      foreach item $skipEntities {if {[lsearch $skipPerm $item] == -1} {puts $skipFile $item}}
      if {[lsearch $skipEntities $thisEntType] == -1 && [lsearch $skipPerm $thisEntType] == -1} {puts $skipFile $thisEntType}
      close $skipFile
    } emsg]} {
      errorMsg "ERROR processing 'skip' file: $emsg"
    }
    update idletasks

# -------------------------------------------------------------------------------------------------
# entity worksheet already open
  } else {
    incr row($thisEntType)
    if {$row($thisEntType) > $rowmax} {outputMsg " Maximum Rows exceeded ([expr {$rowmax-3}])" red}
    set heading($thisEntType) 0
  }

# -------------------------------------------------------------------------------------------------
# start filling in the cells

# if less than max allowed rows
  if {$row($thisEntType) <= $rowmax} {
    set col($thisEntType) 1
    incr count($thisEntType)
    
# show progress with > 50000 entities
    if {$entCount($thisEntType) >= 50000} {
      set c1 [expr {$count($thisEntType)%20000}]
      if {$c1 == 0} {
        outputMsg " $count($thisEntType) of $entCount($thisEntType) processed"
        update idletasks
      }
    }

# entity ID
    set p21id [$objEntity P21ID]
    $cells($thisEntType) Item $row($thisEntType) 1 $p21id
    [$worksheet($thisEntType) Range A$row($thisEntType)] NumberFormat "0"
      
# keep track of property_defintion or annotation occurrence rows in propDefIDRow, gpmiIDRow
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
      #outputMsg "$p21id $attrName [$objAttribute NodeType]" red

      if {[catch {
        if {![info exists badAttributes($thisEntType)]} {
          set objValue [$objAttribute Value]

# look for bad attributes that cause a crash
        } else {
          set ok 1
          foreach ba $badAttributes($thisEntType) {if {$ba == $attrName} {set ok 0}}
          if {$ok} {
            set objValue [$objAttribute Value]
          } else {
            set objValue "???"
            if {[llength $badAttributes($thisEntType)] == 1} {
              errorMsg " Skipping attribute '$attrName' on [formatComplexEnt $thisEntType] - '???' will appear in spreadsheet for this attribute" red
            } else {
              set str $badAttributes($thisEntType)
              regsub -all " " $str "' '" str
              errorMsg " Skipping attributes '$str' on [formatComplexEnt $thisEntType]\n '???' will appear in spreadsheet for these attributes" red
            }
          }
        }

# error getting attribute value
      } emsgv]} {
        set msg "ERROR processing #[$objEntity P21ID]=[$objEntity Type] '$attrName' attribute: $emsgv"
        if {[string first "datum_reference_compartment 'modifiers' attribute" $msg] != -1 || \
            [string first "datum_reference_element 'modifiers' attribute" $msg] != -1 || \
            [string first "annotation_plane 'elements' attribute" $msg] != -1} {
          errorMsg "Syntax Error: On '[$objEntity Type]' entities change the '$attrName' attribute with\n '()' to '$' where applicable.  The attribute is an OPTIONAL SET\[1:?\] and '()' is not valid."
        }
        errorMsg $msg
        set objValue ""
        catch {raise .}
      }

      incr nattr

# -------------------------------------------------------------------------------------------------
# headings in first row only for first instance of an entity
      if {$heading($thisEntType) != 0} {
        set ihead 1
        if {$ihead} {
          $cells($thisEntType) Item 3 [incr heading($thisEntType)] $attrName
          set attrType($heading($thisEntType)) [$objAttribute Type]
          if {[$objAttribute Type] == "STR" || [$objAttribute Type] == "RoseBoolean" || [$objAttribute Type] == "RoseLogical"} {
            #outputMsg "  $attrName  [$objAttribute Type]"  
            set letters ABCDEFGHIJKLMNOPQRSTUVWXYZ
            set c $heading($thisEntType)
            set inc [expr {int(double($c-1.)/26.)}]
            if {$inc == 0} {
              set c [string index $letters [expr {$c-1}]]
            } else {
              set c [string index $letters [expr {$inc-1}]][string index $letters [expr {$c-$inc*26-1}]]
            }
            #set range [$worksheet($thisEntType) Range "$c:$c"]
            #[$range Columns] NumberFormat "@"
          } 
        }
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
          if {!$opt(XL_FPREC)} {
            $cells($thisEntType) Item $row($thisEntType) $col($thisEntType) $ov
          } elseif {$attrType($col($thisEntType)) != "double" && $attrType($col($thisEntType)) != "measure_value"} {
            $cells($thisEntType) Item $row($thisEntType) $col($thisEntType) $ov
          } elseif {[string length $ov] < 12} {
            $cells($thisEntType) Item $row($thisEntType) $col($thisEntType) $ov

# no rounding, show as text '
          } else {
            $cells($thisEntType) Item $row($thisEntType) $col($thisEntType) "'$ov"
          }
        }

# -------------------------------------------------------------------------------------------------
# if attribute is reference to another entity
      } else {
        
# node type 18=ENTITY, 19=SELECT TYPE  (node type is 20 for SET or LIST is processed below)
        if {[$objAttribute NodeType] == 18 || [$objAttribute NodeType] == 19} {
          set refEntity [$objAttribute Value]

# get refType, however, sometimes this is not a single reference, but rather a list
#  which causes an error and it has to be processed like a list below
          if {[catch {
            set refType [$refEntity Type]
            set valnotlist 1
          } emsg2]} {

# process like a list which is very unusual
            #if {$developer} {errorMsg " Attribute reference is a List: $emsg2"}
            catch {foreach idx [array names cellval] {unset cellval($idx)}}
            ::tcom::foreach val $refEntity {
              append cellval([$val Type]) "[$val P21ID] "
            }
            set str ""
            set size 0
            catch {set size [array size cellval]}

            if {$size > 0} {
              foreach idx [lsort [array names cellval]] {
                set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
                if {$ncell > 1 || $size > 1} {
                  if {$ncell < 30} {
                    append str "($ncell) [formatComplexEnt $idx 1] $cellval($idx)  "
                  } else {
                    append str "($ncell) [formatComplexEnt $idx 1]  "
                  }
                } else {
                  append str "(1) [formatComplexEnt $idx 1] $cellval($idx)  "
                }
              }
            }
            $cells($thisEntType) Item $row($thisEntType) $col($thisEntType) [string trim $str]
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

            $cells($thisEntType) Item $row($thisEntType) $col($thisEntType) $str
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
                if {$ncell < 30} {
                  append str "($ncell) [formatComplexEnt $idx 1] $cellval($idx)  "
                } else {
                  append str "($ncell) [formatComplexEnt $idx 1]  "
                }
              } else {
                append str "(1) [formatComplexEnt $idx 1] $cellval($idx)  "
              }
            }
          }
          $cells($thisEntType) Item $row($thisEntType) $col($thisEntType) [string trim $str]
        }
      }
    }

# -------------------------------------------------------------------------------------------------
# report inverses    
    if {$invLen > 0} {invReport}

# rows exceeded
  } else {
    return 0
  }  

# clean up variables to hopefully release some memory
  foreach var {objAttributes attrName refEntity refType} {
    if {[info exists $var]} {unset $var}
  }
  update idletasks
  return 1
}

# -------------------------------------------------------------------------------
# keep track of property_defintion, annotation occurrence, or semantic PMI rows in
# propDefIDRow, gpmiIDRow, spmiIDRow

proc setIDRow {entType p21id} {
  global gpmiEnts gpmiIDRow propDefIDRow row spmiEnts spmiIDRow
  #outputMsg "setIDRow [info exists spmiEnts($entType)] $entType $p21id $row($entType)" red
  
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
proc getEntityCSV {objEntity} {
  global badAttributes count csvdirnam csvfile csvstr entCount fcsv skipEntities skipPerm localName roseLogical thisEntType 
  
# get entity type
  set thisEntType [$objEntity Type]

# -------------------------------------------------------------------------------------------------
# csv file for each entity if it does not already exist
  if {![info exists csvfile($thisEntType)]} {
    set msg "[formatComplexEnt $thisEntType] ($entCount($thisEntType))"
    outputMsg $msg

# open csv file
    set csvfile($thisEntType) 1
    set csvfname [file join $csvdirnam $thisEntType.csv]
    set fcsv [open $csvfname w]
    puts $fcsv "[formatComplexEnt $thisEntType] ($entCount($thisEntType))"
    #outputMsg $fcsv red

# headings in first row
    set csvstr "ID"
    ::tcom::foreach objAttribute [$objEntity Attributes] {append csvstr ",[$objAttribute Name]"}
    puts $fcsv $csvstr
    unset csvstr

    set count($thisEntType) 0

# file of entities not to process
    set cfile [file rootname $localName]
    append cfile "-skip.dat"
    if {[catch {
      set skipFile [open $cfile w]
      foreach item $skipEntities {if {[lsearch $skipPerm $item] == -1} {puts $skipFile $item}}
      if {[lsearch $skipEntities $thisEntType] == -1 && [lsearch $skipPerm $thisEntType] == -1} {puts $skipFile $thisEntType}
      close $skipFile
    } emsg]} {
      errorMsg "ERROR processing 'skip' file: $emsg"
    }
    update idletasks
  }

# -------------------------------------------------------------------------------------------------
# start filling in the cells
  incr count($thisEntType)
  
# show progress with > 50000 entities
  if {$entCount($thisEntType) >= 50000} {
    set c1 [expr {$count($thisEntType)%20000}]
    if {$c1 == 0} {
      outputMsg " $count($thisEntType) of $entCount($thisEntType) processed"
      update idletasks
    }
  }

# entity ID
  set p21id [$objEntity P21ID]

# -------------------------------------------------------------------------------------------------
# for all attributes of the entity
  set nattr 0
  set csvstr $p21id
  set objAttributes [$objEntity Attributes]
  ::tcom::foreach objAttribute $objAttributes {
    set attrName [$objAttribute Name]
    #outputMsg "$p21id $attrName [$objAttribute NodeType]" red

    if {[catch {
      if {![info exists badAttributes($thisEntType)]} {
        set objValue [$objAttribute Value]

# look for bad attributes that cause a crash
      } else {
        set ok 1
        foreach ba $badAttributes($thisEntType) {if {$ba == $attrName} {set ok 0}}
        if {$ok} {
          set objValue [$objAttribute Value]
        } else {
          set objValue "???"
          errorMsg " Skipping '$attrName' attribute on [formatComplexEnt $thisEntType] - '???' will appear in the CSV file for this attribute" red
        }
      }

# error getting attribute value
    } emsgv]} {
      set msg "ERROR processing #[$objEntity P21ID]=[$objEntity Type] '$attrName' attribute: $emsgv"
      if {[string first "datum_reference_compartment 'modifiers' attribute" $msg] != -1 || \
          [string first "datum_reference_element 'modifiers' attribute" $msg] != -1 || \
          [string first "annotation_plane 'elements' attribute" $msg] != -1} {
        errorMsg "Syntax Error: On '[$objEntity Type]' entities change the '$attrName' attribute with\n '()' to '$' where applicable.  The attribute is an OPTIONAL SET\[1:?\] and '()' is not valid."
      }
      errorMsg $msg
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

      append csvstr ",$ov"

# -------------------------------------------------------------------------------------------------
# if attribute is reference to another entity
    } else {
      
# node type 18=ENTITY, 19=SELECT TYPE  (node type is 20 for SET or LIST is processed below)
      if {[$objAttribute NodeType] == 18 || [$objAttribute NodeType] == 19} {
        set refEntity [$objAttribute Value]

# get refType, however, sometimes this is not a single reference, but rather a list
#  which causes an error and it has to be processed like a list below
        if {[catch {
          set refType [$refEntity Type]
          set valnotlist 1
        } emsg2]} {

# process like a list which is very unusual
          catch {foreach idx [array names cellval] {unset cellval($idx)}}
          ::tcom::foreach val $refEntity {
            append cellval([$val Type]) "[$val P21ID] "
          }
          set str ""
          set size 0
          catch {set size [array size cellval]}

          if {$size > 0} {
            foreach idx [lsort [array names cellval]] {
              set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
              if {$ncell > 1 || $size > 1} {
                if {$ncell < 30} {
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
              if {$ncell < 30} {
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
  }
  #outputMsg "$fcsv $csvstr"

# write to CSV file
  if {[catch {
    puts $fcsv $csvstr
  } emsg]} {
    errorMsg "Error writing to CSV file for: $thisEntType"
  }

# -------------------------------------------------------------------------------------------------
# clean up variables to hopefully release some memory
  foreach var {objAttributes attrName refEntity refType} {if {[info exists $var]} {unset $var}}
  update idletasks
  return 1
}
