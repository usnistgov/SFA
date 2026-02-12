proc reportUnknownEntities {} {
  global cells col count entCount entName entRows gpmiEnts heading localName numUnknownEnts row rowmax
  global sheetLast spmiEnts unknownEntityID unknownEnts worksheets worksheet wsCount wsNames ws_name

  set debug 0
  set result [parseStepEntities $localName $unknownEnts]

# get IDs for unknown entities
  set sortUnknownEnts {}
  for {set i 0} {$i < [llength $result]} {incr i} {
    set r0 [lindex $result $i]
    if {[expr {$i%2}] == 0} {
      set ent [string tolower $r0]
      lappend sortUnknownEnts $ent
    }
    for {set j 0} {$j < [llength $r0]} {incr j} {
      set r1 [lindex $r0 $j]
      set idx [lindex $r1 0]
      if {[string is integer $idx]} {set unknownEntityID($idx) $ent}
    }
  }

# process results to spreadsheet
  for {set i 0} {$i < [llength $result]} {incr i} {
    if {[expr {$i%2}] == 0} {
      set r0 [lsort [lindex $result $i]]
    } else {
      set r0 [lsort -integer -index 0 [lindex $result $i]]
    }
    if {$debug} {
      if {[expr {$i%2}] == 0} {outputMsg " "}
      outputMsg "$i / [llength $r0] / $r0"
    }

    set rm [expr {min([llength $r0],$rowmax-3)}]
    for {set j 0} {$j < $rm} {incr j} {
      set r1 [lindex $r0 $j]
      if {$debug} {if {$j < 20} {outputMsg "$j / [llength $r1] [llength [lindex $r1 1]] / $r1" green}}

# entity name, start new worksheet
      if {[catch {
        if {[llength $r1] == 1} {
          set ent [string tolower [lindex $r1 0]]
          set count($ent) $numUnknownEnts($ent)
          set entCount($ent) $numUnknownEnts($ent)
          set entRows($ent) [expr {$numUnknownEnts($ent)+3}]
          set gpmiEnts($ent) 0
          set spmiEnts($ent) 0

          set wsCount [$worksheets Count]
          if {$wsCount < 1} {
            set worksheet($ent) [$worksheets Item [expr [incr wsCount]]]
          } else {
            set worksheet($ent) [$worksheets Add [::tcom::na] $sheetLast]
          }
          $worksheet($ent) Activate
          set sheetLast $worksheet($ent)
          set name $ent
          if {[string length $name] > 31} {
            set name [string range $name 0 30]
            for {set n 1} {$n < 10} {incr n} {
              if {[info exists entName($name)]} {set name "[string range $name 0 29]$n"}
            }
          }
          set wsNames($name) $ent
          set ws_name($ent) [$worksheet($ent) Name $name]
          set cells($ent)   [$worksheet($ent) Cells]
          set heading($ent) 1
          set row($ent) 3
          set col($ent) 1

          $cells($ent) Item 3 1 ID
          $cells($ent) VerticalAlignment [expr -4160]
          set entName($name) $ent

# entity ID and attributes
        } elseif {[llength $r1] == 2} {
          lappend rowList [lindex $r1 0]
          foreach item [lindex $r1 1] {
            if {[string first "_MEASURE" $item] == -1 && [string first "COMMON_DATUM_LIST" $item] == -1} {
              set attributes [join $item]
              if {$attributes == "*" || $attributes == "\$"} {set attributes ""}

# group IDs for list of attributes
              catch {unset entIDs}
              if {[string first "\{" $attributes] == 0} {
                set missingAttributes ""
                foreach attr $attributes {
                  if {[llength $attr] == 2} {
                    lappend entIDs([lindex $attr 0]) [lindex $attr 1]
                  } elseif {[llength $attr] == 1} {
                    set attr1 [string range $attr 1 end]
                    if {[info exists unknownEntityID($attr1)]} {
                      lappend entIDs($unknownEntityID($attr1)) $attr1
                    } elseif {[string index $attr 0] == "\#"} {
                      append missingAttributes "(1) $attr  "
                    }
                  } else {
                    errorMsg "  Unexpected attribute on $ent: $attr" red
                  }
                }
                set attributes ""
                foreach idx [lsort [array names entIDs]]  {append attributes "([llength $entIDs($idx)]) $idx $entIDs($idx)  "}
                if {$missingAttributes != ""} {append attributes $missingAttributes}
                set attributes [string trim $attributes]
              }

# substitute unknown entity name
              set c1 [string first "\#" $attributes]
              if {$c1 == 0} {
                set c2 [string last  "\#" $attributes]
                if {$c2 == 0} {
                  set idx [string range $attributes 1 end]
                  if {[info exists unknownEntityID($idx)]} {set attributes "$unknownEntityID($idx) $idx"}
                }
              }

# add attributes to row list
              lappend rowList $attributes
              if {$row($ent) == 3} {
                incr col($ent)
                $cells($ent) Item 3 $col($ent) "attr[expr {$col($ent)-1}]"
              }
            } else {
              errorMsg "  Skipping '$item' text on [string toupper $ent]" red
            }
          }

# add rows to matrix
          incr row($ent)
          lappend matrixList $rowList
          unset rowList
        }
      } emsg]} {
        errorMsg "Error processing unknown entity: $emsg"
      }
    }

# write all rows (matrixList) at once
    if {[info exists matrixList]} {
      set str $numUnknownEnts($ent)
      if {$numUnknownEnts($ent) > $rm} {set str "$rm of $numUnknownEnts($ent)"}
      outputMsg " $ent ($str)"

      if {[catch {
        set range [$worksheet($ent) Range [cellRange 4 1] [cellRange [expr {[llength $matrixList]+3}] [llength [lindex $matrixList 0]]]]
        $range Value2 $matrixList
      } emsg]} {
        errorMsg "  Error writing worksheet for '$ent': $emsg" red
      }
      if {[llength $matrixList] != $numUnknownEnts($ent)} {set entRows($ent) [expr {[llength $matrixList]+3}]}
      unset matrixList
    }
  }

# move unknown entity worksheets to beginning
  set wsCount [$worksheets Count]
  set sortUnknownEnts [lsort -decreasing $sortUnknownEnts]
  foreach ent $sortUnknownEnts {
    foreach idx [array names entName] {
      if {$entName($idx) == $ent} {
        for {set i $wsCount} {$i > 0} {incr i -1} {
          if {$idx == [[$worksheets Item [expr $i]] Name]} {
            [$worksheets Item [expr $i]] -namedarg Move Before [$worksheets Item [expr 2]]
            break
          }
        }
      }
    }
  }
  set sheetLast [$worksheets Item [$worksheets Count]]
}

# -------------------------------------------------------------------------------
# code to parse STEP entities - based on ChatGPT https://chatgpt.com/share/694a0c30-8944-8005-ae26-eb0eacff81c7
proc parseStepEntities {filename typeList} {
  global ncomplex

  set fh [open $filename r]
  set data [read $fh]
  close $fh

  set result {}
  set buffer ""

  foreach line [split $data "\n"] {
    set line [string trim $line]
    if {$line eq ""} continue

    append buffer $line
    if {[string match *\; $line]} {
      set entity [string trim $buffer]
      set buffer ""

      if {[regexp {^#([0-9]+)\s*=\s*([A-Z0-9_]+)\s*\((.*)\)\s*;} $entity -> id type params]} {
        if {[lsearch -exact $typeList $type] >= 0} {
          dict lappend result $type [list $id [parseStepParams $params]]
        }
      } else {

# check for complex entities
        if {![info exists ncomplex]} {
          foreach ent $typeList {
            if {[string first $ent $entity] != -1 && [string first "\#" [string trim $entity]] == 0} {
              errorMsg " Unknown complex entities are not yet supported, for example:" red
              outputMsg "  $entity"
              set ncomplex 1
            }
          }
        }
      }
    }
  }
  catch {unset ncomplex}
  return $result
}

# -------------------------------------------------------------------------------
proc parseStepParams {paramString} {
  set tokens [stepTokenize $paramString]
  set idx 0
  return [stepParseTokens $tokens idx]
}

# -------------------------------------------------------------------------------
proc stepTokenize {s} {
  set tokens {}
  set token ""
  set inString 0
  set len [string length $s]

  for {set i 0} {$i < $len} {incr i} {
    set c [string index $s $i]

    if {$inString} {
      append token $c
      if {$c eq "'"} {
        set inString 0
        lappend tokens $token
        set token ""
      }
      continue
    }

    switch -- $c {
      "'" {
        set inString 1
        set token "'"
      }
      "(" - ")" - "," {
        if {$token ne ""} {
          lappend tokens [string trim $token]
          set token ""
        }
        lappend tokens $c
      }
      default {
        append token $c
      }
    }
  }

  if {$token ne ""} {lappend tokens [string trim $token]}
  return $tokens
}

# -------------------------------------------------------------------------------
proc stepParseTokens {tokens idxVar} {
  global objDesign
  upvar $idxVar idx

  set result {}
  while {$idx < [llength $tokens]} {
    set tok [lindex $tokens $idx]
    incr idx

    switch -- $tok {
      "(" {
        # Start nested list
        lappend result [stepParseTokens $tokens idx]
      }
      ")" {
        # End current list
        return $result
      }
      "," {
        # Comma only separates elements
        continue
      }
      default {
        # Atom
        if {$tok eq "$" || $tok eq "*"} {
          lappend result $tok
        } elseif {[regexp {^#([0-9]+)$} $tok -> ref]} {

# add known entity name for an ID (ref)
          set objValue [$objDesign FindObjectByP21Id [expr {int($ref)}]]
          if {$objValue != ""} {
            set ref "[formatComplexEnt [$objValue Type]] $ref"
          } else {
            set ref "\#$ref"
          }
          lappend result [list $ref]
        } elseif {[regexp {^'.*'$} $tok]} {
          lappend result [string range $tok 1 end-1]
        } elseif {[string is double -strict $tok]} {
          lappend result [expr {double($tok)}]
        } else {
          lappend result $tok
        }
      }
    }
  }
  return $result
}
