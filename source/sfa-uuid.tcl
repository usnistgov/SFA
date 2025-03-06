# add UUIDs from uuid_attribute entities (AP242 Edition >= 4)
proc uuidGetAttributes {totalUUID entsUUID} {
  global cells idRow localName syntaxErr uuid uuidEnts
  global objDesign

  errorMsg "\nProcessing UUID attributes" blue
  set nUUID 0
  set allUUID {}
  set noUUIDent {}
  catch {unset uuid}

# read STEP file for UUID entities because one of the attributes is a LIST of LIST
  set f [open $localName r]
  while {[gets $f line] >= 0} {
    set ok 0
    foreach ent $entsUUID {if {[string first $ent $line] != -1} {set ok 1; break}}
    if {$ok} {
      set ent [string tolower $ent]

# get rest of entity if one multiple lines
      while {1} {
        if {[string first ";" $line] == -1} {
          gets $f line1
          append line $line1
        } else {
          break
        }
      }

# entity ID
      if {[catch {
        set entid [string range $line 1 [string first "=" $line]-1]

# get UUID (pid)
        set pid [string range $line [string first "'" $line]+1 [string last "'" $line]-1]
        regsub -all {[0123456789abcdefABCDEF-]} $pid "" npid
        if {[string length $npid] > 0} {
          set msg "Unexpected characters ($npid) in UUID"
          errorMsg " $msg"
          lappend syntaxErr($ent) [list $entid identifier $msg]
        }
        if {[string length $pid] == 36 && [string first "-" $pid] == 8 && [string last "-" $pid] == 23} {

# check for duplicate UUIDs
          if {[lsearch $allUUID $pid] == -1} {
            lappend allUUID $pid
          } else {
            set msg "UUID is assigned to multiple identified_item.  Only one UUID entity is needed that lists all entities in identified_item."
            lappend syntaxErr($ent) [list $entid identifier $msg]
            errorMsg " $msg" red
          }
          set uuidstr $pid
          if {[string index $ent 0] == "v"} {append uuidstr " ([string range $ent 0 1])"}
          if {[string first "HASH" $line] != -1} {append uuidstr " (hash v5)"}
          if {[string first "LOCATION" $line] != -1} {append uuidstr " (location)"}

# get identified_items
          set line1 [string range $line [string last "'" $line]+3 end]
          set c1 1
          if {[string index $line1 0] == "\#"} {set c1 0}
          set items [split [string range $line1 $c1 [string first "))" $line1]-1] ","]
          set iditem ""

# loop over all items
          foreach item $items {
            set eid [string range $item [string first "\#" $item]+1 end]
            set c2 [string first ")" $eid]
            if {$c2 != -1} {set eid [string range $eid 0 $c2-1]}
            set c2 [string first "(" $eid]
            if {$c2 != -1} {set eid [string range $eid $c2+1 end]}

            set e1 [$objDesign FindObjectByP21Id [expr $eid]]
            set uuidEnt [$e1 Type]
            if {[lsearch $uuidEnts $uuidEnt] == -1} {lappend uuidEnts $uuidEnt}
            if {$uuidEnt == "id_attribute"} {
              set msg "Error: identified_item should refer directly to entities assigned a UUID and not id_attribute"
              errorMsg " $msg"
              lappend syntaxErr($ent) [list $entid identified_item $msg]
            }

            if {[info exists uuid($uuidEnt,[$e1 P21ID])]} {
              set msg "Error: Multiple UUIDs are associated with the same entity"
              errorMsg " $msg"
              lappend syntaxErr($ent) [list $entid identified_item $msg]
            }
            set uuid($uuidEnt,[$e1 P21ID]) $uuidstr
            set okid 1
            if {![info exist cells($uuidEnt)]} {lappend noUUIDent $uuidEnt}
            if {$iditem != ""} {
              errorMsg " Some UUIDs are associated with multiple entities" red
              if {[info exists idRow($ent,$entid)]} {
                addCellComment $ent $idRow($ent,$entid) 3 "UUID is associated with multiple entities"
              }
            }
            append iditem "[formatComplexEnt $uuidEnt] [$e1 P21ID]   "
          }

# write identified_items to uuid_attribute entity
          if {[info exists idRow($ent,$entid)]} {
            $cells($ent) Item $idRow($ent,$entid) 3 $iditem
          }

        } else {
          set msg "Error with UUID format"
          errorMsg " $msg"
          lappend syntaxErr($ent) [list $entid identifier $msg]
        }
      } emsg2]} {
        errorMsg "Error getting UUID: $emsg2"
      }

      incr nUUID
      if {$nUUID == $totalUUID} {break}
    }
  }
  close $f

  set noUUIDent [lrmdups $noUUIDent]
  if {[llength $noUUIDent] > 0} {
    regsub -all " " [join [lrmdups $noUUIDent]] ", " str
    outputMsg " UUIDs are also associated with: $str" red
    unset noUUIDent
  }
}

# -------------------------------------------------------------------------------------------------
# add UUIDs with uuid_attribute (idType=2) OR add worksheets for Part 21 edition 3 sections (idType=1)
proc uuidReportAttributes {idType {uuidEnt ""}} {
  global cells entCount entName fileSumRow idRow legendColor nistName p21e3Section spmiSumRowID sumHeaderRow uuid worksheet worksheets
  global objDesign

# add UUIDs with uuid_attribute
  if {$idType == 2} {
    set heading "UUID"
    foreach idx [array names uuid] {
      set uuidval $uuid($idx)
      set idx [split $idx ","]
      set anchorEnt [lindex $idx 0]
      if {$uuidEnt == "" || $anchorEnt == $uuidEnt} {
        set anchorID  [lindex $idx 1]
        if {[info exists worksheet($anchorEnt)]} {
          if {![info exists urow($anchorEnt)]} {set urow($anchorEnt) [[[$worksheet($anchorEnt) UsedRange] Rows] Count]}
          if {![info exists ucol($anchorEnt)]} {set ucol($anchorEnt) [getNextUnusedColumn $anchorEnt]}
          if {[info exists idRow($anchorEnt,$anchorID)]} {
            set ur $idRow($anchorEnt,$anchorID)
            $cells($anchorEnt) Item $ur $ucol($anchorEnt) $uuidval
            set range [$worksheet($anchorEnt) Range [cellRange $ur $ucol($anchorEnt)]]
            [$range Interior] ColorIndex [expr 40]
            catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 1]}}
          }
          if {[info exists spmiSumRowID($anchorID)]} {set anchorSum($spmiSumRowID($anchorID)) $uuidval}
        }
      }
    }

# -------------------------------------------------------------------------------------------------
# look for three section types possible in Part 21 Edition 3
  } elseif {$idType == 1} {
    set heading "ANCHOR ID"
    foreach line $p21e3Section {
      if {$line == "ANCHOR" || $line == "REFERENCE" || $line == "SIGNATURE"} {
        set sect $line
        set worksheet($sect) [$worksheets Add [::tcom::na] [$worksheets Item [$worksheets Count]]]
        set n [$worksheets Count]
        [$worksheets Item [expr $n]] -namedarg Move Before [$worksheets Item [expr 3]]
        $worksheet($sect) Activate
        $worksheet($sect) Name $sect
        set hlink [$worksheet($sect) Hyperlinks]
        set cells($sect) [$worksheet($sect) Cells]
        set r 0
        outputMsg " Adding $line worksheet" blue
      }

# add to worksheet
      incr r
      set line1 $line
      if {$r == 1} {addCellComment $sect 1 1 "See Help > User Guide section 5.6."}
      $cells($sect) Item $r 1 $line1

# process anchor section persistent IDs
      if {$sect == "ANCHOR"} {
        if {$r == 1} {$cells($sect) Item $r 2 "Entity"}
        set c2 [string first ";" $line]
        if {$c2 != -1} {set line [string range $line 0 $c2-1]}

        set c1 [string first "\#" $line]
        if {$c1 != -1} {
          set badEnt 0
          set anchorID [string range $line $c1+1 end]
          if {[string is integer $anchorID]} {
            if {[catch {
              set objValue  [$objDesign FindObjectByP21Id [expr {int($anchorID)}]]
              set anchorEnt [$objValue Type]

# add anchor ID to entity worksheet and PMI summary
              if {$anchorEnt != ""} {
                $cells($sect) Item $r 2 $anchorEnt

                if {[info exist fileSumRow($anchorEnt)]} {
                  set fsrow [expr {$fileSumRow($anchorEnt)+$sumHeaderRow+1}]
                  set val [[$cells(Summary) Item $fsrow 1] Value]
                  if {[string first "Anchor" $val] == -1} {
                    $cells(Summary) Item $fsrow 1 "$val  \[Anchor\]"
                    set range [$worksheet(Summary) Range [cellRange $fsrow 1]]
                    [$range Font] Bold [expr 1]
                  }
                }

# add anchor ID to entity worksheet
                if {[info exists worksheet($anchorEnt)]} {
                  set c3 [string first ">" $line]
                  if {$c3 == -1} {set c3 [string first "=" $line]}
                  set uuidval [string range $line 1 $c3-1]
                  if {![info exists urow($anchorEnt)]} {set urow($anchorEnt) [[[$worksheet($anchorEnt) UsedRange] Rows] Count]}
                  if {![info exists ucol($anchorEnt)]} {set ucol($anchorEnt) [getNextUnusedColumn $anchorEnt]}
                  if {[info exists idRow($anchorEnt,$anchorID)]} {
                    set ur $idRow($anchorEnt,$anchorID)
                    set val [[$cells($anchorEnt) Item $ur $ucol($anchorEnt)] Value]
                    if {$val == ""} {
                      $cells($anchorEnt) Item $ur $ucol($anchorEnt) $uuidval
                    } else {
                      $cells($anchorEnt) Item $ur $ucol($anchorEnt) "$val   $uuidval"
                    }
                    set range [$worksheet($anchorEnt) Range [cellRange $ur $ucol($anchorEnt)]]
                    [$range Interior] ColorIndex [expr 40]
                    catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 1]}}
                  }

# link to entity worksheet
                  set anchor [$worksheet($sect) Range "B$r"]
                  set hlsheet $anchorEnt
                  if {[string length $anchorEnt] > 31} {
                    foreach item [array names entName] {if {$entName($item) == $anchorEnt} {set hlsheet $item}}
                  }
                  catch {$hlink Add $anchor [string trim ""] "$hlsheet!A1" "Go to $anchorEnt"}

# add anchor ID to PMI summary
                  if {[info exists spmiSumRowID($anchorID)]} {
                    set anchorSum($spmiSumRowID($anchorID)) $uuidval
                  } elseif {[string first "dimensional_size" $anchorEnt] != -1 || [string first "dimensional_location" $anchorEnt] != -1} {
                    set dcrs [$objValue GetUsedIn [string trim dimensional_characteristic_representation] [string trim dimension]]
                    ::tcom::foreach dcr $dcrs {
                      set id1 [[[[$dcr Attributes] Item [expr 1]] Value] P21ID]
                      if {$id1 == $anchorID} {
                        set id2 [$dcr P21ID]
                        if {[info exists spmiSumRowID($id2)]} {
                          set anchorSum($spmiSumRowID($id2)) $uuidval
                        }
                      }
                    }
                  }
                }
              } else {
                set badEnt 1
              }
            } emsg]} {
              errorMsg "Error missing entity #$anchorID for ANCHOR section."
            }
          } else {
            set badEnt 1
          }

# bad ID in anchor section
          if {$badEnt} {
            [[$worksheet($sect) Range [cellRange $r 1] [cellRange $r 1]] Interior] Color $legendColor(red)
            errorMsg "Syntax Error: Bad format for entity ID in ANCHOR section."
          }
        }
      }
      if {$line == "ENDSEC"} {[$worksheet($sect) Columns] AutoFit}
    }
  }

# -------------------------------------------------------------------------------------------------
# add UUIDs or anchor IDs to semantic PMI summary worksheet
  if {[info exists anchorSum]} {
    set spmiSumName "Semantic PMI Summary"
    set c 4
    if {$nistName != ""} {set c 5}

    if {[[$cells($spmiSumName) Item 3 $c] Value] == ""} {
      $cells($spmiSumName) Item 3 $c $heading
      set range [$worksheet($spmiSumName) Range [cellRange 3 $c]]
      addCellComment $spmiSumName 3 $c "See Help > User Guide (section 5.6)\n\nUUIDs for dimensional_characteristic_representation are found on the corresponding dimensional_location or dimensional_size entities."
      catch {foreach i {8 9} {[[$range Borders] Item $i] Weight [expr 2]}}
      [$range Font] Bold [expr 1]
      $range HorizontalAlignment [expr -4108]
    }

    set rmax 0
    foreach r [array names anchorSum] {
      $cells($spmiSumName) Item $r $c $anchorSum($r)
      if {$r > $rmax} {set rmax $r}
    }
    set range [$worksheet($spmiSumName) Range [cellRange 3 $c] [cellRange $rmax $c]]
    [$range Columns] AutoFit
  }

  foreach ent [array names urow] {
    $cells($ent) Item 3 $ucol($ent) $heading
    set msg "See ANCHOR worksheet and Help > User Guide (section 5.6)"
    if {[info exists entCount(v4_uuid_attribute)] || [info exists entCount(v5_uuid_attribute)]} {
      set msg "See Recommended Practices for Persistent IDs for Design Iteration and Downstream Exchange"
      incr urow($ent)
    }
    addCellComment $ent 3 $ucol($ent) $msg
    set range [$worksheet($ent) Range [cellRange 3 $ucol($ent)] [cellRange $urow($ent) $ucol($ent)]]
    [$range Columns] AutoFit
    set range [$worksheet($ent) Range [cellRange 3 $ucol($ent)]]
    [$range Interior] ColorIndex [expr 40]
    catch {[[$range Borders] Item [expr 8]] Weight [expr 3]}
    [$range Font] Bold [expr 1]
    $range HorizontalAlignment [expr -4108]
    catch {[[[$worksheet($ent) Range [cellRange $urow($ent) $ucol($ent)]] Borders] Item [expr 9]] Weight [expr 3]}
  }
}
