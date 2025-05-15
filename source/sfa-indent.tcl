proc indentPutLine {line {comment ""} {putstat 1}} {
  global indentWriteFile idlist idpatr lpatr npatr opt spatr

  if {$putstat != 1} {return 0}
  set stat 1

  set ll [string length $line]
  set t1 [string first "\#" $line]
  set t2 [string first "=" [string range $line [expr {$t1 + 1}] end]]
  set indent [expr {$t1 + $t2 + 2}]
  #outputMsg "($indent $t1 $t2) [string range $line $t1 [expr {$t1+$t2}]] $line"

# reset id list
  if {$t1 == 0} {
    set idlist ""
    catch {unset idpatr}
    catch {unset lpatr}
    catch {unset spatr}
    set npatr -1
  }
  set id [string range $line $t1 [expr {$t1+$t2}]]

# look for repeating patterns of entities and stop
  if {![info exists idlist]} {set idlist ""}
  if {[string first $id $idlist] != -1} {
    set id1 [string first $id $idpatr]
    #outputMsg "\n$id1 $id"
    if {$id1 != -1} {
      if {[info exists npatr]} {
        if {$id1 == 0} {
          for {set i 0} {$i < $npatr} {incr i} {
            #outputMsg $spatr($i)
            #outputMsg $idpatr
            #outputMsg [string first $spatr($i) [string range $idpatr [string length $spatr($i)] end]]
            #outputMsg "$i ($lpatr($i)) - $npatr ($lpatr($npatr))"
            if {$lpatr($i) == $lpatr($npatr)} {
              #outputMsg $idpatr
              errorMsg "Repeating pattern of entities"
              puts $indentWriteFile "(Repeating pattern of entities)"
              set stat -1
              return $stat
            }
          }
          incr npatr
        }
        lappend lpatr($npatr) $id1
        append spatr($npatr) $id
        #outputMsg "$npatr - $lpatr($npatr)"
      }
    }
    append idpatr $id
    #outputMsg PATR$idpatr
  } else {
    #outputMsg "RESET PATTERN"
    set idpatr ""
  }
  append idlist [string range $line $t1 [expr {$t1+$t2}]]
  #outputMsg LIST$idlist

# stop when indentation is out of control
  if {$indent > 40 && [string range $line 0 1] != "/*"} {
    outputMsg $line
    errorMsg "Indentation greater than 40 spaces" red
    puts $indentWriteFile "(Indentation greater than 40 spaces)"
    set stat 0
    return $stat
  }

# write short lines
  if {$ll <= 120} {
    #if {$comment != ""} {puts $indentWriteFile [string range $line 0 [expr {$t1-1}]]$comment}

# check for styled_item
    set ok 1
    if {[info exists opt(indentStyledItem)]} {if {!$opt(indentStyledItem) && [string first "STYLED_ITEM" $line] != -1} {set ok 0}}
    if {$ok} {puts $indentWriteFile $line}

# write longer lines
  } else {
    set line1 [string range $line 90 end]
    set p1 [string first ")"  $line1]
    set p2 [string first ");" $line1]
    if {$p1 == $p2} {
      puts $indentWriteFile $line

    } else {
      set line2 [string range $line1 [expr {$p1 + 1}] end]
      if {[string length $line2] > 9} {
        puts $indentWriteFile [string range $line 0 [expr {$p1 + 90}]]
        set line3 ""
        for {set i 0} {$i < $indent} {incr i} {append line3 " "}
        append line3 $line2

        set ll [string length $line3]
        if {$ll <= 120} {
          puts $indentWriteFile $line3

        } else {
          set line4 [string trim [string range $line3 90 end]]
          set p1 [string first ")"  $line4]
          set p2 [string first ");" $line4]
          if {$p1 == $p2} {
            puts $indentWriteFile $line3

          } else {
            set line5 [string range $line4 [expr {$p1 + 1}] end]
            if {[string length $line5] > 9} {
              puts $indentWriteFile [string range $line3 0 [expr {$p1 + 90}]]
              set line6 ""
              for {set i 0} {$i < $indent} {incr i} {append line6 " "}
              append line6 $line5
              puts $indentWriteFile $line6
            } else {
              puts $indentWriteFile $line3
            }
          }
        }
      } else {
        puts $indentWriteFile $line
      }
    }
  }
  return $stat
}

#-------------------------------------------------------------------------------
proc indentSearchLine {line ndent} {
  global comment indentEntity indentMissing indentstat indentStop

  incr ndent
  #outputMsg "------------- $ndent"

  set sline [split $line "\#"]
  for {set i 1} {$i < [llength $sline]} {incr i} {
    set str [lindex $sline $i]
    set id [indentGetID $str]
    if {$id > 2147483647} {errorMsg "An entity ID >= 2147483648 (2^31)"}

    set str "  "
    for {set j 1} {$j < $ndent} {incr j} {append str "  "}
    if {[info exists indentEntity($id)]} {
      append str $indentEntity($id)
      set cmnt ""
      if {[info exists comment($id)]} {set cmnt $comment($id)}
      if {![info exists indentstat]} {set indentstat 1}
      set indentstat [indentPutLine $str $cmnt $indentstat]
      if {$indentstat ==  0} {break}
      set line1 [string range $indentEntity($id) 1 end]
      if {[string first "\#" $line1] != -1} {
        set ok 1

# check for entities that stop the indentation
        foreach idx $indentStop {
          if {[string first $idx $line1] != -1} {
            set ok 0
            break
          }
        }
        if {$indentstat == -1} {set ok 0}

        if {$ok} {
          set stat [indentSearchLine $line1 $ndent]
          if {$stat == 0} {break}
        }
      }
    } elseif {$id != ""} {
      if {[lsearch $indentMissing "#$id"] == -1} {lappend indentMissing "#$id"}
    }
  }
  return 1
}

#-------------------------------------------------------------------------------
proc indentFile {ifile} {
  global editorCmd errmsg indentEntity indentMissing indentPass indentReadFile indentstat indentStart indentStop indentWriteFile opt writeDir

# more entities to start and stop with
  if {[info exists opt(indentStyledItem)]} {if {$opt(indentStyledItem)} {lappend indentStart STYLED_ITEM}}

  if {[info exists opt(indentGeometry)]} {
    if {!$opt(indentGeometry)} {
      lappend indentStop ADVANCED_FACE CLOSED_SHELL GEOMETRIC_CURVE_SET GEOMETRIC_SET CONSTRUCTIVE_GEOMETRY_REPRESENTATION TESSELLATED_SOLID
    }
  }
  if {[info exists opt(indentStyledItem)]} {if {!$opt(indentStyledItem)} {lappend indentStop STYLED_ITEM}}

  catch {.tnb select .tnb.status}
  outputMsg "Processing: [truncFileName [file nativename $ifile] 1]"
  outputMsg " Pass 1 of 2"
  set indentPass 1
  set indentReadFile [open $ifile r]

# same directory as file
  if {$opt(writeDirType) != 2} {
    set indentFileName [file rootname $ifile]

# user-defined directory
  } else {
    set indentFileName [file join $writeDir [file rootname [file tail $ifile]]]
  }
  append indentFileName "-sfa.txt"
  set indentWriteFile [open $indentFileName w]
  puts $indentWriteFile "Tree View generated by the NIST STEP File Analyzer and Viewer (v[getVersion])  [clock format [clock seconds]]\nSTEP file: [file nativename $ifile]\n"

  set cmnt ""
  foreach var {indentEntity idlist idpatr npatr lpatr spatr indentstat} {
    if {[info exists $var]} {unset $var}
  }
  if {[info exists errmsg]} {unset errmsg}
  set indentMissing {}

# read all entities
  set ihead 1
  while {[gets $indentReadFile line] >= 0} {
    set line [indentCheckLine $line]

    if {[string first "ENDSEC" $line] != -1} {set ihead 0}
    if {$ihead} {puts $indentWriteFile $line}

    if {[string first "\#" $line] == 0} {
      set id [string trim [string range $line 1 [expr {[string first "\=" $line] - 1}]]]
      set indentEntity($id) $line
      if {$cmnt != ""} {set comment($id) $cmnt}
      set cmnt ""
    } elseif {[string first "\/\*" $line] == 0} {
      set cmnt $line
    }
  }
  close $indentReadFile

  outputMsg " Pass 2 of 2"
  set indentPass 2
  set indentReadFile [open $ifile r]

# check for entities that start an indentation
  while {[gets $indentReadFile line] >= 0} {
    foreach var {idlist idpatr npatr lpatr spatr} {if {[info exists $var]} {unset $var}}
    set indentstat 1

    set line [indentCheckLine $line]
    set line1 [string range $line 1 end]
    if {[string first "\#" $line1] != -1} {
      foreach idx $indentStart {
        if {[string first $idx $line] != -1} {
          puts $indentWriteFile "\n"
          #outputMsg "PUT LINE 2"
          set stat [indentPutLine $line]
          if {$stat == 0} {break}
          #outputMsg "SEARCH LINE 2"
          set stat [indentSearchLine $line1 0]
          break
        }
      }
    }
  }
  close $indentReadFile
  close $indentWriteFile

  if {[llength $indentMissing] > 0} {errorMsg "Missing STEP entities: [lsort $indentMissing]"}

  if {$editorCmd != "" && [file size $indentFileName] < 100000000} {
    outputMsg "Opening Tree View STEP file:"
    exec $editorCmd [file nativename $indentFileName] &
  } else {
    outputMsg "Tree View STEP file written:"
  }
  outputMsg " [truncFileName [file nativename $indentFileName] 1] ([fileSize $indentFileName])" blue
}

#-------------------------------------------------------------------------------
proc indentCheckLine {line} {
  global indentReadFile

  if {([string last ";" $line] == -1 && [string last "*/" $line] == -1) || \
       [string range $line end end] == "(" || [string range $line end end] == ")" || [string range $line end end] == ","} {

# read more if the line is not complete
    if {[gets $indentReadFile line1] != -1} {
      if {[string length $line] < 500} {
        append line $line1
      } else {
        set iline [string range $line 0 500]
        set c1 [string last "," $iline]
        set iline "[string range $iline 0 $c1] (truncated)"
        if {[string last ";" $line1] != -1} {
          return $iline
        } else {
          while {1} {
            gets $indentReadFile line2
            if {[string last ";" $line2] != -1} {return $iline}
          }
        }
      }
      if {[catch {set line [indentCheckLine $line]} err]} {errorMsg $err}
      return $line
    } else {
      return $line
    }
  } else {
    set line [string trim $line]
    return $line
  }
}

#-------------------------------------------------------------------------------
proc indentGetID {id} {
  set p1 [string first "," $id]
  set p2 [string first "\)" $id]
  if {$p1 != -1 && $p2 != -1} {
    if {$p1 < $p2} {set id [string range $id 0 [expr {$p1-1}]]}
    if {$p1 > $p2} {set id [string range $id 0 [expr {$p2-1}]]}
  } elseif {$p1 != -1} {
    set id [string range $id 0 [expr {$p1-1}]]
  } elseif {$p2 != -1} {
    set id [string range $id 0 [expr {$p2-1}]]
  }
  set id [string trim $id]
  if {[string first "'" $id] != -1} {set id ""}
  return $id
}
