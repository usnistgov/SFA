proc tessPart {entType} {
  global objDesign
  global ao entLevel ent entAttrList opt tessCoordID

  if {$opt(DEBUG1)} {outputMsg "START tessPart $entType" red}
  set msg " Adding Tessellated Part Visualization"
  if {$opt(XLSCSV) == "None"} {append msg " ($entType)"}
  outputMsg $msg green

# faces
  set triangulated_face          [list triangulated_face name]
  set triangulated_surface_set   [list triangulated_surface_set name]
  set complex_triangulated_face  [list complex_triangulated_face name]
  set complex_triangulated_surface_set [list complex_triangulated_surface_set name]
  set PMIP(tessellated_solid) [list tessellated_solid name items \
    $triangulated_face $complex_triangulated_face $complex_triangulated_surface_set $triangulated_surface_set]
  set PMIP(tessellated_shell) [list tessellated_shell name items \
    $triangulated_face $complex_triangulated_face $complex_triangulated_surface_set $triangulated_surface_set]
   
# initialize
  set ao $entType
  set entAttrList {}
  set tessCoordID {}
  if {[info exists ent]} {unset ent}
  set entLevel 0
  setEntAttrList $PMIP($ao)
  if {$opt(DEBUG1)} {outputMsg "entattrlist $entAttrList"}

# process entities, call tessPartGeometry
  ::tcom::foreach objEntity [$objDesign FindObjects [join $entType]] {
    if {[$objEntity Type] == $entType} {tessPartGeometry $objEntity}
  }
}

# -------------------------------------------------------------------------------
proc tessPartGeometry {objEntity} {
  global ao badAttributes entLevel ent entAttrList objEntity1 opt tessCoord tessIndex tessIndexCoord x3dStartFile

  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]

  if {[string first "handle" $objEntity] != -1} {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set objAttributes [$objEntity Attributes]
    set ent($entLevel) $objType
    if {$entLevel == 1} {set objEntity1 $objEntity}

    if {$opt(DEBUG1)} {outputMsg "$ind ENT $entLevel #$objID=$objType (ATR=[$objAttributes Count])" blue}

    ::tcom::foreach objAttribute $objAttributes {
      set objName  [$objAttribute Name]
      set ent1 "$ent($entLevel) $objName"

# look for entities with bad attributes that cause a crash
      set okattr 1
      if {[info exists badAttributes($objType)]} {foreach ba $badAttributes($objType) {if {$ba == $objName} {set okattr 0}}}

      if {$okattr} {
        set objValue    [$objAttribute Value]
        set objNodeType [$objAttribute NodeType]
        set objSize     [$objAttribute Size]
        set objAttrType [$objAttribute Type]
        set idx [lsearch $entAttrList $ent1]

        if {$objNodeType == 20} {
          if {$idx != -1} {
            if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}
            ::tcom::foreach val1 $objValue {tessPartGeometry $val1}
          }
        } elseif {$objNodeType == 5} {
          if {[catch {
            if {$idx != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objAttrType)  ($ent1)"}
    
# get values for these entity and attribute pairs
              switch -glob $ent1 {
                "tessellated_shell name" -
                "tessellated_solid name" {
# start X3DOM file, read tessellated geometry
                  if {$x3dStartFile} {x3dFileStart}
                }
                "*triangulated_face name" -
                "*triangulated_surface_set name" -
                "tessellated_curve_set name" {
# write tessellated coords and index for part geometry
                  if {$ao == "tessellated_solid" || $ao == "tessellated_shell"} {
                    if {[info exists tessIndex($objID)] && [info exists tessCoord($tessIndexCoord($objID))]} {
                      x3dTessGeom $objID $objEntity1 $ent1
                    } else {
                      errorMsg "Missing tessellated coordinates and index for $objID"
                    }
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing Tessellated Part Geometry: $emsg3"
            set entLevel 2
          }
        }
      }
    }
  }
  incr entLevel -1
}

# -------------------------------------------------------------------------------
proc tessReadGeometry {} {
  global entCount localName tessCoord tessIndex tessIndexCoord developer x3dMin x3dMax x3dMsg
  
  set tg [open $localName r]
  set ncl  0
  set ntc  0
  set ntc1 0
  foreach ent {tessellated_curve_set \
              complex_triangulated_surface_set triangulated_surface_set \
              complex_triangulated_face triangulated_face} {
    if {[info exists entCount($ent)]} {set ntc1 [expr {$ntc1+$entCount($ent)}]}
  }
  outputMsg " Reading tessellated geometry" green
  
# read step
  while {[gets $tg line] >= 0} {
    if {[string first "COORDINATES_LIST" $line] != -1 || \
        [string first "TESSELLATED_CURVE_SET" $line] != -1 || \
        [string first "TRIANGULATED_FACE" $line] != -1 || \
        [string first "COMPLEX_TRIANGULATED_FACE" $line] != -1 || \
        [string first "TRIANGULATED_SURFACE_SET" $line] != -1 || \
        [string first "COMPLEX_TRIANGULATED_SURFACE_SET" $line] != -1} {

# get rest of entity if one multiple lines
      while {1} {
        if {[string first ";" $line] == -1} {
          gets $tg line1
          append line $line1
        } else {
          break
        }
      }

# entity ID
      set id [string range $line 1 [string first "=" $line]-1]

# coordinates_list
      if {[string first "COORDINATES_LIST" $line] != -1} {
        set ncoord [string range $line [string first "," $line]+1 [string first "((" $line]-2]
        if {$ncoord > 0} {
          set line [string range $line [string first "((" $line] end-1]
          regsub -all {[(),]} $line " " line
          set line [string trim $line]
          regsub -all "  " $line " " line
          regsub -all "  " $line " " line
          regsub -all "  " $line " " line
          regsub -all "  " $line " " line
          set tessCoord($id) " "

          set sline [split $line " "]
          for {set j 0} {$j < [llength $sline]} {incr j 3} {
            set x3dPoint(x) [trimNum [lindex $sline $j] 4]
            set x3dPoint(y) [trimNum [lindex $sline [expr {$j+1}]] 4]
            set x3dPoint(z) [trimNum [lindex $sline [expr {$j+2}]] 4]
            foreach idx {x y z} {
              append tessCoord($id) "$x3dPoint($idx) "
              if {$x3dPoint($idx) > $x3dMax($idx)} {set x3dMax($idx) $x3dPoint($idx)}
              if {$x3dPoint($idx) < $x3dMin($idx)} {set x3dMin($idx) $x3dPoint($idx)}
            }
          }
        } else {
          set msg "Syntax Error: #$id=COORDINATES_LIST has zero coordinates"
          errorMsg $msg
          set msg "<i>#$id=COORDINATES_LIST has no coordinates</i>"
          if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
        }
        incr ncl

# tessellated_curve_set, *triangulated_surface_set, *triangulated_face
      } elseif {[string first "TESSELLATED_CURVE_SET" $line] != -1 || [string first "TRIANGULATED" $line] != -1} {
        set complex 0
        if {[string first "COMPLEX" $line] != -1} {set complex 1}

# id of coordinate list
        set c1 [string first "," $line]
        set c2 [string first "((" $line]
        set tessIndexCoord($id) [string range $line $c1+2 $c2-2]
        set c1 [string first "," $tessIndexCoord($id)]
        if {$c1 != -1} {set tessIndexCoord($id) [string range $tessIndexCoord($id) 0 $c1-1]}
        
# pnindex (optional)
        catch {unset pnindex}
        set c1 [string first "))" $line]
        set nline [string range $line $c1+3 end]
        set nline [string range $nline [string first "(" $nline]+1 [string first ")" $nline]-1]
        set nline [split $nline ","]
        set n 0
        foreach item $nline {set pnindex([incr n]) [string trim $item]}

# tessellated curve set
        if {[string first "TESSELLATED_CURVE_SET" $line] != -1} {
          set line [string range $line [string last "((" $line] end-1]
          regsub -all " " $line "" line
          regsub -all {[(),]} $line "x" line
          regsub -all "xxx" $line " 0 " line
          regsub -all "x" $line " " line

          set tessIndex($id) ""
          foreach j [split $line " "] {if {$j != ""} {append tessIndex($id) "[expr {$j-1}] "}}

# *triangulated surface set, *triangulated face
        } else {
          set line [string range $line [string first "((" $line]+2 end-1]
          set line [string trim [string range $line [string first "((" $line] end]]
          
# regsub is very important to distill line into something usable
          regsub -all " " $line "" line
          regsub -all {[(),]} $line "x" line
          regsub -all "xxx" $line " 0 " line
          regsub -all "x" $line " " line
          set line [string trim $line]
          set c1 [string first "   " $line]
          if {$c1 == -1} {set c1 [string first "  " $line]}

          set c2 1
          if {$complex && $c1 != -1} {set c2 3}
          set strip [string trim [string range $line 0 $c1-1]]
          set fan   [string trim [string range $line $c1+$c2 end]]
          set tessIndex($id) ""

# triangle strip
          if {[string length $strip] > 0} {
            set strip [split $strip " "]
            for {set j 0} {$j < [llength $strip]} {incr j} {
              foreach k {0 1 2} {
                set s($k) [lindex $strip [expr {$j+$k}]]
                if {$s($k) != 0 && [info exists pnindex($s($k))]} {set s($k) $pnindex($s($k))}
              }
              if {$s(0) != "0" && $s(1) != "0" && $s(2) != "0"} {
                append tessIndex($id) "[expr {$s(0)-1}] [expr {$s(1)-1}] [expr {$s(2)-1}] -1 "
              }
            }
          }

# triangle fan
          if {[string length $fan] > 0} {
            if {$complex} {
              regsub -all " 0 " $fan "x" fan
              set fan [split $fan "x"]
              foreach f $fan {
                for {set j 1} {$j < [llength $f]} {incr j} {
                  if {[lindex $f $j+1] != ""} {
                    set f1(0) [lindex $f 0]
                    set f1(1) [lindex $f $j]
                    set f1(2) [lindex $f $j+1]
                    if {$f1(0) != "0" && $f1(1) != "0" && $f1(2) != "0"} {
                      foreach k {0 1 2} {if {[info exists pnindex($f1($k))]} {set f1($k) $pnindex($f1($k))}}
                      append tessIndex($id) "[expr {$f1(0)-1}] [expr {$f1(1)-1}] [expr {$f1(2)-1}] -1 "
                    }
                  }
                }
              }

# triangle face              
            } else {
              if {[llength [array names pnindex]] > 0} {set pnindex(0) 0}
              set fan [split $fan " "]
              foreach f $fan {
                if {[info exists pnindex($f)]} {
                  append tessIndex($id) "[expr {$pnindex($f)-1}] "
                } else {
                  append tessIndex($id) "[expr {$f-1}] "
                }
              }
            }
          }
        }
        incr ntc
      }
    }
    
# done reading tessellated geometry    
    if {$ncl == $entCount(coordinates_list) && $ntc == $ntc1} {
      outputMsg "  [expr {$ncl+$ntc}] tessellated geometry entities"
      close $tg
      return
    }
  }
}

# -------------------------------------------------------------------------------
proc tessCountColors {} {
  global objDesign
  global entCount

# count unique colors in colour_rgb and draughting_pre_defined_colour
  set colors {}
  if {[catch {
    if {[info exists entCount(colour_rgb)]} {
      ::tcom::foreach e0 [$objDesign FindObjects [string trim colour_rgb]] {
        set a0 [$e0 Attributes]
        set color "[[$a0 Item 2] Value] [[$a0 Item 3] Value] [[$a0 Item 4] Value]"
        if {[lsearch $colors $color] == -1} {lappend colors $color}
      }
    }
    if {[info exists entCount(draughting_pre_defined_colour)]} {
      ::tcom::foreach e0 [$objDesign FindObjects [string trim draughting_pre_defined_colour]] {
        set color [[[$e0 Attributes] Item 1] Value]
        if {[lsearch $colors $color] == -1} {lappend colors $color}
      }
    }
  } emsg]} {
    errorMsg " ERROR counting unique colors: $emsg"
  }
  return [llength $colors]
}

# -------------------------------------------------------------------------------
proc tessSetColor {tsID} {
  global objDesign
  global tessColor x3dColor entCount recPracNames

  set ok 0
  if {[info exists entCount(styled_item)]} {
    if {$entCount(styled_item) > 0} {set ok 1}
  }

  if {[info exists tessColor($tsID)]} {
    set x3dColor $tessColor($tsID)
  } elseif {$ok} {

# get color from styled_item reference to tessellated_solid
    if {[catch {
      ::tcom::foreach e0 [$objDesign FindObjects [string trim styled_item]] {
        set a0 [[$e0 Attributes] Item 3]
# presentation_style_assignment
        set e1 [$a0 Value]
        if {[$e1 P21ID] == $tsID} {
          set a1 [[$e0 Attributes] Item 2]
# surface_style_usage
          ::tcom::foreach e2 [$a1 Value] {
            set a2 [[$e2 Attributes] Item 1]

            set e3 [$a2 Value]
            set a3 [[$e3 Attributes] Item 2]
# surface side style
            set e4 [$a3 Value]
            set a4 [[$e4 Attributes] Item 2]
# surface style fill area
            set e5 [$a4 Value]
            set a5 [[$e5 Attributes] Item 1]
# fill area style
            set e6 [$a5 Value]
            set a6 [[$e6 Attributes] Item 2]
# fill area style colour
            set e7 [$a6 Value]
            set a7 [[$e7 Attributes] Item 2]
# color
            set e8 [$a7 Value]
            if {[$e8 Type] == "colour_rgb"} {
              set x3dColor ""
              set j 0
              ::tcom::foreach a8 [$e8 Attributes] {
                if {$j > 0} {append x3dColor "[trimNum [$a8 Value] 3] "}
                incr j
              }
              set x3dColor [string trim $x3dColor]
              set tessColor($tsID) $x3dColor
            } elseif {[$e8 Type] == "draughting_pre_defined_colour"} {
              ::tcom::foreach a8 [$e8 Attributes] {
                switch [$a8 Value] {
                  black   {set x3dColor "0 0 0"}
                  white   {set x3dColor "1 1 1"}
                  red     {set x3dColor "1 0 0"}
                  yellow  {set x3dColor "1 1 0"}
                  green   {set x3dColor "0 1 0"}
                  cyan    {set x3dColor "0 1 1"}
                  blue    {set x3dColor "0 0 1"}
                  magenta {set x3dColor "1 0 1"}
                  default {
                    set x3dColor ".7 .7 .7"
                    errorMsg "Syntax Error: Unknown draughting_pre_defined_colour name '$objValue' (using gray)\n[string repeat " " 14]($recPracNames(model), Sec. 4.2.3, Table 2)"
                  }
                }
              }
              set tessColor($tsID) $x3dColor
            } else {
              errorMsg "  Unexpected color type ([$e8 Type]) for $ao"
              set tessColor($tsID) $x3dColor
            }
          }
        }
      }
    } emsg]} {
      errorMsg " Error setting tessellated geometry color, using gray"
      set tessColor($tsID) $x3dColor
      update idletasks
    }
  } else {
    errorMsg "Syntax Error: No 'styled_item' found to assign color to tessellated geometry (using gray)\n[string repeat " " 14]($recPracNames(model), Sec. 4.2.2, Fig. 2)"
  }
}

# -------------------------------------------------------------------------------
proc tessSetPlacement {tsID} {
  global objDesign
  global tessPlacement tessRepo shapeRepName
  
  if {[catch {
    set tessRepo 0
    catch {unset tessPlacement}
    set debug 0
    
# tessellated_shape_representation
    ::tcom::foreach e0 [$objDesign FindObjects [string trim tessellated_shape_representation]] {

# get TSR.items
      set a0 [[$e0 Attributes] Item 2]

# for each entity in TSR.items
      ::tcom::foreach e1 [$a0 Value] {
        
# compare entity ID to tessellated_solid ID        
        if {[$e1 P21ID] == $tsID} {
          if {$debug} {errorMsg "\n[$e1 Type] [$e1 P21ID] ([$a0 Name])" red}

          #foreach repSRR {rep_1 rep_2} 
          foreach repSRR {rep_1} {
            set e2s [$e0 GetUsedIn [string trim shape_representation_relationship] [string trim $repSRR]]

            ::tcom::foreach e2 $e2s {
              if {$debug} {errorMsg " [$e2 Type] [$e2 P21ID]" red}
              set a2 [[$e2 Attributes] Item 4]
              set e3 [$a2 Value]
              if {$debug} {errorMsg "  [$e3 Type] [$e3 P21ID] ([$a2 Name])" red}

              if {[$e3 Type] != "shape_representation"} {
                set a2 [[$e2 Attributes] Item 3]
                set e3 [$a2 Value]
                if {$debug} {errorMsg "  [$e3 Type] [$e3 P21ID] ([$a2 Name])" blue}
              }

              set shapeRepName [string trim [[[$e3 Attributes] Item 1] Value]]
              if {$shapeRepName != ""} {
                if {$debug} {errorMsg "  $shapeRepName" blue}
              } else {
                unset shapeRepName
              }

              set e4s [$e3 GetUsedIn [string trim shape_definition_representation] [string trim used_representation]]
              ::tcom::foreach e4 $e4s {
                set a4 [[$e4 Attributes] Item 1]
                set e5 [$a4 Value]
                if {$debug} {errorMsg "   [$e5 Type] [$e5 P21ID] ([$a4 Name])" green}
                set a5 [[$e5 Attributes] Item 3]
                set e6 [$a5 Value]
                if {$debug} {errorMsg "    [$e6 Type] [$e6 P21ID] ([$a5 Name])" green}
                foreach rel {relating_product_defintion related_product_definition} {
                  set e7s [$e6 GetUsedIn [string trim next_assembly_usage_occurrence] [string trim $rel]]
                  ::tcom::foreach e7 $e7s {
                    set a7 [[$e7 Attributes] Item 5]
                    set e8 [$a7 Value]
                    if {[$e8 P21ID] == [$e6 P21ID]} {
                      if {$debug} {errorMsg "     [$e7 Type] [$e7 P21ID] ([$a7 Name])" green}
                    }
                  }
                }
              }
              
              foreach repRRWT {rep_1 rep_2} {
                set e4s [$e3 GetUsedIn [string trim representation_relationship_with_transformation_and_shape_representation_relationship] [string trim $repRRWT]]
                ::tcom::foreach e4 $e4s {
                  if {$debug} {errorMsg "   [$e4 Type] [$e4 P21ID]" red}
                  set a4 [[$e4 Attributes] Item 5]
                  set e5 [$a4 Value]
                  if {$debug} {errorMsg "    [$e5 Type] [$e5 P21ID] ([$a4 Name])" red}
                  set a5 [[$e5 Attributes] Item 4]
                  set e6 [$a5 Value]
                  if {$debug} {errorMsg "     [$e6 Type] [$e6 P21ID] ([$a5 Name])" red}

# a2p3d origin
                  set a6 [[$e6 Attributes] Item 2]
                  set e7 [$a6 Value]
                  set val ""
                  foreach n [[[$e7 Attributes] Item 2] Value] {append val "[trimNum $n 5] "}
                  lappend tessPlacement(origin) [string trim $val]
                  if {$debug} {errorMsg "      [$e7 Type] [$e7 P21ID] ([$a6 Name]) [string trim $val]" red}

# a2p3d axis
                  set a6 [[$e6 Attributes] Item 3]
                  set e7 [$a6 Value]
                  lappend tessPlacement(axis) [[[$e7 Attributes] Item 2] Value]
                  if {$debug} {errorMsg "      [$e7 Type] [$e7 P21ID] ([$a6 Name]) [[[$e7 Attributes] Item 2] Value]" red}

# a2p3d reference direction
                  set a6 [[$e6 Attributes] Item 4]
                  set e7 [$a6 Value]
                  lappend tessPlacement(refdir) [[[$e7 Attributes] Item 2] Value]
                  set tessRepo 1
                  if {$debug} {errorMsg "      [$e7 Type] [$e7 P21ID] ([$a6 Name]) [[[$e7 Attributes] Item 2] Value]" red}
                }
              }
            }
          }
        }
      }
    }
  } emsg]} {
    errorMsg " ERROR getting tessellated geometry placement: $emsg"
  }
}

# -------------------------------------------------------------------------------
# generate x3d rotation from axis2_placement_3d
proc x3dRotation {{a {0 0 1}} {r {1 0 0}}} {
  
# convert from geometry coordinate system to x3dom coordinate, Zv = -Yc, Yv = Zc
  set axis   [list [lindex $a 0] [lindex $a 2] [expr {0.-[lindex $a 1]}]]
  set refdir [list [lindex $r 0] [lindex $r 2] [expr {0.-[lindex $r 1]}]]
  #set axis $a
  #set refdir $r
    
# construct rotation matrix u, must normalize to use with quaternion
  set u2 [vecnorm $axis]
  set u1 [vecnorm [vecsub $refdir [vecmult $u2 [vecdot $refdir $u2]]]]
  set u3 [vecnorm [veccross $u1 $u2]]

# extract quaternion
  if {[lindex $u1 0] >= 0.0} {
    set tmp [expr {[lindex $u2 1] + [lindex $u3 2]}]
    if {$tmp >=  0.0} {
      set q(0) [expr {[lindex $u1 0] + $tmp + 1.}]
      set q(1) [expr {[lindex $u3 1] - [lindex $u2 2]}]
      set q(2) [expr {[lindex $u1 2] - [lindex $u3 0]}]
      set q(3) [expr {[lindex $u2 0] - [lindex $u1 1]}]
    } else {
      set q(0) [expr {[lindex $u3 1] - [lindex $u2 2]}]
      set q(1) [expr {[lindex $u1 0] - $tmp + 1.}]
      set q(2) [expr {[lindex $u2 0] + [lindex $u1 1]}]
      set q(3) [expr {[lindex $u1 2] + [lindex $u3 0]}]
    }
  } else {
    set tmp [expr {[lindex $u2 1] - [lindex $u3 2]}]
    if {$tmp >= 0.0} {
      set q(0) [expr {[lindex $u1 2] - [lindex $u3 0]}]
      set q(1) [expr {[lindex $u2 0] + [lindex $u1 1]}]
      set q(2) [expr {1. - [lindex $u1 0] + $tmp}]
      set q(3) [expr {[lindex $u3 1] + [lindex $u2 2]}]
    } else {
      set q(0) [expr {[lindex $u2 0] - [lindex $u1 1]}]
      set q(1) [expr {[lindex $u1 2] + [lindex $u3 0]}]
      set q(2) [expr {[lindex $u3 1] + [lindex $u2 2]}]
      set q(3) [expr {1. - [lindex $u1 0] - $tmp}]
    }
  }

# normalize quaternion
  set lenq [expr {sqrt($q(0)*$q(0) + $q(1)*$q(1) + $q(2)*$q(2) + $q(3)*$q(3))}]
  if {$lenq != 0.} {
    foreach i {0 1 2 3} {set q($i) [expr {$q($i) / $lenq}]}
  } else {
    foreach i {0 1 2 3} {set q($i) 0.}
  }

# convert from quaterion to rotation
  set rotation_changed {0 1 0 0}
  set angle [expr {acos($q(0))*2.0}]
  if {$angle != 0.} {
    set sina [expr {sin($angle*0.5)}]
    set axm 0.
    foreach i {0 1 2} {
      set i1 [expr {$i+1}]
      set ax [expr {-$q($i1) / $sina}]
      lset rotation_changed $i $ax
      set axa [expr {abs($ax)}]
      if {$axa > $axm} {set axm $axa}
    }
    if {$axm > 0. && $axm < 1.} {
      foreach i {0 1 2} {lset rotation_changed $i [expr {[lindex $rotation_changed $i]/$axm}]}
    }
    lset rotation_changed 3 $angle
    foreach i {0 1 2 3} {lset rotation_changed $i [trimNum [lindex $rotation_changed $i] 4]}
  }
  return $rotation_changed  
}

#-------------------------------------------------------------------------------
# dot - calculate scalar dot product of two vectors
proc vecdot {v1 v2} {
  set v3 0.0
  foreach c1 $v1 c2 $v2 {set v3 [expr {$v3+$c1*$c2}]}
  return $v3
}

# sub - multiply two vectors
proc vecmult {v1 v2} {
  foreach c1 $v1 {lappend v3 [expr {$c1*$v2}]}
  return $v3
}

# sub - subtract one vector from another
proc vecsub {v1 v2} {
  foreach c1 $v1 c2 $v2 {lappend v3 [expr {$c1-$c2}]}
  return $v3
}

# cross - cross product between two 3d-vectors
proc veccross {v1 v2} {
  set v1x [lindex $v1 0]
  set v1y [lindex $v1 1]
  set v1z [lindex $v1 2]
  set v2x [lindex $v2 0]
  set v2y [lindex $v2 1]
  set v2z [lindex $v2 2]
  set v3 [list [expr {$v1y*$v2z-$v1z*$v2y}] [expr {$v1z*$v2x-$v1x*$v2z}] [expr {$v1x*$v2y-$v1y*$v2x}]]
  return $v3
}

# len - get scalar length of a vector
proc veclen {v1} {
 set l 0.
 foreach c1 $v1 {set l [expr {$l + $c1*$c1}]}
 return [expr {sqrt($l)}]
}

# norm - normalize a vector
proc vecnorm {v1} {
  set l [veclen $v1]
  if {$l != 0.} {
    set s [expr {1./$l}]
    foreach c1 $v1 {lappend v2 [expr {$c1*$s}]}
  } else {
    set v2 $v1
  }
  return $v2
}
