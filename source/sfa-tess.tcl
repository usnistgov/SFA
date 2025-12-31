proc tessPart {entType} {
  global objDesign
  global ao ent entAttrList entCount entLevel mytemp opt
  global tessPartFile tessPartFileName

  if {$opt(DEBUG1)} {outputMsg "START tessPart $entType" red}
  set msg " Adding Tessellated Part Geometry"
  if {$opt(xlFormat) == "None"} {append msg " ($entType)"}
  if {[string first "annotation" $msg] == -1} {outputMsg $msg green}

# edge, faces
  set tessellated_edge            [list tessellated_edge name coordinates line_strip]
  set tessellated_connecting_edge [list tessellated_connecting_edge name coordinates line_strip]

  set triangulated_face                [list triangulated_face name]
  set triangulated_surface_set         [list triangulated_surface_set name]
  set complex_triangulated_face        [list complex_triangulated_face name]
  set complex_triangulated_surface_set [list complex_triangulated_surface_set name]

  set PMIP(tessellated_solid) [list tessellated_solid name items \
    $tessellated_edge $tessellated_connecting_edge $triangulated_face $complex_triangulated_face $complex_triangulated_surface_set $triangulated_surface_set]
  set PMIP(tessellated_shell) [list tessellated_shell name items \
    $tessellated_edge $tessellated_connecting_edge $triangulated_face $complex_triangulated_face $complex_triangulated_surface_set $triangulated_surface_set]
  set PMIP(tessellated_wire)  [list tessellated_wire name items $tessellated_edge $tessellated_connecting_edge]

  set tessellated_face_surface [list tessellated_face_surface name tessellated_face_geometry $triangulated_face $complex_triangulated_face]
  set PMIP(tessellated_closed_shell) [list tessellated_closed_shell name cfs_faces $tessellated_face_surface]
  set PMIP(tessellated_open_shell)   [list tessellated_open_shell   name cfs_faces $tessellated_face_surface]

# initialize
  set ao $entType
  set entAttrList {}
  catch {unset ent}
  set entLevel 0
  setEntAttrList $PMIP($ao)
  if {$opt(DEBUG1)} {outputMsg "entattrlist $entAttrList"}

# open file for tessellated parts
  checkTempDir
  foreach tess {tessellated_solid tessellated_shell tessellated_closed_shell tessellated_open_shell tessellated_wire} {
    if {[info exists entCount($tess)] && ![info exists tessPartFileName]} {
      if {$entCount($tess) > 0} {
        set tessPartFileName [file join $mytemp tessPart.txt]
        catch {file delete -force -- $tessPartFileName}
        set tessPartFile [open $tessPartFileName w]
      }
    }
  }

# process entities, call tessPartGeometry
  ::tcom::foreach objEntity [$objDesign FindObjects [join $entType]] {
    if {[$objEntity Type] == $entType} {tessPartGeometry $objEntity}
  }
}

# -------------------------------------------------------------------------------
proc tessPartGeometry {objEntity} {
  global ao badAttributes ent entAttrList entLevel gen objEntity1 opt tessCoord
  global tessEdgeCoord tessEdges tessIndex tessIndexCoord x3dFile x3dStartFile

  if {$opt(DEBUG1)} {outputMsg "tessPartGeometry [$objEntity Type] [$objEntity P21ID]" red}
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
            switch -glob $ent1 {
              "tessellated_*edge line_strip" {
# tessellated edges, store to write to x3dFile later
                set line_strip {}
                foreach item $objValue {lappend line_strip [expr {$item-1}]}
                lappend tessEdges($tessEdgeCoord) "[join $line_strip] -1"
              }
            }

# recursively get the entities that are referred to
            if {[catch {
              ::tcom::foreach val1 $objValue {tessPartGeometry $val1}
            } emsg]} {
              foreach val2 $objValue {tessPartGeometry $val2}
            }
          }

        } elseif {$objNodeType == 5} {
          if {[catch {
            if {$idx != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objAttrType)  ($ent1)"}

# get values for these entity and attribute pairs
              switch -glob $ent1 {
                "tessellated_shell name" -
                "tessellated_closed_shell name" -
                "tessellated_open_shell name" -
                "tessellated_solid name" {
# start x3dom file, read tessellated geometry
                  if {$gen(View) && $x3dStartFile} {x3dFileStart}
                }
                "*triangulated_face name" -
                "*triangulated_surface_set name" -
                "tessellated_curve_set name" {
# write tessellated coords and index for part geometry (objEntity1 - tessellated_* entity, objEntity - triangulated face entity)
                  if {$ao == "tessellated_solid" || $ao == "tessellated_shell" || $ao == "tessellated_closed_shell" || $ao == "tessellated_open_shell"} {
                    if {[info exists tessIndex($objID)] && [info exists tessCoord($tessIndexCoord($objID))]} {
                      x3dTessGeom $objID $objEntity1 $objEntity
                    } else {
                      errorMsg "Missing tessellated coordinates and index"
                    }
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "Error processing Tessellated Part Geometry: $emsg3"
            set entLevel 2
          }

        } elseif {$objNodeType == 18 || $objNodeType == 19} {
          if {$idx != -1} {
            if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType-[$objAttribute AggNodeType], $objSize, $objAttrType)"}
            switch -glob $ent1 {
              "tessellated_*edge coordinates" {
# tessellated edge coordinate list
                set tessEdgeCoord [$objValue P21ID]
              }
            }

# recursively get the entities that are referred to
            if {[catch {
              ::tcom::foreach val1 $objValue {tessPartGeometry $val1}
            } emsg]} {
              foreach val2 $objValue {tessPartGeometry $val2}
            }
          }
        }
      }
    }
  }
  incr entLevel -1
}

# -------------------------------------------------------------------------------
# generate coordinates and faces by reading tessellated geometry entities line-by-line from STEP file, necessary because of limitations of the IFCsvr toolkit
# if coordOnly=1, then only read coordinates_list, used for saved view pmi validation properties
# if coordOnly=2, then only read point_cloud_dataset
proc tessReadGeometry {{coordOnly 0}} {
  global entCount localName opt recPracNames syntaxErr spaces
  global tessBrep tessCoord tessCoordName tessIndex tessIndexCoord tessMinMax tessReadOnce x3dBbox x3dMax x3dMin
  global objDesign

  if {[info exists tessReadOnce]} {return}
  set tessReadOnce 1

  set tg [open $localName r]
  set ncl  0
  set ntc  0
  set ntc1 0
  set tessellated 0

# read everything
  if {$coordOnly == 0} {
    foreach ent {tessellated_curve_set complex_triangulated_surface_set triangulated_surface_set complex_triangulated_face triangulated_face} {
      if {[info exists entCount($ent)]} {set ntc1 [expr {$ntc1+$entCount($ent)}]}
    }
    foreach ent [list COORDINATES_LIST TESSELLATED_CURVE_SET TRIANGULATED_FACE COMPLEX_TRIANGULATED_FACE TRIANGULATED_SURFACE_SET COMPLEX_TRIANGULATED_SURFACE_SET] {
      if {[info exists entCount([string tolower $ent])]} {lappend ents $ent}
    }
    outputMsg " Reading tessellated geometry" green

# which coordinates_list are associated with tessellated geometry entities for checking xyz min/max
    if {$opt(tessPartOld) || $tessBrep || [info exists entCount(tessellated_closed_shell)]} {
      set tsCoordID {}
      set tessMinMax 1
      foreach item [list tessellated_solid tessellated_shell] {
        if {[info exists entCount($item)] && $entCount($item) > 0} {
          ::tcom::foreach ts [$objDesign FindObjects [string trim $item]] {
            set e0s [[[$ts Attributes] Item [expr 2]] Value]
            ::tcom::foreach e0 $e0s {lappend tsCoordID [[[[$e0 Attributes] Item [expr 2]] Value] P21ID]}
          }
        }
      }
      if {[info exists entCount(tessellated_closed_shell)] && $entCount(tessellated_closed_shell) > 0} {
        ::tcom::foreach tcs [$objDesign FindObjects [string trim tessellated_closed_shell]] {
          set e0s [[[$tcs Attributes] Item [expr 2]] Value]
          ::tcom::foreach e0 $e0s {
            set e1 [[[$e0 Attributes] Item [expr 3]] Value]
            lappend tsCoordID [[[[$e1 Attributes] Item [expr 2]] Value] P21ID]
          }
        }
      }
      set tsCoordID [lrmdups $tsCoordID]
      if {[llength $tsCoordID] == 0} {unset tsCoordID}
    }

# read only coordinates list
  } elseif {$coordOnly == 1} {
    foreach ent {coordinates_list} {if {[info exists entCount($ent)]} {set ntc1 [expr {$ntc1+$entCount($ent)}]}}
    set ents [list COORDINATES_LIST]
    outputMsg " Reading coordinates list" green

# read only point cloud dataset
  } elseif {$coordOnly == 2} {
    foreach ent {point_cloud_dataset} {if {[info exists entCount($ent)]} {set ntc1 [expr {$ntc1+$entCount($ent)}]}}
    set ents [list POINT_CLOUD_DATASET]
    outputMsg " Processing point cloud" green
    catch {unset tessCoord}
  }

# read step
  while {[gets $tg line] >= 0} {
    set ok 0
    foreach ent $ents {if {[string first $ent $line] != -1} {set ok 1; break}}
    if {$ok} {

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
      if {[string first "TRIANGULATED" $line] != -1 || [string first "CURVE" $line] != -1} {set tessellated 1}

# coordinates_list or point_cloud_dataset
      if {[string first "COORDINATES_LIST" $line] != -1 || [string first "POINT_CLOUD_DATASET" $line] != -1} {
        set tessCoordName($id) ""
        set c1 [string first "'" $line]
        set c2 [string last  "'" $line]
        if {$c2 != [expr {$c1+1}]} {set tessCoordName($id) [string range $line $c1+1 $c2-1]}

        set ncoord 0
        if {$coordOnly != 2} {set ncoord [string range $line [string first "," $line]+1 [string first "((" $line]-2]}

        if {$ncoord > 0 || $coordOnly == 2} {
          set line [string range $line [string first "((" $line] end-1]

# regsub is very important to distill line into something usable
          regsub -all {[(),]} $line " " line
          set line [string trim $line]
          for {set i 0} {$i < 4} {incr i} {regsub -all "  " $line " " line}
          set tessCoord($id) " "

          set sline [split $line " "]
          set prec 3
          for {set j 0} {$j < [llength $sline]} {incr j 3} {
            set x3dPoint(x) [lindex $sline $j]
            set x3dPoint(y) [lindex $sline [expr {$j+1}]]
            set x3dPoint(z) [lindex $sline [expr {$j+2}]]
            set tc ""
            foreach idx {x y z} {
              set prec 3
              if {[expr {abs($x3dPoint($idx))}] < 0.01} {set prec 4}
              append tc "[trimNum $x3dPoint($idx) $prec] "

# get min/max
              if {$tessMinMax} {
                set ok1 0
                if {![info exists tsCoordID] || [lsearch $tsCoordID $id] != -1} {set ok1 1}
                if {$coordOnly == 0 && $ok1} {
                  if {$x3dPoint($idx) > $x3dMax($idx)} {set x3dMax($idx) $x3dPoint($idx)}
                  if {$x3dPoint($idx) < $x3dMin($idx)} {set x3dMin($idx) $x3dPoint($idx)}
                }
              }
            }
            append tessCoord($id) $tc
          }
        } elseif {$coordOnly != 2} {
          set msg "Error missing coordinates on \#$id=COORDINATES_LIST"
          errorMsg $msg
        }
        incr ncl

# tessellated_curve_set, *triangulated_surface_set, *triangulated_face
      } elseif {[string first "TESSELLATED_CURVE_SET" $line] != -1 || [string first "TRIANGULATED" $line] != -1} {
        set complex 0
        if {[string first "COMPLEX" $line] != -1} {set complex 1}

# id of coordinate list
        regsub -all " " $line "" line
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

# regsub is very important to distill line into something usable
          set c1 [string last "((" $line]
          if {$c1 != -1} {
            set tessIndex($id) ""
            set line [string range $line $c1 end-1]
            regsub -all {[(),]} $line "x" line
            regsub -all "xxx" $line " 0 " line
            regsub -all "x" $line " " line
            foreach j [split $line " "] {if {$j != ""} {append tessIndex($id) "[expr {$j-1}] "}}

# error reading line strips
          } else {
            if {[string first "$,$" $line] != -1} {
              set msg "Syntax Error: Missing 'coordinates' and 'line_strips' on tessellated_curve_set.$spaces\($recPracNames(pmi242), Sec. 8.2)"
              errorMsg $msg
              lappend syntaxErr(tessellated_curve_set) [list $id coordinates $msg]
              lappend syntaxErr(tessellated_curve_set) [list $id line_strips $msg]
            } else {
              errorMsg "Error reading line_strips on \#$id=TESSELLATED_CURVE_SET"
            }
          }

# *triangulated surface set, *triangulated face
        } else {
          set line [string range $line [string first "((" $line]+2 end-1]
          set line [string trim [string range $line [string first "((" $line] end]]

# regsub is very important to distill line into something usable
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
    set ok 0
    if {[info exists entCount(coordinates_list)]}    {if {$ncl == $entCount(coordinates_list)}    {set ok 1}}
    if {[info exists entCount(point_cloud_dataset)]} {if {$ncl == $entCount(point_cloud_dataset)} {set ok 1}}
    if {$ok} {
      if {($coordOnly == 0 && $ntc == $ntc1) || $coordOnly != 0} {
        if {$coordOnly == 0} {
          outputMsg "  [expr {$ncl+$ntc}] tessellated geometry entities"
          if {!$tessellated} {errorMsg " No tessellated curves, faces, or surfaces found."}
        }
        close $tg

# bounding box
        if {[info exists x3dMin(x)] && ($opt(tessPartOld) || $tessBrep)} {set x3dBbox 1}
        return
      }
    }
  }
}

# -------------------------------------------------------------------------------
proc tessSetColor {tessEnt faceEnt} {
  global defaultColor entCount recPracNames spaces tessColor x3dColor
  global objDesign

# color already exists for the tessellation
  set id(tess) [$tessEnt P21ID]
  set id(face) [$faceEnt P21ID]
  foreach idx {face tess} {
    if {[info exists tessColor($idx)]} {
      set x3dColor $tessColor($idx)
      return
    }
  }

# get styled_item.item for triangulated faces
  set tsID $id(face)
  set objEntity $faceEnt
  set e0s [$objEntity GetUsedIn [string trim styled_item] [string trim item]]

# get styled_item.item for tessellated_solid/shell if none for faces
  set n 0
  ::tcom::foreach e0 $e0s {incr n}
  if {$n == 0} {
    set tsID $id(tess)
    set objEntity $tessEnt
    set e0s [$objEntity GetUsedIn [string trim styled_item] [string trim item]]
  }

# get styled_item.item for tessellated_closed_shell (tessellated brep)
  set n 0
  ::tcom::foreach e0 $e0s {incr n}
  if {$n == 0} {
    if {[info exists entCount(manifold_solid_brep)] && $entCount(manifold_solid_brep) > 0} {
      ::tcom::foreach msb [$objDesign FindObjects [string trim "manifold_solid_brep"]] {
        set tcs [[[$msb Attributes] Item [expr 2]] Value]
        set tcsType [$tcs Type]
        if {$tcsType == "tessellated_closed_shell" || $tcsType == "tessellated_open_shell"} {
          if {[$tcs P21ID] == $id(tess)} {set e0s [$msb GetUsedIn [string trim styled_item] [string trim item]]}
        }
      }
    }
  }

# no styled_item for either
  set n 0
  ::tcom::foreach e0 $e0s {incr n}
  if {$n == 0} {
    errorMsg "Syntax Error: No 'styled_item' found for tessellated geometry$spaces\($recPracNames(model), Sec. 4.2.2, Fig. 2)"
    set tessColor($id(face)) $defaultColor
    set tessColor($id(tess)) $defaultColor
    return
  }

# get color from styled_item
  if {[catch {
    set debug 0
    ::tcom::foreach e0 $e0s {
      if {$debug} {errorMsg "[$e0 Type] [$e0 P21ID]" green}

# styled_item.styles
      set a1 [[$e0 Attributes] Item [expr 2]]
# presentation_style.styles
      ::tcom::foreach e2 [$a1 Value] {
        set a2 [[$e2 Attributes] Item [expr 1]]
        if {$debug} {errorMsg " [$e2 Type] [$e2 P21ID]" red}

        set e3 [$a2 Value]
        if {$e3 != "null" && $e3 != ""} {
          set a3 [[$e3 Attributes] Item [expr 2]]
          if {$debug} {errorMsg "  [$e3 Type] [$e3 P21ID]" red}
# surface side style
          set e4 [$a3 Value]
          set a4s [[$e4 Attributes] Item [expr 2]]
# surface style fill area
          foreach e5 [$a4s Value] {
            if {$debug} {errorMsg "   [$e5 Type] [$e5 P21ID]" red}
            if {[$e5 Type] == "surface_style_fill_area"} {
              set a5 [[$e5 Attributes] Item [expr 1]]
# fill area style
              set e6 [$a5 Value]
              set a6 [[$e6 Attributes] Item [expr 2]]
# fill area style colour
              set e7 [$a6 Value]
              set a7 [[$e7 Attributes] Item [expr 2]]
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
                set x3dColor [x3dPreDefinedColor [[[$e8 Attributes] Item [expr 1]] Value]]
                set tessColor($tsID) $x3dColor
              } else {
                errorMsg "  Tessellated part color type '[$e8 Type]' is not supported."
                set tessColor($tsID) $x3dColor
              }
            }
          }
        }
      }
    }

# color not found in styled_item.item
    if {![info exists tessColor($tsID)]} {errorMsg "Syntax Error: Problem setting tessellated geometry color$spaces\($recPracNames(model), Sec. 4.2.2, Fig. 2)"}

  } emsg]} {
    errorMsg " Error setting tessellated geometry color: $emsg"
    set tessColor($tsID) $defaultColor
  }
}

# -------------------------------------------------------------------------------
proc tessCountColors {} {
  global defaultColor entCount
  global objDesign

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
        set color [[[$e0 Attributes] Item [expr 1]] Value]
        if {[lsearch $colors $color] == -1} {lappend colors $color}
      }
    }
  } emsg]} {
    errorMsg " Error counting unique colors: $emsg"
  }

  if {[llength $colors] == 0 && ([info exists entCount(tessellated_solid)] || [info exists entCount(tessellated_shell)] || \
                                 [info exists entCount(tessellated_closed_shell)] || [info exists entCount(tessellated_open_shell)])} {
    lappend colors $defaultColor
  }
  return [llength $colors]
}

# -------------------------------------------------------------------------------
# this works for repositioned tessellated geometry used for PMI annotations, it does not work for tessellated part geometry in assemblies
proc tessSetPlacement {objEntity tsID} {
  global shapeRepName tessPlacement tessRepo

  set debug 0
  if {[catch {
    set tessRepo 0
    catch {unset tessPlacement}

# find tessellated_solid in tessellated_shape_representation.items
    ::tcom::foreach e0 [$objEntity GetUsedIn [string trim tessellated_shape_representation] [string trim items]] {
      if {$debug} {errorMsg "[$e0 Type] [$e0 P21ID]" green}

# find TSR in SRR rep_1 or rep_2
      foreach repSRR {rep_1 rep_2} {
        set e2s [$e0 GetUsedIn [string trim shape_representation_relationship] [string trim $repSRR]]

        ::tcom::foreach e2 $e2s {
          if {$debug} {errorMsg " [$e2 Type] [$e2 P21ID]" red}
          set a2 [[$e2 Attributes] Item [expr 4]]
          set e3 [$a2 Value]
          if {$debug} {errorMsg "  [$e3 Type] [$e3 P21ID] ([$a2 Name])" red}

          if {[$e3 Type] != "shape_representation"} {
            set a2 [[$e2 Attributes] Item [expr 3]]
            set e3 [$a2 Value]
            if {$debug} {errorMsg "  [$e3 Type] [$e3 P21ID] ([$a2 Name])" blue}
          }

          set shapeRepName [string trim [[[$e3 Attributes] Item [expr 1]] Value]]
          if {$shapeRepName != ""} {
            if {$debug} {errorMsg "  $shapeRepName" blue}
          } else {
            unset shapeRepName
          }

          set e4s [$e3 GetUsedIn [string trim shape_definition_representation] [string trim used_representation]]
          ::tcom::foreach e4 $e4s {
            set a4 [[$e4 Attributes] Item [expr 1]]
            set e5 [$a4 Value]
            if {$debug} {errorMsg "   [$e5 Type] [$e5 P21ID] ([$a4 Name])" green}
            set a5 [[$e5 Attributes] Item [expr 3]]
            set e6 [$a5 Value]
            if {$debug} {errorMsg "    [$e6 Type] [$e6 P21ID] ([$a5 Name])" green}
            foreach rel {relating_product_definition related_product_definition} {
              set e7s [$e6 GetUsedIn [string trim next_assembly_usage_occurrence] [string trim $rel]]
              ::tcom::foreach e7 $e7s {
                set a7 [[$e7 Attributes] Item [expr 5]]
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
              set a4 [[$e4 Attributes] Item [expr 5]]
              set e5 [$a4 Value]
              if {$debug} {errorMsg "    [$e5 Type] [$e5 P21ID] ([$a4 Name])" red}
              set a5 [[$e5 Attributes] Item [expr 4]]
              set e6 [$a5 Value]
              if {$debug} {errorMsg "     [$e6 Type] [$e6 P21ID] ([$a5 Name])" red}

# a2p3d
              set a2p3d [x3dGetA2P3D $e6]
              lappend tessPlacement(origin) [lindex $a2p3d 0]
              lappend tessPlacement(axis)   [lindex $a2p3d 1]
              lappend tessPlacement(refdir) [lindex $a2p3d 2]
              set tessRepo 1
            }
          }
        }
      }
    }
  } emsg]} {
    errorMsg " Error getting tessellated geometry placement: $emsg"
  }
}
