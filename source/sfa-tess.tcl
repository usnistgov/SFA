proc tessPart {entType} {
  global objDesign
  global ao entLevel ent entAttrList opt tessCoordID x3dFile entCount mytemp
  global tessEdgeFile tessEdgeFileName tessPartFile tessPartFileName tessSuppGeomFile tessSuppGeomFileName

  if {$opt(DEBUG1)} {outputMsg "START tessPart $entType" red}
  set msg " Adding Tessellated Part View"
  if {$opt(XLSCSV) == "None"} {append msg " ($entType)"}

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
   
# initialize
  set ao $entType
  set entAttrList {}
  set tessCoordID {}
  catch {unset ent}
  set entLevel 0
  setEntAttrList $PMIP($ao)
  if {$opt(DEBUG1)} {outputMsg "entattrlist $entAttrList"}
  
# open file for tessellated parts
  foreach tess {tessellated_solid tessellated_shell tessellated_wire} {
    if {[info exist entCount($tess)] && ![info exists tessPartFileName]} {
      if {$entCount($tess) > 0} {
        set tessPartFileName [file join $mytemp tessPart.txt]
        catch {file delete -force -- $tessPartFileName}
        set tessPartFile [open $tessPartFileName w]
      }
    }
  }
  
# open file for tessellated edges
  foreach edge {tessellated_edge tessellated_connecting_edge} {
    if {[info exist entCount($edge)] && ![info exists tessEdgeFileName]} {
      if {$entCount($edge) > 0} {
        set tessEdgeFileName [file join $mytemp tessEdge.txt]
        catch {file delete -force -- $tessEdgeFileName}
        set tessEdgeFile [open $tessEdgeFileName w]
      }
    }
  }
  
# open file for tessellated supplemental geometry, might not be used
  foreach wire {tessellated_wire tessellated_shell} {
    if {[info exist entCount($wire)] && ![info exists tessSuppGeomFileName]} {
      if {$entCount($wire) > 0} {
        set tessSuppGeomFileName [file join $mytemp tessSuppGeom.txt]
        catch {file delete -force -- $tessSuppGeomFileName}
        set tessSuppGeomFile [open $tessSuppGeomFileName w]
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
  global ao badAttributes entLevel ent entAttrList objEntity1 opt
  global tessCoord tessEdgeCoord tessIndex tessIndexCoord tessEdgeFile tessCoordID tessEdgeCoordDef tessSuppGeomFile tessEdge shellSuppGeom
  global x3dStartFile x3dFile entCount
  #outputMsg "tessPartGeometry [$objEntity Type] [$objEntity P21ID]" red

  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]

  if {[string first "handle" $objEntity] != -1} {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set objAttributes [$objEntity Attributes]
    set ent($entLevel) $objType
    if {$entLevel == 1} {
      set objEntity1 $objEntity
      
# check if tessellated wire is an item of constructive_geometry_representation (supplemental geometry)   
      catch {set tessEdge $tessEdgeFile}
      if {[$objEntity1 Type] == "tessellated_wire"} {
        set e0s [$objEntity1 GetUsedIn [string trim constructive_geometry_representation] [string trim items]]
        ::tcom::foreach e0 $e0s {
          ::tcom::foreach e1 [[[$e0 Attributes] Item [expr 2]] Value] {
            if {$objID == [$e1 P21ID]} {set tessEdge $tessSuppGeomFile}
          }
        }
      }
      
# check if tessellated shell is an item of constructive_geometry_representation (supplemental geometry)     
      set shellSuppGeom 0
      if {[$objEntity1 Type] == "tessellated_shell"} {
        set e0s [$objEntity1 GetUsedIn [string trim constructive_geometry_representation] [string trim items]]
        ::tcom::foreach e0 $e0s {set shellSuppGeom 1}
      }
    }
  
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
# tessellated edges
                set line_strip {}
                foreach item $objValue {lappend line_strip [expr {$item-1}]}
                puts $tessEdge "<Shape><Appearance><Material emissiveColor='0 0 0'></Material></Appearance>"
                if {![info exists tessEdgeCoordDef($tessEdgeCoord)]} {
                  lappend tessCoordID $tessEdgeCoord
                  set tessEdgeCoordDef($tessEdgeCoord) 1
                  puts $tessEdge " <IndexedLineSet coordIndex='$line_strip -1'>\n  <Coordinate DEF='coord$tessEdgeCoord' point='$tessCoord($tessEdgeCoord)'></Coordinate></IndexedLineSet></Shape>"
                } else {
                  puts $tessEdge " <IndexedLineSet coordIndex='$line_strip -1'><Coordinate USE='coord$tessEdgeCoord'></Coordinate></IndexedLineSet></Shape>"
                }
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
                      errorMsg "Missing tessellated coordinates and index for \#$objID"
                    }
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing Tessellated Part Geometry: $emsg3"
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
          }
        }
      }
    }
  }
  incr entLevel -1
}

# -------------------------------------------------------------------------------
proc tessReadGeometry {} {
  global entCount localName tessCoord tessIndex tessIndexCoord x3dMin x3dMax opt
  global coordinatesList lineStrips normals triangles
  
  set tg [open $localName r]
  set ncl  0
  set ntc  0
  set ntc1 0
  set ndup 0
  set tessellated 0
  
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
      if {[string first "TRIANGULATED" $line] != -1 || [string first "CURVE" $line] != -1} {set tessellated 1}      

# coordinates_list
      if {[string first "COORDINATES_LIST" $line] != -1} {
        if {$opt(PR_STEP_CPNT)} {regsub -all " " [string range $line [string first "((" $line]+1 end-3] "" coordinatesList($id)}
        
        set ncoord [string range $line [string first "," $line]+1 [string first "((" $line]-2]
        if {$ncoord > 50000} {errorMsg "COORDINATES_LIST #$id has $ncoord coordinates."}
        
        if {$ncoord > 0} {
          set nc 0
          set line [string range $line [string first "((" $line] end-1]

# regsub is very important to distill line into something usable
          regsub -all {[(),]} $line " " line
          set line [string trim $line]
          regsub -all "  " $line " " line
          regsub -all "  " $line " " line
          regsub -all "  " $line " " line
          regsub -all "  " $line " " line
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
              if {$x3dPoint($idx) > $x3dMax($idx)} {set x3dMax($idx) $x3dPoint($idx)}
              if {$x3dPoint($idx) < $x3dMin($idx)} {set x3dMin($idx) $x3dPoint($idx)}
            }
            
# check for one duplicate            
            incr nc
            if {$nc == 8} {if {[string first " $tc" $tessCoord($id)] != -1} {incr ndup}}
            append tessCoord($id) $tc
          }
        } else {
          set msg "COORDINATES_LIST #$id has no coordinates."
          errorMsg $msg
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
          if {$opt(PR_STEP_GEOM) || $opt(PMIGRF)} {regsub -all " " [string range $line [string first "((" $line]+1 end-3] "" lineStrips($id)}
          
# regsub is very important to distill line into something usable
          set c1 [string last "((" $line]
          if {$c1 != -1} {
            set line [string range $line $c1 end-1]
            regsub -all " " $line "" line
            regsub -all {[(),]} $line "x" line
            regsub -all "xxx" $line " 0 " line
            regsub -all "x" $line " " line
  
            set tessIndex($id) ""
            foreach j [split $line " "] {if {$j != ""} {append tessIndex($id) "[expr {$j-1}] "}}
          } else {
            errorMsg "ERROR reading 'line_strips' on \#$id=TESSELLATED_CURVE_SET"
          }

# *triangulated surface set, *triangulated face
        } else {
          #if {$opt(PR_STEP_GEOM)} {
          #  if {[string first "_FACE" $line] != -1} {
          #    set c1 [string first "))" $line]
          #    regsub -all " " [string range $line [string first "((" $line]+1 $c1] "" normals($id)
          #    set nline [string range $line $c1+2 end]
          #    regsub -all " " [string range $nline [string first "((" $nline]+1 [string first "))" $nline]] "" triangles($id)
          #  }
          #}
          
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
      if {$ndup > 0 && [expr {double($ndup)/double($ncl)}] > 0.4} {errorMsg "At least $ndup of $ncl COORDINATE_LIST have duplicate coordinates."}
      outputMsg "  [expr {$ncl+$ntc}] tessellated geometry entities"
      if {!$tessellated} {errorMsg " No tessellated curves, faces, or surfaces found."}
      close $tg
      return
    }
  }
}

# -------------------------------------------------------------------------------
proc tessSetColor {objEntity tsID} {
  global tessColor x3dColor entCount recPracNames defaultColor
  #outputMsg "tessSetColor [$objEntity Type] [$objEntity P21ID] ($tsID)" red

  set ok  0
  set ok1 1
  if {[info exists entCount(styled_item)]} {if {$entCount(styled_item) > 0} {set ok 1}}

# color already exists for the tessellation
  if {[info exists tessColor($tsID)]} {
    set x3dColor $tessColor($tsID)

# get color from styled_item
  } elseif {$ok} {
    if {[catch {
      set debug 0
      
# get styled_item.item for tessellated_solid/shell or triangulated faces
      ::tcom::foreach e0 [$objEntity GetUsedIn [string trim styled_item] [string trim item]] {
        if {$debug} {errorMsg "[$e0 Type] [$e0 P21ID]" green}
        set ok1 0
        
# styled_item.styles
        set a1 [[$e0 Attributes] Item [expr 2]]
# presentation_style.styles
        ::tcom::foreach e2 [$a1 Value] {
          set a2 [[$e2 Attributes] Item [expr 1]]
          if {$debug} {errorMsg " [$e2 Type] [$e2 P21ID]" red}

          set e3 [$a2 Value]
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
                errorMsg "  Unexpected color type ([$e8 Type]) for $ao"
                set tessColor($tsID) $x3dColor
              }
            }
          }
        }
      }
      
# tessellated_solid and _shell not found in styled_item.item, try the first item of TS, usually a face      
      if {$ok1} {
        if {[string first "tessellated" [$objEntity Type]] == 0} {set missingStyledItem [$objEntity Type]}
        set n 0
        catch {
          ::tcom::foreach e1 [[[$objEntity Attributes] Item [expr 2]] Value] {
            incr n
            if {$n == 1} {tessSetColor $e1 $tsID}
          }
        }
      }
    } emsg]} {
      errorMsg " ERROR setting Tessellated Geometry Color for [$objEntity Type] (using [lindex $defaultColor 1])\n $emsg"
      set tessColor($tsID) [lindex $defaultColor 0]
    }
  } else {
    errorMsg "Syntax Error: No 'styled_item' found for tessellated geometry color (using [lindex $defaultColor 1])\n[string repeat " " 14]($recPracNames(model), Sec. 4.2.2, Fig. 2)"
  }
  
# color not found in styled_item.item
  if {![info exists tessColor($tsID)]} {
    if {[info exists missingStyledItem]} {
      errorMsg " For tessellated geometry color, '$missingStyledItem' was not found in 'styled_item.item' (using [lindex $defaultColor 1])"
    }
  }
}

# -------------------------------------------------------------------------------
proc tessCountColors {} {
  global objDesign
  global entCount defaultColor

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
    errorMsg " ERROR counting unique colors: $emsg"
  }

  if {[llength $colors] == 0 && ([info exists entCount(tessellated_solid)] || [info exists entCount(tessellated_shell)])} {
    lappend colors [lindex $defaultColor 0]
  }
  return [llength $colors]
}

# -------------------------------------------------------------------------------
proc tessSetPlacement {objEntity tsID} {
  global objDesign
  global tessPlacement tessRepo shapeRepName

  set debug 0
  #outputMsg "[$objEntity Type] $tsID" blue
  
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
            foreach rel {relating_product_defintion related_product_definition} {
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
    errorMsg " ERROR getting tessellated geometry placement: $emsg"
  }
}

# -------------------------------------------------------------------------------
# generate A2P3D origin, axis, refdir
proc x3dGetA2P3D {e0} {

  set origin "0 0 0"
  set axis   "0 0 1"
  set refdir "1 0 0"
  set debug 0

# a2p3d origin
  set a2 [[$e0 Attributes] Item [expr 2]]
  set e2 [$a2 Value]
  if {$e2 != ""} {
    set origin [vectrim [[[$e2 Attributes] Item [expr 2]] Value]]
    if {$debug} {errorMsg "      [$e2 Type] [$e2 P21ID] ([$a2 Name]) $origin" red}
  }

# a2p3d axis
  set a3 [[$e0 Attributes] Item [expr 3]]
  set e3 [$a3 Value]
  if {$e3 != ""} {
    set axis [[[$e3 Attributes] Item [expr 2]] Value]
    if {$debug} {errorMsg "      [$e3 Type] [$e3 P21ID] ([$a3 Name]) $axis" red}
  }

# a2p3d reference direction
  set a4 [[$e0 Attributes] Item [expr 4]]
  set e4 [$a4 Value]
  if {$e4 != ""} {
    set refdir [[[$e4 Attributes] Item [expr 2]] Value]
    if {$debug} {errorMsg "      [$e4 Type] [$e4 P21ID] ([$a4 Name]) $refdir" red}
  }
  
  return [list $origin $axis $refdir]
}

# -------------------------------------------------------------------------------
# generate x3d rotation from axis2_placement_3d
proc x3dRotation {axis refdir} {
    
# construct rotation matrix u, must normalize to use with quaternion
  set u3 [vecnorm $axis]
  set u1 [vecnorm [vecsub $refdir [vecmult $u3 [vecdot $refdir $u3]]]]
  set u2 [vecnorm [veccross $u3 $u1]]

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

# convert from quaterion to x3d rotation
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

# mult - multiply vector by scalar
proc vecmult {v1 scalar} {
  foreach c1 $v1 {lappend v2 [expr {$c1*$scalar}]}
  return $v2
}

# sub - subtract one vector from another
proc vecsub {v1 v2} {
  foreach c1 $v1 c2 $v2 {lappend v3 [expr {$c1-$c2}]}
  return $v3
}

# add - add one vector to another
proc vecadd {v1 v2} {
  foreach c1 $v1 c2 $v2 {lappend v3 [expr {$c1+$c2}]}
  return $v3
}

# reverse - reverse vector direction
proc vecrev {v1} {
  foreach c1 $v1 {
    if {$c1 != 0.} {
      lappend v2 [expr {$c1*-1.}]
    } else {
      lappend v2 $c1
    }
  }
  return $v2
}

# trim - truncate values in a vector
proc vectrim {v1} {
  foreach c1 $v1 {
    set prec 3
    if {[expr {abs($c1)}] < 0.01} {set prec 4}
    lappend v2 [trimNum $c1 $prec]
  }
  return $v2
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
