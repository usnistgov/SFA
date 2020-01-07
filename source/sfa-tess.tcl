proc tessPart {entType} {
  global objDesign
  global ao ent entAttrList entCount entLevel mytemp opt tessCoordID
  global tessPartFile tessPartFileName tessSuppGeomFile tessSuppGeomFileName

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
  checkTempDir
  foreach tess {tessellated_solid tessellated_shell tessellated_wire} {
    if {[info exist entCount($tess)] && ![info exists tessPartFileName]} {
      if {$entCount($tess) > 0} {
        set tessPartFileName [file join $mytemp tessPart.txt]
        catch {file delete -force -- $tessPartFileName}
        set tessPartFile [open $tessPartFileName w]
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
  global ao badAttributes ent entAttrList entLevel objEntity1 opt shellSuppGeom tessCoord
  global tessEdgeCoord tessEdges tessIndex tessIndexCoord x3dFile x3dStartFile

  if {$opt(DEBUG1)} {outputMsg "tessPartGeometry [$objEntity Type] [$objEntity P21ID]" red}
  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]

  if {[string first "handle" $objEntity] != -1} {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set objAttributes [$objEntity Attributes]
    set ent($entLevel) $objType
    if {$entLevel == 1} {
      set objEntity1 $objEntity
      
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
# generate coordinates and faces by reading tessellated geometry entities line-by-line from STEP file, necessary because of limitations of the IFCsvr toolkit
# if coordOnly=1, then only read coordinates_list, used for saved view pmi validation properties
proc tessReadGeometry {{coordOnly 0}} {
  global coordinatesList entCount lineStrips localName normals opt tessCoord tessCoordName tessIndex tessIndexCoord triangles x3dMax x3dMin
  
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
    set ents [list COORDINATES_LIST TESSELLATED_CURVE_SET TRIANGULATED_FACE COMPLEX_TRIANGULATED_FACE TRIANGULATED_SURFACE_SET COMPLEX_TRIANGULATED_SURFACE_SET]
    outputMsg " Reading tessellated geometry" green

# read only coordinates list
  } else {
    foreach ent {coordinates_list} {if {[info exists entCount($ent)]} {set ntc1 [expr {$ntc1+$entCount($ent)}]}}
    set ents [list COORDINATES_LIST]
    outputMsg " Reading coordinates_list" green
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

# coordinates_list
      if {[string first "COORDINATES_LIST" $line] != -1} {
        set tessCoordName($id) ""
        set c1 [string first "'" $line]
        set c2 [string last  "'" $line]
        if {$c2 != [expr {$c1+1}]} {set tessCoordName($id) [string range $line $c1+1 $c2-1]} 
        
        if {$opt(PR_STEP_CPNT)} {regsub -all " " [string range $line [string first "((" $line]+1 end-3] "" coordinatesList($id)}
        
        set ncoord [string range $line [string first "," $line]+1 [string first "((" $line]-2]
        if {$ncoord > 50000} {errorMsg "COORDINATES_LIST #$id has $ncoord coordinates."}
        
        if {$ncoord > 0} {
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
              if {$coordOnly == 0} {
                if {$x3dPoint($idx) > $x3dMax($idx)} {set x3dMax($idx) $x3dPoint($idx)}
                if {$x3dPoint($idx) < $x3dMin($idx)} {set x3dMin($idx) $x3dPoint($idx)}
              }
            }
            append tessCoord($id) $tc
          }
        } else {
          set msg "ERROR missing coordinates on \#$id=COORDINATES_LIST"
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
            set tessIndex($id) ""
            set line [string range $line $c1 end-1]
            regsub -all " " $line "" line
            regsub -all {[(),]} $line "x" line
            regsub -all "xxx" $line " 0 " line
            regsub -all "x" $line " " line
            foreach j [split $line " "] {if {$j != ""} {append tessIndex($id) "[expr {$j-1}] "}}

# error reading line strips
          } else {
            set str "reading"
            if {[string first "$,$" $line] != -1} {set str "missing"}
            errorMsg "ERROR $str line_strips on \#$id=TESSELLATED_CURVE_SET"
          }

# *triangulated surface set, *triangulated face
        } else {

# try to read normals and triangles for spreadsheet
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
    if {$ncl == $entCount(coordinates_list)} {
      if {($coordOnly == 0 && $ntc == $ntc1) || $coordOnly} {
        if {$coordOnly == 0} {
          outputMsg "  [expr {$ncl+$ntc}] tessellated geometry entities"
          if {!$tessellated} {errorMsg " No tessellated curves, faces, or surfaces found."}
        }
        close $tg
        return
      }
    }
  }
}

# -------------------------------------------------------------------------------
proc tessSetColor {objEntity tsID} {
  global defaultColor entCount recPracNames tessColor x3dColor
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
      errorMsg "Syntax Error: '$missingStyledItem' was not found in 'styled_item.item' (using [lindex $defaultColor 1])\n[string repeat " " 14]($recPracNames(model), Sec. 4.2.2, Fig. 2)"
    }
  }
}

# -------------------------------------------------------------------------------
proc tessCountColors {} {
  global objDesign
  global defaultColor entCount

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
# this works for repositioned tessellated geometry used for PMI annotations, it does not work for tessellated part geometry in assemblies
proc tessSetPlacement {objEntity tsID} {
  global shapeRepName tessPlacement tessRepo

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
