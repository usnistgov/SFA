proc feaStart {objDesign entType} {
  global ent entAttrList entCount entLevel opt rowmax nprogEnts count
  global x3domCoord x3domFile x3domIndex x3domIndex1 x3domFile x3domMin x3domMax x3domMsg
  global feaType feaTypes feaElemTypes feaNodes feaFaces x3dwire x3delements

  if {$opt(DEBUG1)} {outputMsg "START feaStart $entType\n" red}

# finite elements
  set cartesian_point [list cartesian_point coordinates]
  set node            [list node items $cartesian_point]
  set dummy_node      [list dummy_node name items]
  set curve_3d_element_descriptor   [list curve_3d_element_descriptor topology_order description]
  set surface_3d_element_descriptor [list surface_3d_element_descriptor topology_order description Shape]
  set volume_3d_element_descriptor  [list volume_3d_element_descriptor topology_order description purpose Shape]
  
  set FEA(curve_3d_element_representation)   [list curve_3d_element_representation   name node_list $node element_descriptor $curve_3d_element_descriptor]
  set FEA(surface_3d_element_representation) [list surface_3d_element_representation name node_list $node $dummy_node element_descriptor $surface_3d_element_descriptor]
  set FEA(volume_3d_element_representation)  [list volume_3d_element_representation  name node_list $node $dummy_node element_descriptor $volume_3d_element_descriptor]

  if {[info exists ent]} {unset ent}
  set entLevel 0
  set entAttrList {}
  setEntAttrList $FEA($entType)
  
# check number of elements to see if faces will be displayed
  set feaFaces(2D) 1
  if {$entType == "surface_3d_element_representation"} {
    if {$entCount(surface_3d_element_representation) > 2000000} {
      set feaFaces(2D) 0
      errorMsg "For surface_3d_element_representation, too many elements ($entCount(surface_3d_element_representation)) to display faces"
    }
  }
  set feaFaces(3D) 1
  if {$entType == "volume_3d_element_representation"} {
    if {$entCount(volume_3d_element_representation) > 200000} {
      set feaFaces(3D) 0
      errorMsg "For volume_3d_element_representation, too many elements ($entCount(volume_3d_element_representation)) to display faces"
    }
  }
  catch {unset feaTypes}

# process all *_element_representation entities, call feaReport
  set n 0
  set startent [lindex $FEA($entType) 0]
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    if {[$objEntity Type] == $startent} {
      if {$n < 10000000} {
        if {[expr {$n%2000}] == 0} {
          if {$n > 0} {outputMsg "  $n"}
          update idletasks
        }
        if {$n > $rowmax} {
         incr nprogEnts
         if {[expr {$nprogEnts%1000}] == 0} {update idletasks}
        }
        feaReport $objEntity
        if {$opt(DEBUG1)} {outputMsg \n}
      }
      incr n
    }
  }

# done processing entities, write fem  
# coordinate min, max for node and axes size
  foreach idx {x y z} {
    set delt($idx) [expr {$x3domMax($idx)-$x3domMin($idx)}]
    set xyzcen($idx) [format "%.4f" [expr {0.5*$delt($idx) + $x3domMin($idx)}]]
  }

  set maxxyz $delt(x)
  if {$delt(y) > $maxxyz} {set maxxyz $delt(y)}
  if {$delt(z) > $maxxyz} {set maxxyz $delt(z)}

# save nodes
  if {[llength $x3domCoord] > 4000 && $opt(feaNodeType) == 2} {
    errorMsg "Too many Nodes to display as Cubes ([llength $x3domCoord]) for '$feaType\_element_representation'\n Switching to Nodes displayed as Points"
    set opt(feaNodeType) 1
  }
  foreach coord $x3domCoord {lappend feaNodes $coord}

# 1D elements
  if {$feaType == "curve_3d"} {
    append x3delements "\n<Switch whichChoice='0' id='sw1D'>"
    append x3delements "<Shape>\n <Appearance><Material emissiveColor='1 0 1'></Material></Appearance>"
    append x3delements " <IndexedLineSet coordIndex='$x3domIndex'>"
    foreach coord $x3domCoord {append coords "$coord "}
    append x3delements "  <Coordinate point='$coords'></Coordinate></IndexedLineSet>\n</Shape>"
    append x3delements "</Switch>"
    unset x3domIndex    

# 2D, 3D elements
  } else {

# wireframe
    append x3dwire "<Shape>\n <Appearance><Material emissiveColor='0 0 0'></Material></Appearance>"
    append x3dwire " <IndexedLineSet coordIndex='$x3domIndex'>"
    foreach coord $x3domCoord {append coords "$coord "}
    append x3dwire "  <Coordinate point='$coords'></Coordinate></IndexedLineSet>\n</Shape>"
    unset x3domIndex

# faces
    if {[info exists x3domIndex1]} {
      if {$feaType == "surface_3d"} {
        append x3delements "\n<Switch whichChoice='0' id='sw2D'><Group>"
        set dc "0 1 1"
        set id "mat2D"
      } else {
        append x3delements "\n<Switch whichChoice='0' id='sw3D'><Group>"      
        set dc "1 1 0"
        set id "mat3D"
      }
      append x3delements "<Shape>\n <Appearance><Material id='$id' diffuseColor='$dc'></Material></Appearance>"
      append x3delements " <IndexedFaceSet solid='FALSE' coordIndex='$x3domIndex1'>"
      foreach coord $x3domCoord {append coords "$coord "}
      append x3delements "  <Coordinate point='$coords'></Coordinate></IndexedFaceSet>\n</Shape>"
      unset x3domIndex1
      append x3delements "</Group></Switch>"
    }
  }

# check if done processing all element types
  set writeX3DOM 0
  if {$feaType == "volume_3d"} {set writeX3DOM 1}
  if {$feaType == "surface_3d" && ![info exists entCount(volume_3d_element_representation)]} {set writeX3DOM 1}
  if {$feaType == "curve_3d" && ![info exists entCount(surface_3d_element_representation)] && \
                                ![info exists entCount(volume_3d_element_representation)]} {set writeX3DOM 1}

# write wireframe, nodes, axes at the end of processing all element types
  if {$writeX3DOM} {

# coordinate axes    
    set asize [trimNum [expr {$maxxyz/30.}]]
    puts $x3domFile "\n<Shape><Appearance><Material emissiveColor='1 0 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. $asize 0. 0.'></Coordinate></IndexedLineSet></Shape>"
    puts $x3domFile "<Shape><Appearance><Material emissiveColor='0 1 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. $asize 0.'></Coordinate></IndexedLineSet></Shape>"
    puts $x3domFile "<Shape><Appearance><Material emissiveColor='0 0 1'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. 0. $asize'></Coordinate></IndexedLineSet></Shape>"

# nodes
    if {[info exists feaNodes]} {
      set nc 1
      set gsize [expr {$maxxyz/500.}]
      puts $x3domFile "<Switch whichChoice='0' id='swND'><Group>"
      if {$opt(feaNodeType) == 2} {
        foreach coord $feaNodes {
          if {$nc != 1} {
            puts $x3domFile "<Transform translation='$coord'><Shape USE='node'></Shape></Transform>"
          } else {
            puts $x3domFile "<Transform translation='$coord'>\n <Shape DEF='node'>\n  <Appearance><Material diffuseColor='0 1 0'></Material></Appearance>\n  <Box size='$gsize $gsize $gsize'></Box>\n </Shape>\n</Transform>"
            #puts $x3domFile "<Transform translation='$coord'>\n <Shape DEF='sphere'>\n  <Appearance><Material diffuseColor='0 1 0'></Material></Appearance>\n  <Sphere radius='$gsize'></Sphere>\n </Shape>\n</Transform>"
          }
          incr nc
        }
      } elseif {$opt(feaNodeType) == 1} {
        puts $x3domFile "<Shape>\n <Appearance><Material emissiveColor='0 0.7 0'></Material></Appearance>"
        puts $x3domFile " <PointSet>"
        foreach coord $x3domCoord {append coords "$coord "}
        puts $x3domFile "  <Coordinate point='$coords'></Coordinate></PointSet>\n</Shape>"
      }
      puts $x3domFile "</Group></Switch>"
    }
    
# write wireframe
    if {[info exists x3dwire]} {
      puts $x3domFile "<Switch whichChoice='0' id='swWF'><Group>"
      puts $x3domFile $x3dwire
      puts $x3domFile "</Group></Switch>"
      unset x3dwire
    }
    
# write elements
    if {[info exists x3delements]} {
      if {$x3delements != ""} {
        puts $x3domFile $x3delements
        unset x3delements
      }
    }

    lappend x3domMsg "$entCount(node) - Nodes"
    catch {unset feaNodes}
  }  
  unset x3domCoord
  
  if {[info exists feaTypes]} {
    foreach item [array names feaTypes] {lappend x3domMsg "$feaTypes($item) - $item"}
  }
}

# -------------------------------------------------------------------------------

proc feaReport {objEntity} {
  global badAttributes ent entAttrList entCount entLevel localName nistVersion opt
  global x3domCoord x3domFile x3domFileName x3domFileOpen x3domIndex x3domIndex1 x3domMax x3domMin x3domPoint
  global idx feaIndex feaType feaTypes firstID nnode nnodes feaFaces

# entLevel is very important, keeps track level of entity in hierarchy
  #outputMsg [$objEntity Type] red
  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]

  if {[string first "handle" $objEntity] != -1} {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set objAttributes [$objEntity Attributes]
    set ent($entLevel) $objType

    if {$opt(DEBUG1)} {outputMsg "$ind ENT $entLevel #$objID=$objType (ATR=[$objAttributes Count])" blue}
    
    ::tcom::foreach objAttribute $objAttributes {
      set objName  [$objAttribute Name]
      set ent1 "$ent($entLevel) $objName"
      set ent2 "$ent($entLevel).$objName"

# look for entities with bad attributes that cause a crash
      set okattr 1
      if {[info exists badAttributes($objType)]} {foreach ba $badAttributes($objType) {if {$ba == $objName} {set okattr 0}}}

      if {$okattr} {
        set objValue    [$objAttribute Value]
        set objNodeType [$objAttribute NodeType]
        set objSize     [$objAttribute Size]
        set objAttrType [$objAttribute Type]
        set entAttrOK   [lsearch $entAttrList $ent1]

# -----------------
# nodeType = 18,19
        if {$objNodeType == 18 || $objNodeType == 19} {
          if {[catch {
            if {$entAttrOK != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}

# referenced entities
              if {[string first "handle" $objEntity] != -1} {feaReport $objValue}
            }
          } emsg3]} {
            errorMsg "ERROR processing Analysis Model ($objNodeType $ent2): $emsg3"
          }

# --------------
# nodeType = 20
        } elseif {$objNodeType == 20} {
          if {[catch {
            if {$entAttrOK != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}
    
              switch -glob $ent1 {
                "*_element_representation node_list" {set nnodes $objSize}
                "dummy_node items" -
                "cartesian_point coordinates" {
# node index
                  if {$objType == "cartesian_point"} {
                    if {[info exists x3domCoord]} {
                      set id [lsearch $x3domCoord $objValue]
                      if {$id == -1} {
                        lappend x3domCoord $objValue
                        set id [expr {[llength $x3domCoord]-1}]
                      }
                    } else {
                      lappend x3domCoord $objValue
                      set id 0
                    }
                    set idx($nnode) $id
                  }

# element connectivity based on part 104 ordering               
                  incr nnode
                  if {$nnode == 1} {set firstID $id}
  
# curve_3d
                  if {$nnode == $nnodes} {
                    if {$feaType == "curve_3d"} {
                      if {$nnodes == 3} {
                        foreach id {1 3 2} {append x3domIndex "$idx([expr {$id-1}]) "}
                      } elseif {$nnodes == 4} {
                        foreach id {1 3 4 2} {append x3domIndex "$idx([expr {$id-1}]) "}
                      } else {
                        foreach id [lsort [array names idx]] {append x3domIndex "$idx($id) "}
                        if {$nnodes > 4} {errorMsg "Unexpect number of nodes ($nnodes) for a $feaType element"}
                      }
                      append x3domIndex "-1 "
                      unset idx

# surface_3d, volume_3d (feaIndex)
                    } elseif {$feaType == "surface_3d" || $feaType == "volume_3d"} {
                      if {[info exists feaIndex($feaType,$nnodes,line)]} {
                        foreach id $feaIndex($feaType,$nnodes,line) {
                          if {$id != -1} {append x3domIndex "$idx([expr {$id-1}]) "} else {append x3domIndex "$id "}
                        }
                        if {($feaFaces(2D) && $feaType == "surface_3d") || ($feaFaces(3D) && $feaType == "volume_3d")} {
                          foreach id $feaIndex($feaType,$nnodes,surf) {
                            if {$id != -1} {append x3domIndex1 "$idx([expr {$id-1}]) "} else {append x3domIndex1 "$id "}
                          }
                        }
                      } else {
                        foreach id [lsort [array names idx]] {append x3domIndex "$idx($id) "}
                        set x3domIndex1 $x3domIndex
                        append x3domIndex  "$firstID -1 "
                        append x3domIndex1 "-1 "
                        if {$nnodes > 4} {errorMsg "Unexpect number of nodes ($nnodes) for a $feaType element, cubic elements are not supported"}
                      }
                      unset idx

                    } else {
                      errorMsg "Unexpect element type $feaType"
                    }
                  }

# min,max of points
                  if {$objType == "cartesian_point"} {
                    set x3domPoint(x) [lindex $objValue 0]
                    set x3domPoint(y) [lindex $objValue 1]
                    set x3domPoint(z) [lindex $objValue 2]
                    foreach xyz {x y z} {
                      if {$x3domPoint($xyz) > $x3domMax($xyz)} {set x3domMax($xyz) $x3domPoint($xyz)}
                      if {$x3domPoint($xyz) < $x3domMin($xyz)} {set x3domMin($xyz) $x3domPoint($xyz)}
                    }
                  }
                }
              }

# referenced entities
              if {[catch {
                ::tcom::foreach val1 $objValue {feaReport $val1}
              } emsg]} {
                foreach val2 $objValue {feaReport $val2}
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing Analysis Model ($objNodeType $ent2): $emsg3"
          }

# ---------------------
# nodeType = 5 (!= 18,19,20)
        } else {
          if {[catch {
            if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objAttrType)  ($ent1)"}
            if {$entAttrOK != -1} {
              switch -glob $ent1 {
                "*_element_representation name" {
                  set feaType [string range $ent1 0 [string first "_element" $ent1]-1]
                  set nnode 0
# start X3DOM file                
                  if {$x3domFileOpen} {
                    set x3domFileOpen 0
                    set x3domFileName [file rootname $localName]_x3dom.html
                    catch {file delete -force $x3domFileName}
                    set x3domFile [open $x3domFileName w]
                    outputMsg " Writing Analysis Model to: [truncFileName [file nativename $x3domFileName]]" green
                    puts $x3domFile "<!DOCTYPE html>\n<html>\n<head>\n<title>[file tail $localName] | AP209 Analysis Model</title>\n<base target=\"_blank\">\n<meta http-equiv='Content-Type' content='text/html;charset=utf-8'/>\n<link rel='stylesheet' type='text/css' href='https://www.x3dom.org/x3dom/release/x3dom.css'/>\n<script type='text/javascript' src='https://www.x3dom.org/x3dom/release/x3dom.js'></script>"

# node, element checkbox script
                    x3domScript ND
                    if {[info exists entCount(surface_3d_element_representation)] || \
                        [info exists entCount(volume_3d_element_representation)]}  {x3domScript WF}
                    if {[info exists entCount(curve_3d_element_representation)]}   {x3domScript 1D}
                    if {[info exists entCount(surface_3d_element_representation)]} {x3domScript 2D}
                    if {[info exists entCount(volume_3d_element_representation)]}  {x3domScript 3D}

# transparency script
                    if {[info exists entCount(surface_3d_element_representation)] || [info exists entCount(volume_3d_element_representation)]} {
                      puts $x3domFile "<script>function matTrans(trans){"
                      if {[info exists entCount(surface_3d_element_representation)]} {puts $x3domFile " document.getElementById('mat2D').setAttribute('transparency', trans);"}
                      if {[info exists entCount(volume_3d_element_representation)]}  {puts $x3domFile " document.getElementById('mat3D').setAttribute('transparency', trans);"}
                      puts $x3domFile "}\n</script>"
                    }
                    puts $x3domFile "</head>"

                    puts $x3domFile "\n<body><font face=\"arial\">\n<h3>AP209 Analysis Model:  [file tail $localName]</h3>"
                    puts $x3domFile "<ul><li><a href=\"https://www.x3dom.org/documentation/interaction/\">Use the mouse</a> to rotate, pan, and zoom.  Use Page Down to switch between perspective and orthographic views."
                    puts $x3domFile "<li>The viewer is experimental and not optimized for large models."
                    puts $x3domFile "</ul>"

# node, element checkboxes
                    puts $x3domFile "<input type='checkbox' checked onclick='togND(this.value)'/>Nodes&nbsp;&nbsp;"
                    if {[info exists entCount(surface_3d_element_representation)] || \
                        [info exists entCount(volume_3d_element_representation)]}  {puts $x3domFile "<input type='checkbox' checked onclick='togWF(this.value)'/>Wireframe&nbsp;&nbsp;"}
                    if {[info exists entCount(curve_3d_element_representation)]}   {puts $x3domFile "<input type='checkbox' checked onclick='tog1D(this.value)'/>Curve_3D&nbsp;&nbsp;"}
                    if {[info exists entCount(surface_3d_element_representation)]} {puts $x3domFile "<input type='checkbox' checked onclick='tog2D(this.value)'/>Surface_3D&nbsp;&nbsp;"}
                    if {[info exists entCount(volume_3d_element_representation)]}  {puts $x3domFile "<input type='checkbox' checked onclick='tog3D(this.value)'/>Volume_3D&nbsp;&nbsp;"}

# transparency slider
                    if {[info exists entCount(surface_3d_element_representation)] || [info exists entCount(volume_3d_element_representation)]} {
                      puts $x3domFile "<input style='width:80px' type='range' min='0' max='1' step='0.25' value='0' onchange='matTrans(this.value)'/>Transparency (might not appear correct)"
                    }
                    puts $x3domFile "<table><tr><td>"
                    
# x3d window size
                    set height 800
                    set width [expr {int($height*1.5)}]
                    catch {
                      set height [expr {int([winfo screenheight .]*0.7)}]
                      set width [expr {int($height*[winfo screenwidth .]/[winfo screenheight .])}]
                    }
                    puts $x3domFile "\n<X3D id='someUniqueId' showStat='false' showLog='false' x='0px' y='0px' width='$width\px' height='$height\px'>\n<Scene DEF='scene'>"
                  }
                }
                "*_element_descriptor description" {
# can also use with topology_order, Shape                
                  incr feaTypes($objValue)
                }
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing Analysis Model ($objNodeType $ent2): $emsg3"
            set entLevel 1
          }
        }
      }
    }
  }
  incr entLevel -1
}

# -------------------------------------------------------------------------------
# script for switch node

proc x3domScript {type} {
  global x3domFile
  
  puts $x3domFile "<script>\nfunction tog$type\(choice){"
  puts $x3domFile " if (!document.getElementById('sw$type').checked) {\n  document.getElementById('sw$type').setAttribute('whichChoice', -1);\n } else {\n  document.getElementById('sw$type').setAttribute('whichChoice', 0);\n }"
  puts $x3domFile " document.getElementById('sw$type').checked = !document.getElementById('sw$type').checked;\n}\n</script>"
}