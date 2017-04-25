proc feaModel {objDesign entType} {
  global ent entAttrList entCount entLevel opt rowmax nprogEnts count localName mytemp sfaPID
  global x3domFile x3domMin x3domMax x3domMsg x3domStartFile x3domFileName
  global feaType feaTypes feaElemTypes feaFaces nodeArr nfeaElem feaFile feaFileName feaFaceList feaFaceOrig

  if {$opt(DEBUG1)} {outputMsg "START feaModel $entType\n" red}

# finite elements
  set node            [list node name]
  set dummy_node      [list dummy_node name]
  set curve_3d_element_descriptor   [list curve_3d_element_descriptor description]
  set surface_3d_element_descriptor [list surface_3d_element_descriptor description]
  set volume_3d_element_descriptor  [list volume_3d_element_descriptor description]
  
  set FEA(curve_3d_element_representation)   [list curve_3d_element_representation   name node_list $node $dummy_node element_descriptor $curve_3d_element_descriptor]
  set FEA(surface_3d_element_representation) [list surface_3d_element_representation name node_list $node $dummy_node element_descriptor $surface_3d_element_descriptor]
  set FEA(volume_3d_element_representation)  [list volume_3d_element_representation  name node_list $node $dummy_node element_descriptor $volume_3d_element_descriptor]

  if {[info exists ent]} {unset ent}
  set entLevel 0
  set entAttrList {}
  setEntAttrList $FEA($entType)
  
# check number of elements to see if faces will be displayed
  set feaFaces(2D) 1
  set feaFaces(3D) 1

  #if {$entType == "surface_3d_element_representation"} {
  #  if {$entCount(surface_3d_element_representation) > 2000000} {
  #    set feaFaces(2D) 0
  #    errorMsg "Too many 'surface_3d_element_representation' elements ($entCount(surface_3d_element_representation)) to display faces"
  #  }
  #}
  #if {$entType == "volume_3d_element_representation"} {
  #  if {$entCount(volume_3d_element_representation) > 400000} {
  #    set feaFaces(3D) 0
  #    errorMsg "Too many 'volume_3d_element_representation' elements ($entCount(volume_3d_element_representation)) to display faces"
  #  }
  #}
  catch {unset feaTypes}

# ---------- 
# start X3DOM file                
  if {$x3domStartFile} {
    set x3domStartFile 0
    set x3domFileName [file rootname $localName]_x3dom.html
    catch {file delete -force $x3domFileName}
    set x3domFile [open $x3domFileName w]
    puts $x3domFile "<!DOCTYPE html>\n<html>\n<head>\n<title>[file tail $localName] | STEP AP209 Finite Element Model</title>\n<base target=\"_blank\">\n<meta http-equiv='Content-Type' content='text/html;charset=utf-8'/>\n<link rel='stylesheet' type='text/css' href='https://www.x3dom.org/x3dom/release/x3dom.css'/>\n<script type='text/javascript' src='https://www.x3dom.org/x3dom/release/x3dom.js'></script>"

# node, element checkbox script
    feaSwitch Nodes
    if {[info exists entCount(surface_3d_element_representation)] || \
        [info exists entCount(volume_3d_element_representation)]}  {feaSwitch Mesh}
    if {[info exists entCount(curve_3d_element_representation)]}   {feaSwitch 1DElements}
    if {[info exists entCount(surface_3d_element_representation)]} {feaSwitch 2DElements}
    if {[info exists entCount(volume_3d_element_representation)]}  {feaSwitch 3DElements}

# transparency script
    if {[info exists entCount(surface_3d_element_representation)] || [info exists entCount(volume_3d_element_representation)]} {
      puts $x3domFile "<script>function matTrans(trans){"
      if {[info exists entCount(surface_3d_element_representation)]} {puts $x3domFile " document.getElementById('mat2D').setAttribute('transparency', trans);"}
      if {[info exists entCount(volume_3d_element_representation)]}  {
        puts $x3domFile " document.getElementById('mat3D').setAttribute('transparency', trans);"
        puts $x3domFile " if (trans > 0) {document.getElementById('faces').setAttribute('solid', true);} else {document.getElementById('faces').setAttribute('solid', false);}"
      }
      puts $x3domFile "}\n</script>"
    }
    puts $x3domFile "</head>"

    puts $x3domFile "\n<body><font face=\"arial\">\n<h3>STEP AP209 Finite Element Model:  [file tail $localName]</h3>"
    puts $x3domFile "<ul><li><a href=\"https://www.x3dom.org/documentation/interaction/\">Use the mouse</a> to rotate, pan, and zoom.  Use Page Down to switch between perspective and orthographic views."
    puts $x3domFile "</ul>"

# node, element checkboxes
    puts $x3domFile "<input type='checkbox' checked onclick='togNodes(this.value)'/>Nodes&nbsp;&nbsp;"
    if {[info exists entCount(surface_3d_element_representation)] || \
        [info exists entCount(volume_3d_element_representation)]}  {puts $x3domFile "<input type='checkbox' checked onclick='togMesh(this.value)'/>Mesh&nbsp;&nbsp;"}
    if {[info exists entCount(curve_3d_element_representation)]}   {puts $x3domFile "<input type='checkbox' checked onclick='tog1DElements(this.value)'/>1D Elements&nbsp;&nbsp;"}
    if {[info exists entCount(surface_3d_element_representation)]} {puts $x3domFile "<input type='checkbox' checked onclick='tog2DElements(this.value)'/>2D Elements&nbsp;&nbsp;"}
    if {[info exists entCount(volume_3d_element_representation)]}  {puts $x3domFile "<input type='checkbox' checked onclick='tog3DElements(this.value)'/>3D Elements&nbsp;&nbsp;"}

# transparency slider
    if {[info exists entCount(surface_3d_element_representation)] || [info exists entCount(volume_3d_element_representation)]} {
      puts $x3domFile "<input style='width:80px' type='range' min='0' max='0.8' step='0.2' value='0' onchange='matTrans(this.value)'/>Transparency (might not appear correct)"
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

# nodes    
    feaGetNodes $objDesign
    outputMsg " Writing FEM to: [truncFileName [file nativename $x3domFileName]]" blue

# coordinate axes    
    foreach xyz {x y z} {
      set delt($xyz) [expr {$x3domMax($xyz)-$x3domMin($xyz)}]
      set xyzcen($xyz) [format "%.4f" [expr {0.5*$delt($xyz) + $x3domMin($xyz)}]]
    }
    set maxxyz $delt(x)
    if {$delt(y) > $maxxyz} {set maxxyz $delt(y)}
    if {$delt(z) > $maxxyz} {set maxxyz $delt(z)}
    set asize [trimNum [expr {$maxxyz/30.}]]

    puts $x3domFile "\n<Shape><Appearance><Material emissiveColor='1 0 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. $asize 0. 0.'></Coordinate></IndexedLineSet></Shape>"
    puts $x3domFile "<Shape><Appearance><Material emissiveColor='0 1 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. $asize 0.'></Coordinate></IndexedLineSet></Shape>"
    puts $x3domFile "<Shape><Appearance><Material emissiveColor='0 0 1'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. 0. $asize'></Coordinate></IndexedLineSet></Shape>"
    update idletasks
  }

# create temp files so that large models can be processed
  foreach f {elements mesh meshIndex faceIndex} {
    set feaFileName($f) [file join $mytemp $f.txt]
    if {![file exists $feaFileName($f)]} {set feaFile($f) [open $feaFileName($f) w]}
  }

# ----------  
# process all *_element_representation entities, call feaElements  
  set nfeaElem 0
  catch {unset feaFaceList}
  catch {unset feaFaceOrig}
  
  set startent [lindex $FEA($entType) 0]
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    if {[$objEntity Type] == $startent} {
      if {$nfeaElem < 10000000} {
        if {[expr {$nfeaElem%2000}] == 0} {
          if {$nfeaElem > 0} {outputMsg "  $nfeaElem"}
          set mem [expr {[lindex [twapi::get_process_info $sfaPID -pagefilebytes] 1]/1048576}]
          if {$mem > 1780} {
            errorMsg "Insufficient memory to process all of the elements"
            lappend x3domMsg "Some elements are not displayed"
            update idletasks
            break
          }
          update idletasks
        }
        if {$nfeaElem > $rowmax} {incr nprogEnts}
        feaElements $objEntity
        if {$opt(DEBUG1)} {outputMsg \n}
      }
      incr nfeaElem
    }
  }

# ----------  
# 1D elements, get from mesh
  if {$feaType == "curve_3d"} {
    puts $feaFile(elements) "<Switch whichChoice='0' id='sw1DElements'>"
    puts $feaFile(elements) " <Shape><Appearance><Material emissiveColor='1 0 1'></Material></Appearance>"
    puts $feaFile(elements) "  <IndexedLineSet coordIndex='"
    feaWriteIndex meshIndex elements
    puts $feaFile(elements) "  '>"
    puts $feaFile(elements) "   <Coordinate USE='coords'></Coordinate></IndexedLineSet></Shape>"
    puts $feaFile(elements) "</Switch>"

# 2D, 3D elements
  } else {

# write index file to mesh file
    puts $feaFile(mesh) " <Shape id='$feaType'><Appearance><Material emissiveColor='0 0 0'></Material></Appearance>"
    puts $feaFile(mesh) "  <IndexedLineSet coordIndex='"
    feaWriteIndex meshIndex mesh
    puts $feaFile(mesh) "  '>"
    puts $feaFile(mesh) "   <Coordinate USE='coords'></Coordinate></IndexedLineSet></Shape>"

# write faces index file to elements file
    if {$feaType == "surface_3d"} {
      puts $feaFile(elements) "<Switch whichChoice='0' id='sw2DElements'>"
      puts $feaFile(elements) " <Shape><Appearance><Material id='mat2D' diffuseColor='0 1 1'></Material></Appearance>"
      puts $feaFile(elements) "  <IndexedFaceSet solid='false' coordIndex='"
    } else {
      puts $feaFile(elements) "<Switch whichChoice='0' id='sw3DElements'>"      
      puts $feaFile(elements) " <Shape><Appearance><Material id='mat3D' diffuseColor='1 1 0'></Material></Appearance>"
      puts $feaFile(elements) "  <IndexedFaceSet id='faces' solid='false' coordIndex='"
    }
    if {[info exists feaFaceList]} {
      foreach face [array name feaFaceList] {

# add original face to file if only one sorted face, more means duplicate internal faces that can be eliminated
        if {$feaFaceList($face) == 1} {
# subdivide 4 noded faces
          if {[llength $face] == 4} {
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 0] [lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 2] -1"
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 0] [lindex $feaFaceOrig($face) 2] [lindex $feaFaceOrig($face) 3] -1"
# subdivide 6 noded faces
          } elseif {[llength $face] == 6} {
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 0] [lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 5] -1"
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 3] [lindex $feaFaceOrig($face) 5] -1"
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 2] [lindex $feaFaceOrig($face) 3] -1"
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 3] [lindex $feaFaceOrig($face) 4] [lindex $feaFaceOrig($face) 5] -1"
# subdivide 8 noded faces
          } elseif {[llength $face] == 8} {
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 0] [lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 7] -1"
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 2] [lindex $feaFaceOrig($face) 3] -1"
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 3] [lindex $feaFaceOrig($face) 7] -1"
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 7] [lindex $feaFaceOrig($face) 3] [lindex $feaFaceOrig($face) 5] -1"
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 7] [lindex $feaFaceOrig($face) 5] [lindex $feaFaceOrig($face) 6] -1"
            puts $feaFile(elements) "[lindex $feaFaceOrig($face) 3] [lindex $feaFaceOrig($face) 4] [lindex $feaFaceOrig($face) 5] -1"
# other faces that are not subdivided (3, 4)
          } else {
            puts $feaFile(elements) "$feaFaceOrig($face) -1"
          }
        }
      }
    }
    feaWriteIndex faceIndex elements
    puts $feaFile(elements) "  '>"
    puts $feaFile(elements) "   <Coordinate USE='coords'></Coordinate></IndexedFaceSet></Shape>"
    puts $feaFile(elements) "</Switch>"
  }

# check if done processing all element types
  set writeX3DOM 0
  if {$feaType == "volume_3d"} {set writeX3DOM 1}
  if {$feaType == "surface_3d" && ![info exists entCount(volume_3d_element_representation)]} {set writeX3DOM 1}
  if {$feaType == "curve_3d"   && ![info exists entCount(surface_3d_element_representation)] && \
                                  ![info exists entCount(volume_3d_element_representation)]} {set writeX3DOM 1}

# write mesh and elements at the end of processing an element type
  if {$writeX3DOM} {
    
# write mesh
    close $feaFile(mesh)
    if {[file size $feaFileName(mesh)] > 0} {
      puts $x3domFile "\n<Switch whichChoice='0' id='swMesh'><Group>"
      set feaFile(mesh) [open $feaFileName(mesh) r]
      while {[gets $feaFile(mesh) line] >= 0} {puts $x3domFile $line}
      puts $x3domFile "</Group></Switch>"
      close $feaFile(mesh)
    }
    
# write all elements
    close $feaFile(elements)
    if {[file size $feaFileName(elements)] > 0} {
      set feaFile(elements) [open $feaFileName(elements) r]
      while {[gets $feaFile(elements) line] >= 0} {puts $x3domFile $line}
      close $feaFile(elements)
    }

# close temp files
    foreach f {elements mesh meshIndex faceIndex} {
      catch {close $feaFileName($f)}
      catch {file delete -force $feaFileName($f)}
    }
    outputMsg " Finished writing FEM" blue
    catch {unset feaFaceList}
    catch {unset feaFaceOrig}
  }  

# messages
  if {[info exists feaTypes]} {
    foreach item [array names feaTypes] {lappend x3domMsg "$feaTypes($item) - $item"}
  }
}

# -------------------------------------------------------------------------------
proc feaElements {objEntity} {
  global badAttributes ent entAttrList entCount entLevel localName nistVersion opt
  global x3domFile x3domFileName x3domStartFile feaMeshIndex feaFaceIndex x3domMax x3domMin x3domMsg
  global idx feaIndex feaType feaTypes firstID nnode nnodes feaFaces nodeID nodeArr nfeaElem feaFile feaFaceList

# entLevel is very important, keeps track level of entity in hierarchy
  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]
  set idx(-1) -1

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
              if {[string first "handle" $objEntity] != -1} {feaElements $objValue}
            }
          } emsg3]} {
            errorMsg "ERROR processing FEM ($objNodeType $ent2)\n $emsg3"
          }

# --------------
# nodeType = 20
        } elseif {$objNodeType == 20} {
          if {[catch {
            if {$entAttrOK != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}
    
              switch -glob $ent1 {
                "*_element_representation node_list" {set nnodes $objSize}
              }

# referenced entities
              if {[catch {
                ::tcom::foreach val1 $objValue {feaElements $val1}
              } emsg]} {
                foreach val2 $objValue {feaElements $val2}
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing FEM ($objNodeType $ent2): $emsg3"
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
                }
                "*_element_descriptor description" {
# can also use with topology_order, Shape                
                  incr feaTypes($objValue)
                }
                "*node name" {
# node index
                  if {$objType == "node"} {
                    set nodeID $objID
                    set idx($nnode) $nodeArr($nodeID)
                    incr nnode
                    if {$nnode == 1} {set firstID $nodeArr($nodeID)}

# account for dummy nodes
                  } else {
                    incr nnodes -1
                  }
  
# done reading nodes, write index
# curve_3d
                  if {$nnode == $nnodes} {
                    set feaMeshIndex  ""
                    set feaFaceIndex ""
                    if {$feaType == "curve_3d"} {
                      if {$nnodes == 2} {
                        foreach id [array names idx] {append feaMeshIndex "$idx($id) "}
                      } elseif {$nnodes == 3} {
                        foreach id {0 2 1} {append feaMeshIndex "$idx($id) "}
                      } elseif {$nnodes == 4} {
                        foreach id {0 2 3 1} {append feaMeshIndex "$idx($id) "}
                      } else {
                        foreach id [lsort [array names idx]] {append feaMeshIndex "$idx($id) "}
                        errorMsg "Unexpected number of nodes ($nnodes) for a $feaType element"
                      }
                      #append feaMeshIndex "-1 "
                      unset idx

# surface_3d, volume_3d (feaIndex)
                    } elseif {$feaType == "surface_3d" || $feaType == "volume_3d"} {
                      if {[info exists feaIndex($feaType,$nnodes,line)]} {
# mesh
                        foreach id $feaIndex($feaType,$nnodes,line) {append feaMeshIndex "$idx($id) "}
# 2D faces
                        if {$feaFaces(2D) && $feaType == "surface_3d"} {
                          foreach id $feaIndex($feaType,$nnodes,surf) {append feaFaceIndex "$idx($id) "}
# 3D faces
                        } elseif {$feaFaces(3D) && $feaType == "volume_3d"} {
# write faces                          
                          if {[info exist feaIndex($feaType,$nnodes,face)]} {
                            set j 0
                            catch {unset nidx}
                            foreach id $feaIndex($feaType,$nnodes,face) {
                              if {$id != -1} {
                                lappend nidx($j) $idx($id)
                              } else {
                                incr j
                              }
                            }
                            for {set k 0} {$k < $j} {incr k} {
                              set f1 [feaFaceSort $nidx($k)]
                              incr feaFaceList($f1)
                            }
# write index
                          } elseif {[info exist feaIndex($feaType,$nnodes,surf)]} {
                            foreach id $feaIndex($feaType,$nnodes,surf) {append feaFaceIndex "$idx($id) "}
                          } else {
                            errorMsg "Unexpected number of nodes ($nnodes) for a $feaType element"
                          }
                        }
# unexpected element
                      } else {
                        foreach id [lsort [array names idx]] {append feaMeshIndex "$idx($id) "}
                        append feaMeshIndex  "$firstID -1 "
                        if {$nnodes > 4} {
                          errorMsg "Unexpected number of nodes ($nnodes) for a $feaType element"
                          lappend x3domMsg "Unexpected number of nodes ($nnodes) for a $feaType element.  Faces not displayed."
                        }
                      }
                      unset idx

                    } else {
                      errorMsg "Unexpected element type $feaType"
                    }
                    
# write index to file
                    if {$feaMeshIndex  != ""} {puts $feaFile(meshIndex) $feaMeshIndex}
                    if {$feaFaceIndex != ""} {puts $feaFile(faceIndex) $feaFaceIndex}
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing FEM ($objNodeType $ent2): $emsg3"
            set entLevel 1
          }
        }
      }
    }
  }
  incr entLevel -1
}

# -------------------------------------------------------------------------------
# sort face list
proc feaFaceSort {face} {
  global feaFaceOrig
  set saveFace $face
  
# simple sort for 3 noded faces
  if {[llength $face] == 3} {
    set face [lsort $face]

# find smallest node id
  } else {
    set sid [lindex [lsort $face] 0]
    set ps [lsearch $face $sid]

# if smallest is not first, then shift order
    if {$ps != 0} {set face [concat [lrange $face $ps end] [lrange $face 0 $ps-1]]}

# if last id less than second id, then reverse second through last
    if {[lindex $face end] < [lindex $face 1]} {
      set f1 [lrange $face 1 end]
      set nf {}
      set i [llength $f1]
      while {[incr i -1]} {lappend nf [lindex $f1 $i]}
      lappend nf [lindex $f1 0]
      set face [concat [lindex $face 0] $nf]
    }
  }
  
# save original face ordering to preserve outward normal  
  set feaFaceOrig($face) $saveFace
  return $face
}

# -------------------------------------------------------------------------------
# get and write nodes
proc feaGetNodes {objDesign} {
  global entCount nodeArr x3domMax x3domMin x3domFile x3domMsg
  catch {unset nodeArr}

  set nodeIndex -1
  outputMsg " Reading nodes ($entCount(node))" blue

# start node output
  puts $x3domFile "\n<Switch whichChoice='0' id='swNodes'>"
  puts $x3domFile " <Shape><Appearance><Material emissiveColor='0 0 1'></Material></Appearance>"
  puts $x3domFile "  <PointSet><Coordinate DEF='coords' point='"
  
# get all nodes
 ::tcom::foreach e0 [$objDesign FindObjects [string trim node]] {
    set p21id [$e0 P21ID]
    set a0 [[$e0 Attributes] Item 2]
    ::tcom::foreach e1 [$a0 Value] {
      set coordList [[[$e1 Attributes] Item 2] Value]
      set nodeArr($p21id) [incr nodeIndex]
  
# xyz coordinates
      set coord(x) [lindex $coordList 0]
      set coord(y) [lindex $coordList 1]
      set coord(z) [lindex $coordList 2]
      puts $x3domFile "[trimNum $coord(x) 5] [trimNum $coord(y) 5] [trimNum $coord(z) 5]"

# min max      
      foreach xyz {x y z} {
        if {$coord($xyz) > $x3domMax($xyz)} {set x3domMax($xyz) $coord($xyz)}
        if {$coord($xyz) < $x3domMin($xyz)} {set x3domMin($xyz) $coord($xyz)}
      }
      if {[expr {$nodeIndex%50000}] == 0} {
        if {$nodeIndex > 0} {outputMsg "  $nodeIndex"}
        update idletasks
      }
    }
  }

# finish node output
  puts $x3domFile "   '></Coordinate></PointSet></Shape>"
  puts $x3domFile "</Switch>"

  lappend x3domMsg "$entCount(node) - Nodes"
  if {[info exists entCount(dummy_node)]} {lappend x3domMsg "$entCount(dummy_node) - Dummy Node"}
}

# -------------------------------------------------------------------------------
# write index to file
proc feaWriteIndex {idx x3d} {
  global feaFile feaFileName
  close $feaFile($idx)
  if {[file size $feaFileName($idx)] > 0} {
    set feaFile($idx) [open $feaFileName($idx) r]
    while {[gets $feaFile($idx) line] >= 0} {puts $feaFile($x3d) $line}
    close $feaFile($idx)
    catch {file delete -force $feaFileName($idx)}
  }
}

# -------------------------------------------------------------------------------
# script for switch node
proc feaSwitch {type} {
  global x3domFile
  
  puts $x3domFile "<script>\nfunction tog$type\(choice){"
  puts $x3domFile " if (!document.getElementById('sw$type').checked) {\n  document.getElementById('sw$type').setAttribute('whichChoice', -1);\n } else {\n  document.getElementById('sw$type').setAttribute('whichChoice', 0);\n }"
  puts $x3domFile " document.getElementById('sw$type').checked = !document.getElementById('sw$type').checked;\n}\n</script>"
}