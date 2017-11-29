proc feaModel {entType} {
  global objDesign
  global ent entAttrList entCount entLevel opt rowmax nprogBarEnts count localName mytemp sfaPID cadSystem
  global x3dFile x3dMin x3dMax x3dMsg x3dStartFile x3dFileName x3dAxesSize
  global feaType feaTypes feaElemTypes nfeaElem feaFile feaFileName feaFaceList feaFaceOrig feaBC feaLD feaMeshIndex

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

  set cartesian_point [list cartesian_point coordinates]
  set FEA(single_point_constraint_element_values) [list single_point_constraint_element_values \
                                                    defined_state [list specified_state state_id description] \
                                                    element [list single_point_constraint_element required_node [list node name items $cartesian_point]]]
  set FEA(nodal_freedom_action_definition) [list nodal_freedom_action_definition \
                                              defined_state [list specified_state state_id description] \
                                              node [list node name items $cartesian_point]]
  
  if {[info exists ent]} {unset ent}
  set entLevel 0
  set entAttrList {}
  setEntAttrList $FEA($entType)
  catch {unset feaTypes}

# ---------- 
# start X3DOM file                
  if {$x3dStartFile} {
    set x3dStartFile 0
    catch {file delete -force -- "[file rootname $localName]_x3dom.html"}
    catch {file delete -force -- "[file rootname $localName]-x3dom.html"}
    set x3dFileName [file rootname $localName]-sfa.html
    catch {file delete -force -- $x3dFileName}
    set x3dFile [open $x3dFileName w]
    puts $x3dFile "<!DOCTYPE html>\n<html>\n<head>\n<title>[file tail $localName] | STEP AP209 Finite Element Model</title>\n<base target=\"_blank\">\n<meta http-equiv='Content-Type' content='text/html;charset=utf-8'/>"
    puts $x3dFile "<link rel='stylesheet' type='text/css' href='https://www.x3dom.org/x3dom/release/x3dom.css'/>\n<script type='text/javascript' src='https://www.x3dom.org/x3dom/release/x3dom.js'></script>\n"

# node, element checkbox script
    x3dSwitchScript Nodes
    if {[info exists entCount(surface_3d_element_representation)] || \
        [info exists entCount(volume_3d_element_representation)]}  {x3dSwitchScript Mesh}
    if {[info exists entCount(curve_3d_element_representation)]}   {x3dSwitchScript 1DElements}
    if {[info exists entCount(surface_3d_element_representation)]} {x3dSwitchScript 2DElements}
    if {[info exists entCount(volume_3d_element_representation)]}  {x3dSwitchScript 3DElements}

# transparency script
    if {[info exists entCount(surface_3d_element_representation)] || [info exists entCount(volume_3d_element_representation)]} {
      puts $x3dFile "<script>function matTrans(trans){"
      if {[info exists entCount(surface_3d_element_representation)]} {puts $x3dFile " document.getElementById('mat2D').setAttribute('transparency', trans);"}
      if {[info exists entCount(volume_3d_element_representation)]}  {
        puts $x3dFile " document.getElementById('mat3D').setAttribute('transparency', trans);"
        puts $x3dFile " if (trans > 0) {document.getElementById('faces').setAttribute('solid', true);} else {document.getElementById('faces').setAttribute('solid', false);}"
      }
      puts $x3dFile "}\n</script>"
    }
    puts $x3dFile "</head>"

    set name [file tail $localName]
    if {$cadSystem != ""} {append name "  ($cadSystem)"}
    puts $x3dFile "\n<body><font face=\"arial\">\n<h3>STEP AP209 Finite Element Model:  $name</h3>"
    puts $x3dFile "<table><tr><td>"
    
# x3d window size
    set height 800
    set width [expr {int($height*1.5)}]
    catch {
      set height [expr {int([winfo screenheight .]*0.75)}]
      set width [expr {int($height*[winfo screenwidth .]/[winfo screenheight .])}]
    }
    puts $x3dFile "\n<X3D id='someUniqueId' showStat='false' showLog='false' x='0px' y='0px' width='$width\px' height='$height\px'>\n<Scene DEF='scene'>"

# nodes    
    feaGetNodes
    outputMsg " Writing FEM to: [truncFileName [file nativename $x3dFileName]]" green

# coordinate min, max, center  
    foreach xyz {x y z} {
      set delt($xyz) [expr {$x3dMax($xyz)-$x3dMin($xyz)}]
      set xyzcen($xyz) [format "%.4f" [expr {0.5*$delt($xyz) + $x3dMin($xyz)}]]
    }
    set maxxyz $delt(x)
    if {$delt(y) > $maxxyz} {set maxxyz $delt(y)}
    if {$delt(z) > $maxxyz} {set maxxyz $delt(z)}

# coordinate axes    
    set x3dAxesSize [trimNum [expr {$maxxyz*0.05}]]
    x3dCoordAxes $x3dAxesSize
  }

# create temp files so that a lot of elements can be processed
  if {[string first "_3d" $entType] != -1} {
    foreach f {elements mesh meshIndex faceIndex} {
      set feaFileName($f) [file join $mytemp $f.txt]
      if {![file exists $feaFileName($f)]} {set feaFile($f) [open $feaFileName($f) w]}
    }
  }

# ----------  
# process all *_element_representation entities, call feaElements  
  set nfeaElem 0
  catch {unset feaFaceList}
  catch {unset feaFaceOrig}

  set startent [lindex $FEA($entType) 0]
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    if {[$objEntity Type] == $startent} {
      if {$opt(XLSCSV) == "None"} {incr nprogBarEnts}
      
      if {$nfeaElem < 10000000} {
        if {[expr {$nfeaElem%2000}] == 0} {

# check memory and gracefully exit
          set mem [expr {[lindex [twapi::get_process_info $sfaPID -pagefilebytes] 1]/1048576}]
          if {$mem > 1600} {
            errorMsg "Insufficient memory to process all of the elements"
            lappend x3dMsg "Some elements were not processed."
            update idletasks
            break
          }
          update
        }
        if {$nfeaElem > $rowmax && $opt(XLSCSV) != "None"} {incr nprogBarEnts}

# process the entity
        feaElements $objEntity
        if {$opt(DEBUG1)} {outputMsg \n}
      }
      incr nfeaElem
    }
  }
  update idletasks

# ----------  
# 1D elements, get from mesh
  if {[string first "_3d" $startent] != -1} {
    puts $feaFile(elements) "\n<!-- $feaType elements -->"
    if {$feaType == "curve_3d"} {
      puts $feaFile(elements) "<Switch whichChoice='0' id='sw1DElements'>"
      puts $feaFile(elements) " <Shape><Appearance><Material emissiveColor='1 0 1'></Material></Appearance>"
      puts $feaFile(elements) "  <IndexedLineSet coordIndex='"
      feaWriteIndex meshIndex elements
      puts $feaFile(elements) "  '><Coordinate USE='coords'></Coordinate></IndexedLineSet></Shape>"
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
    if {[file exists $feaFileName(mesh)]} {
      close $feaFile(mesh)
      if {[file size $feaFileName(mesh)] > 0} {
        puts $x3dFile "<!-- wireframe mesh -->\n<Switch whichChoice='0' id='swMesh'><Group>"
        set feaFile(mesh) [open $feaFileName(mesh) r]
        while {[gets $feaFile(mesh) line] >= 0} {puts $x3dFile $line}
        puts $x3dFile "</Group></Switch>"
        close $feaFile(mesh)
        unset feaMeshIndex
      }
    }
    
# write all elements
    if {[file exists $feaFileName(elements)]} {
      close $feaFile(elements)
      if {[file size $feaFileName(elements)] > 0} {
        set feaFile(elements) [open $feaFileName(elements) r]
        while {[gets $feaFile(elements) line] >= 0} {puts $x3dFile $line}
        close $feaFile(elements)
        outputMsg " Finished writing elements" green
      }
    }

# write loads
    if {[info exists feaLD] && $entType == "nodal_freedom_action_definition"} {
      set size [trimNum [expr {$x3dAxesSize*0.15}]]
      puts $x3dFile " "
      set n 0
      set i 0
      puts $x3dFile "<!-- loads -->"
      foreach load [lsort [array names feaLD]] {
        incr n
        puts $x3dFile "<Switch whichChoice='-1' id='load$n'><Group>"
        foreach bc $feaLD($load) {
          if {$i == 0} {
            puts $x3dFile "<Transform translation='$bc'><Group DEF='LOAD-shape'>"
            puts $x3dFile " <Transform translation='0 [expr {2.*$size}] 0'><Shape>\n  <Appearance><Material diffuseColor='.6 .6 .6'></Material></Appearance>\n  <Cone bottomRadius='[expr {$size*0.5}]' height='[expr {$size*2.}]'>\n </Shape><\Transform>"
            puts $x3dFile " <Shape>\n  <Appearance><Material diffuseColor='1 1 1'></Material></Appearance>\n  <IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0 [expr {$size*-2.}] 0 0 -$size 0'></Coordinate></IndexedLineSet>\n </Shape>\n</Group></Transform>"
          } else {
            puts $x3dFile "<Transform translation='$bc'><Group USE='LOAD-shape'></Group></Transform>"
          }
          incr i
        }
        puts $x3dFile "</Group></Switch>\n"
      }
    }

# write boundary conditions
    if {[info exists feaBC] && $entType == "single_point_constraint_element_values"} {
      set size [trimNum [expr {$x3dAxesSize*0.2}]]
      puts $x3dFile " "
      set n 0
      set i 0
      puts $x3dFile "<!-- boundary conditions -->"
      foreach spc [lsort [array names feaBC]] {
        incr n
        puts $x3dFile "<Switch whichChoice='-1' id='spc$n'><Group>"
        foreach bc $feaBC($spc) {
          if {$i == 0} {
            puts $x3dFile "<Transform translation='$bc'>\n <Shape DEF='SPC-shape'>\n  <Appearance><Material diffuseColor='1 0 0'></Material></Appearance>\n  <Box size='$size $size $size'>\n </Shape>\n</Transform>"
          } else {
            puts $x3dFile "<Transform translation='$bc'><Shape USE='SPC-shape'></Shape></Transform>"
          }
          incr i
        }
        puts $x3dFile "</Group></Switch>\n"
      }
    }
    
# close temp files
    if {[string first "_3d" $entType] != -1} {
      foreach f {elements mesh meshIndex faceIndex} {
        if {[file exists $feaFileName($f)]} {
          catch {close $feaFileName($f)}
          catch {file delete -force $feaFileName($f)}
        }
      }
      catch {unset feaFaceList}
      catch {unset feaFaceOrig}
    }
  }  

# messages
  if {[info exists feaTypes]} {
    foreach item [array names feaTypes] {lappend x3dMsg "$feaTypes($item) - [string tolower $item]"}
  }
}

# -------------------------------------------------------------------------------
proc feaElements {objEntity} {
  global badAttributes ent entAttrList entCount entLevel localName nistVersion opt
  global x3dFile x3dFileName x3dStartFile feaMeshIndex feaFaceIndex x3dMsg feaStateID feaEntity
  global feaidx feaIndex feaType feaTypes firstID nnode nnodes nodeID nfeaElem feaFile feaFaceList feaBC feaLD
  global nodeArr

# entLevel is very important, keeps track level of entity in hierarchy
  incr entLevel
  set ind [string repeat " " [expr {4*($entLevel-1)}]]
  set feaidx(-1) -1

  if {[string first "handle" $objEntity] != -1} {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set objAttributes [$objEntity Attributes]
    set ent($entLevel) $objType
    if {$entLevel == 1} {set feaEntity $objType}

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
                "cartesian_point coordinates" {
                  if {$feaEntity == "nodal_freedom_action_definition"} {
                    lappend feaLD($feaStateID) [vectrim $objValue]
                  } elseif {$feaEntity == "single_point_constraint_element_values"} {
                    lappend feaBC($feaStateID) [vectrim $objValue]
                  }
                }
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
                  if {[string first "_3d" $feaEntity] != -1} {
                    if {$objType == "node"} {
                      set nodeID $objID
                      set feaidx($nnode) $nodeArr($nodeID)
                      if {[info exists nnode]} {
                        incr nnode
                        if {$nnode == 1} {set firstID $nodeArr($nodeID)}
                      }
  
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
                          foreach id [array names feaidx] {append feaMeshIndex "$feaidx($id) "}
                        } elseif {$nnodes == 3} {
                          foreach id {0 2 1} {append feaMeshIndex "$feaidx($id) "}
                        } elseif {$nnodes == 4} {
                          foreach id {0 2 3 1} {append feaMeshIndex "$feaidx($id) "}
                        } else {
                          foreach id [lsort [array names feaidx]] {append feaMeshIndex "$feaidx($id) "}
                          errorMsg "Unexpected number of nodes ($nnodes) for a $feaType element"
                        }
                        unset feaidx
  
# surface_3d, volume_3d (feaIndex)
                      } elseif {$feaType == "surface_3d" || $feaType == "volume_3d"} {
                        if {[info exists feaIndex($feaType,$nnodes,line)]} {
# mesh
                          foreach id $feaIndex($feaType,$nnodes,line) {append feaMeshIndex "$feaidx($id) "}
# 2D faces
                          if {$feaType == "surface_3d"} {
                            foreach id $feaIndex($feaType,$nnodes,surf) {append feaFaceIndex "$feaidx($id) "}
# 3D faces
                          } elseif {$feaType == "volume_3d"} {
# write faces                          
                            if {[info exist feaIndex($feaType,$nnodes,face)]} {
                              set j 0
                              catch {unset nidx}
                              foreach id $feaIndex($feaType,$nnodes,face) {
                                if {$id != -1} {
                                  lappend nidx($j) $feaidx($id)
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
                              foreach id $feaIndex($feaType,$nnodes,surf) {append feaFaceIndex "$feaidx($id) "}
                            } else {
                              errorMsg "Unexpected number of nodes ($nnodes) for a $feaType element"
                            }
                          }
# unexpected element
                        } else {
                          foreach id [lsort [array names feaidx]] {append feaMeshIndex "$feaidx($id) "}
                          append feaMeshIndex  "$firstID -1 "
                          if {$nnodes > 4} {
                            errorMsg "Unexpected number of nodes ($nnodes) for a $feaType element"
                            lappend x3dMsg "$feaType element with an unexpected number of nodes ($nnodes)"
                          }
                        }
                        unset feaidx
  
                      } else {
                        errorMsg "Unexpected element type $feaType"
                      }
                      
# write index to file
                      if {$feaMeshIndex != ""} {puts $feaFile(meshIndex) $feaMeshIndex}
                      if {$feaFaceIndex != ""} {puts $feaFile(faceIndex) $feaFaceIndex}
                    }
                  }
                }
                "specified_state state_id" {
                  set feaStateID $objValue
                }
                "specified_state description" {
                  if {$feaStateID == ""} {set feaStateID $objValue}
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
  
# simple sort for 3 noded faces
  set saveFace $face
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
proc feaGetNodes {} {
  global objDesign
  global entCount nodeArr x3dMax x3dMin x3dFile x3dMsg
  catch {unset nodeArr}

  set nodeIndex -1
  outputMsg " Reading nodes ($entCount(node))" green

# start node output
  puts $x3dFile "\n<!-- nodes -->\n<Switch whichChoice='0' id='swNodes'>"
  puts $x3dFile " <Shape><Appearance><Material emissiveColor='0 0 1'></Material></Appearance>"
  puts $x3dFile "  <PointSet><Coordinate DEF='coords' point='"
  
# get all nodes
  set n 0
  set nline ""
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
      append nline "[trimNum $coord(x) 5] [trimNum $coord(y) 5] [trimNum $coord(z) 5] "
      incr n
      if {$n == 10} {
        puts $x3dFile $nline
        set n 0
        set nline ""
      }

# min max      
      foreach xyz {x y z} {
        if {$coord($xyz) > $x3dMax($xyz)} {set x3dMax($xyz) $coord($xyz)}
        if {$coord($xyz) < $x3dMin($xyz)} {set x3dMin($xyz) $coord($xyz)}
      }

      if {[expr {$nodeIndex%50000}] == 0} {
        if {$nodeIndex > 0} {outputMsg "  $nodeIndex"}
        update
      }
    }
  }
  if {[string length $nline] > 0} {puts $x3dFile $nline}

# finish node output
  puts $x3dFile "  '></Coordinate></PointSet></Shape>"
  puts $x3dFile "</Switch>"

  lappend x3dMsg "$entCount(node) - nodes"
  if {[info exists entCount(dummy_node)]} {lappend x3dMsg "$entCount(dummy_node) - dummy node"}
}

# -------------------------------------------------------------------------------
# write index to file
proc feaWriteIndex {ftyp x3d} {
  global feaFile feaFileName

  close $feaFile($ftyp)
  if {[file size $feaFileName($ftyp)] > 0} {
    set feaFile($ftyp) [open $feaFileName($ftyp) r]
    
    set n 0
    set nline ""
    while {[gets $feaFile($ftyp) line] >= 0} {
      #outputMsg $line blue
      append nline "[string trim $line] "
      incr n
      if {$n == 10} {
        puts $feaFile($x3d) $nline
        set n 0
        set nline ""
      }
    }
    if {[string length $nline] > 0} {puts $feaFile($x3d) $nline}
      
    close $feaFile($ftyp)
    catch {file delete -force $feaFileName($ftyp)}
  }
}
