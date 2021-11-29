proc feaModel {entType} {
  global objDesign
  global cadSystem ent entAttrList entCount entLevel feaBoundary feaDisp feaFaceList feaFaceOrig feaFile feaFileName
  global feaFirstEntity feaLastEntity feaLoad feaMeshIndex feaType feaTypes localName mytemp nfeaElem nprogBarEnts opt rowmax
  global stepAP timeStamp x3dAxesSize x3dFile x3dFileName x3dMax x3dMin x3dMsg x3dStartFile x3dTitle x3dViewOK

  if {$opt(DEBUG1)} {outputMsg "START feaModel $entType\n" red}

# finite elements
  set node            [list node name]
  set dummy_node      [list dummy_node name]
  set curve_3d_element_descriptor   [list curve_3d_element_descriptor   description]
  set surface_3d_element_descriptor [list surface_3d_element_descriptor description]
  set volume_3d_element_descriptor  [list volume_3d_element_descriptor  description]

  set FEA(curve_3d_element_representation)   [list curve_3d_element_representation   name node_list $node $dummy_node element_descriptor $curve_3d_element_descriptor]
  set FEA(surface_3d_element_representation) [list surface_3d_element_representation name node_list $node $dummy_node element_descriptor $surface_3d_element_descriptor]
  set FEA(volume_3d_element_representation)  [list volume_3d_element_representation  name node_list $node $dummy_node element_descriptor $volume_3d_element_descriptor]

  set cartesian_point [list cartesian_point coordinates]
  set node [list node name items $cartesian_point]
  set node_group [list node_group nodes]
  set freedom [list freedoms_list freedoms]
  set single_point_constraint_element \
    [list single_point_constraint_element required_node $node $node_group freedoms_and_values [list freedom_and_coefficient freedom a]]

  set state_id1 [list specified_state state_id description]
  set state_id2 [list state state_id description]
  set state_id3 [list calculated_state state_id description]
  set surface_nodes [list surface_3d_element_representation node_list $node]
  set volume_nodes  [list volume_3d_element_representation node_list $node]

# boundary conditions
  set FEA(single_point_constraint_element_values) \
    [list single_point_constraint_element_values \
      defined_state $state_id1 $state_id2 \
      element $single_point_constraint_element \
      degrees_of_freedom $freedom b]

# nodal loads
  set FEA(nodal_freedom_action_definition) \
    [list nodal_freedom_action_definition \
      defined_state $state_id1 $state_id2 \
      node $node $node_group \
      degrees_of_freedom $freedom values]

# element loads
  set FEA(surface_3d_element_boundary_constant_specified_surface_variable_value) \
    [list surface_3d_element_boundary_constant_specified_surface_variable_value \
      defined_state $state_id1 $state_id2 \
      element $surface_nodes [list surface_3d_element_group elements $surface_nodes] \
      simple_value variable element_face]

  set FEA(volume_3d_element_boundary_constant_specified_variable_value) \
    [list volume_3d_element_boundary_constant_specified_variable_value \
      defined_state $state_id1 $state_id2 \
      element $volume_nodes [list volume_3d_element_group elements $volume_nodes] \
      simple_value variable element_face]

# nodal results
  set FEA(nodal_freedom_values) \
    [list nodal_freedom_values \
      defined_state $state_id3 $state_id1 $state_id2 \
      node $node \
      degrees_of_freedom $freedom values \
      values]

# element nodal results (disabled for now in sfa-gen.tcl, sfa-step.tcl)
  set FEA(element_nodal_freedom_actions) \
    [list element_nodal_freedom_actions \
      defined_state [list specified_state state_id description] [list calculated_state state_id description] [list state state_id description] \
      element [list curve_3d_element_representation   node_list $node $node_group] \
              [list surface_3d_element_representation node_list $node $node_group] \
              [list volume_3d_element_representation  node_list $node $node_group] \
      nodal_action [list element_nodal_freedom_terms degrees_of_freedom $freedom values]]

  if {[info exists ent]} {unset ent}
  set entLevel 0
  set entAttrList {}
  setEntAttrList $FEA($entType)
  catch {unset feaTypes}

# ------------------------------------------------------------------------------------------------
# start X3DOM file
  if {$x3dStartFile} {
    set x3dStartFile 0
    set x3dViewOK 1
    catch {file delete -force -- "[file rootname $localName]_x3dom.html"}
    catch {file delete -force -- "[file rootname $localName]-x3dom.html"}
    set x3dFileName [file rootname $localName]-sfa.html
    catch {file delete -force -- $x3dFileName}
    set x3dFile [open $x3dFileName w]

    set title [file tail $localName]
    if {$stepAP != "" && [string range $stepAP 0 1] == "AP"} {append title " | $stepAP"}
    puts $x3dFile "<!DOCTYPE html>\n<html>\n<head>\n<title>$title</title>\n<base target=\"_blank\">\n<meta http-equiv='Content-Type' content='text/html;charset=utf-8'/>"

# use x3dom 1.8.1 because 1.8.2 breaks transparency
    puts $x3dFile "<link rel='stylesheet' type='text/css' href='https://www.x3dom.org/download/1.8.1/x3dom.css'/>\n<script type='text/javascript' src='https://www.x3dom.org/download/1.8.1/x3dom.js'></script>"
    #puts $x3dFile "<link rel='stylesheet' type='text/css' href='https://www.x3dom.org/x3dom/release/x3dom.css'/>\n<script type='text/javascript' src='https://www.x3dom.org/x3dom/release/x3dom.js'></script>"
    puts $x3dFile "</head>"

    set x3dTitle [file tail $localName]
    if {$stepAP != "" && [string range $stepAP 0 1] == "AP"} {append x3dTitle "&nbsp;&nbsp;&nbsp;$stepAP"}
    if {$timeStamp != ""} {
      set ts [fixTimeStamp $timeStamp]
      append x3dTitle "&nbsp;&nbsp;&nbsp;$ts"
    }
    if {$cadSystem != ""} {
      regsub -all "_" $cadSystem " " cs
      append x3dTitle "&nbsp;&nbsp;&nbsp;$cs"
    }
    puts $x3dFile "\n<body><font face=\"sans-serif\">\n<h3>$x3dTitle</h3>"
    puts $x3dFile "\n<table><tr><td valign='top' width='85%'>\n<table><tr><td style='border:1px solid black'>\n<noscript>JavaScript must be enabled in the web browser</noscript>"

# x3d window size
    set height 800
    set width [expr {int($height*1.5)}]
    catch {
      set height [expr {int([winfo screenheight .]*0.85)}]
      set width [expr {int($height*[winfo screenwidth .]/[winfo screenheight .])}]
    }

# start x3d with flat-to-screen hud for viewpoint text
    puts $x3dFile "\n<X3D id='x3d' showStat='false' showLog='false' x='0px' y='0px' width='$width' height='$height'>"
    puts $x3dFile "<div id='HUDs_Div'><div class='group' style='margin:2px; margin-top:0px; padding:4px; background-color:rgba(0,0,0,1.); position:absolute; float:center; z-index:1000;'>Viewpoint: <span id='clickedView'></span></div></div>"
    puts $x3dFile "<Scene DEF='scene'>"
    puts $x3dFile "<!-- X3D generated by the NIST STEP File Analyzer and Viewer [getVersion] -->"

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

# create temp files
  if {[string first "element_representation" $entType] != -1} {
    checkTempDir
    foreach f {faceIndex meshIndex} {
      set feaFileName($f) [file join $mytemp $f.txt]
      set feaFile($f) [open $feaFileName($f) w]
    }
    if {$entType == $feaFirstEntity} {
      foreach f {bcs elements loads mesh} {
        set feaFileName($f) [file join $mytemp $f.txt]
        set feaFile($f) [open $feaFileName($f) w]
      }
    }
  }

# ------------------------------------------------------------------------------------------------
# process all *_element_representation, load, and BC entities > call feaEntities
  set nfeaElem 0
  catch {unset feaFaceList}
  catch {unset feaFaceOrig}

  set skipElements 0
  set startent [lindex $FEA($entType) 0]
  set memMsg "<i>Some elements were not processed.</i>"

  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    if {[$objEntity Type] == $startent && (!$skipElements || [string first "element_representation" $startent] == -1)} {
      if {$opt(xlFormat) == "None"} {incr nprogBarEnts}
      if {$nfeaElem < 10000000} {

# check memory and gracefully exit
        if {[expr {$nfeaElem%2000}] == 0} {
          set mem [expr {[lindex [twapi::get_process_info [pid] -pagefilebytes] 1]/1048576}]
          if {$mem > 1700} {
            errorMsg "Insufficient memory to process all elements"
            if {[lsearch $x3dMsg $memMsg] == -1} {
              lappend x3dMsg $memMsg
              set skipElements 1
            }
          }
          update idletasks
        }
        if {$nfeaElem > $rowmax && $opt(xlFormat) != "None"} {incr nprogBarEnts}

# process the entity
        feaEntities $objEntity
        if {$opt(DEBUG1)} {outputMsg \n}
      }
      incr nfeaElem
    }
  }
  update idletasks

# ------------------------------------------------------------------------------------------------
# 1D elements, get from mesh
  if {[string first "element_representation" $entType] != -1 && [info exists feaType]} {
    puts $feaFile(elements) "\n<!-- [string toupper $feaType] ELEMENTS -->"
    if {$feaType == "curve_3d"} {
      puts $feaFile(elements) "<Switch whichChoice='0' id='sw1DElements'>"
      puts $feaFile(elements) " <Shape><Appearance><Material emissiveColor='1 0 1'/></Appearance>"
      puts $feaFile(elements) "  <IndexedLineSet coordIndex='"
      feaWriteIndex meshIndex elements
      puts $feaFile(elements) "  '><Coordinate USE='nodes'/></IndexedLineSet></Shape>"
      puts $feaFile(elements) "</Switch>"

# 2D, 3D elements
    } else {

# write index file to mesh file
      puts $feaFile(mesh) " <Shape id='$feaType'><Appearance><Material emissiveColor='0 0 0'/></Appearance>"
      puts $feaFile(mesh) "  <IndexedLineSet coordIndex='"
      feaWriteIndex meshIndex mesh
      puts $feaFile(mesh) "  '>"
      puts $feaFile(mesh) "   <Coordinate USE='nodes'/></IndexedLineSet></Shape>"

# write faces index file to elements file
      if {$feaType == "surface_3d"} {
        puts $feaFile(elements) "<Switch whichChoice='0' id='sw2DElements'>"
        puts $feaFile(elements) " <Shape><Appearance><Material id='mat2Dfem' diffuseColor='0 1 1'/></Appearance>"
        puts $feaFile(elements) "  <IndexedFaceSet solid='false' coordIndex='"
      } else {
        puts $feaFile(elements) "<Switch whichChoice='0' id='sw3DElements'>"
        puts $feaFile(elements) " <Shape><Appearance><Material id='mat3Dfem' diffuseColor='1 1 0'/></Appearance>"
        puts $feaFile(elements) "  <IndexedFaceSet id='faces' solid='false' coordIndex='"
      }
      if {[info exists feaFaceList]} {
        set n 0
        set nline ""
        foreach face [array name feaFaceList] {

# add original face to file if only one sorted face, more means duplicate internal faces that can be eliminated
          if {$feaFaceList($face) == 1} {
            incr n

# subdivide 4 noded faces
            if {[llength $face] == 4} {
              append nline "[lindex $feaFaceOrig($face) 0] [lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 2] -1 "
              append nline "[lindex $feaFaceOrig($face) 0] [lindex $feaFaceOrig($face) 2] [lindex $feaFaceOrig($face) 3] -1 "
# subdivide 6 noded faces
            } elseif {[llength $face] == 6} {
              append nline "[lindex $feaFaceOrig($face) 0] [lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 5] -1 "
              append nline "[lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 3] [lindex $feaFaceOrig($face) 5] -1 "
              append nline "[lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 2] [lindex $feaFaceOrig($face) 3] -1 "
              append nline "[lindex $feaFaceOrig($face) 3] [lindex $feaFaceOrig($face) 4] [lindex $feaFaceOrig($face) 5] -1 "
# subdivide 8 noded faces
            } elseif {[llength $face] == 8} {
              append nline "[lindex $feaFaceOrig($face) 0] [lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 7] -1 "
              append nline "[lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 2] [lindex $feaFaceOrig($face) 3] -1 "
              append nline "[lindex $feaFaceOrig($face) 1] [lindex $feaFaceOrig($face) 3] [lindex $feaFaceOrig($face) 7] -1 "
              append nline "[lindex $feaFaceOrig($face) 7] [lindex $feaFaceOrig($face) 3] [lindex $feaFaceOrig($face) 5] -1 "
              append nline "[lindex $feaFaceOrig($face) 7] [lindex $feaFaceOrig($face) 5] [lindex $feaFaceOrig($face) 6] -1 "
              append nline "[lindex $feaFaceOrig($face) 3] [lindex $feaFaceOrig($face) 4] [lindex $feaFaceOrig($face) 5] -1 "
# other faces that are not subdivided (3, 4)
            } else {
              append nline "$feaFaceOrig($face) -1 "
            }
            if {$n == 10} {
              puts $feaFile(elements) $nline
              set n 0
              set nline ""
            }
          }
        }
        if {[string length $nline] > 0} {puts $feaFile(elements) $nline}
      }
      feaWriteIndex faceIndex elements
      puts $feaFile(elements) "  '>"
      puts $feaFile(elements) "   <Coordinate USE='nodes'/></IndexedFaceSet></Shape>"
      puts $feaFile(elements) "</Switch>"
    }
  }

# ------------------------------------------------------------------------------------------------
# write loads after all load entity types are processed
  set ecload3 0
  set ecload4 0
  if {[info exists feaLoad] && $opt(feaLoads)} {
    set ecload3 [info exists entCount(surface_3d_element_boundary_constant_specified_surface_variable_value)]
    set ecload4 [info exists entCount(volume_3d_element_boundary_constant_specified_variable_value)]

    if {($entType == "nodal_freedom_action_definition" && !$ecload3 && !$ecload4) || \
        ($entType == "surface_3d_element_boundary_constant_specified_surface_variable_value" && !$ecload4) || \
         $entType == "volume_3d_element_boundary_constant_specified_variable_value"} {
      feaLoads $entType
    }
  }

# write displacements
  if {[info exists feaDisp] && $opt(feaDisp) && $entType == "nodal_freedom_values"} {feaLoads $entType}

# write boundary conditions
  if {[info exists feaBoundary] && $opt(feaBounds) && $entType == "single_point_constraint_element_values"} {feaBCs $entType}

# ------------------------------------------------------------------------------------------------
# write mesh, elements, loads, bcs after last element type is processed
  if {$entType == $feaLastEntity} {

# mesh
    if {[info exists feaFileName(mesh)] && [info exists feaMeshIndex]} {
      if {[file exists $feaFileName(mesh)]} {
        if {[info exists feaFile(mesh)]} {
          catch {close $feaFile(mesh)}
          if {[file size $feaFileName(mesh)] > 0} {
            puts $x3dFile "<!-- WIREFRAME -->\n<Switch whichChoice='0' id='swMesh'><Group>"
            set feaFile(mesh) [open $feaFileName(mesh) r]
            while {[gets $feaFile(mesh) line] >= 0} {puts $x3dFile $line}
            puts $x3dFile "</Group></Switch>\n"
            close $feaFile(mesh)
            foreach f {mesh meshIndex} {
              if {[file exists $feaFileName($f)]} {
                catch {close $feaFile($f)}
                catch {file delete -force -- $feaFileName($f)}
              }
            }
            unset feaMeshIndex
          }
        }
      }
    }

# loads (displacements), bcs, elements (order is important)
    foreach f {loads bcs elements} {
      if {[info exists feaFileName($f)]} {
        if {[file exists $feaFileName($f)]} {
          close $feaFile($f)
          if {[file size $feaFileName($f)] > 0} {
            set feaFile($f) [open $feaFileName($f) r]
            while {[gets $feaFile($f) line] >= 0} {puts $x3dFile $line}
            close $feaFile($f)
            if {$f == "elements"} {outputMsg " Finished writing FEM" green}
          }
        }
      }
    }

# delete temp files
    foreach f {bcs elements faceIndex loads} {
      if {[file exists $feaFileName($f)]} {
        catch {close $feaFile($f)}
        catch {file delete -force -- $feaFileName($f)}
      }
    }
    catch {unset feaFaceList}
    catch {unset feaFaceOrig}
  }

# messages
  if {[info exists feaTypes]} {
    foreach item [array names feaTypes] {lappend x3dMsg "$feaTypes($item) - [string tolower $item]"}
  }
}

# -------------------------------------------------------------------------------
proc feaEntities {objEntity} {
  global badAttributes elemID elemLoadValue elemLoadVariable elemLoadVec ent entAttrList entLevel feaBCNode feaBoundary feaDisp feaDispMag
  global feaDispNode feaDOFR feaDOFT feaElemFace feaEntity feaFaceIndex feaFaceList feaFile feaidx feaIndex feaLoad feaLoadMag feaLoadNode feaMeshIndex
  global feaNodes feaStateID feaType feaTypes firstID nnode nnodes nodeID opt surfaceNodes volumeNodes x3dMsg

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
    if {[string first "_3d_element_representation" $objType] != -1} {set elemID [$objEntity P21ID]}

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

              switch -glob $ent1 {
                "freedom_and_coefficient a" {
                  if {$objValue != 1} {
                    errorMsg "Unexpected freedom_and_coefficient 'a' attribute not equal to 1 ($objValue)"
                  }
                }
                "surface_3d_element_boundary_constant_specified_surface_variable_value simple_value" -
                "volume_3d_element_boundary_constant_specified_variable_value simple_value" {
                  if {$objValue != ""} {
                    set elemLoadValue $objValue
                  } else {
                    errorMsg "Syntax Error: Missing required 'simple_value' attribute on [lindex $ent1 0]"
                    catch {unset elemLoadValue}
                  }
                }
                "surface_3d_element_boundary_constant_specified_surface_variable_value variable" -
                "volume_3d_element_boundary_constant_specified_variable_value variable" {
                  if {$objValue != ""} {
                    set elemLoadVariable $objValue
                  } else {
                    errorMsg "Syntax Error: Missing required 'variable' attribute on [lindex $ent1 0]"
                    catch {unset elemLoadVariable}
                  }
                }
              }

# referenced entities
              if {[string first "handle" $objEntity] != -1} {feaEntities $objValue}
            }
          } emsg3]} {
            errorMsg "Error processing FEM ($objNodeType $ent2)\n $emsg3"
          }

# --------------
# nodeType = 20
        } elseif {$objNodeType == 20} {
          if {[catch {
            if {$entAttrOK != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $entLevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}

              switch -glob $ent1 {
                "*_element_representation node_list" {
                  set nnodes $objSize
                }

                "cartesian_point coordinates" {
# nodes for loads
                  if {$feaEntity == "nodal_freedom_action_definition" || $feaEntity == "element_nodal_freedom_actions"} {
                    lappend feaLoadNode [vectrim $objValue]

# nodes for displacements
                  } elseif {$feaEntity == "nodal_freedom_values"} {
                    lappend feaDispNode [vectrim $objValue]

# create node for surface load on surface element
                  } elseif {$feaEntity == "surface_3d_element_boundary_constant_specified_surface_variable_value"} {
                    lappend surfaceNodes $objValue
                    if {($nnodes <= 6 && [llength $surfaceNodes] == $nnodes) || ($nnodes == 9 && [llength $surfaceNodes] == 8)} {

# add all nodes together to get average node where load vector can be attached to
                      set cnode {0 0 0}
                      foreach n $surfaceNodes {set cnode [vecadd $cnode $n]}
                      set cnode [vecmult $cnode [expr {1./[llength $surfaceNodes]}]]
                      lappend feaLoadNode [vectrim $cnode]

# use cross product between two 'perpendicular' edges to get load vector for surface3d elements
                      set v1 [vecsub [lindex $surfaceNodes 1] [lindex $surfaceNodes 0]]
                      if {$nnodes <= 4} {
                        set v2 [vecsub [lindex $surfaceNodes [expr {$nnodes-1}]] [lindex $surfaceNodes 0]]
                      } elseif {$nnodes == 6} {
                        set v2 [vecsub [lindex $surfaceNodes 2] [lindex $surfaceNodes 0]]
                      } elseif {$nnodes == 9} {
                        set v2 [vecsub [lindex $surfaceNodes 3] [lindex $surfaceNodes 0]]
                      } elseif {$nnodes == 8 || $nnodes == 20 || $nnodes == 27} {

                      }
                      lappend elemLoadVec [vecnorm [veccross $v1 $v2]]
                      unset surfaceNodes
                    }

# save nodes for surface load on volume element
                  } elseif {$feaEntity == "volume_3d_element_boundary_constant_specified_variable_value"} {
                    lappend volumeNodes($elemID) $objValue

# nodes for bcs
                  } elseif {$feaEntity == "single_point_constraint_element_values"} {
                    lappend feaBCNode [vectrim $objValue]
                  }
                }

                "element_nodal_freedom_terms values" -
                "nodal_freedom_action_definition values" {
# loads
                  set n 0
                  set feaDOF [list $feaDOFT $feaDOFR]
                  if {[string first "unspecified" $objValue] == -1} {
                    foreach dof $feaDOF {
                      set ov $objValue
                      incr n
                      switch $n {
                        1 {
                          set type "force"
                          if {$feaEntity != "nodal_freedom_action_definition"} {set ov [join [lrange $objValue 0 2]]}
                        }
                        2 {
                          set type "moment"
                          if {$feaEntity != "nodal_freedom_action_definition"} {set ov [join [lrange $objValue 3 5]]}
                        }
                      }
                      if {$dof != ""} {
                        switch $dof {
                          x   {set lv "$ov 0. 0."}
                          xy  {set lv "$ov 0."}
                          xz  {set lv "[lindex $ov 0] 0. [lindex $ov 1]"}
                          xyz {set lv $ov}
                          y   {set lv "0. $ov 0."}
                          yz  {set lv "0. $ov"}
                          z   {set lv "0. 0. $ov"}
                        }
                        if {[info exists feaLoadNode]} {
                          foreach ld $feaLoadNode {
                            set lv [vectrim $lv]
                            if {$lv != "0.0 0.0 0.0"} {lappend feaLoad($feaStateID) "$ld,$lv,$type"}
                          }
                        }
                      }
                    }
                    catch {unset feaLoadNode}

                    if {![info exists feaLoadMag(min)]} {set feaLoadMag(min) 1.e+10}
                    if {![info exists feaLoadMag(max)]} {set feaLoadMag(max) -1.e+10}
                    set mag [veclen $objValue]
                    if {$mag < $feaLoadMag(min)} {set feaLoadMag(min) $mag}
                    if {$mag > $feaLoadMag(max)} {set feaLoadMag(max) $mag}
                  }
                }

                "nodal_freedom_values values" {
# displacements
                  if {[string first "unspecified" $objValue] == -1} {
                    foreach dof $feaDOFT {
                      set ov $objValue
                      if {$dof != ""} {
                        switch $dof {
                          x   {set lv "$ov 0. 0."}
                          xy  {set lv "$ov 0."}
                          xz  {set lv "[lindex $ov 0] 0. [lindex $ov 1]"}
                          xyz {set lv $ov}
                          y   {set lv "0. $ov 0."}
                          yz  {set lv "0. $ov"}
                          z   {set lv "0. 0. $ov"}
                        }
                        if {[info exists feaDispNode]} {
                          foreach ld $feaDispNode {
                            set lv [lrange $lv 0 2]
                            if {$lv != "0.0 0.0 0.0"} {lappend feaDisp($feaStateID) "$ld,$lv,displacement"}
                          }
                        }
                      }
                    }
                    catch {unset feaDispNode}

                    set feaDispMag(min) 0.
                    if {![info exists feaDispMag(max)]} {set feaDispMag(max) -1.e+10}
                    set mag [veclen $objValue]
                    if {$mag > $feaDispMag(max)} {set feaDispMag(max) $mag}
                  }
                }

                "freedoms_list freedoms" {
                  set feaDOFT ""
                  set feaDOFR ""
                  foreach dof $objValue {
                    if {[string first "trans" $dof] != -1} {
                      append feaDOFT [string index $dof 0]
                    } else {
                      append feaDOFR [string index $dof 0]
                    }
                  }
                  if {$feaEntity == "single_point_constraint_element_values"} {
                    if {[info exists feaBCNode]} {
                      foreach bc $feaBCNode {lappend feaBoundary($feaStateID) "$bc,$feaDOFT,$feaDOFR"}
                      unset feaBCNode
                    }
                  }
                }
                "single_point_constraint_element_values b" {
                  set sum 0.
                  foreach item $objValue {set sum [expr {$sum+$item}]}
                  if {$sum != 0.} {
                    errorMsg "Unexpected single_point_constraint_element_values 'b' attribute not equal to zero ($objValue)"
                  }
                }
              }

# referenced entities
              if {[catch {
                ::tcom::foreach val1 $objValue {feaEntities $val1}
              } emsg]} {
                foreach val2 $objValue {feaEntities $val2}
              }
            }
          } emsg3]} {
            errorMsg "Error processing FEM ($objNodeType $ent2): $emsg3"
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
                  if {[string first "element_representation" $feaEntity] != -1} {
                    if {$objType == "node"} {
                      set nodeID $objID
                      set feaidx($nnode) $feaNodes($nodeID)
                      if {[info exists nnode]} {
                        incr nnode
                        if {$nnode == 1} {set firstID $feaNodes($nodeID)}
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

                "state state_id" -
                "calculated_state state_id" -
                "specified_state state_id" {
                  if {$objValue == ""} {errorMsg "Syntax Error: Missing required 'state_id' on [lindex $ent1 0]."}
                  set feaStateID $objValue
                }
                "state description" -
                "calculated_state description" -
                "specified_state description" {
                  if {$feaStateID == ""} {set feaStateID $objValue}
                }

                "surface_3d_element_boundary_constant_specified_surface_variable_value element_face" {
# surface load on surface element
                  if {[info exists elemLoadVec] && [info exists elemLoadValue]} {
                    set ltype "ppressure"
                    if {$elemLoadValue < 0.} {set ltype "npressure"}
                    if {$objValue < 1 || $objValue > 2} {errorMsg "Unexpected [lindex $ent1 0] element_face 'element_face' attribute ($objValue)"}
                    foreach node $feaLoadNode lvec $elemLoadVec {
                      set lv [vecmult $lvec $elemLoadValue]
                      if {$objValue == 1} {set lv [vecrev $lv]}
                      lappend feaLoad($feaStateID) "[join $node],$lv,$ltype"

                      if {![info exists feaLoadMag(min)]} {set feaLoadMag(min) 1.e+10}
                      if {![info exists feaLoadMag(max)]} {set feaLoadMag(max) -1.e+10}
                      set mag [trimNum [veclen $lv] 5]
                      if {$mag < $feaLoadMag(min)} {set feaLoadMag(min) $mag}
                      if {$mag > $feaLoadMag(max)} {set feaLoadMag(max) $mag}
                    }
                  }
                  unset feaLoadNode
                  unset elemLoadVec
                }

                "volume_3d_element_boundary_constant_specified_variable_value element_face" {
# surface load on volume element face
                  foreach id [array names volumeNodes] {
                    set nnodes [llength $volumeNodes($id)]
                    if {[info exists feaElemFace($nnodes,1)]} {

# add all nodes together to get average node where load vector can be attached to
                      set node {0 0 0}
                      foreach n $feaElemFace($nnode,$objValue) {set node [vecadd $node [lindex $volumeNodes($id) $n]]}
                      set node [vecmult $node [expr {1./4.}]]

                      set v1 [vecsub [lindex $volumeNodes($id) [lindex $feaElemFace($nnode,$objValue) 1]] [lindex $volumeNodes($id) [lindex $feaElemFace($nnode,$objValue) 0]]]
                      set v2 [vecsub [lindex $volumeNodes($id) [lindex $feaElemFace($nnode,$objValue) 2]] [lindex $volumeNodes($id) [lindex $feaElemFace($nnode,$objValue) 0]]]

                      set ltype "ppressure"
                      if {$elemLoadValue < 0.} {set ltype "npressure"}
                      set lv [vecmult [vecnorm [veccross $v2 $v1]] $elemLoadValue]
                      lappend feaLoad($feaStateID) "[join [vectrim $node]],$lv,$ltype"

                      if {![info exists feaLoadMag(min)]} {set feaLoadMag(min) 1.e+10}
                      if {![info exists feaLoadMag(max)]} {set feaLoadMag(max) -1.e+10}
                      set mag [trimNum [veclen $lv] 5]
                      if {$mag < $feaLoadMag(min)} {set feaLoadMag(min) $mag}
                      if {$mag > $feaLoadMag(max)} {set feaLoadMag(max) $mag}
                    } else {
                      errorMsg "Missing 'feaElemFace' values for $nnodes noded volume_3d elements."
                    }
                  }
                  unset volumeNodes
                }
              }
            }
          } emsg3]} {
            errorMsg "Error processing FEM ($objNodeType $ent2): $emsg3"
            set entLevel 1
          }
        }
      }
    }
  }
  incr entLevel -1
}

# -------------------------------------------------------------------------------
proc feaLoads {entType} {
  global dsScaleSwitch feaDisp feaDispMag feaFile feaLoad feaLoadMag ldScaleSwitch opt x3dAxesSize

# initialize
  catch {unset def}

# loads
  if {$entType != "nodal_freedom_values"} {
    set type "loads"
    foreach i [array names feaLoad] {set fld($i) $feaLoad($i)}
    set magMin $feaLoadMag(min)
    set magMax $feaLoadMag(max)
    foreach idx {force moment npressure ppressure} {
      set def($idx) -1
      set def(color$idx) {}
    }
    set ldScaleSwitch {}

# displacements
  } else {
    set type "displacements"
    foreach i [array names feaDisp] {set fld($i) $feaDisp($i)}
    set magMin $feaDispMag(min)
    set magMax $feaDispMag(max)
    set def(displacement) -1
    set def(colordisplacement) {}
    set dsScaleSwitch {}
  }
  outputMsg " Writing $type" green

  if {[catch {
    set size [expr {($x3dAxesSize*0.15)/0.12}]
    set range [expr {$magMax-$magMin}]
    set n 0

    puts $feaFile(loads) "\n<!-- [string toupper $type] -->"
    foreach load [lsort [array names fld]] {
      incr n
      puts $feaFile(loads) "<Switch whichChoice='-1' id='[string range $type 0 3]$n'><Group>"

      foreach fl $fld($load) {
        set fl [split $fl ","]
        set xyz [lindex $fl 0]
        set mag [veclen [lindex $fl 1]]
        set ltype [lindex $fl 2]
        set nsize [trimNum $size]

# vector scale
        if {$opt(feaLoadScale) && $type == "loads"} {
          if {$feaLoadMag(max) != 0} {
            set nsize [trimNum [expr {($mag/abs($feaLoadMag(max)))*$size}]]
            set nsize [trimNum [expr {$nsize*0.9+$size*0.1}]]
          }
        } elseif {$type == "displacements"} {
          if {$feaDispMag(max) != 0 && !$opt(feaDispNoTail)} {
            set nsize [trimNum [expr {($mag/abs($feaDispMag(max)))*$size}] 5]
          }
        }

# load vector color (https://www.particleincell.com/2014/colormap/)
        if {$range != 0} {
          set s [expr {($mag-$magMin)/$range}]
          set a [expr {(1.-$s)/.25}]
          set x [expr {floor($a)}]
          set y [trimNum [expr {$a-$x}]]
          set x [expr {int($x)}]
          switch -- $x {
            0 {set r 1; set g $y; set b 0}
            1 {set r [trimNum [expr {1-$y}]]; set g 1; set b 0}
            2 {set r 0; set g 1; set b $y}
            3 {set r 0; set g [trimNum [expr {1-$y}]]; set b 1}
            4 {set r 0; set g 0; set b 1}
          }
        } else {
          set r 1; set g 0; set b 0
        }

# load vector rotation
        set rot1 ""
        set rot "1 0 0 0"
        set vec [vectrim [vecnorm [lindex $fl 1]]]

# default vector direction is 1 0 0
        switch -- $vec {
          "-1. 0. 0." {set rot "0 1 0 3.1416"}
          "0. 1. 0."  {set rot "0 0 1 1.5708"}
          "0. -1. 0." {set rot "0 0 1 -1.5708"}
          "0. 0. 1."  {set rot "0 1 0 -1.5708"}
          "0. 0. -1." {set rot "0 1 0 1.5708"}
          default {
# arbitrary rotation
            set lvec [split $vec " "]

# rotation in xy plane
            set ang [trimNum [expr {atan2([lindex $lvec 1],[lindex $lvec 0])}]]
            set rot "0 0 1 $ang"
            set vz [lindex $lvec 2]
# rotation also in z
            if {$vz != 0.} {
              set vec1 "[lindex $lvec 0] [lindex $lvec 1] 0"
              set ang1 [vecangle $vec $vec1]
              #set ang1 [trimNum [expr {acos([vecdot $vec $vec1] / ([veclen $vec]*[veclen $vec1]))}]]
              if {$vz > 0.} {set ang1 [expr {$ang1*-1.}]}
              set rot1 "[expr {-[lindex $lvec 1]}] [lindex $lvec 0] 0 $ang1"
            }
          }
        }
        if {$rot1 == ""} {
          set ldtxt "<Transform translation='$xyz' rotation='$rot'"
        } else {
          set ldtxt "<Transform translation='$xyz' rotation='$rot1'><Transform rotation='$rot'"
        }
        if {$nsize != 1.} {append ldtxt " scale='$nsize $nsize $nsize'"}
        append ldtxt ">"

# tail and arrow head, double head for moment
        if {$ltype == "force" || $ltype == "displacement" || $ltype == "moment" || [string first "pressure" $ltype] != -1} {
          set p1 [lsearch $def(color$ltype) "$r $g $b"]
          if {$p1 == -1} {
            incr def($ltype)
            set id "scale[string totitle $ltype]$def($ltype)"
            append ldtxt "\n <Transform id='$id' DEF='$ltype$def($ltype)'>"
            append ldtxt [feaArrow $r $g $b $ltype $def($ltype)]
            append ldtxt "\n </Transform>\n</Transform>"
            if {$ltype != "displacement"} {
              lappend ldScaleSwitch $id
            } else {
              lappend dsScaleSwitch $id
            }
            lappend def(color$ltype) "$r $g $b"
          } else {
            append ldtxt "<Transform USE='$ltype$p1'></Transform></Transform>"
          }
        }
        if {$rot1 != ""} {append ldtxt "</Transform>"}
        puts $feaFile(loads) $ldtxt
      }
      puts $feaFile(loads) "</Group></Switch>\n"
      update idletasks
    }
  } emsg]} {
    errorMsg "Error processing FEM $type: $emsg"
  }
}

# -------------------------------------------------------------------------------
proc feaArrow {r g b type num} {
  global opt

  set t1 -1
  set t2 0
  set h1 -.2
  switch -- $type {
    npressure -
    displacement {
      if {!$opt(feaDispNoTail) || $type == "npressure"} {
        set t1 0
        set t2 1
        set h1 [expr {$t2+$h1}]
      } else {
        set t2 .2
        set h1 0
      }
    }
    moment {
      set t1 -1.05
      set h1 -.19
    }
  }
  set arrow \n

# tail
  if {!$opt(feaDispNoTail) || $type != "displacement"} {
    append arrow "  <Shape><Appearance><Material emissiveColor='$r $g $b'/></Appearance>"
    append arrow "<IndexedLineSet coordIndex='0 1 -1'><Coordinate point='$t1 0 0 $t2 0 0'/></IndexedLineSet></Shape>"
  }

# head
  append arrow "\n  <Shape"
  if {$type == "moment"} {append arrow " DEF='head$num'"}
  append arrow ">"
  append arrow "<Appearance><Material diffuseColor='$r $g $b'/></Appearance>"
  append arrow "<IndexedFaceSet coordIndex='0 1 2 -1 0 2 3 -1 0 3 4 -1 0 4 1 -1' solid='FALSE'><Coordinate point='$t2 0 0 $h1 .1 0 $h1 0 .1 $h1 -.1 0 $h1 0 -.1'/></IndexedFaceSet></Shape>"

# second head
  if {$type == "moment"} {append arrow "\n  <Transform translation='-.1 0 0'><Shape USE='head$num'></Shape></Transform>"}

  return $arrow
}

# -------------------------------------------------------------------------------
proc feaBCs {entType} {
  global bcScaleSwitch feaBoundary feaFile x3dAxesSize

  outputMsg " Writing boundary conditions" green
  if {[catch {
    set size [trimNum [expr {$x3dAxesSize*0.3}]]
    set crd(x) "$size 0 0"
    set crd(y) "0 $size 0"
    set crd(z) "0 0 $size"
    set clr(x) "1 0 0"
    set clr(y) "0 .5 0"
    set clr(z) "0 0 1"
    set defUSEt(x) 0
    set defUSEt(y) 0
    set defUSEt(z) 0
    set defUSEt(xyz) 0
    set defUSEr(x) 0
    set defUSEr(y) 0
    set defUSEr(z) 0
    set defUSEr(xyz) 0
    set bcScaleSwitch {}

    set ns 24
    set angle 0
    set dlt [expr {6.28319/$ns}]
    for {set i 0} {$i < $ns} {incr i} {append circleIndex "$i "}
    append circleIndex "0 -1 "

    set n 0
    set bctxt "<!-- BOUNDARY CONDITIONS -->"
    foreach spc [lsort [array names feaBoundary]] {
      incr n
      append bctxt "\n<Switch whichChoice='-1' id='spc$n'><Group>"

      foreach fbc $feaBoundary($spc) {
        set fbc [split $fbc ","]
        set xyz [lindex $fbc 0]
        set bctrn [lindex $fbc 1]
        set bcrot [lindex $fbc 2]
        if {[string length $bctrn] > 0 || [string length $bcrot] > 0} {
          append bctxt "\n<Transform translation='$xyz'>"

# fixed (box) all six DOF
          if {[string length $bctrn] == 3 && [string length $bcrot] == 3} {
            if {$defUSEt($bctrn) == 0} {
              append bctxt "\n <Group DEF='BCfixed'><Transform id='BCfixedScale'>"
              append bctxt "<Shape><Appearance><Material diffuseColor='.7 .7 .7'/></Appearance><Box size='$size $size $size'></Box></Shape>"
              append bctxt "</Transform></Group>\n"
              incr defUSEt(xyz)
              lappend bcScaleSwitch "BCfixedScale"
            } else {
              append bctxt "<Group USE='BCfixed'/>"
            }
          } else {

# translation
            if {[string length $bctrn] > 0} {
              if {[string length $bctrn] == 3} {

# pinned (pyramid) all three DOF
                if {$defUSEt($bctrn) == 0} {
                  set num(2) $size
                  set num(1) [expr {$size*0.5}]
                  append bctxt "\n <Group DEF='BCT$bctrn'><Transform id='BCTxyzScale'>"
                  append bctxt "<Shape><Appearance><Material diffuseColor='.7 .7 .7'/></Appearance><IndexedFaceSet coordIndex='0 1 2 -1 0 2 3 -1 0 3 4 -1 0 4 1 -1 1 4 3 2 -1'><Coordinate point='0 0 0 $num(1) $num(1) -$num(2) -$num(1) $num(1) -$num(2) -$num(1) -$num(1) -$num(2) $num(1) -$num(1) -$num(2)'/></IndexedFaceSet></Shape>"
                  append bctxt "</Transform></Group>\n"
                  incr defUSEt(xyz)
                  lappend bcScaleSwitch "BCTxyzScale"
                } else {
                  append bctxt "<Group USE='BCT$bctrn'/>"
                }

# other translation DOF constraints
              } else {
                for {set j 0} {$j < [string length $bctrn]} {incr j} {
                  set t [string index $bctrn $j]
                  if {$defUSEt($t) == 0} {
                    append bctxt "\n <Group DEF='BCT$t'><Transform id='BCT$t\Scale'>"
                    append bctxt "<Shape><Appearance><Material emissiveColor='$clr($t)'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0 0 0 $crd($t)'/></IndexedLineSet></Shape>"
                    append bctxt "</Transform></Group>\n"
                    incr defUSEt($t)
                    lappend bcScaleSwitch "BCT$t\Scale"
                  } else {
                    append bctxt "<Group USE='BCT$t'/>"
                  }
                }
              }
            }

# rotation
            if {[string length $bcrot] > 0} {
              if {[string length $bcrot] == 3} {

# fixed rotation (sphere) all three DOF
                if {$defUSEr($bcrot) == 0} {
                  append bctxt "\n <Group DEF='BCR$bcrot'><Transform id='BCRxyzScale'>"
                  append bctxt "<Shape><Appearance><Material diffuseColor='.7 .7 .7'/></Appearance><Sphere radius='[expr {0.3*$size}]'></Sphere></Shape>"
                  append bctxt "</Transform></Group>\n"
                  incr defUSEr(xyz)
                  lappend bcScaleSwitch "BCRxyzScale"
                } else {
                  append bctxt "<Group USE='BCR$bcrot'/>"
                }

# other rotation DOF constraints
              } else {
                for {set j 0} {$j < [string length $bcrot]} {incr j} {
                  set r [string index $bcrot $j]
                  if {$defUSEr($r) == 0} {
                    set circlePoints ""
                    for {set i 0} {$i < $ns} {incr i} {
                      set x [trimNum [expr {0.3*$size*cos($angle)}]]
                      set y [trimNum [expr {0.3*$size*sin($angle)}]]
                      switch -- $r {
                        x {append circlePoints "0 $x $y "}
                        y {append circlePoints "$x 0 $y "}
                        z {append circlePoints "$x $y 0 "}
                      }
                      set angle [expr {$angle+$dlt}]
                    }
                    append bctxt "\n <Group DEF='BCR$r'><Transform id='BCR$r\Scale'>"
                    append bctxt "<Shape><Appearance><Material emissiveColor='$clr($r)'/></Appearance><IndexedLineSet coordIndex='$circleIndex'><Coordinate point='$circlePoints'/></IndexedLineSet></Shape>"
                    append bctxt "</Transform></Group>\n"
                    incr defUSEr($r)
                    lappend bcScaleSwitch "BCR$r\Scale"
                  } else {
                    append bctxt "<Group USE='BCR$r'/>"
                  }
                }
              }
            }
          }
          append bctxt "</Transform>"
        }
      }
      append bctxt "\n</Group></Switch>\n"
    }
    puts $feaFile(bcs) $bctxt
    update idletasks
  } emsg]} {
    errorMsg "Error processing FEM boundary conditions: $emsg"
  }
}

# -------------------------------------------------------------------------------
proc feaButtons {type} {
  global bcScaleSwitch dsScaleSwitch entCount feaBoundary feaDisp feaDispMag feaLoad feaLoadMag ldScaleSwitch opt x3dFile

# node, mesh, element checkboxes
  if {$type == 1} {
    puts $x3dFile "\n<!-- FEM buttons -->\n<input type='checkbox' checked onclick='togNodes(this.value)'/>Nodes<br>"
    if {[info exists entCount(surface_3d_element_representation)] || \
        [info exists entCount(volume_3d_element_representation)]}  {puts $x3dFile "<input type='checkbox' checked onclick='togMesh(this.value)'/>Mesh<br>"}
    if {[info exists entCount(curve_3d_element_representation)]}   {puts $x3dFile "<input type='checkbox' checked onclick='tog1DElements(this.value)'/>1D Elements<br>"}
    if {[info exists entCount(surface_3d_element_representation)]} {puts $x3dFile "<input type='checkbox' checked onclick='tog2DElements(this.value)'/>2D Elements<br>"}
    if {[info exists entCount(volume_3d_element_representation)]}  {puts $x3dFile "<input type='checkbox' checked onclick='tog3DElements(this.value)'/>3D Elements<br>"}

# boundary condition checkboxes
    if {[info exists feaBoundary] && $opt(feaBounds)} {
      puts $x3dFile "\n<!-- BC checkbox and slider -->\n<p>Boundary Conditions<br>"
      set n 0
      foreach spc [lsort [array names feaBoundary]] {
        incr n
        puts $x3dFile "<input type='checkbox' name='spc' id='SPC$n' onclick='togSPC(this.value)'/>$spc<br>"
      }
      puts $x3dFile "<input style='width:80px' type='range' min='-2' max='4' step='0.25' value='1' onchange='bcScale(this.value)'/> Scale"
    }

# loads
    if {[info exists feaLoad] && $opt(feaLoads)} {
      puts $x3dFile "\n<!-- LOAD checkbox and slider -->\n<p>"
      set n 0
      if {$feaLoadMag(min) != $feaLoadMag(max)} {
        puts $x3dFile "<table border=0 cellpadding=0 cellspacing=0><tr><td>Loads</td></tr>"
      } else {
        set val $feaLoadMag(max)
        if {$val >= 100.} {set val [trimNum $val 0]} else {set val [trimNum $val]}
        puts $x3dFile "<table border=0 cellpadding=0 cellspacing=0><tr><td>Load = <font color='red'>$val</font></td></tr>"
      }

# load checkboxes
      foreach load [lsort [array names feaLoad]] {
        incr n
        puts $x3dFile "<tr><td><input type='checkbox' name='load' id='LOAD$n' onclick='togLOAD(this.value)'/>$load</td></tr>"
      }

# load color scale
      feaColorScale $feaLoadMag(min) $feaLoadMag(max)
      puts $x3dFile "</table>"

# load scale slider
      puts $x3dFile "<input style='width:80px' type='range' min='-2' max='4' step='0.25' value='1' onchange='ldScale(this.value)'/> Scale"
    }
    catch {unset feaLoadMag}

# displacements
    if {[info exists feaDisp] && $opt(feaDisp)} {
      puts $x3dFile "\n<!-- DISPLACEMENT checkbox and slider -->\n<p>"
      set n 0
      if {$feaDispMag(min) != $feaDispMag(max)} {
        puts $x3dFile "<table border=0 cellpadding=0 cellspacing=0><tr><td>Displacements</td></tr>"
      } else {
        puts $x3dFile "<table border=0 cellpadding=0 cellspacing=0><tr><td>Load = <font color='red'>[format "%.3e" $feaDispMag(max)]</font></td></tr>"
      }

# displacement checkboxes
      foreach disp [lsort [array names feaDisp]] {
        incr n
        puts $x3dFile "<tr><td><input type='checkbox' name='DISP' id='DISP$n' onclick='togDISPLACEMENT(this.value)'/>$disp</td></tr>"
      }

# displacement color scale
      feaColorScale $feaDispMag(min) $feaDispMag(max)
      puts $x3dFile "</table>"

# displacement scale slider
      puts $x3dFile "<input style='width:80px' type='range' min='-2' max='4' step='0.25' value='1' onchange='dsScale(this.value)'/> Scale"
    }
    catch {unset feaDispMag}

# functions for BC switches and scale
  } elseif {$type == 2} {
    if {[info exists feaBoundary] && $opt(feaBounds)} {
      puts $x3dFile "\n<!-- BC switch -->\n<script>function togSPC(val)\{"
      puts $x3dFile "  for(var i=1; i<=[llength [array names feaBoundary]]; i++) \{"
      puts $x3dFile "    if (!document.getElementById('SPC' + i).checked) \{document.getElementById('spc' + i).setAttribute('whichChoice', -1);\} else \{document.getElementById('spc' + i).setAttribute('whichChoice', 0);\}"
      puts $x3dFile "  \}"
      puts $x3dFile "\}</script>"
      unset feaBoundary
      puts $x3dFile "\n<!-- BC scale -->\n<script>function bcScale(scale)\{"
      puts $x3dFile " if (scale < 1) {scale = 0.7 + scale*0.3;}"
      puts $x3dFile " nscale = new x3dom.fields.SFVec3f(scale,scale,scale);"
      foreach element $bcScaleSwitch {
        puts $x3dFile " document.getElementById('$element').setFieldValue('scale', nscale);"
      }
      puts $x3dFile "\}</script>"
    }

# functions for load toggle switch
    if {[info exists feaLoad] && $opt(feaLoads)} {
      puts $x3dFile "\n<!-- LOAD switch -->\n<script>function togLOAD(val)\{"
      puts $x3dFile "  for(var i=1; i<=[llength [array names feaLoad]]; i++) \{"
      puts $x3dFile "    if (!document.getElementById('LOAD' + i).checked) \{document.getElementById('load' + i).setAttribute('whichChoice', -1);\} else \{document.getElementById('load' + i).setAttribute('whichChoice', 0);\}"
      puts $x3dFile "  \}"
      puts $x3dFile "\}</script>"
      unset feaLoad
      puts $x3dFile "\n<!-- LOAD scale -->\n<script>function ldScale(scale)\{"
      puts $x3dFile " if (scale < 1) {scale = 0.7 + scale*0.3;}"
      puts $x3dFile " nscale = new x3dom.fields.SFVec3f(scale,scale,scale);"
      foreach element $ldScaleSwitch {
        puts $x3dFile " document.getElementById('$element').setFieldValue('scale', nscale);"
      }
      puts $x3dFile "\}</script>"
    }

# functions for displacement toggle switch
    if {[info exists feaDisp] && $opt(feaDisp)} {
      puts $x3dFile "\n<!-- DISPLACEMENT switch -->\n<script>function togDISPLACEMENT(val)\{"
      puts $x3dFile "  for(var i=1; i<=[llength [array names feaDisp]]; i++) \{"
      puts $x3dFile "    if (!document.getElementById('DISP' + i).checked) \{document.getElementById('disp' + i).setAttribute('whichChoice', -1);\} else \{document.getElementById('disp' + i).setAttribute('whichChoice', 0);\}"
      puts $x3dFile "  \}"
      puts $x3dFile "\}</script>"
      unset feaDisp
      puts $x3dFile "\n<!-- DISPLACEMENT scale -->\n<script>function dsScale(scale)\{"
      puts $x3dFile " if (scale < 1) {scale = 0.7 + scale*0.3;}"
      puts $x3dFile " nscale = new x3dom.fields.SFVec3f(scale,scale,scale);"
      foreach element $dsScaleSwitch {
        puts $x3dFile " document.getElementById('$element').setFieldValue('scale', nscale);"
      }
      puts $x3dFile "\}</script>"
    }
  }
}

# -------------------------------------------------------------------------------
proc feaColorScale {min max} {
  global x3dFile

  if {$min != $max} {
    puts $x3dFile "<tr><td><table border=0 cellpadding=0 cellspacing=0>"
    puts $x3dFile "<tr><td colspan=3><img src='https://www.nist.gov/sites/default/files/styles/480_x_480_limit/public/images/2021/11/01/red-blue-scale.png' alt='Red-blue color scale' width='197' height='21'></td></tr>"

    puts $x3dFile "<tr><td align='left' width='33%'><font color='blue'>[feaColorScaleNum $min]</font></td>"
    puts $x3dFile "<td align='middle' width='34%'><font color='green'>[feaColorScaleNum [expr {($min+$max)/2.}]]</font></td>"
    puts $x3dFile "<td align='right' width='33%'><font color='red'>[feaColorScaleNum $max]</font></td></tr>"
    puts $x3dFile "</table></td></tr>"
  }
}

# -------------------------------------------------------------------------------
proc feaColorScaleNum {num} {

  if {$num >= 100.} {
    set num [trimNum $num 0]
  } elseif {$num < 0.01 && $num != 0.} {
    set num [format "%.2e" $num]
    for {set i 0} {$i < 2} {incr i} {regsub -all {\-0} $num {-} num}
  } else {
    set num [trimNum $num]
    if {[string index $num end] == "."} {set num [string range $num 0 end-1]}
  }
  return $num
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
  global entCount feaNodes x3dFile x3dMax x3dMin x3dMsg

  catch {unset feaNodes}
  set nodeIndex -1
  outputMsg " Reading nodes ($entCount(node))" green

# start node output
  puts $x3dFile "\n<!-- NODES -->\n<Switch whichChoice='0' id='swNodes'>"
  puts $x3dFile " <Shape><Appearance><Material emissiveColor='0 0 1'/></Appearance>"
  puts $x3dFile "  <PointSet><Coordinate DEF='nodes' point='"

# get all nodes
  set n 0
  set nline ""
  ::tcom::foreach e0 [$objDesign FindObjects [string trim node]] {
    set p21id [$e0 P21ID]
    set a0 [[$e0 Attributes] Item [expr 2]]

    ::tcom::foreach e1 [$a0 Value] {
      set coordList [[[$e1 Attributes] Item [expr 2]] Value]
      set feaNodes($p21id) [incr nodeIndex]

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
  puts $x3dFile "  '/></PointSet></Shape>"
  puts $x3dFile "</Switch>"

  lappend x3dMsg "$entCount(node) - nodes"
  if {[info exists entCount(dummy_node)]} {lappend x3dMsg "$entCount(dummy_node) - dummy node"}
}

# -------------------------------------------------------------------------------
# write index to file
proc feaWriteIndex {ftyp file} {
  global feaFile feaFileName

  close $feaFile($ftyp)
  if {[file size $feaFileName($ftyp)] > 0} {
    set feaFile($ftyp) [open $feaFileName($ftyp) r]

    set n 0
    set nline ""
    while {[gets $feaFile($ftyp) line] >= 0} {
      append nline "[string trim $line] "
      incr n
      if {$n == 10} {
        puts $feaFile($file) $nline
        set n 0
        set nline ""
      }
    }
    if {[string length $nline] > 0} {puts $feaFile($file) $nline}

    close $feaFile($ftyp)
    catch {file delete -force -- $feaFileName($ftyp)}
  }
}
