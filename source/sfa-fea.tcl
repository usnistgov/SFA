proc feaModel {entType} {
  global objDesign
  global ent entAttrList entCount entLevel opt rowmax nprogBarEnts count localName mytemp sfaPID cadSystem
  global x3dFile x3dMin x3dMax x3dMsg x3dStartFile x3dFileName x3dAxesSize
  global feaType feaTypes feaElemTypes nfeaElem feaFile feaFileName feaFaceList feaFaceOrig feaBoundary feaLoad feaMeshIndex feaLoadMag

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
  set node [list node name items $cartesian_point]
  set node_group [list node_group nodes]
  set freedom [list freedoms_list freedoms]
  set single_point_constraint_element [list single_point_constraint_element \
                                        required_node $node $node_group freedoms_and_values [list freedom_and_coefficient freedom a]]

  set FEA(single_point_constraint_element_values) [list single_point_constraint_element_values \
                                                    defined_state [list specified_state state_id description] [list state state_id description] \
                                                    element $single_point_constraint_element \
                                                    degrees_of_freedom $freedom b]
  set FEA(nodal_freedom_action_definition) [list nodal_freedom_action_definition \
                                              defined_state [list specified_state state_id description] [list state state_id description] \
                                              node $node $node_group \
                                              degrees_of_freedom $freedom values]
  
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
    puts $x3dFile "\n<table><tr><td width='85%'>"
    
# x3d window size
    set height 800
    set width [expr {int($height*1.5)}]
    catch {
      set height [expr {int([winfo screenheight .]*0.85)}]
      set width [expr {int($height*[winfo screenwidth .]/[winfo screenheight .])}]
    }
    puts $x3dFile "<X3D id='x3d' showStat='false' showLog='false' x='0px' y='0px' width='$width' height='$height'>\n<Scene DEF='scene'>"
    #puts $x3dFile "<X3D id='someUniqueId' showStat='false' showLog='false' x='0px' y='0px' width='$width\px' height='$height\px'>\n<Scene DEF='scene'>"

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

  set skipElements 0
  set startent [lindex $FEA($entType) 0]
  set memMsg "<i>Some elements were not processed.</i>"
  
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    if {[$objEntity Type] == $startent && (!$skipElements || [string first "element_representation" $startent] == -1)} {
      if {$opt(XLSCSV) == "None"} {incr nprogBarEnts}
      
      if {$nfeaElem < 10000000} {
        if {[expr {$nfeaElem%2000}] == 0} {

# check memory and gracefully exit
          set mem [expr {[lindex [twapi::get_process_info $sfaPID -pagefilebytes] 1]/1048576}]
          if {$mem > 1600} {
            errorMsg "Insufficient memory to process all elements"
            if {[lsearch $x3dMsg $memMsg] == -1} {
              lappend x3dMsg $memMsg
              set skipElements 1
            }
            update idletasks
            #break
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

# ------------------------------------------------------------------------------------------------  
# 1D elements, get from mesh
  if {[string first "_3d" $startent] != -1} {
    puts $feaFile(elements) "\n<!-- [string toupper $feaType] ELEMENTS -->"
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
        puts $x3dFile "<!-- WIREFRAME -->\n<Switch whichChoice='0' id='swMesh'><Group>"
        set feaFile(mesh) [open $feaFileName(mesh) r]
        while {[gets $feaFile(mesh) line] >= 0} {puts $x3dFile $line}
        puts $x3dFile "</Group></Switch>"
        close $feaFile(mesh)
        unset feaMeshIndex
      }
    }
    
# write all elements (must be before loads and boundary conditions)
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
    if {[info exists feaLoad] && $entType == "nodal_freedom_action_definition"} {
      set size [expr {($x3dAxesSize*0.15)/0.12}]
      set range [expr {$feaLoadMag(max)-$feaLoadMag(min)}]
      set n 0
      set i 0
      puts $x3dFile "\n<!-- LOADS -->"
      foreach load [lsort [array names feaLoad]] {
        incr n
        puts $x3dFile "<Switch whichChoice='-1' id='load$n'><Group>"
        
        foreach fl $feaLoad($load) {
          set fl [split $fl ","]
          set xyz [lindex $fl 0]
          set mag [veclen [lindex $fl 1]]
          set nsize [trimNum $size]
          if {$feaLoadMag(max) != 0} {
            set nsize [trimNum [expr {($mag/abs($feaLoadMag(max)))*$size}]]
            set nsize [trimNum [expr {$nsize*0.9+$size*0.1}]]
          }
          
# load vector color (https://www.particleincell.com/2014/colormap/)
          if {$range != 0} {
            set s [expr {($mag-$feaLoadMag(min))/$range}]
            set a [expr {(1.-$s)/.25}]
            set x [expr {floor($a)}]
            set y [trimNum [expr {$a-$x}]]
            set x [expr {int($x)}]
            switch $x {
              0 {set r 1; set g $y; set b 0}
              1 {set r [trimNum [expr {1.-$y}]]; set g 1; set b 0}
              2 {set r 0; set g 1; set b $y}
              3 {set r 0; set g [trimNum [expr {1.-$y}]]; set b 1}
              4 {set r 0; set g 0; set b 1}
            }
          } else {
            set r 1; set g 0; set b 0
          }

# load vector rotation
          set rot "1 0 0 0"
          set rot1 ""
          set vec [vectrim [vecnorm [lindex $fl 1]]]
          switch $vec {
            "-1. 0. 0." {set rot "0 1 0 3.1416"}
            "0. 1. 0."  {set rot "0 0 1 1.5708"}
            "0. -1. 0." {set rot "0 0 1 -1.5708"}
            "0. 0. 1."  {set rot "0 1 0 -1.5708"}
            "0. 0. -1." {set rot "0 1 0 1.5708"}
            default {
              set lvec [split $vec " "]
              set ang [trimNum [expr {atan2([lindex $lvec 1],[lindex $lvec 0])}]]
              set rot "0 0 1 $ang"
              if {[string range $vec end-2 end] != " 0."} {
                set vec1 "[lindex $lvec 0] [lindex $lvec 1] 0"
                set ang1 [trimNum [expr {acos([vecdot $vec $vec1] / ([veclen $vec]*[veclen $vec1]))}]]
                set rot1 "[expr {-[lindex $lvec 1]}] [lindex $lvec 0] 0 $ang1"
              }
            }
          }
          if {$rot1 == ""} {
            set str "<Transform translation='$xyz' rotation='$rot' "
          } else {
            set str "<Transform translation='$xyz' rotation='$rot1'><Transform rotation='$rot' "
          }
          append str "scale='$nsize $nsize $nsize'>"
          if {$i == 0} {
            append str "\n <Shape>\n  <Appearance><Material emissiveColor='$r $g $b'></Material></Appearance>"
            append str "\n  <IndexedLineSet DEF='Arrow' coordIndex='0 1 -1 1 2 -1 1 3 -1 1 4 -1 1 5 -1 2 3 4 5 2 -1'><Coordinate point='-1 0 0 0 0 0 -.2 .1 0 -.2 0 .1 -.2 -.1 0 -.2 0 -.1'></Coordinate></IndexedLineSet>\n </Shape>\n</Transform>"
            if {$rot1 != ""} {append str "</Transform>"}
          } else {
            append str "<Shape><Appearance><Material emissiveColor='$r $g $b'></Material></Appearance><IndexedLineSet USE='Arrow'></IndexedLineSet></Shape></Transform>"
            if {$rot1 != ""} {append str "</Transform>"} 
          }
          puts $x3dFile $str
          incr i
        }
        puts $x3dFile "</Group></Switch>\n"
      }
    }

# write boundary conditions
    if {[info exists feaBoundary] && $entType == "single_point_constraint_element_values"} {
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
      set defUSEr(x) 0
      set defUSEr(y) 0
      set defUSEr(z) 0
      set defUSEr(xyz) 0
      set ns 24
      set angle 0
      set dlt [expr {6.28319/$ns}]
      for {set i 0} {$i < $ns} {incr i} {append circleIndex "$i "}
      append circleIndex "0 -1 "

      set n 0
      puts $x3dFile "<!-- BOUNDARY CONDITIONS -->"
      foreach spc [lsort [array names feaBoundary]] {
        incr n
        puts $x3dFile "<Switch whichChoice='-1' id='spc$n'><Group>"
        
        foreach fbc $feaBoundary($spc) {
          set fbc [split $fbc ","]
          set xyz [lindex $fbc 0]
          set bctrn [lindex $fbc 1]
          set bcrot [lindex $fbc 2]
          if {[string length $bctrn] > 0 || [string length $bcrot] > 0} {
            puts $x3dFile "<Transform translation='$xyz'><Group>"

# translation
            if {[string length $bctrn] > 0} {
              for {set j 0} {$j < [string length $bctrn]} {incr j} {
                set t [string index $bctrn $j]
                if {$defUSEt($t) == 0} {
                  puts $x3dFile " <Group DEF='BCT$t'><Transform id='bct$t\Scale'>"
                  puts $x3dFile "  <Shape><Appearance><Material emissiveColor='$clr($t)'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0 0 0 $crd($t)'></Coordinate></IndexedLineSet></Shape>"
                  puts $x3dFile " </Transform></Group>"
                  incr defUSEt($t)
                } else {
                  puts $x3dFile " <Group USE='BCT$t'></Group>"
                }
              }
            }
          
# rotation          
            if {[string length $bcrot] > 0} {
              if {[string length $bcrot] == 3} {

# fixed rotation
                if {$defUSEr($bcrot) == 0} {
                  puts $x3dFile " <Group DEF='BCR$bcrot'><Transform id='bcSphere\Scale'>"
                  puts $x3dFile "  <Shape><Appearance><Material diffuseColor='.7 .7 .7' transparency='.5'></Material></Appearance><Sphere radius='[expr {0.3*$size}]'></Sphere></Shape>"
                  puts $x3dFile " </Transform></Group>"
                  incr defUSEr(xyz)
                } else {
                  puts $x3dFile " <Group USE='BCR$bcrot'></Group>"
                }

# other
              } else {
                for {set j 0} {$j < [string length $bcrot]} {incr j} {
                  set r [string index $bcrot $j]
                  if {$defUSEr($r) == 0} {
                    set circlePoints ""
                    for {set i 0} {$i < $ns} {incr i} {
                      set x [trimNum [expr {0.3*$size*cos($angle)}]]
                      set y [trimNum [expr {-0.3*$size*sin($angle)}]]
                      switch $r {
                        x {append circlePoints "0 $x $y "}
                        y {append circlePoints "$x 0 $y "}
                        z {append circlePoints "$x $y 0 "}
                      }
                      set angle [expr {$angle+$dlt}]
                    }
                    puts $x3dFile "<Group DEF='BCR$r'><Transform id='bcr$r\Scale'>"
                    puts $x3dFile " <Shape><Appearance><Material emissiveColor='$clr($r)'></Material></Appearance><IndexedLineSet coordIndex='$circleIndex'><Coordinate point='$circlePoints'></Coordinate></IndexedLineSet></Shape>"
                    puts $x3dFile "</Transform></Group>"
                    incr defUSEr($r)
                  } else {
                    puts $x3dFile "<Group USE='BCR$r'></Group>"
                  }
                }
              }
            }
            puts $x3dFile "</Group></Transform>"
          }

# pyramid          
          #set i 0
          #set size [trimNum [expr {$x3dAxesSize*0.15}]]
          #if {$i == 0} {
          #  puts $x3dFile "<Transform translation='$xyz'><Group DEF='Pinned'>\n <Transform id='bcTransform'><Shape>"
          #  puts $x3dFile "  <Appearance><Material diffuseColor='1 0 0'></Material></Appearance>\n  <IndexedFaceSet coordIndex='0 1 2 -1 0 2 3 -1 0 3 4 -1 0 4 1 -1 1 4 3 2 -1'><Coordinate point='0 0 0 .1 .1 -.2 -.1 .1 -.2 -.1 -.1 -.2 .1 -.1 -.2'></Coordinate></IndexedFaceSet>\n </Shape></Transform>\n</Group></Transform>"
          #} else {
          #  puts $x3dFile "<Transform translation='$xyz'><Group USE='Pinned'></Group></Transform>"
          #}
          #incr i
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
proc feaButtons {type} {
  global x3dFile feaBoundary feaLoad feaLoadMag entCount
    
# node, mesh, element checkboxes
  if {$type == 1} {
    puts $x3dFile "\n<input type='checkbox' checked onclick='togNodes(this.value)'/>Nodes<br>"
    if {[info exists entCount(surface_3d_element_representation)] || \
        [info exists entCount(volume_3d_element_representation)]}  {puts $x3dFile "<input type='checkbox' checked onclick='togMesh(this.value)'/>Mesh<br>"}
    if {[info exists entCount(curve_3d_element_representation)]}   {puts $x3dFile "<input type='checkbox' checked onclick='tog1DElements(this.value)'/>1D Elements<br>"}
    if {[info exists entCount(surface_3d_element_representation)]} {puts $x3dFile "<input type='checkbox' checked onclick='tog2DElements(this.value)'/>2D Elements<br>"}
    if {[info exists entCount(volume_3d_element_representation)]}  {puts $x3dFile "<input type='checkbox' checked onclick='tog3DElements(this.value)'/>3D Elements<br>"}

# boundary and load radiobuttons
    if {[info exists feaBoundary]} {
      puts $x3dFile "\n<p>Boundary Conditions\n<br>"
      set n 0
      foreach spc [lsort [array names feaBoundary]] {
        incr n
        puts $x3dFile "<input type='checkbox' name='spc' id='SPC$n' onclick='togSPC(this.value)'/>$spc<br>"
      }
      if {$n > 0} {puts $x3dFile "<input style='width:80px' type='range' min='0.25' max='4' step='0.25' value='1' onchange='bcScale(this.value)'/> Scale"}
    }
    if {[info exists feaLoad]} {
      puts $x3dFile "\n<p>Load Range: <font color='blue'>[trimNum $feaLoadMag(min)]</font> - <font color='red'>[trimNum $feaLoadMag(max)]</font>\n<br>"
      set n 0
      foreach load [lsort [array names feaLoad]] {
        incr n
        puts $x3dFile "<input type='checkbox' name='load' id='LOAD$n' onclick='togLOAD(this.value)'/>$load<br>"
      }
      unset feaLoadMag
    }
    
# switches
  } elseif {$type == 2} {
    if {[info exists feaBoundary]} {
      puts $x3dFile "\n<script>function togSPC(val)\{"
      puts $x3dFile "  for(var i=1; i<=[llength [array names feaBoundary]]; i++) \{"
      puts $x3dFile "    if (!document.getElementById('SPC' + i).checked) \{"
      puts $x3dFile "     document.getElementById('spc' + i).setAttribute('whichChoice', -1);"
      puts $x3dFile "    \} else \{"
      puts $x3dFile "     document.getElementById('spc' + i).setAttribute('whichChoice', 0);"
      puts $x3dFile "    \}"
      puts $x3dFile "  \}"
      puts $x3dFile "\}</script>"
      unset feaBoundary
      puts $x3dFile "<script>function bcScale(scale)\{"
      puts $x3dFile " nscale = new x3dom.fields.SFVec3f(scale,scale,scale);"
      puts $x3dFile " document.getElementById('bctxScale').setFieldValue('scale', nscale);"
      puts $x3dFile " document.getElementById('bctyScale').setFieldValue('scale', nscale);"
      puts $x3dFile " document.getElementById('bctzScale').setFieldValue('scale', nscale);"
      puts $x3dFile " document.getElementById('bcrxScale').setFieldValue('scale', nscale);"
      puts $x3dFile " document.getElementById('bcryScale').setFieldValue('scale', nscale);"
      puts $x3dFile " document.getElementById('bcrzScale').setFieldValue('scale', nscale);"
      puts $x3dFile " document.getElementById('bcSphere').setFieldValue('scale', nscale);"
      puts $x3dFile "\}</script>"
    }

# functions for load buttons
    if {[info exists feaLoad]} {
      puts $x3dFile "\n<script>function togLOAD(val)\{"
      puts $x3dFile "  for(var i=1; i<=[llength [array names feaLoad]]; i++) \{"
      puts $x3dFile "    if (!document.getElementById('LOAD' + i).checked) \{"
      puts $x3dFile "     document.getElementById('load' + i).setAttribute('whichChoice', -1);"
      puts $x3dFile "    \} else \{"
      puts $x3dFile "     document.getElementById('load' + i).setAttribute('whichChoice', 0);"
      puts $x3dFile "    \}"
      puts $x3dFile "  \}"
      puts $x3dFile "\}</script>"
      unset feaLoad
    }
  }
}

# -------------------------------------------------------------------------------
proc feaElements {objEntity} {
  global badAttributes ent entAttrList entCount entLevel localName nistVersion opt
  global x3dFile x3dFileName x3dStartFile feaMeshIndex feaFaceIndex x3dMsg feaStateID feaEntity
  global feaidx feaIndex feaType feaTypes firstID nnode nnodes nodeID nfeaElem feaFile feaFaceList feaBoundary feaLoad feaLoadNode feaLoadMag
  global feaNodes feaDOFT feaDOFR

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

              switch -glob $ent1 {
                "freedom_and_coefficient a" {
                  if {$objValue != 1} {
                    errorMsg "Unexpected (freedom)(coefficient) 'a' attribute not equal to 1 ($objValue)"
                  }
                }
              }

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
                "*_element_representation node_list" {
                  set nnodes $objSize
                }
                "cartesian_point coordinates" {
                  if {$feaEntity == "nodal_freedom_action_definition" || $feaEntity == "single_point_constraint_element_values"} {
                    set feaLoadNode [vectrim $objValue]
                  }
                }
                "nodal_freedom_action_definition values" {
                  switch $feaDOFT {
                    x  {set objValue "$objValue 0. 0."}
                    xy {set objValue "$objValue 0."}
                    xz {set objValue "[lindex $objValue 0] 0. [lindex $objValue 1]"}
                    y  {set objValue "0. $objValue 0."}
                    yz {set objValue "0. $objValue"}
                    z  {set objValue "0. 0. $objValue"}
                  }
                  lappend feaLoad($feaStateID) "$feaLoadNode,[vectrim $objValue]"
                  if {![info exists feaLoadMag(min)]} {set feaLoadMag(min) 1.e+10}
                  if {![info exists feaLoadMag(max)]} {set feaLoadMag(max) -1.e+10}
                  set mag [veclen $objValue]
                  if {$mag < $feaLoadMag(min)} {set feaLoadMag(min) $mag}
                  if {$mag > $feaLoadMag(max)} {set feaLoadMag(max) $mag}
                  if {$feaDOFR != ""} {errorMsg "Moment loads are not displayed."}
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
                  if {$feaEntity == "single_point_constraint_element_values"} {lappend feaBoundary($feaStateID) "$feaLoadNode,$feaDOFT,$feaDOFR"}
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
                "specified_state state_id" {
                  set feaStateID $objValue
                }
                "state description" -
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
  global entCount feaNodes x3dMax x3dMin x3dFile x3dMsg
  catch {unset feaNodes}

  set nodeIndex -1
  outputMsg " Reading nodes ($entCount(node))" green

# start node output
  puts $x3dFile "\n<!-- NODES -->\n<Switch whichChoice='0' id='swNodes'>"
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
  puts $x3dFile "  '></Coordinate></PointSet></Shape>"
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
    catch {file delete -force $feaFileName($ftyp)}
  }
}
