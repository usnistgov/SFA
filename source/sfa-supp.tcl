# supplemental geometry
proc x3dSuppGeom {} {
  global cgrObjects developer maxxyz recPracNames skipEntities syntaxErr tessSuppGeomFile tessSuppGeomFileName trimVal x3dFile
  global objDesign
  if {![info exists objDesign]} {return}

  set size [trimNum [expr {$maxxyz*0.025}]]
  set tsize [trimNum [expr {$size*0.33}]]
  set debugSG 0

  outputMsg " Processing supplemental geometry" green
  puts $x3dFile "\n<!-- SUPPLEMENTAL GEOMETRY -->\n<Switch whichChoice='0' id='swSMG'><Group>"

# get all constructive_geometry_representation
  if {![info exists cgrObjects]} {set cgrObjects [$objDesign FindObjects [string trim constructive_geometry_representation]]}
  ::tcom::foreach e0 $cgrObjects {
    if {$debugSG} {outputMsg "[$e0 Type] [$e0 P21ID]" green; puts $x3dFile "<!-- CGR [$e0 P21ID] -->"}

# for SG in assemblies, find related shape_representation > item_defined_transformation > a2p3d > transform
    if {[catch {
      set cgrrs [$e0 GetUsedIn [string trim constructive_geometry_representation_relationship] [string trim rep_2]]
      ::tcom::foreach cgrr $cgrrs {
        set shapeRep [[[$cgrr Attributes] Item [expr 3]] Value]
        set rrwts [$shapeRep GetUsedIn [string trim representation_relationship_with_transformation_and_shape_representation_relationship] [string trim rep_1]]
        if {$debugSG} {outputMsg " [$cgrr Type] [$cgrr P21ID]\n  [$shapeRep Type] [$shapeRep P21ID] [[[$shapeRep Attributes] Item [expr 1]] Value]" red}

# count NRRWT, if none use rep_2
        set rep_2 0
        set nrrwt 0
        ::tcom::foreach rrwt $rrwts {incr nrrwt}
        if {$nrrwt == 0} {
          set rrwts [$shapeRep GetUsedIn [string trim representation_relationship_with_transformation_and_shape_representation_relationship] [string trim rep_2]]
          ::tcom::foreach rrwt $rrwts {incr nrrwt}
          if {$developer && $nrrwt > 0} {errorMsg " Using rep_2 on (RRWT)(SRR) to find item_defined_transformation" red; set rep_2 1}
        }

# get transform (assumes no nesting)
        if {$nrrwt > 0} {
          ::tcom::foreach rrwt $rrwts {
            set idt [[[$rrwt Attributes] Item [expr 5]] Value]
            switch -- $rep_2 {
              0 {set a2p3d [[[$idt Attributes] Item [expr 4]] Value]}
              1 {set a2p3d [[[$idt Attributes] Item [expr 3]] Value]}
            }
            if {$debugSG} {outputMsg "   [$idt Type] [$idt P21ID]\n    [$a2p3d Type] [$a2p3d P21ID]" red}

# transform
            set a2p3d [x3dGetA2P3D $a2p3d]
            set transform [x3dTransform [lindex $a2p3d 0] [lindex $a2p3d 1] [lindex $a2p3d 2] "supplemental geometry 'transform'"]
            lappend transforms $transform
            if {$debugSG} {outputMsg "     $transform" red}
          }
        }
      }
    } emsg]} {
      errorMsg " Error getting transform for supplemental geometry in an assembly: $emsg"
    }
    if {![info exists transforms]} {
      lappend transforms "<Transform>"
    } elseif {$debugSG} {
      outputMsg "    Transforms ([llength $transforms])" blue
    }

# process all items in a constructive_geometry_representation, considering transforms used for assemblies
    foreach tf $transforms {
      if {$tf != "<Transform>"} {
        set sgTransform "<Transform id='idt [$idt P21ID]'[string range $tf 10 end]"
        puts $x3dFile $sgTransform
        if {$debugSG} {outputMsg "     $sgTransform" blue}
      }

      set a1 [[$e0 Attributes] Item [expr 2]]
      ::tcom::foreach e2 [$a1 Value] {

        if {[catch {
          set ename [$e2 Type]
          if {$debugSG} {outputMsg "      $ename [$e2 P21ID]"}

          switch $ename {
            line -
            polyline {x3dSuppGeomLine $e2 $tsize $ename}
            circle -
            ellipse  {x3dSuppGeomCircle $e2 $tsize $ename}
            plane    {x3dSuppGeomPlane $e2 $size}
            cartesian_point     {x3dSuppGeomPoint $e2 $tsize}
            axis2_placement_3d  {x3dSuppGeomAxis $e2 $size $tsize}
            cylindrical_surface {x3dSuppGeomCylinder $e2 $tsize}

            trimmed_curve -
            composite_curve -
            geometric_curve_set {
              catch {unset trimVal}
              set trimmedCurves {}

# get trimmed curves
              if {[lsearch $skipEntities $ename] == -1} {
                if {$ename == "trimmed_curve"} {
                  lappend trimmedCurves $e2

# composite_curve -> composite_curve_segment -> trimmed_curve
                } elseif {$ename == "composite_curve"} {
                  ::tcom::foreach ccs [[[$e2 Attributes] Item [expr 2]] Value] {
                    lappend trimmedCurves [[[$ccs Attributes] Item [expr 3]] Value]
                  }

# geometric_curve_set -> list of trimmed_curve or composite_curve -> trimmed_curve
                } elseif {$ename == "geometric_curve_set"} {
                  set e3s [[[$e2 Attributes] Item [expr 2]] Value]
                  foreach e3 $e3s {
                    set ename1 [$e3 Type]
                    switch $ename1 {
                      line -
                      polyline {x3dSuppGeomLine $e3 $tsize $ename1}
                      circle -
                      ellipse {x3dSuppGeomCircle $e3 $tsize $ename1}
                      cartesian_point {x3dSuppGeomPoint $e3 $tsize}
                      trimmed_curve {lappend trimmedCurves $e3}
                      composite_curve {::tcom::foreach ccs [[[$e3 Attributes] Item [expr 2]] Value] {lappend trimmedCurves [[[$ccs Attributes] Item [expr 3]] Value]}}
                      default {
                        errorMsg " Supplemental geometry for '[formatComplexEnt $ename1]' in 'geometric_curve_set' is not supported."
                      }
                    }
                  }
                }
              } else {
                errorMsg " Supplemental geometry for '[formatComplexEnt $ename]' is not supported."
              }

# process trimmed curves collected from above
              foreach tc $trimmedCurves {

# trimming with values OK (do not delete the meaningless 'catch')
                set trimVal(1) [[[$tc Attributes] Item [expr 3]] Value]
                set trimVal(2) [[[$tc Attributes] Item [expr 4]] Value]
                catch {set tmp "[$trimVal(1) Type][$trimVal(2) Type]"}

                foreach idx [list 1 2] {
                  if {[llength $trimVal($idx)] == 2} {
                    if {[string is double [lindex $trimVal($idx) 0]]} {set trimVal($idx) [lindex $trimVal($idx) 0]}
                    if {[string is double [lindex $trimVal($idx) 1]]} {set trimVal($idx) [lindex $trimVal($idx) 1]}
                  }
                  if {[string first "handle" $trimVal($idx)] == -1} {
                    if {[expr {abs($trimVal($idx))}] > 1000.} {
                      set nval [trimNum [expr {10.*$trimVal($idx)/abs($trimVal($idx))}]]
                      errorMsg "Trim value [trimNum $trimVal($idx)] for a 'trimmed_curve' is very large, using $nval instead."
                      set trimVal($idx) $nval
                    }
                  }
                }

# line, polyline, circle, ellipse trimmed curves
                set e3 [[[$tc Attributes] Item [expr 2]] Value]
                set ename2 [$e3 Type]
                switch $ename2 {
                  line -
                  polyline {x3dSuppGeomLine $e3 $tsize $ename2}
                  circle -
                  ellipse {x3dSuppGeomCircle $e3 $tsize $ename2}
                  default {
                    errorMsg " Supplemental geometry for '[formatComplexEnt [$e3 Type]]' in 'trimmed_curve' is not supported."
                  }
                }
              }
            }

            shell_based_surface_model {
              set cylIDs {}
              set e3 [lindex [[[$e2 Attributes] Item [expr 2]] Value] 0]
              set e4s [[[$e3 Attributes] Item [expr 2]] Value]
              ::tcom::foreach e4 $e4s {
                set e5 [lindex [[[$e4 Attributes] Item [expr 3]] Value] 0]
                set ename5 [$e5 Type]
                switch $ename5 {
                  plane {x3dSuppGeomPlane $e5 $size}
                  cylindrical_surface {
                    if {[lsearch $cylIDs [$e5 P21ID]] == -1} {
                      lappend cylIDs [$e5 P21ID]
                      x3dSuppGeomCylinder $e5 $tsize
                    }
                  }
                  default {
                    errorMsg " Supplemental geometry for '[formatComplexEnt [$e5 Type]]' in 'shell_based_surface_model' is not supported."
                  }
                }
              }
            }

            default {
              if {$ename != "tessellated_shell" && $ename != "tessellated_wire"} {
                if {$ename == "direction"} {
                  set msg "Syntax Error: Supplemental geometry for '$ename' is not valid.  ($recPracNames(suppgeom), Sec. 4.2)"
                  errorMsg $msg
                  lappend syntaxErr(constructive_geometry_representation) [list [$e0 P21ID] "items" $msg]
                } else {
                  errorMsg " Supplemental geometry for '[formatComplexEnt $ename]' is not supported."
                }
              }
            }
          }

  # error
        } emsg]} {
          errorMsg " Error adding '$ename' Supplemental Geometry: $emsg"
        }
      }

# close transform
      if {$tf != "<Transform>"} {puts $x3dFile "</Transform>"}
    }
    catch {unset transforms}
  }

# check for tessellated edges that are supplemental geometry
  if {[info exists tessSuppGeomFile]} {
    close $tessSuppGeomFile
    if {[file size $tessSuppGeomFileName] > 0} {
      set f [open $tessSuppGeomFileName r]
      puts $x3dFile "<!-- TESSELLATED GEOMETRY that is SUPPLEMENTAL GEOMETRY -->"
      while {[gets $f line] >= 0} {puts $x3dFile $line}
      close $f
    }
    catch {file delete -force -- $tessSuppGeomFileName}
    unset tessSuppGeomFile
    unset tessSuppGeomFileName
  }
  puts $x3dFile "</Group></Switch>"
}

# -------------------------------------------------------------------------------
# supplemental geometry for axis
proc x3dSuppGeomAxis {e2 size tsize} {
  global axesDef spmiTypesPerFile viz x3dFile

  set e3 $e2
  set a2p3d [x3dGetA2P3D $e3]
  set transform [x3dTransform [lindex $a2p3d 0] [lindex $a2p3d 1] [lindex $a2p3d 2] "supplemental geometry 'axes'"]

# check for axis color
  set axisColor [lindex [x3dSuppGeomColor $e3 "axes"] 0]

# red, green, blue axes
  if {$axisColor == ""} {
    set id [lsearch $axesDef $size]
    if {$id != -1} {
      puts $x3dFile "$transform<Group USE='axes$id'></Group>"
    } else {
      lappend axesDef $size
      puts $x3dFile $transform
      puts $x3dFile " <Group DEF='axes[expr {[llength $axesDef]-1}]'><Shape><Appearance><Material emissiveColor='1 0 0'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. $size 0. 0.'/></IndexedLineSet></Shape>"
      puts $x3dFile " <Shape><Appearance><Material emissiveColor='0 .5 0'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. $size 0.'/></IndexedLineSet></Shape>"
      puts $x3dFile " <Shape><Appearance><Material emissiveColor='0 0 1'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. 0. $size'/></IndexedLineSet></Shape></Group>"
    }

# colored axes
  } else {
    set id [lsearch $axesDef "$size $axisColor"]
    if {$id != -1} {
      puts $x3dFile "$transform<Group USE='axes$id'></Group>"
    } else {
      lappend axesDef "$size $axisColor"
      set sz [trimNum [expr {$size*1.5}]]
      set tsize [trimNum [expr {$sz*0.33}]]
      puts $x3dFile $transform
      puts $x3dFile " <Group DEF='axes[expr {[llength $axesDef]-1}]'><Shape><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. $sz 0. 0.'/></IndexedLineSet><Appearance><Material emissiveColor='$axisColor'/></Appearance></Shape>"
      puts $x3dFile " <Shape><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. $sz 0.'/></IndexedLineSet><Appearance><Material emissiveColor='$axisColor'/></Appearance></Shape>"
      puts $x3dFile " <Shape><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. 0. $sz'/></IndexedLineSet><Appearance><Material emissiveColor='$axisColor'/></Appearance></Shape>"
      puts $x3dFile " <Transform translation='$sz 0 0' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='X'><FontStyle family='SANS'/></Text><Appearance><Material diffuseColor='$axisColor'/></Appearance></Shape></Billboard></Transform>"
      puts $x3dFile " <Transform translation='0 $sz 0' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='Y'><FontStyle family='SANS'/></Text><Appearance><Material diffuseColor='$axisColor'/></Appearance></Shape></Billboard></Transform>"
      puts $x3dFile " <Transform translation='0 0 $sz' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='Z'><FontStyle family='SANS'/></Text><Appearance><Material diffuseColor='$axisColor'/></Appearance></Shape></Billboard></Transform></Group>"
    }
  }

# text string color and name
  set nsize [trimNum [expr {$tsize*0.5}]]
  set tcolor "1 0 0"
  if {$axisColor != ""} {set tcolor $axisColor}
  set name [[[$e2 Attributes] Item [expr 1]] Value]
  if {$name != ""} {
    regsub -all "'" $name "" name
    puts $x3dFile " <Transform scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$name'><FontStyle family='SANS' justify='BEGIN'/></Text><Appearance><Material diffuseColor='$tcolor'/></Appearance></Shape></Billboard></Transform>"
  }
  puts $x3dFile "</Transform>"
  set viz(SUPPGEOM) 1
  lappend spmiTypesPerFile "supplemental geometry"
}

# -------------------------------------------------------------------------------
# supplemental geometry for point and the origin of a hole
proc x3dSuppGeomPoint {e2 tsize {thruHole ""} {holeName ""}} {
  global sphereDef spmiTypesPerFile viz x3dFile

# check for point color
  if {$thruHole != "" || $holeName != ""} {
    set pointColor "0 0 0"
  } else {
    set pointColor [lindex [x3dSuppGeomColor $e2 "point"] 0]
    if {$pointColor == ""} {set pointColor "0 0 0"}
  }

  if {[catch {

# get cartesian_point name attribute or use hole name
    set name [[[$e2 Attributes] Item [expr 1]] Value]
    if {$holeName != ""} {set name $holeName}
    set name [string trim $name]

# append THRU for thru holes
    if {$thruHole == 1} {
      if {[string length $name] > 0} {
        append name " (THRU)"
      } else {
        append name "THRU"
      }
    }
    set coord1 [[[$e2 Attributes] Item [expr 2]] Value]

# point is a black emissive sphere
    set id [lsearch $sphereDef "$tsize $pointColor"]
    if {$id != -1} {
      puts $x3dFile "<Transform translation='[vectrim $coord1]'><Shape USE='point$id'></Shape></Transform>"
    } else {
      lappend sphereDef "$tsize $pointColor"
      puts $x3dFile "<Transform translation='[vectrim $coord1]'><Shape DEF='point[expr {[llength $sphereDef]-1}]'><Sphere radius='[trimNum [expr {$tsize*0.05}]]'></Sphere><Appearance><Material diffuseColor='$pointColor' emissiveColor='$pointColor'/></Appearance></Shape></Transform>"
    }

# point name
    if {$name != ""} {
      set nsize [trimNum [expr {$tsize*0.5}]]
      puts $x3dFile " <Transform translation='[vectrim $coord1]' scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$name'><FontStyle family='SANS' justify='BEGIN'/></Text><Appearance><Material diffuseColor='$pointColor'/></Appearance></Shape></Billboard></Transform>"
    }
    set viz(SUPPGEOM) 1
    if {$thruHole == ""} {lappend spmiTypesPerFile "supplemental geometry"}
  } emsg]} {
    errorMsg "Error adding 'point' Supplemental Geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry for line, polyline
proc x3dSuppGeomLine {e3 tsize {type "line"}} {
  global spmiTypesPerFile trimVal viz x3dFile

# check for line color
  set lineColor [lindex [x3dSuppGeomColor $e3 $type] 0]
  if {$lineColor == ""} {set lineColor "1 0 1"}

  if {[catch {
    if {$type == "line"} {
      set e4 [[[$e3 Attributes] Item [expr 2]] Value]
      set coord1 [vectrim [[[$e4 Attributes] Item [expr 2]] Value]]
      set e5 [[[$e3 Attributes] Item [expr 3]] Value]
      set mag [[[$e5 Attributes] Item [expr 3]] Value]
      set e6 [[[$e5 Attributes] Item [expr 2]] Value]
      set dir [[[$e6 Attributes] Item [expr 2]] Value]
      set coord2 [vectrim [vecmult $dir $mag]]

# trim line
      if {[info exists trimVal(2)]} {

# trim with real number
        if {[string first "handle" $trimVal(2)] == -1} {
          if {$trimVal(1) != 0.} {set origin [vectrim [vecmult $dir [expr {$trimVal(1)*$mag}]]]}
          set coord2 [vectrim [vecmult $dir [expr {$trimVal(2)*$mag}]]]

# trim with cartesian points
        } else {
          foreach idx [list 1 2] {set trim($idx) [[[$trimVal($idx) Attributes] Item [expr 2]] Value]}
          set coord1 [vectrim $trim(1)]
          set coord2 [vectrim [vecsub $trim(2) $trim(1)]]
        }
      }

      set origin $coord1
      set coord2 [vectrim [vecadd $coord1 $coord2]]
      set points "$coord1 $coord2"
      set npoints 2

# polyline
    } else {
      set e4s [[[$e3 Attributes] Item [expr 2]] Value]
      set points ""
      set npoints 0
      ::tcom::foreach e4 $e4s {
        append points "[vectrim [[[$e4 Attributes] Item [expr 2]] Value]] "
        incr npoints
        if {$npoints == 1} {set origin $points}
      }
    }

# index
    set index ""
    for {set i 0} {$i < $npoints} {incr i} {append index "$i "}
    append index "-1"

# line geometry
    puts $x3dFile "<Shape><IndexedLineSet coordIndex='$index'><Coordinate point='$points'/></IndexedLineSet><Appearance><Material emissiveColor='$lineColor'/></Appearance></Shape>"

# line name at beginning
    set name [[[$e3 Attributes] Item [expr 1]] Value]
    if {$name != ""} {
      set nsize [trimNum [expr {$tsize*0.5}]]
      puts $x3dFile " <Transform translation='$origin' scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$name'><FontStyle family='SANS' justify='BEGIN'/></Text><Appearance><Material diffuseColor='$lineColor'/></Appearance></Shape></Billboard></Transform>"
    }
    set viz(SUPPGEOM) 1
    lappend spmiTypesPerFile "supplemental geometry"

  } emsg]} {
    errorMsg "Error adding '$type' Supplemental Geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry for circle, ellipse
proc x3dSuppGeomCircle {e3 tsize {type "circle"}} {
  global DTR spmiTypesPerFile trimVal viz x3dFile

  if {$tsize == 0} {
    set circleColor "0 0 1"

# check for circle color
  } else {
    set circleColor [lindex [x3dSuppGeomColor $e3 $type] 0]
    if {$circleColor == ""} {set circleColor "1 0 1"}
  }

  if {[catch {
    set e4 [[[$e3 Attributes] Item [expr 2]] Value]
    set rad [[[$e3 Attributes] Item [expr 3]] Value]

    set scale ""
    if {$type == "ellipse"} {
      set rad1 [[[$e3 Attributes] Item [expr 4]] Value]
      set sy [expr {$rad1/$rad}]
      set scale "1 $sy 1"
      set dsy [trimNum [expr {abs($sy-1.)}]]
      if {$dsy > 0. && $dsy <= 0.05} {errorMsg " Supplemental geometry $type axes ($rad,$rad1) are almost identical."}
    }

# circle position and orientation
    set a2p3d [x3dGetA2P3D $e4]
    set transform [x3dTransform [lindex $a2p3d 0] [lindex $a2p3d 1] [lindex $a2p3d 2] "supplemental geometry '$type'" $scale]
    puts $x3dFile $transform

# generate circle, account for trimming
# lim is the limit on an angle before deciding it is in degrees to convert to radians
    set ns 48
    set angle 0.
    set dlt [expr {6.28319/$ns}]
    set trimmed 0
    set lim 6.28319

# trim with angles, assume radians, convert degrees to radians
    if {[info exists trimVal(1)]} {
      if {[string first "handle" $trimVal(1)] == -1} {
        set angle $trimVal(1)
        set conv 1.
        if {[expr {abs($trimVal(1))}] > $lim || [expr {abs($trimVal(2))}] > $lim} {
          set conv $DTR
          set angle [expr {$angle*$conv}]
        }
        set dlt [expr {$conv*($trimVal(2)-$trimVal(1))/$ns}]
        incr ns
        set trimmed 1

# trim with cartesian points
      } else {

# compute angles from cartesian points (doesn't work yet)
        #foreach idx [list 1 2] {
        #  set trim($idx) [[[$trimVal($idx) Attributes] Item [expr 2]] Value]
        #  set vec($idx) [vecnorm [vecsub $trim($idx) $origin]]
        #  outputMsg "$idx / trim point [vectrim $trim($idx)] / origin [vectrim $origin] / vector [vectrim $vec($idx)] / axis [vectrim $axis] / refdir [vectrim $refdir]" red
        #  outputMsg "angle [vecangle $vec($idx) $axis]"
        #}
        if {$tsize == 0} {errorMsg " Trimming supplemental geometry '$type' with 'cartesian_point' is not supported."}
      }
    }
    set index ""
    for {set i 0} {$i < $ns} {incr i} {append index "$i "}
    if {!$trimmed} {append index "0 "}
    append index "-1"

    set coord ""
    for {set i 0} {$i < $ns} {incr i} {
      append coord "[trimNum [expr {$rad*cos($angle)}]] "
      append coord "[trimNum [expr {$rad*sin($angle)}]] "
      append coord "0 "
      set angle [expr {$angle+$dlt}]
      if {$i == 0} {set origin $coord}
    }

# circle geometry and possible name
    puts $x3dFile " <Shape><IndexedLineSet coordIndex='$index'><Coordinate point='$coord'/></IndexedLineSet><Appearance><Material emissiveColor='$circleColor'/></Appearance></Shape>"
    set name [[[$e3 Attributes] Item [expr 1]] Value]
    if {$name != "" && $tsize != 0} {
      set nsize [trimNum [expr {$tsize*0.5}]]
      puts $x3dFile " <Transform translation='$origin' scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$name'><FontStyle family='SANS' justify='BEGIN'/></Text><Appearance><Material diffuseColor='$circleColor'/></Appearance></Shape></Billboard></Transform>"
    }
    puts $x3dFile "</Transform>"
    set viz(SUPPGEOM) 1
    lappend spmiTypesPerFile "supplemental geometry"

  } emsg]} {
    errorMsg "Error adding '$type' Supplemental Geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry for plane, also used for section view clipping plane
proc x3dSuppGeomPlane {e2 size {type ""} {name ""}} {
  global clippingDef maxxyz planeDef spmiTypesPerFile viz x3dFile

# check for plane color and transparency
  set tmp [x3dSuppGeomColor $e2 "plane"]
  set planeColor [lindex $tmp 0]
  if {$planeColor == ""} {set planeColor "0 0 1"}
  set planeTrans [lindex $tmp 1]
  if {$planeTrans == ""} {set planeTrans 0.8}

  if {$type == ""} {
    set type "supplemental geometry"
  } else {
    set planeColor "1 1 1"
    set planeTrans 0.8
    set size [expr {$maxxyz*0.25}]
  }

  if {[catch {
    if {$name == ""} {set name [[[$e2 Attributes] Item [expr 1]] Value]}
    set e3 [[[$e2 Attributes] Item [expr 2]] Value]

# plane position and orientation
    set a2p3d [x3dGetA2P3D $e3]
    set transform [x3dTransform [lindex $a2p3d 0] [lindex $a2p3d 1] [lindex $a2p3d 2] "$type 'plane'"]

# plane geometry
    if {![info exists planeDef]} {set planeDef {}}
    set nsize [trimNum [expr {$size*2.}]]
    set id [lsearch $planeDef "$nsize $planeColor"]

# all supplemental geometry planes, only unique clipping planes
    set ok 1
    if {$type == "clipping plane"} {
      set cpdef "$name $transform"
      set planeColor "0 0 0"
      if {[lsearch $clippingDef $cpdef] != -1} {set ok 0}
    }
    if {$ok} {

# use DEF'd plane
      if {$id != -1} {
        puts $x3dFile "$transform<Group USE='plane$id'></Group>"
        if {$type == "supplemental geometry"} {lappend spmiTypesPerFile $type}
      } else {

# look for line corners
        set lp(2) ""
        set lp(3) ""
        set corners {}
        ::tcom::foreach e0 [$e2 GetUsedIn [string trim advanced_face] [string trim face_geometry]] {
          ::tcom::foreach e3 [[[$e0 Attributes] Item [expr 2]] Value] {
            set e4 [[[$e3 Attributes] Item [expr 2]] Value]
            ::tcom::foreach e5 [[[$e4 Attributes] Item [expr 2]] Value] {
              set e6 [[[$e5 Attributes] Item [expr 4]] Value]
              set e7 [[[$e6 Attributes] Item [expr 4]] Value]
              #outputMsg "0 [$e0 Type] [$e0 P21ID]\n3  [$e3 Type] [$e3 P21ID]\n4   [$e4 Type] [$e4 P21ID]\n5    [$e5 Type] [$e5 P21ID]\n6     [$e6 Type] [$e6 P21ID]\n7      [$e7 Type] [$e7 P21ID]"
              if {[$e7 Type] == "line"} {
                foreach i {2 3} {
                  set e8 [[[$e6 Attributes] Item [expr $i]] Value]
                  set e9 [[[$e8 Attributes] Item [expr 2]] Value]
                  set p($i) [vectrim [[[$e9 Attributes] Item [expr 2]] Value]]
                  #outputMsg "8 [$e8 Type] [$e8 P21ID]\n9  [$e9 Type] [$e9 P21ID]  $p($i)"
                }
                if {$p(2) != $lp(2) && $p(3) != $lp(3)} {
                  lappend corners $p(2)
                  lappend corners $p(3)
                } elseif {$p(2) == $lp(3)} {
                  lappend corners $p(3)
                } else {
                  lappend corners $p(2)
                  lappend corners $p(3)
                }
              } else {
                errorMsg " Bounding edges defined by '[formatComplexEnt [$e7 Type]]' for $type 'plane' are not supported."
              }
            }
          }

# fix order
          if {[llength $corners] == 8} {
            if {[lindex $corners 0] == [lindex $corners 2] && [lindex $corners 3] == [lindex $corners 4] && [lindex $corners 5] == [lindex $corners 7]} {
              set corners [list [lindex $corners 1] [lindex $corners 2] [lindex $corners 3] [lindex $corners 5]]
            }
          }

# use corners
          set ncorners [llength $corners]
          if {$ncorners > 2} {
            set txtpos [lindex $corners 0]
            set corners [join $corners " "]
            errorMsg " Using bounding edges for $type 'plane'" red
            if {$type == "supplemental geometry"} {lappend spmiTypesPerFile "bounded supplemental geometry"}
          } else {
            set corners {}
          }
        }

# no corners found
        if {[llength $corners] == 0} {
          set corners "-$nsize -$nsize 0. $nsize -$nsize 0. $nsize $nsize 0. -$nsize $nsize 0."
          set txtpos "-$nsize -$nsize 0."
          set ncorners 4
          lappend planeDef "$nsize $planeColor"
          if {$type == "supplemental geometry"} {lappend spmiTypesPerFile "supplemental geometry"}
        }

# set index
        set idx ""
        for {set i 0} {$i < $ncorners} {incr i} {append idx "$i "}
        set idx [string trim $idx]

        puts $x3dFile $transform
        set def ""
        if {[llength $planeDef] > 0} {set def " DEF='plane[expr {[llength $planeDef]-1}]'"}
        set str " <Group$def><Shape><IndexedLineSet coordIndex='$idx 0 -1'><Coordinate point='$corners'/></IndexedLineSet><Appearance><Material emissiveColor='$planeColor'/></Appearance></Shape>"
        if {$type != "clipping plane"} {
          append str "\n <Shape><IndexedFaceSet solid='false' coordIndex='$idx -1'><Coordinate point='$corners'/></IndexedFaceSet><Appearance><Material diffuseColor='$planeColor' transparency='$planeTrans'/></Appearance></Shape></Group>"
        } else {
          append str "</Group>"
        }
        puts $x3dFile $str
      }

# plane name at one corner
      set tsize [trimNum [expr {$size*0.33}]]
      if {$type != "supplemental geometry"} {
        set tsize [trimNum [expr {$tsize*0.25}]]
      }
      if {$name != ""} {
        if {![info exists txtpos]} {set txtpos "-$nsize -$nsize 0."}
        puts $x3dFile " <Transform translation='$txtpos' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$name'><FontStyle family='SANS' justify='BEGIN'/></Text><Appearance><Material diffuseColor='$planeColor'/></Appearance></Shape></Billboard></Transform>"
      }
      puts $x3dFile "</Transform>"
    }

    if {$type == "supplemental geometry"} {set viz(SUPPGEOM) 1}
    if {$type == "clipping plane"} {lappend clippingDef "$name $transform"}

  } emsg]} {
    errorMsg "Error adding 'plane' $type: $emsg"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry for cylinder
proc x3dSuppGeomCylinder {e2 size} {
  global viz x3dFile

# check for cylinder color and transparency
  set tmp [x3dSuppGeomColor $e2 "cylinder"]
  set cylColor [lindex $tmp 0]
  if {$cylColor == ""} {set cylColor "0 0 1"}
  set cylTrans [lindex $tmp 1]
  if {$cylTrans == ""} {set cylTrans 0.8}

  if {[catch {
    set e3 [[[$e2 Attributes] Item [expr 2]] Value]
    set rad [[[$e2 Attributes] Item [expr 3]] Value]

# cylinder position and orientation
    set a2p3d [x3dGetA2P3D $e3]
    set origin [lindex $a2p3d 0]
    set transform [x3dTransform $origin [lindex $a2p3d 1] [lindex $a2p3d 2] "supplemental geometry 'cylinder'"]
    puts $x3dFile "$transform<Transform rotation='1 0 0 1.5708'>"

# check if the cylinder is bounded by circles, get height by length between circle origins, better to compute with vertex_point
    set nco 0
    set circles {}
    ::tcom::foreach e0 [$e2 GetUsedIn [string trim advanced_face] [string trim face_geometry]] {
      ::tcom::foreach e3 [[[$e0 Attributes] Item [expr 2]] Value] {
        set e4 [[[$e3 Attributes] Item [expr 2]] Value]
        ::tcom::foreach e5 [[[$e4 Attributes] Item [expr 2]] Value] {
          set e6 [[[$e5 Attributes] Item [expr 4]] Value]
          foreach idx {2 3 4} {
            set e7 [[[$e6 Attributes] Item [expr $idx]] Value]
            if {[$e7 Type] == "circle"} {
              set e8 [[[$e7 Attributes] Item [expr 2]] Value]
              set e9 [[[$e8 Attributes] Item [expr 2]] Value]
              incr nco
              set co($nco) [[[$e9 Attributes] Item [expr 2]] Value]
              if {$nco == 2 && ![info exists height]} {
                set height 0
                foreach idx {0 1 2} {set height [expr {$height + [lindex [split $co(2) " "] $idx] - [lindex [split $co(1) " "] $idx]}]}
                if {$height == 0} {unset height}
              }
              lappend circles $e7
            } elseif {[$e7 Type] != "vertex_point" && [$e7 Type] != "line"} {
              errorMsg " Bounding edges '[formatComplexEnt [$e7 Type]]' for supplemental geometry 'cylindrical_surface' are not supported."
            }
          }
        }
      }
    }

# cylinder geometry
    if {![info exists height]} {set height [expr {$size*10.}]}
    puts $x3dFile "  <Shape><Cylinder radius='[trimNum $rad]' height='[trimNum $height]' top='false' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='$cylColor' transparency='$cylTrans'/></Appearance></Shape>"
    puts $x3dFile "</Transform></Transform>"

# cylinder name at origin
    set name [[[$e2 Attributes] Item [expr 1]] Value]
    if {$name != ""} {
      set tsize [trimNum [expr {$size*0.33}]]
      puts $x3dFile " <Transform translation='$origin' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$name'><FontStyle family='SANS' justify='BEGIN'/></Text><Appearance><Material diffuseColor='$cylColor'/></Appearance></Shape></Billboard></Transform>"
    }
    set viz(SUPPGEOM) 1

# circles at cylinder ends
    if {[llength $circles] == 0} {
      lappend spmiTypesPerFile "supplemental geometry"
    } else {
      lappend spmiTypesPerFile "bounded supplemental geometry"
      foreach circle $circles {x3dSuppGeomCircle $circle 0}
    }

  } emsg]} {
    errorMsg "Error adding 'cylinder' Supplemental Geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry color and transparency
proc x3dSuppGeomColor {e0 type} {
  global grayBackground

  set sgColor ""
  set sgTrans ""

# get color
  if {[catch {
    set e1s [$e0 GetUsedIn [string trim styled_item] [string trim item]]
    ::tcom::foreach e1 $e1s {
      set e2s [[[$e1 Attributes] Item [expr 2]] Value]
      ::tcom::foreach e2 $e2s {
        set e3 [[[$e2 Attributes] Item [expr 1]] Value]
        if {$e3 != "null"} {

# curve or point style
          if {[$e3 Type] == "curve_style" || [$e3 Type] == "point_style"} {
            set e4 [[[$e3 Attributes] Item [expr 4]] Value]
            if {$e4 != ""} {
              if {[$e4 Type] == "colour_rgb"} {
                set j 0
                ::tcom::foreach a4 [$e4 Attributes] {
                  if {$j > 0} {append sgColor "[trimNum [$a4 Value] 3] "}
                  incr j
                }
                set sgColor [string trim $sgColor]
              } elseif {[$e4 Type] == "draughting_pre_defined_colour"} {
                set sgColor [x3dPreDefinedColor [[[$e4 Attributes] Item [expr 1]] Value]]
              }
            }
            if {$type == "plane" || $type == "cylinder"} {errorMsg " Wrong type of style ([$e3 Type]) for a supplemental geometry '$type'."}

# surface style
          } elseif {[$e3 Type] == "surface_style_usage"} {
            set e4 [[[$e3 Attributes] Item [expr 2]] Value]
            set e5s [[[$e4 Attributes] Item [expr 2]] Value]
            foreach e5 $e5s {
              if {[$e5 Type] == "surface_style_fill_area"} {
                set e6 [[[$e5 Attributes] Item [expr 1]] Value]
                set e7s [[[$e6 Attributes] Item [expr 2]] Value]
                foreach e7 $e7s {
                  set e8 [[[$e7 Attributes] Item [expr 2]] Value]
                  if {$e8 != ""} {
                    set sgColor ""
                    if {[$e8 Type] == "colour_rgb"} {
                      set j 0
                      ::tcom::foreach a8 [$e8 Attributes] {
                        if {$j > 0} {append sgColor "[trimNum [$a8 Value] 3] "}
                        incr j
                      }
                      set sgColor [string trim $sgColor]
                    } elseif {[$e8 Type] == "draughting_pre_defined_colour"} {
                      set sgColor [x3dPreDefinedColor [[[$e8 Attributes] Item [expr 1]] Value]]
                    }
                  }
                }

# surface transparency
              } elseif {[$e5 Type] == "surface_style_rendering_with_properties"} {
                set e6s [[[$e5 Attributes] Item [expr 3]] Value]
                foreach e6 $e6s {set sgTrans [[[$e6 Attributes] Item [expr 1]] Value]}
              }
            }

            if {$type != "plane" && $type != "cylinder"} {errorMsg " Wrong type of style ([$e3 Type]) for a supplemental geometry '$type'."}
          } else {
            errorMsg " Color defined with '[$e3 Type]' referred from 'presentation_style_assignment' for supplemental geometry '$type' is not allowed."
          }
        }
      }
    }
  } emsg]} {
    errorMsg " Error getting color for '$type' Supplemental Geometry: $emsg"
  }

  foreach color [list "1. 1. 1." "1 1 1" "1. 1. 0." "1 1 0"] {if {$sgColor == $color} {set grayBackground 1}}
  return [list $sgColor $sgTrans]
}
