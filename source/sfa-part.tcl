# -------------------------------------------------------------------------------
# B-rep part geometry, new stp2x3d in SFA 5.10 also processes tessellated geometry
proc x3dBrepGeom {} {
  global brepFile brepFileName brepScale buttons cadSystem clippingCap developer DTR edgeMatID entCount localName
  global matTrans maxxyz mytemp nistVersion nsketch opt rawBytes rosetteGeom tessSolid viz x3dApps x3dBbox x3dMax x3dMin x3dMsg
  global x3dMsgColor x3dParts
  global objDesign

  if {[catch {
    if {$opt(debugX3D)} {getTiming x3dBrepGeom}

# get stp2x3d executable
    x3dCopySTP2X3D

# generate x3d from b-rep geometry with stp2x3d-part
    set stp2x3d [file join $mytemp stp2x3d-part.exe]
    if {[file exists $stp2x3d]} {

# output .x3d file name
      set stpx3dFileName [string range $localName 0 [string last "." $localName]]
      append stpx3dFileName "x3d"
      catch {file delete -force -- $stpx3dFileName}

# output dir for stp2x3d-part, delete if it exists
      set ftail [file tail [file rootname $localName]]
      set msg " Processing STEP part geometry for the Viewer"
      if {[info exists buttons]} {
        set fsize [file size $localName]
        if {$fsize > 200000000} {
          append msg ".  Please wait, it could take several minutes for large STEP files."
        } elseif {$fsize > 50000000} {
          append msg ", please wait."
        }
      }
      outputMsg $msg $x3dMsgColor
      if {$opt(brepAlt)} {outputMsg " Using alternative geometry processing, see More tab" red}

# check if invisibility is applied to surfaces
      if {[info exists entCount(invisibility)] && $opt(debugX3D)} {
        if {$entCount(invisibility) > 0} {
          ::tcom::foreach e0 [$objDesign FindObjects [string trim invisibility]] {
            catch {
              foreach e1 [[[$e0 Attributes] Item [expr 1]] Value] {
                if {[$e1 Type] == "styled_item"} {
                  set e2 [[[$e1 Attributes] Item [expr 3]] Value]
                  if {[$e2 Type] == "shell_based_surface_model"} {errorMsg " Surface invisibility is not supported" red}
                }
              }
            }
          }
        }
      }

# check for tessellated edges
      if {[info exists tessSolid] && $tessSolid} {
        if {[info exists entCount(tessellated_connecting_edge)]} {
          if {$entCount(tessellated_connecting_edge) > 0} {errorMsg " Tessellated edges are not supported" red}
        }
      }

# check for composite rosette curve_11
      set rosetteOpt  0
      set rosetteGeom 0
      foreach ent {composite_curve_and_curve_11 composite_curve_and_curve_11_and_measure_representation_item} {
        if {[info exists entCount($ent)]} {
          if {$entCount($ent) > 0} {
            set rosetteOpt 1
            if {[info exists objDesign] && $ent == "composite_curve_and_curve_11_and_measure_representation_item"} {
              ::tcom::foreach c11 [$objDesign FindObjects [string trim $ent]] {
                set id [$c11 P21ID]
                set curve11($id) [[[$c11 Attributes] Item [expr 4]] Value]
                set curve11($id) [trimNum [expr {$curve11($id)/$DTR}]]
              }
            }
          }
        }
      }

# check for clipping planes to cap
      set clippingCap 0
      if {$opt(partCap)} {
        set ent "camera_model_d3_multi_clipping"
        if {[info exists entCount($ent)]} {
          if {$entCount($ent) > 0 && $entCount($ent) < 17} {
            if {![info exists entCount(camera_model_d3_multi_clipping_intersection)] && \
                ![info exists entCount(camera_model_d3_multi_clipping_union)]} {set clippingCap 1}
          }
        }
      }

# run stp2x3d-part.exe
      if {$opt(debugX3D)} {getTiming stp2x3d}
      if {![info exists tessSolid]} {set tessSolid 1}
      if {$opt(debugX3D)} {
        outputMsg "--quality $opt(partQuality) --edge $opt(partEdges) --sketch $opt(partSketch) --normal $opt(partNormals) --rosette $rosetteOpt --tess $opt(brepAlt) --cap $clippingCap --tsolid $tessSolid"
      }
      catch {
        exec $stp2x3d --input [file nativename $localName] --quality $opt(partQuality) --edge $opt(partEdges) --sketch $opt(partSketch) --normal $opt(partNormals) --rosette $rosetteOpt --tess $opt(brepAlt) --cap $clippingCap --tsolid $tessSolid
      } errs
      if {$opt(debugX3D)} {getTiming done; outputMsg $errs}

# done processing
      if {[string first "STEP to X3D completed!" $errs] != -1} {
        if {[file exists $stpx3dFileName]} {
          if {[file size $stpx3dFileName] > 0} {
            set sketch 0
            set nind {}
            set x3dApps {}

# check for conversion units, mm > inch
            set brepScale [x3dBrepUnits]

# get min and max, number of materials, indents used to add Switch nodes
            set x3dBbox ""
            catch {unset indents}
            foreach line [split $errs "\n"] {
              if {[string first "No color will be supported." $line] != -1} {outputMsg "  Using gray for the part color" red}

              set sline [split [string trim $line] " "]
              if {[string first "MinXYZ" $line] != -1} {
                append x3dBbox "<br>Min:"
                foreach id1 {1 2 3} id2 {x y z} {
                  set num [expr {[lindex $sline $id1]}]
                  regsub -all "," $num "." num
                  set x3dMin($id2) [expr {$brepScale*$num}]
                  set prec 3
                  if {[expr {abs($num)}] >= 100.} {set prec 2}
                  set num [trimNum $num $prec]
                  if {[expr {abs($num)}] > 1.e8} {
                    set num "?"
                    errorMsg " Part min/max XYZ coordinate too small/large" red
                  }
                  append x3dBbox "&nbsp;&nbsp;$num"
                }
              } elseif {[string first "MaxXYZ" $line] != -1} {
                append x3dBbox "<br>Max:"
                foreach id1 {1 2 3} id2 {x y z} {
                  set num [expr {[lindex $sline $id1]}]
                  regsub -all "," $num "." num
                  set x3dMax($id2) [expr {$brepScale*$num}]
                  set prec 3
                  if {[expr {abs($num)}] >= 100.} {set prec 2}
                  set num [trimNum $num $prec]
                  if {[expr {abs($num)}] > 1.e8} {
                    set num "?"
                    errorMsg " Part min/max XYZ coordinate too small/large" red
                  }
                  append x3dBbox "&nbsp;&nbsp;$num"
                }
              } elseif {[string first "Number of Materials" $line] != -1} {
                set napps [string trim [string range $line [string last " " $line] end]]
                for {set i 0} {$i < $napps} {incr i} {lappend x3dApps $i}
              } elseif {[string first "Number of Rosettes" $line] != -1} {
                set rosetteGeom 1
              } elseif {[string first "indent" $line] != -1} {
                set indents([lindex $sline 1]) [lindex $sline 3]
              } elseif {[string first "Sketch geometry" $line] != -1} {
                set sketch 1

# STEP file errors
              } elseif {([string first "ERR StepFile" $line] != -1 || [string first "ERR StepReaderData" $line] != -1) && \
                         [string first "Fails Count : 1 " $line] == -1} {
                outputMsg $line red
                errorMsg " Use F8 to run the Syntax Checker to check for possible STEP file errors.  See Help > Syntax Checker" red

# other stp2x3d error messages
              } elseif {$developer && [string first "*" $line] == 0} {
                outputMsg $line red
              }

# coordinate min, max
              if {[info exists x3dMax(x)] && [info exists x3dMin(x)]} {
                foreach idx {x y z} {
                  if {$x3dMax($idx) == -1.e8 || $x3dMax($idx) > 1.e8} {set x3dMax($idx) 500.}
                  if {$x3dMin($idx) == 1.e8 || $x3dMin($idx) < -1.e8} {set x3dMin($idx) -500.}
                  set delt($idx) [expr {$x3dMax($idx)-$x3dMin($idx)}]
                }
                set maxxyz [expr {max($delt(x),$delt(y),$delt(z))}]
              }
            }
            if {$x3dBbox != ""} {set x3dBbox "Bounding Box$x3dBbox"}

# determine assembly level to insert Switch nodes
            foreach idx [lsort -integer [array names indents]] {lappend nind $indents($idx)}
            set ilast 1
            set lastLevel 1
            for {set i 0} {$i < [llength $nind]} {incr i} {
              set j $i
              set level [lindex $nind $i]
              if {$level < $lastLevel || ($level > 100 && $i > 1)} {
                set level $lastLevel
                set j [expr {$i-1}]
                break
              }
              set lastLevel $level
            }
            incr j
            set space "[string repeat " " $j]<"
            if {$opt(debugX3D)} {outputMsg "$nind\n$j $level" red}

# open temp file
            set brepFileName [file join $mytemp brep.txt]
            set brepFile [open $brepFileName w]

# integrate x3d from stp2x3d-part with existing x3dom file
            set str "\n<!-- PART GEOMETRY -->\n<Switch whichChoice='0' id='swPRT'>"
            if {$brepScale != 1} {
              append str "<Transform scale='$brepScale $brepScale $brepScale' onclick='handleGroupClick(event)'>"
            } else {
              append str "<Group onclick='handleGroupClick(event)'>"
            }
            puts $brepFile $str
            set stpx3dFile [open $stpx3dFileName r]
            set write 0
            set shape 0
            set npart(PRT) -1
            set nsketch -1
            set oksketch 0
            set lastline ""
            set close 0
            catch {unset parts}
            catch {unset matTrans}
            if {![info exists viz(EDGE)]} {set viz(EDGE) 0}

# process all lines in file
            if {$opt(debugX3D)} {getTiming start}
            outputMsg " Processing X3D output" $x3dMsgColor; update
            while {[gets $stpx3dFile line] >= 0} {
              if {$write} {
                if {[string first "Scene" $line] != -1} {set write 0}
              } elseif {[string first "Scene" $line] != -1} {
                gets $stpx3dFile line
                set write 1
              }
              if {!$shape} {
                if {[string first "Shape" $line] != -1} {set shape 1}
              }

# write
              if {$write} {

# Shape
                if {[string first "<Shape" $line] != -1} {

# check for edges
                  if {!$viz(EDGE)} {if {[string first "edge" $line] != -1} {set viz(EDGE) 1; set checkMatID 1}}

# check for composites curve 11
                  if {$rosetteGeom == 1} {
                    if {[string first "curve 11" $line] != -1} {
                      if {[catch {
                        set id11 [string range [lindex [split $line " "] 6] 0 end-2]
                        if {[info exists curve11($id11)]} {
                          puts $brepFile $line

# add label to end of curve
                          for {set i 0} {$i < 5} {incr i} {
                            gets $stpx3dFile line1
                            puts $brepFile $line1
                            if {[string first "Coordinate" $line1] != -1} {
                              set line1 [split $line1 " "]
                              set p11(x1) [string range [lindex $line1 6] 7 end]
                              set p11(y1) [lindex $line1 7]
                              set p11(z1) [lindex $line1 8]
                              set p11(x2) [lindex $line1 end-2]
                              set p11(y2) [lindex $line1 end-1]
                              set p11(z2) [string range [lindex $line1 end] 0 end-15]
                            }
                          }
                          set nsize [trimNum [expr {$maxxyz*0.02}]]
                          puts $brepFile "   <Transform translation='$p11(x1) $p11(y1) $p11(z1)' scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$curve11($id11)'><FontStyle family='SANS' justify='BEGIN'/></Text><Appearance><Material diffuseColor='0 0 0'/></Appearance></Shape></Billboard></Transform>"
                          puts $brepFile "   <Transform translation='$p11(x2) $p11(y2) $p11(z2)' scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$curve11($id11)'><FontStyle family='SANS' justify='BEGIN'/></Text><Appearance><Material diffuseColor='0 0 0'/></Appearance></Shape></Billboard></Transform>"
                          set line ""
                        }
                      } emsg2]} {
                        errorMsg "Error processing curve_11 for composite rosettes: $emsg2"
                      }
                    }
                  }

# add Switch for sketch geometry
                  if {$sketch} {
                    set c1 [string first "<Shape>" $line]
                    if {$c1 != -1} {
                      if {[string first "swComposite" $lastline] == -1} {
                        incr nsketch
                        set line "\n<!-- sketch geometry $nsketch -->\n[string repeat " " $c1]<Switch id='swSketch$nsketch' whichChoice='0'>[string range $line $c1 end]"
                        set oksketch 1
                      } else {
                        set sketch 0
                        set oksketch 0
                      }
                    }
                  }
                }

# check for mat id and transparency
                if {[string first "<Appearance DEF" $line] != -1} {
                  set c2 [string first "'mat" $line]
                  set id [string range $line $c2+4 $c2+7]
                  set id [string range $id 0 [string first "'" $id]-1]
                  if {[info exists checkMatID]} {set edgeMatID "mat$id"; unset checkMatID}

                  set c1 [string first "transparency=" $line]
                  if {$c1 != -1} {
                    set trans [string range $line $c1+14 end]
                    set trans [string range $trans 0 [string first "'" $trans]-1]
                    if {$trans == 1} {
                      set msg " Some surfaces are clear and not visible"
                      if {$opt(partEdges)} {
                        append msg " except for their edges"
                        if {[lsearch $x3dMsg [string trim $msg]] == -1} {lappend x3dMsg [string trim $msg]}
                      } else {
                        append msg ".  To view the surfaces, select 'Edges' in the Viewer section."
                      }
                      errorMsg $msg red
                    }
                    if {$trans > 0.} {
                      set matTrans($id) $trans
                      if {$trans != 1} {errorMsg " Some surfaces are transparent" red}
                    }
                  }
                }

# only check lines with correct number of spaces at beginning of line
                if {[string first $space $line] == 0} {

# close the Switch-Group
                  if {$close && [string first "$space/" $line] == 0} {
                    if {[string first "</Group></Switch>" $line] == -1} {append line "\n$space\/Group></Switch>"}
                    set close 0

# get id name of Transform and use for Switch
                  } elseif {[string first "Transform" $line] != -1} {
                    set c1 [string first "'" $line]
                    set c2 [string first "'" [string range $line $c1+1 end]]
                    if {$c2 == -1} {
                      errorMsg " Error reading a text string in the X3D file.  Parts might be missing in the the Viewer.\n$line"
                      append line "'>"
                      set c2 [string first "'" [string range $line $c1+1 end]]
                      lappend x3dMsg "Some part geometry might be missing"
                    }
                    incr npart(PRT)

# get id name based on there being a rotation= or translation= after the id of the Transform
                    set ct [string first "translation=" $line]
                    set cr [string first "rotation=" $line]
                    if {$ct != -1} {
                      set c2 $ct
                    } elseif {$cr != -1} {
                      set c2 $cr
                    } else {
                      set c2 [string length $line]
                    }
                    set id [string range $line $c1+1 $c2-3]
                    set cx [string first "\\X" $id]
                    if {$cx != -1} {set id [getUnicode $id]}
                    #if {$opt(debugX3D)} {outputMsg $id blue}
                    set parts($id) $npart(PRT)
                    set line "$space\Switch id='swPart$npart(PRT)' whichChoice='0'><Group>\n$line"
                    set close 1

# get DEF name of Group and use for Switch
                  } elseif {[string first "Group" $line] != -1} {
                    set close1 0
                    if {[string first "Group" $line] != [string last "Group" $line]} {
                      if {[string first "<Group" $line] != -1 && [string first "</Group>" $line] != -1} {set close1 1}
                    }
                    set c1 [string first "DEF" $line]
                    if {$c1 != -1} {
                      set c1 [expr {$c1+4}]
                    } else {
                      set c1 [string first "'" $line]
                    }
                    set c2 [string last "'" $line]
                    if {$c1 == $c2} {
                      errorMsg " Error reading a text string in the X3D file.  Parts might be missing in the the Viewer.\n$line"
                      append line "'>"
                      set c2 [string first "'" [string range $line $c1+1 end]]
                      lappend x3dMsg "Some part geometry might be missing"
                    }
                    incr npart(PRT)
                    set id [string range $line $c1+1 $c2-1]

# increment Group name _n
                    if {[info exists parts($id)]} {
                      if {$opt(debugX3D)} {outputMsg $id green}
                      for {set i 1} {$i < 99} {incr i} {
                        set c1 [string last "_" $id]
                        if {$c1 != -1} {
                          set nid "[string range $id 0 $c1]$i"
                        } else {
                          set nid "$id\_$i"
                        }
                        if {![info exists parts($nid)]} {
                          set id $nid
                          #if {$opt(debugX3D)} {outputMsg $id red}
                          break
                        }
                      }
                    }

# part switch
                    if {[string first "swSketch" $line] == -1 && [string first "swClipping" $line] == -1 && [string first "swComposites1" $line] == -1} {
                      set cx [string first "\\X" $id]
                      if {$cx != -1} {set id [getUnicode $id]}
                      set parts($id) $npart(PRT)
                      set line "$space\Switch id='swPart$npart(PRT)' whichChoice='0'><Group>\n$line"
                    }

# close group, switch
                    if {$close1} {
                      append line "\n$space\/Group></Switch>"
                    } else {
                      set close 1
                    }
                  }
                }

# end Switch for sketch geometry
                if {$oksketch} {
                  set c1 [string first "</Shape>" $line]
                  if {$c1 != -1} {
                    set line "$line</Switch>"
                    set oksketch 0
                  }
                }

# single quotes in quotes, change outer single quotes to double
                if {[string first "'" $line] != -1} {
                  if {[string first "<Group " $line] != -1 || [string first "<Shape " $line] != -1 || [string first "<Transform " $line] != -1} {
                    set n 0
                    set ok 0
                    set sline1 {}
                    set sline [split $line "="]
                    foreach str $sline {
                      if {$n > 0} {
                        set c1 [string last "'" $str]
                        set str1 [string range $str 1 $c1-1]
                        if {[string first "'" $str1] != -1} {
                          set str "\"$str1\"[string range $str $c1+1 end]"
                          set ok 1
                        }
                      }
                      lappend sline1 $str
                      incr n
                    }
                    if {$ok} {set line [join $sline1 "="]}
                  }
                }

# write line
                puts $brepFile $line
                set lastline $line
              }
            }

# check for duplicate part names in parts for x3dParts
            catch {unset x3dParts}
            if {[info exists parts]} {

# group duplicate parts
              if {!$opt(partNoGroup)} {
                foreach name [lsort [array names parts]] {
                  if {$opt(debugX3D)} {outputMsg "$name $parts($name)"}

# S control directive
                  if {[string first "\\S\\" $name] != -1} {errorMsg " The \\S\\ control directive is not supported for accented characters.  See Help > Text Strings and Numbers" red}

# check for _n at end of name
                  if {[string index $name end-1] == "_" || [string index $name end-2] == "_" || [string index $name end-3] == "_"} {

# remove _n
                    set c1 [string last "_" $name]
                    set name1 [string range $name 0 $c1-1]
                    if {$opt(debugX3D)} {outputMsg " $name1" red}

# add to x3dParts
                    if {[string range $name $c1 end] == "_1" && ![info exists parts($name1)]} {
                      set x3dParts($name1) $parts($name)
                    } else {
                      if {[lsearch [array names x3dParts] $name1] != -1} {
                        append x3dParts($name1) " $parts($name)"
                      } else {
                        set x3dParts($name) $parts($name)
                      }
                    }
                  } else {
                    set x3dParts($name) $parts($name)
                  }
                }

# do not group duplicate parts
              } else {
                foreach name [lsort [array names parts]] {set x3dParts($name) $parts($name)}
              }
            }

# check for non-English characters due to STEP file not having utf-8 encoding
            set okcad 1
            if {[info exists cadSystem]} {
              if {[string first "Autodesk" $cadSystem] != -1 || [string first "Siemens" $cadSystem] != -1} {set okcad 0}
            }
            if {[llength [array names x3dParts]] > 1 && $okcad} {
              set err 0
              foreach idx [array names x3dParts] {
                if {!$err} {
                  if {[string first "\;" $idx] != -1} {
                    set lnames [split $idx "\;"]
                    foreach name $lnames {
                      set c1 [string first "&" $name]
                      if {$c1 != -1} {set name [string range $name $c1 end]}
                      if {[string length $name] == 7 && [string range $name 0 2] == "&#5"} {
                        set err 1
                        break
                      }
                    }
                  } elseif {[regexp -all {[§¥Œ¿œ»º¤‡‰]} $idx] > 0 && [string first "Ð" $idx] == -1} {
                    set err 1
                    break
                  }
                }
              }
              if {$err && ![info exists rawBytes]} {
                set msg "The list of Assembly/Part names in the Viewer might use the wrong characters, possibly caused by the encoding of the STEP file"
                if {[info exists cadSystem]} {
                  if {[string first "SolidWorks" $cadSystem] != -1 && [string first "MBD" $cadSystem] == -1} {append msg " or improper characters on PRODUCT entities in the STEP file"}
                }
                append msg ".\nIf possible convert the encoding of the STEP file to UTF-8 with the Notepad++ text editor or other software.  See Text Strings and Numbers"
                errorMsg $msg
              }
            }
            if {$opt(debugX3D)} {foreach idx [array names x3dParts] {outputMsg "$idx $x3dParts($idx)" blue}}

# no shapes
            set viz(PART) 1
            if {!$shape} {
              set viz(PART) 0
              errorMsg " There is no geometry (Shape nodes) in the X3D file."
            }

# end the brep file
            if {$brepScale == 1} {
              puts $brepFile "</Group></Switch>"
            } else {
              puts $brepFile "</Transform></Switch>"
            }

            close $stpx3dFile
          }
          if {$opt(debugX3D)} {getTiming done}

# no X3D output
        } else {
          errorMsg " Cannot find the part geometry (X3D file) generated by stp2x3d-part.exe"
        }
        catch {file delete -force -- $stpx3dFileName}

# errors running stp2x3d
      } else {
        outputMsg $errs red

# missing Microsoft Visual C++ Redistributable
        if {$errs == "child killed: unknown signal"} {
          set msg "To process STEP part geometry, you have to install the Microsoft Visual C++ Redistributable.\n Follow the instructions in the SFA-README-FIRST.pdf included with the SFA zip file that you\n downloaded from the NIST website.  After the Redistributable is installed the Viewer should be\n able to process part geometry."

# other errors
        } elseif {[string first "No such file or directory" $errs] == 0} {
          set msg "Change the file or directory name and process the file again."
        } elseif {[string first "permission denied" $errs] != -1} {
          set msg "Antivirus software might be blocking stp2x3d-part.exe from running in $mytemp"
        } elseif {[info exists entCount(tessellated_brep_shape_representation)]} {
          if {$entCount(tessellated_brep_shape_representation)} {set msg "Polyhedral b-rep geometry (tessellated_brep_shape_representation) is not supported."}

# crash with STEP file
        } else {
          set msg "Error processing STEP part geometry."
          if {$tessSolid && !$opt(partOnly)} {append msg "\n Try the option for 'Alternative processing of tessellated geometry' (More tab)"}
          set ename "camera_model_d3_multi_clipping"
          if {[info exists entCount($ename)] && $opt(partCap)} {
            if {$entCount($ename) > 0} {append msg "\n Turn off generating capped surfaces for clipping planes' (More tab)"}
          }
          append msg "\n Try opening the file in another STEP viewer.  See Websites > STEP > STEP File Viewers"
          append msg "\n Use F8 to run the Syntax Checker to check for STEP file errors.  See Help > Syntax Checker"
        }
        if {[info exists msg]} {
          errorMsg $msg
          outputMsg " "
          lappend x3dMsg "Error generating STEP part geometry"
        }
      }
    } else {
      set msg " The program (stp2x3d-part.exe) to convert STEP part geometry to X3D was not found in $mytemp"
      if {!$nistVersion} {append msg "\n  You must first run the NIST version of the STEP File Analyzer and Viewer before generating a View."}
      errorMsg $msg
    }
  } emsg]} {
    errorMsg " Error generating Part Geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# check for conversion units, mm > inch
proc x3dBrepUnits {} {
  global objDesign
  if {![info exists objDesign]} {return 1}

  set scale 1.
  set a2 ""
  ::tcom::foreach e0 [$objDesign FindObjects [string trim advanced_brep_shape_representation]] {
    set e1 [[[$e0 Attributes] Item [expr 3]] Value]
    set a2 [[$e1 Attributes] Item [expr 5]]
    if {$a2 == ""} {set a2 [[$e1 Attributes] Item [expr 4]]}
  }
  if {$a2 == "" && ![info exists entCount(advanced_brep_shape_representation)]} {
    ::tcom::foreach e0 [$objDesign FindObjects [string trim geometric_representation_context_and_global_uncertainty_assigned_context_and_global_unit_assigned_context]] {
      set a2 [[$e0 Attributes] Item [expr 5]]
    }
  }
  if {$a2 != ""} {
    foreach e3 [$a2 Value] {
      if {[$e3 Type] == "conversion_based_unit_and_length_unit"} {
        set e4 [[[$e3 Attributes] Item [expr 3]] Value]
        set cf [[[$e4 Attributes] Item [expr 1]] Value]
        set scale [trimNum [expr {1./$cf}] 5]
      }
    }
  }
  return $scale
}

# -------------------------------------------------------------------------------
# copy stp2x3d files to temp directory, DLLs in sp2x3d-dll.zip, exe in stp2x3d-part.exe
proc x3dCopySTP2X3D {} {
  global nistVersion mytemp opt wdir

  if {!$nistVersion} {return}

  if {[catch {
    foreach fn {stp2x3d-dll.zip stp2x3d-part.exe} {
      set internal [file join $wdir exe $fn]
      set stp2x3d [file join $mytemp $fn]
      if {[file exists $internal]} {
        set copy 0
        if {![file exists $stp2x3d]} {
          set copy 1
        } elseif {[file mtime $internal] > [file mtime $stp2x3d]} {
          set copy 2
        }
        if {$copy > 0} {
          errorMsg " Copying Viewer software to [file nativename [file dirname $stp2x3d]]" red
          file copy -force -- $internal $stp2x3d
        }
      }
    }
  } emsg]} {
    errorMsg " Error copying stp2x3d-part.exe: $emsg"
  }

# extract DLLs from zip file
  set stp2x3dz [file join $mytemp stp2x3d-dll.zip]
  if {[file exists $stp2x3dz]} {
    if {[catch {
      vfs::zip::Mount $stp2x3dz stp2x3d-dll
      foreach file [glob -nocomplain stp2x3d-dll/*] {
        set fn [file join $mytemp [file tail $file]]
        set copy 0
        if {![file exists $fn]} {
          set copy 1
        } elseif {[file mtime $file] > [file mtime $fn]} {
          set copy 2
        }
        if {$copy > 0} {
          errorMsg " Copying Viewer software to $mytemp" red
          file copy -force -- $file $fn
        }
      }
      if {$opt(debugX3D)} {getTiming "copy and extract"}
    } emsg]} {
      errorMsg " Error extracting DLLs for stp2x3d-part.exe: $emsg"
    }
  }
}

# -------------------------------------------------------------------------------
# generate STEP AP242 tessellated geometry from STL file for viewing
proc STL2STEP {} {
  global localName

  catch {.tnb select .tnb.status}
  if {[catch {
    set f1 $localName
    set r1 [open $f1 r]

# check ascii or binary
    set aorb "Binary"
    for {set i 0} {$i < 9} {incr i} {
  		set line [string tolower [gets $r1]]
      if {[string first "facet" $line] != -1} {set aorb "ASCII"; break}
    }
    close $r1
    catch {unset r1}
    set r1 [open $f1 r]
    outputMsg "Generating STEP AP242 tessellated geometry from $aorb STL file" blue

# output file
    set f2 "[file rootname $f1]-stl.stp"
    file delete -force -- $f2
    set w2 [open $f2 w]

# header and entities
    puts $w2 "ISO-10303-21;
HEADER;
FILE_DESCRIPTION(('','CAx-IF Rec.Pracs.---3D Tessellated Geometry---1.1---2019-08-22'),'2;1');
FILE_NAME('[file tail $f1]','[clock format [clock seconds] -format "%Y-%m-%d\T%T"]',(' '),(' '),' ','NIST SFA [getVersion]',' ');
FILE_SCHEMA(('AP242_MANAGED_MODEL_BASED_3D_ENGINEERING_MIM_LF { 1 0 10303 442 3 1 4 }'));
ENDSEC;\n
DATA;
#1=APPLICATION_CONTEXT('managed model based 3d engineering') ;
#2=PRODUCT_CONTEXT(' ',#1,'mechanical') ;
#3=PRODUCT_DEFINITION_CONTEXT('part definition',#1,' ') ;
#4=APPLICATION_PROTOCOL_DEFINITION('international standard','ap242_managed_model_based_3d_engineering',2020,#1) ;
#5=PRODUCT('[file tail $f1]','[file tail $f1]','',(#2)) ;
#6=PRODUCT_DEFINITION_FORMATION_WITH_SPECIFIED_SOURCE('',' ',#5,.NOT_KNOWN.) ;
#7=PRODUCT_CATEGORY('part','specification') ;
#8=PRODUCT_RELATED_PRODUCT_CATEGORY('part',$,(#5)) ;
#9=PRODUCT_CATEGORY_RELATIONSHIP(' ',' ',#7,#8) ;
#10=PRODUCT_DEFINITION('[file tail $f1]',' ',#6,#3) ;
#11=PRODUCT_DEFINITION_SHAPE(' ',' ',#10) ;
#12=(LENGTH_UNIT()NAMED_UNIT(*)SI_UNIT(.MILLI.,.METRE.)) ;
#13=(NAMED_UNIT(*)PLANE_ANGLE_UNIT()SI_UNIT($,.RADIAN.)) ;
#14=(NAMED_UNIT(*)SI_UNIT($,.STERADIAN.)SOLID_ANGLE_UNIT()) ;
#15=UNCERTAINTY_MEASURE_WITH_UNIT(LENGTH_MEASURE(0.005),#12,'distance_accuracy_value','CONFUSED CURVE UNCERTAINTY') ;
#16=(GEOMETRIC_REPRESENTATION_CONTEXT(3)GLOBAL_UNCERTAINTY_ASSIGNED_CONTEXT((#15))GLOBAL_UNIT_ASSIGNED_CONTEXT((#12,#13,#14))REPRESENTATION_CONTEXT(' ',' ')) ;
#17=CARTESIAN_POINT(' ',(0.,0.,0.)) ;
#18=AXIS2_PLACEMENT_3D(' ',#17,$,$) ;
#19=SHAPE_REPRESENTATION(' ',(#18),#16) ;
#20=SHAPE_DEFINITION_REPRESENTATION(#11,#19) ;
#28=COLOUR_RGB('Colour',0.55,0.55,0.6) ;
#29=FILL_AREA_STYLE_COLOUR(' ',#28) ;
#30=FILL_AREA_STYLE(' ',(#29)) ;
#31=SURFACE_STYLE_FILL_AREA(#30) ;
#32=SURFACE_SIDE_STYLE(' ',(#31)) ;
#33=SURFACE_STYLE_USAGE(.BOTH.,#32) ;
#34=PRESENTATION_STYLE_ASSIGNMENT((#33)) ;
#35=STYLED_ITEM(' ',(#34),#102) ;
#92=SHAPE_REPRESENTATION_RELATIONSHIP(' ',' ',#19,#103) ;
#94=MECHANICAL_DESIGN_GEOMETRIC_PRESENTATION_REPRESENTATION(' ',(#35),#16) ;"

    set clist {}
    set index {}
    set nlist {}
    set cidx -1
    set midx -1
    set nface 0
    set id 100
    set lasttime [clock clicks -milliseconds]

# read ASCII stl file
    if {$aorb == "ASCII"} {
      set ntriangles 0
      while {[gets $r1 line] >= 0} {

# normals
        if {[string first "normal" $line] != -1} {
          incr ntriangles
          gets $r1 $line
          set idx "("
          for {set i 0} {$i < 3} {incr i} {
            gets $r1 line

# coordinates
            set c1 [expr {[string first "vertex" $line]+7}]
            set coord [string trim [string range $line $c1 end]]
            foreach c $coord {append cn "[trimNum $c],"}
            set coord "([string trim [string range $cn 0 end-1]]),"
            unset cn

# index
            if {![info exists clist1($coord)]} {
              lappend clist $coord
              set clist1($coord) [incr cidx]
              incr midx
              set jdx [expr {$midx+1}]
            } else {
              set jdx [expr {$clist1($coord)+1}]
            }
            set str $jdx
            if {$i < 2} {
              append str ","
            } else {
              append str "),"
            }
            lappend idx $str
          }
          incr nface
          lappend index $idx
        }
      }

# read binary STL (https://www.johann-oberdorfer.eu/blog/2018/01/12/18-01-12_stl_files_convert_from_ascii2binary/)
    } else {
      set ilen [GetByteLength "UINT32"]
      set rlen [expr {[GetByteLength "REAL"] * 3}]
      fconfigure $r1 -translation binary
      seek $r1 [expr {[GetByteLength "UINT8"] * 80}] start

# number of triangles
      binary scan [read $r1 $ilen] i ntriangles
      outputMsg " Reading [expr {$ntriangles*3}] coordinates, $ntriangles faces"
      for {set ntri 0} {$ntri < $ntriangles} {incr ntri} {

# normal
        binary scan [read $r1 $rlen] rrr normalX normalY normalZ
        set idx "("

# vertices
        for {set i 0} {$i < 3} {incr i} {
          binary scan [read $r1 $rlen] rrr X Y Z
          set coord "$X $Y $Z"
          foreach c $coord {append cn "[trimNum $c],"}
          set coord "([string trim [string range $cn 0 end-1]]),"
          unset cn

# index
          if {![info exists clist1($coord)]} {
            lappend clist $coord
            set clist1($coord) [incr cidx]
            incr midx
            set jdx [expr {$midx+1}]
          } else {
            set jdx [expr {$clist1($coord)+1}]
          }
          set str $jdx
          if {$i < 2} {
            append str ","
          } else {
            append str "),"
          }
          lappend idx $str
        }
        incr nface
        lappend index $idx
        seek $r1 [GetByteLength "UINT16"] current
      }
    }

    close $r1
    catch {unset clist1}

# finish STEP file
    if {$nface > 0} {
      if {[info exists ntriangles]} {outputMsg " Removed [expr {$ntriangles*3 - [llength $clist]}] duplicate coordinates" red}
      outputMsg " Writing [llength $clist] coordinates, $nface faces"
      set str "#$id=COORDINATES_LIST('',[llength $clist],([string range [join $clist] 0 end-1]));"
      regsub -all " " $str "" str
      puts $w2 $str

      set str "#[expr {$id+1}]=TRIANGULATED_FACE('',#$id,0,(),$,(),\n([string range [join $index] 0 end-1]));"
      regsub -all " " $str "" str
      puts $w2 $str

      incr id 2
      set str "#$id=TESSELLATED_SOLID('',("
      for {set i 101} {$i < $id} {incr i 2} {append str "#$i,"}
      set str [string range $str 0 end-1]
      append str "),$);"
      puts $w2 $str
      puts $w2 "#[expr {$id+1}]=TESSELLATED_SHAPE_REPRESENTATION('',(#$id),#16);"
      puts $w2 "ENDSEC;\nEND-ISO-10303-21;"

      foreach var {clist index} {catch {unset -- $var}}
      update idletasks
      close $w2
      outputMsg " [truncFileName [file nativename $f2]] ([fileSize $f2])"
      set localName $f2

      set cc [clock clicks -milliseconds]
      set proctime [expr {($cc - $lasttime)/1000}]
      if {$proctime <= 60} {set proctime [expr {(($cc - $lasttime)/100)/10.}]}
      outputMsg "Processing time: $proctime seconds" blue

# no index
    } else {
      close $w2
      errorMsg "Error generating STEP from STL file"
      set localName ""
      catch {file delete -force -- $f2}
    }

# errors
  } emsg]} {
    errorMsg "Error generating STEP from STL file: $emsg"
  }
}

# -------------------------------------------------------------------------------
proc GetByteLength {type} {
  switch -- $type {
    "UINT8"  {return [string length [binary format c 0]]}
    "UINT16" {return [string length [binary format s 0]]}
    "UINT32" {return [string length [binary format i 0]]}
    "REAL"   {return [string length [binary format r 0.000000E+00]]}
  }
}
