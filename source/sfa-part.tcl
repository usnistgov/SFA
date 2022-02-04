# write tessellated geometry for annotations and parts
proc x3dTessGeom {objID objEntity1 ent1} {
  global ao defaultColor draftModelCameras entCount lastID mytemp opt recPracNames savedViewFile savedViewFileName savedViewNames shapeRepName
  global shellSuppGeom spaces srNames syntaxErr tessCoord tessCoordID tessGeomTxt tessIndex tessIndexCoord tessPartFile tessPlacement
  global tessRepo tessSuppGeomFile tsName x3dColor x3dColorFile x3dColors x3dCoord x3dFile x3dIndex

  set x3dIndex $tessIndex($objID)
  set x3dCoord $tessCoord($tessIndexCoord($objID))

  if {$x3dColor == ""} {
    set x3dColor "0 0 0"
    if {[string first "annotation" [$objEntity1 Type]] != -1} {
      set msg "Syntax Error: Missing PMI Presentation color (using black).$spaces\($recPracNames(pmi242), Sec. 8.5, Fig. 84)"
      errorMsg $msg
      lappend syntaxErr([$objEntity1 Type]) [list [$objEntity1 P21ID] "color" $msg]
    }
  }
  set x3dIndexType "line"
  set solid ""
  set emit "emissiveColor='$x3dColor'"
  set spec ""
  set x3dSolid 0

# faces
  if {[string first "face" $ent1] != -1} {
    set x3dIndexType "face"
    set solid "solid='false'"

# tessellated part geometry
    if {$ao == "tessellated_solid" || $ao == "tessellated_shell"} {
      set tsID [$objEntity1 P21ID]
      set tessRepo 0
      set x3dSolid 1
      set tsName($tsID) [[[$objEntity1 Attributes] Item [expr 1]] Value]

# find name linked to product
      if {[info exists entCount(product)]} {
        if {$entCount(product) > 1} {
          if {$tsName($tsID) == ""} {
            set e0s [$objEntity1 GetUsedIn [string trim geometric_item_specific_usage] [string trim identified_item]]
            ::tcom::foreach e0 $e0s {
              for {set i 0} {$i < 5} {incr i} {set e0 [[[$e0 Attributes] Item [expr 3]] Value]}
              set tsName($tsID) [[[$e0 Attributes] Item [expr 1]] Value]
            }
          }
        }
      }

# set default color
      set x3dColor [lindex $defaultColor 0]
      tessSetColor $objEntity1 $tsID
      set spec "specularColor='[vectrim [vecmult $x3dColor 0.2]]'"
      set emit ""

# set placement for tessellated part geometry in assemblies (axis and ref_direction)
      if {[info exists entCount(item_defined_transformation)]} {tessSetPlacement $objEntity1 $tsID}
    }
  }

# write transform based on placement
  catch {unset endTransform}
  set nplace 0
  if {[info exists tessRepo]} {
    if {$tessRepo && [info exists tessPlacement(origin)]} {set nplace [llength $tessPlacement(origin)]}
  }
  if {$nplace == 0} {set nplace 1}

# file list where to write geometry
  set flist $x3dFile
  if {$ao == "tessellated_solid" || $ao == "tessellated_shell"} {
    set flist $tessPartFile
    if {$ao == "tessellated_shell" && [info exists shellSuppGeom]} {if {$shellSuppGeom} {set flist $tessSuppGeomFile}}
  }

# get saved view name
  if {[info exists draftModelCameras] && $ao == "tessellated_annotation_occurrence"} {set savedViewName [x3dGetSavedViewName $objEntity1]}

# no savedViewName, i.e., PMI not in a Saved View
  if {$ao != "tessellated_solid" && $ao != "tessellated_shell"} {
    if {![info exists savedViewName]} {set savedViewName ""}
    if {$savedViewName == ""} {
      set svn "Not in a Saved View"
      lappend savedViewName $svn
      if {[lsearch $savedViewNames $svn] == -1} {lappend savedViewNames $svn}
      set svn1 "View[lsearch $savedViewNames $svn]"
      if {![info exists savedViewFile($svn1)]} {
        catch {file delete -force -- $savedViewFileName($svn1)}
        set fn [file join $mytemp $svn1.txt]
        set savedViewFile($svn1) [open $fn w]
        set savedViewFileName($svn1) $fn
      }
    }

    if {[llength $savedViewName] > 0} {
      set numView {}
      foreach svn $savedViewName {lappend numView [lsearch $savedViewNames $svn]}
      set flist {}
      foreach num [lsort -integer $numView] {
        set svn1 "View$num"
        if {[info exists savedViewFile($svn1)]} {lappend flist $savedViewFile($svn1)}
      }
    }
  }

# -------------------------------------------------------------------------------
# loop over list of files from above
  foreach f $flist {

# annotation name
    catch {unset idshape}
    set txt [[[$objEntity1 Attributes] Item [expr 1]] Value]
    regsub -all "'" $txt "\"" idshape

# group annotations for AR workflow
    if {$opt(viewPMIAR)} {
      if {[string first "annotation" $ao] != -1} {
        set aoID [$objEntity1 P21ID]
        if {![info exists lastID($f)] || $aoID != $lastID($f)} {
          if {[info exists lastID($f)] && !$tessRepo} {puts $f "</Group>"}
          set tessGeomTxt "TAO $aoID | "
          set e0 [[[$objEntity1 Attributes] Item [expr 3]] Value]
          append tessGeomTxt "[[[$e0 Attributes] Item [expr 1]] Value] | $idshape"
          if {!$tessRepo} {puts $f "<Group id='$tessGeomTxt'>"}
        }
        set lastID($f) $aoID
      }
    }

# multiple saved view color
    if {[info exists savedViewName]} {
      if {$opt(gpmiColor) == 3 && [llength $savedViewNames] > 1} {
        if {![info exists x3dColorFile($f)]} {set x3dColorFile($f) [x3dSetPMIColor $opt(gpmiColor) 1]}
        set x3dColor $x3dColorFile($f)
        set emit "emissiveColor='$x3dColor'"
      }
    }

# -------------------------------------------------------------------------------
# loop over placements, if any
    for {set np 0} {$np < $nplace} {incr np} {
      set srName ""
      if {![info exists shapeRepName]} {
        set shapeRepName $x3dIndexType
      } elseif {$shapeRepName != "line" && $shapeRepName != "face"} {
        set srName $shapeRepName
      }

# for tessellated shell or solid name
      if {[info exists tsID]} {
        if {$tsName($tsID) != ""} {
          set srName $tsName($tsID)
        } elseif {$srName == ""} {
          set srName "[string toupper $ao] $tsID"
        }
      }

# name of shape, solid, or shell
      if {$srName != ""} {
        incr srNames($srName)
        if {$srNames($srName) == 1} {puts $f "<!-- $srName -->"}
      }

# translation and rotation (sometimes PMI and usually assemblies)
      if {$tessRepo && [info exists tessPlacement(origin)]} {
        if {![info exists tessGeomTxt]} {set tessGeomTxt ""}
        set transform [x3dTransform [lindex $tessPlacement(origin) $np] [lindex $tessPlacement(axis) $np] [lindex $tessPlacement(refdir) $np] "tessellated geometry" "" $tessGeomTxt]
        puts $f $transform
        set endTransform 1
      }

# write tessellated face or line
      if {$np == 0} {
        set defstr ""
        if {$nplace > 1} {set defstr " DEF='$shapeRepName$objID'"}

# shape
        set idstr ""
        if {[info exists idshape]} {if {$idshape != ""} {set idstr " id='$idshape'"}}
        if {$emit == ""} {
          set matID ""
          set colorID [lsearch $x3dColors $x3dColor]
          if {$colorID == -1} {
            lappend x3dColors $x3dColor
            puts $f "<Shape$idstr$defstr><Appearance DEF='appTess[llength $x3dColors]'><Material id='matTess[llength $x3dColors]' diffuseColor='$x3dColor' $spec/></Appearance>"
          } else {
            puts $f "<Shape$idstr$defstr><Appearance USE='appTess[incr colorID]'></Appearance>"
          }
        } else {
          if {$x3dIndexType == "face"} {
            puts $f "<Shape$idstr$defstr><Appearance><Material diffuseColor='$x3dColor' emissiveColor='$x3dColor' shininess='0'/></Appearance>"
          } else {
            puts $f "<Shape$idstr$defstr><Appearance><Material $emit/></Appearance>"
          }
        }

# coordinate index
        set indexedSet "<Indexed[string totitle $x3dIndexType]\Set $solid coordIndex='[string trim $x3dIndex]'>"

# coordinates
        if {[lsearch $tessCoordID $tessIndexCoord($objID)] == -1} {
          lappend tessCoordID $tessIndexCoord($objID)
          puts $f " $indexedSet\n  <Coordinate DEF='coord$tessIndexCoord($objID)' point='[string trim $x3dCoord]'/></Indexed[string totitle $x3dIndexType]\Set></Shape>"
        } else {
          puts $f " $indexedSet<Coordinate USE='coord$tessIndexCoord($objID)'/></Indexed[string totitle $x3dIndexType]\Set></Shape>"
        }

# reuse shape
      } else {
        puts $f "<Shape USE='$shapeRepName$objID'></Shape>"
      }

# -------------------------------------------------------------------------------
# for tessellated part geometry only, write mesh based on faces
      if {$opt(tessPartMesh)} {
        if {$x3dIndexType == "face" && ($ao == "tessellated_solid" || $ao == "tessellated_shell")} {
          if {$np == 0} {
            set x3dMesh ""

# write individual edges
            set edges {}
            for {set i 0} {$i < [llength $x3dIndex]} {incr i 4} {
              lappend edges [lsort "[lindex $x3dIndex $i] [lindex $x3dIndex $i+1]"]
              lappend edges [lsort "[lindex $x3dIndex $i+1] [lindex $x3dIndex $i+2]"]
              lappend edges [lsort "[lindex $x3dIndex $i] [lindex $x3dIndex $i+2]"]
            }

# try to combine some edges and write mesh
            set edges [lsort [lrmdups $edges]]
            for {set i 0} {$i < [llength $edges]} {incr i} {
              set edge [lindex $edges $i]
              set nedge [lindex $edges $i+1]
              if {[lindex $edge 1] == [lindex $nedge 0]} {
                set edge [lappend edge [lindex $nedge 1]]
                incr i
              } elseif {[lindex $edge 0] == [lindex $nedge 0]} {
                set edge [concat [lindex $nedge 1] $edge]
                incr i
              }
              append x3dMesh "$edge -1 "
            }

# write mesh
            set ecolor ""
            foreach c [split $x3dColor] {append ecolor "[expr {$c*.5}] "}
            set defstr ""
            if {$nplace > 1} {set defstr " DEF='mesh$objID'"}
            puts $f "<Shape$idstr$defstr><Appearance><Material emissiveColor='$ecolor'/></Appearance>"
            puts $f " <IndexedLineSet coordIndex='[string trim $x3dMesh]'><Coordinate USE='coord$tessIndexCoord($objID)'/></IndexedLineSet></Shape>"
          } else {
            puts $f "<Shape USE='mesh$objID'></Shape>"
          }
        }
      }

# end transform
      if {[info exists endTransform]} {puts $f "</Transform>"}
    }
  }
  set x3dCoord ""
  set x3dIndex ""
  catch {unset tessGeomTxt}
  update idletasks
}

# -------------------------------------------------------------------------------
# get saved view names
proc x3dGetSavedViewName {objEntity} {
  global draughtingModels draftModelCameraNames draftModelCameras savedsavedViewNames savedViewName

# saved view name already saved
  if {[info exists savedsavedViewNames([$objEntity P21ID])]} {return $savedsavedViewNames([$objEntity P21ID])}

  set savedViewName {}
  foreach dm $draughtingModels {
    set entDraughtingModels [$objEntity GetUsedIn [string trim $dm] [string trim items]]
    set entDraughtingCallouts [$objEntity GetUsedIn [string trim draughting_callout] [string trim contents]]
    ::tcom::foreach entDraughtingCallout $entDraughtingCallouts {
      set entDraughtingModels [$entDraughtingCallout GetUsedIn [string trim $dm] [string trim items]]
    }

    ::tcom::foreach entDraughtingModel $entDraughtingModels {
      if {[info exists draftModelCameras([$entDraughtingModel P21ID])]} {
        set dmcn $draftModelCameraNames([$entDraughtingModel P21ID])
        if {[lsearch $savedViewName $dmcn] == -1} {lappend savedViewName $dmcn}
      }
    }
  }

# save saved view name
  if {![info exists savedsavedViewNames([$objEntity P21ID])]} {set savedsavedViewNames([$objEntity P21ID]) $savedViewName}
  return $savedViewName
}

# -------------------------------------------------------------------------------
# B-rep part geometry
proc x3dBrepGeom {} {
  global brepFile brepFileName buttons cadSystem defaultColor developer grayBackground localName matTrans mytemp nistVersion nsketch opt viz
  global x3dApps x3dBbox x3dMax x3dMin x3dMsg x3dMsgColor x3dParts

  if {[catch {
    if {$opt(DEBUGX3D)} {getTiming x3dBrepGeom}

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
      set msg " Processing STEP part geometry"
      if {[info exists buttons]} {
        set fsize [file size $localName]
        if {$fsize > 50000000} {
          append msg ".  Please wait, it could take several minutes for large STEP files."
        } elseif {$fsize > 10000000} {
          append msg ", please wait."
        }
      }
      outputMsg $msg $x3dMsgColor

# run stp2x3d-part.exe
      if {$opt(DEBUGX3D)} {getTiming stp2x3d}
      catch {exec $stp2x3d --input [file nativename $localName] --quality $opt(partQuality) --edge $opt(partEdges) --sketch $opt(partSketch) --normal $opt(partNormals)} errs
      if {$opt(DEBUGX3D)} {getTiming done; outputMsg $errs}

# done processing
      if {[string first "STEP to X3D completed!" $errs] != -1} {
        if {[file exists $stpx3dFileName]} {
          if {[file size $stpx3dFileName] > 0} {
            set sketch 0
            set nind {}
            set x3dApps {}

# check for conversion units, mm > inch
            set sc [x3dBrepUnits]

# get min and max, number of materials, indents used to add Switch nodes
            set x3dBbox ""
            catch {unset indents}
            foreach line [split $errs "\n"] {
              if {[string first "No color will be supported." $line] != -1} {outputMsg "  Using [lindex $defaultColor 1] for the part color" red}

              set sline [split [string trim $line] " "]
              if {[string first "MinXYZ" $line] != -1} {
                append x3dBbox "<br>Min:"
                foreach id1 {1 2 3} id2 {x y z} {
                  set num [expr {[lindex $sline $id1]}]
                  regsub -all "," $num "." num
                  set x3dMin($id2) [expr {$sc*$num}]
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
                  set x3dMax($id2) [expr {$sc*$num}]
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
              } elseif {[string first "indent" $line] != -1} {
                set indents([lindex $sline 1]) [lindex $sline 3]
              } elseif {[string first "Sketch geometry" $line] != -1} {
                set sketch 1

# error messages from stp2x3d
              } elseif {$developer && [string first "*" $line] == 0} {
                outputMsg $line red
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
            if {$opt(DEBUGX3D)} {outputMsg "$nind\n$j $level" red}

# open temp file
            set brepFileName [file join $mytemp brep.txt]
            set brepFile [open $brepFileName w]

# integrate x3d from stp2x3d-part with existing x3dom file
            set str "\n<!-- PART GEOMETRY -->\n<Switch whichChoice='0' id='swPRT'>"
            if {$sc != 1} {
              append str "<Transform scale='$sc $sc $sc' onclick='handleGroupClick(event)'>"
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
            set close 0
            catch {unset parts}
            catch {unset matTrans}
            if {![info exists viz(EDGE)]} {set viz(EDGE) 0}

# process all lines in file
            if {$opt(DEBUGX3D)} {getTiming start}
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
                  if {!$viz(EDGE)} {if {[string first "edge" $line] != -1} {set viz(EDGE) 1}}

# add Switch for sketch geometry
                  if {$sketch} {
                    set c1 [string first "<Shape>" $line]
                    if {$c1 != -1} {
                      incr nsketch
                      set line "\n<!-- sketch geometry $nsketch -->\n[string repeat " " $c1]<Switch id='swSketch$nsketch' whichChoice='0'>[string range $line $c1 end]"
                      set oksketch 1
                    }
                  }
                }

# check for transparency
                if {[string first "<Appearance DEF" $line] != -1} {
                  set c1 [string first "transparency=" $line]
                  if {$c1 != -1} {
                    set trans [string range $line $c1+14 end]
                    set trans [string range $trans 0 [string first "'" $trans]-1]
                    set c2 [string first "'mat" $line]
                    set id [string range $line $c2+4 $c2+7]
                    set id [string range $id 0 [string first "'" $id]-1]
                    if {$trans == 1} {
                      set msg "  Some surfaces are clear and not visible"
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
                      if {$developer && $trans != 1} {errorMsg "  Some surfaces are transparent" red}
                    }
                  }
                }

# check if any sketch geometry is white
                if {$oksketch && ![info exists grayBackground]} {
                  if {[string first "<Appearance" $line] != -1} {
                    if {[string first "'1 1 1'" $line] != -1} {set grayBackground 1}
                  } elseif {[string first "<Color" $line] != -1} {
                    if {[string first "1 1 1" $line] != -1} {set grayBackground 1}
                  }
                }

# only check lines with correct number of spaces at beginning of line
                if {[string first $space $line] == 0} {

# close the Switch-Group
                  if {$close && [string first "$space/" $line] == 0} {
                    append line "\n$space\/Group></Switch>"
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
                    #if {$opt(DEBUGX3D)} {outputMsg $id blue}
                    set parts($id) $npart(PRT)
                    set line "$space\Switch id='swPart$npart(PRT)' whichChoice='0'><Group>\n$line"
                    set close 1

# get DEF name of Group and use for Switch
                  } elseif {[string first "Group" $line] != -1} {
                    set close1 0
                    if {[string first "Group" $line] != [string last "Group" $line]} {set close1 1}
                    set c1 [string first "DEF" $line]
                    if {$c1 != -1} {
                      set c1 [expr {$c1+4}]
                    } else {
                      set c1 [string first "'" $line]
                    }
                    set c2 [string last  "'" $line]
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
                      if {$opt(DEBUGX3D)} {outputMsg $id green}
                      for {set i 1} {$i < 99} {incr i} {
                        set c1 [string last "_" $id]
                        if {$c1 != -1} {
                          set nid "[string range $id 0 $c1]$i"
                        } else {
                          set nid "$id\_$i"
                        }
                        if {![info exists parts($nid)]} {
                          set id $nid
                          #if {$opt(DEBUGX3D)} {outputMsg $id red}
                          break
                        }
                      }
                    }

                    if {[string first "swSketch" $line] == -1} {
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
              }
            }

# check for duplicate part names in parts for x3dParts
            catch {unset x3dParts}
            if {[info exists parts]} {
              foreach name [lsort [array names parts]] {
                if {$opt(DEBUGX3D)} {outputMsg "$name $parts($name)"}

# S control directive
                if {[string first "\\S\\" $name] != -1} {errorMsg " The \\S\\ control directive is not supported for accented characters.  See Help > Text Strings and Numbers" red}

# check for _n at end of name
                if {[string index $name end-1] == "_" || [string index $name end-2] == "_" || [string index $name end-3] == "_"} {

# remove _n
                  set c1 [string last "_" $name]
                  set name1 [string range $name 0 $c1-1]
                  if {$opt(DEBUGX3D)} {outputMsg " $name1" red}

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
              if {$err} {errorMsg "The list of Assembly/Part names in the Viewer might have some wrong characters.  This is due to the\nencoding of the STEP file.  If possible convert the encoding of the STEP file to UTF-8 with the\nNotepad++ text editor or other software.  See Text Strings and Numbers"}
            }
            if {$opt(DEBUGX3D)} {foreach idx [array names x3dParts] {outputMsg "$idx $x3dParts($idx)" blue}}

# no shapes
            set viz(PART) 1
            if {!$shape} {
              set viz(PART) 0
              errorMsg " There is no B-rep Part Geometry in the STEP file.  There might be Tessellated Part Geometry.  Check the Viewer selections."
            }

# end the brep file
            if {$sc == 1} {
              puts $brepFile "</Group></Switch>"
            } else {
              puts $brepFile "</Transform></Switch>"
            }

            close $stpx3dFile
          }
          if {$opt(DEBUGX3D)} {getTiming done}

# no X3D output
        } else {
          errorMsg " Cannot find the part geometry (X3D file) generated by stp2x3d-part.exe"
        }
        catch {file delete -force -- $stpx3dFileName}

# errors running stp2x3d
      } elseif {[string first "Nothing to translate" $errs] != -1} {
        set msg "Part geometry cannot be processed"
        errorMsg " $msg"
        lappend x3dMsg $msg
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
        } else {
          set msg "Error processing STEP part geometry.\n Use F8 to run the Syntax Checker to check for STEP file errors.  See Help > Syntax Checker\n Try another STEP file viewer.  See Websites > STEP Software > STEP File Viewers"
        }
        errorMsg $msg
        outputMsg " "
        lappend x3dMsg "Error generating STEP part geometry"
      }

# missing stp2x3d
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

  set sc 1.
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
        set sc [trimNum [expr {1./$cf}] 5]
      }
    }
  }
  return $sc
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
          set copy 1
        }
        if {$copy} {if {$opt(DEBUGX3D)} {outputMsg "copy $stp2x3d"}; file copy -force -- $internal $stp2x3d}
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
          set copy 1
        }
        if {$copy} {if {$opt(DEBUGX3D)} {outputMsg "copy $file"}; file copy -force -- $file $fn}
      }
      if {$opt(DEBUGX3D)} {getTiming "copy and extract"}
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
  outputMsg "Generating STEP AP242 tessellated geometry from STL file" blue

  if {[catch {
    set f1 $localName
    set f2 "[file rootname $f1]-stl.stp"
    file delete -force -- $f2
    set r1 [open $f1 r]
    set w2 [open $f2 w]

# header and entities
    puts $w2 "ISO-10303-21;
HEADER;
FILE_DESCRIPTION(('[file tail $f1]'),'2;1');
FILE_NAME(' ','[clock format [clock seconds] -format "%Y-%m-%d\T%T"]',(' '),(' '),' ','NIST SFA [getVersion]',' ');
FILE_SCHEMA(('AP242_MANAGED_MODEL_BASED_3D_ENGINEERING_MIM_LF'));
ENDSEC;\n
DATA;
#1=APPLICATION_CONTEXT('managed model based 3d engineering') ;
#2=PRODUCT_CONTEXT(' ',#1,'mechanical') ;
#3=PRODUCT_DEFINITION_CONTEXT('part definition',#1,' ') ;
#4=APPLICATION_PROTOCOL_DEFINITION('international standard','ap242_managed_model_based_3d_engineering',2014,#1) ;
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
    set pnindex {}
    set nlist {}
    set cidx -1
    set midx -1
    set nface 0
    set id 100
    set binarySTL 0
    set lasttime [clock clicks -milliseconds]

# read stl file
    while {[gets $r1 line] >= 0} {

# normals
      if {[string first "normal" $line] != -1} {
        incr num
        set c1 [expr {[string first "normal" $line]+7}]
        set norm [string trim [string range $line $c1 end]]
        foreach n $norm {append n1 "[trimNum $n],"}
        lappend nlist "([string trim [string range $n1 0 end-1]]),"
        unset n1

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
          if {![info exists pnindex1($jdx)]} {
            set pnindex1($jdx) 1
            lappend pnindex $jdx
          }
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
    close $r1
    unset clist1
    unset pnindex1

# finish STEP file
    if {[llength $pnindex] > 0} {
      set str "#$id=COORDINATES_LIST('',[llength $clist],([string range [join $clist] 0 end-1]));"
      regsub -all " " $str "" str
      puts $w2 $str
      unset clist

      regsub -all " " [join [lsort -integer $pnindex]] "," pnindex
      set str "#[expr {$id+1}]=TRIANGULATED_FACE('',#$id,$nface,([string range [join $nlist] 0 end-1]),$,\n($pnindex),\n([string range [join $index] 0 end-1]));"
      regsub -all " " $str "" str
      puts $w2 $str
      unset nlist
      unset pnindex
      unset index

      incr id 2
      set str "#$id=TESSELLATED_SHELL('',("
      for {set i 101} {$i < $id} {incr i 2} {append str "#$i,"}
      set str [string range $str 0 end-1]
      append str "),$);"
      puts $w2 $str
      puts $w2 "#[expr {$id+1}]=TESSELLATED_SHAPE_REPRESENTATION('',(#$id),$);"
      puts $w2 "ENDSEC;\nEND-ISO-10303-21;"

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
      errorMsg "Binary STL files are not supported"
      set localName ""
      catch {file delete -force -- $f2}
    }

# errors
  } emsg]} {
    errorMsg "Error generating STEP from STL file: $emsg"
  }
}
