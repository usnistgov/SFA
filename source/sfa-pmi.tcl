# write tessellated geometry for PMI annotations and parts
proc x3dTessGeom {objID tessEnt faceEnt {aoname ""}} {
  global ao assemTransform defaultColor draftModelCameras entCount leaderCoords mytemp noGroupTransform opt placeCoords placeSavedView
  global recPracNames savedViewFile savedViewFileName savedViewNames shapeRepName shellSuppGeom spaces srNames syntaxErr tessCoord tessCoordID
  global tessGeomTxt tessIndex tessIndexCoord tessPartFile tessPlacement tessRepo tessSuppGeomFile tsName x3dColor x3dColorFile x3dColors
  global x3dCoord x3dFile x3dIndex

  set x3dIndex $tessIndex($objID)
  set x3dCoord $tessCoord($tessIndexCoord($objID))

  if {$x3dColor == ""} {
    set x3dColor "0 0 0"
    if {[string first "annotation" [$tessEnt Type]] != -1} {
      set msg "Syntax Error: Missing Graphic PMI color (using black).$spaces\($recPracNames(pmi242), Sec. 8.5, Fig. 85)"
      errorMsg $msg
      lappend syntaxErr([$tessEnt Type]) [list [$tessEnt P21ID] "color" $msg]
    }
  }
  set x3dIndexType "line"
  set solid ""
  set emit "emissiveColor='$x3dColor'"
  set spec ""
  set x3dSolid 0

# faces
  set ent1 $faceEnt
  if {[string first "handle" $faceEnt] != -1} {set ent1 [$faceEnt Type]}

  if {[string first "face" $ent1] != -1} {
    set x3dIndexType "face"
    set solid "solid='false'"

# tessellated part geometry
    if {$ao == "tessellated_solid" || $ao == "tessellated_shell"} {
      set tsID [$tessEnt P21ID]
      set tessRepo 0
      set x3dSolid 1
      set tsName($tsID) [[[$tessEnt Attributes] Item [expr 1]] Value]

# find name linked to product
      if {[info exists entCount(product)]} {
        if {$entCount(product) > 1} {
          if {$tsName($tsID) == ""} {
            set e0s [$tessEnt GetUsedIn [string trim geometric_item_specific_usage] [string trim identified_item]]
            ::tcom::foreach e0 $e0s {
              for {set i 0} {$i < 5} {incr i} {set e0 [[[$e0 Attributes] Item [expr 3]] Value]}
              set tsName($tsID) [[[$e0 Attributes] Item [expr 1]] Value]
            }
          }
        }
      }

# set default color
      set x3dColor $defaultColor
      tessSetColor $tessEnt $faceEnt
      set spec "specularColor='[vectrim [vecmult $x3dColor 0.2]]'"
      set emit ""

# set placement for tessellated part geometry in assemblies (axis and ref_direction)
      if {[info exists entCount(item_defined_transformation)]} {tessSetPlacement $tessEnt $tsID}
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
  if {[info exists draftModelCameras] && $ao == "tessellated_annotation_occurrence"} {set savedViewName [x3dGetSavedViewName $tessEnt]}

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

# TAO check for transform related to assembly
  set taoID [$tessEnt P21ID]
  x3dAssemblyTransform $tessEnt

# -------------------------------------------------------------------------------
# loop over list of files from above
  foreach f $flist {

# annotation name
    catch {unset idshape}
    set txt [[[$tessEnt Attributes] Item [expr 1]] Value]
    regsub -all "'" $txt "\"" idshape

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

# transform for PMI in assemblies
      if {[info exists assemTransform($taoID)]} {puts $f $assemTransform($taoID)}

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
        if {[info exists idshape]} {if {$idshape != "" && [lsearch $savedViewNames $idshape] == -1} {set idstr " id='$idshape'"}}
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
        if {![info exists tessCoordID($f)] || [lsearch $tessCoordID($f) $tessIndexCoord($objID)] == -1} {
          lappend tessCoordID($f) $tessIndexCoord($objID)
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
            foreach c [split $x3dColor] {append ecolor "[expr {$c*.25}] "}
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

# generate placeholder if in saved view
      if {[info exists placeSavedView($aoname)] && ([info exists placeCoords($aoname)] || [info exists leaderCoords($aoname)])} {
        x3dPlaceholder $aoname $f
        set noGroupTransform 1
      }

# close transform for PMI in assemblies
      if {[info exists assemTransform($taoID)]} {puts $f "</Transform>"}
    }
  }
  set x3dCoord ""
  set x3dIndex ""
  catch {unset tessGeomTxt}
  update idletasks
}

# -------------------------------------------------------------------------------
# TAO check for transform related to assembly
proc x3dAssemblyTransform {tessEnt} {
  global ao assemTransform entCount noGroupTransform opt syntaxErr taoLastID x3dMsg

  set debugTAO 0
  set taoID [$tessEnt P21ID]

  if {[catch {
    set dc "draughting_callout"
    set rrwt "representation_relationship_with_transformation_and_shape_representation_relationship"
    if {[info exists entCount($rrwt)] && [info exists entCount($dc)]} {
      if {$ao == "tessellated_annotation_occurrence" && $entCount($rrwt) > 0 && $entCount($dc) > 0} {
        if {![info exists taoLastID] || $taoID != $taoLastID} {
          if {$debugTAO} {outputMsg "\n$ao $taoID [[[$tessEnt Attributes] Item [expr 1]] Value]"}

# check for TAO in draughting callout
          set e0s [$tessEnt GetUsedIn [string trim draughting_callout] [string trim contents]]
          ::tcom::foreach e0 $e0s {
            if {$debugTAO} {outputMsg [$e0 Type][$e0 P21ID]}
            set okTransform 0

# DMIA
            set e1s [$e0 GetUsedIn [string trim draughting_model_item_association] [string trim identified_item]]
            ::tcom::foreach e1 $e1s {
              if {$okTransform == 0 && [string first "placeholder" [$e1 Type]] == -1} {
                if {$debugTAO} {outputMsg "1 [$e1 Type][$e1 P21ID]"}

# shape aspect from SAR
                set e2 [[[$e1 Attributes] Item [expr 3]] Value]
                if {$debugTAO} {outputMsg "2  [$e2 Type][$e2 P21ID]"}
                set checkCPSA 1

# composite shape aspect
                if {[$e2 Type] == "composite_group_shape_aspect" || [$e2 Type] == "composite_shape_aspect"} {
                  set e3s [$e2 GetUsedIn [string trim shape_aspect_relationship] [string trim relating_shape_aspect]]

# get shape aspect from composite
                  set n 0
                  ::tcom::foreach e3 $e3s {
                    set e2 [[[$e3 Attributes] Item [expr 4]] Value]
                    if {$debugTAO} {outputMsg "2a [$e2 Type][$e2 P21ID]"}
                    incr n
                    if {$n} {break}
                  }
                  if {[$e2 Type] == "composite_shape_aspect"} {
                    set e3s [$e2 GetUsedIn [string trim shape_aspect_relationship] [string trim relating_shape_aspect]]
                    set n 0
                    ::tcom::foreach e3 $e3s {
                      set e2 [[[$e3 Attributes] Item [expr 4]] Value]
                      if {$debugTAO} {outputMsg "2c [$e2 Type][$e2 P21ID]"}
                      incr n
                      if {$n} {break}
                    }
                  }

# dimensional location refers to two SA
                } elseif {[$e2 Type] == "dimensional_location"} {
                  foreach idx {3 4} {
                    set dlsa($idx) [[[$e2 Attributes] Item [expr $idx]] Value]

# check that it is not in CPSA
                    if {[string first "shape_aspect" [$dlsa($idx) Type]] != -1} {
                      set e3s [$dlsa($idx) GetUsedIn [string trim component_path_shape_aspect] [string trim component_shape_aspect]]
                      set ncpsa 0
                      ::tcom::foreach e3 $e3s {incr ncpsa}
                      if {$ncpsa == 0} {set e2 $dlsa($idx); break}
                    }
                  }

# dimensional size refers to one SA
                } elseif {[$e2 Type] == "dimensional_size"} {
                  set e2 [[[$e2 Attributes] Item [expr 1]] Value]
                  if {$debugTAO} {outputMsg "2a [$e2 Type][$e2 P21ID]"}
                  set checkCPSA 0
                }

# check if shape aspect is used in component_path_shape_aspect
                set oksa 1
                if {[string first "shape_aspect" [$e2 Type]] != -1 && $checkCPSA} {
                  set e3s [$e2 GetUsedIn [string trim component_path_shape_aspect] [string trim component_shape_aspect]]
                  set ncpsa 0
                  ::tcom::foreach e3 $e3s {incr ncpsa}
                  if {$ncpsa > 0} {
                    set oksa 0
                    if {$debugTAO} {outputMsg "    SA in CPSA"}
                  }
                }

                set msg "Graphic PMI on parts in an assembly might have the wrong position and orientation"
                if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}

# GISU
                if {[string first "shape_aspect" [$e2 Type]] != -1 && $oksa} {
                  set e3s [$e2 GetUsedIn [string trim geometric_item_specific_usage] [string trim definition]]
                  set ngisu 0
                  ::tcom::foreach e3 $e3s {incr ngisu}
                  if {$debugTAO && $ngisu == 0} {outputMsg "    SA not in GISU"}

                  ::tcom::foreach e3 $e3s {
                    if {$debugTAO} {outputMsg "3   [$e3 Type][$e3 P21ID]"}
                    set e4 [[[$e3 Attributes] Item [expr 4]] Value]
                    if {$e4 != ""} {
                      if {$debugTAO} {outputMsg "4    [$e4 Type][$e4 P21ID]"}

# check for ABSR in SSR rep_2 or rep_1
                      set nssr 0
                      set ssrRep 3
                      set e5s [$e4 GetUsedIn [string trim shape_representation_relationship] [string trim rep_2]]
                      ::tcom::foreach e5 $e5s {incr nssr}
                      if {$nssr == 0} {
                        set ssrRep 4
                        set e5s [$e4 GetUsedIn [string trim shape_representation_relationship] [string trim rep_1]]
                        if {$opt(debugX3D)} {errorMsg " Error getting ABSR in SSR rep_2, checking for ABSR in rep_1" red}
                      }

# shape representation
                      ::tcom::foreach e5 $e5s {
                        if {$debugTAO} {outputMsg "5     [$e5 Type][$e5 P21ID]"}
                        set e6 [[[$e5 Attributes] Item [expr $ssrRep]] Value]
                        set srID [$e6 P21ID]
                        if {$debugTAO} {outputMsg "6      [$e6 Type][$e6 P21ID]"}

# check SR in RRWT rep_1 > item defined transformation
                        set e7s [$e6 GetUsedIn [string trim $rrwt] [string trim rep_1]]
                        ::tcom::foreach e7 $e7s {
                          if {$debugTAO} {outputMsg "7       [$e7 Type][$e7 P21ID]"}
                          set e8 [[[$e7 Attributes] Item [expr 5]] Value]
                          if {$debugTAO} {outputMsg "8        [$e8 Type][$e8 P21ID]"}

# IDT transform 2 (item 4)
                          set e9 [[[$e8 Attributes] Item [expr 4]] Value]
                          if {$debugTAO} {outputMsg "9         [$e9 Type][$e9 P21ID]"}
                          set a2p3d [x3dGetA2P3D $e9]
                          if {$debugTAO} {outputMsg "t1         $a2p3d"}
                          if {$debugTAO} {outputMsg [x3dTransform [lindex $a2p3d 0] [lindex $a2p3d 1] [lindex $a2p3d 2]] red}
                          set assemTransform($taoID) [x3dTransform [lindex $a2p3d 0] [lindex $a2p3d 1] [lindex $a2p3d 2]]
                          set noGroupTransform 1

# IDT transform 1 (item 3)
                          set e10 [[[$e8 Attributes] Item [expr 3]] Value]
                          if {$debugTAO} {outputMsg "10        [$e10 Type][$e10 P21ID]"}
                          set a2p3d [x3dGetA2P3D $e10]
                          if {$debugTAO} {outputMsg "t2         $a2p3d"}
                          set okTransform 1
                        }
                      }
                    } else {
                      set msg "Syntax Error: Missing 'used_representation' attribute on geometric_item_specific_usage"
                      errorMsg $msg
                      lappend syntaxErr(geometric_item_specific_usage) [list [$e3 P21ID] "used_representation" $msg]
                    }
                  }
                } else {
                  if {$debugTAO} {outputMsg "2  [$e2 Type][$e2 P21ID]" red}
                }
              }
            }
          }
        }
        set taoLastID $taoID
      }
    }
  } emsg3]} {
    errorMsg "Error getting TAO transform: $emsg3"
  }
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
    set ndm 0
    ::tcom::foreach ent $entDraughtingModels {incr ndm}
    if {$ndm == 0} {
      set entDraughtingCallouts [$objEntity GetUsedIn [string trim draughting_callout] [string trim contents]]
      ::tcom::foreach entDraughtingCallout $entDraughtingCallouts {
        set entDraughtingModels [$entDraughtingCallout GetUsedIn [string trim $dm] [string trim items]]
      }
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
# set x3d color for PMI
proc x3dSetPMIColor {type {mode 0}} {
  global idxColor

# black
  if {$type == 1} {
    set color "0 0 0"

# random
  } elseif {$type == 2 || $type == 3} {
    incr idxColor($mode)
    switch -- $idxColor($mode) {
      1 {set color "1 0 0"}
      2 {set color "0 0 1"}
      3 {set color "0 .5 0"}
      4 {set color "1 0 1"}
      5 {set color "0 .5 .5"}
    }
    if {$idxColor($mode) == 5} {set idxColor($mode) 0}
  }
  return $color
}

# -------------------------------------------------------------------------------
# write geometry for polyline annotations
proc x3dPolylinePMI {{objEntity1 ""}} {
  global ao mytemp opt recPracNames savedViewFile savedViewFileName savedViewName savedViewNames
  global spaces x3dColor x3dColorFile x3dCoord x3dFile x3dIndex x3dIndexType x3dShape

  if {[catch {
    if {[info exists x3dCoord] || $x3dShape} {
      set flist $x3dFile

# no savedViewName, i.e., PMI not in a Saved View
      if {[llength $savedViewName] == 0} {
        set svn  "Not in a Saved View"
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

# multiple saved views, write to individual files
      if {[llength $savedViewName] > 0} {
        set flist {}
        foreach svn $savedViewName {
          set svn1 "View[lsearch $savedViewNames $svn]"
          if {[info exists savedViewFile($svn1)]} {lappend flist $savedViewFile($svn1)}
        }
      }

# loop over list of files from above
      foreach f $flist {

# multiple saved view color
        if {$opt(gpmiColor) == 3 && [llength $savedViewNames] > 1} {
          if {![info exists x3dColorFile($f)]} {set x3dColorFile($f) [x3dSetPMIColor $opt(gpmiColor) 1]}
          set x3dColor $x3dColorFile($f)
        }

# start shape
        if {[string length $x3dCoord] > 0} {
          set idstr ""
          if {[info exists idshape]} {if {$idshape != "" && [lsearch $savedViewNames $idshape] == -1} {set idstr " id='$idshape'"}}
          if {$x3dColor != ""} {
            puts $f "<Shape$idstr><Appearance><Material emissiveColor='$x3dColor'/></Appearance>"
          } else {
            puts $f "<Shape$idstr><Appearance><Material emissiveColor='0 0 0'/></Appearance>"
            errorMsg "Syntax Error: Missing Graphic PMI color for [formatComplexEnt $ao] (using black)$spaces\($recPracNames(pmi242), Sec. 8.5, Fig. 85)"
          }
          catch {unset idshape}

# index and coordinates
          puts $f " <IndexedLineSet coordIndex='[string trim $x3dIndex]'>\n  <Coordinate point='[string trim $x3dCoord]'/></IndexedLineSet></Shape>"

# end shape
        } elseif {$x3dShape && [info exists x3dIndexType]} {
          puts $f "</Indexed$x3dIndexType\Set></Shape>"
        }
      }
      set x3dCoord ""
      set x3dIndex ""
      set x3dColor ""
      set x3dShape 0
    }
  } emsg3]} {
    errorMsg "Error writing polyline annotation graphics: $emsg3"
  }
  update idletasks
}
