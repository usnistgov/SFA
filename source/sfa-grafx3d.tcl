# start x3dom file for non-FEM graphics
proc x3dFileStart {} {
  global brepEnts cadSystem entCount localName opt stepAP timeStamp viz
  global x3dColorBrep x3dColorStyle x3dFile x3dFileName x3dMax x3dMin x3dStartFile
  global objDesign

  if {$x3dStartFile == 0} {return}
  set x3dStartFile 0
  checkTempDir

  catch {file delete -force -- "[file rootname $localName]_x3dom.html"}
  catch {file delete -force -- "[file rootname $localName]-x3dom.html"}
  set x3dFileName [file rootname $localName]-sfa.html
  catch {file delete -force -- $x3dFileName}
  set x3dFile [open $x3dFileName w]

  set title [file tail $localName]
  if {$stepAP != "" && [string range $stepAP 0 1] == "AP"} {append title " | $stepAP"}
  puts $x3dFile "<!DOCTYPE html>\n<html>\n<head>\n<title>$title</title>\n<base target=\"_blank\">\n<meta http-equiv='Content-Type' content='text/html;charset=utf-8'/>"
  #puts $x3dFile "<meta http-equiv=\"Cache-Control\" content=\"no-cache, no-store, must-revalidate\" />\n<meta http-equiv=\"Pragma\" content=\"no-cache\" />\n<meta http-equiv=\"Expires\" content=\"0\" />"
  puts $x3dFile "<link rel='stylesheet' type='text/css' href='https://www.x3dom.org/x3dom/release/x3dom.css'/>\n<script type='text/javascript' src='https://www.x3dom.org/x3dom/release/x3dom.js'></script>\n</head>"

  set name [file tail $localName]
  if {$stepAP != "" && [string range $stepAP 0 1] == "AP"} {append name "&nbsp;&nbsp;&nbsp;$stepAP"}
  if {$timeStamp != ""} {
    set ts [fixTimeStamp $timeStamp]
    append name "&nbsp;&nbsp;&nbsp;$ts"
  }
  if {$cadSystem != ""} {
    regsub -all "_" $cadSystem " " cs
    append name "&nbsp;&nbsp;&nbsp;$cs"
  }
  puts $x3dFile "\n<body><font face=\"arial\">\n<h3>$name</h3>"
  puts $x3dFile "\n<table><tr><td valign='top' width='85%'>"

  if {$opt(VIZBRP)} {
    set ok 0
    foreach item $brepEnts {if {[info exists entCount($item)]} {set ok 1}}
    set x3dColorBrep [x3dBrepColor]
    if {$ok} {

# report multiple colors
      set ok1 0
      if {$x3dColorBrep == "" && [info exists entCount(styled_item)] && $x3dColorStyle} {set ok1 1}
      if {$ok1} {
        puts $x3dFile "Multiple part colors are ignored.  "

# report overriding colors
      } elseif {[info exists entCount(over_riding_styled_item)]} {
        if {$entCount(over_riding_styled_item) > 0} {
          set overriding 0
          ::tcom::foreach e0 [$objDesign FindObjects [string trim over_riding_styled_item]] {
            set item [[[[$e0 Attributes] Item [expr 3]] Value] Type]
            if {$item == "advanced_face"} {set overriding 1}
          }
          if {$overriding} {puts $x3dFile "Overriding part colors are ignored.  "}
        }
      }
    }
  }
  if {$viz(PMI)} {
    puts $x3dFile "$viz(PMIMSG)  "
  } elseif {$opt(VIZPMI)} {
    if {[string first "Some Graphical PMI" $viz(PMIMSG)] == 0} {puts $x3dFile "The STEP file contains only Semantic PMI and no Graphical PMI.  "}
  }
  if {$viz(TPG) && [info exist entCount(next_assembly_usage_occurrence)]} {puts $x3dFile "Tessellated parts in an assembly might have the wrong position and orientation or be missing."}
  puts $x3dFile "</td><td></td><tr><td valign='top' width='85%'>"

# x3d window size
  set height 900
  set width [expr {int($height*1.78)}]
  catch {
    set height [expr {int([winfo screenheight .]*0.85)}]
    set width [expr {int($height*[winfo screenwidth .]/[winfo screenheight .])}]
  }
  puts $x3dFile "\n<X3D id='x3d' showStat='false' showLog='false' x='0px' y='0px' width='$width' height='$height'>\n<Scene DEF='scene'>"

# read tessellated geometry separately because of IFCsvr limitations
  if {($viz(PMI) && [info exists entCount(tessellated_annotation_occurrence)]) || $viz(TPG)} {tessReadGeometry}
  outputMsg " Writing View to: [truncFileName [file nativename $x3dFileName]]" green

# coordinate min, max, center
  if {$x3dMax(x) != -1.e10} {
    foreach xyz {x y z} {
      set delt($xyz) [expr {$x3dMax($xyz)-$x3dMin($xyz)}]
      set xyzcen($xyz) [format "%.4f" [expr {0.5*$delt($xyz) + $x3dMin($xyz)}]]
    }
    set maxxyz $delt(x)
    if {$delt(y) > $maxxyz} {set maxxyz $delt(y)}
    if {$delt(z) > $maxxyz} {set maxxyz $delt(z)}
  }
  update idletasks
}

# -------------------------------------------------------------------------------
# write tessellated geometry for annotations and parts
proc x3dTessGeom {objID objEntity1 ent1} {
  global ao defaultColor draftModelCameras entCount nshape opt recPracNames savedViewFile savedViewNames shapeRepName shellSuppGeom srNames
  global tessCoord tessCoordID tessIndex tessIndexCoord tessPartFile tessPlacement tessRepo tessSuppGeomFile
  global x3dColor x3dColorFile x3dColors x3dColorsUsed x3dCoord x3dFile x3dIndex x3dMsg
  #outputMsg "x3dTessGeom $objID"

  set x3dIndex $tessIndex($objID)
  set x3dCoord $tessCoord($tessIndexCoord($objID))

  if {$x3dColor == ""} {
    set x3dColor "0 0 0"
    if {[string first "annotation" [$objEntity1 Type]] != -1} {
      errorMsg "Syntax Error: Missing PMI Presentation color (using black).\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 8.4, Figure 75)"
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

# multiple saved views, write PMI to individual files, collected in x3dViewpoint below
  if {[info exists draftModelCameras] && $ao == "tessellated_annotation_occurrence"} {
    set savedViewName [getSavedViewName $objEntity1]
    if {[llength $savedViewName] > 0} {
      set flist {}
      foreach svn $savedViewName {lappend flist $savedViewFile($svn)}
    }
  }

# loop over list of files from above
  foreach f $flist {

# multiple saved view color
    if {[info exists savedViewName]} {
      if {$opt(gpmiColor) == 3 && [llength $savedViewNames] > 1} {
        if {![info exists x3dColorFile($f)]} {set x3dColorFile($f) [x3dSetColor $opt(gpmiColor) 1]}
        set x3dColor $x3dColorFile($f)
        set emit "emissiveColor='$x3dColor'"
      }
    }

# loop over placements, if any
    for {set np 0} {$np < $nplace} {incr np} {
      set srName ""
      if {![info exists shapeRepName]} {
        set shapeRepName $x3dIndexType
        if {[info exists tsID]} {set srName "[string toupper $ao] $tsID"}
      } elseif {$shapeRepName != "line" && $shapeRepName != "face"} {
        set srName $shapeRepName
      }
      if {$srName != ""} {
        incr srNames($srName)
        if {$srNames($srName) == 1} {puts $f "\n<!-- $srName -->"}
      }

# translation and rotation (sometimes PMI and usually assemblies)
      if {$tessRepo && [info exists tessPlacement(origin)]} {
        set transform [x3dTransform [lindex $tessPlacement(origin) $np] [lindex $tessPlacement(axis) $np] [lindex $tessPlacement(refdir) $np] "tessellated geometry"]
        puts $f $transform
        set endTransform "</Transform>"
      }

# write tessellated face or line
      if {$np == 0} {
        set defstr ""
        if {$nplace > 1} {set defstr " DEF='$shapeRepName$objID'"}

        if {$emit == ""} {
          set matID ""
          set colorID [lsearch $x3dColors $x3dColor]
          if {$colorID == -1} {
            lappend x3dColors $x3dColor
            puts $f "<Shape$defstr><Appearance DEF='app[llength $x3dColors]'><Material id='mat[llength $x3dColors]' diffuseColor='$x3dColor' $spec></Material></Appearance>"
          } else {
            puts $f "<Shape$defstr><Appearance USE='app[incr colorID]'></Appearance>"
          }
        } else {
          puts $f "<Shape$defstr><Appearance><Material diffuseColor='$x3dColor' $emit></Material></Appearance>"
        }
        lappend x3dColorsUsed $x3dColor

        set indexedSet "<Indexed[string totitle $x3dIndexType]\Set $solid coordIndex='[string trim $x3dIndex]'>"

        if {[lsearch $tessCoordID $tessIndexCoord($objID)] == -1} {
          lappend tessCoordID $tessIndexCoord($objID)
          puts $f " $indexedSet\n  <Coordinate DEF='coord$tessIndexCoord($objID)' point='[string trim $x3dCoord]'></Coordinate></Indexed[string totitle $x3dIndexType]\Set></Shape>"
        } else {
          puts $f " $indexedSet<Coordinate USE='coord$tessIndexCoord($objID)'></Coordinate></Indexed[string totitle $x3dIndexType]\Set></Shape>"
        }

# reuse shape
      } else {
        puts $f "<Shape USE='$shapeRepName$objID'></Shape>"
      }

# for tessellated part geometry only, write mesh based on faces
      if {$opt(VIZTPGMSH) || ($ao == "tessellated_shell" && [info exists entCount(tessellated_solid)])} {
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
            if {$ao == "tessellated_shell" && [info exists entCount(tessellated_solid)]} {
              set ecolor "0 0 0"
              set msg "Triangular faces in tessellated shells are outlined in black."
              if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
            }

            set defstr ""
            if {$nplace > 1} {set defstr " DEF='mesh$objID'"}
            puts $f "<Shape$defstr><Appearance><Material emissiveColor='$ecolor'></Material></Appearance>"
            puts $f " <IndexedLineSet coordIndex='[string trim $x3dMesh]'><Coordinate USE='coord$tessIndexCoord($objID)'></Coordinate></IndexedLineSet></Shape>"
          } else {
            puts $f "<Shape USE='mesh$objID'></Shape>"
          }
        }
      }

      incr nshape
      if {[expr {$nshape%1000}] == 0} {outputMsg "  $nshape"}

# end transform
      if {[info exists endTransform]} {puts $f $endTransform}
    }
  }
  set x3dCoord ""
  set x3dIndex ""
  update idletasks
}

# -------------------------------------------------------------------------------
# finish x3d file, write tessellated edges, PMI saved view geometry, set viewpoints, add navigation and background color, and close X3DOM file
proc x3dFileEnd {} {
  global ao brepEnts brepFile brepFileName entCount nistModelURLs mytemp nistName nistVersion
  global nistNumSavedViews numTessColor opt savedViewButtons savedViewFile savedViewFileName savedViewItems savedViewNames savedViewpoint
  global sphereDef stepAP tessCoord tessEdges tessPartFile tessPartFileName viz x3dAxes x3dColorsUsed x3dFile x3dMax x3dMin x3dMsg

# PMI is already written to file
# generate b-rep part geometry based on pythonOCC and OpenCascade
  set viz(BRP) 0
  if {$opt(VIZBRP)} {
    set ok 0
    foreach item $brepEnts {if {[info exists entCount($item)]} {set ok 1}}
    if {$ok} {x3dBrepGeom}
  }

# coordinate min, max, center
  foreach idx {x y z} {
    set delt($idx) [expr {$x3dMax($idx)-$x3dMin($idx)}]
    set xyzcen($idx) [trimNum [format "%.4f" [expr {0.5*$delt($idx) + $x3dMin($idx)}]]]
  }
  set maxxyz $delt(x)
  if {$delt(z) > $maxxyz} {set maxxyz $delt(z)}
  set maxxz $maxxyz
  if {$delt(y) > $maxxyz} {set maxxyz $delt(y)}

# -------------------------------------------------------------------------------
# write tessellated edges
  set viz(EDG) 0
  if {[info exists tessEdges]} {
    puts $x3dFile "\n<!-- TESSELLATED EDGES -->\n<Switch whichChoice='0' id='swTED'><Group>"
    foreach cid [array names tessEdges] {
      puts $x3dFile "<Shape><Appearance><Material emissiveColor='0 0 0'></Material></Appearance>"
      puts $x3dFile " <IndexedLineSet coordIndex='[join $tessEdges($cid)]'>"
      puts $x3dFile "  <Coordinate DEF='coord$cid' point='$tessCoord($cid)'></Coordinate></IndexedLineSet></Shape>"
    }
    puts $x3dFile "</Group></Switch>"
    set viz(EDG) 1
    unset tessEdges
  }

# -------------------------------------------------------------------------------
# holes
  set ok 0
  set viz(HOL) 0
  set sphereDef {}
  foreach ent [list counterbore_hole_occurrence counterdrill_hole_occurrence countersink_hole_occurrence spotface_hole_occurrence] {
    if {[info exists entCount($ent)]} {set ok 1}
  }
  if {$ok} {x3dHoles $maxxyz}

# -------------------------------------------------------------------------------
# supplemental geometry
  set viz(SMG) 0
  if {[info exists entCount(constructive_geometry_representation)]} {x3dSuppGeom $maxxyz}

# -------------------------------------------------------------------------------
# write any PMI saved view geometry for multiple saved views
  set savedViewButtons {}
  if {[llength $savedViewNames] > 0} {
    for {set i 0} {$i < [llength $savedViewNames]} {incr i} {
      set svn [lindex $savedViewNames $i]
      catch {close $savedViewFile($svn)}
      if {[file size $savedViewFileName($svn)] > 0} {
        set svMap($svn) $svn
        set svWrite 1

# check if same saved view graphics already written
        if {[info exists savedViewItems($svn)]} {
          for {set j 0} {$j < $i} {incr j} {
            set svn1 [lindex $savedViewNames $j]
            if {[info exists savedViewItems($svn1)]} {
              if {$savedViewItems($svn) == $savedViewItems($svn1)} {
                set svMap($svn) $svn1
                set svWrite 0
                break
              }
            }
          }
        }

        set svn2 $svn
        if {$svn2 == ""} {
          set svn2 "Missing_name"
          set svMap($svn2) $svn2
        }
        lappend savedViewButtons $svn2

        puts $x3dFile "\n<!-- SAVED VIEW $svn2 -->"
        puts $x3dFile "<Switch whichChoice='0' id='sw$svn2'><Group>"
        if {$svWrite} {

# get saved view graphics from file
          set lastTransform ""
          set f [open $savedViewFileName($svn) r]
          while {[gets $f line] >= 0} {

# check for similar transforms
            if {[string first "<Transform" $line] == -1 && [string first "</Transform>" $line] == -1} {
              puts $x3dFile $line
            } elseif {[string first "<Transform" $line] == 0} {
              if {$line != $lastTransform} {
                if {$lastTransform != ""} {puts $x3dFile "</Transform>"}
                puts $x3dFile $line
                set lastTransform $line
              }
            }
          }
          if {$lastTransform != ""} {puts $x3dFile "</Transform>"}

          close $f
          catch {unset savedViewFile($svn)}
        } else {
          puts $x3dFile "<!-- SAME AS $svMap($svn) -->"
        }
        puts $x3dFile "</Group></Switch>"
      } else {
        catch {close $savedViewFile($svn)}
      }
      catch {file delete -force -- $savedViewFileName($svn)}
    }
  }

# -------------------------------------------------------------------------------
# coordinate axes, if not already written
  if {$x3dAxes} {
    set asize [trimNum [expr {$maxxyz*0.05}]]
    x3dCoordAxes $asize
  }

# -------------------------------------------------------------------------------
# write tessellated part
  if {[info exists tessPartFile]} {
    puts $x3dFile "\n<!-- TESSELLATED PART GEOMETRY -->\n<Switch whichChoice='0' id='swTPG'><Group>"
    catch {close $tessPartFile}

    set lastTransform ""
    set f [open $tessPartFileName r]
    #set ftmp [open [file join $mytemp tpgtmp.txt] w]

# first check for similar transforms, write to tmp file
    while {[gets $f line] >= 0} {
      if {[string first "<Transform" $line] == -1 && [string first "</Transform>" $line] == -1} {
        puts $x3dFile $line
      } elseif {[string first "<Transform" $line] == 0} {
        if {$line != $lastTransform} {
          if {$lastTransform != ""} {puts $x3dFile "</Transform>"}
          puts $x3dFile $line
          set lastTransform $line
        }
      }
    }
    if {$lastTransform != ""} {puts $x3dFile "</Transform>"}
    #close $ftmp

# then from tmp file, consolidate faces that use the same coordinates
    #catch {unset coord}
    #catch {unset index}
    #set ftmp [open [file join $mytemp tpgtmp.txt] r]
    #
    #while {[gets $ftmp line] >= 0} {
#      if {[string first "IndexedFaceSet" $line] != -1} {
#        set sline [split $line "'"]
#        #outputMsg "[llength $sline] $sline"
#
## check for IFS with new coord def
#        if {[llength $sline] == 5} {
#
## get the coordinates
#          gets $ftmp line1
#          set cidx [lindex [split $line1 "'"] 1]
#          set coord($cidx) $line1
#          #outputMsg "coordinates $cidx [string range $coord($cidx) 0 100]" green
#        }
#        lappend index($cidx) [lindex $sline 3]
#        set shape($cidx) $shp
#
## get shape
#      } elseif {[string first "<Shape>" $line] != -1 && [string first "Appearance DEF" $line] != -1} {
#        set shp $line
#        #outputMsg $shape red
#      } elseif {([string first "<Shape" $line] != -1 && [string first "emissiveColor" $line] != -1) || [string first "IndexedLineSet" $line] != -1} {
#        puts $x3dFile $line
#        outputMsg $line blue
#      } else {
#        outputMsg $line red
#      }
    #}

    close $f
    #close $ftmp
    #catch {[file delete -force -- [file join $mytemp tpgtmp.txt]]}

# write faces for each coord with index grouped together
    #foreach idx [array names coord] {
    #  puts $x3dFile $shape($idx)
    #  puts $x3dFile " <IndexedFaceSet solid='false' coordIndex='[join $index($idx)]'>"
    #  puts $x3dFile $coord($idx)
    #}

    puts $x3dFile "</Group></Switch>"
    catch {file delete -force -- $tessPartFileName}
    unset tessPartFile
    unset tessPartFileName
  }

# -------------------------------------------------------------------------------
# add b-rep part geometry from temp file
  if {$viz(BRP)} {
    if {[info exists brepFileName]} {
      if {[file exists $brepFileName]} {
        close $brepFile
        if {[file size $brepFileName] > 0} {
          set brepFile [open $brepFileName r]
          while {[gets $brepFile line] >= 0} {puts $x3dFile $line}
          close $brepFile
          catch {file delete -force -- $brepFileName}
        }
      }
    }
  }

# -------------------------------------------------------------------------------
# default and saved viewpoints
  puts $x3dFile "\n<!-- VIEWPOINTS -->"
  set cor "centerOfRotation='$xyzcen(x) $xyzcen(y) $xyzcen(z)'"
  set fov [trimNum [expr {$delt(x)*0.5 + $delt(z)*0.5}]]
  set psy [trimNum [expr {$x3dMin(y) - 1.4*$maxxz}]]

  puts $x3dFile "<Viewpoint id='Front' position='$xyzcen(x) $psy $xyzcen(z)' orientation='1 0 0 1.5708' $cor></Viewpoint>"
  puts $x3dFile "<OrthoViewpoint id='Ortho' position='$xyzcen(x) $psy $xyzcen(z)' orientation='1 0 0 1.5708' $cor fieldOfView='\[-$fov,-$fov,$fov,$fov\]'></OrthoViewpoint>"

# viewpoint orientation
  #if {[llength $savedViewNames] > 0 && $opt(VIZPMIVP)} {
  #  foreach svn $savedViewNames {
  #    if {[info exists savedViewpoint($svn)] && [lsearch $savedViewButtons $svn] != -1} {
  #      puts $x3dFile "<Transform translation='[lindex $savedViewpoint($svn) 0]'><Viewpoint id='vp$svn' position='[lindex $savedViewpoint($svn) 0]' orientation='[lindex $savedViewpoint($svn) 1]' $cor></Viewpoint></Transform>"
  #    }
  #  }
  #}

# navigation, background color
  set bgc "1 1 1"
  if {[info exists x3dColorsUsed]} {
    set x3dColorsUsed [lrmdups $x3dColorsUsed]
    foreach color {"1 1 0" "1 1 1" "1. 1. 1."} {
      if {[lsearch $x3dColorsUsed $color] != -1} {
        set bgc ".8 .8 .8"
        break
      }
    }
  }
  puts $x3dFile "\n<!-- BACKGROUND -->"
  puts $x3dFile "<Background id='BG' skyColor='$bgc'></Background>"
  puts $x3dFile "<NavigationInfo type='\"EXAMINE\" \"ANY\"'></NavigationInfo>"
  puts $x3dFile "</Scene></X3D>"

# credits
  set ver "NIST "
  set url "https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer"
  if {!$nistVersion} {
    set ver ""
    set url "https://github.com/usnistgov/SFA"
  }
  set str "\n<p>&nbsp;<br>Generated by the <a href=\"$url\">$ver\STEP File Analyzer and Viewer (v[getVersion])</a> and displayed with <a href=\"https://www.x3dom.org/\">x3dom</a>"
  if {$opt(VIZBRP)} {
    set ok 0
    foreach item $brepEnts {if {[info exists entCount($item)]} {set ok 1}}
    if {$ok} {
      append str " and <a href=\"https://github.com/tpaviot/pythonocc\">pythonOCC</a>"
      if {$viz(TPG) || $viz(PMI) || $viz(FEA) || $viz(SMG)} {append str " for part geometry"}
      append str ".&nbsp;&nbsp;&nbsp;Part geometry can also be viewed with <a href=\"https://www.cax-if.org/cax/step_viewers.php\">STEP file viewers</a>"
    }
  }
  append str ".&nbsp;&nbsp;&nbsp;<a href=\"https://www.nist.gov/disclaimer\">NIST Disclaimer</a>&nbsp;&nbsp;&nbsp;[clock format [clock seconds] -format "%d %b %G %H:%M"]"
  puts $x3dFile $str
  #puts $x3dFile "<script>if (!navigator.onLine) {document.write(\"<B><P>You must have an Internet connection to view any of the graphics.</B>\");}</script>"

# start right column
  puts $x3dFile "</td>\n\n<!-- START RIGHT COLUMN -->\n<td valign='top'>"

# -------------------------------------------------------------------------------
# for NIST model - link to drawing
  if {$nistName != ""} {
    foreach item $nistModelURLs {
      if {[string first $nistName $item] == 0} {puts $x3dFile "<a href=\"https://s3.amazonaws.com/nist-el/mfg_digitalthread/$item\">NIST Test Case Drawing</a><p>"}
    }
  }

# BRP button
  if {$viz(BRP)} {
    puts $x3dFile "\n<!-- BRP button -->\n<input type='checkbox' checked onclick='togBRP(this.value)'/>Part Geometry<p>"
  }

# TPG button
  if {$viz(TPG)} {
    puts $x3dFile "\n<!-- TPG button -->\n<input type='checkbox' checked onclick='togTPG(this.value)'/>Tessellated Part Geometry"
    if {$viz(EDG)} {puts $x3dFile "<!-- TED button -->\n<br><input type='checkbox' checked onclick='togTED(this.value)'/>Lines (Tessellated Edges)"}
    puts $x3dFile "<p>"
  }

# SMG button
  if {$viz(SMG)} {
    puts $x3dFile "\n<!-- SMG button -->\n<input type='checkbox' checked onclick='togSMG(this.value)'/>Supplemental Geometry<p>"
  }

# HOLE button
  if {$viz(HOL)} {
    puts $x3dFile "\n<!-- Hole button -->\n<input type='checkbox' checked onclick='togHole(this.value)'/>Holes<p>"
  }

# for PMI annotations - checkboxes for toggling saved view graphics
  set svmsg {}
  if {[info exists nistNumSavedViews($nistName)]} {
    if {$viz(PMI) && $nistName != "" && [llength $savedViewButtons] != $nistNumSavedViews($nistName)} {
      lappend svmsg "For the NIST test case, expecting $nistNumSavedViews($nistName) Graphical PMI Saved Views, found [llength $savedViewButtons]."
    }
  }
  if {$viz(PMI) && [llength $savedViewButtons] > 0} {
    puts $x3dFile "\n<!-- Saved View buttons -->\nSaved View PMI"
    set ok 1
    foreach svn $savedViewButtons {
      regsub -all {\-} $svn "_" svn2
      puts $x3dFile "<br><input type='checkbox' checked onclick='tog$svn2\(this.value)'/>$svn"
      if {[string first "MBD" [string toupper $svn]] == -1 && $nistName != ""} {set ok 0}
    }
    if {!$ok && [info exists nistNumSavedViews($nistName)] && [llength $savedViewButtons] <= $nistNumSavedViews($nistName)} {
      lappend svmsg "For the NIST test case, some unexpected Graphical PMI Saved View names were found."
    }
    if {$opt(VIZPMIVP)} {
      puts $x3dFile "<p>Selecting a Saved View above changes the viewpoint.  Viewpoints usually have the correct orientation but are not centered.  Use pan and zoom to center the PMI."
    }
    puts $x3dFile "<hr><p>"
  }
  if {[llength $svmsg] > 0 && $viz(PMI)} {foreach msg $svmsg {errorMsg $msg red}}

# FEM buttons
  if {$viz(FEA)} {feaButtons 1}

# extra text messages
  if {[info exists x3dMsg]} {
    if {[llength $x3dMsg] > 0} {
      puts $x3dFile "\n<!-- Messages -->"
      puts $x3dFile "<ul style=\"padding-left:20px\">"
      foreach item $x3dMsg {puts $x3dFile "<li>$item"}
      puts $x3dFile "</ul><hr><p>"
      unset x3dMsg
    }
  }

# axes button
  puts $x3dFile "\n<!-- Axes button -->\n<p><input type='checkbox' checked onclick='togAxes(this.value)'/>Origin<p>"

# background color buttons
  puts $x3dFile "\n<!-- Background buttons -->\nBackground Color<br>"
  set check1 "checked"
  set check2 ""
  if {$bgc == ".8 .8 .8"} {
    set check2 "checked"
    set check1 ""
  }
  puts $x3dFile "<input type='radio' name='bgcolor' value='1 1 1' $check1 onclick='BGcolor(this.value)'/>White<br>"
  puts $x3dFile "<input type='radio' name='bgcolor' value='.8 .8 .8' $check2 onclick='BGcolor(this.value)'/>Gray<br>"
  puts $x3dFile "<input type='radio' name='bgcolor' value='0 0 0' onclick='BGcolor(this.value)'/>Black"

# transparency slider
  set max 0
  set transFunc 0
  if {$viz(FEA) && ([info exists entCount(surface_3d_element_representation)] || [info exists entCount(volume_3d_element_representation)])} {
    set max 0.9
  } elseif {$viz(BRP)} {
    set max 0.9
  } elseif {$viz(TPG)} {
    set max 0.9
    if {$opt(VIZTPGMSH)} {set max 1}
  }
  if {$max > 0} {
    puts $x3dFile "\n<!-- Transparency slider -->\n<p>Transparency<br>(approximate)<br>"
    puts $x3dFile "<input style='width:80px' type='range' min='0' max='$max' step='0.1' value='0' onchange='matTrans(this.value)'/>"
    set transFunc 1
  }

# mouse message
  puts $x3dFile "\n<p>Key 'a' to view all, 'r' to restore, Page Up for orthographic.  <a href=\"https://www.x3dom.org/documentation/interaction/\">Use the mouse</a> in 'Examine Mode' to rotate, pan, zoom."
  puts $x3dFile "</td></tr></table>"

# -------------------------------------------------------------------------------
# function for BRP
  if {$viz(BRP)} {x3dSwitchScript BRP}

# function for TPG
  if {$viz(TPG)} {
    if {[string first "occurrence" $ao] == -1} {
      x3dSwitchScript TPG
      if {$viz(EDG)} {x3dSwitchScript TED}
    }
  }

# function for SMG
  if {$viz(SMG)} {x3dSwitchScript SMG}

# function for holes
  if {$viz(HOL)} {x3dSwitchScript Hole}

# functions for PMI
  if {$viz(PMI)} {
    if {[llength $savedViewButtons] > 0} {
      puts $x3dFile " "
      foreach svn $savedViewButtons {
        x3dSwitchScript $svMap($svn) $svn $opt(VIZPMIVP)
      }
    }
  }

# functions for FEA buttons
  if {$viz(FEA)} {feaButtons 2}

# background function
  puts $x3dFile "\n<!-- Background function -->\n<script>function BGcolor(color){document.getElementById('BG').setAttribute('skyColor', color);}</script>"

# axes function
  x3dSwitchScript Axes

# transparency function (function for fea in sfa-fea.tcl)
  set numTessColor 0
  if {$viz(TPG)} {set numTessColor [tessCountColors]}
  if {$transFunc && [string first "AP209" $stepAP] == -1} {
    puts $x3dFile "\n<!-- Transparency function -->\n<script>function matTrans(trans){"
    if {$viz(BRP)} {
      puts $x3dFile " document.getElementById('color').setAttribute('transparency', trans);"
    }
    for {set i 1} {$i <= $numTessColor} {incr i} {
      puts $x3dFile " document.getElementById('mat$i').setAttribute('transparency', trans);"
    }
    puts $x3dFile "}\n</script>"
  }

  puts $x3dFile "</font></body></html>"
  close $x3dFile
  update idletasks

  unset x3dMax
  unset x3dMin
}

# -------------------------------------------------------------------------------
# B-rep part geometry
proc x3dBrepGeom {} {
  global brepFile brepFileName buttons entCount localName mytemp viz wdir x3dColorBrep x3dMax x3dMin x3dMsg
  global objDesign

# copy stp2x3d executable to temp directory
  if {[catch {
    set stp2x3d [file join $mytemp stp2x3d.exe]
    if {[file exists [file join $wdir exe stp2x3d.exe]]} {
      set copy 0
      if {![file exists $stp2x3d]} {
        set copy 1
      } elseif {[file mtime [file join $wdir exe stp2x3d.exe]] > [file mtime $stp2x3d]} {
        set copy 1
      }
      if {$copy} {file copy -force -- [file join $wdir exe stp2x3d.exe] $mytemp}
    }

# run stp2x3d
    if {[file exists $stp2x3d]} {

# output .x3d file name, account for extra '.' characters
      set stpx3dFileName [string range $localName 0 [string first "." $localName]]
      append stpx3dFileName "x3d"
      catch {file delete -force -- $stpx3dFileName}
      set msg " Processing STEP part geometry"
      if {[info exists buttons]} {append msg ".  Wait for the popup program (stp2x3d.exe) to complete."}
      outputMsg $msg green
      catch {exec $stp2x3d [file nativename $localName]} errs

# done processing
      if {[string first "DONE!" $errs] != -1} {
        if {[file exists $stpx3dFileName]} {
          if {[file size $stpx3dFileName] > 0} {

# check for conversion units, mm > inch
            set sc 1
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

# open temp file
            set brepFileName [file join $mytemp brep.txt]
            set brepFile [open $brepFileName w]

# integrate x3d from stp2x3d with existing x3dom file
            puts $brepFile "\n<!-- B-REP PART GEOMETRY -->\n<Switch whichChoice='0' id='swBRP'><Transform scale='$sc $sc $sc'>"
            set stpx3dFile [open $stpx3dFileName r]
            update idletasks
            set write 0

# process all lines in file
            outputMsg " Processing X3D output from stp2x3d.exe" green
            update
            while {[gets $stpx3dFile line] >= 0} {
              if {[string first "<Shape" $line] != -1} {
                set line "<Shape><Appearance>"
                set write 1
              }

# loop over all coordinates
              if {[string first "<Coordinate" $line] != -1} {
                set n 0
                foreach idx {0 1 2} {set min($idx) 1.e10; set max($idx) -1.e10}
                set getMinMax 1
                while {[gets $stpx3dFile nline] >= 0} {

# done reading all coordinates, save min and max, compute ratio
                  if {[string first "/Coordinate" $nline] != -1} {
                    set getMinMax 0
                    set ratio 0
                    foreach id1 {0 1 2} id2 {x y z} {
                      if {[expr {abs($x3dMin($id2))}] > 0.1} {set ratio [expr {max($ratio, $min($id1)/$x3dMin($id2))}]}
                      if {[expr {abs($x3dMax($id2))}] > 0.1} {set ratio [expr {max($ratio, $max($id1)/$x3dMax($id2))}]}
                      set x3dMin($id2) $min($id1)
                      set x3dMax($id2) $max($id1)
                    }
                    if {$ratio > 1000.} {
                      set msg "Part geometry XYZ dimensions are much greater than the dimensions of the graphical PMI."
                      errorMsg " $msg"
                      if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
                    }
                  }

# get xyz min and max
                  if {$getMinMax} {
                    set vals [split $nline " "]
                    foreach idx {0 1 2} {
                      set val [lindex $vals $idx]
                      if {$sc != 1} {set val [expr {$val*$sc}]}
                      set min($idx) [expr {min($val, $min($idx))}]
                      set max($idx) [expr {max($val, $max($idx))}]
                    }
                  }

                  if {[string first "<Normal" $nline] != -1} {append line "\n"}
                  append line " $nline"

# write line to file every 3000 lines
                  incr n
                  if {[expr {$n%3000}] == 0} {
                    puts $brepFile $line
                    set line ""
                    set n 0
                  }
                  if {[string first "</Normal" $nline] != -1} {break}
                }

# adjust color
              } elseif {[string first "<Material" $line] != -1} {
                if {![info exists x3dColorBrep]} {set x3dColorBrep "0.65 0.65 0.7"}
                if {$x3dColorBrep == ""} {set x3dColorBrep "0.65 0.65 0.7"}
                set line "<Material id='color' diffuseColor='$x3dColorBrep' specularColor='.2 .2 .2'>"
              }

# write remaining line to file
              if {$write} {puts $brepFile $line}
              if {[string first "</Shape" $line] != -1} {set write 0}
            }
            close $stpx3dFile
            puts $brepFile "</Transform></Switch>"
            set viz(BRP) 1
          }

# no X3D output
        } else {
          errorMsg " ERROR: Cannot find the part geometry (X3D file) output from stp2x3d.exe"
        }

# delete X3D file
        catch {file delete -force -- $stpx3dFileName}

# stp2x3d did not finish
      } else {
        errorMsg " ERROR generating X3D from the part geometry.  Try another viewer, see Websites > STEP File Viewers"
        lappend x3dMsg "B-rep part geometry cannot be viewed."
      }
    } else {
      errorMsg " ERROR: The program (stp2x3d.exe) to convert b-rep part geometry to X3D was not found in $mytemp"
    }
  } emsg]} {
    errorMsg " ERROR adding Part Geometry ($emsg)"
  }
}

# -------------------------------------------------------------------------------
proc x3dBrepColor {} {
  global x3dColorBrepAdjusted x3dColorStyle x3dColorsUsed
  global objDesign

  set debug 0
  set x3dColorBrep ""
  set x3dColorStyle 0
  foreach item {manifold_solid_brep shell_based_surface_model advanced_face} {set colors($item) {}}

# get styled_item
  catch {
    ::tcom::foreach e0 [$objDesign FindObjects [string trim styled_item]] {
      if {$debug} {errorMsg "[$e0 Type] [$e0 P21ID]" green}

# styled_item.styles
      if {[$e0 Type] == "styled_item"} {
        if {[[[$e0 Attributes] Item [expr 3]] Value] != ""} {
          set item [[[[$e0 Attributes] Item [expr 3]] Value] Type]
        } else {
          errorMsg "Syntax Error: Required styled_item 'item' attribute is blank."
          set item ""
        }

        if {$item == "manifold_solid_brep" || $item == "shell_based_surface_model" || $item == "advanced_face"} {
          set x3dColorStyle 1

# presentation_style.styles
          set a1 [[$e0 Attributes] Item [expr 2]]
          ::tcom::foreach e2 [$a1 Value] {
            set a2 [[$e2 Attributes] Item [expr 1]]
            set e3 [$a2 Value]
            set a3 [[$e3 Attributes] Item [expr 2]]

# surface side style
            set e4 [$a3 Value]
            set a4s [[$e4 Attributes] Item [expr 2]]

# surface style fill area
            foreach e5 [$a4s Value] {
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
                  set x3dColorBrep ""
                  set j 0
                  ::tcom::foreach a8 [$e8 Attributes] {
                    if {$j > 0} {append x3dColorBrep "[trimNum [$a8 Value] 3] "}
                    incr j
                  }
                  set x3dColorBrep [string trim $x3dColorBrep]
                  lappend colors($item) $x3dColorBrep
                } elseif {[$e8 Type] == "draughting_pre_defined_colour"} {
                  set x3dColorBrep [x3dPreDefinedColor [[[$e8 Attributes] Item [expr 1]] Value]]
                  lappend colors($item) $x3dColorBrep
                } else {
                  errorMsg "  Unexpected color type ([$e8 Type])"
                }
                if {$debug} {errorMsg "$x3dColorBrep  $item" red}
              }
            }
          }
        }
      }
    }
  }

# ignore advanced_face if other colors exist
  if {[llength $colors(advanced_face)] > 0} {
    if {[llength $colors(manifold_solid_brep)] > 0 || [llength $colors(shell_based_surface_model)] > 0} {
      set colors(advanced_face) {}
      #errorMsg " Part colors assigned to 'advanced_face' are ignored." red
    }
  }

  set allcolors [lrmdups [concat $colors(manifold_solid_brep) $colors(shell_based_surface_model) $colors(advanced_face)]]
  if {$debug} {outputMsg $allcolors green}
  if {[llength $allcolors] > 1} {
    set x3dColorBrep ""
    errorMsg " Part colors are ignored if multiple colors are specified (using gray)."

# adjust color for comparison to PMI color
  } else {
    lappend x3dColorsUsed $x3dColorBrep
    set x3dColorBrepAdjusted ""
    foreach val $x3dColorBrep {
      set nval $val
      if {$val < 0.3} {set nval 0}
      if {$val > 0.7} {set nval 1}
      append x3dColorBrepAdjusted "$nval "
    }
    set x3dColorBrepAdjusted [string trim $x3dColorBrepAdjusted]
  }

  if {$x3dColorBrep == "0 0 0"} {errorMsg " The STEP part geometry is colored black."}
  if {$debug} {outputMsg $x3dColorBrep blue}
  return $x3dColorBrep
}

# -------------------------------------------------------------------------------
# supplemental geometry
proc x3dSuppGeom {maxxyz} {
  global cgrObjects planeDef recPracNames syntaxErr tessSuppGeomFile tessSuppGeomFileName trimVal viz x3dColorsUsed x3dFile x3dMsg
  global objDesign

  set size [trimNum [expr {$maxxyz*0.025}]]
  set tsize [trimNum [expr {$size*0.33}]]
  set axesDef {}
  set planeDef {}

  outputMsg " Processing supplemental geometry" green
  puts $x3dFile "\n<!-- SUPPLEMENTAL GEOMETRY -->\n<Switch whichChoice='0' id='swSMG'><Group>"
  if {![info exists cgrObjects]} {set cgrObjects [$objDesign FindObjects [string trim constructive_geometry_representation]]}
  ::tcom::foreach e0 $cgrObjects {
    set a1 [[$e0 Attributes] Item [expr 2]]

# process all items
    ::tcom::foreach e2 [$a1 Value] {
      if {[catch {
        set ename [$e2 Type]
        switch $ename {
          plane -
          axis2_placement_3d {
            set name [[[$e2 Attributes] Item [expr 1]] Value]
            set e3 $e2
            if {$ename == "plane"} {set e3 [[[$e2 Attributes] Item [expr 2]] Value]}

# axis
            if {$ename == "axis2_placement_3d"} {
              set a2p3d [x3dGetA2P3D $e3]
              set origin [lindex $a2p3d 0]
              set axis   [lindex $a2p3d 1]
              set refdir [lindex $a2p3d 2]
              set transform [x3dTransform $origin $axis $refdir "supplemental geometry axes"]

# check for axis color
              set axisColor ""
              if {[catch {
                set e4s [$e3 GetUsedIn [string trim styled_item] [string trim item]]
                ::tcom::foreach e4 $e4s {
                  set e5s [[[$e4 Attributes] Item [expr 2]] Value]
                  ::tcom::foreach e5 $e5s {
                    set e6 [[[$e5 Attributes] Item [expr 1]] Value]
                    set e7 [[[$e6 Attributes] Item [expr 4]] Value]
                    if {$e7 != ""} {
                      if {[$e7 Type] == "colour_rgb"} {
                        set j 0
                        ::tcom::foreach a7 [$e7 Attributes] {
                          if {$j > 0} {append axisColor "[trimNum [$a7 Value] 3] "}
                          incr j
                        }
                        set axisColor [string trim $axisColor]
                      } elseif {[$e7 Type] == "draughting_pre_defined_colour"} {
                        set axisColor [x3dPreDefinedColor [[[$e7 Attributes] Item [expr 1]] Value]]
                      } else {
                        errorMsg " Unexpected color '[$e7 Type]' for '$ename' supplemental geometry."
                      }
                      #outputMsg color$axisColor
                    }
                  }
                }
              } emsg]} {
                errorMsg " ERROR getting color for '$ename' supplemental geometry."
              }

              set closeTransform 1
              if {$axisColor == ""} {
                set id [lsearch $axesDef $size]
                if {$id != -1} {
                  puts $x3dFile "$transform<Group USE='axes$id'></Group></Transform>"
                  set closeTransform 0
                } else {
                  lappend axesDef $size
                  puts $x3dFile $transform
                  puts $x3dFile " <Group DEF='axes[expr {[llength $axesDef]-1}]'><Shape><Appearance><Material emissiveColor='1 0 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. $size 0. 0.'></Coordinate></IndexedLineSet></Shape>"
                  puts $x3dFile " <Shape><Appearance><Material emissiveColor='0 .5 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. $size 0.'></Coordinate></IndexedLineSet></Shape>"
                  puts $x3dFile " <Shape><Appearance><Material emissiveColor='0 0 1'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. 0. $size'></Coordinate></IndexedLineSet></Shape></Group>"
                }
              } else {
                set id [lsearch $axesDef "$size $axisColor"]
                if {$id != -1} {
                  puts $x3dFile "$transform<Group USE='axes$id'></Group></Transform>"
                  set closeTransform 0
                } else {
                  lappend axesDef "$size $axisColor"
                  set sz [trimNum [expr {$size*1.5}]]
                  set tsize [trimNum [expr {$sz*0.33}]]
                  puts $x3dFile $transform
                  puts $x3dFile " <Group DEF='axes[expr {[llength $axesDef]-1}]'><Shape><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. $sz 0. 0.'></Coordinate></IndexedLineSet><Appearance><Material emissiveColor='$axisColor'></Material></Appearance></Shape>"
                  puts $x3dFile " <Shape><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. $sz 0.'></Coordinate></IndexedLineSet><Appearance><Material emissiveColor='$axisColor'></Material></Appearance></Shape>"
                  puts $x3dFile " <Shape><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. 0. $sz'></Coordinate></IndexedLineSet><Appearance><Material emissiveColor='$axisColor'></Material></Appearance></Shape>"
                  puts $x3dFile " <Transform translation='$sz 0 0' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='\"X\"'><FontStyle family='\"SANS\"'></FontStyle></Text><Appearance><Material diffuseColor='$axisColor'></Material></Appearance></Shape></Billboard></Transform>"
                  puts $x3dFile " <Transform translation='0 $sz 0' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='\"Y\"'><FontStyle family='\"SANS\"'></FontStyle></Text><Appearance><Material diffuseColor='$axisColor'></Material></Appearance></Shape></Billboard></Transform>"
                  puts $x3dFile " <Transform translation='0 0 $sz' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='\"Z\"'><FontStyle family='\"SANS\"'></FontStyle></Text><Appearance><Material diffuseColor='$axisColor'></Material></Appearance></Shape></Billboard></Transform></Group>"
                  lappend x3dColorsUsed $axisColor
                }
              }

              set nsize [trimNum [expr {$tsize*1.5}]]
              set tcolor "1 0 0"
              if {$axisColor != ""} {set tcolor $axisColor}
              if {$name != ""} {puts $x3dFile " <Transform scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='\"$name\"'><FontStyle family='\"SANS\"' justify='\"BEGIN\"'></FontStyle></Text><Appearance><Material diffuseColor='$tcolor'></Material></Appearance></Shape></Billboard></Transform>"}
              if {$closeTransform} {puts $x3dFile "</Transform>"}
              set viz(SMG) 1

# plane
            } elseif {$ename == "plane"} {
              x3dSuppGeomPlane $e2 $size $name
            }
          }

          trimmed_curve -
          composite_curve -
          geometric_curve_set {

# get all trimmed curves
            #outputMsg "$ename [$e2 P21ID]" red
            catch {unset trimVal}
            set trimmedCurves {}
            if {$ename == "trimmed_curve"} {
              lappend trimmedCurves $e2

# composite_curve -> composite_curve_segment -> trimmed_curve
            } elseif {$ename == "composite_curve"} {
              ::tcom::foreach ccs [[[$e2 Attributes] Item [expr 2]] Value] {
                lappend trimmedCurves [[[$ccs Attributes] Item [expr 3]] Value]
              }

# geometric_curve_set -> list of trimmed_ or composite_curve -> trimmed_curve or cartesian_point
            } elseif {$ename == "geometric_curve_set"} {
              set e3s [[[$e2 Attributes] Item [expr 2]] Value]
              foreach e3 $e3s {
                set ename1 [$e3 Type]
                switch $ename1 {
                  trimmed_curve   {lappend trimmedCurves $e3}
                  composite_curve {::tcom::foreach ccs [[[$e3 Attributes] Item [expr 2]] Value] {lappend trimmedCurves [[[$ccs Attributes] Item [expr 3]] Value]}}
                  line            {x3dSuppGeomLine   $e3 $tsize $name}
                  circle          {x3dSuppGeomCircle $e3 $tsize $name}
                  cartesian_point {x3dSuppGeomPoint  $e3 $tsize}
                  default {
                    set msg "Supplemental geometry for '[$e3 Type]' in 'geometric_curve_set' is not shown."
                    errorMsg " $msg"
                    if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
                  }
                }
              }
            }

# loop on trimmed curves, process lines and circles
            foreach tc $trimmedCurves {

# trimming with values OK, cartesian_points for lines, but circles NG, do not delete the meaningless 'catch'
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

              set name [[[$tc Attributes] Item [expr 1]] Value]
              set e3 [[[$tc Attributes] Item [expr 2]] Value]
              if {$name == "BRepFtrWithContract"} {set name ""}

# line, circle
              if {[$e3 Type] == "line"} {
                x3dSuppGeomLine $e3 $tsize $name
              } elseif {[$e3 Type] == "circle"} {
                x3dSuppGeomCircle $e3 $tsize $name
              } else {
                set msg "Supplemental geometry for '[$e3 Type]' in 'trimmed_curve' is not shown."
                errorMsg " $msg"
                if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
              }
            }
          }

          polyline {
            set e2s [[[$e2 Attributes] Item [expr 2]] Value]
            set polyline ""
            set ncoord 0
            ::tcom::foreach e3 $e2s {
              append polyline "[vectrim [[[$e3 Attributes] Item [expr 2]] Value]] "
              incr ncoord
              if {$ncoord == 1} {set origin $polyline}
            }
            set index ""
            for {set i 0} {$i < $ncoord} {incr i} {append index "$i "}
            append index "-1"
            puts $x3dFile " <Shape><IndexedLineSet coordIndex='$index'><Coordinate point='$polyline'></Coordinate></IndexedLineSet><Appearance><Material emissiveColor='1 0 1'></Material></Appearance></Shape>"
            if {$name != ""} {
              set nsize [trimNum [expr {$tsize*0.5}]]
              puts $x3dFile " <Transform scale='$nsize $nsize $nsize' translation='$origin'><Billboard axisOfRotation='0 0 0'><Shape><Text string='\"$name\"'><FontStyle family='\"SANS\"' justify='\"BEGIN\"'></FontStyle></Text><Appearance><Material diffuseColor='1 0 1'></Material></Appearance></Shape></Billboard></Transform>"
            }
            set viz(SMG) 1
          }

          cartesian_point {
            x3dSuppGeomPoint $e2 $tsize
          }

          cylindrical_surface {
            x3dSuppGeomCylinder $e2 $tsize
          }

          shell_based_surface_model {
            set cylIDs {}
            set e3 [lindex [[[$e2 Attributes] Item [expr 2]] Value] 0]
            set e4s [[[$e3 Attributes] Item [expr 2]] Value]
            ::tcom::foreach e4 $e4s {
              set e5 [lindex [[[$e4 Attributes] Item [expr 3]] Value] 0]
              if {[$e5 Type] == "cylindrical_surface"} {
                if {[lsearch $cylIDs [$e5 P21ID]] == -1} {
                  lappend cylIDs [$e5 P21ID]
                  x3dSuppGeomCylinder $e5 $tsize
                }
              } elseif {[$e5 Type] == "plane"} {
                x3dSuppGeomPlane $e5 $size
              } else {
                set msg "Supplemental geometry for '[$e5 Type]' in 'shell_based_surface_model' is not shown."
                errorMsg " $msg"
                if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
              }
            }
          }

          default {
            if {$ename != "tessellated_shell" && $ename != "tessellated_wire"} {
              if {$ename != "direction"} {
                set msg "Supplemental geometry for '$ename' is not shown."
                errorMsg " $msg"
                if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
              } else {
                set msg "Supplemental geometry for '$ename' is not valid."
                if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
                set msg "Syntax Error: Supplemental geometry for '$ename' is not valid.\n[string repeat " " 14]\($recPracNames(suppgeom), Sec. 4.2)"
                errorMsg $msg
                lappend syntaxErr(constructive_geometry_representation) [list [$e0 P21ID] "items" $msg]
              }
            }
          }
        }
      } emsg]} {
        errorMsg " ERROR adding '$ename' Supplemental Geometry: $emsg"
      }
    }
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
  puts $x3dFile "</Group></Switch>\n"
}

# -------------------------------------------------------------------------------
# supplemental geometry for point and the origin of a hole
proc x3dSuppGeomPoint {e2 tsize {thruHole ""} {holeName ""}} {
  global sphereDef viz x3dFile

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

# point is a black non-reflecting sphere
    set id [lsearch $sphereDef $tsize]
    if {$id != -1} {
      puts $x3dFile "<Transform translation='[vectrim $coord1]'><Shape USE='point$id'></Shape></Transform>"
    } else {
      lappend sphereDef $tsize
      puts $x3dFile "<Transform translation='[vectrim $coord1]'><Shape DEF='point[expr {[llength $sphereDef]-1}]'><Sphere radius='[trimNum [expr {$tsize*0.05}]]'></Sphere><Appearance><Material diffuseColor='0 0 0' emissiveColor='0 0 0'></Material></Appearance></Shape></Transform>"
    }

# point name
    if {$name != ""} {
      set nsize [trimNum [expr {$tsize*0.25}]]
      puts $x3dFile " <Transform translation='[vectrim $coord1]' scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='\"$name\"'><FontStyle family='\"SANS\"' justify='\"BEGIN\"'></FontStyle></Text><Appearance><Material diffuseColor='0 0 0'></Material></Appearance></Shape></Billboard></Transform>"
    }
    set viz(SMG) 1
  } emsg]} {
    errorMsg "ERROR adding 'point' supplemental geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry for line
proc x3dSuppGeomLine {e3 tsize {name ""}} {
  global trimVal viz x3dFile

  if {[catch {
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

# line geometry
    set coord2 [vectrim [vecadd $coord1 $coord2]]
    puts $x3dFile "<Shape><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='$coord1 $coord2'></Coordinate></IndexedLineSet><Appearance><Material emissiveColor='1 0 1'></Material></Appearance></Shape>"

# line name at beginning
    if {$name != ""} {
      set nsize [trimNum [expr {$tsize*0.5}]]
      puts $x3dFile " <Transform translation='$coord1' scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='\"$name\"'><FontStyle family='\"SANS\"' justify='\"BEGIN\"'></FontStyle></Text><Appearance><Material diffuseColor='1 0 1'></Material></Appearance></Shape></Billboard></Transform>"
    }
    set viz(SMG) 1

  } emsg]} {
    errorMsg "ERROR adding 'line' supplemental geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry for circle
proc x3dSuppGeomCircle {e3 tsize {name ""}} {
  global DTR trimVal viz x3dFile x3dMsg

  if {[catch {
    set e4 [[[$e3 Attributes] Item [expr 2]] Value]
    set rad [[[$e3 Attributes] Item [expr 3]] Value]

# circle position and orientation
    set a2p3d [x3dGetA2P3D $e4]
    set origin [lindex $a2p3d 0]
    set axis   [lindex $a2p3d 1]
    set refdir [lindex $a2p3d 2]
    set transform [x3dTransform $origin $axis $refdir "supplemental geometry circle"]
    puts $x3dFile $transform

# generate circle, account for trimming
# lim is the limit on an angle before deciding it is in degrees to convert to radians
    set ns 24
    set angle 0.
    set dlt [expr {6.28319/$ns}]
    set trimmed 0
    set lim 6.28319

# trim with angles
    if {[info exists trimVal(1)]} {
      if {[string first "handle" $trimVal(1)] == -1} {
        set angle $trimVal(1)
        set conv 1.
        if {$trimVal(1) > $lim && $trimVal(2) > $lim} {
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
        set msg "Supplemental geometry 'circle' trimmed by 'cartesian_point' are not trimmed."
        errorMsg " $msg"
        if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
      }
    }
    set index ""
    for {set i 0} {$i < $ns} {incr i} {append index "$i "}
    if {!$trimmed} {append index "0 "}
    append index "-1"

    set coord ""
    for {set i 0} {$i < $ns} {incr i} {
      append coord "[trimNum [expr {$rad*cos($angle)}]] "
      append coord "[trimNum [expr {-1.*$rad*sin($angle)}]] "
      append coord "0 "
      set angle [expr {$angle+$dlt}]
      if {$i == 0} {set coord1 $coord}
    }

# circle geometry and possible name
    puts $x3dFile " <Shape><IndexedLineSet coordIndex='$index'><Coordinate point='$coord'></Coordinate></IndexedLineSet><Appearance><Material emissiveColor='1 0 1'></Material></Appearance></Shape>"
    if {$name != ""} {puts $x3dFile " <Transform translation='$coord1' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='\"$name\"'><FontStyle family='\"SANS\"' justify='\"BEGIN\"'></FontStyle></Text><Appearance><Material diffuseColor='1 0 1'></Material></Appearance></Shape></Billboard></Transform>"}
    puts $x3dFile "</Transform>"
    set viz(SMG) 1

  } emsg]} {
    errorMsg "ERROR adding 'circle' supplemental geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry for plane
proc x3dSuppGeomPlane {e2 size {name ""}} {
  global planeDef viz x3dFile x3dMsg

  if {[catch {
    set e3 [[[$e2 Attributes] Item [expr 2]] Value]

# plane position and orientation
    set a2p3d [x3dGetA2P3D $e3]
    set origin [lindex $a2p3d 0]
    set axis   [lindex $a2p3d 1]
    set refdir [lindex $a2p3d 2]
    set transform [x3dTransform $origin $axis $refdir "supplemental geometry plane"]

# plane geometry
    set nsize [trimNum [expr {$size*2.}]]
    set id [lsearch $planeDef $nsize]
    if {$id != -1} {
      puts $x3dFile "$transform<Group USE='plane$id'></Group>"
    } else {
      lappend planeDef $nsize
      puts $x3dFile $transform
      puts $x3dFile " <Group DEF='plane[expr {[llength $planeDef]-1}]'><Shape><IndexedLineSet coordIndex='0 1 2 3 0 -1'><Coordinate point='-$nsize -$nsize 0. $nsize -$nsize 0. $nsize $nsize 0. -$nsize $nsize 0.'></Coordinate></IndexedLineSet><Appearance><Material emissiveColor='0 0 1'></Material></Appearance></Shape>"
      puts $x3dFile " <Shape><IndexedFaceSet solid='false' coordIndex='0 1 2 3 -1'><Coordinate point='-$nsize -$nsize 0. $nsize -$nsize 0. $nsize $nsize 0. -$nsize $nsize 0.'></Coordinate></IndexedFaceSet><Appearance><Material diffuseColor='0 0 1' transparency='0.8'></Material></Appearance></Shape></Group>"
    }

# plane name at one corner
    if {$name != ""} {
      set tsize [trimNum [expr {$size*0.33}]]
      puts $x3dFile " <Transform translation='-$nsize -$nsize 0.' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='\"$name\"'><FontStyle family='\"SANS\"' justify='\"BEGIN\"'></FontStyle></Text><Appearance><Material diffuseColor='0 0 1'></Material></Appearance></Shape></Billboard></Transform>"
    }
    puts $x3dFile "</Transform>"
    set viz(SMG) 1

# check if the plane is bounded
    set bnds [$e2 GetUsedIn [string trim advanced_face] [string trim faced_geometry]]
    set bound 0
    ::tcom::foreach bnd $bnds {set bound 1}
    if {$bound} {
      set msg "Bounding edges for supplemental geometry 'plane' are ignored."
      errorMsg " $msg"
      if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
    }

  } emsg]} {
    errorMsg "ERROR adding 'plane' supplemental geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry for cylinder
proc x3dSuppGeomCylinder {e2 tsize {name ""}} {
  global viz x3dFile x3dMsg

  if {[catch {
    set e3 [[[$e2 Attributes] Item [expr 2]] Value]
    set rad [[[$e2 Attributes] Item [expr 3]] Value]

# cylinder position and orientation
    set a2p3d [x3dGetA2P3D $e3]
    set origin [lindex $a2p3d 0]
    set axis   [lindex $a2p3d 1]
    set refdir [lindex $a2p3d 2]
    set transform [x3dTransform $origin $axis $refdir "supplemental geometry cylinder"]
    puts $x3dFile "$transform<Transform rotation='1 0 0 1.5708'>"

# cylinder geometry
    puts $x3dFile "  <Shape><Cylinder radius='$rad' height='[trimNum [expr {$tsize*10.}]]' top='false' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 0 1' transparency='0.8'></Material></Appearance></Shape>"
    puts $x3dFile "</Transform></Transform>"
    set viz(SMG) 1

# check if the cylinder is bounded
    set bnds [$e2 GetUsedIn [string trim advanced_face] [string trim face_geometry]]
    set bound 0
    ::tcom::foreach e0 $bnds {set bound 1}
    if {$bound} {
      set msg "Bounding edges for supplemental geometry 'cylindrical_surface' are ignored."
      errorMsg " $msg"
      if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
    }

  } emsg]} {
    errorMsg "ERROR adding 'cylinder' supplemental geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# holes counter and spotface
proc x3dHoles {maxxyz} {
  global dim entCount holeDefinitions x3dFile viz syntaxErr DTR
  global objDesign

  set drillPoint [trimNum [expr {$maxxyz*0.02}]]
  set head 1
  set holeDEF {}

  set scale 1.
  if {$dim(unit) == "INCH"} {set scale 25.4}

  ::tcom::foreach e0 [$objDesign FindObjects [string trim item_identified_representation_usage]] {
    if {[catch {
      set e1 [[[$e0 Attributes] Item [expr 3]] Value]
      set e2 [[[$e0 Attributes] Item [expr 5]] Value]
      if {[string first "occurrence" [$e1 Type]] != -1 && [$e2 Type] == "mapped_item"} {
        set defID   [[[[$e1 Attributes] Item [expr 5]] Value] P21ID]
        set defType [[[[$e1 Attributes] Item [expr 5]] Value] Type]

        set holeName [split $defType "_"]
        foreach idx {0 1} {
          if {[string first "counter" [lindex $holeName $idx]] != -1 || [string first "spotface" [lindex $holeName $idx]] != -1} {set holeName [lindex $holeName $idx]}
        }

# check if there is an a2p3d associated with a hole occurrence
        set e3 [[[$e2 Attributes] Item [expr 3]] Value]
        if {[$e3 Type] == "axis2_placement_3d"} {
          if {$head} {
            outputMsg " Processing hole geometry" green
            puts $x3dFile "\n<!-- HOLES -->\n<Switch whichChoice='0' id='swHole'><Group>"
            set head 0
            set viz(HOL) 1
          }
          if {[lsearch $holeDEF $defID] == -1} {puts $x3dFile "<!-- $defType $defID -->"}

# hole geometry
          if {[info exists holeDefinitions($defID)]} {
            #outputMsg "$defID $holeDefinitions($defID)" red

# hole origin and axis transform
            set a2p3d [x3dGetA2P3D $e3]
            set origin [lindex $a2p3d 0]
            set axis   [lindex $a2p3d 1]
            set refdir [lindex $a2p3d 2]
            set transform [x3dTransform $origin $axis $refdir $holeName]

# drilled hole dimensions
            set drill [lindex $holeDefinitions($defID) 0]
            set drillRad [trimNum [expr {[lindex $drill 1]*0.5*$scale}] 5]
            set drillPoint $drillRad
            catch {unset drillDep}
            if {[llength $drill] > 2} {set drillDep [expr {[lindex $drill 2]*$scale}]}

# through hole
            set holeTop "true"
            set thruHole [lindex $holeDefinitions($defID) end-1]
            if {$thruHole == 1} {set holeTop "false"}

# hole name
            set holeName [lindex $holeDefinitions($defID) end]

            catch {unset sink}
            catch {unset bore}

            if {[llength $holeDefinitions($defID)] > 1} {
              set holeType [lindex [lindex $holeDefinitions($defID) 1] 0]

# countersink hole (cylinder, cone)
              if {$holeType == "countersink"} {
                set sink [lindex $holeDefinitions($defID) 1]

# compute length of countersink from angle and radius
                set sinkRad [trimNum [expr {[lindex $sink 1]*0.5*$scale}] 5]
                set sinkAng [expr {[lindex $sink 2]*0.5}]
                set sinkDep [expr {($sinkRad-$drillRad)/tan($sinkAng*$DTR)}]

# check for bad radius and depth
                if {$sinkRad <= $drillRad} {
                  set msg "Syntax Error: $holeType diameter <= drill diameter"
                  errorMsg $msg
                  foreach ent [list $holeType\_hole_definition simplified_$holeType\_hole_definition] {
                    if {[info exists entCount($ent)]} {
                      lappend syntaxErr($ent) [list $defID "countersink_diameter" $msg]
                      lappend syntaxErr($ent) [list $defID "drilled_hole_diameter" $msg]
                    }
                  }
                }
                if {[info exist drillDep]} {
                  if {$sinkDep >= $drillDep} {
                    set msg "Syntax Error: $holeType computed 'depth' >= drill depth"
                    errorMsg $msg
                    foreach ent [list $holeType\_hole_definition simplified_$holeType\_hole_definition] {
                      if {[info exists entCount($ent)]} {lappend syntaxErr($ent) [list $defID "drilled_hole_depth" $msg]}
                    }
                  }
                }

                if {[lsearch $holeDEF $defID] == -1} {
                  puts $x3dFile "$transform<Group DEF='$holeName$defID'>"
                  if {[info exists drillDep]} {
                    puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {($drillDep+$sinkDep)*0.5}] 5]'>"
                    puts $x3dFile "  <Shape><Cylinder radius='$drillRad' height='[trimNum [expr {$drillDep-$sinkDep}] 5]' top='$holeTop' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 1'></Material></Appearance></Shape></Transform>"
                  }
                  puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {$sinkDep*0.5}] 5]'>"
                  puts $x3dFile "  <Shape><Cone bottomRadius='$sinkRad' topRadius='$drillRad' height='[trimNum $sinkDep 5]' top='false' bottom='false' solid='false'></Cone><Appearance><Material diffuseColor='0 1 1'></Material></Appearance></Shape></Transform>"
                  puts $x3dFile "</Group></Transform>"
                  lappend holeDEF $defID
                } else {
                  puts $x3dFile "$transform<Group USE='$holeName$defID'></Group></Transform>"
                }

# counterbore or spotface hole (2 cylinders, flat cone)
              } elseif {$holeType == "counterbore" || $holeType == "spotface"} {
                set bore [lindex $holeDefinitions($defID) 1]
                set boreRad [expr {[lindex $bore 1]*0.5*$scale}]
                set boreDep [expr {[lindex $bore 2]*$scale}]

# check for bad radius and depth
                if {$boreRad <= $drillRad} {
                  set msg "Syntax Error: $holeType diameter <= drill diameter"
                  errorMsg $msg
                  foreach ent [list $holeType\_hole_definition simplified_$holeType\_hole_definition] {
                    if {[info exists entCount($ent)]} {
                      lappend syntaxErr($ent) [list $defID "counterbore" $msg]
                      lappend syntaxErr($ent) [list $defID "drilled_hole_diameter" $msg]
                    }
                  }
                }
                if {[info exist drillDep]} {
                  if {$boreDep >= $drillDep} {
                    set msg "Syntax Error: $holeType depth >= drill depth"
                    errorMsg $msg
                    foreach ent [list $holeType\_hole_definition simplified_$holeType\_hole_definition] {
                      if {[info exists entCount($ent)]} {
                        lappend syntaxErr($ent) [list $defID "counterbore" $msg]
                        lappend syntaxErr($ent) [list $defID "drilled_hole_depth" $msg]
                      }
                    }
                  }
                }

                if {[lsearch $holeDEF $defID] == -1} {
                  puts $x3dFile "$transform<Group DEF='$holeName$defID'>"
                  if {[info exists drillDep]} {
                    puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {($drillDep+$boreDep)*0.5}] 5]'>"
                    puts $x3dFile "  <Shape><Cylinder radius='$drillRad' height='[trimNum [expr {$drillDep-$boreDep}] 5]' top='$holeTop' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 0'></Material></Appearance></Shape></Transform>"
                  }
                  puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum $boreDep 5]'>"
                  puts $x3dFile "  <Shape><Cone bottomRadius='$boreRad' topRadius='$drillRad' height='0.001' top='false' bottom='false' solid='false'></Cone><Appearance><Material diffuseColor='0 1 0'></Material></Appearance></Shape></Transform>"
                  puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {$boreDep*0.5}] 5]'>"
                  puts $x3dFile "  <Shape><Cylinder radius='$boreRad' height='[trimNum $boreDep 5]' top='false' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 0'></Material></Appearance></Shape></Transform>"
                  puts $x3dFile "</Group></Transform>"
                  lappend holeDEF $defID
                } else {
                  puts $x3dFile "$transform<Group USE='$holeName$defID'></Group></Transform>"
                }
              }
            }
          } else {
            errorMsg "Only drill entry points for holes are shown when no spreadsheet\n is generated with the report for Semantic PMI (See Options tab)."
            if {[lsearch $holeDEF $defID] == -1} {lappend holeDEF $defID}
          }

# point at origin of hole
          set e4 [[[$e3 Attributes] Item [expr 2]] Value]
          if {![info exists thruHole]} {set thruHole 0}
          x3dSuppGeomPoint $e4 $drillPoint $thruHole $holeName
        }
      }
    } emsg]} {
      errorMsg "ERROR adding 'hole' geometry: $emsg"
    }
  }
  if {$viz(HOL)} {puts $x3dFile "</Group></Switch>\n"}
  catch {unset holeDefinitions}
}

# -------------------------------------------------------------------------------
# set predefined color
proc x3dPreDefinedColor {name} {
  global defaultColor recPracNames

  switch -- $name {
    black   {set color "0 0 0"}
    white   {set color "1 1 1"}
    red     {set color "1 0 0"}
    yellow  {set color "1 1 0"}
    green   {set color "0 1 0"}
    cyan    {set color "0 1 1"}
    blue    {set color "0 0 1"}
    magenta {set color "1 0 1"}
    default {
      set color [lindex $defaultColor 0]
      errorMsg "Syntax Error: Unexpected draughting_pre_defined_colour name '$name' (using [lindex $defaultColor 1])\n[string repeat " " 14]($recPracNames(model), Sec. 4.2.3, Table 2)"
    }
  }
  return $color
}

# -------------------------------------------------------------------------------
# write geometry for polyline annotations
proc x3dPolylinePMI {} {
  global ao gpmiPlacement opt placeAnchor placeOrigin recPracNames savedViewFile savedViewName savedViewNames
  global x3dColor x3dColorFile x3dColorsUsed x3dCoord x3dFile x3dIndex x3dIndexType x3dShape

  if {[catch {
    if {[info exists x3dCoord] || $x3dShape} {

# multiple saved views, write to individual files, collected in x3dViewpoint below
      set flist $x3dFile
      if {[llength $savedViewName] > 0} {
        set flist {}
        foreach svn $savedViewName {if {[info exists savedViewFile($svn)]} {lappend flist $savedViewFile($svn)}}
      }

      foreach f $flist {

# multiple saved view color
        if {$opt(gpmiColor) == 3 && [llength $savedViewNames] > 1} {
          if {![info exists x3dColorFile($f)]} {set x3dColorFile($f) [x3dSetColor $opt(gpmiColor) 1]}
          set x3dColor $x3dColorFile($f)
        }

        if {[string length $x3dCoord] > 0} {

# placeholder transform
          if {[string first "placeholder" $ao] != -1} {
            set transform [x3dTransform $gpmiPlacement(origin) $gpmiPlacement(axis) $gpmiPlacement(refdir) "annotation placeholder"]
            puts $f $transform
          }

# start shape
          if {$x3dColor != ""} {
            puts $f "<Shape><Appearance><Material diffuseColor='$x3dColor' emissiveColor='$x3dColor'></Material></Appearance>"
            lappend x3dColorsUsed $x3dColor

          } else {
            puts $f "<Shape><Appearance><Material diffuseColor='0 0 0' emissiveColor='0 0 0'></Material></Appearance>"
            errorMsg "Syntax Error: Missing PMI Presentation color for [formatComplexEnt $ao] (using black)\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 8.4, Figure 75)"
          }

# index and coordinates
          puts $f " <IndexedLineSet coordIndex='[string trim $x3dIndex]'>\n  <Coordinate point='[string trim $x3dCoord]'></Coordinate></IndexedLineSet></Shape>"

# end placeholder transform, add leader line
          if {[string first "placeholder" $ao] != -1} {
            puts $f "</Transform>"
            puts $f "<Shape><Appearance><Material emissiveColor='$x3dColor'></Material></Appearance>"
            puts $f " <IndexedLineSet coordIndex='0 1 -1'>\n  <Coordinate point='$placeOrigin $placeAnchor'></Coordinate></IndexedLineSet></Shape>"
          }

# end shape
        } elseif {$x3dShape} {
          puts $f "</Indexed$x3dIndexType\Set></Shape>"
        }
      }
      set x3dCoord ""
      set x3dIndex ""
      set x3dColor ""
      set x3dShape 0
    }
  } emsg3]} {
    errorMsg "ERROR writing polyline annotation graphics: $emsg3"
  }
  update idletasks
}

# -------------------------------------------------------------------------------
# write coordinate axes
proc x3dCoordAxes {size} {
  global x3dAxes x3dFile

# axes
  if {$x3dAxes} {
    puts $x3dFile "\n<!-- COORDINATE AXIS -->\n<Switch whichChoice='0' id='swAxes'><Group>"
    puts $x3dFile "<Shape id='x_axis'><Appearance><Material emissiveColor='1 0 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0 0 0 $size 0 0'></Coordinate></IndexedLineSet></Shape>"
    puts $x3dFile "<Shape id='y_axis'><Appearance><Material emissiveColor='0 .5 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0 0 0 0 $size 0'></Coordinate></IndexedLineSet></Shape>"
    puts $x3dFile "<Shape id='z_axis'><Appearance><Material emissiveColor='0 0 1'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0 0 0 0 0 $size'></Coordinate></IndexedLineSet></Shape>"

# xyz labels
    set tsize [trimNum [expr {$size*0.33}]]
    puts $x3dFile "<Transform translation='$size 0 0' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='1 0 0'></Material></Appearance><Text string='\"X\"'><FontStyle family='\"SANS\"'></FontStyle></Text></Shape></Billboard></Transform>"
    puts $x3dFile "<Transform translation='0 $size 0' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='0 .5 0'></Material></Appearance><Text string='\"Y\"'><FontStyle family='\"SANS\"'></FontStyle></Text></Shape></Billboard></Transform>"
    puts $x3dFile "<Transform translation='0 0 $size' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='0 0 1'></Material></Appearance><Text string='\"Z\"'><FontStyle family='\"SANS\"'></FontStyle></Text></Shape></Billboard></Transform>"
    puts $x3dFile "</Group></Switch>\n"
    set x3dAxes 0
  }
}

# -------------------------------------------------------------------------------
# get A2P3D origin, axis, refdir
proc x3dGetA2P3D {e0} {

  set origin "0 0 0"
  set axis   "0 0 1"
  set refdir "1 0 0"
  set debug 0

# a2p3d origin
  set a2 [[$e0 Attributes] Item [expr 2]]
  set e2 [$a2 Value]
  if {$e2 != ""} {
    set origin [vectrim [[[$e2 Attributes] Item [expr 2]] Value]]
    if {$debug} {errorMsg "      [$e2 Type] [$e2 P21ID] ([$a2 Name]) $origin" red}
  }

# a2p3d axis
  set a3 [[$e0 Attributes] Item [expr 3]]
  set e3 [$a3 Value]
  if {$e3 != ""} {
    set axis [[[$e3 Attributes] Item [expr 2]] Value]
    if {$debug} {errorMsg "      [$e3 Type] [$e3 P21ID] ([$a3 Name]) $axis" red}
  }

# a2p3d reference direction
  set a4 [[$e0 Attributes] Item [expr 4]]
  set e4 [$a4 Value]
  if {$e4 != ""} {
    set refdir [[[$e4 Attributes] Item [expr 2]] Value]
    if {$debug} {errorMsg "      [$e4 Type] [$e4 P21ID] ([$a4 Name]) $refdir" red}
  }

  return [list $origin $axis $refdir]
}

# -------------------------------------------------------------------------------
# generate transform
proc x3dTransform {origin axis refdir {text ""}} {

  set transform "<Transform"
  if {$origin != "0. 0. 0."} {append transform " translation='$origin'"}
  set rot [x3dGetRotation $axis $refdir $text]
  if {[lindex $rot 3] != 0} {append transform " rotation='$rot'"}
  append transform ">"
  return $transform
}

# -------------------------------------------------------------------------------
# set x3d color
proc x3dSetColor {type {mode 0}} {
  global idxColor x3dColorBrepAdjusted

# black
  if {$type == 1} {return "0 0 0"}

# random
  if {$type == 2 || $type == 3} {
    incr idxColor($mode)
    switch -- $idxColor($mode) {
      1 {set color "1 0 0"}
      2 {set color "0 0 1"}
      3 {set color "0 .5 0"}
      4 {set color "1 0 1"}
      5 {set color "0 .5 .5"}
      6 {set color ".5 .25 0"}
      7 {set color "0 0 0"}
      8 {set color "1 1 0"}
      9 {set color "1 1 1"}
    }
    if {$idxColor($mode) == 9} {set idxColor($mode) 0}
  }

# change color if is it the same as brep color
  if {[info exists x3dColorBrepAdjusted]} {
    set color1 $color
    if {$color1 == "0 .5 0"}  {set color1 "0 1 0"}
    if {$color1 == "0 .5 .5"} {set color1 "0 1 1"}
    if {$color1 == $x3dColorBrepAdjusted} {set color [x3dSetColor $type $mode]}
  }
  return $color
}

# -------------------------------------------------------------------------------------------------
# open X3DOM file
proc openX3DOM {{fn ""} {numFile 0}} {
  global lastX3DOM multiFile opt scriptName x3dFileName viz

# f3 is for opening last x3dom file with function key F3
  set f3 1
  if {$fn == ""} {
    set f3 0
    set ok 0

# check that there is a file to view
    if {[info exists x3dFileName]} {if {[file exists $x3dFileName]} {set ok 1}}
    if {$ok} {
      set fn $x3dFileName

# no file, show message
    } elseif {$opt(VIZPMI) || $opt(VIZTPG) || $opt(VIZFEA) || $opt(VIZBRP)} {
      if {$opt(XLSCSV) == "None"} {errorMsg "There is nothing in the STEP file to view based on the View selections (Options tab)."}
      return
    }
  }
  if {[file exists $fn] != 1} {return}
  if {![info exists multiFile]} {set multiFile 0}

  set open 0
  if {![info exists viz(BRP)]} {set viz(BRP) 0}
  if {$f3} {
    set open 1
  } elseif {($viz(PMI) || $viz(TPG) || $viz(FEA) || $viz(BRP)) && $fn != "" && $multiFile == 0} {
    if {$opt(XL_OPEN)} {set open 1}
  }

# open file (.html) in web browser
  set lastX3DOM $fn
  if {$open} {
    outputMsg "\nOpening View in the default Web Browser: [file tail $fn]" green
    catch {.tnb select .tnb.status}
    if {[catch {
      exec {*}[auto_execok start] "" [file nativename $fn]
    } emsg]} {
      if {[string first "UNC" $emsg] != -1} {set emsg [fixErrorMsg $emsg]}
      if {$emsg != ""} {
        errorMsg "ERROR opening View file ($emsg)\n Open [truncFileName [file nativename $fn]]\n in a web browser that supports x3dom https://www.x3dom.org"
      }
    }
    update idletasks
  } elseif {$numFile == 0 && [string first "STEP-File-Analyzer.exe" $scriptName] != -1} {
    outputMsg " Use F3 to open the View (see Options tab)" red
  }
}

# -------------------------------------------------------------------------------
# get saved view names
proc getSavedViewName {objEntity} {
  global draftModelCameraNames draftModelCameras entCount savedsavedViewNames savedViewName

# saved view name already saved
  if {[info exists savedsavedViewNames([$objEntity P21ID])]} {return $savedsavedViewNames([$objEntity P21ID])}

  set savedViewName {}
  set dmlist {}
  foreach dms [list draughting_model characterized_object_and_draughting_model characterized_representation_and_draughting_model characterized_representation_and_draughting_model_and_representation] {
    if {[info exists entCount($dms)]} {if {$entCount($dms) > 0} {lappend dmlist $dms}}
  }
  foreach dm $dmlist {
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
# script for switch node
proc x3dSwitchScript {name {name1 ""} {vp 0}} {
  global x3dFile

  if {$name1 == ""} {set name1 $name}
  regsub -all {\-} $name "_" name
  regsub -all {\-} $name1 "_" name1

  puts $x3dFile "\n<!-- [string toupper $name1] switch -->\n<script>function tog$name1\(choice){"
  puts $x3dFile " if (!document.getElementById('sw$name1').checked) {"
  puts $x3dFile "  document.getElementById('sw$name').setAttribute('whichChoice', -1);"
  puts $x3dFile " } else {"
  puts $x3dFile "  document.getElementById('sw$name').setAttribute('whichChoice', 0);"
  if {$vp} {puts $x3dFile "  document.getElementById('vp$name1').setAttribute('set_bind','true');"}
  puts $x3dFile " }"
  puts $x3dFile " document.getElementById('sw$name1').checked = !document.getElementById('sw$name1').checked;\n}</script>"
}

# -------------------------------------------------------------------------------
# generate x3d rotation (axis angle format) from axis2_placement_3d
proc x3dGetRotation {axis refdir {type ""}} {
  global x3dMsg

# check axis and refdir, both are congruent
  set msg ""
  if {[veclen [veccross $axis $refdir]] == 0} {
    set msg "Syntax Error: For an axis2_placement_3d"
    if {$type != ""} {append msg " related to '$type',"}
    append msg " the 'axis' and 'ref_direction' are congruent."

# one of the axes is zero length
  } elseif {[veclen $axis] == 0 || [veclen $refdir] == 0} {
    set msg "Syntax Error: For an axis2_placement_3d"
    if {$type != ""} {append msg " related to '$type',"}
    append msg " the magnitude of the 'axis' or 'ref_direction' is zero."
  }

  if {$msg != ""} {
    errorMsg $msg
    set msg1 [string range $msg 14 end]
    if {[lsearch $x3dMsg $msg1] == -1 && [string first "supplemental" $msg1] != -1} {lappend x3dMsg $msg1}
  }

# construct rotation matrix u, must normalize to use with quaternion
  set u3 [vecnorm $axis]
  set u1 [vecnorm [vecsub $refdir [vecmult $u3 [vecdot $refdir $u3]]]]
  set u2 [vecnorm [veccross $u3 $u1]]

# extract quaternion
  if {[lindex $u1 0] >= 0.0} {
    set tmp [expr {[lindex $u2 1] + [lindex $u3 2]}]
    if {$tmp >=  0.0} {
      set q(0) [expr {[lindex $u1 0] + $tmp + 1.}]
      set q(1) [expr {[lindex $u3 1] - [lindex $u2 2]}]
      set q(2) [expr {[lindex $u1 2] - [lindex $u3 0]}]
      set q(3) [expr {[lindex $u2 0] - [lindex $u1 1]}]
    } else {
      set q(0) [expr {[lindex $u3 1] - [lindex $u2 2]}]
      set q(1) [expr {[lindex $u1 0] - $tmp + 1.}]
      set q(2) [expr {[lindex $u2 0] + [lindex $u1 1]}]
      set q(3) [expr {[lindex $u1 2] + [lindex $u3 0]}]
    }
  } else {
    set tmp [expr {[lindex $u2 1] - [lindex $u3 2]}]
    if {$tmp >= 0.0} {
      set q(0) [expr {[lindex $u1 2] - [lindex $u3 0]}]
      set q(1) [expr {[lindex $u2 0] + [lindex $u1 1]}]
      set q(2) [expr {1. - [lindex $u1 0] + $tmp}]
      set q(3) [expr {[lindex $u3 1] + [lindex $u2 2]}]
    } else {
      set q(0) [expr {[lindex $u2 0] - [lindex $u1 1]}]
      set q(1) [expr {[lindex $u1 2] + [lindex $u3 0]}]
      set q(2) [expr {[lindex $u3 1] + [lindex $u2 2]}]
      set q(3) [expr {1. - [lindex $u1 0] - $tmp}]
    }
  }

# normalize quaternion
  set lenq [expr {sqrt($q(0)*$q(0) + $q(1)*$q(1) + $q(2)*$q(2) + $q(3)*$q(3))}]
  if {$lenq != 0.} {
    foreach i {0 1 2 3} {set q($i) [expr {$q($i) / $lenq}]}
  } else {
    foreach i {0 1 2 3} {set q($i) 0.}
  }

# convert from quaterion to x3d rotation
  set rotation_changed {0 1 0 0}
  set angle [expr {acos($q(0))*2.0}]
  if {$angle != 0.} {
    set sina [expr {sin($angle*0.5)}]
    set axm 0.
    foreach i {0 1 2} {
      set i1 [expr {$i+1}]
      set ax [expr {-$q($i1) / $sina}]
      lset rotation_changed $i $ax
      set axa [expr {abs($ax)}]
      if {$axa > $axm} {set axm $axa}
    }
    if {$axm > 0. && $axm < 1.} {
      foreach i {0 1 2} {lset rotation_changed $i [expr {[lindex $rotation_changed $i]/$axm}]}
    }
    lset rotation_changed 3 $angle
    foreach i {0 1 2 3} {lset rotation_changed $i [trimNum [lindex $rotation_changed $i] 4]}
  }
  return $rotation_changed
}
