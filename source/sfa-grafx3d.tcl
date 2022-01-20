# start x3dom file for non-FEM graphics
proc x3dFileStart {} {
  global ap242XML cadSystem entCount gen localName opt stepAP timeStamp viz writeDir writeDirType x3dFile x3dFileName x3dFiles
  global x3dFileSave x3dFileNameSave x3dHeight x3dMax x3dMin x3dPartClick x3dStartFile x3dTitle x3dViewOK x3dWidth

  if {!$gen(View)} {return}
  set x3dViewOK 1
  if {$x3dStartFile == 0} {return}

  if {![info exists stepAP]} {set stepAP [getStepAP $localName]}
  if {[string first "IFC" $stepAP] == 0 || [string first "ISO" $stepAP] == 0 || $stepAP == "AP210" || \
      $stepAP == "CUTTING_TOOL_SCHEMA_ARM" || $stepAP == "STRUCTURAL_FRAME_SCHEMA"} {
    set msg "The Viewer only works with STEP AP203, AP209, AP214, AP238, and AP242 files.  See Help > Support STEP APs"
    if {$stepAP == "STRUCTURAL_FRAME_SCHEMA"} {append msg "\n Use the NIST SteelVis viewer for CIS/2 files."}
    errorMsg $msg
    set x3dViewOK 0
    return
  }

  set x3dStartFile 0
  checkTempDir

# x3d output file name
  set x3dir [file rootname $localName]
  if {$ap242XML} {append x3dir "-stpx"}
  if {$opt(writeDirType) == 2} {set x3dir [file join $writeDir [file rootname [file tail $localName]]]}
  set x3dFileName $x3dir\-sfa.html
  set x3dFile [open $x3dFileName w]
  catch {file delete -force -- $x3dFileName}

# start x3d file
  set title [encoding convertto utf-8 [file tail $localName]]
  if {$stepAP != "" && [string range $stepAP 0 1] == "AP"} {append title " | $stepAP"}
  puts $x3dFile "<!DOCTYPE html>\n<html>\n<head>\n<title>$title</title>\n<base target=\"_blank\">\n<meta http-equiv='Content-Type' content='text/html;charset=utf-8'/>"

# use x3dom 1.8.1 because 1.8.2 breaks transparency
  puts $x3dFile "<link rel='stylesheet' type='text/css' href='https://www.x3dom.org/download/1.8.1/x3dom.css'/>\n<script type='text/javascript' src='https://www.x3dom.org/download/1.8.1/x3dom.js'></script>"
  #puts $x3dFile "<link rel='stylesheet' type='text/css' href='https://www.x3dom.org/x3dom/release/x3dom.css'/>\n<script type='text/javascript' src='https://www.x3dom.org/x3dom/release/x3dom.js'></script>"

# scripts and functions for selecting parts
  set callback {}
  set x3dPartClick 0
  if {[info exists viz(PART)]} {if {$viz(PART)} {set x3dPartClick 1; lappend callback "shape"}}
  if {!$x3dPartClick && $opt(viewPart)} {set x3dPartClick 1; lappend callback "shape"}
  if {$viz(TESSPART)} {set x3dPartClick 1; lappend callback "group"}

  if {$x3dPartClick} {
    puts $x3dFile "\n<script type='text/javascript' src='https://code.jquery.com/jquery-2.1.0.min.js'></script>"
    puts $x3dFile "<script>\n// Mark selection point"
    puts $x3dFile "function handleGroupClick(event) {\$('#marker').attr('translation', event.hitPnt);}"

    puts $x3dFile "// Handle click on '[join $callback]'"
    foreach type $callback {puts $x3dFile "function handleSingleClick($type) {\$('#clickedObject').html(\$($type).attr('id'));}"}

    puts $x3dFile "// Add onclick callback to every '[join $callback]'"
    puts $x3dFile "\$(document).ready(function() \{"
    foreach type $callback {
      puts $x3dFile "  \$('$type').each(function() \{"
      puts $x3dFile "    \$(this).attr('onclick','handleSingleClick(this)');"
      puts $x3dFile "  \});"
    }
    puts $x3dFile "\});"
    puts $x3dFile "</script>"
  }
  puts $x3dFile "</head>"

# x3d title
  set x3dTitle [encoding convertto utf-8 [file tail $localName]]
  if {$stepAP != "" && [string range $stepAP 0 1] == "AP"} {append x3dTitle "&nbsp;&nbsp;&nbsp;$stepAP"}
  if {[info exists timeStamp]} {
    if {$timeStamp != ""} {
      set ts [fixTimeStamp $timeStamp]
      append x3dTitle "&nbsp;&nbsp;&nbsp;$ts"
    }
  }
  if {[info exists cadSystem]} {
    if {$cadSystem != ""} {
      regsub -all "_" $cadSystem " " cs
      append x3dTitle "&nbsp;&nbsp;&nbsp;$cs"
    }
  }
  puts $x3dFile "\n<body><font face=\"sans-serif\">\n<h3>$x3dTitle</h3>"
  puts $x3dFile "\n<table>"

# messages above the x3d
  set msg ""
  if {$viz(PMI)} {
    append msg "$viz(PMIMSG)  "
  } elseif {$opt(viewPMI)} {
    if {[string first "Some Graphical PMI" $viz(PMIMSG)] == 0} {append msg "The STEP file contains only Semantic PMI and no supported Graphical PMI.  "}
  }
  if {$msg != ""} {puts $x3dFile "<tr><td valign='top' width='85%'>[string trim $msg]</td><td></td></tr>"}

# x3d window size
  puts $x3dFile "<tr><td valign='top' width='85%'>\n<table><tr><td style='border:1px solid black'>\n<noscript>JavaScript must be enabled in the web browser</noscript>"
  set x3dHeight 900
  set x3dWidth [expr {int($x3dHeight*1.78)}]
  catch {
    set x3dHeight [expr {int([winfo screenheight .]*0.85)}]
    set x3Width [expr {int($x3dHeight*[winfo screenwidth .]/[winfo screenheight .])}]
  }

# start x3d with flat-to-screen hud for viewpoint text
  puts $x3dFile "\n<X3D id='x3d' showStat='false' showLog='false' x='0px' y='0px' width='$x3dWidth' height='$x3dHeight'>"
  set txt "<div id='HUDs_Div'><div class='group' style='margin:2px; margin-top:0px; padding:4px; background-color:rgba(0,0,0,1.); position:absolute; float:center; z-index:1000;'>Viewpoint: <span id='clickedView'></span>"
  if {$x3dPartClick} {append txt "<br>Part: <span id='clickedObject'></span>"}
  append txt "</div></div>"
  puts $x3dFile $txt
  puts $x3dFile "<Scene DEF='scene'>"
  puts $x3dFile "<!-- X3D generated by the NIST STEP File Analyzer and Viewer [getVersion] -->"

# read tessellated geometry separately because of IFCsvr limitations
  if {($viz(PMI) && [info exists entCount(tessellated_annotation_occurrence)]) || $viz(TESSPART)} {tessReadGeometry}
  outputMsg " Writing Viewer file to: [truncFileName [file nativename $x3dFileName]]" blue

# coordinate min, max, center
  if {$x3dMax(x) != -1.e8} {
    foreach xyz {x y z} {
      set delt($xyz) [expr {$x3dMax($xyz)-$x3dMin($xyz)}]
      set xyzcen($xyz) [format "%.4f" [expr {0.5*$delt($xyz) + $x3dMin($xyz)}]]
    }
    set maxxyz [expr {max($delt(x),$delt(y),$delt(z))}]
  }

# keep X3D file, write saved views and brep geom
  catch {unset x3dFiles}
  lappend x3dFiles $x3dFile
  if {$opt(x3dSave)} {
    set x3dFileNameSave "[file rootname $localName].x3d"
    catch {file delete -force -- $x3dFileNameSave}
    set x3dFileSave [open $x3dFileNameSave w]
    lappend x3dFiles $x3dFileSave
    puts $x3dFileSave "<?xml version='1.0' encoding='UTF-8'?>\n<X3D>\n<head><meta name='Generator' content='NIST STEP File and Viewer [getVersion]'/></head>\n<Scene>"
  }
  update idletasks
}

# -------------------------------------------------------------------------------
# finish x3d file, write tessellated edges, PMI saved view geometry, set viewpoints, add navigation and background color, and close x3dom file
proc x3dFileEnd {} {
  global ao ap242XML brepFile brepFileName datumTargetView entCount grayBackground matTrans maxxyz nistName nsketch numTessColor opt parts partstg
  global samplingPoints savedViewButtons savedViewFile savedViewFileName savedViewItems savedViewNames savedViewpoint savedViewVP sphereDef stepAP
  global tessCoord tessEdges tessPartFile tessPartFileName tessRepo tsName viz x3dApps x3dAxes x3dBbox x3dCoord x3dFile x3dFileNameSave x3dFiles
  global x3dFileSave x3dIndex x3dMax x3dMin x3dMsg x3dPartClick x3dParts x3dShape x3dStartFile x3dTessParts x3dTitle x3dViewOK viewsWithPMI

  if {!$x3dViewOK} {
    foreach var [list x3dCoord x3dFile x3dIndex x3dMax x3dMin x3dShape x3dStartFile] {catch {unset -- $var}}
    return
  }

# PMI is already written to file, generate b-rep part geometry
  if {!$ap242XML} {
    set viz(PART) 0
    if {$opt(viewPart)} {
      if {!$opt(partOnly)} {
        set ok 0
        foreach item [list advanced_brep_shape_representation geometrically_bounded_surface_shape_representation geometrically_bounded_wireframe_shape_representation manifold_solid_brep manifold_surface_shape_representation shell_based_surface_model document_file] {
          if {[info exists entCount($item)]} {set ok 1}
        }
      } else {
        set ok 1
      }
      if {$ok} {x3dBrepGeom}
    }
  }

# coordinate min, max, center
  foreach idx {x y z} {
    if {$x3dMax($idx) == -1.e8 || $x3dMax($idx) > 1.e8} {set x3dMax($idx) 500.}
    if {$x3dMin($idx) == 1.e8 || $x3dMin($idx) < -1.e8} {set x3dMin($idx) -500.}
    set delt($idx) [expr {$x3dMax($idx)-$x3dMin($idx)}]
    set xyzcen($idx) [trimNum [format "%.4f" [expr {0.5*$delt($idx) + $x3dMin($idx)}]]]
  }
  set maxxyz [expr {max($delt(x),$delt(y),$delt(z))}]

# -------------------------------------------------------------------------------
# write tessellated edges
  set viz(TESSEDGE) 0
  if {[info exists tessEdges]} {
    puts $x3dFile "\n<!-- TESSELLATED EDGES -->\n<Switch whichChoice='0' id='swTED'><Group>"
    foreach cid [array names tessEdges] {
      puts $x3dFile "<Shape><Appearance><Material emissiveColor='0 0 0'/></Appearance>"
      puts $x3dFile " <IndexedLineSet coordIndex='[join $tessEdges($cid)]'>"
      puts $x3dFile "  <Coordinate DEF='coord$cid' point='$tessCoord($cid)'/></IndexedLineSet></Shape>"
    }
    puts $x3dFile "</Group></Switch>"
    set viz(TESSEDGE) 1
    unset tessEdges
  }

# -------------------------------------------------------------------------------
# holes
  set ok 0
  set viz(HOLE) 0
  set sphereDef {}
  foreach ent [list basic_round_hole_occurrence counterbore_hole_occurrence counterdrill_hole_occurrence countersink_hole_occurrence spotface_occurrence] {
    if {[info exists entCount($ent)]} {set ok 1}
    set ent1 "$ent\_in_assembly"
    if {[info exists entCount($ent1)]} {set ok 1}
  }
  if {$ok} {x3dHoles $maxxyz}

# -------------------------------------------------------------------------------
# supplemental geometry
  set viz(SUPPGEOM) 0
  if {[info exists entCount(constructive_geometry_representation)]} {x3dSuppGeom $maxxyz}

# -------------------------------------------------------------------------------
# datum targets
  set viz(DTMTAR) 0
  if {[info exists datumTargetView]} {
    x3dDatumTarget $maxxyz
  } elseif {[info exists entCount(placed_datum_target_feature)] || [info exists entCount(datum_target)]} {
    set msg " Datum targets cannot be shown without "
    if {$opt(xlFormat) != "Excel"} {
      append msg "generating a spreadsheet"
    } elseif {!$opt(PMISEM) || $opt(PMISEMDIM)} {
      append msg "selecting Semantic PMI in the Analyzer section"
    }
    append msg ".  See Help > Viewer > Datum Targets"
    outputMsg $msg red
  }
  catch {unset datumTargetView}

# -------------------------------------------------------------------------------
# points
  set viz(POINTS) 0
  set pointsLabel ""

# validation property sampling points
  if {[info exists samplingPoints]} {
    set viz(POINTS) 1
    set pointsLabel "Cloud of Points"
    puts $x3dFile "\n<!-- [string toupper $pointsLabel] -->"
    puts $x3dFile "<Switch whichChoice='0' id='swPoints'>"
    puts $x3dFile "<Shape><Appearance><Material emissiveColor='0 0 1'/></Appearance><PointSet><Coordinate point='$samplingPoints'/></PointSet></Shape>\n</Switch>"

# point cloud
  } elseif {[info exists entCount(point_cloud_dataset)]} {
    if {$entCount(point_cloud_dataset) > 0} {
      tessReadGeometry 2
      set viz(POINTS) 1
      set pointsLabel "Point Cloud"
      puts $x3dFile "\n<!-- [string toupper $pointsLabel] -->"
      puts $x3dFile "<Switch whichChoice='0' id='swPoints'><Group>"
      foreach idx [array names tessCoord] {
        puts $x3dFile "<Shape><Appearance><Material emissiveColor='0 0 1'/></Appearance><PointSet><Coordinate point='$tessCoord($idx)'/></PointSet></Shape>"
      }
      puts $x3dFile "</Group></Switch>"
    }
  }

# -------------------------------------------------------------------------------
# write any PMI saved view geometry for multiple saved views
  set torg ""
  set savedViewButtons {}
  if {[info exists savedViewNames]} {
    if {[llength $savedViewNames] > 0} {
      for {set i 0} {$i < [llength $savedViewNames]} {incr i} {
        set svn [lindex $savedViewNames $i]
        set svnfn "View$i"
        catch {close $savedViewFile($svnfn)}
        if {[info exists savedViewFileName($svnfn)]} {
          if {[file size $savedViewFileName($svnfn)] > 0} {
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
              set svn2 "Missing name"
              set svMap($svn2) $svn2
            }
            lappend savedViewButtons $svn2
            foreach xf $x3dFiles {puts $xf "\n<!-- SAVED VIEW$i $svn2 -->\n<Switch whichChoice='0' id='sw$svnfn'><Group>"}

# show camera viewpoint
            if {[info exists savedViewpoint($svn2)]} {x3dSavedViewpoint $svn2}

# get saved view graphics from file
            if {$svWrite} {
              set lastTransform ""
              set f [open $savedViewFileName($svnfn) r]
              while {[gets $f line] >= 0} {

# check for similar transforms
                if {[string first "<Transform" $line] == -1 && [string first "</Transform>" $line] == -1} {
                  foreach xf $x3dFiles {puts $xf $line}
                } elseif {[string first "<Transform" $line] == 0} {
                  if {$line != $lastTransform} {
                    if {$lastTransform != ""} {foreach xf $x3dFiles {puts $xf "</Transform>"}}
                    foreach xf $x3dFiles {puts $xf $line}
                    set lastTransform $line
                  }
                  set torg "Transform"
                }
                if {[string first "<Group" $line]   == 0} {set torg "Group"}
              }
              if {$lastTransform != ""} {foreach xf $x3dFiles {puts $xf "</Transform>"}}

              close $f
              catch {unset savedViewFile($svnfn)}

# duplicate saved views
            } else {
              foreach xf $x3dFiles {puts $xf "<!-- SAME AS $svMap($svn) -->"}
              errorMsg " Two or more Saved Views have the exact same graphical PMI" red
              set torg ""
            }

# ending group and switch
            foreach xf $x3dFiles {
              if {$torg == "Group"} {puts $xf "</Group>"}
              puts $xf "</Group></Switch>"
            }
            set torg ""
          } else {
            catch {close $savedViewFile($svnfn)}
          }
        }
        catch {file delete -force -- $savedViewFileName($svnfn)}
      }
    }
  }

# viewpoints without PMI
  if {![info exists savedViewVP]} {
    if {([info exists entCount(camera_model_d3)] || [info exists entCount(camera_model_d3_multi_clipping)]) && \
         [info exists entCount(view_volume)] && [info exists entCount(planar_box)]} {
      for {set i 0} {$i < [llength $savedViewNames]} {incr i} {
        set svn [lindex $savedViewNames $i]
        if {[info exists savedViewpoint($svn)]} {x3dSavedViewpoint $svn}
      }
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
  set oktpg 0
  if {[info exists tessPartFile]} {
    if {[file size $tessPartFileName] > 0} {
      set oktpg 1
    } else {
      set viz(TESSPART) 0
    }
  }
  if {$oktpg} {
    foreach xf $x3dFiles {puts $xf "\n<!-- TESSELLATED PART GEOMETRY -->\n<Switch whichChoice='0' id='swTPG'><Group>"}
    catch {close $tessPartFile}
    set f [open $tessPartFileName r]
    set npart(TESSPART) -1

# for parts with a transform, append to lines for each part name and transform to group by transform
    if {![info exists tessRepo]} {set tessRepo 0}
    if {$tessRepo} {
      set tgparts {}
      while {[gets $f line] >= 0} {
        if {[string first "<!--" $line] == 0} {
          set part $line
          lappend tgparts $line
        } elseif {[string first "<Transform" $line] == 0} {
          set transform $line
        } elseif {$line != "</Transform>"} {
          if {![info exists transform]} {set transform "<Transform>"}
          lappend lines($part,$transform) $line
        }
      }
      close $f

# write parts for each transform
      set items [lreverse [array names lines]]
      foreach part $tgparts {
        foreach xf $x3dFiles {puts $xf $part}

# set partname
        set partname [string range $part [string first " " $part]+1 [string last " " $part]-1]
        if {[string first "TESSELLATED" $partname] == 0} {set partname [string tolower [string range $partname 12 end]]}
        incr npart(TESSPART)
        set x3dTessParts($partname) $npart(TESSPART)

# switch if more than one part
        if {[llength $tgparts] > 1} {
          regsub -all "'" $partname "\"" txt
          foreach xf $x3dFiles {puts $xf "<Switch id='swTessPart$npart(TESSPART)' whichChoice='0'><Group id='$txt'>"}
        }

# write
        foreach item $items {
          if {[string first $part $item] == 0} {
            set transform [string range $item [string last "," $item]+1 end]
            if {$transform != "<Transform>"} {foreach xf $x3dFiles {puts $xf $transform}}
            foreach line $lines($item) {foreach xf $x3dFiles {puts $xf $line}}
            if {$transform != "<Transform>"} {foreach xf $x3dFiles {puts $xf "</Transform>"}}
          }
        }
        if {[llength $tgparts] > 1} {foreach xf $x3dFiles {puts $xf "</Group></Switch>"}}
      }

# no grouping if no transforms, add switch
    } else {
      while {[gets $f line] >= 0} {
        if {[string first "<!--" $line] == 0} {
          set partname [string range $line [string first " " $line]+1 [string last " " $line]-1]
          if {[string first "TESSELLATED" $partname] == 0} {set partname [string tolower [string range $partname 12 end]]}
          incr npart(TESSPART)
          set x3dTessParts($partname) $npart(TESSPART)
          if {$npart(TESSPART) > 0} {foreach xf $x3dFiles {puts $xf "</Group></Switch>"}}
          regsub -all "'" $partname "\"" txt
          foreach xf $x3dFiles {puts $xf "$line\n<Switch id='swTessPart$npart(TESSPART)' whichChoice='0'><Group id='$txt'>"}
        } else {
          foreach xf $x3dFiles {puts $xf $line}
        }
      }
      foreach xf $x3dFiles {puts $xf "</Group></Switch>"}
      close $f
    }

# close overall switch
    foreach xf $x3dFiles {puts $xf "</Group></Switch>"}
  }
  catch {file delete -force -- $tessPartFileName}
  foreach var {tessPartFile tessPartFileName tsName} {catch {unset -- $var}}

# -------------------------------------------------------------------------------
# part geometry
  if {![info exists x3dFiles]} {set x3dFiles [list $x3dFile]}
  if {$viz(PART)} {

# bounding box
    if {[info exists x3dBbox]} {
      if {$x3dBbox != ""} {
        set p(0) "$x3dMin(x) $x3dMin(y) $x3dMin(z)"
        set p(1) "$x3dMax(x) $x3dMin(y) $x3dMin(z)"
        set p(2) "$x3dMax(x) $x3dMax(y) $x3dMin(z)"
        set p(3) "$x3dMin(x) $x3dMax(y) $x3dMin(z)"
        set p(4) "$x3dMin(x) $x3dMin(y) $x3dMax(z)"
        set p(5) "$x3dMax(x) $x3dMin(y) $x3dMax(z)"
        set p(6) "$x3dMax(x) $x3dMax(y) $x3dMax(z)"
        set p(7) "$x3dMin(x) $x3dMax(y) $x3dMax(z)"
        puts $x3dFile "\n<!-- BOUNDING BOX -->"
        puts $x3dFile "<Switch whichChoice='-1' id='swBbox'><Group>"
        puts $x3dFile " <Shape><Appearance><Material emissiveColor='0 0 0'/></Appearance>"
        puts $x3dFile "  <IndexedLineSet coordIndex='0 1 2 3 0 -1 4 5 6 7 4 -1 0 4 -1 1 5 -1 2 6 -1 3 7 -1'><Coordinate point='$p(0) $p(1) $p(2) $p(3) $p(4) $p(5) $p(6) $p(7)'/></IndexedLineSet></Shape>"
        puts $x3dFile "</Group></Switch>"
      }
    }

# add b-rep part geometry from temp file
    if {[info exists brepFileName]} {
      if {[file exists $brepFileName]} {
        close $brepFile
        if {[file size $brepFileName] > 0} {
          foreach xf $x3dFiles {
            set brepFile [open $brepFileName r]
            while {[gets $brepFile line] >= 0} {puts $xf $line}
            close $brepFile
          }
          if {!$opt(DEBUGX3D)} {catch {file delete -force -- $brepFileName}}
        }
      }
    }
  }

# -------------------------------------------------------------------------------
# viewpoints
  foreach xf $x3dFiles {puts $xf "\n<!-- VIEWPOINTS -->"}

# default
  set cor "centerOfRotation='$xyzcen(x) $xyzcen(y) $xyzcen(z)'"
  set xmin [trimNum [expr {$x3dMin(x) - 1.4*max($delt(y),$delt(z))}]]
  set xmax [trimNum [expr {$x3dMax(x) + 1.4*max($delt(y),$delt(z))}]]
  set ymin [trimNum [expr {$x3dMin(y) - 1.4*max($delt(x),$delt(z))}]]
  set ymax [trimNum [expr {$x3dMax(y) + 1.4*max($delt(x),$delt(z))}]]
  set zmin [trimNum [expr {$x3dMin(z) - 1.4*max($delt(x),$delt(y))}]]
  set zmax [trimNum [expr {$x3dMax(z) + 1.4*max($delt(x),$delt(y))}]]

# z axis up
  set sfastr ""
  if {[info exists savedViewVP]} {set sfastr " (SFA)"}
  foreach xf $x3dFiles {puts $xf "<Viewpoint id='Front 1$sfastr' position='$xyzcen(x) $ymin $xyzcen(z)' $cor orientation='1 0 0 1.5708'></Viewpoint>"}

# other front/side/top/isometric viewpoints if not saved views
  if {![info exists savedViewVP]} {
    puts $x3dFile "<Viewpoint id='Side 1' position='$xmax $xyzcen(y) $xyzcen(z)' $cor orientation='1 1 1 2.094'></Viewpoint>"
    puts $x3dFile "<Viewpoint id='Top 1' position='$xyzcen(x) $xyzcen(y) $zmax' $cor></Viewpoint>"
    puts $x3dFile "<Viewpoint id='Front 2' position='$xyzcen(x) $xyzcen(y) $zmin' $cor orientation='0 1 0 3.1416'></Viewpoint>"
    puts $x3dFile "<Viewpoint id='Side 2' position='$xmax $xyzcen(y) $xyzcen(z)' $cor orientation='0 1 0 1.5708'></Viewpoint>"
    puts $x3dFile "<Viewpoint id='Top 2' position='$xyzcen(x) $ymax $xyzcen(z)' $cor orientation='1 0 0 -1.5708'></Viewpoint>"

    puts $x3dFile "<Transform rotation='0 0 1 0.5236'><Transform rotation='1 0 0 -0.5236'>"
    puts $x3dFile " <Viewpoint id='Isometric' position='$xyzcen(x) $ymin $xyzcen(z)' $cor orientation='1 0 0 1.5708'></Viewpoint>"
    puts $x3dFile "</Transform></Transform>"

# saved views and other viewpoints
  } else {
    outputMsg " Use PageDown in the Viewer for ([llength $savedViewVP]) Viewpoints" red
    foreach xf $x3dFiles {foreach line $savedViewVP {puts $xf $line}}
  }

# orthographic
  set fov [trimNum [expr {0.55*max($delt(x),$delt(z))}]]
  puts $x3dFile "<OrthoViewpoint id='Orthographic$sfastr' position='$xyzcen(x) [trimNum [expr {$x3dMin(y) - 1.4*max($delt(x),$delt(z))}]] $xyzcen(z)' $cor orientation='1 0 0 1.5708' fieldOfView='\[-$fov,-$fov,$fov,$fov\]'></OrthoViewpoint>"

# background color, default gray
  set skyBlue ".53 .81 .92"
  set bgcheck1 ""
  set bgcheck2 ""
  set bgcheck3 "checked"
  set bgcolor ".8 .8 .8"

# blue background
  if {!$viz(PMI) && !$viz(SUPPGEOM) && !$viz(DTMTAR) && !$viz(HOLE) && [string first "AP209" $stepAP] == -1} {
    set bgcheck2 "checked"
    set bgcheck3 ""
    set bgcolor $skyBlue

# white background
  } elseif {![info exists grayBackground]} {
    set bgcheck1 "checked"
    set bgcheck3 ""
    set bgcolor "1 1 1"
  }

# blue background w/o sketch geometry controlled by CSS instead of BACKGROUND node
  set bgcss 0
  if {![info exists nsketch]} {set nsketch -1}
  if {$bgcolor == $skyBlue && $nsketch == -1} {set bgcss 1}
  catch {unset grayBackground}

# background, navigation, world info
  foreach xf $x3dFiles {puts $xf "\n<!-- BACKGROUND, NAVIGATION, WORLD INFO -->"}
  if {!$bgcss} {puts $x3dFile "<Background id='BG' skyColor='$bgcolor'/>"}
  if {$opt(x3dSave)} {puts $x3dFileSave "<Background skyColor='0.9 0.9 0.9'/>"}

  puts $x3dFile "<NavigationInfo type='\"EXAMINE\",\"ANY\"'/>"

  regsub -all "&nbsp;" $x3dTitle " " title
  foreach xf $x3dFiles {puts $xf "<WorldInfo title='$title' info='Generated by the NIST STEP File Analyzer and Viewer [getVersion]'/>"}
  foreach xf $x3dFiles {puts $xf "</Scene></X3D>"}
  puts $x3dFile "</td></tr></table>"

# close saved x3d file
  if {$opt(x3dSave)} {
    close $x3dFileSave
    set x3dfn "[file rootname $x3dFileNameSave]-sfa.x3d"
    catch {file delete -force -- $x3dfn}
    catch {
      file copy -force -- $x3dFileNameSave $x3dfn
      outputMsg " Saving X3D file: [truncFileName [file nativename $x3dfn]]" blue
      file delete -force -- $x3dFileNameSave
    }
  }

# credits
  set str "\nGenerated by the <a href=\"https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer\">NIST STEP File Analyzer and Viewer [getVersion]</a>"
  append str "&nbsp;&nbsp;[clock format [clock seconds] -format "%d %b %G %H:%M"]"
  append str "&nbsp;&nbsp;<a href=\"https://www.nist.gov/disclaimer\">NIST Disclaimer</a>"
  puts $x3dFile $str

# -------------------------------------------------------------------------------
# start right column
  puts $x3dFile "</td>\n\n<!-- RIGHT COLUMN BUTTONS -->\n<td valign='top'>"

# for NIST model - link to drawing
  if {[info exists nistName]} {
    if {$nistName != ""} {
      regsub -all "_" $nistName "-" name
      set name "nist-cad-model-[string range $name 5 end]"
      puts $x3dFile "<a href=\"https://www.nist.gov/document/$name\">NIST Test Case Drawing</a><p>"
    }
  }

# part geometry, sketch geometry, edges checkboxes
  set pcb 0
  if {$viz(PART) && !$ap242XML} {
    puts $x3dFile "\n<!-- Part geometry checkbox -->\n<input type='checkbox' checked onclick='togPRT(this.value)'/>Part Geometry"
    if {[info exists nsketch]} {
      if {$nsketch > -1} {puts $x3dFile "<!-- Sketch geometry checkbox -->\n<br><input type='checkbox' checked onclick='togSKH(this.value)'/>Sketch Geometry"}
      if {$nsketch > 1000} {errorMsg " Sketch geometry ([expr {$nsketch+1}]) might take too long to view.  Turn off Sketch and regenerate the View."}
    }
    if {$opt(partEdges) && $viz(EDGE)} {puts $x3dFile "<!-- Edges checkbox -->\n<br><input type='checkbox' checked onclick='togEDG(this.value)' id='swEDG'/>Edges"}
  }

# part checkboxes
  if {$viz(PART)} {
    if {[info exists x3dParts]} {if {[llength [array names x3dParts]] > 1} {x3dPartCheckbox "Part"; set pcb 1}}
    puts $x3dFile "<p>"
  }

# tessellated part geometry checkbox
  if {$viz(TESSPART)} {
    if {$pcb} {if {[info exists x3dTessParts]} {if {[llength [array names x3dTessParts]] > 1} {puts $x3dFile "<hr>"}}}
    puts $x3dFile "\n<!-- Tessellated part geometry checkbox -->\n<input type='checkbox' checked onclick='togTPG(this.value)'/>Tessellated Part Geometry"
    if {$viz(TESSEDGE)} {puts $x3dFile "<!-- Tessellated edges checkbox -->\n<br><input type='checkbox' checked onclick='togTED(this.value)'/>Edges"}

    if {[info exists entCount(next_assembly_usage_occurrence)] || [info exists entCount(repositioned_tessellated_item_and_tessellated_geometric_set)]} {
      puts $x3dFile "<p><font size='-1'>Tessellated Parts in an assembly might be in the wrong position and orientation or be missing.</font>"
    }

# tessellated part checkboxes
    if {[info exists x3dTessParts]} {if {[llength [array names x3dTessParts]] > 1} {x3dPartCheckbox "Tess"}}
    puts $x3dFile "<p>"
  }

# supplemental geometry checkbox
  if {$viz(SUPPGEOM)} {
    puts $x3dFile "\n<!-- Supplemental geometry checkbox -->\n<input type='checkbox' checked onclick='togSMG(this.value)'/>Supplemental Geometry"
    if {$viz(DTMTAR)} {
      puts $x3dFile "<br>"
    } else {
      puts $x3dFile "<p>"
    }
  }

# datum targets checkbox
  if {$viz(DTMTAR)} {
    puts $x3dFile "\n<!-- Datum targets checkbox -->\n<input type='checkbox' checked onclick='togDTR(this.value)'/>Datum Targets"
    if {$viz(POINTS)} {
      puts $x3dFile "<br>"
    } else {
      puts $x3dFile "<p>"
    }
  }

# sampling points or point cloud
  if {$viz(POINTS)} {
    puts $x3dFile "\n<!-- $pointsLabel checkbox -->\n<input type='checkbox' checked onclick='togPoints(this.value)'/>$pointsLabel"
    if {$viz(HOLE)} {
      puts $x3dFile "<br>"
    } else {
      puts $x3dFile "<p>"
    }
  }

# holes checkbox
  if {$viz(HOLE)} {
    puts $x3dFile "\n<!-- Holes checkbox -->\n<input type='checkbox' checked onclick='togHole(this.value)'/>Holes<p>"
  }

# for PMI annotations - checkboxes for toggling saved view PMI
  if {$viz(PMI) && [llength $savedViewButtons] > 0} {
    set sv 1
    if {[llength $savedViewButtons] == 1 && [lindex $savedViewNames 0] == "Not in a Saved View"} {
      set sv 0
      set name "Graphical PMI"
      set savedViewButtons [list $name]
      set savedViewNames $savedViewButtons
      set svMap($name) $name
    }
    puts $x3dFile "\n<!-- Saved view checkboxes -->"
    if {$sv} {puts $x3dFile "Saved View Graphical PMI"}
    if {[info exists savedViewVP] && $opt(viewPMIVP)} {puts $x3dFile "<br><font size='-1'>(PageDown to switch Saved Views)</font>"}

    foreach svn $savedViewButtons {
      set str ""
      if {$sv} {append str "<br>"}
      set id [lsearch $savedViewNames $svn]
      append str "<input type='checkbox' id='cbView$id' checked onclick='togView$id\(this.value)'/>$svn"
      puts $x3dFile $str
    }
  }

# FEM checkboxes
  if {$viz(FEA)} {feaButtons 1}

# extra text messages
  if {[info exists x3dMsg]} {
    if {[llength $x3dMsg] > 0} {
      puts $x3dFile "\n<!-- Messages -->"
      puts $x3dFile "<ul style=\"padding-left:20px\">"
      foreach item $x3dMsg {puts $x3dFile "<li>$item"}
      puts $x3dFile "</ul>"
      unset x3dMsg
    }
  }

# bounding box
  puts $x3dFile "\n<p><hr>"
  if {$viz(PART) && [info exists x3dBbox]} {
    if {$x3dBbox != ""} {puts $x3dFile "\n<p><input type='checkbox' onclick='togBbox(this.value)'/>$x3dBbox"}
    if {$viz(FEA)} {puts $x3dFile "<p>"}
  }

# axes checkbox
  set check "checked"
  if {$viz(SUPPGEOM)} {set check ""}
  puts $x3dFile "\n<!-- Axes checkbox -->\n<p><input type='checkbox' $check onclick='togAxes(this.value)'/>Origin<p>"

# background color radio buttons
  puts $x3dFile "\n<!-- Background radio button -->\nBackground Color<br>"
  if {!$bgcss} {
    puts $x3dFile "<input type='radio' name='bgcolor' value='1 1 1' $bgcheck1 onclick='BGcolor(this.value)'/>White<br>"
    puts $x3dFile "<input type='radio' name='bgcolor' value='$skyBlue' $bgcheck2 onclick='BGcolor(this.value)'/>Blue<br>"
    puts $x3dFile "<input type='radio' name='bgcolor' value='.8 .8 .8' $bgcheck3 onclick='BGcolor(this.value)'/>Gray<br>"
    puts $x3dFile "<input type='radio' name='bgcolor' value='0 0 0' onclick='BGcolor(this.value)'/>Black"
  } else {
    puts $x3dFile "<input type='radio' name='bgcolor' value='white' $bgcheck1 onclick='BGcolor(this.value)'/>White<br>"
    puts $x3dFile "<input type='radio' name='bgcolor' value='blue' $bgcheck2 onclick='BGcolor(this.value)'/>Blue<br>"
    puts $x3dFile "<input type='radio' name='bgcolor' value='gray' $bgcheck3 onclick='BGcolor(this.value)'/>Gray<br>"
    puts $x3dFile "<input type='radio' name='bgcolor' value='black' onclick='BGcolor(this.value)'/>Black"
  }

# transparency slider
  set transFunc 0
  if {!$ap242XML} {
    set max 0
    if {$viz(PART) || $viz(TESSPART) || \
       ($viz(FEA) && ([info exists entCount(surface_3d_element_representation)] || [info exists entCount(volume_3d_element_representation)]))} {
      set max 1
    }
    if {$max == 1} {
      puts $x3dFile "\n<!-- Transparency slider -->\n<p>Transparency<br>(approximate)<br>"
      puts $x3dFile "<input style='width:80px' type='range' min='0' max='$max' step='0.1' value='0' onchange='matTrans(this.value)'/>"
      set transFunc 1
    }
  }

# mouse message
  puts $x3dFile "\n<p>PageDown for Viewpoints.  Key 'r' to restore, 'a' to view all.  <a href=\"https://www.x3dom.org/documentation/interaction/\">Use the mouse</a> in 'Examine Mode' to rotate, pan, zoom."
  puts $x3dFile "</td></tr></table>"

# -------------------------------------------------------------------------------
# function for PRT, sketch, EDG, part names
  if {$viz(PART) && !$ap242XML} {
    x3dSwitchScript PRT

    if {[info exists nsketch]} {
      if {$nsketch > -1} {puts $x3dFile "\n<!-- SKH switch -->\n<script>function togSKH\(choice\)\{"}
      for {set i 0} {$i <= $nsketch} {incr i} {
        puts $x3dFile " if (!document.getElementById('swSketch$i').checked) \{document.getElementById('swSketch$i').setAttribute('whichChoice', -1);\} else \{document.getElementById('swSketch$i').setAttribute('whichChoice', 0);\}"
        puts $x3dFile " document.getElementById('swSketch$i').checked = !document.getElementById('swSketch$i').checked;"
      }
      if {$nsketch > -1} {puts $x3dFile "\}</script>"}
    }

    if {$opt(partEdges)} {
      puts $x3dFile "\n<!-- EDG switch -->\n<script>function togEDG\(choice\)\{"
      puts $x3dFile " if \(!document.getElementById\('swEDG'\).checked\) \{document.getElementById\('mat1'\).setAttribute\('transparency', 1\);\} else \{document.getElementById\('mat1'\).setAttribute\('transparency', 0\);\}\n\}</script>"
    }
  }

# part names, bounding box
  if {$viz(PART)} {
    if {[info exists x3dParts]} {
      if {[llength [array names x3dParts]] > 1} {
        foreach item [array names x3dParts] {x3dSwitchScript Part$x3dParts($item)}
      }
      catch {unset x3dParts}
      if {[llength [array names parts]] > 2} {
        puts $x3dFile "\n<!-- All Parts Show/Hide switch -->\n<script>function togPartAll(choice)\{"
        foreach name [lsort -nocase [array names parts]] {
          puts $x3dFile " togPart[lindex $parts($name) 0](choice);"
        }
        puts $x3dFile "\}</script>"
      }
      catch {unset parts}
    }

# bounding box
    if {[info exists x3dBbox]} {if {$x3dBbox != ""} {x3dSwitchScript Bbox}}
  }

# switch functions for fem
  if {$viz(FEA)} {
    x3dSwitchScript Nodes
    if {[info exists entCount(surface_3d_element_representation)] || \
        [info exists entCount(volume_3d_element_representation)]}  {x3dSwitchScript Mesh}
    if {[info exists entCount(curve_3d_element_representation)]}   {x3dSwitchScript 1DElements}
    if {[info exists entCount(surface_3d_element_representation)]} {x3dSwitchScript 2DElements}
    if {[info exists entCount(volume_3d_element_representation)]}  {x3dSwitchScript 3DElements}
  }

# function for TPG
  if {$viz(TESSPART)} {
    if {[string first "occurrence" $ao] == -1} {
      x3dSwitchScript TPG
      if {$viz(TESSEDGE)} {x3dSwitchScript TED}

      if {[info exists x3dTessParts]} {
        if {[llength [array names x3dTessParts]] > 1} {
          foreach item [array names x3dTessParts] {x3dSwitchScript TessPart$x3dTessParts($item)}
        }
        catch {unset x3dTessParts}
        if {[llength [array names partstg]] > 2} {
          puts $x3dFile "\n<!-- All Tessellated Parts Show/Hide switch -->\n<script>function togTessPartAll(choice)\{"
          foreach name [lsort -nocase [array names partstg]] {
            puts $x3dFile " togTessPart[lindex $partstg($name) 0](choice);"
          }
          puts $x3dFile "\}</script>"
        }
      }
    }
  }

# function for SMG
  if {$viz(SUPPGEOM)} {x3dSwitchScript SMG}

# function for DTR
  if {$viz(DTMTAR)} {x3dSwitchScript DTR}

# function for DTR
  if {$viz(POINTS)} {x3dSwitchScript Points}

# function for holes
  if {$viz(HOLE)} {x3dSwitchScript Hole}

# onload
  set onload {}
  if {[info exists x3dPartClick]} {if {$x3dPartClick} {lappend onload " document.getElementById('clickedObject').innerHTML = 'click on a part';"}}

# functions for viewpoint names and PMI
  if {[llength $savedViewButtons] > 0 || [info exists savedViewVP]} {
    puts $x3dFile " "
    if {$viz(PMI)} {foreach svn $savedViewButtons {x3dSwitchScript View[lsearch $savedViewNames $svn] $svMap($svn)}}

    if {[info exists savedViewVP]} {
      set id 0
      foreach svn [list "Front 1 (SFA)" "Orthographic (SFA)"] {
        lappend onload "\n var view$id = document.getElementById('$svn');\n view$id.addEventListener('outputchange', function(event) \{"
        lappend onload "  document.getElementById('clickedView').innerHTML = '$svn';"
        incr id
        if {$viz(PMI) && $opt(viewPMIVP)} {
          foreach svn1 $savedViewButtons {
            if {[info exists viewsWithPMI($svn1)]} {
              lappend onload "  document.getElementById('swView$viewsWithPMI($svn1)').setAttribute('whichChoice', 0);"
              lappend onload "  document.getElementById('swView$viewsWithPMI($svn1)').checked = false;"
              lappend onload "  document.getElementById('cbView$viewsWithPMI($svn1)').checked = true;"
            }
          }
        }
        lappend onload " \}, false);"
      }

      set svb $savedViewButtons
      if {!$viz(PMI)} {set svb [array names savedViewpoint]}

      foreach svn $svb {
        lappend onload "\n var view$id = document.getElementById('$svn');\n view$id.addEventListener('outputchange', function(event) \{"
        lappend onload "  document.getElementById('clickedView').innerHTML = '$svn';"
        incr id
        if {$viz(PMI) && $opt(viewPMIVP)} {
          foreach svn1 $savedViewButtons {
            if {[info exists viewsWithPMI($svn1)]} {
              set wc -1
              set ch1 "true"
              set ch2 "false"
              if {$svn == $svn1} {set wc 0; set ch1 "false"; set ch2 "true"}
              if {$svn == $svMap($svn) || $svn == $svn1} {
                lappend onload "  document.getElementById('swView$viewsWithPMI($svn1)').setAttribute('whichChoice', $wc);"
                lappend onload "  document.getElementById('swView$viewsWithPMI($svn1)').checked = $ch1;"
              }
              lappend onload "  document.getElementById('cbView$viewsWithPMI($svn1)').checked = $ch2;"
            }
          }
        }
        lappend onload " \}, false);"
      }
    }
  }

# functions for eventListener for viewpoint if no saved views
  if {![info exists savedViewVP]} {
    set id 0
    foreach svn [list "Front 1" "Side 1" "Top 1" "Front 2" "Side 2" "Top 2" "Isometric" "Orthographic"] {
      lappend onload " var view$id = document.getElementById('$svn');\n view$id.addEventListener('outputchange', function(event) \{document.getElementById('clickedView').innerHTML = '$svn';\}, false);"
      incr id
    }
  }
  catch {unset savedViewVP}

# functions for FEA buttons
  if {$viz(FEA)} {feaButtons 2}

# background function
  if {!$bgcss} {
    puts $x3dFile "\n<!-- Background function -->\n<script>function BGcolor(color){document.getElementById('BG').setAttribute('skyColor', color);}</script>"
  } else {
    puts $x3dFile "\n<!-- Background functions -->
<script>function BGcolor(color){
 if (color == 'blue') {
  document.getElementById('x3d').style.backgroundImage = 'linear-gradient(skyBlue, white)';
 } else if (color == 'gray') {
  document.getElementById('x3d').style.backgroundImage = 'linear-gradient(darkgray, lightgray)';
 } else {
  document.getElementById('x3d').style.background = color;
 }
}
</script>"

# background onload select checked background
    if {[llength $onload] > 0} {lappend onload " "}
    lappend onload " var items = document.getElementsByName('bgcolor');"
    lappend onload " for (var i=0; i<items.length; i++) {if (items\[i\].checked == true) {BGcolor(items\[i\].value);}}"
  }

# axes function
  x3dSwitchScript Axes

# transparency function
  set numTessColor 0
  if {$viz(TESSPART)} {set numTessColor [tessCountColors]}
  if {$transFunc} {
    puts $x3dFile "\n<!-- Transparency function -->\n<script>function matTrans(trans){"

# part transparency
    if {$viz(PART)} {
      if {[info exists x3dApps]} {
        foreach n [lrmdups [lsort -integer $x3dApps]] {
          if {!$opt(partEdges) || $n != 1} {
            if {![info exists matTrans($n)]} {
              puts $x3dFile " document.getElementById('mat$n').setAttribute('transparency', trans);"
            } elseif {$matTrans($n) < 1.} {
              puts $x3dFile " if (trans > $matTrans($n)) {document.getElementById('mat$n').setAttribute('transparency', trans);} else {document.getElementById('mat$n').setAttribute('transparency', $matTrans($n));}"
            }
          }
        }
      }
    }

# tessellated geometry transparency
    for {set i 1} {$i <= $numTessColor} {incr i} {puts $x3dFile " document.getElementById('matTess$i').setAttribute('transparency', trans);"}

# finite element model transparency
    if {$viz(FEA)} {
      if {[info exists entCount(surface_3d_element_representation)]} {
        puts $x3dFile " document.getElementById('mat2Dfem').setAttribute('transparency', trans);"
      }
      if {[info exists entCount(volume_3d_element_representation)]}  {
        puts $x3dFile " document.getElementById('mat3Dfem').setAttribute('transparency', trans);"
        puts $x3dFile " if (trans > 0) {document.getElementById('faces').setAttribute('solid', true);} else {document.getElementById('faces').setAttribute('solid', false);}"
      }
    }
    puts $x3dFile "}</script>"
  }

# onload functions
  if {[llength $onload] > 0} {
    puts $x3dFile "\n<!-- onload functions -->\n<script>document.onload = function() \{\n document.getElementById('clickedView').innerHTML = 'Front 1$sfastr';"
    foreach line $onload {puts $x3dFile $line}
    puts $x3dFile "\}\n</script>"
  }

  puts $x3dFile "</font></body></html>"
  close $x3dFile
  update idletasks

# unset variables
  foreach var [list x3dCoord x3dFile x3dFiles x3dIndex x3dMax x3dMin x3dShape x3dStartFile] {catch {unset -- $var}}
}

# -------------------------------------------------------------------------------
# saved view viewpoints
proc x3dSavedViewpoint {name} {
  global maxxyz opt recPracNames savedViewpoint savedViewVP spaces x3dFiles

# check for errors
  set msg ""
  set pp [lindex $savedViewpoint($name) 3]
  set pbaxis [lindex $savedViewpoint($name) 8]

  if {$pp != "0. 0. 0." || $pbaxis != "0.0 0.0 1.0"} {
    if {$pp != "0. 0. 0."} {append msg " The projection_point should be '0 0 0'."}
    if {$pbaxis != "0.0 0.0 1.0"} {append msg " The planar_box a2p3d axis should be '0 0 1'."}
  }
  set diff [expr {abs([lindex [lindex $savedViewpoint($name) 6] 2]-[lindex $savedViewpoint($name) 2])}]
  if {$diff > 1.} {append msg " The view_plane_distance and the planar_box a2p3d origin Z value should be equal."}
  if {$msg != ""} {
    append msg "$spaces\($recPracNames(pmi242), Sec. 9.4.2.6)"
    errorMsg "Syntax Error: Camera model viewpoint is not modeled correctly.$msg"
  }

# default viewpoint with transform
  set n 0
  foreach xf $x3dFiles {
    incr n
    if {$n == 1} {lappend savedViewVP "<Transform translation='[lindex $savedViewpoint($name) 0]' rotation='[lindex $savedViewpoint($name) 1]'><Viewpoint id='$name' position='0 0 0' orientation='0 1 0 3.14156'/></Transform>"}

# show camera model for debugging
    if {$opt(DEBUGVP)} {
      set scale [trimNum [expr {$maxxyz*0.08}]]
      puts $xf "<Transform translation='[lindex $savedViewpoint($name) 0]' rotation='[lindex $savedViewpoint($name) 1]' scale='$scale $scale $scale'>"
      puts $xf " <Shape><Appearance><Material emissiveColor='1 0 0'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 1. 0. 0.'/></IndexedLineSet></Shape>"
      puts $xf " <Shape><Appearance><Material emissiveColor='0 1 0'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. 1. 0.'/></IndexedLineSet></Shape>"

# line from to pp to vpd (pp2vpd)
      set pp2vpd "[lindex $savedViewpoint($name) 3] [lindex [lindex $savedViewpoint($name) 3] 0] [lindex [lindex $savedViewpoint($name) 3] 1] [trimNum [expr {[lindex [lindex $savedViewpoint($name) 3] 2]+[lindex $savedViewpoint($name) 2]}]]"
      puts $xf " <Shape><Appearance><Material emissiveColor='0 0 1'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='$pp2vpd'/></IndexedLineSet></Shape>"

# planar box a2p3d
      puts $xf " <Transform translation='[lindex $savedViewpoint($name) 6]' rotation='[lindex $savedViewpoint($name) 7]'>"
      puts $xf "  <Shape><Appearance><Material emissiveColor='0 0 0'/></Appearance><IndexedLineSet coordIndex='0 1 2 3 0 -1 0 2 -1 1 3 -1'><Coordinate point='0. 0. 0. [lindex $savedViewpoint($name) 4] 0. 0. [lindex $savedViewpoint($name) 4] [lindex $savedViewpoint($name) 5] 0. 0. [lindex $savedViewpoint($name) 5] 0.'/></IndexedLineSet></Shape>"
      puts $xf " </Transform>"
      puts $xf " <Shape><Appearance><Material emissiveColor='0 0 0'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='[lindex $savedViewpoint($name) 3] [lindex $savedViewpoint($name) 6]'/></IndexedLineSet></Shape>"
      puts $xf " <Transform scale='0.5 0.5 0.5'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$name'><FontStyle family='SANS' justify='BEGIN'/></Text><Appearance><Material diffuseColor='0 0 0'/></Appearance></Shape></Billboard></Transform>"

# label
      set trans [vectrim [vecadd [lindex $savedViewpoint($name) 6] [list [expr {[lindex $savedViewpoint($name) 4]*0.5}] [lindex $savedViewpoint($name) 5] 0.]]]
      set scale [trimNum [expr {[lindex $savedViewpoint($name) 2]/36.}]]
      puts $xf " <Transform translation='$trans' scale='$scale $scale $scale'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$name'><FontStyle family='SANS' justify='MIDDLE'/></Text><Appearance><Material diffuseColor='0 0 0'/></Appearance></Shape></Billboard></Transform>"
      puts $xf "</Transform>\n"
    }
  }
}

# -------------------------------------------------------------------------------
# part checkboxes
proc x3dPartCheckbox {type} {
  global parts partstg x3dFile x3dHeight x3dParts x3dTessParts x3dWidth

  switch -- $type {
    Part {
      set name "Assembly/Part"
      set tog "togPart"
      catch {unset parts}
      foreach idx [array names x3dParts] {set parts($idx) $x3dParts($idx)}
      set arparts [array names parts]
    }
    Tess {
      set name "Tessellated Parts"
      set tog "togTessPart"
      catch {unset partstg}
      foreach idx [array names x3dTessParts] {set partstg($idx) $x3dTessParts($idx)}
      set arparts [array names partstg]
    }
  }

  set txt ""
  set nparts [llength $arparts]
  if {$nparts > 2} {set txt "&nbsp;&nbsp;<button onclick='$tog\All\(this.value)'>Show/Hide</button>"}
  puts $x3dFile "\n<!-- $name checkboxes -->\n<p>$name$txt\n<br><font size='-1'>"

  set lenname 0
  foreach name $arparts {if {[string length $name] > $lenname} {set lenname [string length $name]}}
  set div ""
  set max 40
  if {$nparts > $max || $lenname > $max} {
    append div "<style>div.$type \{overflow: scroll;"
    if {$lenname > $max} {append div " width: [expr {int($x3dWidth*.2)}]px;"}
    if {$nparts > $max} {append div " height: [expr {int($x3dHeight*.75)}]px;"}
    append div "\}</style>"
  }
  if {$div != ""} {puts $x3dFile "$div\n<div class='$type'>"}
  foreach name [lsort -nocase $arparts] {
    switch -- $type {
      Part {set pname [lindex $parts($name) 0]}
      Tess {set pname [lindex $partstg($name) 0]}
    }
    puts $x3dFile "<nobr><input id='cb[string range $tog 3 end]$pname' type='checkbox' checked onclick='$tog$pname\(this.value)'/>$name </nobr><br>"
  }
  if {$div != ""} {puts $x3dFile "</div>"}
  puts $x3dFile "</font>"
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
                    if {$cx != -1} {set id [x3dUnicode $id]}
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
                      if {$cx != -1} {set id [x3dUnicode $id]}
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
# process X and X2 control directives with Unicode characters
proc x3dUnicode {id {type "view"}} {
  set x "&#x"
  set u "\\u"
  set z "00"

  foreach xl [list X X2] {
    set cx [string first "\\$xl\\" $id]
    if {$cx != -1} {
      switch -- $xl {
        X {
          while {$cx != -1} {
            set xu ""
            set uc [string range $id $cx+3 $cx+4]
            switch -- $type {
              view {append xu "$x$uc;"}
              attr {append xu [join [eval list $u$z$uc]]}
            }
            set id [string range $id 0 $cx-1]$xu[string range $id $cx+5 end]
            set cx [string first "\\$xl\\" $id]
          }
        }
        X2 {
          while {$cx != -1} {
            set xu ""
            for {set i 4} {$i < 200} {incr i 4} {
              set uc [string range $id $cx+$i [expr {$cx+$i+3}]]
              if {$type == "attr" && $uc == "000A"} {
                set xu [format "%c" 10]
              } elseif {[string first "\\" $uc] == -1} {
                switch -- $type {
                  view {append xu "$x$uc;"}
                  attr {append xu [join [eval list $u$uc]]}
                }
              } else {
                set cx0 [string first "\\X0\\" $id]
                if {$cx0 != -1} {
                  set id "[string range $id 0 $cx-1]$xu[string range $id $cx0+4 end]"
                  break
                } else {
                  errorMsg " Missing \\X0\\ for \\X2\\"
                  return $id
                }
              }
            }
            set cx [string first "\\$xl\\" $id]
          }
        }
      }
    }
  }
  return $id
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

  if {[info exists draftModelCameras] && $ao == "tessellated_annotation_occurrence"} {set savedViewName [getSavedViewName $objEntity1]}

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

# group annotations
    catch {unset idshape}
    set txt [[[$objEntity1 Attributes] Item [expr 1]] Value]
    regsub -all "'" $txt "\"" idshape

# used for augmented reality workflow
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
# datum targets
proc x3dDatumTarget {maxxyz} {
  global datumTargetView dttype recPracNames spaces viz x3dFile x3dMsg

  outputMsg " Processing datum targets" green
  puts $x3dFile "\n<!-- DATUM TARGETS -->\n<Switch whichChoice='0' id='swDTR'><Group>"

  foreach idx [array names datumTargetView] {
    set dttype [lindex $datumTargetView($idx) 0]
    set shape  [lindex $datumTargetView($idx) 1]
    set color "1 0 0"
    set feat ""
    if {[string first "feature" $idx] != -1} {
      set color "0 .5 0"
      set feat " feature"
    }
    set endTransform 0

# check for handle
    if {[string first "handle" $shape] == -1} {
      set e3 ""

# position and orientation
      set origin [lindex [lindex $datumTargetView($idx) 1] 0]
      set axis   [lindex [lindex $datumTargetView($idx) 1] 1]
      set refdir [lindex [lindex $datumTargetView($idx) 1] 2]
      if {$origin == "0. 0. 0."} {
        set msg "Datum target(s) located at the origin."
        if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
      }
      set shape $dttype

# handle, then shape is with geometric entity (cartesian_point, line, and circle are supported)
    } else {
      set e3 [lindex $datumTargetView($idx) 1]
      set shape [$e3 Type]
      if {$shape == "trimmed_curve"} {
        set e3 [[[$e3 Attributes] Item [expr 2]] Value]
        if {[$e3 Type] == "line"}   {set shape [$e3 Type]}
        if {[$e3 Type] == "circle"} {set shape "circular curve"}
      } elseif {$shape == "circle"} {
        set shape "circular curve"
      }
    }

# text
    set textOrigin "0 0 0"
    set target [lindex $datumTargetView($idx) end]
    set len [string length $target]
    if {$len < 2 || $len > 5 || ![string is alpha [string index $target 0]]} {set target ""}
    set textJustify "BEGIN"
    if {$e3 != ""} {set textJustify "END"}
    if {$target != ""} {puts $x3dFile "<!-- $target -->"}

# process different shapes
    if {[catch {
      switch -- $shape {
        point -
        vertex_point -
        cartesian_point {
# generate point
          set rad [trimNum [expr {$maxxyz*0.00125}]]
          if {$e3 != ""} {
            if {$shape == "vertex_point"} {
              set e3 [[[$e3 Attributes] Item [expr 2]] Value]
              if {[$e3 Type] != "cartesian_point"} {errorMsg "Datum target vertex_point defined by '[$e3 Type]' is not supported."}
            }
            set origin [vectrim [[[$e3 Attributes] Item [expr 2]] Value]]
          }
          puts $x3dFile "<Transform translation='$origin'><Shape><Appearance><Material diffuseColor='$color' emissiveColor='$color'/></Appearance><Sphere radius='$rad'></Sphere></Shape>"
          set target " $target"
          set viz(DTMTAR) 1
          set endTransform 1
        }

        line -
        edge_curve {
# generate line
          if {$e3 == ""} {
            puts $x3dFile [x3dTransform $origin $axis $refdir "$shape datum target"]
            set x [trimNum [lindex [lindex $datumTargetView($idx) 2] 1]]
            puts $x3dFile " <Shape><Appearance><Material emissiveColor='$color'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0 0 0 $x 0 0'/></IndexedLineSet></Shape>"
            set textOrigin "[trimNum [expr {$x*0.5}]] 0 0"
            set endTransform 1
          } else {
            if {$shape == "line"} {
              set e4 [[[$e3 Attributes] Item [expr 2]] Value]
              set coord1 [vectrim [[[$e4 Attributes] Item [expr 2]] Value]]
              set e5 [[[$e3 Attributes] Item [expr 3]] Value]
              set mag [[[$e5 Attributes] Item [expr 3]] Value]
              set e6 [[[$e5 Attributes] Item [expr 2]] Value]
              set dir [[[$e6 Attributes] Item [expr 2]] Value]
              set coord2 [vectrim [vecadd $coord1 [vecmult $dir $mag]]]
            } elseif {$shape == "edge_curve"} {
              set vp [[[$e3 Attributes] Item [expr 2]] Value]
              set cp [[[$vp Attributes] Item [expr 2]] Value]
              set coord1 [vectrim [[[$cp Attributes] Item [expr 2]] Value]]
              set vp [[[$e3 Attributes] Item [expr 3]] Value]
              set cp [[[$vp Attributes] Item [expr 2]] Value]
              set coord2 [vectrim [[[$cp Attributes] Item [expr 2]] Value]]
            }
            puts $x3dFile "<Shape><Appearance><Material emissiveColor='$color'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='$coord1 $coord2'/></IndexedLineSet></Shape>"
            set textOrigin [vectrim [vecmult [vecadd $coord1 $coord2] 0.5]]
          }
          set viz(DTMTAR) 1
        }

        rectangle {
# generate rectangle
          puts $x3dFile [x3dTransform $origin $axis $refdir "$shape datum target"]
          foreach i {2 3} {
            set type [lindex $datumTargetView($idx) $i]
            switch -- [lindex $type 0] {
              "target length" {set x [trimNum [expr {[lindex $type 1]*0.5}]]}
              "target width"  {set y [trimNum [expr {[lindex $type 1]*0.5}]]}
            }
          }
          puts $x3dFile " <Shape><Appearance><Material emissiveColor='$color'/></Appearance><IndexedLineSet coordIndex='0 1 2 3 0 -1'><Coordinate point='-$x -$y 0 $x -$y 0 $x $y 0 -$x $y 0'/></IndexedLineSet></Shape>"
          puts $x3dFile " <Shape><Appearance><Material diffuseColor='$color' transparency='0.8'/></Appearance><IndexedFaceSet solid='false' coordIndex='0 1 2 3 -1'><Coordinate point='-$x -$y 0 $x -$y 0 $x $y 0 -$x $y 0'/></IndexedFaceSet></Shape>"
          set endTransform 1
          set viz(DTMTAR) 1
        }

        circle -
        "circular curve" {
# generate circle
          if {$e3 == ""} {
            set rad [trimNum [expr {[lindex [lindex $datumTargetView($idx) 2] 1]*0.5}]]
          } else {
            set e4 [[[$e3 Attributes] Item [expr 2]] Value]
            set rad [[[$e3 Attributes] Item [expr 3]] Value]
            set a2p3d [x3dGetA2P3D $e4]
            set origin [lindex $a2p3d 0]
            set axis   [lindex $a2p3d 1]
            set refdir [lindex $a2p3d 2]
          }
          puts $x3dFile [x3dTransform $origin $axis $refdir "$shape datum target"]
          set ns 48
          set angle 0.
          set dlt [expr {6.28319/$ns}]
          set index ""
          for {set i 0} {$i < $ns} {incr i} {append index "$i "}
          set coord ""
          for {set i 0} {$i < $ns} {incr i} {
            append coord "[trimNum [expr {$rad*cos($angle)}]] "
            append coord "[trimNum [expr {$rad*sin($angle)}]] "
            append coord "0 "
            set angle [expr {$angle+$dlt}]
          }
          puts $x3dFile " <Shape><Appearance><Material emissiveColor='$color'/></Appearance><IndexedLineSet coordIndex='$index 0 -1'><Coordinate point='$coord'/></IndexedLineSet></Shape>"
          if {$shape == "circle"} {
            puts $x3dFile " <Shape><Appearance><Material diffuseColor='$color' transparency='0.8'/></Appearance><IndexedFaceSet solid='false' coordIndex='$index -1'><Coordinate point='$coord'/></IndexedFaceSet></Shape>"
          } else {
            set textOrigin "$rad 0 0"
          }
          set endTransform 1
          set viz(DTMTAR) 1
        }

        advanced_face {
# for advanced face, look for circles and lines
          set e1 $e3
          set e2 [[[$e1 Attributes] Item [expr 3]] Value]

# if in a plane, follow face_outer_bounds and face_bounds to ...
          if {[$e2 Type] == "plane"} {
            set e2s [[[$e1 Attributes] Item [expr 2]] Value]
            set igeom 0
            set coord ""
            set ncoord 0

# get number of face bounds
            set nbound 0
            ::tcom::foreach e2 $e2s {incr nbound}

            ::tcom::foreach e2 $e2s {
              set e3 [[[$e2 Attributes] Item [expr 2]] Value]
              set e4s [[[$e3 Attributes] Item [expr 2]] Value]

# get number and types of geometric entities defining the edges
              set ngeom 0
              set gtypes {}
              ::tcom::foreach e4 $e4s {
                incr ngeom
                set e5 [[[$e4 Attributes] Item [expr 4]] Value]
                set e6 [[[$e5 Attributes] Item [expr 4]] Value]
                if {[lsearch $gtypes [$e6 Type]] == -1} {lappend gtypes [$e6 Type]}
              }

# check for only multiple circles or ellipses
              set onlyCircle 0
              if {[llength $gtypes] == 1} {if {$gtypes == "circle" || $gtypes == "ellipse"} {set onlyCircle 1}}

              ::tcom::foreach e4 $e4s {
                set e5 [[[$e4 Attributes] Item [expr 4]] Value]
                set e6 [[[$e5 Attributes] Item [expr 4]] Value]
                incr igeom

# advanced face circle and ellipse edges
                if {[$e6 Type] == "circle" || [$e6 Type] == "ellipse"} {
                  if {$nbound == 1 && ($ngeom == 1 || $onlyCircle)} {
                    set rad [[[$e6 Attributes] Item [expr 3]] Value]
                    set scale ""

# check ellipse axes
                    if {[$e6 Type] == "ellipse"} {
                      set rad1 [[[$e6 Attributes] Item [expr 4]] Value]
                      set sy [expr {$rad1/$rad}]
                      set scale "1 $sy 1"
                      set dsy [trimNum [expr {abs($sy-1.)}]]
                      if {$dsy <= 0.05} {errorMsg " Datum target ($dttype) '[$e6 Type]' axes ($rad, $rad1) are almost identical."}
                    }

# transform for circle
                    if {!$onlyCircle || $igeom == 1} {
                      set a2p3d [x3dGetA2P3D [[[$e6 Attributes] Item [expr 2]] Value]]
                      puts $x3dFile [x3dTransform [lindex $a2p3d 0] [lindex $a2p3d 1] [lindex $a2p3d 2] "$shape circle datum target" $scale]
                    }

# generate coordinates
                    incr ncoord 48
                    set angle 0.
                    set dlt [expr {6.28319/$ncoord}]
                    for {set i 0} {$i < $ncoord} {incr i} {
                      append coord "[trimNum [expr {$rad*cos($angle)}]] "
                      append coord "[trimNum [expr {$rad*sin($angle)}]] "
                      append coord "0 "
                      set angle [expr {$angle+$dlt}]
                      if {$i == 0 && $igeom == 1} {set textOrigin $coord}
                    }
                    set endTransform 1
                  } else {
                    errorMsg "[string totitle $dttype] datum target$feat edge defined by multiple types of curves is not supported."
                  }

# advanced face line edges
                } elseif {[$e6 Type] == "line"} {
                  set e7 [[[$e6 Attributes] Item [expr 2]] Value]
                  set pt [vectrim [[[$e7 Attributes] Item [expr 2]] Value]]
                  append coord "$pt "
                  incr ncoord
                  if {$ncoord == 1 && $igeom == 1} {set textOrigin $pt}

# not a circle or line
                } else {
                  set target ""
                  errorMsg "[string totitle $dttype] datum target$feat edge defined by '[$e2 Type]' is not supported."
                }
              }
            }

# shape for circles and lines
            if {$coord != ""} {
              set index ""
              for {set i 0} {$i < $ncoord} {incr i} {append index "$i "}
              puts $x3dFile " <Shape><Appearance><Material emissiveColor='$color'/></Appearance><IndexedLineSet coordIndex='$index 0 -1'><Coordinate point='$coord'/></IndexedLineSet></Shape>"
              puts $x3dFile " <Shape><Appearance><Material diffuseColor='$color' transparency='0.8'/></Appearance><IndexedFaceSet solid='false' coordIndex='$index -1'><Coordinate point='$coord'/></IndexedFaceSet></Shape>"
              set viz(DTMTAR) 1
            }

# non planes are not supported
          } else {
            set target ""
            errorMsg "[string totitle $dttype] datum target$feat face defined by '[$e2 Type]' is not supported."
          }
        }

        default {
          set target ""
          errorMsg "Syntax Error: [string totitle $dttype] datum target$feat defined by '$shape' should use an 'advanced_face'.$spaces\($recPracNames(pmi242), Sec. 6.6.2, Fig. 44)"
        }
      }

# small coordinate triad
      if {$shape != "point" && $shape != "cartesian_point" && $shape != "advanced_face" && [string first "feature" $idx] == -1} {
        set size [trimNum [expr {$maxxyz*0.005}]]
        puts $x3dFile " <Shape><Appearance><Material emissiveColor='1 0 0'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. $size 0. 0.'/></IndexedLineSet></Shape>"
        puts $x3dFile " <Shape><Appearance><Material emissiveColor='0 .5 0'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. $size 0.'/></IndexedLineSet></Shape>"
        puts $x3dFile " <Shape><Appearance><Material emissiveColor='0 0 1'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. 0. $size'/></IndexedLineSet></Shape>"
      }

# datum target label
      if {$target != ""} {
        set size [trimNum [expr {$maxxyz*0.01}]]
        set trans ""
        if {$textOrigin != "0 0 0"} {set trans " translation='$textOrigin'"}
        puts $x3dFile " <Transform$trans scale='$size $size $size'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$target'><FontStyle family='SANS' justify='$textJustify'/></Text><Appearance><Material diffuseColor='$color'/></Appearance></Shape></Billboard></Transform>"
      }

# end transform
      if {$endTransform} {puts $x3dFile "</Transform>"}

    } emsg]} {
      errorMsg "Error viewing a '$dttype' datum target$feat ($target): $emsg"
    }
  }
  puts $x3dFile "</Group></Switch>"
  catch {unset datumTargetView}
}

# -------------------------------------------------------------------------------
# supplemental geometry
proc x3dSuppGeom {maxxyz} {
  global cgrObjects axesDef planeDef recPracNames skipEntities syntaxErr tessSuppGeomFile tessSuppGeomFileName trimVal x3dFile
  global objDesign
  if {![info exists objDesign]} {return}

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
  set origin [lindex $a2p3d 0]
  set axis   [lindex $a2p3d 1]
  set refdir [lindex $a2p3d 2]
  set transform [x3dTransform $origin $axis $refdir "supplemental geometry 'axes'"]

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

  set nsize [trimNum [expr {$tsize*0.5}]]
  set tcolor "1 0 0"
  set name [[[$e2 Attributes] Item [expr 1]] Value]
  if {$axisColor != ""} {set tcolor $axisColor}
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
    errorMsg "Error adding 'point' supplemental geometry: $emsg"
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
    errorMsg "Error adding '$type' supplemental geometry: $emsg"
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
      if {$dsy <= 0.05} {errorMsg " Supplemental geometry $type axes ($rad,$rad1) are almost identical."}
    }

# circle position and orientation
    set a2p3d [x3dGetA2P3D $e4]
    set origin [lindex $a2p3d 0]
    set axis   [lindex $a2p3d 1]
    set refdir [lindex $a2p3d 2]
    set transform [x3dTransform $origin $axis $refdir "supplemental geometry '$type'" $scale]
    puts $x3dFile $transform

# generate circle, account for trimming
# lim is the limit on an angle before deciding it is in degrees to convert to radians
    set ns 48
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
    errorMsg "Error adding '$type' supplemental geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry for plane
proc x3dSuppGeomPlane {e2 size} {
  global planeDef spmiTypesPerFile viz x3dFile

# check for plane color and transparency
  set tmp [x3dSuppGeomColor $e2 "plane"]
  set planeColor [lindex $tmp 0]
  if {$planeColor == ""} {set planeColor "0 0 1"}
  set planeTrans [lindex $tmp 1]
  if {$planeTrans == ""} {set planeTrans 0.8}

  if {[catch {
    set e3 [[[$e2 Attributes] Item [expr 2]] Value]

# plane position and orientation
    set a2p3d [x3dGetA2P3D $e3]
    set origin [lindex $a2p3d 0]
    set axis   [lindex $a2p3d 1]
    set refdir [lindex $a2p3d 2]
    set transform [x3dTransform $origin $axis $refdir "supplemental geometry 'plane'"]

# plane geometry
    set nsize [trimNum [expr {$size*2.}]]
    set id [lsearch $planeDef "$nsize $planeColor"]
    if {$id != -1} {
      puts $x3dFile "$transform<Group USE='plane$id'></Group>"
      lappend spmiTypesPerFile "supplemental geometry"
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
              errorMsg " Bounding edges defined by '[formatComplexEnt [$e7 Type]]' for supplemental geometry 'plane' are not supported."
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
          errorMsg " Using bounding edges for supplemental geometry 'plane'" red
          lappend spmiTypesPerFile "bounded supplemental geometry"
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
        lappend spmiTypesPerFile "supplemental geometry"
      }

# set index
      set idx ""
      for {set i 0} {$i < $ncorners} {incr i} {append idx "$i "}
      set idx [string trim $idx]

      puts $x3dFile $transform
      set def ""
      if {[llength $planeDef] > 0} {set def " DEF='plane[expr {[llength $planeDef]-1}]'"}
      puts $x3dFile " <Group$def><Shape><IndexedLineSet coordIndex='$idx 0 -1'><Coordinate point='$corners'/></IndexedLineSet><Appearance><Material emissiveColor='$planeColor'/></Appearance></Shape>"
      puts $x3dFile " <Shape><IndexedFaceSet solid='false' coordIndex='$idx -1'><Coordinate point='$corners'/></IndexedFaceSet><Appearance><Material diffuseColor='$planeColor' transparency='$planeTrans'/></Appearance></Shape></Group>"
    }

# plane name at one corner
    set name [[[$e2 Attributes] Item [expr 1]] Value]
    if {$name != ""} {
      set tsize [trimNum [expr {$size*0.33}]]
      if {![info exists txtpos]} {set txtpos "-$nsize -$nsize 0."}
      puts $x3dFile " <Transform translation='$txtpos' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Text string='$name'><FontStyle family='SANS' justify='BEGIN'/></Text><Appearance><Material diffuseColor='$planeColor'/></Appearance></Shape></Billboard></Transform>"
    }
    puts $x3dFile "</Transform>"
    set viz(SUPPGEOM) 1

  } emsg]} {
    errorMsg "Error adding 'plane' supplemental geometry: $emsg"
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
    set axis   [lindex $a2p3d 1]
    set refdir [lindex $a2p3d 2]
    set transform [x3dTransform $origin $axis $refdir "supplemental geometry 'cylinder'"]
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
    puts $x3dFile "  <Shape><Cylinder radius='$rad' height='[trimNum $height]' top='false' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='$cylColor' transparency='$cylTrans'/></Appearance></Shape>"
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
    errorMsg "Error adding 'cylinder' supplemental geometry: $emsg"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry color and transparency
proc x3dSuppGeomColor {e0 type} {
  global grayBackground recPracNames

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

            if {$type == "plane" || $type == "cylinder"} {
              errorMsg "Syntax Error: Wrong type of style ([$e3 Type]) for a supplemental geometry '$type'.  ($recPracNames(model), Sec. 4.2.2)"
            }

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

            if {$type != "plane" && $type != "cylinder"} {
              errorMsg "Syntax Error: Wrong type of style ([$e3 Type]) for a supplemental geometry '$type'.  ($recPracNames(model), Sec. 4.2.2)"
            }
          } else {
            errorMsg "Syntax Error: Color defined with '[$e3 Type]' referred from 'presentation_style_assignment' for supplemental geometry '$type' is not allowed.  ($recPracNames(model), Sec. 4.2.2)"
          }
        }
      }
    }
  } emsg]} {
    errorMsg " Error getting color for '$type' supplemental geometry: $emsg"
  }

  foreach color [list "1. 1. 1." "1 1 1" "1. 1. 0." "1 1 0"] {if {$sgColor == $color} {set grayBackground 1}}
  return [list $sgColor $sgTrans]
}

# -------------------------------------------------------------------------------
# holes counter and spotface
proc x3dHoles {maxxyz} {
  global dim DTR entCount gen holeDefinitions opt recPracNames spaces syntaxErr viz x3dFile
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
      if {[catch {
        set e2type [$e2 Type]
      } emsg1]} {
        ::tcom::foreach e2a $e2 {set e2 $e2a; break}
      }
      if {[string first "occurrence" [$e1 Type]] != -1 && [$e2 Type] == "mapped_item"} {
        set defID   [[[[$e1 Attributes] Item [expr 5]] Value] P21ID]
        set defType [[[[$e1 Attributes] Item [expr 5]] Value] Type]

# hole name
        set holeName [split $defType "_"]
        foreach idx {0 1} {
          if {[string first "counter" [lindex $holeName $idx]] != -1 || [string first "spotface" [lindex $holeName $idx]] != -1} {set holeName [lindex $holeName $idx]}
        }
        if {$defType == "basic_round_hole"} {set holeName $defType}

# check if there is an a2p3d associated with a hole occurrence
        set e3 [[[$e2 Attributes] Item [expr 3]] Value]
        if {[$e3 Type] == "axis2_placement_3d"} {
          if {$head} {
            outputMsg " Processing hole geometry" green
            puts $x3dFile "\n<!-- HOLES -->\n<Switch whichChoice='0' id='swHole'><Group>"
            set head 0
            set viz(HOLE) 1
          }
          if {[lsearch $holeDEF $defID] == -1} {puts $x3dFile "<!-- $defType $defID -->"}

# hole geometry
          if {[info exists holeDefinitions($defID)]} {

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
            set lhd [llength $holeDefinitions($defID)]
            if {$lhd > 1} {
              set holeType [lindex [lindex $holeDefinitions($defID) [expr {$lhd-3}]] 0]

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
                if {[info exists drillDep]} {
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
                    puts $x3dFile "  <Shape><Cylinder radius='$drillRad' height='[trimNum [expr {$drillDep-$sinkDep}] 5]' top='$holeTop' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 1'/></Appearance></Shape></Transform>"
                  }
                  puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {$sinkDep*0.5}] 5]'>"
                  puts $x3dFile "  <Shape><Cone bottomRadius='$sinkRad' topRadius='$drillRad' height='[trimNum $sinkDep 5]' top='false' bottom='false' solid='false'></Cone><Appearance><Material diffuseColor='0 1 1'/></Appearance></Shape></Transform>"
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
                if {[info exists drillDep]} {
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
                    puts $x3dFile "  <Shape><Cylinder radius='$drillRad' height='[trimNum [expr {$drillDep-$boreDep}] 5]' top='$holeTop' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                  }
                  puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum $boreDep 5]'>"
                  puts $x3dFile "  <Shape><Cone bottomRadius='$boreRad' topRadius='$drillRad' height='0.001' top='false' bottom='false' solid='false'></Cone><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                  puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {$boreDep*0.5}] 5]'>"
                  puts $x3dFile "  <Shape><Cylinder radius='$boreRad' height='[trimNum $boreDep 5]' top='false' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                  puts $x3dFile "</Group></Transform>"
                  lappend holeDEF $defID
                } else {
                  puts $x3dFile "$transform<Group USE='$holeName$defID'></Group></Transform>"
                }

# basic round hole
              } elseif {$holeType == "round_hole"} {
                set hole [lindex $holeDefinitions($defID) 0]
                set holeRad [expr {[lindex $hole 1]*0.5*$scale}]
                if {[lindex $hole 2] != ""} {
                  set holeDep [expr {[lindex $hole 2]*$scale}]
                } else {
                  set holeDep [expr {[lindex $hole 1]*0.01*$scale}]
                }
                if {[lsearch $holeDEF $defID] == -1} {
                  puts $x3dFile "$transform<Group DEF='$holeName$defID'>"
                  if {!$thruHole && [lindex $hole 2] != ""} {
                    puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum $holeDep 5]'>"
                    puts $x3dFile "  <Shape><Cone bottomRadius='$holeRad' topRadius='0' height='0.001' top='false' bottom='false' solid='false'></Cone><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                  }
                  puts $x3dFile " <Transform rotation='1 0 0 1.5708' translation='0 0 [trimNum [expr {$holeDep*0.5}] 5]'>"
                  puts $x3dFile "  <Shape><Cylinder radius='$holeRad' height='[trimNum $holeDep 5]' top='false' bottom='false' solid='false'></Cylinder><Appearance><Material diffuseColor='0 1 0'/></Appearance></Shape></Transform>"
                  puts $x3dFile "</Group></Transform>"
                  lappend holeDEF $defID
                } else {
                  puts $x3dFile "$transform<Group USE='$holeName$defID'></Group></Transform>"
                }
              }
            }
          } elseif {!$opt(PMISEM) || $gen(None)} {
            errorMsg " Only hole drill entry points are shown when the Analyzer report for Semantic PMI is not selected."
            if {[lsearch $holeDEF $defID] == -1} {lappend holeDEF $defID}
          }

# point at origin of hole
          set e4 [[[$e3 Attributes] Item [expr 2]] Value]
          if {![info exists thruHole]} {set thruHole 0}
          x3dSuppGeomPoint $e4 $drillPoint $thruHole $holeName
        }
      }
    } emsg]} {
      errorMsg "Error adding 'hole' geometry: $emsg"
    }
  }
  if {$viz(HOLE)} {puts $x3dFile "</Group></Switch>\n"}
  catch {unset holeDefinitions}

  set ok 0
  if {![info exists entCount(item_identified_representation_usage)]} {set ok 1} elseif {$entCount(item_identified_representation_usage) == 0} {set ok 1}
  if {$ok} {errorMsg "Syntax Error: Missing IIRU to link hole with explicit geometry.$spaces\($recPracNames(holes), Sec. 5.1.1.2)"}
}

# -------------------------------------------------------------------------------
# set predefined color
proc x3dPreDefinedColor {name} {
  global defaultColor recPracNames spaces

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
      errorMsg "Syntax Error: draughting_pre_defined_colour name '$name' is not supported (using [lindex $defaultColor 1])$spaces\($recPracNames(model), Sec. 4.2.3, Table 2)"
    }
  }
  return $color
}

# -------------------------------------------------------------------------------
# write geometry for polyline annotations
proc x3dPolylinePMI {{objEntity1 ""}} {
  global ao gpmiPlacement lastID mytemp opt placeAnchor placeOrigin polylineTxt recPracNames savedViewFile savedViewFileName
  global savedViewName savedViewNames spaces x3dColor x3dColorFile x3dCoord x3dFile x3dIndex x3dIndexType x3dShape

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

# group annotations
        if {$objEntity1 != "" && [string first "placeholder" $ao] == -1} {
          set aoID [$objEntity1 P21ID]
          if {![info exists lastID($f)] || $aoID != $lastID($f)} {
            if {[info exists lastID($f)]} {puts $f "</Group>"}
            set polylineTxt "AO $aoID"
            set e0 [[[$objEntity1 Attributes] Item [expr 3]] Value]
            set txt [[[$e0 Attributes] Item [expr 1]] Value]
            if {$txt != ""} {append polylineTxt " | $txt"}
            set txt [[[$objEntity1 Attributes] Item [expr 1]] Value]
            regsub -all "'" $txt "\"" idshape
            if {$txt != ""} {append polylineTxt " | $idshape"}
            puts $f "<Group id='$polylineTxt'>"
          }
          set lastID($f) $aoID
        }

# multiple saved view color
        if {$opt(gpmiColor) == 3 && [llength $savedViewNames] > 1} {
          if {![info exists x3dColorFile($f)]} {set x3dColorFile($f) [x3dSetPMIColor $opt(gpmiColor) 1]}
          set x3dColor $x3dColorFile($f)
        }

        if {[string length $x3dCoord] > 0} {

# placeholder transform
          if {[string first "placeholder" $ao] != -1} {
            if {![info exists polylineTxt]} {set polylineTxt ""}
            set transform [x3dTransform $gpmiPlacement(origin) $gpmiPlacement(axis) $gpmiPlacement(refdir) "annotation placeholder" "" $polylineTxt]
            puts $f $transform
          }

# start shape
          set idstr ""
          if {[info exists idshape]} {if {$idshape != ""} {set idstr " id='$idshape'"}}
          if {$x3dColor != ""} {
            puts $f "<Shape$idstr><Appearance><Material emissiveColor='$x3dColor'/></Appearance>"
          } else {
            puts $f "<Shape$idstr><Appearance><Material emissiveColor='0 0 0'/></Appearance>"
            errorMsg "Syntax Error: Missing PMI Presentation color for [formatComplexEnt $ao] (using black)$spaces\($recPracNames(pmi242), Sec. 8.5, Fig. 84)"
          }
          catch {unset idshape}

# index and coordinates
          puts $f " <IndexedLineSet coordIndex='[string trim $x3dIndex]'>\n  <Coordinate point='[string trim $x3dCoord]'/></IndexedLineSet></Shape>"

# end placeholder transform, add leader line
          if {[string first "placeholder" $ao] != -1} {
            puts $f "</Transform>"
            puts $f "<Shape><Appearance><Material emissiveColor='$x3dColor'/></Appearance>"
            puts $f " <IndexedLineSet coordIndex='0 1 -1'>\n  <Coordinate point='$placeOrigin $placeAnchor'/></IndexedLineSet></Shape>"
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
    errorMsg "Error writing polyline annotation graphics: $emsg3"
  }
  update idletasks
}

# -------------------------------------------------------------------------------
# write coordinate axes
proc x3dCoordAxes {size} {
  global viz x3dAxes x3dFile

  set choice 0
  catch {if {$viz(SUPPGEOM)} {set choice -1}}

# axes
  if {$x3dAxes} {
    puts $x3dFile "\n<!-- COORDINATE AXIS -->\n<Switch whichChoice='$choice' id='swAxes'><Group>"
    puts $x3dFile "<Shape><Appearance><Material emissiveColor='1 0 0'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0 0 0 $size 0 0'/></IndexedLineSet></Shape>"
    puts $x3dFile "<Shape><Appearance><Material emissiveColor='0 .5 0'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0 0 0 0 $size 0'/></IndexedLineSet></Shape>"
    puts $x3dFile "<Shape><Appearance><Material emissiveColor='0 0 1'/></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0 0 0 0 0 $size'/></IndexedLineSet></Shape>"

# xyz labels
    set tsize [trimNum [expr {$size*0.33}]]
    puts $x3dFile "<Transform translation='$size 0 0' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='1 0 0'/></Appearance><Text string='X'><FontStyle family='SANS'/></Text></Shape></Billboard></Transform>"
    puts $x3dFile "<Transform translation='0 $size 0' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='0 .5 0'/></Appearance><Text string='Y'><FontStyle family='SANS'/></Text></Shape></Billboard></Transform>"
    puts $x3dFile "<Transform translation='0 0 $size' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='0 0 1'/></Appearance><Text string='Z'><FontStyle family='SANS'/></Text></Shape></Billboard></Transform>"

# credits
    set tsize1 [trimNum [expr {$tsize*0.05}] 3]
    puts $x3dFile "<Transform scale='$tsize1 $tsize1 $tsize1'><Billboard axisOfRotation='0 0 0'><Shape><Text string='\"Generated by the\",\"NIST STEP File Analyzer and Viewer [getVersion]\"'><FontStyle family='SANS'/></Text><Appearance><Material diffuseColor='0 0 0'/></Appearance></Shape></Billboard></Transform>"

# marker for selection point
    puts $x3dFile "\n<!-- SELECTION POINT -->"
    puts $x3dFile "<Transform id='marker'><Shape><PointSet><Coordinate point='0 0 0'/></PointSet></Shape></Transform>"
    puts $x3dFile "</Group></Switch>"
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

# if refdir not specified, do not use default 1 0 0 if axis is 1 0 0
  } elseif {$axis == "1.0 0.0 0.0"} {
    set refdir "0 0 1"
  }

  return [list $origin $axis $refdir]
}

# -------------------------------------------------------------------------------
# generate transform
proc x3dTransform {origin axis refdir {text ""} {scale ""} {id ""}} {

  set transform "<Transform"
  if {$id != ""} {append transform " id='$id'"}
  if {$origin != "0. 0. 0."} {append transform " translation='$origin'"}

# get rotation from axis and refdir
  set rot [x3dGetRotation $axis $refdir $text]
  if {[lindex $rot 3] != 0} {append transform " rotation='$rot'"}
  if {$scale != ""} {append transform " scale='$scale'"}
  append transform ">"
  return $transform
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

# -------------------------------------------------------------------------------------------------
# open x3dom file
proc openX3DOM {{fn ""} {numFile 0}} {
  global lastX3DOM multiFile opt scriptName x3dMsgColor x3dFileName viz

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
    } elseif {$opt(viewPMI) || $opt(viewTessPart) || $opt(viewFEA) || $opt(viewPart)} {
      if {$opt(xlFormat) == "None"} {errorMsg "There is nothing in the STEP file for the Viewer to show based on the selections on the Options tab."}
      return
    }
  }

  if {[file exists $fn] != 1} {return}
  if {![info exists multiFile]} {set multiFile 0}

  set open 0
  if {![info exists viz(PART)]} {set viz(PART) 0}
  if {$f3} {
    set open 1
  } elseif {($viz(PMI) || $viz(TESSPART) || $viz(FEA) || $viz(PART)) && $fn != "" && $multiFile == 0} {
    if {$opt(outputOpen)} {set open 1}
  }

# open file (.html) in web browser
  set lastX3DOM $fn
  if {$open} {
    if {![info exists x3dMsgColor]} {set x3dMsgColor blue}
    catch {.tnb select .tnb.status}
    outputMsg "\nOpening Viewer file: [file tail $fn] ([fileSize $fn])" $x3dMsgColor
    openURL [file nativename $fn]
    update idletasks
  } elseif {$numFile == 0 && [string first "STEP-File-Analyzer.exe" $scriptName] != -1} {
    outputMsg " Use F3 to open the Viewer file" red
  }
}

# -------------------------------------------------------------------------------
# get saved view names
proc getSavedViewName {objEntity} {
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
# script for switch node
proc x3dSwitchScript {name {name1 ""}} {
  global savedViewNames viz x3dFile

# not parts
  if {[string first "Part" $name] != 0 && [string first "TessPart" $name] != 0} {

# adjust for saved views
    if {$name1 == ""} {set name1 $name}
    set viewName ""
    if {[string first "View" $name] == 0} {
      set viewName " [lindex $savedViewNames [string range $name end end]]"
      if {$name1 != ""} {set name1 "View[lsearch $savedViewNames $name1]"}
    }

# controls if checking/unchecking box is before or after changing whichChoice
    set ok 1
    if {$name == "Axes"} {catch {if {$viz(SUPPGEOM)} {set ok 0}}}
    if {$name == "Bbox"} {set ok 0}

    puts $x3dFile "\n<!-- $name$viewName switch -->\n<script>function tog$name\(choice)\{"
    if {!$ok} {puts $x3dFile " document.getElementById('sw$name').checked = !document.getElementById('sw$name').checked;"}
    puts $x3dFile " if (!document.getElementById('sw$name').checked) \{document.getElementById('sw$name1').setAttribute('whichChoice', -1);\} else \{document.getElementById('sw$name1').setAttribute('whichChoice', 0);\}"
    if {$ok}  {puts $x3dFile " document.getElementById('sw$name').checked = !document.getElementById('sw$name').checked;"}
    puts $x3dFile "\}</script>"

# parts
  } else {
    set c1 [string first "Part" $name]
    set ids [string range $name $c1+4 end]
    set name1 [string range $name 0 $c1+3]

    puts $x3dFile "\n<!-- $name switch -->\n<script>function tog$name1[lindex $ids 0]\(choice)\{"
    if {[llength $ids] == 1} {
      puts $x3dFile " if (!document.getElementById('sw$name').checked) \{document.getElementById('sw$name').setAttribute('whichChoice', -1);\} else \{document.getElementById('sw$name').setAttribute('whichChoice', 0);\}"
      puts $x3dFile " document.getElementById('sw$name').checked = !document.getElementById('sw$name').checked;"
      puts $x3dFile " document.getElementById('cb$name').checked = !document.getElementById('sw$name').checked;\n\}</script>"
    } else {
      puts $x3dFile " if (!document.getElementById('sw$name1[lindex $ids 0]').checked) \{"
      foreach id $ids {puts $x3dFile "  document.getElementById('sw$name1$id').setAttribute('whichChoice', -1);"}
      puts $x3dFile " \} else \{"
      foreach id $ids {puts $x3dFile "  document.getElementById('sw$name1$id').setAttribute('whichChoice', 0);"}
      puts $x3dFile " \}"
      puts $x3dFile " document.getElementById('sw$name1[lindex $ids 0]').checked = !document.getElementById('sw$name1[lindex $ids 0]').checked;"
      puts $x3dFile " document.getElementById('cb$name1[lindex $ids 0]').checked = !document.getElementById('sw$name1[lindex $ids 0]').checked;\n\}</script>"
    }
  }
}

# -------------------------------------------------------------------------------
# generate x3d rotation (axis angle format) from axis2_placement_3d
proc x3dGetRotation {axis refdir {type ""}} {

# check if one of the vectors is zero length, i.e., '0 0 0'
  set msg ""
  if {[veclen $axis] == 0 || [veclen $refdir] == 0} {
    set msg "Syntax Error: The axis2_placement_3d axis or ref_direction vector is '0 0 0'"
    if {$type != ""} {append msg " for a $type"}
    append msg "."

# check if axis and refdir are parallel
  } elseif {[veclen [veccross $axis $refdir]] == 0} {
    set msg "Syntax Error: The axis2_placement_3d axis and ref_direction vectors '$refdir' are parallel"
    if {$type != ""} {append msg " for a $type"}
    append msg "."
  }
  if {$msg != "" && [string first "counter" $msg] == -1 && [string first "hole" $msg] == -1} {errorMsg $msg}

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

# -------------------------------------------------------------------------------
# generate STEP AP242 tessellated geometry from STL file for viewing
proc stl2STEP {} {
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
