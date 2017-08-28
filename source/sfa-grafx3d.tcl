# start x3dom file for PMI annotations                        
proc x3dFileStart {} {
  global ao entCount localName opt savedViewNames x3dFile x3dFileName x3dStartFile numTessColor x3dMin x3dMax
  
  set x3dStartFile 0
  catch {file delete -force -- "[file rootname $localName]_x3dom.html"}
  set x3dFileName [file rootname $localName]-x3dom.html
  catch {file delete -force -- $x3dFileName}
  set x3dFile [open $x3dFileName w]

  if {[string first "occurrence" $ao] != -1} {
    if {([info exists entCount(tessellated_solid)] || [info exists entCount(tessellated_shell)]) && $opt(VIZTES)} {
      set title "Graphical PMI and Tessellated Part Geometry"
    } else {
      set title "Graphical PMI"
    }
  } else {
    set title "Tessellated Part Geometry"
  }
  
  puts $x3dFile "<!DOCTYPE html>\n<html>\n<head>\n<title>[file tail $localName] | $title</title>\n<base target=\"_blank\">\n<meta http-equiv='Content-Type' content='text/html;charset=utf-8'/>"
  puts $x3dFile "<link rel='stylesheet' type='text/css' href='https://www.x3dom.org/x3dom/release/x3dom.css'/>\n<script type='text/javascript' src='https://www.x3dom.org/x3dom/release/x3dom.js'></script>\n</head>"

# transparency script
  set numTessColor 0
  if {[string first "Tessellated" $title] != -1} {
    set numTessColor [tessCountColors]
    if {$numTessColor > 0} {
      puts $x3dFile "<script>function matTrans(trans){"
      for {set i 1} {$i <= $numTessColor} {incr i} {
        puts $x3dFile " document.getElementById('mat$i').setAttribute('transparency', trans);"
        puts $x3dFile " if (trans > 0) {document.getElementById('mat$i').setAttribute('solid', true);} else {document.getElementById('mat$i').setAttribute('solid', false);}"
      }
      puts $x3dFile "}\n</script>"
    }
  }

  puts $x3dFile "\n<body><font face=\"arial\">\n<h3>$title:  [file tail $localName]</h3>"
  puts $x3dFile "<ul><li>Only $title is shown.  Boundary representation (b-rep) part geometry can be viewed with <a href=\"https://www.cax-if.org/step_viewers.html\">STEP file viewers</a>."
  if {[string first "Tessellated" $title] != -1 && [info exist entCount(next_assembly_usage_occurrence)]} {
    puts $x3dFile "<li>Parts in an assembly might have the wrong position and orientation or be missing."
  }
  puts $x3dFile "</ul>\n<table><tr><td>"

# x3d window size
  set height 800
  set width [expr {int($height*1.5)}]
  catch {
    set height [expr {int([winfo screenheight .]*0.75)}]
    set width [expr {int($height*[winfo screenwidth .]/[winfo screenheight .])}]
  }
  puts $x3dFile "\n<X3D id='someUniqueId' showStat='false' showLog='false' x='0px' y='0px' width='$width\px' height='$height\px'>\n<Scene DEF='scene'>"
  
# read tessellated geometry separately because of IFCsvr limitations
  if {($opt(VIZPMI) && [info exists entCount(tessellated_annotation_occurrence)]) || \
      ($opt(VIZTES) && ([info exists entCount(tessellated_solid)] || [info exists entCount(tessellated_shell)]))} {
    tessReadGeometry
  }
  outputMsg " Writing Visualization to: [truncFileName [file nativename $x3dFileName]]" green

# coordinate min, max, center
  if {$x3dMax(x) != -1.e10} {
    foreach xyz {x y z} {
      set delt($xyz) [expr {$x3dMax($xyz)-$x3dMin($xyz)}]
      set xyzcen($xyz) [format "%.4f" [expr {0.5*$delt($xyz) + $x3dMin($xyz)}]]
    }
    set maxxyz $delt(x)
    if {$delt(y) > $maxxyz} {set maxxyz $delt(y)}
    if {$delt(z) > $maxxyz} {set maxxyz $delt(z)}

# coordinate axes    
    set asize [trimNum [expr {$maxxyz*0.05}]]
    x3dCoordAxes $asize
  }
  update idletasks
}

# -------------------------------------------------------------------------------
# write tessellated geometry for annotations and parts
proc x3dTessGeom {objID objEntity1 ent1} {
  global ao draftModelCameras entCount nshape shapeRepName savedViewFile tessIndex tessIndexCoord tessCoord tessCoordID
  global tessPlacement tessRepo x3dColor x3dCoord x3dIndex x3dFile x3dColors x3dMsg
  
  set x3dIndex $tessIndex($objID)
  set x3dCoord $tessCoord($tessIndexCoord($objID))

  if {$x3dColor == ""} {set x3dColor "1 1 0"}
  set x3dIndexType "Line"
  set solid ""
  set emit "emissiveColor='$x3dColor'"
  set spec ""

# faces
  if {[string first "face" $ent1] != -1} {
    set x3dIndexType "Face"
    set solid "solid='false'"

# tessellated part geometry
    if {$ao == "tessellated_solid" || $ao == "tessellated_shell"} {
      set tsID [$objEntity1 P21ID]
      set tessRepo 0

# set color
      set x3dColor ".7 .7 .7"
      tessSetColor $tsID
      set spec "specularColor='.5 .5 .5'"
      set emit ""

# set placement for tessellated part geometry in assemblies (axis and ref_direction)
      if {[info exists entCount(item_defined_transformation)]} {tessSetPlacement $tsID}
    }
  }

# write transform based on placement
  catch {unset endTransform}
  set nplace 0
  if {[info exists tessRepo]} {
    if {$tessRepo && [info exists tessPlacement(origin)]} {set nplace [llength $tessPlacement(origin)]}
  }
  if {$nplace == 0} {set nplace 1}

# multiple saved views, write to individual files, collected in x3dViewpoint below
  set flist $x3dFile
  if {[info exists draftModelCameras] && $ao == "tessellated_annotation_occurrence"} {
    set savedViewName [getSavedViewName $objEntity1]
    if {[llength $savedViewName] > 0} {
      set flist {}
      foreach svn $savedViewName {lappend flist $savedViewFile($svn)}
    }
  }

# loop over list of files from above
  foreach f $flist {
    for {set np 0} {$np < $nplace} {incr np} {

# translation and rotation (sometimes PMI and usually assemblies)
      if {$tessRepo && [info exists tessPlacement(origin)]} {
        puts $f "<Transform translation='[lindex $tessPlacement(origin) $np]' rotation='[x3dRotation [lindex $tessPlacement(axis) $np] [lindex $tessPlacement(refdir) $np]]'>"
        set endTransform "</Transform>"
        #errorMsg "$np / [lindex $tessPlacement(origin) $np] / [lindex $tessPlacement(axis) $np] / [lindex $tessPlacement(refdir) $np] / [x3dRotation [lindex $tessPlacement(axis) $np] [lindex $tessPlacement(refdir) $np]] / $ao"
      }

# write tessellated face or line
      if {$ao == "tessellated_annotation_occurrence" || ![info exists shapeRepName]} {set shapeRepName $x3dIndexType}
      if {$np == 0} {
        if {$emit == ""} {
          set matID ""
          set colorID [lsearch $x3dColors $x3dColor]
          if {$colorID == -1} {
            lappend x3dColors $x3dColor
            puts $f "<Shape DEF='$shapeRepName$objID'>\n <Appearance DEF='app[llength $x3dColors]'><Material id='mat[llength $x3dColors]' diffuseColor='$x3dColor' $spec></Material></Appearance>"
          } else {
            puts $f "<Shape DEF='$shapeRepName$objID'>\n <Appearance USE='app[incr colorID]'></Appearance>"
          }
        } else {
          puts $f "<Shape DEF='$shapeRepName$objID'>\n <Appearance><Material diffuseColor='$x3dColor' $emit></Material></Appearance>"
        }
        puts $f " <Indexed$x3dIndexType\Set $solid coordIndex='[string trim $x3dIndex]'>"
        if {[lsearch $tessCoordID $tessIndexCoord($objID)] == -1} { 
          lappend tessCoordID $tessIndexCoord($objID)
          puts $f "  <Coordinate DEF='coord$tessIndexCoord($objID)' point='[string trim $x3dCoord]'></Coordinate>"
        } else {
          puts $f "  <Coordinate USE='coord$tessIndexCoord($objID)'></Coordinate>"
        }
        puts $f " </Indexed$x3dIndexType\Set>\n</Shape>"
      } else {
        puts $f " <Shape USE='$shapeRepName$objID'></Shape>"
      }

# for tessellated part geometry only, write mesh based on faces
      set mesh 0
      if {[info exists entCount(triangulated_face)]}         {if {$entCount(triangulated_face)         < 10000} {set mesh 1}}
      if {[info exists entCount(complex_triangulated_face)]} {if {$entCount(complex_triangulated_face) < 10000} {set mesh 1}}

      if {$x3dIndexType == "Face" && ($ao == "tessellated_solid" || $ao == "tessellated_shell") && $mesh} {
        if {$np == 0} {
          set x3dMesh ""
          set firstID [lindex $x3dIndex 0]
          set getFirst 0
          foreach id [split $x3dIndex " "] {
            if {$id == -1} {
              append x3dMesh "$firstID "
              set getFirst 1
            }
            append x3dMesh "$id "
            if {$id != -1 && $getFirst} {
              set firstID $id
              set getFirst 0
            }
          }
          
          set ecolor ""
          foreach c [split $x3dColor] {append ecolor "[expr {$c*.5}] "}
          puts $x3dFile "<Shape DEF='Mesh$objID'>\n <Appearance><Material emissiveColor='$ecolor'></Material></Appearance>"
          puts $x3dFile " <IndexedLineSet coordIndex='[string trim $x3dMesh]'>\n  <Coordinate USE='coord$tessIndexCoord($objID)'></Coordinate>"
          puts $x3dFile " </IndexedLineSet>\n</Shape>"
        } else {
          puts $x3dFile " <Shape USE='Mesh$objID'></Shape>"
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
# write geometry for polyline annotations
proc x3dPolylinePMI {} {
  global ao x3dCoord x3dShape x3dIndex x3dIndexType x3dFile x3dColor gpmiPlacement placeOrigin placeAnchor boxSize
  global savedViewName savedViewFile

  if {[catch {
    if {[info exists x3dCoord] || $x3dShape} {

# multiple saved views, write to individual files, collected in x3dViewpoint below
      set flist $x3dFile
      if {[llength $savedViewName] > 0} {
        set flist {}
        foreach svn $savedViewName {lappend flist $savedViewFile($svn)}
      }
  
      foreach f $flist {
        if {[string length $x3dCoord] > 0} {

# placeholder transform
          if {[string first "placeholder" $ao] != -1} {
            puts $f "<Transform translation='$gpmiPlacement(origin)' rotation='[x3dRotation $gpmiPlacement(axis) $gpmiPlacement(refdir)]'>"
          }  

# start shape
          if {$x3dColor != ""} {
            puts $f "<Shape>\n <Appearance><Material diffuseColor='$x3dColor' emissiveColor='$x3dColor'></Material></Appearance>"
          } elseif {[string first "annotation_occurrence" $ao] == 0 || [string first "annotation_fill_area_occurrence" $ao] == 0} {
            puts $f "<Shape>\n <Appearance><Material diffuseColor='1 0.5 0' emissiveColor='1 0.5 0'></Material></Appearance>"
            errorMsg "Syntax Error: Color not specified for PMI Presentation (using orange)"
          }

# index and coordinates
          puts $f " <IndexedLineSet coordIndex='[string trim $x3dIndex]'>\n  <Coordinate point='[string trim $x3dCoord]'></Coordinate>\n </IndexedLineSet>\n</Shape>"

# end placeholder transform, add leader line
          if {[string first "placeholder" $ao] != -1} {
            puts $f "</Transform>"
            puts $f "<Shape>\n <Appearance><Material emissiveColor='$x3dColor'></Material></Appearance>"
            puts $f " <IndexedLineSet coordIndex='0 1 -1'>\n  <Coordinate point='$placeOrigin $placeAnchor '></Coordinate>\n </IndexedLineSet>\n</Shape>"
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
  global x3dFile x3dAxes
  
# axes
  if {$x3dAxes} {
    puts $x3dFile "\n<Shape id='x_axis'><Appearance><Material emissiveColor='1 0 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. $size 0. 0.'></Coordinate></IndexedLineSet></Shape>"
    puts $x3dFile "<Shape id='y_axis'><Appearance><Material emissiveColor='0 .5 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. $size 0.'></Coordinate></IndexedLineSet></Shape>"
    puts $x3dFile "<Shape id='z_axis'><Appearance><Material emissiveColor='0 0 1'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. 0. $size'></Coordinate></IndexedLineSet></Shape>"

# xyz labels
    set tsize [trimNum [expr {$size*0.33}]]
    puts $x3dFile "<Transform translation='$size 0. 0.' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='1 0 0'></Material></Appearance><Text string='\"X\"'><FontStyle family='\"SANS\"'></FontStyle></Text></Shape></Billboard></Transform>"
    puts $x3dFile "<Transform translation='0. $size 0.' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='0 .5 0'></Material></Appearance><Text string='\"Y\"'><FontStyle family='\"SANS\"'></FontStyle></Text></Shape></Billboard></Transform>"
    puts $x3dFile "<Transform translation='0. 0. $size' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='0 0 1'></Material></Appearance><Text string='\"Z\"'><FontStyle family='\"SANS\"'></FontStyle></Text></Shape></Billboard></Transform>\n"
    set x3dAxes 0
  }
}

# -------------------------------------------------------------------------------
# write PMI saved view geometry, set viewpoints, add navigation and background color, and close X3DOM file
proc x3dFileEnd {} {
  global modelURLs nistName opt stepAP x3dAxes x3dMax x3dMin x3dFile x3dMsg stepAP entCount nistVersion numTessColor
  global savedViewButtons savedViewFileName savedViewFile savedViewNames savedViewpoint
  
# write any PMI saved view geometry for multiple saved views
  set savedViewButtons {}
  if {[llength $savedViewNames] > 0} {
    foreach svn $savedViewNames {
      if {[file size $savedViewFileName($svn)] > 0} {
        lappend savedViewButtons $svn
        close $savedViewFile($svn)
        set f [open $savedViewFileName($svn) r]
        puts $x3dFile "\n<Switch whichChoice='0' id='sw$svn'><Group>"
        while {[gets $f line] >= 0} {puts $x3dFile $line}
        puts $x3dFile "</Group></Switch>"
        close $f
        unset savedViewFile($svn)
      }
      catch {file delete -force $savedViewFileName($svn)}
    }
  }

# coordinate min, max, center    
  foreach xyz {x y z} {
    set delt($xyz) [expr {$x3dMax($xyz)-$x3dMin($xyz)}]
    set xyzcen($xyz) [trimNum [format "%.4f" [expr {0.5*$delt($xyz) + $x3dMin($xyz)}]]]
  }
  set maxxyz $delt(x)
  if {$delt(y) > $maxxyz} {set maxxyz $delt(y)}
  if {$delt(z) > $maxxyz} {set maxxyz $delt(z)}

# coordinate axes, if not already written
  if {$x3dAxes} {
    set asize [trimNum [expr {$maxxyz*0.05}]]
    x3dCoordAxes $asize
  }

# default and saved viewpoints
  set cor "centerOfRotation='$xyzcen(x) $xyzcen(y) $xyzcen(z)'"
  set fov [trimNum [expr {$delt(z)*0.5 + $delt(y)*0.5}]]
  set psy [trimNum [expr {0. - ($xyzcen(y) + 1.4*$maxxyz)}]]

  puts $x3dFile "\n<Viewpoint id='Front' position='$xyzcen(x) $psy $xyzcen(z)' orientation='1 0 0 1.5708' $cor></Viewpoint>"
  if {[llength $savedViewNames] > 0} {
    foreach svn $savedViewNames {
      if {[info exists savedViewpoint($svn)] && [lsearch $savedViewButtons $svn] != -1} {
        puts $x3dFile "<Transform translation='[lindex $savedViewpoint($svn) 0]'><Viewpoint id='$svn' position='[lindex $savedViewpoint($svn) 0]' orientation='[lindex $savedViewpoint($svn) 1]' $cor></Viewpoint></Transform>"  
        #set px [trimNum [expr {[lindex [lindex $savedViewpoint($svn) 0] 0]+$xyzcen(x)}]]
        #set py [trimNum [expr {[lindex [lindex $savedViewpoint($svn) 0] 1]+$psy}]]
        #set pz [trimNum [expr {[lindex [lindex $savedViewpoint($svn) 0] 2]+$xyzcen(z)}]]
        #puts $x3dFile "<OrthoViewpoint description='$svn' position='$px $py $pz' orientation='[lindex $savedViewpoint($svn) 1]' $cor fieldOfView='\[-$fov,-$fov,$fov,$fov\]'></OrthoViewpoint>"  
      }
    }
  }
  puts $x3dFile "<OrthoViewpoint id='Ortho' position='$xyzcen(x) $psy $xyzcen(z)' orientation='1 0 0 1.5708' $cor fieldOfView='\[-$fov,-$fov,$fov,$fov\]'></OrthoViewpoint>"  

# navigation, background color
  puts $x3dFile "\n<NavigationInfo type='\"EXAMINE\" \"ANY\"'></NavigationInfo>"
  if {[string first "AP209" $stepAP] == -1} {
    puts $x3dFile "<Background skyColor='.8 .8 .8'></Background>"
  } else {
    puts $x3dFile "<Background skyColor='1 1 1'></Background>"
  }
  puts $x3dFile "</Scene></X3D>\n\n</td><td valign='top'>"

# for NIST model - link to drawing 
  if {$nistName != ""} {
    foreach item $modelURLs {
      if {[string first $nistName $item] == 0} {puts $x3dFile "<a href=\"https://s3.amazonaws.com/nist-el/mfg_digitalthread/$item\">Test Case Drawing</a><p>"}
    }
  }

# for PMI annotations - checkboxes for toggling saved view graphics
  if {$opt(VIZPMI) && [llength $savedViewButtons] > 0} {
    puts $x3dFile "\nSaved View PMI"
    foreach svn $savedViewButtons {puts $x3dFile "<br><input type='checkbox' checked onclick='tog$svn\(this.value)'/>$svn"}
    puts $x3dFile "<p>Selecting a Saved View changes the viewpoint or Page Up for the next viewpoint.  Zoom and pan to view all PMI."
  }

# for FEM - node, element checkboxes
  if {$opt(VIZFEA) && [string first "AP209" $stepAP] == 0} {
    puts $x3dFile "\n<input type='checkbox' checked onclick='togNodes(this.value)'/>Nodes<br>"
    if {[info exists entCount(surface_3d_element_representation)] || \
        [info exists entCount(volume_3d_element_representation)]}  {puts $x3dFile "<input type='checkbox' checked onclick='togMesh(this.value)'/>Mesh<br>"}
    if {[info exists entCount(curve_3d_element_representation)]}   {puts $x3dFile "<input type='checkbox' checked onclick='tog1DElements(this.value)'/>1D Elements<br>"}
    if {[info exists entCount(surface_3d_element_representation)]} {puts $x3dFile "<input type='checkbox' checked onclick='tog2DElements(this.value)'/>2D Elements<br>"}
    if {[info exists entCount(volume_3d_element_representation)]}  {puts $x3dFile "<input type='checkbox' checked onclick='tog3DElements(this.value)'/>3D Elements<br>"}

# transparency slider
    if {[info exists entCount(surface_3d_element_representation)] || [info exists entCount(volume_3d_element_representation)]} {
      puts $x3dFile "\n<p>Transparency (approximate)<br>"
      puts $x3dFile "<input style='width:80px' type='range' min='0' max='0.8' step='0.2' value='0' onchange='matTrans(this.value)'/>"
    }

# different transparency slider
  } elseif {$numTessColor > 0} {
    puts $x3dFile "\n<p>Transparency (approximate)<br>"
    puts $x3dFile "<input style='width:80px' type='range' min='0' max='1' step='0.25' value='0' onchange='matTrans(this.value)'/>"
  }
  
# extra text messages
  if {[info exists x3dMsg]} {
    if {[llength $x3dMsg] > 0} {
      puts $x3dFile "\n<ul style=\"padding-left:20px\">"
      foreach item $x3dMsg {puts $x3dFile "<li>$item"}
      puts $x3dFile "</ul>"
      unset x3dMsg
    }
  }
  puts $x3dFile "\n<p><ul style=\"padding-left:20px\">"
  puts $x3dFile "<li><a href=\"https://www.x3dom.org/documentation/interaction/\">Use the mouse</a> in 'Examine Mode' to rotate, pan, zoom.<li>Use Page Up to switch between views.<p>"
  puts $x3dFile "</ul></td></tr></table>"

# toggle switches for PMI views
  if {[llength $savedViewButtons] > 0} {
    puts $x3dFile " "
    foreach svn $savedViewButtons {x3dSwitchScript $svn 1}
  }
                          
  set str "NIST "
  set url "https://www.nist.gov/services-resources/software/step-file-analyzer"
  if {!$nistVersion} {
    set str ""
    set url "https://github.com/usnistgov/SFA"
  }
  puts $x3dFile "\n<p>Generated by the <a href=\"$url\">$str\STEP File Analyzer (v[getVersion])</a> and rendered with <a href=\"https://www.x3dom.org/\">X3DOM</a>."
  puts $x3dFile "[clock format [clock seconds]]"

  puts $x3dFile "</font></body></html>"
  close $x3dFile
  update idletasks
  
  unset x3dMax
  unset x3dMin
}

# -------------------------------------------------------------------------------
# set x3d color
proc x3dSetColor {type} {
  global idxColor

  if {$type == 1} {return "0 0 0"}

  if {$type == 2} {
    incr idxColor
    switch $idxColor {
      1 {set color "0 0 0"}
      2 {set color "1 1 1"}
      3 {set color "1 0 0"}
      4 {set color ".5 .25 0"}
      5 {set color "1 1 0"}
      6 {set color "0 .5 0"}
      7 {set color "0 .5 .5"}
      8 {set color "0 0 1"}
      9 {set color "1 0 1"}
    }
    if {$idxColor == 9} {set idxColor 0}
  }
  return $color
} 

# -------------------------------------------------------------------------------------------------
# open X3DOM file 
proc openX3DOM {{fn ""}} {
  global opt x3dFileName multiFile stepAP lastX3DOM
  
  set f7 1  
  if {$fn == ""} {
    set f7 0
    set ok 0
    if {[info exists x3dFileName]} {if {[file exists $x3dFileName]} {set ok 1}}
    if {$ok} {
      set fn $x3dFileName
    } else {
      if {$opt(XLSCSV) == "None"} {errorMsg " There is nothing to Visualize (Options tab)."}
      return
    }
  }
  if {[file exists $fn] != 1} {return}
  if {![info exists multiFile]} {set multiFile 0}

  if {(($opt(VIZPMI) || $opt(VIZFEA) || $opt(VIZTES)) && $fn != "" && $multiFile == 0) || $f7} {
    outputMsg "\nOpening Visualization in the default Web Browser" green
    catch {.tnb select .tnb.status}
    set lastX3DOM $fn
    if {[catch {
      exec {*}[auto_execok start] "" $fn
    } emsg]} {
      errorMsg "No web browser is associated with HTML files.\n Open [truncFileName [file nativename $fn]] in a web browser that supports X3DOM.  https://www.x3dom.org/check/\n $emsg"
    }
    update idletasks
  }
}

# -------------------------------------------------------------------------------
# get saved view names
proc getSavedViewName {objEntity} {
  global savedViewName draftModelCameras draftModelCameraNames entCount

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
        lappend savedViewName $draftModelCameraNames([$entDraughtingModel P21ID])
      }
    }
  }
  return $savedViewName
}

# -------------------------------------------------------------------------------
# script for switch node
proc x3dSwitchScript {type {vp 0}} {
  global x3dFile
  
  puts $x3dFile "<script>function tog$type\(choice){"
  puts $x3dFile " if (!document.getElementById('sw$type').checked) {"
  puts $x3dFile "  document.getElementById('sw$type').setAttribute('whichChoice', -1);"
  puts $x3dFile " } else {"
  puts $x3dFile "  document.getElementById('sw$type').setAttribute('whichChoice', 0);"
  if {$vp} {puts $x3dFile "  document.getElementById('$type').setAttribute('set_bind','true');"}
  puts $x3dFile " }"
  puts $x3dFile " document.getElementById('sw$type').checked = !document.getElementById('sw$type').checked;\n}</script>"
}
