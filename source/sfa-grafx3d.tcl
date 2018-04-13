# start x3dom file for PMI annotations                        
proc x3dFileStart {} {
  global ao entCount localName opt x3dFile x3dFileName x3dStartFile numTessColor x3dMin x3dMax cadSystem viz
  
  set x3dStartFile 0
  catch {file delete -force -- "[file rootname $localName]_x3dom.html"}
  catch {file delete -force -- "[file rootname $localName]-x3dom.html"}
  set x3dFileName [file rootname $localName]-sfa.html
  catch {file delete -force -- $x3dFileName}
  set x3dFile [open $x3dFileName w]

  if {$viz(PMI) && $viz(TPG)} {
    set title "Graphical PMI and Tessellated Part Geometry"
  } elseif {$viz(PMI)} {
    set title "Graphical PMI"
  } elseif {$viz(TPG)} {
    set title "Tessellated Part Geometry"
  }
  
  puts $x3dFile "<!DOCTYPE html>\n<html>\n<head>\n<title>[file tail $localName] | $title</title>\n<base target=\"_blank\">\n<meta http-equiv='Content-Type' content='text/html;charset=utf-8'/>"
  puts $x3dFile "<link rel='stylesheet' type='text/css' href='https://www.x3dom.org/x3dom/release/x3dom.css'/>\n<script type='text/javascript' src='https://www.x3dom.org/x3dom/release/x3dom.js'></script>\n</head>"

  set name [file tail $localName]
  if {$cadSystem != ""} {append name "  ($cadSystem)"}
  puts $x3dFile "\n<body><font face=\"arial\">\n<h3>$title:  $name</h3>"
  puts $x3dFile "Boundary representation (b-rep) part geometry can also be viewed with <a href=\"https://www.cax-if.org/step_viewers.html\">STEP file viewers</a>."
  puts $x3dFile "  B-rep geometry might include supplemental geometry."
  if {$viz(PMI)} {puts $x3dFile "<br>Some Graphical PMI might not have equivalent Semantic PMI in the STEP file."}
  if {$viz(TPG) && [info exist entCount(next_assembly_usage_occurrence)]} {
    puts $x3dFile "<br>Parts in an assembly might have the wrong position and orientation or be missing."
  }
  puts $x3dFile "\n<p><table><tr><td valign='top' width='85%'>"

# x3d window size
  set height 800
  set width [expr {int($height*1.5)}]
  catch {
    set height [expr {int([winfo screenheight .]*0.85)}]
    set width [expr {int($height*[winfo screenwidth .]/[winfo screenheight .])}]
  }
  puts $x3dFile "\n<X3D id='x3d' showStat='false' showLog='false' x='0px' y='0px' width='$width' height='$height'>\n<Scene DEF='scene'>"
  
# read tessellated geometry separately because of IFCsvr limitations
  if {($viz(PMI) && [info exists entCount(tessellated_annotation_occurrence)]) || $viz(TPG)} {tessReadGeometry}
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
  global ao draftModelCameras entCount nshape recPracNames shapeRepName savedViewFile tessIndex tessIndexCoord tessCoord tessCoordID
  global tessPlacement tessRepo x3dColor x3dCoord x3dIndex x3dFile x3dColors x3dMsg opt defaultColor tessPartFile tessSuppGeomFile shellSuppGeom
  global savedViewNames savedViewFileName mytemp
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
      set spec "specularColor='"
      foreach c [lindex $defaultColor 0] {append spec "[trimNum [expr {$c*0.7}]] "}
      set spec [string range $spec 0 end-1]'
      set emit ""
      tessSetColor $objEntity1 $tsID

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
    for {set np 0} {$np < $nplace} {incr np} {

# translation and rotation (sometimes PMI and usually assemblies)
      if {$tessRepo && [info exists tessPlacement(origin)]} {
        puts $f "<Transform translation='[lindex $tessPlacement(origin) $np]' rotation='[x3dRotation [lindex $tessPlacement(axis) $np] [lindex $tessPlacement(refdir) $np]]'>"
        set endTransform "</Transform>"
        #errorMsg "$np / [lindex $tessPlacement(origin) $np] / [lindex $tessPlacement(axis) $np] / [lindex $tessPlacement(refdir) $np] / [x3dRotation [lindex $tessPlacement(axis) $np] [lindex $tessPlacement(refdir) $np]] / $ao"
      }

# write tessellated face or line
      if {![info exists shapeRepName]} {set shapeRepName $x3dIndexType}
      if {$np == 0} {

        if {$emit == ""} {
          set matID ""
          set colorID [lsearch $x3dColors $x3dColor]
          if {$colorID == -1} {
            lappend x3dColors $x3dColor
            puts $f "<Shape DEF='$shapeRepName$objID'><Appearance DEF='app[llength $x3dColors]'><Material id='mat[llength $x3dColors]' diffuseColor='$x3dColor' $spec></Material></Appearance>"
          } else {
            puts $f "<Shape DEF='$shapeRepName$objID'><Appearance USE='app[incr colorID]'></Appearance>"
          }
        } else {
          puts $f "<Shape DEF='$shapeRepName$objID'><Appearance><Material diffuseColor='$x3dColor' $emit></Material></Appearance>"
        }
        
        set indexedSet "<Indexed[string totitle $x3dIndexType]\Set $solid coordIndex='[string trim $x3dIndex]'>"
        
        if {[lsearch $tessCoordID $tessIndexCoord($objID)] == -1} { 
          lappend tessCoordID $tessIndexCoord($objID)
          puts $f " $indexedSet\n  <Coordinate DEF='coord$tessIndexCoord($objID)' point='[string trim $x3dCoord]'></Coordinate></Indexed[string totitle $x3dIndexType]\Set></Shape>"
        } else {
          puts $f " $indexedSet<Coordinate USE='coord$tessIndexCoord($objID)'></Coordinate></Indexed[string totitle $x3dIndexType]\Set></Shape>"
        }
      } else {
        puts $f "<Shape USE='$shapeRepName$objID'></Shape>"
      }

# for tessellated part geometry only, write mesh based on faces
      if {$opt(VIZTPGMSH) || $ao == "tessellated_shell"} {
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
            if {$ao == "tessellated_shell"} {
              set ecolor "0 0 0"
              errorMsg " Triangular faces in 'tessellated_shell' are outlined in black."
              set msg "Triangular faces in tessellated shells are outlined in black."
              if {[lsearch $x3dMsg $msg] == -1} {lappend x3dMsg $msg}
            }
            puts $f "<Shape DEF='mesh$objID'><Appearance><Material emissiveColor='$ecolor'></Material></Appearance>"
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
# write tessellated edges, PMI saved view geometry, set viewpoints, add navigation and background color, and close X3DOM file
proc x3dFileEnd {} {
  global ao modelURLs nistName opt stepAP x3dAxes x3dMax x3dMin x3dFile x3dMsg stepAP entCount nistVersion numSavedViews numTessColor viz
  global savedViewButtons savedViewFileName savedViewFile savedViewNames savedViewpoint feaBoundary feaLoad savedViewItems feaLoadMsg feaLoadMag
  global tessEdgeFile tessEdgeFileName tessPartFile tessPartFileName tessEdgeCoordDef
  global wdir mytemp localName
  global objDesign

# coordinate min, max, center    
  foreach xyz {x y z} {
    set delt($xyz) [expr {$x3dMax($xyz)-$x3dMin($xyz)}]
    set xyzcen($xyz) [trimNum [format "%.4f" [expr {0.5*$delt($xyz) + $x3dMin($xyz)}]]]
  }
  set maxxyz $delt(x)
  if {$delt(y) > $maxxyz} {set maxxyz $delt(y)}
  if {$delt(z) > $maxxyz} {set maxxyz $delt(z)}

# write tessellated edges
  set viz(EDG) 0
  if {[info exists tessEdgeFile]} {
    close $tessEdgeFile
    if {[file size $tessEdgeFileName] > 0} {
      puts $x3dFile "\n<!-- TESSELLATED EDGES -->\n<Switch whichChoice='0' id='swTED'><Group>"
      set f [open $tessEdgeFileName r]
      while {[gets $f line] >= 0} {puts $x3dFile $line}
      close $f
      puts $x3dFile "</Group></Switch>"
      set viz(EDG) 1
      unset tessEdgeCoordDef
    }
    catch {file delete -force -- $tessEdgeFileName}
    unset tessEdgeFile
    unset tessEdgeFileName
  }

# supplemental geometry
  set viz(SMG) 0
  if {[info exists entCount(constructive_geometry_representation)]} {x3dSuppGeom $maxxyz}

# write any PMI saved view geometry for multiple saved views
  set savedViewButtons {}
  if {[llength $savedViewNames] > 0} {
    for {set i 0} {$i < [llength $savedViewNames]} {incr i} {
      set svn [lindex $savedViewNames $i]
      if {[file size $savedViewFileName($svn)] > 0} {
        set svMap($svn) $svn
        set svWrite 1
        close $savedViewFile($svn)
        
# check if same saved view graphics already written
        if {[info exists savedViewItems($svn)]} {
          #outputMsg "$svn $savedViewItems($svn)" blue
          for {set j 0} {$j < $i} {incr j} {
            set svn1 [lindex $savedViewNames $j]
            if {[info exists savedViewItems($svn1)]} {
              if {$savedViewItems($svn) == $savedViewItems($svn1)} {
                #outputMsg "$svn1 $savedViewItems($svn1)" green  
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
          set f [open $savedViewFileName($svn) r]
          while {[gets $f line] >= 0} {puts $x3dFile $line}
          close $f
          catch {unset savedViewFile($svn)}
        } else {
          puts $x3dFile "<!-- SAME AS $svMap($svn) -->"
        }
        puts $x3dFile "</Group></Switch>"
      } else {
        close $savedViewFile($svn)
      }
      catch {file delete -force -- $savedViewFileName($svn)}
    }
  }
  
# coordinate axes, if not already written
  if {$x3dAxes} {
    set asize [trimNum [expr {$maxxyz*0.05}]]
    x3dCoordAxes $asize
  }

# write tessellated part
  if {[info exists tessPartFile]} {
    puts $x3dFile "\n<!-- TESSELLATED PART GEOMETRY -->\n<Switch whichChoice='0' id='swTPG'><Group>"
    catch {close $tessPartFile}
    set f [open $tessPartFileName r]
    while {[gets $f line] >= 0} {puts $x3dFile $line}
    close $f
    puts $x3dFile "</Group></Switch>"
    catch {file delete -force -- $tessPartFileName}
    unset tessPartFile
    unset tessPartFileName
  }
  
# -------------------------------------------------------------------------------
# add b-rep geometry based on pythonOCC and OpenCascade
  set viz(BRP) 0
  if {([info exists entCount(advanced_brep_shape_representation)] || \
       [info exists entCount(manifold_surface_shape_representation)] || \
       [info exists entCount(manifold_solid_brep)]) && $opt(VIZBRP)} {x3dBrepGeom}

# -------------------------------------------------------------------------------
# default and saved viewpoints
  puts $x3dFile "\n<!-- VIEWPOINTS -->"
  set cor "centerOfRotation='$xyzcen(x) $xyzcen(y) $xyzcen(z)'"
  set fov [trimNum [expr {$delt(z)*0.5 + $delt(y)*0.5}]]
  set psy [trimNum [expr {0. - ($xyzcen(y) + 1.4*$maxxyz)}]]

  puts $x3dFile "<Viewpoint id='Front' position='$xyzcen(x) $psy $xyzcen(z)' orientation='1 0 0 1.5708' $cor></Viewpoint>"
  if {[llength $savedViewNames] > 0 && $opt(VIZPMIVP)} {
    foreach svn $savedViewNames {
      if {[info exists savedViewpoint($svn)] && [lsearch $savedViewButtons $svn] != -1} {
        puts $x3dFile "<Transform translation='[lindex $savedViewpoint($svn) 0]'><Viewpoint id='vp$svn' position='[lindex $savedViewpoint($svn) 0]' orientation='[lindex $savedViewpoint($svn) 1]' $cor></Viewpoint></Transform>"  
      }
    }
  }
  puts $x3dFile "<OrthoViewpoint id='Ortho' position='$xyzcen(x) $psy $xyzcen(z)' orientation='1 0 0 1.5708' $cor fieldOfView='\[-$fov,-$fov,$fov,$fov\]'></OrthoViewpoint>"  

# navigation, background color
  set bgc "1 1 1"
  if {$viz(PMI)} {set bgc ".8 .8 .85"}
  puts $x3dFile "\n<!-- BACKGROUND -->"
  puts $x3dFile "<Background id='BG' skyColor='$bgc'></Background>"
  puts $x3dFile "<NavigationInfo type='\"EXAMINE\" \"ANY\"'></NavigationInfo>"
  puts $x3dFile "</Scene></X3D>\n</td>\n\n<!-- START RIGHT COLUMN -->\n<td valign='top'>"

# -------------------------------------------------------------------------------
# for NIST model - link to drawing 
  if {$nistName != ""} {
    foreach item $modelURLs {
      if {[string first $nistName $item] == 0} {puts $x3dFile "<a href=\"https://s3.amazonaws.com/nist-el/mfg_digitalthread/$item\">NIST Test Case Drawing</a><p>"}
    }
  }
    
# BRP button
  if {$viz(BRP)} {
    puts $x3dFile "\n<!-- BRP button -->\n<input type='checkbox' checked onclick='togBRP(this.value)'/>B-rep Geometry<p>"
  }

# TPG button
  if {$viz(TPG) && ($viz(PMI) || $viz(SMG) || $viz(EDG))} {
    puts $x3dFile "\n<!-- TPG button -->\n<input type='checkbox' checked onclick='togTPG(this.value)'/>Tessellated Part Geometry"
    if {$viz(EDG)} {puts $x3dFile "<!-- TED button -->\n<br><input type='checkbox' checked onclick='togTED(this.value)'/>Lines (Tessellated Edges)"}
    puts $x3dFile "<p>"
  }
  
# SMG button
  if {$viz(SMG)} {
    puts $x3dFile "\n<!-- SMG button -->\n<input type='checkbox' checked onclick='togSMG(this.value)'/>Supplemental Geometry<p>"
  }

# for PMI annotations - checkboxes for toggling saved view graphics
  set svmsg {}
  if {[info exists numSavedViews($nistName)]} {
    if {$viz(PMI) && $nistName != "" && [llength $savedViewButtons] != $numSavedViews($nistName)} {
      lappend svmsg "For the NIST test case, expecting $numSavedViews($nistName) Graphical PMI Saved Views, found [llength $savedViewButtons]."
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
    if {!$ok && [info exists numSavedViews($nistName)] && [llength $savedViewButtons] <= $numSavedViews($nistName)} {
      lappend svmsg "For the NIST test case, some unexpected Graphical PMI Saved View names were found."
    }
    if {$opt(VIZPMIVP)} {
      puts $x3dFile "<p>Selecting a Saved View above changes the viewpoint.  Viewpoints usually have the correct orientation but are not centered.  Use pan and zoom to center the PMI."
    }
    puts $x3dFile "<hr><p>"
  }
  if {[llength $svmsg] > 0 && $viz(PMI)} {foreach msg $svmsg {errorMsg $msg}}

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
  
# background color buttons
  puts $x3dFile "\n<!-- Background buttons -->\nBackground Color<br>"
  set check1 "checked"
  set check2 ""
  if {$viz(PMI)} {
    set check2 "checked"
    set check1 ""
  }
  puts $x3dFile "<input type='radio' name='bgcolor' value='1 1 1' $check1 onclick='BGcolor(this.value)'/>White<br>"
  puts $x3dFile "<input type='radio' name='bgcolor' value='.8 .8 .8' $check2 onclick='BGcolor(this.value)'/>Gray<br>"
  puts $x3dFile "<input type='radio' name='bgcolor' value='0 0 0' onclick='BGcolor(this.value)'/>Black"
  
# axes button
  puts $x3dFile "\n<!-- Axes button -->\n<p><input type='checkbox' checked onclick='togAxes(this.value)'/>Origin"

# transparency slider
  set max 0
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
  }
  
# mouse message  
  puts $x3dFile "\n<p><a href=\"https://www.x3dom.org/documentation/interaction/\">Use the mouse</a> in 'Examine Mode' to rotate, pan, zoom.  Use Page Up to switch between views.  Use 'a' to show all."
  puts $x3dFile "</td></tr></table>"
  
# -------------------------------------------------------------------------------
# function for BRP 
  if {$viz(BRP)} {x3dSwitchScript BRP}
  
# function for TPG
  if {$viz(TPG) && ($viz(PMI) || $viz(SMG) || $viz(EDG))} {
    if {[string first "occurrence" $ao] == -1} {
      x3dSwitchScript TPG
      if {$viz(EDG)} {x3dSwitchScript TED}
    }
  }
  
# function for SMG 
  if {$viz(SMG)} {x3dSwitchScript SMG}

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

# transparency function
  set numTessColor 0
  if {$viz(TPG)} {set numTessColor [tessCountColors]}
  if {($numTessColor > 0 || $viz(BRP)) && [string first "AP209" $stepAP] == -1} {
    puts $x3dFile "\n<!-- Transparency function -->\n<script>function matTrans(trans){"
    if {$viz(BRP)} {
      puts $x3dFile " document.getElementById('color').setAttribute('transparency', trans);"
      #puts $x3dFile " if (trans > 0) {document.getElementById('color').setAttribute('solid', true);} else {document.getElementById('color').setAttribute('solid', false);}"
    }
    for {set i 1} {$i <= $numTessColor} {incr i} {
      puts $x3dFile " document.getElementById('mat$i').setAttribute('transparency', trans);"
    }
    puts $x3dFile "}\n</script>"
  }
                          
# credits
  set str "NIST "
  set url "https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer"
  if {!$nistVersion} {
    set str ""
    set url "https://github.com/usnistgov/SFA"
  }
  puts $x3dFile "\n<p>Generated by the <a href=\"$url\">$str\STEP File Analyzer and Viewer (v[getVersion])</a> and displayed with <a href=\"https://www.x3dom.org/\">X3DOM</a>."
  puts $x3dFile "[clock format [clock seconds]]"

  puts $x3dFile "</font></body></html>"
  close $x3dFile
  update idletasks
  
  unset x3dMax
  unset x3dMin
}

# -------------------------------------------------------------------------------
# B-rep geometry
proc x3dBrepGeom {} {
  global entCount opt viz mytemp wdir localName objDesign x3dFile

# copy executable
  if {[catch {
    set stp2x3d [file join $mytemp stp2x3d.exe]
    if {[file exists [file join $wdir stp2x3d stp2x3d.exe]]} {
      set copy 0
      if {![file exists $stp2x3d]} {
        set copy 1
      } elseif {[file mtime [file join $wdir stp2x3d stp2x3d.exe]] > [file mtime $stp2x3d]} {
        set copy 1
      }
      if {$copy} {file copy -force [file join $wdir stp2x3d stp2x3d.exe] $mytemp}
    }
              
# run stp2x3d    
    if {[file exists $stp2x3d]} {
      set stpx3dFileName [file rootname $localName].x3d
      catch {file delete -force $stpx3dFileName}
      outputMsg " Processing B-rep geometry. Wait for the popup program (stp2x3d.exe) to complete. See Options tab." green
      catch {exec $stp2x3d $localName} errs
      #outputMsg $errs red
      
# done processing      
      if {[string first "DONE!" $errs] != -1} {
        if {[file exists $stpx3dFileName]} {
          if {[file size $stpx3dFileName] > 0} {
            
# check for conversion units, mm > inch            
            set sc 1
            ::tcom::foreach e0 [$objDesign FindObjects [string trim advanced_brep_shape_representation]] {
              set e1 [[[$e0 Attributes] Item 3] Value]
              set a2 [[$e1 Attributes] Item 5]
              foreach e3 [$a2 Value] {
                if {[$e3 Type] == "conversion_based_unit_and_length_unit"} {
                  set e4 [[[$e3 Attributes] Item 3] Value]
                  set cf [[[$e4 Attributes] Item 1] Value]
                  set sc [trimNum [expr {1./$cf}] 5]
                }
              }
            }
            if {$sc == 1 && ![info exists entCount(advanced_brep_shape_representation)]} {
              ::tcom::foreach e0 [$objDesign FindObjects [string trim geometric_representation_context_and_global_uncertainty_assigned_context_and_global_unit_assigned_context]] {
                set a2 [[$e0 Attributes] Item 5]
                foreach e3 [$a2 Value] {
                  if {[$e3 Type] == "conversion_based_unit_and_length_unit"} {
                    set e4 [[[$e3 Attributes] Item 3] Value]
                    set cf [[[$e4 Attributes] Item 1] Value]
                    set sc [trimNum [expr {1./$cf}] 5]
                  }
                }
              }
            }
            
# integrate x3d from pythonOCC with existing x3dom file
            puts $x3dFile "\n<!-- B-REP GEOMETRY -->\n<Switch whichChoice='0' id='swBRP'><Transform scale='$sc $sc $sc'>"
            set stpx3dFile [open $stpx3dFileName r]
            set write 0
            while {[gets $stpx3dFile line] >= 0} {
              if {[string first "<Shape" $line] != -1} {
                set line "<Shape><Appearance>"
                set write 1
              }
              if {[string first "<Coordinate" $line] != -1} {
                set n 0
                while {[gets $stpx3dFile nline] >= 0} {
                  if {[string first "<Normal" $nline] != -1} {append line "\n"}
                  append line " $nline"
                  incr n
                  if {[expr {$n%3000}] == 0} {append line "\n"; set n 0}
                  if {[string first "</Normal" $nline] != -1} {break}
                }
              }
              if {$write} {puts $x3dFile $line}
              if {[string first "</Shape" $line] != -1} {set write 0}
            }
            close $stpx3dFile
            puts $x3dFile "</Transform></Switch>"
            set viz(BRP) 1
          }
        } else {
          set msg " ERROR: Cannot find b-rep geometry X3D file."
          if {[string first "." $stpx3dFileName] != [string last "." $stpx3dFileName]} {append msg "  Rename STEP file or pathname to remove '.' character."}
          errorMsg $msg
        }
        catch {file delete -force -- $stpx3dFileName}
      } else {
        errorMsg " ERROR generating visualization of B-rep geometry"
      }
    }
  } emsg]} {
    errorMsg " ERROR adding B-rep Geometry ($emsg)"
  }
}

# -------------------------------------------------------------------------------
# supplemental geometry
proc x3dSuppGeom {maxxyz} {
  global x3dFile objDesign viz developer tessSuppGeomFile tessSuppGeomFileName
  
  set defAxes  0
  set defPlane 0
  set size [trimNum [expr {$maxxyz*0.025}]]
  set tsize [trimNum [expr {$size*0.33}]]

  outputMsg " Processing supplemental geometry" green
  puts $x3dFile "\n<!-- SUPPLEMENTAL GEOMETRY -->\n<Switch whichChoice='0' id='swSMG'><Group>"
  ::tcom::foreach e0 [$objDesign FindObjects [string trim constructive_geometry_representation]] {
    set a1 [[$e0 Attributes] Item 2]
    ::tcom::foreach e2 [$a1 Value] {
      if {[catch {
        set ename [$e2 Type]
        switch $ename {
          plane -
          axis2_placement_3d {
            set name [[[$e2 Attributes] Item 1] Value]
            
            set e3 $e2
            if {$ename == "plane"} {set e3 [[[$e2 Attributes] Item 2] Value]}

# a2p3d
            set a2p3d [x3dGetA2P3D $e3]
            set origin [lindex $a2p3d 0]
            set axis   [lindex $a2p3d 1]
            set refdir [lindex $a2p3d 2]
            puts $x3dFile "<Transform translation='$origin' rotation='[x3dRotation $axis $refdir]'>"
          
# axes              
            if {$ename == "axis2_placement_3d"} {
              if {!$defAxes} {
                puts $x3dFile " <Group DEF='sgAxes'>"
                puts $x3dFile "  <Shape><Appearance><Material emissiveColor='1 0 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. $size 0. 0.'></Coordinate></IndexedLineSet></Shape>"
                puts $x3dFile "  <Shape><Appearance><Material emissiveColor='0 .5 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. $size 0.'></Coordinate></IndexedLineSet></Shape>"
                puts $x3dFile "  <Shape><Appearance><Material emissiveColor='0 0 1'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. 0. $size'></Coordinate></IndexedLineSet></Shape>"
                puts $x3dFile " </Group>"
                set defAxes 1
              } else {
                puts $x3dFile " <Group USE='sgAxes'></Group>"
              }
              set nsize [expr {$tsize*1.5}]
              if {$name != ""} {puts $x3dFile " <Transform scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='0 0 0'></Material></Appearance><Text string='\"$name\"'><FontStyle family='\"SANS\"' justify='\"BEGIN\"'></FontStyle></Text></Shape></Billboard></Transform>"}

# plane
            } elseif {$ename == "plane" || $ename == "circle"} {
              
# check if plane is bounded
              set bnds [$e2 GetUsedIn [string trim advanced_face] [string trim faced_geometry]]
              set bound 0
              ::tcom::foreach bnd $bnds {set bound 1}

              set nsize [expr {$size*2.}]
              if {!$defPlane} {
                if {$bound} {errorMsg " Bounding edges for supplemental geometry bounded planes are ignored."}
                puts $x3dFile " <Shape DEF='sgPlane'><Appearance><Material emissiveColor='0 0 1'></Material></Appearance><IndexedLineSet coordIndex='0 1 2 3 0 -1'><Coordinate point='-$nsize -$nsize 0. $nsize -$nsize 0. $nsize $nsize 0. -$nsize $nsize 0.'></Coordinate></IndexedLineSet></Shape>"
                set defPlane 1
              } else {
                puts $x3dFile " <Shape USE='sgPlane'></Shape>"
              }
              if {$name != ""} {puts $x3dFile " <Transform translation='-$nsize -$nsize 0.' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='0 0 1'></Material></Appearance><Text string='\"$name\"'><FontStyle family='\"SANS\"' justify='\"BEGIN\"'></FontStyle></Text></Shape></Billboard></Transform>"}
            }
            puts $x3dFile "</Transform>"
            set viz(SMG) 1
          }
          
          trimmed_curve -
          composite_curve -
          geometric_curve_set {
            catch {unset trim}
# curves
            if {$ename == "composite_curve"} {
              set e3s [[[$e2 Attributes] Item 2] Value]
              ::tcom::foreach e3 $e3s {set e4 [[[$e3 Attributes] Item 3] Value]}
              set e2 $e4
            }

# trimming with values OK, cartesian_points NG
            if {$ename == "trimmed_curve" || [$e2 Type] == "trimmed_curve"} {
              set trim(1) [[[$e2 Attributes] Item 3] Value]
              set trim(2) [[[$e2 Attributes] Item 4] Value]
            }
  
            set name [[[$e2 Attributes] Item 1] Value]
            set e3 [[[$e2 Attributes] Item 2] Value]
  
            if {[$e3 Type] == "trimmed_curve"} {
              set name [[[$e3 Attributes] Item 1] Value]
              set trim(1) [[[$e3 Attributes] Item 3] Value]
              set trim(2) [[[$e3 Attributes] Item 4] Value]
              set e4 [[[$e3 Attributes] Item 2] Value]
              set e3 $e4
            }
            if {$name == "BRepFtrWithContract"} {set name ""}

# line        
            if {[$e3 Type] == "line"} {
              
              set e4 [[[$e3 Attributes] Item 2] Value]
              set coord1 [vectrim [[[$e4 Attributes] Item 2] Value]]
              set e5 [[[$e3 Attributes] Item 3] Value]
              set mag [[[$e5 Attributes] Item 3] Value]
              set e6 [[[$e5 Attributes] Item 2] Value]
              set dir [[[$e6 Attributes] Item 2] Value]
              set coord2 [vectrim [vecscale $dir $mag]]
              if {[info exists trim(2)]} {if {[string first "handle" $trim(2)] == -1} {set coord2 [vectrim [vecscale $dir [expr {$trim(2)*$mag}]]]}}
              
              puts $x3dFile "<Transform translation='$coord1'>"
              puts $x3dFile " <Shape><Appearance><Material emissiveColor='1 0 1'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. $coord2'></Coordinate></IndexedLineSet></Shape>"
              if {$name != ""} {
                set nsize [expr {$tsize*0.5}]
                puts $x3dFile " <Transform scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='1 0 1'></Material></Appearance><Text string='\"$name\"'><FontStyle family='\"SANS\"' justify='\"BEGIN\"'></FontStyle></Text></Shape></Billboard></Transform>"
              }
              puts $x3dFile "</Transform>"
              set viz(SMG) 1

# circle        
            } elseif {[$e3 Type] == "circle"} {
              set e4 [[[$e3 Attributes] Item 2] Value]
              set rad [[[$e3 Attributes] Item 3] Value]
              set a2p3d [x3dGetA2P3D $e4]
              set origin [lindex $a2p3d 0]
              set axis   [lindex $a2p3d 1]
              set refdir [lindex $a2p3d 2]
              puts $x3dFile "<Transform translation='$origin' rotation='[x3dRotation $axis $refdir]'>"
              
# generate circle, account for trimming
              set ns 24
              set angle 0
              set dlt [expr {6.28319/$ns}]
              set trimmed 0
              if {[info exists trim(1)]} {
                if {[string first "handle" $trim(1)] == -1} {
                  set angle $trim(1)
                  set conv 1.
                  if {$trim(1) > 6.29 || $trim(2) > 6.29} {
                    set conv 0.01745
                    set angle [expr {$angle*$conv}]
                  }
                  set dlt [expr {$conv*($trim(2)-$trim(1))/$ns}]
                  incr ns
                  set trimmed 1
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
  
              puts $x3dFile " <Shape><Appearance><Material emissiveColor='1 0 1'></Material></Appearance><IndexedLineSet coordIndex='$index'><Coordinate point='$coord'></Coordinate></IndexedLineSet></Shape>"
              if {$name != ""} {puts $x3dFile " <Transform translation='$coord1' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='1 0 1'></Material></Appearance><Text string='\"$name\"'><FontStyle family='\"SANS\"' justify='\"BEGIN\"'></FontStyle></Text></Shape></Billboard></Transform>"}
              puts $x3dFile "</Transform>"
              set viz(SMG) 1

# other
            } else {
              errorMsg " Supplemental geometry for '[$e3 Type]' is not visualized."
            }
          }
        
# points
          cartesian_point {
            set name [[[$e2 Attributes] Item 1] Value]
            set coord1 [[[$e2 Attributes] Item 2] Value]
            set nsize [expr {$tsize*0.1}]
            puts $x3dFile "<Transform translation='$coord1'>"
            puts $x3dFile "  <Shape><Appearance><Material emissiveColor='0 0 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1 2 3 -1 4 5 -1'><Coordinate point='$nsize 0. 0. -$nsize 0. 0. 0. $nsize 0. 0. -$nsize 0. 0. 0. $nsize 0. 0. -$nsize'></Coordinate></IndexedLineSet></Shape>"
            if {$name != ""} {
              set nsize [expr {$tsize*0.25}]
              puts $x3dFile " <Transform scale='$nsize $nsize $nsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='0 0 0'></Material></Appearance><Text string='\"$name\"'><FontStyle family='\"SANS\"' justify='\"BEGIN\"'></FontStyle></Text></Shape></Billboard></Transform>"
            }
            puts $x3dFile "</Transform>"
            set viz(SMG) 1
          }
          
          default {
            if {$ename != "tessellated_shell" && $ename != "tessellated_wire"} {errorMsg " Supplemental geometry for '$ename' is not visualized."}
          }
        }
      } emsg]} {
        errorMsg " ERROR adding Supplemental Geometry for: $ename"
      }
    }
  }
  
# check for tessellated edges that are supplemental geometry  
  if {[info exists tessSuppGeomFile]} {
    close $tessSuppGeomFile
    if {[file size $tessSuppGeomFileName] > 0} {
      set f [open $tessSuppGeomFileName r]
      while {[gets $f line] >= 0} {puts $x3dFile $line}
      close $f
      puts $x3dFile "</Group></Switch>"
    }
    catch {file delete -force -- $tessSuppGeomFileName}
    unset tessSuppGeomFile
    unset tessSuppGeomFileName
  }
  puts $x3dFile "</Group></Switch>\n"
}

# -------------------------------------------------------------------------------
# write geometry for polyline annotations
proc x3dPolylinePMI {} {
  global ao x3dCoord x3dShape x3dIndex x3dIndexType x3dFile x3dColor gpmiPlacement placeOrigin placeAnchor boxSize
  global savedViewName savedViewNames savedViewFile savedViewFileName recPracNames mytemp

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
          } else {
            puts $f "<Shape>\n <Appearance><Material diffuseColor='0 0 0' emissiveColor='0 0 0'></Material></Appearance>"
            errorMsg "Syntax Error: Missing PMI Presentation color for [formatComplexEnt $ao] (using black)\n[string repeat " " 14]\($recPracNames(pmi242), Sec. 8.4, Figure 75)"
          }

# index and coordinates
          puts $f " <IndexedLineSet coordIndex='[string trim $x3dIndex]'>\n  <Coordinate point='[string trim $x3dCoord]'></Coordinate></IndexedLineSet>\n</Shape>"

# end placeholder transform, add leader line
          if {[string first "placeholder" $ao] != -1} {
            puts $f "</Transform>"
            puts $f "<Shape>\n <Appearance><Material emissiveColor='$x3dColor'></Material></Appearance>"
            puts $f " <IndexedLineSet coordIndex='0 1 -1'>\n  <Coordinate point='$placeOrigin $placeAnchor '></Coordinate></IndexedLineSet>\n</Shape>"
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
    puts $x3dFile "\n<!-- COORDINATE AXIS -->\n<Switch whichChoice='0' id='swAxes'><Group>"
    puts $x3dFile "<Shape id='x_axis'><Appearance><Material emissiveColor='1 0 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. $size 0. 0.'></Coordinate></IndexedLineSet></Shape>"
    puts $x3dFile "<Shape id='y_axis'><Appearance><Material emissiveColor='0 .5 0'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. $size 0.'></Coordinate></IndexedLineSet></Shape>"
    puts $x3dFile "<Shape id='z_axis'><Appearance><Material emissiveColor='0 0 1'></Material></Appearance><IndexedLineSet coordIndex='0 1 -1'><Coordinate point='0. 0. 0. 0. 0. $size'></Coordinate></IndexedLineSet></Shape>"

# xyz labels
    set tsize [trimNum [expr {$size*0.33}]]
    puts $x3dFile "<Transform translation='$size 0. 0.' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='1 0 0'></Material></Appearance><Text string='\"X\"'><FontStyle family='\"SANS\"'></FontStyle></Text></Shape></Billboard></Transform>"
    puts $x3dFile "<Transform translation='0. $size 0.' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='0 .5 0'></Material></Appearance><Text string='\"Y\"'><FontStyle family='\"SANS\"'></FontStyle></Text></Shape></Billboard></Transform>"
    puts $x3dFile "<Transform translation='0. 0. $size' scale='$tsize $tsize $tsize'><Billboard axisOfRotation='0 0 0'><Shape><Appearance><Material diffuseColor='0 0 1'></Material></Appearance><Text string='\"Z\"'><FontStyle family='\"SANS\"'></FontStyle></Text></Shape></Billboard></Transform>"
    puts $x3dFile "</Group></Switch>\n"
    set x3dAxes 0
  }
}

# -------------------------------------------------------------------------------
# set x3d color
proc x3dSetColor {type} {
  global idxColor

# black
  if {$type == 1} {return "0 0 0"}

# random
  if {$type == 2} {
    incr idxColor
    switch -- $idxColor {
      1 {set color "1 0 0"}
      2 {set color "1 1 0"}
      3 {set color "0 .5 0"}
      4 {set color "0 .5 .5"}
      5 {set color "0 0 1"}
      6 {set color "1 0 1"}
      7 {set color ".5 .25 0"}
      8 {set color "0 0 0"}
      9 {set color "1 1 1"}
    }
    if {$idxColor == 9} {set idxColor 0}
  }
  return $color
} 

# -------------------------------------------------------------------------------------------------
# open X3DOM file 
proc openX3DOM {{fn ""}} {
  global opt x3dFileName multiFile stepAP lastX3DOM entCount recPracNames viz
  
  set f7 1  
  if {$fn == ""} {
    set f7 0
    set ok 0
    if {[info exists x3dFileName]} {if {[file exists $x3dFileName]} {set ok 1}}
    if {$ok} {
      set fn $x3dFileName
    } else {
      if {$opt(XLSCSV) == "None"} {errorMsg "There is nothing selected to Visualize (Options tab) that is in the STEP file to display.\n See Websites > STEP File Viewers"}
      return
    }
  }
  if {[file exists $fn] != 1} {return}
  if {![info exists multiFile]} {set multiFile 0}
  
  set open 0
  if {$f7} {
    set open 1
  } elseif {($viz(PMI) || $viz(TPG) || $viz(FEA)) && $fn != "" && $multiFile == 0} {
    set open 1
  }
  if {$open} {
    outputMsg "\nOpening Visualization in the default Web Browser: [file tail $fn]" green
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
  global savedViewName draftModelCameras draftModelCameraNames entCount savedsavedViewNames
  
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
