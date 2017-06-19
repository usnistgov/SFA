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
# set X3DOM color
proc gpmiSetColor {type} {
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

# -------------------------------------------------------------------------------
# write geometry for polyline annotations
proc x3dShape {} {
  global ao x3dCoord x3dShape x3dIndex x3dIndexType x3dFile x3dColor gpmiPlacement placeOrigin placeAnchor boxSize
  global savedViewName savedViewNames savedViewFile

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
}

# -------------------------------------------------------------------------------
# write saved view geometry, set viewpoints, add navigation and background color, and close X3DOM file
proc x3dViewpoints {} {
  global modelURLs nistName opt stepAP x3dMax x3dMin x3dFile x3dMsg stepAP entCount nistVersion
  global savedViewNames savedViewFileName savedViewFile
  
# write any saved view geometry for multiple saved views
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
    }
  }
  
# viewpoints
  foreach idx {x y z} {
    set delt($idx) [expr {$x3dMax($idx)-$x3dMin($idx)}]
    set xyzcen($idx) [trimNum [format "%.4f" [expr {0.5*$delt($idx) + $x3dMin($idx)}]]]
  }
  set maxxyz $delt(x)
  if {$delt(y) > $maxxyz} {set maxxyz $delt(y)}
  if {$delt(z) > $maxxyz} {set maxxyz $delt(z)}

  set cor "centerOfRotation='$xyzcen(x) $xyzcen(y) $xyzcen(z)'"
  puts $x3dFile "\n<Viewpoint $cor position='$xyzcen(x) [trimNum [expr {0. - ($xyzcen(y) + 1.4*$maxxyz)}]] $xyzcen(z)' orientation='1 0 0 1.5708' description='Perspective'></Viewpoint>"

  set fov [trimNum [expr {$delt(z)*0.5 + $delt(y)*0.5}]]
  puts $x3dFile "<OrthoViewpoint fieldOfView='\[-$fov,-$fov,$fov,$fov\]' $cor position='$xyzcen(x) [trimNum [expr {0. - ($xyzcen(y) + 1.4*$maxxyz)}]] $xyzcen(z)' orientation='1 0 0 1.5708' description='Orthographic'></OrthoViewpoint>"  

  puts $x3dFile "<NavigationInfo type='\"EXAMINE\" \"ANY\"'></NavigationInfo>"
  if {[string first "AP209" $stepAP] == -1} {
    puts $x3dFile "<Background skyColor='.8 .8 .8'></Background>"
  } else {
    puts $x3dFile "<Background skyColor='1 1 1'></Background>"
  }
  puts $x3dFile "</Scene></X3D>\n</td><td valign='top'>"

# drawing
  if {$nistName != ""} {
    foreach item $modelURLs {
      if {[string first $nistName $item] == 0} {puts $x3dFile "<a href=\"https://s3.amazonaws.com/nist-el/mfg_digitalthread/$item\">Test Case Drawing</a><p>"}
    }
  }

# checkboxes for toggling saved view graphics
  if {[llength $savedViewButtons] > 0} {
    puts $x3dFile "Saved View PMI"
    foreach svn [lsort $savedViewButtons] {puts $x3dFile "<br><input type='checkbox' checked onclick='tog$svn\(this.value)'/>$svn"}
  }
  
# extra text messages
  if {[info exists x3dMsg]} {
    if {[llength $x3dMsg] > 0} {
      puts $x3dFile "<ul>"
      foreach item $x3dMsg {puts $x3dFile "<li>$item"}
      puts $x3dFile "</ul>"
      unset x3dMsg
    }
  }
  puts $x3dFile "</td></tr></table>\n<p>"
                          
  set str "NIST "
  set url "https://www.nist.gov/services-resources/software/step-file-analyzer"
  if {!$nistVersion} {
    set str ""
    set url "https://github.com/usnistgov/SFA"
  }
  puts $x3dFile "Generated by the <a href=\"$url\">$str\STEP File Analyzer (v[getVersion])</a> and rendered with <a href=\"https://www.x3dom.org/\">X3DOM</a>."
  puts $x3dFile "[clock format [clock seconds]]"

  puts $x3dFile "</font></body></html>"
  close $x3dFile
  update idletasks
  
  unset x3dMax
  unset x3dMin
}

# -------------------------------------------------------------------------------------------------
# open X3DOM file 
proc openX3DOM {} {
  global opt x3dFileName multiFile stepAP
  
  if {($opt(VIZPMI) || $opt(VIZFEA) || $opt(VIZTES)) && $x3dFileName != "" && $multiFile == 0} {
    outputMsg "\nOpening Graphics in the default Web Browser"
    if {[catch {
      exec {*}[auto_execok start] "" $x3dFileName
    } emsg]} {
      errorMsg "No application is associated with HTML files.  Open the file in a web browser.  https://www.x3dom.org/check/\n $emsg"
    }
    update idletasks
  }
}
