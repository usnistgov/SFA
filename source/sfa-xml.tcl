# process AP242 XML
proc x3dReadXML {} {
  global ap242XML cadSystem developer localName mytemp opt parts stepAP timeStamp viz x3dBbox x3dFile x3dMax x3dMin x3dParts x3dViewOK

  set x3dViewOK 0
  set stepAP "AP242 XML"
  catch {unset parts}
  catch {unset x3dParts}

  set f [open $localName r]
  fconfigure $f -encoding utf-8
  if {[catch {
    outputMsg "Opening AP242 XML file"
    set xmldoc [dom parse [read $f]]
    close $f

    set cadSystem [[[$xmldoc getElementsByTagName OriginatingSystem] firstChild] nodeValue]
    set timeStamp [[[$xmldoc getElementsByTagName TimeStamp] firstChild] nodeValue]

# get File
    set xmlFiles [$xmldoc getElementsByTagName File]
    set nfiles [llength $xmlFiles]

# stp2x3d executable
    if {$nfiles > 0} {
      x3dCopySTP2X3D
      set stp2x3d [file join $mytemp stp2x3d-part.exe]
      if {![info exists stp2x3d]} {return}
    } else {
      return
    }

# get Document
    set xmlDocVer [$xmldoc getElementsByTagName DocumentVersion]
    foreach docver $xmlDocVer {
      set uid [$docver getAttribute uid]
      set node [$docver selectNodes Views/DocumentDefinition/Files/DigitalFile/attribute::uidRef]
      set uidRef [join [string range $node 8 end-1]]
      set mapuid($uid) $uidRef
      #outputMsg "mapuid $uid $uidRef" green
    }

# start x3d file
    x3dFileStart

# process files
    outputMsg " Generating X3D for [llength $xmlFiles] STEP files" blue
    foreach file $xmlFiles {
      set uid [$file getAttribute uid]
      set node [$file selectNodes Id/Identifier/attribute::id]
      if {$node == ""} {
        errorMsg "  Cannot find STEP file name with File/Id/Identifier/attribute::id - Checking File/Locations/ExternalItem/Id/attribute::id" red
        set node [$file selectNodes Locations/ExternalItem/Id/attribute::id]
      }
      set fname [join [string range $node 4 end-1]]
      set ext [file extension $fname]
      if {$ext == ".stp"} {
        outputMsg "  $fname  ($uid)"
        update idletasks

# generate X3D
        set fname [file join [file dirname $localName] $fname]
        if {[file exists $fname]} {
          set stpx3dFileName [string range $fname 0 [string last "." $fname]]
          append stpx3dFileName "x3d"
          catch {file delete -force -- $stpx3dFileName}
          set x3duid($uid) $stpx3dFileName

# run stp2x3d
          catch {exec $stp2x3d --input [file nativename $fname] --quality $opt(partQuality) --edge $opt(partEdges) --sketch $opt(partSketch) --normal $opt(partNormals)} errs($uid)
          if {[string first "Nothing to translate" $errs($uid)] != -1 || [string first "child" $errs($uid)] != -1} {
            outputMsg "   Error generating X3D\n$errs($uid)" red
            unset x3duid($uid)
          }
        } else {
          outputMsg "   STEP file not found: [truncFileName [file nativename $fname]]" red
        }
      } elseif {[string first "stp" $ext] != -1} {
        errorMsg "  File extension $ext not supported" red
      }
    }

# get Part
    set xmlParts [$xmldoc getElementsByTagName Part]
    outputMsg "\n Writing Parts to Viewer file" blue
    set ipart 0
    foreach part $xmlParts {
      set node [$part selectNodes Versions/PartVersion/Views/PartView/DocumentAssignment/AssignedDocument/attribute::uidRef]
      if {$node != ""} {
        set uidRef [string range $node 8 end-1]
        set uid1 $uidRef
        if {[info exists x3duid($uidRef)]} {
          set uid1 $uidRef
        } elseif {[info exists mapuid($uidRef)]} {
          if {[info exists x3duid($mapuid($uidRef))]} {set uid1 $mapuid($uidRef)}
        }
        #outputMsg "$uidRef $uid1" green

# get min and max
        if {[info exists errs($uid1)]} {
          foreach line [split $errs($uid1) "\n"] {
            set sline [split [string trim $line] " "]
            if {[string first "MinXYZ" $line] != -1} {
              foreach id1 {1 2 3} id2 {x y z} {
                set num [expr {[lindex $sline $id1]}]
                regsub -all "," $num "." num
                set x3dMin($id2) [expr {min($num,$x3dMin($id2))}]
              }
            } elseif {[string first "MaxXYZ" $line] != -1} {
              append x3dBbox "<br>Max:"
              foreach id1 {1 2 3} id2 {x y z} {
                set num [expr {[lindex $sline $id1]}]
                regsub -all "," $num "." num
                set x3dMax($id2) [expr {max($num,$x3dMax($id2))}]
              }
            }
          }

# read X3D and write to x3dom
          set fname $x3duid($uid1)
          set fname1 [file rootname [file tail $fname]]
          puts $x3dFile "\n<!-- $fname1 -->"
          puts $x3dFile "<Switch id='swPart$ipart' whichChoice='0'>"
          set parts($fname1) $ipart
          incr ipart
          outputMsg "  $ipart [file tail $fname]  ($uid1)"

          set fx3d [open $fname r]
          set write 0
          while {[gets $fx3d line] >= 0} {
            if {$line == "</Scene>"} {set write 0}
            if {$write} {puts $x3dFile $line}
            if {$line == "<Scene>"} {set write 1}
          }
          puts $x3dFile "</Switch>"
          close $fx3d
          catch {[file delete -force -- $fname]}
          update idletasks
        } else {
          errorMsg "  No file uidRef found ($uid1)" red
        }
      }
    }

    set x3dViewOK 1
    set viz(PART) 1

# bounding box
    set x3dBbox ""
    append x3dBbox "<br>Min:"
    foreach id2 {x y z} {
      append x3dBbox "&nbsp;&nbsp;$x3dMin($id2)"
    }
    append x3dBbox "<br>Max:"
    foreach id2 {x y z} {
      append x3dBbox "&nbsp;&nbsp;$x3dMax($id2)"
    }
    #if {$nfiles > 1} {set x3dBbox ""}
    if {$x3dBbox != ""} {set x3dBbox "Bounding Box$x3dBbox"}

# part names for list in viewer
    foreach name [lsort [array names parts]] {set x3dParts($name) $parts($name)}

# oops
  } emsg]} {
    errorMsg "Error processing XML file: $emsg"
    set ap242XML 0
    set viz(PART) 0
    set x3dViewOK 0
    catch {close $x3dFile}
  }
}
