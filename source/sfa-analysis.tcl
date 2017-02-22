proc analysisStart {objDesign entType} {
  global ao aoEntTypes cells col elevel ent entAttrList gpmiRow nindex opt pmiCol pmiHeading pmiStartCol
  global recPracNames stepAP syntaxErr x3domColor x3domCoord x3domFile x3domIndex x3domShape

  if {$opt(DEBUG1)} {outputMsg "START analysisStart $entType" red}

# basic geometry
  set cartesian_point [list cartesian_point coordinates]
  set node            [list node items $cartesian_point]

  set PMIP(node) $node
  set ao $entType

  set entAttrList {}
  set pmiCol 0
  set nindex 0
  set x3domShape 0
  set gpmiRow($ao) {}

  if {[info exists pmiHeading]} {unset pmiHeading}
  if {[info exists ent]}        {unset ent}

  outputMsg " Adding Analysis Model" green

  if {$opt(DEBUG1)} {outputMsg \n}
  set elevel 0
  #pmiSetEntAttrList $PMIP($ao)
  if {$opt(DEBUG1)} {outputMsg "entattrlist $entAttrList"}
  if {$opt(DEBUG1)} {outputMsg \n}
    
  set startent [lindex $PMIP(node) 0]
  set n 0
  set elevel 0
  
# get next unused column by checking if there is a colName
  #set pmiStartCol($ao) [getNextUnusedColumn $startent 3]

# process all annotation_occurrence entities, call analysisReport
  ::tcom::foreach objEntity [$objDesign FindObjects [join $startent]] {
    if {[$objEntity Type] == $startent} {
      if {$n < 10000000} {
        if {[expr {$n%2000}] == 0} {
          if {$n > 0} {outputMsg "  $n"}
          update idletasks
        }
        analysisReport $objEntity
        if {$opt(DEBUG1)} {outputMsg \n}
      }
      incr n
    }
  }
  set col($ao) $pmiCol
  
# write any remaining X3DOM
  if {[info exists x3domCoord] || $x3domShape} {
    if {[string length $x3domCoord] > 0} {
      puts $x3domFile " <indexedLineSet coordIndex='[string trim $x3domIndex]'>\n  <coordinate point='[string trim $x3domCoord]'></coordinate>\n </indexedLineSet>\n</shape>"
      set x3domCoord ""
      set x3domIndex ""
    } elseif {$x3domShape} {
      puts $x3domFile "</indexedLineSet></shape>"
    }
    set x3domShape 0
    set x3domColor ""
  }
}

# -------------------------------------------------------------------------------

proc analysisReport {objEntity} {
  global ao aoname assocGeom avgX3domColor badAttributes cells circleCenter col currX3domPointID curveTrim dirRatio dirType draftModelCameras
  global elevel ent entAttrList entCount geomType gpmiEnts gpmiID gpmiIDRow gpmiOK gpmiRow gpmiTypes gpmiTypesInvalid gpmiTypesPerFile gpmiValProp
  global iCompCurve iCompCurveSeg incrcol iPolyline localName nindex numCompCurve numCompCurveSeg numPolyline numX3domPointID
  global objEntity1 opt pmiCol pmiColumns pmiHeading pmiStartCol pointLimit prefix propDefIDS recPracNames savedViewCol stepAP syntaxErr 
  global x3domColor x3domCoord x3domFile x3domFileName x3domFileOpen x3domIndex x3domMax x3domMin x3domPoint x3domPointID x3domShape
  global nistVersion

  #outputMsg "analysisReport" red
  #if {[info exists gpmiOK]} {if {$gpmiOK == 0} {return}}

# elevel is very important, keeps track level of entity in hierarchy
  incr elevel
  set ind [string repeat " " [expr {4*($elevel-1)}]]

  set maxcp $pointLimit

  if {[string first "handle" $objEntity] == -1} {
    #if {$objEntity != ""} {outputMsg "$ind $objEntity"}
    #outputMsg "  $objEntity" red
  } else {
    set objType [$objEntity Type]
    set objID   [$objEntity P21ID]
    set objAttributes [$objEntity Attributes]
    set ent($elevel) $objType
    if {$elevel == 1} {set objEntity1 $objEntity}

    if {$opt(DEBUG1)} {outputMsg "$ind ENT $elevel #$objID=$objType (ATR=[$objAttributes Count])" blue}
    #if {$elevel == 1} {outputMsg "#$objID=$objType" blue}

# write any leftover X3DOM from previous Shape
    if {[info exists x3domCoord] || $x3domShape} {
      if {[string length $x3domCoord] > 0} {
        puts $x3domFile " <indexedLineSet coordIndex='[string trim $x3domIndex]'>\n  <coordinate point='[string trim $x3domCoord]'></coordinate>\n </indexedLineSet>\n</shape>"
        set x3domCoord ""
        set x3domIndex ""
      } elseif {$x3domShape} {
        puts $x3domFile "</indexedLineSet></shape>"
      }
      set x3domShape 0
      set x3domColor ""
    }
    
    ::tcom::foreach objAttribute $objAttributes {
      set objName  [$objAttribute Name]
      set ent1 "$ent($elevel) $objName"
      set ent2 "$ent($elevel).$objName"
      set okattr 1
      #outputMsg "$ent1 $okattr" blue

      if {$okattr} {
        set objValue    [$objAttribute Value]
        set objNodeType [$objAttribute NodeType]
        set objSize     [$objAttribute Size]
        set objAttrType [$objAttribute Type]
  
        set idx [lsearch $entAttrList $ent1]
        #outputMsg "$ent1  $objValue $objNodeType $objSize $objAttrType"

# -----------------
# nodeType = 18,19
        if {$objNodeType == 18 || $objNodeType == 19} {
          if {[catch {
            if {$idx != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $elevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}

# if referred to another, get the entity
              if {[string first "handle" $objEntity] != -1} {analysisReport $objValue}
            }
          } emsg3]} {
            errorMsg "ERROR processing PMI Presentation ($objNodeType $ent2): $emsg3"
          }

# --------------
# nodeType = 20
        } elseif {$objNodeType == 20} {
          if {[catch {
            set idx 0
            if {$idx != -1} {
              if {$opt(DEBUG1)} {outputMsg "$ind   ATR $elevel $objName - $objValue ($objNodeType, $objSize, $objAttrType)"}
          
# start of a list of cartesian points, assuming it is for a polyline, elevel = 3
              if {$objAttrType == "ListOfcartesian_point" && $elevel == 3} {
                #outputMsg 1elevel$elevel red
                if {$maxcp <= 10 && $maxcp < $objSize} {
                  append x3domPointID "($maxcp of $objSize) cartesian_point "
                } else {
                  append x3domPointID "($objSize) cartesian_point "
                }
                set numX3domPointID $objSize
                set currX3domPointID 0
                incr iPolyline
    
                set str ""
                for {set i 0} {$i < $objSize} {incr i} {
                  append x3domIndex "[expr {$i+$nindex}] "
                }
                append x3domIndex "-1 "
                incr nindex $objSize
              }
    
              if {[info exists cells($ao)]} {
                set ok 0

# get values for these entity and attribute pairs
# g_c_s and a_f_a both start keeping track of their polylines
# cartesian_point is need to generated X3DOM
                switch -glob $ent1 {
                  "cartesian_point coordinates" {
                    outputMsg $objValue
                    if {$opt(VIZPMI) && $x3domFileName != ""} {
                      #outputMsg "$elevel $geomType $ent1" red

# elevel = 4 for polyline
                      if {$elevel == 4 && $geomType == "polyline"} {
                        append x3domCoord "[format "%.4f" [lindex $objValue 0]] [format "%.4f" [lindex $objValue 1]] [format "%.4f" [lindex $objValue 2]] " 
    
                        set x3domPoint(x) [lindex $objValue 0]
                        set x3domPoint(y) [lindex $objValue 1]
                        set x3domPoint(z) [lindex $objValue 2]

# min,max of points
                        foreach idx {x y z} {
                          if {$x3domPoint($idx) > $x3domMax($idx)} {set x3domMax($idx) $x3domPoint($idx)}
                          if {$x3domPoint($idx) < $x3domMin($idx)} {set x3domMin($idx) $x3domPoint($idx)}
                        }

# write coord and index to X3DOM file for polyline
                        if {$elevel == 4} {
                          if {$iPolyline == $numPolyline && $currX3domPointID == $numX3domPointID} {
                            outputMsg "polyline" blue
                            puts $x3domFile " <indexedLineSet coordIndex='[string trim $x3domIndex]'>\n  <coordinate point='[string trim $x3domCoord]'></coordinate>\n </indexedLineSet>\n</shape>"
                            set x3domCoord ""
                            set x3domIndex ""
                            set x3domShape 0
                            set x3domColor ""
                          }
                        }             

# circle center
                      } elseif {$geomType == "circle"} {
                        set circleCenter $objValue
                      }
                    }
                  }
                }
              }

# -------------------------------------------------
# recursively get the entities that are referred to
              if {[catch {
                ::tcom::foreach val1 $objValue {analysisReport $val1}
              } emsg]} {
                foreach val2 $objValue {analysisReport $val2}
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing PMI Presentation ($objNodeType $ent2): $emsg3"
          }

# ---------------------
# nodeType = 5 (!= 18,19,20)
        } else {
          if {[catch {
            if {$idx != -1} {
              if {$opt(DEBUG1) && $ent1 != "cartesian_point name"} {outputMsg "$ind   ATR $elevel $objName - $objValue ($objNodeType, $objAttrType)  ($ent1)"}
    
              if {[info exists cells($ao)]} {
                set ok 0
                set colName ""

# get values for these entity and attribute pairs
                switch -glob $ent1 {
                  "cartesian_point name" {
                    if {$elevel == 4} {
                      set ok 1
                      set col($ao) [expr {$pmiStartCol($ao)+2}]
                    }
                  }
                }

# value in spreadsheet
                if {$ok && [info exists gpmiID]} {
                  set c [string index [cellRange 1 $col($ao)] 0]
                  set r $gpmiIDRow($ao,$gpmiID)

# column name
                  if {$colName != ""} {
                    if {![info exists pmiHeading($col($ao))]} {
                      $cells($ao) Item 3 $c $colName
                      set pmiHeading($col($ao)) 1
                      set pmiCol [expr {max($col($ao),$pmiCol)}]
                    }
                  }

# keep track of rows with validation properties
                  if {[lsearch $gpmiRow($ao) $r] == -1} {lappend gpmiRow($ao) $r}

# look for correct PMI name on 
# geometric_curve_set  annotation_fill_area  tessellated_geometric_set  composite_curve
                  if {$ent1 == "geometric_curve_set name"  || \
                      $ent1 == "annotation_fill_area name" || \
                      $ent1 == "tessellated_geometric_set name" || \
                      $ent1 == "repositioned_tessellated_item_and_tessellated_geometric_set name" || \
                      $ent1 == "composite_curve name"} {
                    set ov $objValue
                
# start X3DOM file
                    if {$opt(VIZPMI)} {
                      if  {[string first "tessellated" $ao] == -1} {
                        if {$x3domFileOpen} {
                          set x3domFileOpen 0
                          set x3domFileName [file rootname $localName]_x3dom.html
                          catch {file delete -force $x3domFileName}
                          set x3domFile [open $x3domFileName w]
                          outputMsg " Writing PMI Annotations to: [truncFileName [file nativename $x3domFileName]]" green
                          
                          set str "NIST "
                          set url "http://go.usa.gov/yccx"
                          if {!$nistVersion} {
                            set str ""
                            set url "https://github.com/usnistgov/SFA"
                          }
                          
                          puts $x3domFile "<!DOCTYPE html>\n<html>\n<head>\n<title>[file tail $localName]</title>\n<base target=\"_blank\">\n<meta http-equiv='Content-Type' content='text/html;charset=utf-8'></meta>\n<link rel='stylesheet' type='text/css' href='http://www.x3dom.org/x3dom/release/x3dom.css'></link>\n<script type='text/javascript' src='http://www.x3dom.org/x3dom/release/x3dom.js'></script>\n</head>\n<body>"
                          puts $x3domFile "<FONT FACE=\"Arial\"><H3>PMI Presentation Annotations for:  [file tail $localName]</H3><UL>"
                          puts $x3domFile "<LI>Generated by the <a href=\"$url\">$str\STEP File Analyzer (v[getVersion])</A> on [clock format [clock seconds]] and rendered with <A HREF=\"http://www.x3dom.org/\">X3DOM</A>."
                          puts $x3domFile "<LI>Only the PMI annotations are shown.  Part geometry can be viewed with <A HREF=\"https://www.cax-if.org/step_viewers.html\">STEP viewers</A>."
                          puts $x3domFile "<LI><a href=\"http://www.x3dom.org/documentation/interaction/\">Use the mouse</a>, Page Up/Down keys, or touch gestures to rotate, pan, and zoom the annotations."
                          puts $x3domFile "</UL>"
                          puts $x3domFile "<x3d id='someUniqueId' showStat='false' showLog='false' x='0px' y='0px' width='1200px' height='900px'>\n<scene DEF='scene'>"

                          for {set i 0} {$i < 4} {incr i} {set avgX3domColor($i) 0}
                        }

# start X3DOM Shape node                    
                        if {$ao == "annotation_fill_area_occurrence"} {errorMsg "PMI annotations with filled characters are not filled."}
                        if {$x3domColor != ""} {
                          puts $x3domFile "<shape>\n <appearance><material emissiveColor='$x3domColor'></material></appearance>"
                          set colors [split $x3domColor " "]
                          for {set i 0} {$i < 3} {incr i} {set avgX3domColor($i) [expr {$avgX3domColor($i)+[lindex $colors $i]}]}
                          incr avgX3domColor(3)
                        } elseif {[string first "annotation_occurrence" $ao] == 0} {
                          puts $x3domFile "<shape>\n <appearance><material emissiveColor='1 0.5 0'></material></appearance>"
                          errorMsg "Syntax Error: Color not specified for PMI Presentation (using orange)"
                        } elseif {[string first "annotation_fill_area_occurrence" $ao] == 0} {
                          puts $x3domFile "<shape>\n <appearance><material emissiveColor='1 0.5 0'></material></appearance>"
                          errorMsg "Syntax Error: Color not specified for PMI Presentation (using orange)"
                        }
                        set x3domShape 1
                        update idletasks
                      } else {
                        errorMsg " Tessellated PMI Annotations are not supported." red
                      }
                    }               

# keep track of cartesian point ids (x3domPointID)
                  } elseif {[info exists currX3domPointID] && $ent1 == "cartesian_point name"} {
                    if {$currX3domPointID < $maxcp} {append x3domPointID "$objID "}
                    incr currX3domPointID

# cell value for presentation style or color
                  } else {
                  }
                }
              }
            }
          } emsg3]} {
            errorMsg "ERROR processing PMI Presentation ($objNodeType $ent2): $emsg3"
            set elevel 1
          }
        }
      }
    }
  }
  incr elevel -1
}
