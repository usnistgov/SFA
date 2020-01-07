proc checkValues {} {
  global allNone appName appNames buttons developer edmWhereRules edmWriteToFile eeWriteToFile opt userEntityList useXL

  set butNormal {}
  set butDisabled {}

  if {[info exists buttons(appCombo)]} {
    set ic [lsearch $appNames $appName]
    if {$ic < 0} {set ic 0}
    catch {$buttons(appCombo) current $ic}

# Jotne EDM Model Checker
    if {$developer} {
      catch {
        if {[string first "EDM Model Checker" $appName] == 0} {
          pack $buttons(edmWriteToFile) -side left -anchor w -padx 5
          pack $buttons(edmWhereRules) -side left -anchor w -padx 5
        } else {
          pack forget $buttons(edmWriteToFile)
          pack forget $buttons(edmWhereRules)
        }
      }
    }

# STEP Tools
    catch {
      if {[string first "Conformance Checker" $appName] != -1} {
        pack $buttons(eeWriteToFile) -side left -anchor w -padx 5
      } else {
        pack forget $buttons(eeWriteToFile)
      }
    }
    
    catch {
      if {$appName == "Tree View (for debugging)"} {
        pack $buttons(indentGeometry) -side left -anchor w -padx 5
        pack $buttons(indentStyledItem) -side left -anchor w -padx 5
      } else {
        pack forget $buttons(indentGeometry)
        pack forget $buttons(indentStyledItem)
      }
    }
  }

# configure Excel, CSV, Viz only, Excel or not
  if {$opt(XLSCSV) == "Excel"} {
    catch {$buttons(genExcel) configure -text "Generate Spreadsheet"}
  } elseif {$opt(XLSCSV) == "CSV"} {
    catch {$buttons(genExcel) configure -text "Generate CSV Files"}
  } elseif {$opt(XLSCSV) == "None"} {
    catch {$buttons(genExcel) configure -text "Generate View"}
  }
  if {![info exists useXL]} {set useXL 1}

# no Excel
  if {!$useXL} {
    foreach item {INVERSE PMIGRF PMISEM VALPROP writeDirType} {set opt($item) 0}
    set opt(XL_OPEN) 1
    foreach item [array names opt] {
      if {[string first "PR_STEP" $item] == 0} {lappend butNormal "opt$item"}
    }
    foreach b {optHIDELINKS optINVERSE optPMIGRF optPMISEM optVALPROP optXL_FPREC optXL_SORT allNone2} {lappend butDisabled $b}
    foreach b {optVIZFEA optVIZPMI optVIZTPG optVIZBRP optVIZBRPEDG optVIZBRPNRM} {lappend butNormal $b}
    foreach b {allNone0 allNone1 allNone3 optPR_USER} {lappend butNormal $b}

# Excel
  } else {
    foreach item [array names opt] {
      if {[string first "PR_STEP" $item] == 0} {lappend butNormal "opt$item"}
    }
    foreach b {optHIDELINKS optINVERSE optPMIGRF optPMISEM optVALPROP optXL_FPREC optXL_SORT} {lappend butNormal $b}
    foreach b {optVIZFEA optVIZPMI optVIZTPG optVIZBRP optVIZBRPEDG optVIZBRPNRM} {lappend butNormal $b}
    foreach b {allNone0 allNone1 allNone2 allNone3 optPR_USER} {lappend butNormal $b}
  }

# viz only
  if {$opt(XLSCSV) == "None"} {
    set opt(PMIGRF) 0
    foreach item [array names opt] {
      if {[string first "PR_STEP" $item] == 0} {lappend butDisabled "opt$item"}
    }
    foreach b {optPMIGRF optPMIGRFCOV optPMISEM optPMISEMDIM optVALPROP optPR_USER optINVERSE} {lappend butDisabled $b}
    foreach b {allNone0 allNone1 allNone2} {lappend butDisabled $b}
    foreach b {userentity userentityopen} {lappend butDisabled $b}
    set userEntityList {}
    if {$opt(VIZFEA) == 0 && $opt(VIZPMI) == 0 && $opt(VIZTPG) == 0 && $opt(VIZBRP) == 0} {
      foreach item {VIZFEA VIZPMI VIZTPG VIZBRP} {set opt($item) 1}
    }
  }
  
# graphical PMI report
  if {$opt(PMIGRF)} {
    if {$opt(XLSCSV) != "None"} {
      foreach b {optPR_STEP_AP242 optPR_STEP_PRES optPR_STEP_REPR optPR_STEP_SHAP} {
        set opt([string range $b 3 end]) 1
        lappend butDisabled $b
      }
    }
    lappend butNormal optPMIGRFCOV
  } else {
    lappend butNormal optPR_STEP_PRES
    if {!$opt(VALPROP)} {lappend butNormal optPR_STEP_QUAN}
    if {!$opt(PMISEM)}  {foreach b {optPR_STEP_AP242 optPR_STEP_COMM optPR_STEP_SHAP optPR_STEP_REPR} {lappend butNormal $b}}
    lappend butDisabled optPMIGRFCOV
  }

# validation properties
  if {$opt(VALPROP)} {
    foreach b {optPR_STEP_AP242 optPR_STEP_QUAN optPR_STEP_REPR optPR_STEP_SHAP} {
      set opt([string range $b 3 end]) 1
      lappend butDisabled $b
    }
  } elseif {!$opt(PMIGRF)} {
    lappend butNormal optPR_STEP_QUAN
  }

# graphical PMI view
  if {$opt(VIZPMI)} {
    foreach b {gpmiColor0 gpmiColor1 gpmiColor2 gpmiColor3 linecolor} {lappend butNormal $b}
    if {$opt(XLSCSV) != "None"} {
      set opt(PR_STEP_PRES) 1
      lappend butDisabled optPR_STEP_PRES
    }
  } else {
    foreach b {gpmiColor0 gpmiColor1 gpmiColor2 gpmiColor3 linecolor} {lappend butDisabled $b}
  }

# FEM view
  if {$opt(VIZFEA)} {
    foreach b {optVIZFEABC optVIZFEALV optVIZFEADS} {lappend butNormal $b}
    if {$opt(VIZFEALV)} {
      foreach b {optVIZFEALVS} {lappend butNormal $b}
    } else {
      foreach b {optVIZFEALVS} {lappend butDisabled $b}
    }
    if {$opt(VIZFEADS)} {
      foreach b {optVIZFEADSntail} {lappend butNormal $b}
    } else {
      foreach b {optVIZFEADSntail} {lappend butDisabled $b}
    }
  } else {
    foreach b {optVIZFEABC optVIZFEALV optVIZFEALVS optVIZFEADS optVIZFEADSntail} {lappend butDisabled $b}
  }

# semantic PMI report
  if {$opt(PMISEM)} {
    foreach b {optPR_STEP_AP242 optPR_STEP_REPR optPR_STEP_SHAP optPR_STEP_TOLR optPR_STEP_QUAN optPR_STEP_FEAT} {
      set opt([string range $b 3 end]) 1
      lappend butDisabled $b
    }
    lappend butNormal optPMISEMDIM
  } else {
    foreach b {optPR_STEP_REPR optPR_STEP_TOLR optPR_STEP_FEAT} {lappend butNormal $b}
    if {!$opt(PMIGRF)} {
      if {!$opt(VALPROP)} {lappend butNormal optPR_STEP_QUAN}
      foreach b {optPR_STEP_AP242 optPR_STEP_COMM optPR_STEP_SHAP} {lappend butNormal $b}
    }
    lappend butDisabled optPMISEMDIM
  }

# part geometry view
  if {$opt(VIZBRP)} {
    if {$opt(XLSCSV) != "None"} {
      set opt(PR_STEP_PRES) 1
      lappend butDisabled optPR_STEP_PRES
    }
    foreach b {optVIZBRPEDG optVIZBRPNRM} {lappend butNormal $b}
  } else {
    catch {
      if {!$opt(PMISEM) && !$opt(PMIGRF)} {lappend butNormal optPR_STEP_COMM}
      if {!$opt(PMISEM)} {lappend butNormal optPR_STEP_PRES}
    }
    foreach b {optVIZBRPEDG optVIZBRPNRM} {lappend butDisabled $b}
  }

# tessellated geometry view
  if {$opt(VIZTPG)} {
    if {$opt(XLSCSV) != "None"} {
      set opt(PR_STEP_PRES) 1
      lappend butDisabled optPR_STEP_PRES
    }
    foreach b {optVIZTPGMSH} {lappend butNormal $b}
  } else {
    catch {
      if {!$opt(PMISEM) && !$opt(PMIGRF)} {lappend butNormal optPR_STEP_COMM}
      if {!$opt(PMISEM)} {lappend butNormal optPR_STEP_PRES}
    }
    foreach b {optVIZTPGMSH} {lappend butDisabled $b}
  }
  
# user-defined entity list
  if {$opt(PR_USER)} {
    foreach b {userentity userentityopen} {lappend butNormal $b}
  } else {
    foreach b {userentity userentityopen} {lappend butDisabled $b}
    set userEntityList {}
  }
  
# common for any report  
  if {$opt(PMISEM) || $opt(PMIGRF) || $opt(VALPROP)} {
    set opt(PR_STEP_COMM) 1
    lappend butDisabled optPR_STEP_COMM
  } else {
    lappend butNormal optPR_STEP_COMM
  }
  
  if {$developer} {
    if {$opt(INVERSE)} {
      foreach b {optDEBUGINV} {lappend butNormal $b}
    } else {
      foreach b {optDEBUGINV} {lappend butDisabled $b}
    }
  }
  
  if {$opt(writeDirType) == 0} {
    foreach b {userdir userentry userentry1 userfile} {lappend butDisabled $b}
  } elseif {$opt(writeDirType) == 1} {
    foreach b {userdir userentry}   {lappend butDisabled $b}
    foreach b {userentry1 userfile} {lappend butNormal   $b}
  } elseif {$opt(writeDirType) == 2} {
    foreach b {userdir userentry}   {lappend butNormal   $b}
    foreach b {userentry1 userfile} {lappend butDisabled $b}
  }

# make sure there is some entity type to process
  set nopt 0
  foreach idx [lsort [array names opt]] {
    if {([string first "PR_" $idx] == 0 || $idx == "VALPROP" || $idx == "PMIGRF" || $idx == "PMISEM") && [string first "FEAT" $idx] == -1} {
      incr nopt $opt($idx)
    }
  }
  if {$nopt == 0 && $opt(XLSCSV) != "None"} {set opt(PR_STEP_COMM) 1}
  
# configure buttons
  if {[llength $butNormal]   > 0} {foreach but $butNormal   {catch {$buttons($but) configure -state normal}}}
  if {[llength $butDisabled] > 0} {foreach but $butDisabled {catch {$buttons($but) configure -state disabled}}}
    
# configure all, none, for buttons
  if {[info exists allNone]} {
    if {($allNone == 2 && ($opt(PMISEM) != 1 || $opt(PMIGRF) != 1 || $opt(VALPROP) != 1)) ||
        ($allNone == 3 && ($opt(VIZPMI) != 1 || $opt(VIZTPG) != 1 || $opt(VIZFEA)  != 1 || $opt(VIZBRP) != 1))} {
      set allNone -1
    } elseif {$allNone == 0} {
      foreach item [array names opt] {
        if {[string first "PR_STEP" $item] == 0} {
          if {$item != "PR_STEP_GEOM" && $item != "PR_STEP_CPNT"} {
            if {$opt($item) == 0} {set allNone -1}
          }
        }
      }
    } elseif {$allNone == 1} {
      foreach item [array names opt] {
        if {[string first "PR_STEP" $item] == 0} {
          if {$item != "PR_STEP_COMM" && $item != "PR_STEP_FEAT"} {
            if {$opt($item) == 1} {set allNone -1}
          }
        }
      }
    }
  }
}

# -------------------------------------------------------------------------------------------------
proc setCoordMinMax {coord} {
  global x3dMax x3dMin x3dPoint

  set x3dPoint(x) [lindex $coord 0]
  set x3dPoint(y) [lindex $coord 1]
  set x3dPoint(z) [lindex $coord 2]

# min,max of points
  foreach idx {x y z} {
    if {$x3dPoint($idx) > $x3dMax($idx)} {set x3dMax($idx) $x3dPoint($idx)}
    if {$x3dPoint($idx) < $x3dMin($idx)} {set x3dMin($idx) $x3dPoint($idx)}
  }
}

# -------------------------------------------------------------------------------------------------
# set color based on entColorIndex variable
proc setColorIndex {ent {multi 0}} {
  global andEntAP209 entCategory entColorIndex stepAP
  
# special case
  if {[string first "geometric_representation_context" $ent] != -1} {set ent "geometric_representation_context"}
  
# simple entity, not compound with _and_
  foreach i [array names entCategory] {
    if {[info exist entColorIndex($i)]} {
      if {[lsearch $entCategory($i) $ent] != -1} {
        return $entColorIndex($i)
      }
    }
  }
  
# compound entity with _and_
  set c1 [string first "\_and\_" $ent]
  if {$c1 != -1} {
    set c2 [string last  "\_and\_" $ent]
    set tc1 "1000"
    set tc2 "1000"
    set tc3 "1000"
    
    foreach i [array names entCategory] {
      if {[info exist entColorIndex($i)]} {
        set ent1 [string range $ent 0 $c1-1]
        if {[lsearch $entCategory($i) $ent1] != -1} {
          #outputMsg "1 AND $ent  $ent1  $i  $entColorIndex($i)"
          set tc1 $entColorIndex($i)
        }
        if {$c2 == $c1} {
          set ent2 [string range $ent $c1+5 end]
          if {[lsearch $entCategory($i) $ent2] != -1} {
            #outputMsg "2 AND $ent  $ent2  $i  $entColorIndex($i)"
            set tc2 $entColorIndex($i)
          } 
        } elseif {$c2 != $c1} {
          set ent2 [string range $ent $c1+5 $c2-1]
          if {[lsearch $entCategory($i) $ent2] != -1} {
            #outputMsg "2 AND $ent  $ent2  $i  $entColorIndex($i)"
            set tc2 $entColorIndex($i)
          } 
          set ent3 [string range $ent $c2+5 end]
          if {[lsearch $entCategory($i) $ent3] != -1} {
            #outputMsg "3 AND $ent  $ent3  $i  $entColorIndex($i)"
            set tc3 $entColorIndex($i)
          }
        }
      }
    }
    set tc [expr {min($tc1,$tc2,$tc3)}]

# exception for STEP measures    
    if {$tc1 == $entColorIndex(PR_STEP_QUAN) || $tc2 == $entColorIndex(PR_STEP_QUAN) || $tc3 == $entColorIndex(PR_STEP_QUAN)} {
      set tc $entColorIndex(PR_STEP_QUAN)
    }

# fix some AP209 entities with '_and_'
    if {[string first "AP209" $stepAP] != -1} {foreach str $andEntAP209 {if {[string first $str $ent] != -1} {set tc 19}}}

    #outputMsg "TC $tc"
    if {$tc < 1000} {return $tc}
  }

# entity not in any category, color by AP
  if {[string first "AP209" $stepAP] != -1} {return 19} 
  if {$stepAP == "AP210"} {return 15} 
  if {$stepAP == "AP238"} {return 24}

# entity from other APs (no color)
  return -2      
}

#-------------------------------------------------------------------------------
proc openURL {url} {
  global pf32

# open in whatever is registered for the file extension, except for .cgi for upgrade url
  if {[string first ".cgi" $url] == -1} {
    if {[catch {
      exec {*}[auto_execok start] "" $url
    } emsg]} {
      if {[string first "is not recognized" $emsg] == -1} {
        if {[string first "UNC" $emsg] != -1} {set emsg [fixErrorMsg $emsg]}
        if {$emsg != ""} {errorMsg "ERROR opening $url: $emsg"}
      }
    }

# find web browser command  
  } else {
    set webCmd ""
    catch {
      set reg_wb [registry get {HKEY_CURRENT_USER\Software\Classes\http\shell\open\command} {}]
      set reg_wb [lindex [split $reg_wb "\""] 1]
      set webCmd $reg_wb
    }
    if {$webCmd == "" || ![file exists $webCmd]} {set webCmd [file join $pf32 "Internet Explorer" IEXPLORE.EXE]}
    exec $webCmd $url &
  }
}

#-------------------------------------------------------------------------------
proc openFile {{openName ""}} {
  global buttons editorCmd fileDir localName localNameList

  if {$openName == ""} {
  
# file types for file select dialog (removed .stpnc)
    set typelist [list {"STEP Files" {".stp" ".step" ".p21" ".stpZ"}} {"IFC Files" {".ifc"}}]

# file open dialog
    set localNameList [tk_getOpenFile -title "Open STEP File(s)" -filetypes $typelist -initialdir $fileDir -multiple true]
    if {[llength $localNameList] <= 1} {set localName [lindex $localNameList 0]}
    catch {
      set fext [string tolower [file extension $localName]]
      if {$fext == ".stpnc"} {errorMsg "Rename the file extension to '.stp' to process STEP-NC files."}
    }

# file name passed in as openName
  } else {
    set localName $openName
    set localNameList [list $localName]
  }

# multiple files selected
  if {[llength $localNameList] > 1} {
    set fileDir [file dirname [lindex $localNameList 0]]

    outputMsg "\nReady to process [llength $localNameList] STEP files" green
    if {[info exists buttons]} {
      $buttons(genExcel) configure -state normal
      if {[info exists buttons(appOpen)]} {$buttons(appOpen) configure -state normal}
      focus $buttons(genExcel)
    }

# single file selected
  } elseif {[file exists $localName]} {
    catch {pack forget $buttons(pgb1)}
  
# check for zipped file
    if {[string first ".stpz" [string tolower $localName]] != -1} {unzipFile}  

    set fileDir [file dirname $localName]
    if {[string first "z" [string tolower [file extension $localName]]] == -1} {
      outputMsg "\nReady to process: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)" green
      if {[info exists buttons]} {
        $buttons(genExcel) configure -state normal
        if {[info exists buttons(appOpen)]} {$buttons(appOpen) configure -state normal}
        focus $buttons(genExcel)
        if {$editorCmd != ""} {
          bind . <Key-F5> {
            if {[file exists $localName]} {
              outputMsg "\nOpening STEP file: [file tail $localName]"
              exec $editorCmd [file nativename $localName] &
            }
          }
        }
      }
    }
  
# not found
  } else {
    if {$localName != ""} {errorMsg "File not found: [truncFileName [file nativename $localName]]"}
  }
  catch {.tnb select .tnb.status}
}

#-------------------------------------------------------------------------------
proc unzipFile {} {
  global localName mytemp wdir

  if {[catch {
    outputMsg "\nUnzipping: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)"

# copy gunzip to TEMP
    if {[file exists [file join $wdir exe gunzip.exe]]} {file copy -force -- [file join $wdir exe gunzip.exe] $mytemp}

    set gunzip [file join $mytemp gunzip.exe]
    if {[file exists $gunzip]} {

# copy zipped file to TEMP
      if {[regsub ".stpZ" $localName ".stp.Z" ln] == 0} {regsub ".stpz" $localName ".stp.Z" ln}
      set fzip [file join $mytemp [file tail $ln]]
      file copy -force -- $localName $fzip

# get name of unzipped file
      set gz [exec $gunzip -Nl $fzip]
      set c1 [string first "%" $gz]
      set ftmp [string range $gz $c1+2 end]

# unzip to a tmp file
      if {[file tail $ftmp] != [file tail $fzip]} {outputMsg " Extracting: [file tail $ftmp]" blue}
      exec $gunzip -Nf $fzip

# copy to new stp file
      set fstp [file join [file dirname $localName] [file tail $ftmp]]
      set ok 0
      if {![file exists $fstp]} {
        set ok 1
      } elseif {[file mtime $localName] != [file mtime $fstp]} {
        outputMsg " Overwriting existing unzipped STEP file: [truncFileName [file nativename $fstp]]" red
        set ok 1
      }
      if {$ok} {file copy -force -- $ftmp $fstp}

      set localName $fstp
      file delete $fzip
      file delete $ftmp
    } else {
      errorMsg "ERROR: gunzip.exe not found to unzip compressed STEP file"
    }
  } emsg]} {
    errorMsg "ERROR unzipping file: $emsg"
  }
}

#-------------------------------------------------------------------------------
proc saveState {{ok 1}} {
  global buttons developer dispCmd dispCmds fileDir fileDir1 filesProcessed lastX3DOM lastXLS lastXLS1 mydocs openFileList
  global opt optionsFile sfaVersion statusFont upgrade upgradeIFCsvr userEntityFile userWriteDir userXLSFile

# ok = 0 only after installing IFCsvr from the command-line version  
  if {![info exists buttons] && $ok} {return}
  
  if {[catch {
    if {![file exists $optionsFile]} {outputMsg "\nCreating options file: [file nativename $optionsFile]"}
    set fileOptions [open $optionsFile w]
    puts $fileOptions "# Options file for the NIST STEP File Analyzer and Viewer v[getVersion] ([string trim [clock format [clock seconds]]])\n#\n# DO NOT EDIT OR DELETE FROM USER HOME DIRECTORY $mydocs\n# DOING SO WILL CORRUPT THE CURRENT SETTINGS OR CAUSE ERRORS IN THE SOFTWARE\n#"
    set varlist [list fileDir fileDir1 userWriteDir userEntityFile openFileList dispCmd dispCmds lastXLS lastXLS1 lastX3DOM \
                      userXLSFile statusFont upgrade upgradeIFCsvr sfaVersion filesProcessed]

    foreach var $varlist {
      if {[info exists $var]} {
        set vartmp [set $var]
        if {[string first "/" $vartmp] != -1 || [string first "\\" $vartmp] != -1 || [string first " " $vartmp] != -1} {
          if {$var != "dispCmds" && $var != "openFileList"} {
            regsub -all {\\} $vartmp "/" vartmp
            puts $fileOptions "set $var \"$vartmp\""
          } else {
            regsub -all {\\} $vartmp "/" vartmp
            regsub -all {\[} $vartmp "\\\[" vartmp
            regsub -all {\]} $vartmp "\\\]" vartmp
            for {set i 0} {$i < [llength $vartmp]} {incr i} {
              if {$i == 0} {
                if {[llength $vartmp] > 1} {
                  puts $fileOptions "set $var \"\{[lindex $vartmp $i]\} \\"
                } else {
                  puts $fileOptions "set $var \"\{[lindex $vartmp $i]\}\""
                }
              } elseif {$i == [expr {[llength $vartmp]-1}]} {
                puts $fileOptions "       \{[lindex $vartmp $i]\}\""
              } else {
                puts $fileOptions "       \{[lindex $vartmp $i]\} \\"
              }
            }
          }
        } else {
          if {$vartmp != ""} {
            puts $fileOptions "set $var [set $var]"
          } else {
            puts $fileOptions "set $var \"\""
          }
        }
      }
    }
    
    set winpos "+300+200"
    catch {
      set wg [winfo geometry .]
      set winpos [string range $wg [string first "+" $wg] end]
      set wingeo [string range $wg 0 [expr {[string first "+" $wg]-1}]]
    }
    catch {puts $fileOptions "set wingeo \"$wingeo\""}
    catch {puts $fileOptions "set winpos \"$winpos\""}

    foreach idx [lsort [array names opt]] {
      if {([string first "PR_" $idx] == -1 || [string first "PR_STEP" $idx] == 0 || [string first "PR_USER" $idx] == 0) && [string first "DEBUG" $idx] == -1} {
        set var opt($idx)
        set vartmp [set $var]
        if {[string first "/" $vartmp] != -1 || [string first "\\" $vartmp] != -1 || [string first " " $vartmp] != -1} {
          regsub -all {\\} $vartmp "/" vartmp
          puts $fileOptions "set $var \"$vartmp\""
        } else {
          if {$vartmp != ""} {
            puts $fileOptions "set $var [set $var]"
          } else {
            puts $fileOptions "set $var \"\""
          }
        }
      }
    }

    close $fileOptions
    if {$developer} {
      set f1 $optionsFile
      append f1 " - Copy"
      file copy -force -- $optionsFile $f1
    }

  } emsg]} {
    errorMsg "ERROR writing to options file: $emsg"
    catch {raise .}
  }
}

#-------------------------------------------------------------------------------
proc runOpenProgram {} {
  global appName dispCmd editorCmd edmWhereRules edmWriteToFile eeWriteToFile File localName

  set dispFile $localName
  set idisp [file rootname [file tail $dispCmd]]
  if {[info exists appName]} {if {$appName != ""} {set idisp $appName}}

  .tnb select .tnb.status
  outputMsg "\nOpening STEP file in: $idisp"

# open file
#  (list is programs that CANNOT start up with a file *OR* need specific commands below)
  if {[string first "Conformance"       $idisp] == -1 && \
      [string first "Tree View"         $idisp] == -1 && \
      [string first "Default"           $idisp] == -1 && \
      [string first "QuickStep"         $idisp] == -1 && \
      [string first "EDM Model Checker" $idisp] == -1} {

# start up with a file
    if {[catch {
      exec $dispCmd [file nativename $dispFile] &
    } emsg]} {
      errorMsg $emsg
    }

# default viewer associated with file extension
  } elseif {[string first "Default" $idisp] == 0} {
    if {[catch {
      exec {*}[auto_execok start] "" [file nativename $dispFile]
    } emsg]} {
      if {[string first "UNC" $emsg] != -1} {set emsg [fixErrorMsg $emsg]}
      if {$emsg != ""} {
        errorMsg "No application is associated with STEP files."
        errorMsg " See Websites > STEP File Viewers"
      }
    }

# file tree view
  } elseif {[string first "Tree View" $idisp] != -1} {
    indentFile $dispFile

# QuickStep
  } elseif {[string first "QuickStep" $idisp] != -1} {
    cd [file dirname $dispFile]
    exec $dispCmd [file tail $dispFile] &

#-------------------------------------------------------------------------------
# validate file with ST-Developer Conformance Checkers
  } elseif {[string first "Conformance" $idisp] != -1} {
    set stfile $dispFile
    outputMsg "Ready to validate: [file tail $stfile]" blue
    cd [file dirname $stfile]

# gui version
    if {[string first "gui" $dispCmd] != -1 && !$eeWriteToFile} {
      if {[catch {exec $dispCmd $stfile &} err]} {outputMsg "Conformance Checker error:\n $err" red}

# non-gui version
    } else {
      set stname [file tail $stfile]
      set stlog  "[file rootname $stname]\_stdev.log"
      catch {if {[file exists $stlog]} {file delete -force -- $stlog}}
      outputMsg "ST-Developer log file: [truncFileName [file nativename $stlog]]" blue

      set c1 [string first "gui" $dispCmd]
      set dispCmd1 $dispCmd
      if {$c1 != -1} {set dispCmd1 [string range $dispCmd 0 $c1-1][string range $dispCmd $c1+3 end]}

      if {[string first "apconform" $dispCmd1] != -1} {
        if {[catch {exec $dispCmd1 -syntax -required -unique -bounds -aggruni -arrnotopt -inverse -strwidth -binwidth -realprec -atttypes -refdom $stfile >> $stlog &} err]} {outputMsg "Conformance Checker error:\n $err" red}
      } else {
        if {[catch {exec $dispCmd1 $stfile >> $stlog &} err]} {outputMsg "Conformance Checker error:\n $err" red}
      }  
      if {[string first "Notepad++" $editorCmd] != -1} {
        outputMsg "Opening log file in editor"
        exec $editorCmd $stlog &
      } else {
        outputMsg "Wait until the Conformance Checker has finished and then open the log file"
      }
    }

#-------------------------------------------------------------------------------
# Jotne EDM Model Checker (only for developer)
  } elseif {[string first "EDM Model Checker" $idisp] != -1} {
    set filename $dispFile
    outputMsg "Ready to validate: [file tail $filename]" blue
    cd [file dirname $filename]

# write script file to open database
    set edmScript [file join [file dirname $filename] edm-validate-script.txt]
    catch {file delete -force -- $edmScript}
    set scriptFile [open $edmScript w]
    set okschema 1

    set edmDir [split [file nativename $dispCmd] [file separator]]
    set i [lsearch $edmDir "bin"]
    set edmDir [join [lrange $edmDir 0 [expr {$i-1}]] [file separator]]
    set edmDBopen "ACCUMULATING_COMMAND_OUTPUT,OPEN_SESSION"
    
# open file to find STEP schema name
    set edmPW "NIST@edm[string range $idisp end end]"
    set fschema [getSchemaFromFile $filename]

    if {[string first "AP203_CONFIGURATION_CONTROLLED_3D_DESIGN_OF_MECHANICAL_PARTS_AND_ASSEMBLIES_MIM_LF" $fschema] == 0} {
      puts $scriptFile "Database>Open([file nativename [file join $edmDir Db]], ap203, $edmPW, \"$edmDBopen\")"
    } elseif {$fschema == "CONFIG_CONTROL_DESIGN"} {
      puts $scriptFile "Database>Open([file nativename [file join $edmDir Db]], ap203e1, $edmPW, \"$edmDBopen\")"
    } elseif {[string first "AP209_MULTIDISCIPLINARY_ANALYSIS_AND_DESIGN_MIM_LF" $fschema] == 0} {
      puts $scriptFile "Database>Open([file nativename [file join $edmDir Db]], ap209, $edmPW, \"$edmDBopen\")"
    } elseif {$fschema == "AUTOMOTIVE_DESIGN"} {
      puts $scriptFile "Database>Open([file nativename [file join $edmDir Db]], ap214, $edmPW, \"$edmDBopen\")"
    } elseif {[string first "AP242_MANAGED_MODEL_BASED_3D_ENGINEERING_MIM_LF" $fschema] == 0} {
      set ap242 "ap242"
      if {[string first "442 2 1 4" $fschema] != -1 || [string first "442 3 1 4" $fschema] != -1} {append ap242 "e2"}
      puts $scriptFile "Database>Open([file nativename [file join $edmDir Db]], $ap242, $edmPW, \"$edmDBopen\")"
    } else {
      outputMsg "$idisp cannot be used with:\n $fschema" red
      set okschema 0
    }

# create a temporary file if certain characters appear in the name, copy original to temporary and process that one
    if {$okschema} {
      set tmpfile 0
      set fileRoot [file rootname [file tail $filename]]
      if {[string is integer [string index $fileRoot 0]] || \
        [string first " " $fileRoot] != -1 || \
        [string first "." $fileRoot] != -1 || \
        [string first "+" $fileRoot] != -1 || \
        [string first "%" $fileRoot] != -1 || \
        [string first "(" $fileRoot] != -1 || \
        [string first ")" $fileRoot] != -1} {
        if {[string is integer [string index $fileRoot 0]]} {set fileRoot "a_$fileRoot"}
        regsub -all " " $fileRoot "_" fileRoot
        regsub -all {[\.()+%]} $fileRoot "_" fileRoot
        set edmFile [file join [file dirname $filename] $fileRoot]
        append edmFile [file extension $filename]
        file copy -force -- $filename $edmFile
        set tmpfile 1
      } else {
        set edmFile $filename
      }

# validate everything
      #set edmValidate "FULL_VALIDATION,OUTPUT_STEPID"

# not validating DERIVE, ARRAY_REQUIRED_ELEMENTS
      set edmValidate "GLOBAL_RULES,REQUIRED_ATTRIBUTES,ATTRIBUTE_DATA_TYPE,AGGREGATE_DATA_TYPE,AGGREGATE_SIZE,AGGREGATE_UNIQUENESS,OUTPUT_STEPID"
      if {$edmWhereRules} {append edmValidate ",LOCAL_RULES,UNIQUENESS_RULES,INVERSE_RULES"}

# write script file if not writing output to file, just import model and validate
      set edmImport "ACCUMULATING_COMMAND_OUTPUT,KEEP_STEP_IDENTIFIERS,DELETING_EXISTING_MODEL,LOG_ERRORS_AND_WARNINGS_ONLY"
      if {$edmWriteToFile == 0} {
        puts $scriptFile "Data>ImportModel(DataRepository, $fileRoot, DataRepository, $fileRoot\_HeaderModel, \"[file nativename $edmFile]\", \$, \$, \$, \"$edmImport,LOG_TO_STDOUT\")"
        puts $scriptFile "Data>Validate>Model(DataRepository, $fileRoot, \$, \$, \$, \"ACCUMULATING_COMMAND_OUTPUT,$edmValidate,FULL_OUTPUT\")"

# write script file if writing output to file, create file names, import model, validate, and exit
      } else {
        set edmLog  "[file rootname $filename]_edm.log"
        set edmLogImport "[file rootname $filename]_edm_import.log"
        puts $scriptFile "Data>ImportModel(DataRepository, $fileRoot, DataRepository, $fileRoot\_HeaderModel, \"[file nativename $edmFile]\", \"[file nativename $edmLogImport]\", \$, \$, \"$edmImport,LOG_TO_FILE\")"
        puts $scriptFile "Data>Validate>Model(DataRepository, $fileRoot, \$, \"[file nativename $edmLog]\", \$, \"ACCUMULATING_COMMAND_OUTPUT,$edmValidate,FULL_OUTPUT\")"
        puts $scriptFile "Data>Close>Model(DataRepository, $fileRoot, \" ACCUMULATING_COMMAND_OUTPUT\")"
        puts $scriptFile "Data>Delete>ModelContents(DataRepository, $fileRoot, ACCUMULATING_COMMAND_OUTPUT)"
        puts $scriptFile "Data>Delete>Model(DataRepository, $fileRoot, header_section_schema, \"ACCUMULATING_COMMAND_OUTPUT,DELETE_ALL_MODELS_OF_SCHEMA\")"
        puts $scriptFile "Data>Delete>Model(DataRepository, $fileRoot, \$, ACCUMULATING_COMMAND_OUTPUT)"
        puts $scriptFile "Data>Delete>Model(DataRepository, $fileRoot, \$, \"ACCUMULATING_COMMAND_OUTPUT,CLOSE_MODEL_BEFORE_DELETION\")"
        puts $scriptFile "Exit>Exit()"
      }
      close $scriptFile

# run EDM Model Checker with the script file
      outputMsg "Running $idisp"
      eval exec {$dispCmd} $edmScript &

# if results are written to a file, open output file from the validation (edmLog) and output file if there are import errors (edmLogImport)
      if {$edmWriteToFile} {
        if {[string first "Notepad++" $editorCmd] != -1} {
          outputMsg "Opening log file(s) in editor"
          exec $editorCmd $edmLog &
          after 1000
          if {[file size $edmLogImport] > 0} {
            exec $editorCmd $edmLogImport &
          } else {
            catch {file delete -force -- $edmLogImport}
          }
        } else {
          outputMsg "Wait until the EDM Model Checker has finished and then open the log file"
        }
      }

# attempt to delete the script file
      set nerr 0
      while {[file exists $edmScript]} {
        catch {file delete -force -- $edmScript}
        after 1000
        incr nerr
        if {$nerr > 60} {break}
      }

# if using a temporary file, attempt to delete it
      if {$tmpfile} {
        set nerr 0
        while {[file exists $edmFile]} {
          catch {file delete -force -- $edmFile}
          after 1000
          incr nerr
          if {$nerr > 60} {break}
        }
      }
    }

# all others
  } else {
    .tnb select .tnb.status
    outputMsg "You have to manually import the STEP file to $idisp." red
    exec $dispCmd &
  }
  
    
# add file to menu
  addFileToMenu
  saveState
}

#-------------------------------------------------------------------------------
# open a spreadsheet
proc openXLS {filename {check 0} {multiFile 0}} {
  global buttons developer

  if {[info exists buttons]} {
    .tnb select .tnb.status
    update idletasks
  }

  if {[file exists $filename]} {

# check if instances of Excel are already running
    if {$check} {checkForExcel}
    
# start Excel
    if {[catch {
      set xl [::tcom::ref createobject Excel.Application]
      [$xl ErrorCheckingOptions] TextDate False

# errors
    } emsg]} {
      errorMsg "ERROR starting Excel: $emsg"
    }
    
# open spreadsheet in Excel, works even if Excel not already started above although slower
    if {[catch {
      outputMsg "\nOpening Spreadsheet: [file tail $filename]"
      exec {*}[auto_execok start] "" [file nativename $filename]

# errors
    } emsg]} {
      #if {$developer} {outputMsg $emsg red}
      if {[string first "UNC" $emsg] != -1} {set emsg [fixErrorMsg $emsg]}
      if {$emsg != ""} {
        if {[string first "The process cannot access the file" $emsg] != -1} {
          outputMsg " The Spreadsheet might already be opened." red
        } else {
          outputMsg " Error opening the Spreadsheet: $emsg" red
        }
        catch {raise .}
      }
    }

  } else {
    if {[file tail $filename] != ""} {
      outputMsg "\nOpening Spreadsheet: [file tail $filename]"
      outputMsg " Spreadsheet not found: [truncFileName [file nativename $filename]]" red
    }
    set filename ""
  }
  return $filename
}

#-------------------------------------------------------------------------------
proc checkForExcel {{multFile 0}} {
  global buttons lastXLS localName opt
  
  set pid1 [twapi::get_process_ids -name "EXCEL.EXE"]
  if {![info exists useXL]} {set useXL 1}
  
  if {[llength $pid1] > 0 && $opt(XLSCSV) != "None"} {
    if {[info exists buttons]} {
      if {!$multFile} {
        set msg "There are at least ([llength $pid1]) Excel spreadsheets already opened.\n\nDo you want to close the spreadsheets?"
        set dflt yes
        if {[info exists lastXLS] && [info exists localName]} {
          if {[llength $pid1] == 1} {if {[string first [file nativename [file rootname $localName]] [file nativename $lastXLS]] != 0} {set dflt no}}
        }
        set choice [tk_messageBox -type yesno -default $dflt -message $msg -icon question -title "Close Spreadsheets?"]

        if {$choice == "yes"} {
          for {set i 0} {$i < 5} {incr i} {
            set nnc 0
            foreach pid $pid1 {
              if {[catch {
                twapi::end_process $pid -force
              } emsg]} {
                incr nnc
              }
            }
            set pid1 [twapi::get_process_ids -name "EXCEL.EXE"]
            if {[llength $pid1] == 0} {break}
          }
        }
      }
    } else {
      foreach pid $pid1 {
        if {[catch {
          twapi::end_process $pid -force
        } emsg]} {
          #errorMsg "Some instances of Excel were not closed: $emsg" red
        }
      }
    }
  }
  return $pid1
}

#-------------------------------------------------------------------------------
# get next unused column
proc getNextUnusedColumn {ent} {
  global worksheet
  return [expr {[[[$worksheet($ent) UsedRange] Columns] Count] + 1}]
}

# -------------------------------------------------------------------------------
proc formatComplexEnt {str {space 0}} {
  global andEntAP209 entCategory opt stepAP
  
  set str1 $str

# possibly format for _and_
  if {[string first "_and_" $str1] != -1} {

# check if _and_ is part of the entity name
    set ok 1
    foreach cat {PR_STEP_AP242 PR_STEP_COMM PR_STEP_TOLR PR_STEP_PRES PR_STEP_KINE PR_STEP_COMP} {
      if {$opt($cat)} {if {[lsearch $entCategory($cat) $str] != -1} {set ok 0; break}}
    }
    if {[info exists stepAP]} {
      if {[string first "AP209" $stepAP] != -1} {foreach str2 $andEntAP209 {if {[string first $str2 $str] != -1} {set ok 0}}}
    }

# format a_and_b to (a)(b)
    if {$ok} {
      catch {
        regsub -all "_and_" $str1 ") (" str1
        if {$space == 0} {regsub -all " " $str1 "" str1}
        set str1 "($str1)"
      }
    }
  }
  return $str1
}

#-------------------------------------------------------------------------------
proc cellRange {r c} {
  global letters
  
# correct if 'c' is passed in as a letter
  set cf [string first $c $letters]
  if {$cf != -1} {set c [expr {$cf+1}]}

# a much more elegant solution from the Tcl wiki
  set cr ""
  set n $c
  while {[incr n -1] >= 0} {
    set cr [format %c%s [expr {$n%26+65}] $cr]
    set n [expr {$n/26}]
  }

  if {$r > 0} {
    append cr $r
  } else {
    append cr ":$cr"
  }
  
  return $cr
}

#-------------------------------------------------------------------------------
proc addCellComment {ent r c comment} {
  global recPracNames worksheet

  if {![info exists worksheet($ent)] || [string length $comment] < 2} {return}

# modify comment      
  if {[catch {
    while {[string first "  " $comment] != -1} {regsub -all "  " $comment " " comment}
    if {[string first "Syntax" $comment] == 0} {set comment "[string range $comment 14 end]"}
    if {[string first "GISU" $comment] != -1} {regsub "GISU" $comment "geometric_item_specific_usage"  comment}
    if {[string first "IIRU" $comment] != -1} {regsub "IIRU" $comment "item_identified_representation_usage" comment}

    foreach idx [array names recPracNames] {
      if {[string first $recPracNames($idx) $comment] != -1} {
        append comment "  See Websites > Recommended Practices"
        break
      }
    }
    
# add linefeeds for line length
    set ncomment ""
    set j 0
    for {set i 0} {$i < [string length $comment]} {incr i} {
      incr j
      set char [string index $comment $i]
      if {($j > 50 && $char == " ") || $char == [format "%c" 10]} {
        append ncomment \n
        set j 0
      } else {
        append ncomment $char
      }
    }
      
# add comment
    set comm [[$worksheet($ent) Range [cellRange $r $c]] AddComment]
    $comm Text $ncomment
    catch {[[$comm Shape] TextFrame] AutoSize [expr 1]}

# error
  } emsg]} {
    if {[string first "Unknown error" $emsg] == -1} {errorMsg "ERROR adding Cell Comment: $emsg\n  $ent"}
  }
}

#-------------------------------------------------------------------------------
# color bad cells red, add cell comment with message
proc colorBadCells {ent} {
  global cells count entsWithErrors excelVersion idRow legendColor stepAP syntaxErr worksheet
  
  if {$stepAP == "" || $excelVersion < 11} {return}
      
# color red for syntax errors
  set rmax [expr {$count($ent)+3}]
  set okcomment 0
  
  outputMsg " [formatComplexEnt $ent]" red
  set syntaxErr($ent) [lsort -integer -index 0 [lrmdups $syntaxErr($ent)]]
  foreach err $syntaxErr($ent) {
    #outputMsg "$ent $err" red
    set lastr 4
    if {[catch {

# get row and column number
      set r [lindex $err 0]
      set c [lindex $err 1]
      if {$r > 0 && ![info exists idRow($ent,$r)]} {return}
      
# r is entity ID, get row
      if {$r > 0} {
        if {[info exists idRow($ent,$r)]} {
          set r $idRow($ent,$r)
        } else {
          return
        }

# row number passed as a negative number
      } else {
        set r [expr {abs($r)}]
      }
      
# get message for cell comment
      set msg ""
      set msg [lindex $err 2]

# row and column are integers
      if {[string is integer $c]} {
        if {$r <= $rmax} {
          [[$worksheet($ent) Range [cellRange $r $c] [cellRange $r $c]] Interior] Color $legendColor(red)
          set okcomment 1
        }

# column is attribute name
      } else {
        #outputMsg "$ent / $r / $c / [string is integer $c]" red

# find column based on heading text
        set lastCol [[[$worksheet($ent) UsedRange] Columns] Count]
        if {![info exists nc($c)]} {
          for {set i 2} {$i <= $lastCol} {incr i} {
            set val [[$cells($ent) Item 3 $i] Value]
            if {[string first $c $val] == 0} {
              set nc($c) $i
              break
            }
          }
        }
        
# cannot find heading, use first column        
        if {![info exists nc($c)]} {set nc($c) 1}
        set c $nc($c)

# color cell
        if {$r <= $rmax} {
          [[$worksheet($ent) Range [cellRange $r $c] [cellRange $r $c]] Interior] Color $legendColor(red)
          set okcomment 1
        }
      }
      
# add cell comment
      if {$msg != "" && $okcomment} {if {$r <= $rmax} {addCellComment $ent $r $c $msg}}
      if {$okcomment} {lappend entsWithErrors [formatComplexEnt $ent]}

# error      
    } emsg]} {
      if {$emsg != ""} {
        errorMsg "ERROR setting cell color (red) or comment: $emsg\n  $ent"
        catch {raise .}
      }
    }
  }
}

#-------------------------------------------------------------------------------
proc trimNum {num {prec 3}} {
  global unq_num
  
# check for already trimmed number
  set numsav $num
  if {[info exists unq_num($numsav)]} {
    set num $unq_num($numsav)
  } else {
    
# trim number    
    if {[catch {
      
# format number with 'prec' 
      set form "\%."
      append form $prec
      append form "f"
      set num [format $form $num]

# remove trailing zeros
      if {[string first "." $num] != -1} {
        for {set i 0} {$i < $prec} {incr i} {
          set num [string trimright $num "0"]
        }
        if {$num == "-0"} {set num 0.}
      }
    } errmsg]} {
      errorMsg "# $errmsg ($numsav reset to 0.0)" red
      set num 0.
    }

# save the number for next occurrence
    set unq_num($numsav) $num
  }
  return $num
}

#-------------------------------------------------------------------------------
proc outputMsg {msg {color "black"}} {
  global logFile opt outputWin

  if {[info exists outputWin]} {
    $outputWin issue "$msg " $color
    update idletasks
  } else {
    puts $msg
  }
  if {$opt(LOGFILE) && [info exists logFile]} {puts $logFile $msg}
}

#-------------------------------------------------------------------------------
proc errorMsg {msg {color ""}} {
  global errmsg logFile opt outputWin stepAP

  set oklog 0
  if {$opt(LOGFILE) && [info exists logFile]} {set oklog 1}
  
# check if error message has already been used
  if {![info exists errmsg]} {set errmsg ""}
  if {[string first $msg $errmsg] == -1} {

# save current message to the beginning of errmsg
    set errmsg "$msg\n$errmsg"
    
# this fix is necessary to handle messages related to inverses
    set c1 [string first "DELETETHIS" $msg]
    if {$c1 != -1} {set msg [string range $msg 0 $c1-1]}

# syntax error
    if {$color == ""} {
      if {[string first "syntax error" [string tolower $msg]] != -1} {
        if {$stepAP != ""} {
          set logmsg "*** $msg"
          if {[info exists outputWin]} { 
            $outputWin issue "$msg " syntax
          } else {
            puts $logmsg
          }
        }

# regular error message, ilevel is the procedure the error was generated in
      } else {
        set ilevel ""
        catch {set ilevel "  \[[lindex [info level [expr {[info level]-1}]] 0]\]"}
        if {$ilevel == "  \[errorMsg\]"} {set ilevel ""}
        
        set logmsg "*** $msg$ilevel"
        if {[info exists outputWin]} { 
          $outputWin issue "$msg$ilevel " error
        } else {
          puts $logmsg
        }
      }

# error message with color
    } else {
      set logmsg "*** $msg"
      if {[info exists outputWin]} { 
        $outputWin issue "$msg " $color
      } else {
        puts $logmsg
      }
    }
    update idletasks

# add message to logfile
    if {$oklog && [info exists logmsg]} {
      if {[string first "*" $logmsg] == -1} {
        puts $logFile $logmsg
      } else {
        set newmsg [split [string range $logmsg 4 end] "\n"]
        set logmsg ""
        foreach str $newmsg {append logmsg "\n*** $str"}
        puts $logFile [string range $logmsg 1 end]
      }
    }
    return 1

# error message already used, do nothing
  } else {
    return 0
  }
}

# -------------------------------------------------------------------------------------------------
proc fixErrorMsg {emsg} {
  set emsg [split $emsg "\n"]
  if {[llength $emsg] > 3} {
    set emsg [join [lrange $emsg 3 end] "\n"]
  } else {
    set emsg ""
  }
  return $emsg
}

# -------------------------------------------------------------------------------------------------
proc truncFileName {fname {compact 0}} {
  global mydesk mydocs

  if {[string first $mydocs $fname] == 0} {
    set nname "[string range $fname 0 2]...[string range $fname [string length $mydocs] end]"
  } elseif {[string first $mydesk $fname] == 0 && $mydesk != $fname} {
    set nname "[string range $fname 0 2]...[string range $fname [string length $mydesk] end]"
  }

  if {[info exists nname]} {
    if {$nname != "C:\\..."} {set fname $nname}
  }

  if {$compact} {
    catch {
      while {[string length $fname] > 80} {
        set nname $fname
        set s2 0
        if {[string first "\\\\" $nname] == 0} {
          set nname [string range $nname 2 end]
          set s2 1
        }

        set nname [file nativename $nname]
        set sname [split $nname [file separator]]
        if {[llength $sname] <= 3} {break}

        if {[lindex $sname 1] == "..."} {
          set sname [lreplace $sname 2 2]
        } else {
          set sname [lreplace $sname 1 1 "..."]
        }

        set nname ""
        set nitem 0
        foreach item $sname {
          if {$nitem == 0 && [string length $item] == 2 && [string index $item 1] == ":"} {append item "/"}
          set nname [file join $nname $item]
          incr nitem
        }
        if {$s2} {set nname \\\\$nname}
        set fname [file nativename $nname]
      }
    }
  }
  return $fname
}

#-------------------------------------------------------------------------------
# create new file name if current file already exists
proc incrFileName {fn} {
  set fext [file extension $fn]
  set c1 [string last "." $fn]
  for {set i 1} {$i < 100} {incr i} {
    set fn "[string range $fn 0 $c1-1] ($i)$fext"
    catch {[file delete -force -- $fn]}
    if {![file exists $fn]} {break}
  }
  return $fn
}

#-------------------------------------------------------------------------------
# install IFCsvr (or remove to reinstall)
proc installIFCsvr {{exit 0}} {
  global buttons contact ifcsvrKey mydocs mytemp nistVersion upgradeIFCsvr wdir

# if IFCsvr is alreadly installed, get version from registry, decide to reinstall newer version
  if {[catch {

# get registry value "1.0.0 (NIST Update yyyy-mm-dd)"
    set ifcsvrKey "HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{3C8CE0A4-803B-48A6-96A0-A3DDD5AE5596}"
    set verIFCsvr [registry get $ifcsvrKey {DisplayVersion}]

# format version to be yyyymmdd
    set c1 [string first "20" $verIFCsvr]
    if {$c1 != -1} {
      set verIFCsvr [string range $verIFCsvr $c1 end-1]
      regsub -all {\-} $verIFCsvr "" verIFCsvr
    } else {
      set verIFCsvr 0
    }

# old version, reinstall      
    if {$verIFCsvr < [getVersionIFCsvr]} {
      set reinstall 1

# up-to-date, do nothing    
    } else {
      set reinstall 2
      set upgradeIFCsvr [clock seconds]
    }
    
# IFCsvr not installed or can't read registry    
  } emsg]} {
    set reinstall 0
  }

# up-to-date  
  if {$reinstall == 2} {return}

  set ifcsvr     "ifcsvrr300_setup_1008_en-update.msi"
  set ifcsvrInst [file join $wdir exe $ifcsvr]

  if {[info exists buttons]} {.tnb select .tnb.status}
  outputMsg " "
  
# first time installation
  if {!$reinstall} {
    errorMsg "The IFCsvr toolkit must be installed to read and process STEP files (User Guide section 2.2.1)."
    outputMsg "- You might need administrator privileges (Run as administrator) to install the toolkit.
  Antivirus software might respond that there is a security issue with the toolkit.  The
  toolkit is safe to install.  Use the default installation folder for the toolkit.
- To reinstall the toolkit, run the installation file ifcsvrr300_setup_1008_en-update.msi
  in $mytemp
- If there are problems with this procedure, contact [lindex $contact 0] ([lindex $contact 1])."

    if {[file exists $ifcsvrInst] && [info exists buttons]} {
      set msg "The IFCsvr toolkit must be installed to read and process STEP files (User Guide section 2.2.1).  After clicking OK the IFCsvr toolkit installation will start."
      append msg "\n\nYou might need administrator privileges (Run as administrator) to install the toolkit.  Antivirus software might respond that there is a security issue with the toolkit.  The toolkit is safe to install.  Use the default installation folder for the toolkit."
      append msg "\n\nIf there are problems with this procedure, contact [lindex $contact 0] ([lindex $contact 1])."
      set choice [tk_messageBox -type ok -message $msg -icon info -title "Install IFCsvr"]
      outputMsg "\nWait for the installation to finish before processing a STEP file." red
    } elseif {![info exists buttons]} {
      outputMsg "\nRerun this program after the installation has finished to process a STEP file."
    }

# reinstall
  } else {
    errorMsg "The existing IFCsvr toolkit must be reinstalled to update the STEP schemas."
    outputMsg "- First REMOVE the current installation of the IFCsvr toolkit."
    outputMsg "    In the IFCsvr Setup Wizard select 'REMOVE IFCsvrR300 ActiveX Component' and Finish" red
    outputMsg "    If the REMOVE was not successful, then manually uninstall the 'IFCsvrR300 ActiveX Component'"
    if {[info exists buttons]} {
      outputMsg "- Then restart this software or process a STEP file to install the updated IFCsvr toolkit."
    } else {
      outputMsg "- Then run this software again to install the updated IFCsvr toolkit."
    }
    outputMsg "- If there are problems with this procedure, contact [lindex $contact 0] ([lindex $contact 1])."

    if {[file exists $ifcsvrInst] && [info exists buttons]} {
      set msg "The IFCsvr toolkit must be reinstalled to update the STEP schemas."
      append msg "\n\nFirst REMOVE the current installation of the IFCsvr toolkit."
      append msg "\n\nIn the IFCsvr Setup Wizard (after clicking OK) select 'REMOVE IFCsvrR300 ActiveX Component' and Finish.  If the REMOVE was not successful, then manually uninstall the 'IFCsvrR300 ActiveX Component'"
      append msg "\n\nThen restart this software or process a STEP file to install the updated IFCsvr toolkit."
      append msg "\n\nIf there are problems with this procedure, contact [lindex $contact 0] ([lindex $contact 1])."
      set choice [tk_messageBox -type ok -message $msg -icon warning -title "Reinstall IFCsvr"]
      outputMsg "\nWait for the REMOVE process to finish, then restart this software or process a STEP file to install the updated IFCsvr toolkit." red
    }
  }

# try copying installation file to several locations
  set ifcsvrMsi [file join $mytemp $ifcsvr]
  if {[file exists $ifcsvrInst]} {
    if {[catch {
      file copy -force -- $ifcsvrInst $ifcsvrMsi
    } emsg1]} {
      set ifcsvrMsi [file join $mydocs $ifcsvr]
      if {[catch {
        file copy -force -- $ifcsvrInst $ifcsvrMsi
      } emsg2]} {
        set ifcsvrMsi [file join [pwd] $ifcsvr]
        if {[catch {
          file copy -force -- $ifcsvrInst $ifcsvrMsi
        } emsg3]} {
          errorMsg "ERROR copying the IFCsvr toolkit installation file to a directory."
          outputMsg " $emsg1\n $emsg2\n $emsg3"
        }
      }
    }
  }
  
# delete old installer
  catch {file delete -force -- [file join $mytemp ifcsvrr300_setup_1008_en.msi]}

# ready or not to install
  if {[file exists $ifcsvrMsi]} {
    if {[catch {
      exec {*}[auto_execok start] "" $ifcsvrMsi
      set upgradeIFCsvr [clock seconds]
      saveState 0
      if {$exit} {exit}
    } emsg]} {
      if {[string first "UNC" $emsg] != -1} {set emsg [fixErrorMsg $emsg]}
      if {$emsg != ""} {errorMsg "ERROR installing IFCsvr toolkit: $emsg"}
    }

# cannot find the toolkit
  } else {
    if {[file exists $ifcsvrInst]} {errorMsg "The IFCsvr toolkit cannot be automatically installed."}
    catch {.tnb select .tnb.status}
    update idletasks

# manual install instructions
    if {$nistVersion} {
      outputMsg "To manually install the IFCsvr toolkit:
- The installation file ifcsvrr300_setup_1008_en-update.msi can be found in $mytemp
- Run the installer and follow the instructions.  Use the default installation folder for IFCsvr.
  You might need administrator privileges (Run as administrator) to install the toolkit.
- If there are problems with the IFCsvr installation, contact [lindex $contact 0] ([lindex $contact 1])\n"
      after 1000
      errorMsg "Opening folder: $mytemp"
      if {[catch {
        exec {*}[auto_execok start] [file nativename $mytemp]
        if {$exit} {exit}
      } emsg]} {
        if {[string first "UNC" $emsg] != -1} {set emsg [fixErrorMsg $emsg]}
        if {$emsg != ""} {errorMsg "ERROR opening directory: $emsg"}
      }
    } else {
      outputMsg "To install the IFCsvr toolkit you must run the NIST version of the STEP File Analyzer and Viewer."
      outputMsg "1 - Go to https://concrete.nist.gov/cgi-bin/ctv/sfa_request.cgi"
      outputMsg "2 - Fill out the form, submit it, and follow the instructions."
      outputMsg "3 - The IFCsvr toolkit will be installed when the NIST STEP File Analyzer and Viewer is run."
      after 1000
      openURL https://concrete.nist.gov/cgi-bin/ctv/sfa_request.cgi
    }
  }
}

#-------------------------------------------------------------------------------
# shortcuts
proc setShortcuts {} {
  global mydesk mymenu mytemp nistVersion tcl_platform wdir
  
  set progname [info nameofexecutable]
  if {[string first "AppData/Local/Temp" $progname] != -1 || [string first ".zip" $progname] != -1} {
    errorMsg "For the STEP File Analyzer and Viewer to run properly, it is recommended that you first\n extract all of the files from the ZIP file and run the extracted executable."
    return
  }

  set progstr  "STEP File Analyzer and Viewer"
  if {!$nistVersion} {set progstr "SFA"}
  
  if {[info exists mydesk] || [info exists mymenu]} {
    set ok 1
    set app ""
    
# delete old shortcuts    
    set progstr1 "STEP File Analyzer"
    foreach scut [list "Shortcut to $progstr1.exe.lnk" "$progstr1.exe.lnk" "$progstr1.lnk"] {
      catch {
        if {[file exists [file join $mydesk $scut]]} {
          file delete [file join $mydesk $scut]
          set ok 0
        }
        if {[file exists [file join $mymenu "$progstr1.lnk"]]} {file delete [file join $mymenu "$progstr1.lnk"]}
      }
    }
    if {[file exists [file join $mydesk [file tail [info nameofexecutable]]]]} {set ok 0}

    set msg "Do you want to create or overwrite shortcuts to the $progstr (v[getVersion])"
    if {[info exists mydesk]} {
      append msg " on the Desktop"
      if {[info exists mymenu]} {append msg " and"}
    }
    if {[info exists mymenu]} {append msg " in the Start Menu"}
    append msg "?"
      
    if {[info exists mydesk] || [info exists mymenu]} {
      set choice [tk_messageBox -type yesno -icon question -title "Shortcuts" -message $msg]
      if {$choice == "yes"} {
        outputMsg " "
        if {$nistVersion} {catch {[file copy -force -- [file join $wdir images NIST.ico] [file join $mytemp NIST.ico]]}}
        catch {
          if {[info exists mymenu]} {
            if {[file exists [file join $mymenu "$progstr.lnk"]]} {outputMsg "Existing Start Menu shortcut will be overwritten" red}
            if {$nistVersion} {
              if {$tcl_platform(osVersion) >= 6.2} {
                twapi::write_shortcut [file join $mymenu "$progstr.lnk"] -path [info nameofexecutable] -desc $progstr -iconpath [info nameofexecutable]
              } else {
                twapi::write_shortcut [file join $mymenu "$progstr.lnk"] -path [info nameofexecutable] -desc $progstr -iconpath [file join $mytemp NIST.ico]
              }
            } else {
              twapi::write_shortcut [file join $mymenu "$progstr.lnk"] -path [info nameofexecutable] -desc $progstr
            }
            outputMsg " Shortcut created in Start Menu to [truncFileName [file nativename [info nameofexecutable]]]"
          }
        }
  
        if {$ok} {
          catch {
            if {[info exists mydesk]} {
              if {[file exists [file join $mydesk "$progstr.lnk"]]} {outputMsg "Existing Desktop shortcut will be overwritten" red}
              if {$nistVersion} {
                if {$tcl_platform(osVersion) >= 6.2} {
                  twapi::write_shortcut [file join $mydesk "$progstr.lnk"] -path [info nameofexecutable] -desc $progstr -iconpath [info nameofexecutable]
                } else {
                  twapi::write_shortcut [file join $mydesk "$progstr.lnk"] -path [info nameofexecutable] -desc $progstr -iconpath [file join $mytemp NIST.ico]
                }
              } else {
                twapi::write_shortcut [file join $mydesk "$progstr.lnk"] -path [info nameofexecutable] -desc $progstr
              }
              outputMsg " Shortcut created on Desktop to [truncFileName [file nativename [info nameofexecutable]]]"
            }
          }
        }
      }
    }
  }
}

#-------------------------------------------------------------------------------
# set home, docs, desktop, menu directories
proc setHomeDir {} {
  global drive env mydesk mydocs myhome mymenu mytemp tcl_platform

  set drive "C:/"
  if {[info exists env(SystemDrive)]} {
    set drive $env(SystemDrive)
    append drive "/"
  }
  set myhome $drive

# set mydocs, mydesk, mymenu based on USERPROFILE and registry entries
  if {[info exists env(USERPROFILE)]} {
    set myhome $env(USERPROFILE)
    
    catch {
      set reg_personal [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Personal}]
      if {[string first "%USERPROFILE%" $reg_personal] == 0} {regsub "%USERPROFILE%" $reg_personal $env(USERPROFILE) mydocs}
    }
    catch {
      set reg_desktop  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]
      if {[string first "%USERPROFILE%" $reg_desktop] == 0} {regsub "%USERPROFILE%" $reg_desktop $env(USERPROFILE) mydesk}
    }
    catch {
      set reg_menu [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Programs}]
      if {[string first "%USERPROFILE%" $reg_menu] == 0} {regsub "%USERPROFILE%" $reg_menu $env(USERPROFILE) mymenu}
    }
    
# set mytemp
    catch {
      if {$tcl_platform(osVersion) >= 6.0} {
        set reg_temp [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local AppData}]
      } else {
        set reg_temp [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local Settings}]
      }
      if {[string first "%USERPROFILE%" $reg_temp] == 0} {regsub "%USERPROFILE%" $reg_temp $env(USERPROFILE) mytemp}
      set mytemp [file join $mytemp Temp]

# make mytemp dir
      set mytemp [file nativename [file join $mytemp SFA]]
      checkTempDir
    }

# create myhome if USERPROFILE does not exist 
  } elseif {[info exists env(USERNAME)]} {
    set myhome [file join $drive Users $env(USERNAME)]
    if {$tcl_platform(osVersion) < 6.0} {set myhome [file join $drive "Documents and Settings" $env(USERNAME)]}
  }

  if {![info exists mydocs]} {
    set mydocs $myhome
    set docs "Documents"
    if {$tcl_platform(osVersion) < 6.0} {set docs "My Documents"}
    set docs [file join $mydocs $docs]
    if {[file exists $docs]} {if {[file isdirectory $docs]} {set mydocs $docs}}
  }

  if {![info exists mydesk]} {
    set mydesk $myhome
    set desk "Desktop"
    set desk [file join $mydesk $desk]
    if {[file exists $desk]} {if {[file isdirectory $desk]} {set mydesk $desk}}
  }
  
  if {![info exists mytemp]} {
    set mytemp $myhome
    set temp [file join AppData Local Temp SFA]
    if {$tcl_platform(osVersion) < 6.0} {set temp [file join "Local Settings" Temp SFA]}
    set mytemp [file join $mytemp $temp]
    checkTempDir
  }

  set myhome [file nativename $myhome]
  set mydocs [file nativename $mydocs]
  set mydesk [file nativename $mydesk]
  set mytemp [file nativename $mytemp]
}

#-------------------------------------------------------------------------------
# check if temporary directory 'mytemp' exists
proc checkTempDir {} {
  global mytemp

  if {[info exists mytemp]} {
    if {[file isfile $mytemp]} {file delete -force -- $mytemp}
    if {![file exists $mytemp]} {file mkdir $mytemp}
  } else {
    errorMsg "Temporary directory 'mytemp' does not exist."
  }
}

#-------------------------------------------------------------------------------
proc fixTimeStamp {ts} {
  set c1 [string last "+" $ts]
  if {$c1 != -1} {set ts [string range $ts 0 $c1-1]}
  set c1 [string last "-" $ts]
  if {$c1 > 8} {set ts [string range $ts 0 $c1-1]}
  set c1 [string first ":" $ts]
  set c2 [string last  ":" $ts]
  if {$c1 != $c2 && $c2 != -1} {set ts [string range $ts 0 $c2-1]}
  if {[string index $ts end] == "T"} {set ts [string range $ts 0 end-1]}
  return $ts
}

#-------------------------------------------------------------------------------
proc getTiming {{str ""}} {
  global tlast

  set t [clock clicks -milliseconds]
  if {[info exists tlast]} {outputMsg "Timing: [expr {$t-$tlast}]  $str" red}
  set tlast $t
}

#-------------------------------------------------------------------------------
# From http://wiki.tcl.tk/4021
proc sortlength2 {wordlist} {
  set words {}
  foreach word $wordlist {
    lappend words [list [string length $word] $word]
  }
  set result {}
  foreach pair [lsort -decreasing -integer -index 0 [lsort -ascii -index 1 $words]] {
    lappend result [lindex $pair 1]
  }
  return $result
}

#-------------------------------------------------------------------------------
# Based on http://www.posoft.de/html/extCawt.html
proc GetWorksheetAsMatrix {worksheetId} {
  set cellId [[$worksheetId Cells] Range [GetCellRange 1 1 [[[$worksheetId UsedRange] Rows] Count] [[[$worksheetId UsedRange] Columns] Count]]]
  set matrixList [$cellId Value2]
  return $matrixList
}

proc GetCellRange {row1 col1 row2 col2} {
  set range [format "%s%d:%s%d" [ColumnIntToChar $col1] $row1 [ColumnIntToChar $col2] $row2]
  return $range
}

proc ColumnIntToChar {col} {
  if {$col <= 0} {errorMsg "Column number $col is invalid."}
  set dividend $col
  set columnName ""
  while {$dividend > 0} {
    set modulo [expr {($dividend - 1) % 26} ]
    set columnName [format "%c${columnName}" [expr {65 + $modulo}]]
    set dividend [expr {($dividend - $modulo) / 26}]
  }
  return $columnName
}

#-------------------------------------------------------------------------------
proc compareLists {str l1 l2} {
  set l3 [intersect3 $l1 $l2]
  outputMsg "\n$str" red
  outputMsg "Unique to L1   ([llength [lindex $l3 0]])\n  [lindex $l3 0]"
  outputMsg "Common to both ([llength [lindex $l3 1]])\n  [lindex $l3 1]"
  outputMsg "Unique to L2   ([llength [lindex $l3 2]])\n  [lindex $l3 2]"
}

#-------------------------------------------------------------------------------
# dot - calculate scalar dot product of two vectors
proc vecdot {v1 v2} {
  set v3 0.0
  foreach c1 $v1 c2 $v2 {set v3 [expr {$v3+$c1*$c2}]}
  return $v3
}

# mult - multiply vector by scalar
proc vecmult {v1 scalar} {
  foreach c1 $v1 {lappend v2 [expr {$c1*$scalar}]}
  return $v2
}

# sub - subtract one vector from another
proc vecsub {v1 v2} {
  foreach c1 $v1 c2 $v2 {lappend v3 [expr {$c1-$c2}]}
  return $v3
}

# add - add one vector to another
proc vecadd {v1 v2} {
  foreach c1 $v1 c2 $v2 {lappend v3 [expr {$c1+$c2}]}
  return $v3
}

# reverse - reverse vector direction
proc vecrev {v1} {
  foreach c1 $v1 {
    if {$c1 != 0.} {
      lappend v2 [expr {$c1*-1.}]
    } else {
      lappend v2 $c1
    }
  }
  return $v2
}

# trim - truncate values in a vector
proc vectrim {v1} {
  foreach c1 $v1 {
    set prec 3
    if {[expr {abs($c1)}] < 0.01} {set prec 4}
    lappend v2 [trimNum $c1 $prec]
  }
  return $v2
}

# cross - cross product between two 3d-vectors
proc veccross {v1 v2} {
  set v1x [lindex $v1 0]
  set v1y [lindex $v1 1]
  set v1z [lindex $v1 2]
  set v2x [lindex $v2 0]
  set v2y [lindex $v2 1]
  set v2z [lindex $v2 2]
  set v3 [list [expr {$v1y*$v2z-$v1z*$v2y}] [expr {$v1z*$v2x-$v1x*$v2z}] [expr {$v1x*$v2y-$v1y*$v2x}]]
  return $v3
}

# len - get scalar length of a vector
proc veclen {v1} {
 set l 0.
 foreach c1 $v1 {set l [expr {$l + $c1*$c1}]}
 return [expr {sqrt($l)}]
}

# norm - normalize a vector
proc vecnorm {v1} {
  set l [veclen $v1]
  if {$l != 0.} {
    set s [expr {1./$l}]
    foreach c1 $v1 {lappend v2 [expr {$c1*$s}]}
  } else {
    set v2 $v1
  }
  return $v2
}

# angle - angle between two vectors
proc vecangle {v1 v2} {
  set angle [trimNum [expr {acos([vecdot $v1 $v2] / ([veclen $v1]*[veclen $v2]))}]]
  return $angle
}
