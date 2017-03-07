#-------------------------------------------------------------------------------
proc checkValues {} {
  global opt buttons appNames appName programfiles userEntityList
  global edmWriteToFile edmWhereRules eeWriteToFile

  if {[info exists buttons(appCombo)]} {
    set ic [lsearch $appNames $appName]
    if {$ic < 0} {set ic 0}
    $buttons(appCombo) current $ic

# Jotne EDM Model Checker
    catch {
      if {[string first "EDM Model Checker" $appName] == 0} {
        pack $buttons(edmWriteToFile) -side left -anchor w -padx 5
        pack $buttons(edmWhereRules) -side left -anchor w -padx 5
      } else {
        pack forget $buttons(edmWriteToFile)
        pack forget $buttons(edmWhereRules)
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
      if {$appName == "Indent STEP File (for debugging)"} {
        pack $buttons(indentGeometry) -side left -anchor w -padx 5
        pack $buttons(indentStyledItem) -side left -anchor w -padx 5
      } else {
        pack forget $buttons(indentGeometry)
        pack forget $buttons(indentStyledItem)
      }
    }
  }

  if {$opt(XLSCSV) == "CSV"} {
    set opt(INVERSE) 0
    set opt(PMIGRF)  0
    set opt(PMISEM)  0
    set opt(VALPROP) 0
    set opt(writeDirType) 0
    set opt(XL_OPEN) 1
    $buttons(genExcel)   configure -text "Generate CSV Files"
    $buttons(optINVERSE) configure -state disabled
    $buttons(optPMIGRF)  configure -state disabled
    $buttons(optPMISEM)  configure -state disabled
    $buttons(optVALPROP) configure -state disabled
    $buttons(optXL_FPREC)    configure -state disabled
    $buttons(optXL_KEEPOPEN) configure -state disabled
    $buttons(optXL_LINK1)    configure -state disabled
    $buttons(optXL_SORT)     configure -state disabled
  } else {
    $buttons(genExcel)   configure -text "Generate Spreadsheet"
    $buttons(optINVERSE) configure -state normal
    $buttons(optPMIGRF)  configure -state normal
    $buttons(optPMISEM)  configure -state normal
    $buttons(optVALPROP) configure -state normal
    $buttons(optXL_FPREC)    configure -state normal
    $buttons(optXL_KEEPOPEN) configure -state normal
    $buttons(optXL_LINK1)    configure -state normal
    $buttons(optXL_SORT)     configure -state normal
  }
  
# STEP related
  if {$opt(PMIGRF)} {
    set opt(PR_STEP_AP242) 1
    set opt(PR_STEP_COMM) 1
    set opt(PR_STEP_PRES) 1
    set opt(PR_STEP_QUAN) 1
    set opt(PR_STEP_SHAP) 1
    $buttons(optVIZPMI) configure -state normal
    $buttons(optPR_STEP_AP242) configure -state disabled
    $buttons(optPR_STEP_COMM) configure -state disabled
    $buttons(optPR_STEP_PRES) configure -state disabled
    $buttons(optPR_STEP_QUAN) configure -state disabled
    $buttons(optPR_STEP_SHAP) configure -state disabled
  } else {
    set opt(VIZPMI) 0
    $buttons(optVIZPMI) configure -state disabled
    $buttons(optPR_STEP_PRES) configure -state normal
    if {!$opt(VALPROP)} {$buttons(optPR_STEP_QUAN) configure -state normal}
    if {!$opt(PMISEM)}  {
      $buttons(optPR_STEP_AP242) configure -state normal
      $buttons(optPR_STEP_COMM) configure -state normal
      $buttons(optPR_STEP_SHAP) configure -state normal
    }
  }
  if {$opt(VALPROP)} {
    set opt(PR_STEP_QUAN) 1
    $buttons(optPR_STEP_QUAN) configure -state disabled
  } else {
    if {!$opt(PMIGRF)} {$buttons(optPR_STEP_QUAN) configure -state normal}
  }
  if {$opt(VIZPMI)} {
    $buttons(gpmiColor0) configure -state normal
    $buttons(gpmiColor1) configure -state normal
    $buttons(gpmiColor2) configure -state normal
    $buttons(linecolor)  configure -state normal
  } else {
    $buttons(gpmiColor0) configure -state disabled
    $buttons(gpmiColor1) configure -state disabled
    $buttons(gpmiColor2) configure -state disabled
    $buttons(linecolor)  configure -state disabled
  }
  if {$opt(PMISEM)} {
    set opt(PR_STEP_AP242) 1
    set opt(PR_STEP_COMM) 1
    set opt(PR_STEP_QUAN) 1
    set opt(PR_STEP_REPR) 1
    set opt(PR_STEP_SHAP) 1
    set opt(PR_STEP_TOLR) 1
    $buttons(optPR_STEP_AP242) configure -state disabled
    $buttons(optPR_STEP_COMM) configure -state disabled
    $buttons(optPR_STEP_QUAN) configure -state disabled
    $buttons(optPR_STEP_REPR) configure -state disabled
    $buttons(optPR_STEP_SHAP) configure -state disabled
    $buttons(optPR_STEP_TOLR) configure -state disabled
    catch {$buttons(optDEBUG2) configure -state normal}
  } else {
    $buttons(optPR_STEP_REPR) configure -state normal
    $buttons(optPR_STEP_TOLR) configure -state normal
    catch {
      set opt(DEBUG2) 0
      $buttons(optDEBUG2) configure -state disabled
    }
    if {!$opt(PMIGRF)} {
      if {!$opt(VALPROP)} {$buttons(optPR_STEP_QUAN) configure -state normal}
      $buttons(optPR_STEP_AP242) configure -state normal
      $buttons(optPR_STEP_COMM) configure -state normal
      $buttons(optPR_STEP_SHAP) configure -state normal
    }
  }
  
# user-defined entity list
  if {[info exists opt(PR_USER)]} {
    if {$opt(PR_USER)} {
      $buttons(userentity)     configure -state normal
      $buttons(userentityopen) configure -state normal
    } else {
      $buttons(userentity)     configure -state disabled
      $buttons(userentityopen) configure -state disabled
      set userEntityList {}
    }
  }
  
  if {$opt(writeDirType) == 0} {
    $buttons(userdir)    configure -state disabled
    $buttons(userentry)  configure -state disabled
    $buttons(userentry1) configure -state disabled
    $buttons(userfile)   configure -state disabled
  } elseif {$opt(writeDirType) == 1} {
    $buttons(userdir)    configure -state disabled
    $buttons(userentry)  configure -state disabled
    $buttons(userentry1) configure -state normal
    $buttons(userfile)   configure -state normal
  } elseif {$opt(writeDirType) == 2} {
    $buttons(userdir)    configure -state normal
    $buttons(userentry)  configure -state normal
    $buttons(userentry1) configure -state disabled
    $buttons(userfile)   configure -state disabled
  }

# make sure there is some entity type to process
  set nopt 0
  foreach idx [lsort [array names opt]] {
    if {([string first "PR_" $idx] == 0 || $idx == "VALPROP" || $idx == "PMIGRF" || $idx == "PMISEM") && [string first "FEAT" $idx] == -1} {
      incr nopt $opt($idx)
    }
  }
  if {$nopt == 0} {
    set opt(PR_STEP_AP242) 1
    set opt(PR_STEP_COMM) 1
    set opt(PR_STEP_PRES) 1
    set opt(PR_STEP_QUAN) 1
    set opt(PR_STEP_REPR) 1
    set opt(PR_STEP_SHAP) 1
    set opt(PR_STEP_TOLR) 1
  }
}

# -------------------------------------------------------------------------------------------------
# set color based on entColorIndex variable
proc setColorIndex {ent {multi 0}} {
  global entCategory entColorIndex stepAP
  
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

    #outputMsg "TC $tc"
    if {$tc < 1000} {return $tc}
  }

# entity not in any category, color by AP
  if {!$multi} {
    if {[string first "AP209" $stepAP] != -1} {return 19} 
    if {$stepAP == "AP210"} {return 15} 
    if {$stepAP == "AP238"} {return 24}
  }

# entity from other APs (no color)
  return -2      
}

#-------------------------------------------------------------------------------
proc openURL {url} {
  global programfiles

# open in whatever is registered for the file extension, except for .cgi for upgrade url
  if {[string first ".cgi" $url] == -1} {
    if {[catch {
      exec {*}[auto_execok start] "" $url
    } emsg]} {
      if {[string first "is not recognized" $emsg] == -1} {
        errorMsg "ERROR opening $url\n $emsg"
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
    if {$webCmd == "" || ![file exists $webCmd]} {set webCmd [file join $programfiles "Internet Explorer" IEXPLORE.EXE]}
    exec $webCmd $url &
  }
}

#-------------------------------------------------------------------------------
proc openFile {{openName ""}} {
  global localName localNameList outputWin fileDir buttons lastXLS extXLS

  if {$openName == ""} {
  
# file types for file select dialog (removed .stpnc)
    set typelist {{"STEP Files" {".stp" ".step" ".p21" ".stpZ" ".ifc"}}}
    lappend typelist {"All Files" {*}}

# file open dialog
    set localNameList [tk_getOpenFile -title "Open STEP File(s)" -filetypes $typelist -initialdir $fileDir -multiple true]
    if {[llength $localNameList] <= 1} {set localName [lindex $localNameList 0]}
    catch {
      set fext [string tolower [file extension $localName]]
      if {[string first ".ifc" $fext] != -1} {
        #errorMsg "Use the IFC File Analyzer with IFC files."
        #openURL https://go.usa.gov/xK9gh
      } elseif {$fext == ".stpnc"} {
        errorMsg "Rename the file extension to '.stp' to process STEP-NC files."
      }
    }

# file name passed in as openName
  } else {
    set localName $openName
    set localNameList [list $localName]
  }

# multiple files selected
  if {[llength $localNameList] > 1} {
    set fileDir [file dirname [lindex $localNameList 0]]

    outputMsg "Ready to process [llength $localNameList] STEP files" blue
    $buttons(genExcel) configure -state normal
    if {[info exists buttons(appOpen)]} {$buttons(appOpen) configure -state normal}
    focus $buttons(genExcel)

# single file selected
  } elseif {[file exists $localName]} {
  
# check for zipped file
    if {[string first ".stpz" [string tolower $localName]] != -1} {unzipFile}  

    set fileDir [file dirname $localName]
    if {[string first "z" [string tolower [file extension $localName]]] == -1} {
      outputMsg "Ready to process: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)" blue
      $buttons(genExcel) configure -state normal
      if {[info exists buttons(appOpen)]} {$buttons(appOpen) configure -state normal}
      focus $buttons(genExcel)
      #set lastXLS "[file nativename [file join [file dirname $localName] [file rootname [file tail $localName]]]]_stp.xlsx"
    }
  
# not found
  } else {
    if {$localName != ""} {errorMsg "File not found: [truncFileName [file nativename $localName]]"}
  }
  .tnb select .tnb.status
}

#-------------------------------------------------------------------------------
proc unzipFile {} {
  global localName wdir mytemp

  if {[catch {
    outputMsg " Unzipping: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)" blue

# copy gunzip to TEMP
    if {[file exists [file join $wdir schemas gunzip.exe]]} {file copy -force [file join $wdir schemas gunzip.exe] $mytemp}

    set gunzip [file join $mytemp gunzip.exe]
    if {[file exists $gunzip]} {

# copy zipped file to TEMP
      if {[regsub ".stpZ" $localName ".stp.Z" ln] == 0} {regsub ".stpz" $localName ".stp.Z" ln}
      set fzip [file join $mytemp [file tail $ln]]
      file copy -force $localName $fzip

# get name of unzipped file
      set gz [exec $gunzip -Nl $fzip]
      set c1 [string first "%" $gz]
      set ftmp [string range $gz $c1+2 end]

# unzip
      if {[file tail $ftmp] != [file tail $fzip]} {outputMsg " Extracting: [file tail $ftmp]" blue}
      exec $gunzip -Nf $fzip

# copy to new stp file
      set fstp [file join [file dirname $localName] [file tail $ftmp]]
      set ok 0
      if {[file exists $fstp]} {
        if {[file mtime $localName] != [file mtime $fstp]} {
          outputMsg " Overwriting existing STEP file: [truncFileName [file nativename $fstp]]" red
          set ok 1
        } else {
          outputMsg "  File already extracted" red
        }
      } else {
        set ok 1
      }
      if {$ok} {
        file copy -force $ftmp $fstp
        set localName $fstp
      }
      file delete $fzip
      file delete $ftmp
    } else {
      errorMsg "ERROR: gunzip not found to unzip compressed STEP file"
    }
  } emsg]} {
    errorMsg "ERROR unzipping file: $emsg"
  }
}

#-------------------------------------------------------------------------------
proc saveState {} {
  global optionsFile fileDir openFileList opt userWriteDir dispCmd dispCmds
  global lastXLS lastXLS1 userXLSFile fileDir1 mydocs sfaVersion upgrade
  global excelYear userEntityFile buttons statusFont

  if {![info exists buttons]} {return}
  
  if {[catch {
    if {![file exists $optionsFile]} {
      outputMsg " "
      errorMsg "Creating options file: [truncFileName $optionsFile]"
    }
    set fileOptions [open $optionsFile w]
    puts $fileOptions "# Options file for the STEP File Analyzer v[getVersion] ([string trim [clock format [clock seconds]]])\n#\n# DO NOT EDIT OR DELETE FROM USER HOME DIRECTORY $mydocs\n# DOING SO WILL CORRUPT THE CURRENT SETTINGS OR CAUSE ERRORS IN THE SOFTWARE\n#"
    set varlist [list fileDir fileDir1 userWriteDir userEntityFile openFileList dispCmd dispCmds lastXLS lastXLS1 \
                      userXLSFile statusFont upgrade sfaVersion excelYear]

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
    set wg [winfo geometry .]
    catch {set winpos [string range $wg [string first "+" $wg] end]}
    puts $fileOptions "set winpos \"$winpos\""
    set wingeo [string range $wg 0 [expr {[string first "+" $wg]-1}]]
    puts $fileOptions "set wingeo \"$wingeo\""

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

  } emsg]} {
    errorMsg "ERROR writing to options file: $emsg"
    catch {raise .}
  }
}

#-------------------------------------------------------------------------------
proc runOpenProgram {} {
  global localName dispCmd appName transFile
  global sccmsg model_typ pfbent
  global openFileList File editorCmd
  global edmWriteToFile edmWhereRules eeWriteToFile
  
  set dispFile $localName
  set idisp [file rootname [file tail $dispCmd]]
  if {[info exists appName]} {if {$appName != ""} {set idisp $appName}}

  .tnb select .tnb.status
  outputMsg "Opening STEP file in: $idisp"

# open file
#  (list is programs that CANNOT start up with a file *OR* need specific commands below)
  if {[string first "Conformance"       $idisp] == -1 && \
      [string first "Indent"            $idisp] == -1 && \
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
      exec {*}[auto_execok start] "" $dispFile
    } emsg]} {
      errorMsg "No application is associated with STEP files."
      errorMsg " Go to Websites > STEP File Viewers"
    }

# indent file
  } elseif {[string first "Indent" $idisp] != -1} {
    indentFile $dispFile

# QuickStep
  } elseif {[string first "QuickStep" $idisp] != -1} {
    cd [file dirname $dispFile]
    exec $dispCmd [file tail $dispFile] &

#-------------------------------------------------------------------------------
# validate file with ST-Developer Conformance Checkers
  } elseif {[string first "Conformance" $idisp] != -1} {
    set stfile $dispFile
    outputMsg "Ready to validate:  [truncFileName [file nativename $stfile]] ([expr {[file size $stfile]/1024}] Kb)" blue
    cd [file dirname $stfile]

# gui version
    if {[string first "gui" $dispCmd] != -1 && !$eeWriteToFile} {
      if {[catch {exec $dispCmd $stfile &} err]} {outputMsg "Conformance Checker error:\n $err" red}

# non-gui version
    } else {
      set stname [file tail $stfile]
      set stlog  "[file rootname $stname]\_stdev.log"
      catch {if {[file exists $stlog]} {file delete -force $stlog}}
      outputMsg "ST-Developer log file: [truncFileName [file nativename $stlog]]" blue

      set c1 [string first "gui" $dispCmd]
      set dispCmd1 $dispCmd
      if {$c1 != -1} {set dispCmd1 [string range $dispCmd 0 $c1-1][string range $dispCmd $c1+3 end]}

      if {[string first "apconform" $dispCmd1] != -1} {
        if {[catch {exec $dispCmd1 -syntax -required -unique -bounds -aggruni -arrnotopt -inverse -strwidth -binwidth -realprec -atttypes -refdom $stfile >> $stlog &} err]} {outputMsg "Conformance Checker error:\n $err" red}
      } else {
        if {[catch {exec $dispCmd1 $stfile >> $stlog &} err]} {outputMsg "Conformance Checker error:\n $err" red}
      }  
      if {[string first "TextPad" $editorCmd] != -1 || [string first "Notepad++" $editorCmd] != -1} {
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
    outputMsg "Ready to validate:  [truncFileName [file nativename $filename]] ([expr {[file size $filename]/1024}] Kb)" blue
    cd [file dirname $filename]

# write script file to open database
    set edmScript "[file rootname $filename]_edm.scr"
    set scriptFile [open $edmScript w]
    set okschema 1

    set edmDir [split [file nativename $dispCmd] [file separator]]
    set i [lsearch $edmDir "bin"]
    set edmDir [join [lrange $edmDir 0 [expr {$i-1}]] [file separator]]
    set edmDBopen "ACCUMULATING_COMMAND_OUTPUT,OPEN_SESSION"
    
# open file to find STEP schema name (can't create ap214 and ap210 because the schemas won't compile in EDMS without errors)
    set fschema [getSchemaFromFile $filename]
    if {$fschema == "CONFIG_CONTROL_DESIGN"} {
      puts $scriptFile "Database>Open([file nativename [file join $edmDir Db]], ap203, ap203, \"$edmDBopen\")"
    } elseif {[string first "AP203_CONFIGURATION_CONTROLLED_3D_DESIGN_OF_MECHANICAL_PARTS_AND_ASSEMBLIES_MIM_LF" $fschema] == 0} {
      puts $scriptFile "Database>Open([file nativename [file join $edmDir Db]], ap203e2, ap203e2, \"$edmDBopen\")"
    } elseif {[string first "AP242_MANAGED_MODEL_BASED_3D_ENGINEERING_MIM_LF" $fschema] == 0} {
      puts $scriptFile "Database>Open([file nativename [file join $edmDir Db]], ap242, ap242, \"$edmDBopen\")"
    } elseif {[string first "AM_AP242_MANAGED_MODEL_BASED_3D_ENGINEERING_MIM_LF" $fschema] == 0} {
      puts $scriptFile "Database>Open([file nativename [file join $edmDir Db]], ap242_am, ap242_am, \"$edmDBopen\")"
    } elseif {[string first "AP209_MULTIDISCIPLINARY_ANALYSIS_AND_DESIGN_MIM_LF" $fschema] == 0} {
      puts $scriptFile "Database>Open([file nativename [file join $edmDir Db]], ap209, ap209, \"$edmDBopen\")"
    } else {
      outputMsg "EDM Model Checker cannot be used with:\n $fschema" red
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
        regsub -all {[\.()]} $fileRoot "_" fileRoot
        set edmFile [file join [file dirname $filename] $fileRoot]
        append edmFile [file extension $filename]
        file copy -force $filename $edmFile
        set tmpfile 1
      } else {
        set edmFile $filename
      }

# validate everything
      #set validate "FULL_VALIDATION,OUTPUT_STEPID"

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
      outputMsg "Running EDM Model Checker"
      eval exec {$dispCmd} $edmScript &

# if results are written to a file, open output file from the validation (edmLog) and output file if there are import errors (edmLogImport)
      if {$edmWriteToFile} {
        if {[string first "TextPad" $editorCmd] != -1 || [string first "Notepad++" $editorCmd] != -1} {
          outputMsg "Opening log file(s) in editor"
          exec $editorCmd $edmLog &
          after 1000
          if {[file size $edmLogImport] > 0} {
            exec $editorCmd $edmLogImport &
          } else {
            catch {file delete -force $edmLogImport}
          }
        } else {
          outputMsg "Wait until the EDM Model Checker has finished and then open the log file"
        }
      }

# attempt to delete the script file
      set nerr 0
      while {[file exists $edmScript]} {
        after 1000
        incr nerr
        catch {file delete $edmScript}
        if {$nerr > 60} {break}
      }

# if using a temporary file, attempt to delete it
      if {$tmpfile} {
        set nerr 0
        while {[file exists $edmFile]} {
          after 1000
          incr nerr
          catch {file delete $edmFile}
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
proc getOpenPrograms {} {
  global dispApps dispCmds dispCmd appNames appName env programfiles pf64
  global drive edmexe editorCmd developer myhome

# Including any of the CAD viewers and checkers below does not imply a recommendation or endorsement of them by NIST https://www.nist.gov/disclaimer
# For more STEP viewers, go to https://www.cax-if.org/step_viewers.html
  
  set pflist {}
  set pf [string range $programfiles 3 end]
  foreach drives {C D E F G H I J K L M N O P Q R S T U V W X Y Z} {
    set pf1 $drives
    append pf1 ":/$pf"
    if {[file isdirectory $pf1]} {lappend pflist $pf1}
    if {$pf64 != ""} {
      set pf1 $drives
      append pf1 ":/[string range $pf64 3 end]"
      if {[file isdirectory $pf1]} {lappend pflist $pf1}
    }
  }
  set lastver 0

# Jotne EDM Model Checker
  if {$developer} {
    set edmexe  ""
    set edms [glob -nocomplain -directory [file join $drive edm] -join edm* bin Edms.exe]
    foreach match $edms {
      #set vernum [string range [lindex [split $match "/"] 2] 3 11]
      #set name "EDM Model Checker $vernum"
      set name "EDM Model Checker"
      set edmexe $match
      set dispApps($edmexe) $name
    }
  }

  if {[info exists programfiles]} {
    foreach pf $pflist {

# STEP Tools apps
      if {[file isdirectory [file join $pf "STEP Tools"]]} {
        set applist [list \
          [list stepbrws.exe "STEP File Browser"] \
          [list stview.exe "STEP Viewer"] \
          [list stpcheckgui.exe "STEP Check and Browse"] \
          [list stepcleangui.exe "STEP File Cleaner"] \
          [list ap203checkgui.exe "STEP AP203 Conformance Checker"] \
          [list ap209checkgui.exe "STEP AP209 Conformance Checker"] \
          [list ap214checkgui.exe "STEP AP214 Conformance Checker"] \
          [list apconformgui.exe "STEP AP Conformance Checker"] \
        ]
        foreach app $applist {
          set stmatch ""
          foreach match [glob -nocomplain -directory $pf -join "STEP Tools" "ST-Developer*" bin [lindex $app 0]] {
            if {$stmatch == ""} {
              set stmatch $match
              set lastver [lindex [split [file nativename $match] [file separator]] 3]
            } else {
              set ver [lindex [split [file nativename $match] [file separator]] 3]
              if {$ver > $lastver} {set stmatch $match}
            }
          }
          if {$stmatch != ""} {
            if {![info exists dispApps($stmatch)]} {
              #set vn [lindex [lindex [split [file nativename $stmatch] [file separator]] 3] 1]
              set dispApps($stmatch) [lindex $app 1]
            }
          }
        }
      }

# other STEP file apps
      set applist [list \
        [list {*}[glob -nocomplain -directory [file join $pf Materialise] -join "Magics *" Magics.exe] Magics] \
        [list {*}[glob -nocomplain -directory [file join $pf Actify SpinFire] -join "*" SpinFire.exe] SpinFire] \
        [list {*}[glob -nocomplain -directory [file join $pf] -join "3D-Tool V*" 3D-Tool.exe] 3D-Tool] \
        [list {*}[glob -nocomplain -directory [file join $pf] -join "VariCADViewer *" bin varicad-x64.exe] "VariCAD Viewer"] \
        [list {*}[glob -nocomplain -directory [file join $pf "Stratasys Direct Manufacturing"] -join "SolidView Pro RP *" bin SldView.exe] SolidView] \
        [list {*}[glob -nocomplain -directory [file join $pf "TransMagic Inc"] -join "TransMagic *" System code bin TransMagic.exe] TransMagic] \
        [list {*}[glob -nocomplain -directory [file join $pf CADSoftTools] -join "ABViewer*" ABViewer.exe] ABViewer] \
        [list {*}[glob -nocomplain -directory [file join $pf "Soft Gold"] -join "ABViewer*" ABViewer.exe] ABViewer] \
        [list {*}[glob -nocomplain -directory [file join $pf Autodesk] -join "Navisworks Manage*" roamer.exe] "Navisworks Manage"] \
      ]
      foreach app $applist {
        if {[llength $app] == 2} {
          set match [join [lindex $app 0]]
          if {$match != "" && ![info exists dispApps($match)]} {
            set dispApps($match) [lindex $app 1]
          }
        }
      }
      
      set applist [list \
        [list [file join $pf "STEP Tools" "STEP-NC Machine" STEPNCExplorer.exe] "STEP-NC Machine"] \
        [list [file join $pf "STEP Tools" "STEP-NC Machine" STEPNCExplorer_x86.exe] "STEP-NC Machine"] \
        [list [file join $pf "STEP Tools" "STEP-NC Machine Personal Edition" STEPNCExplorer.exe] "STEP-NC Machine"] \
        [list [file join $pf "STEP Tools" "STEP-NC Machine Personal Edition" STEPNCExplorer_x86.exe] "STEP-NC Machine"] \
        [list [file join $pf Glovius Glovius glovius.exe] Glovius] \
        [list [file join $pf "CAD Exchanger" bin Exchanger.exe] "CAD Exchanger"] \
        [list [file join $pf "CAD Assistant" CADAssistant.exe] "CAD Assistant"] \
        [list [file join $pf "3DJuump X64" 3DJuump.exe] "3DJuump"] \
        [list [file join $pf STPViewer STPViewer.exe] "STP Viewer"] \
        [list [file join $pf CadFaster QuickStep QuickStep.exe] QuickStep] \
        [list [file join $pf Kisters 3DViewStation 3DViewStation.exe] 3DViewStation] \
        [list [file join $pf IFCBrowser IfcQuickBrowser.exe] IfcQuickBrowser] \
        [list [file join $pf "Tekla BIMsight" BIMsight.exe] "Tekla BIMsight"] \
      ]
      foreach app $applist {
        if {[file exists [lindex $app 0]]} {
          set name [lindex $app 1]
          set dispApps([lindex $app 0]) $name
        }
      }

# Tetra4D in Adobe Acrobat
      for {set i 40} {$i > 9} {incr i -1} {
        if {$i > 11} {
          set j "20$i"
        } else {
          set j "$i.0"
        }        
        foreach match [glob -nocomplain -directory $pf -join Adobe "Acrobat $j" Acrobat Acrobat.exe] {
          if {[file exists [file join $pf Adobe "Acrobat $j" Acrobat plug_ins 3DPDFConverter 3DPDFConverter.exe]]} {
            if {![info exists dispApps($match)]} {
              set name "Tetra4D Converter"
              set dispApps($match) $name
            }
          }
        }
        set match [file join $pf Adobe "Acrobat $j" Acrobat plug_ins 3DPDFConverter 3DReviewer.exe]
        if {![info exists dispApps($match)]} {
          set name "Tetra4D Reviewer"
          set dispApps($match) $name
        }
      }
    }
  }

# IDA-STEP  
  set b1 [file join $myhome AppData Local IDA-STEP ida-step.exe]
  if {[file exists $b1]} {
    set name "IDA-STEP Viewer"
    set dispApps($b1) $name
  }

#-------------------------------------------------------------------------------
# default viewer
  set dispApps(Default) "Default STEP Viewer"

# file indenter
  set dispApps(Indent) "Indent STEP File (for debugging)"

#-------------------------------------------------------------------------------
# set text editor command and name
  set editorCmd ""
  set editorName ""

# Notepad
  if {[info exists env(windir)]} {
    set editorCmd [file join $env(windir) Notepad.exe]
    set editorName "Notepad"
    set dispApps($editorCmd) $editorName
    if {![file exists $editorCmd]} {
      set editorCmd [file join $env(windir) system32 Notepad.exe]
      set editorName "Notepad"
      set dispApps($editorCmd) $editorName
    }
  }

# other text editors
  for {set i 9} {$i > 5} {incr i -1} {
    set editorCmd1 [file join $programfiles "TextPad $i" TextPad.exe]
    if {[file exists $editorCmd1]} {
      set editorName1 "TextPad $i"
      set dispApps($editorCmd1) $editorName1
      set editorCmd $editorCmd1
      set editorName $editorName1
    }
  }
  set editorCmd1 [file join $programfiles Notepad++ notepad++.exe]
  if {[file exists $editorCmd1]} {
    set editorName1 "Notepad++"
    set dispApps($editorCmd1) $editorName1
    set editorCmd $editorCmd1
    set editorName $editorName1
  }

#-------------------------------------------------------------------------------
# remove cmd that do not exist in dispCmds and non-executables
  set dispCmds1 {}
  foreach app $dispCmds {
    if {([file exists $app] || [string first "Default" $app] == 0 || [string first "Indent" $app] == 0) && \
         [file tail $app] != "NotePad.exe"} {
      lappend dispCmds1 $app
    }
  }
  set dispCmds $dispCmds1

# check for cmd in dispApps that does not exist in dispCmds and add to list
  foreach app [array names dispApps] {
    if {[file exists $app] || [string first "Default" $app] == 0 || [string first "Indent" $app] == 0} {
      set notInCmds 1
      foreach cmd $dispCmds {if {[string tolower $cmd] == [string tolower $app]} {set notInCmds 0}}
      if {$notInCmds} {lappend dispCmds $app}
    }
  }

# remove duplicates in dispCmds
  if {[llength $dispCmds] != [llength [lrmdups $dispCmds]]} {set dispCmds [lrmdups $dispCmds]}

# clean up list of app viewer commands
  if {[info exists dispCmd]} {
    if {([file exists $dispCmd] || [string first "Default" $dispCmd] == 0 || [string first "Indent" $dispCmd] == 0)} {
      if {[lsearch $dispCmds $dispCmd] == -1 && $dispCmd != ""} {lappend dispCmds $dispCmd}
    } else {
      if {[llength $dispCmds] > 0} {
        foreach dispCmd $dispCmds {
          if {([file exists $dispCmd] || [string first "Default" $dispCmd] == 0 || [string first "Indent" $dispCmd] == 0)} {break}
        }
      } else {
        set dispCmd ""
      }
    }
  } else {
    if {[llength $dispCmds] > 0} {
      set dispCmd [lindex $dispCmds 0]
    }
  }
  for {set i 0} {$i < [llength $dispCmds]} {incr i} {
    if {![file exists [lindex $dispCmds $i]] && [string first "Default" [lindex $dispCmds $i]] == -1 && [string first "Indent" [lindex $dispCmds $i]] == -1} {set dispCmds [lreplace $dispCmds $i $i]}
  }

# put dispCmd at beginning of dispCmds list
  if {[info exists dispCmd]} {
    for {set i 0} {$i < [llength $dispCmds]} {incr i} {
      if {$dispCmd == [lindex $dispCmds $i]} {
        set dispCmds [lreplace $dispCmds $i $i]
        set dispCmds [linsert $dispCmds 0 $dispCmd]
      }
    }
  }

# remove duplicates in dispCmds, again
  if {[llength $dispCmds] != [llength [lrmdups $dispCmds]]} {set dispCmds [lrmdups $dispCmds]}

# set list of STEP viewer names, appNames
  set appNames {}
  set appName  ""
  foreach cmd $dispCmds {
    if {[info exists dispApps($cmd)]} {
      lappend appNames $dispApps($cmd)
    } else {
      set name [file rootname [file tail $cmd]]
      if {$name == "Edms"} {set name "EDM Model Checker"}

      lappend appNames  $name
      set dispApps($cmd) $name
    }
  }
  if {$dispCmd != ""} {
    if {[info exists dispApps($dispCmd)]} {set appName $dispApps($dispCmd)}
  }
}

# -------------------------------------------------------------------------------------------------
proc getFirstFile {} {
  global openFileList remoteName buttons
  
  set localName [lindex $openFileList 0]
  if {$localName != ""} {
    set remoteName $localName
    outputMsg "Ready to process: [file tail $localName] ([expr {[file size $localName]/1024}] Kb)" blue
    if {[info exists buttons(appOpen)]} {$buttons(appOpen) configure -state normal}
  }
  return $localName
}

#-------------------------------------------------------------------------------
proc findFile {startDir {recurse 0}} {
  global fileList
  
  set pwd [pwd]
  if {[catch {cd $startDir} err]} {
    errorMsg $err
    return
  }

  set exts {".stp" ".step" ".p21" ".stpz"}

  foreach match [glob -nocomplain -- *] {
    foreach ext $exts {
      if {[file extension [string tolower $match]] == $ext} {
        if {$ext != ".stpz" || ![file exists [string range $match 0 end-1]]} {
          lappend fileList [file join $startDir $match]
        }
      }
    }
  }
  if {$recurse} {
    foreach file [glob -nocomplain *] {
      if {[file isdirectory $file]} {findFile [file join $startDir $file] $recurse}
    }
  }
  cd $pwd
}

#-------------------------------------------------------------------------------
proc addFileToMenu {} {
  global openFileList localName File buttons

  set lenlist 25
  set filemenuinc 4
  
  if {![info exists buttons]} {return}
  
# change backslash to forward slash, if necessary
  regsub -all {\\} $localName "/" localName

# remove duplicates
  set newlist {}
  set dellist {}
  for {set i 0} {$i < [llength $openFileList]} {incr i} {
    set name [lindex $openFileList $i]
    set ifile [lsearch -all $openFileList $name]
    if {[llength $ifile] == 1 || [lindex $ifile 0] == $i} {
      lappend newlist $name
    } else {
      lappend dellist $i
    }
  }
  set openFileList $newlist
  
# check if file name is already in the menu, if so, delete
  set ifile [lsearch $openFileList $localName]
  if {$ifile > 0} {
    set openFileList [lreplace $openFileList $ifile $ifile]
    $File delete [expr {$ifile+$filemenuinc}] [expr {$ifile+$filemenuinc}]
  }

# insert file name at top of list
  set fext [string tolower [file extension $localName]]
  if {$ifile != 0 && ($fext == ".stp" || $fext == ".step" || $fext == ".p21")} {
    set openFileList [linsert $openFileList 0 $localName]
    $File insert $filemenuinc command -label [truncFileName [file nativename $localName] 1] \
      -command [list openFile $localName] -accelerator "F1"
    catch {$File entryconfigure 5 -accelerator {}}
  }

# check length of file list, delete from the end of the list
  if {[llength $openFileList] > $lenlist} {
    set openFileList [lreplace $openFileList $lenlist $lenlist]
    $File delete [expr {$lenlist+$filemenuinc}] [expr {$lenlist+$filemenuinc}]
  }

# compare file list and menu list
  set llen [llength $openFileList]
  for {set i 0} {$i < $llen} {incr i} {
    set f1 [file tail [lindex $openFileList $i]]
    set f2 ""
    catch {set f2 [file tail [lindex [$File entryconfigure [expr {$i+$filemenuinc}] -label] 4]]}
    #if {$f1 != $f2 && $f2 != ""} {errorMsg "File list and menu out of synch: $i $f1 $f2"}
  }
  
# save the state so that if the program crashes the file list will be already saved
  saveState
  return
}

#-------------------------------------------------------------------------------
# open a spreadsheet
proc openXLS {filename {check 0} {multiFile 0}} {
  global lastXLS pf64 buttons

  if {[info exists buttons]} {.tnb select .tnb.status}

  if {[file exists $filename]} {

# check if instances of Excel are already running
    if {$check} {checkForExcel}
    #outputMsg " "
    
# start Excel
    if {[catch {
      #outputMsg "Starting Excel" green
      set xl [::tcom::ref createobject Excel.Application]
      [$xl ErrorCheckingOptions] TextDate False

# errors
    } emsg]} {
      errorMsg "ERROR starting Excel: $emsg"
    }
    
# open spreadsheet in Excel, works even if Excel not already started above although slower
    if {[catch {
      outputMsg "\nOpening Spreadsheet: [file tail $filename]  ([expr {[file size $filename]/1024}] Kb)" blue
      exec {*}[auto_execok start] "" $filename

# errors
    } emsg]} {
      errorMsg "ERROR opening Spreadsheet: $emsg"
      outputMsg " "
      set funcstr "F2"
      if {$multiFile} {set funcstr "F3"}
      errorMsg "Use $funcstr to open the spreadsheet or\n Go to the directory with the spreadsheet and open it."
      catch {raise .}
    }

  } else {
    if {[file tail $filename] != ""} {errorMsg "Spreadsheet not found: [truncFileName [file nativename $filename]]"}
    set filename ""
  }
  return $filename
}

#-------------------------------------------------------------------------------
proc checkForExcel {{multFile 0}} {
  global lastXLS localName buttons
  
  set pid1 [twapi::get_process_ids -name "EXCEL.EXE"]
  #outputMsg "checkForExcel-$pid1"
  
  if {[llength $pid1] > 0} {
    if {[info exists buttons]} {
      if {!$multFile} {
        set msg "There are ([llength $pid1]) instances of Excel already running.\nThe spreadsheets for the other instances might not be visible but will show up in the Windows Task Manager as EXCEL.EXE"
        append msg "\n\nThey might affect generating, saving, or viewing a new Excel spreadsheet."
        append msg "\n\nDo you want to close the other instances of Excel?"

        set dflt yes
        if {[info exists lastXLS] && [info exists localName]} {
          if {[llength $pid1] == 1} {if {[string first [file nativename [file rootname $localName]] [file nativename $lastXLS]] != 0} {set dflt no}}
        }
        set choice [tk_messageBox -type yesno -default $dflt -message $msg -icon question -title "Close Excel?"]

        if {$choice == "yes"} {
          #outputMsg "Closing Excel" red
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
          #if {$nnc > 0} {errorMsg "Some instances ($nnc) of Excel were not closed: $emsg" red}
        }
      }
    } else {
      #outputMsg "Closing Excel" red
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
proc getNextUnusedColumn {ent r} {
  global worksheet
  return [expr {[[[$worksheet($ent) UsedRange] Columns] Count]} + 1]
}

# -------------------------------------------------------------------------------
proc formatComplexEnt {str {space 0}} {
  global entCategory opt
  
  set str1 $str

# possibly format for _and_
  if {[string first "_and_" $str1] != -1} {

# check if _and_ is part of the entity name
    set ok 1
    foreach cat {PR_STEP_AP242 PR_STEP_COMM PR_STEP_TOLR PR_STEP_PRES PR_STEP_KINE PR_STEP_COMP} {
      if {$opt($cat)} {if {[lsearch $entCategory($cat) $str] != -1} {set ok 0; break}}
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
  set letters ABCDEFGHIJKLMNOPQRSTUVWXYZ
  
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
proc addCellComment {ent r c text} {
  global worksheet
  set comment [[$worksheet($ent) Range [cellRange $r $c]] AddComment]
  $comment Text $text
  [$comment Shape] Width  [expr double(250)]
  [$comment Shape] Height [expr double(100)]
}

#-------------------------------------------------------------------------------
proc colorBadCells {ent} {
  global excelVersion syntaxErr count cells worksheet stepAP
  
  if {$stepAP == ""} {return}
      
# color "Bad" (red) for syntax errors
  set lastr 4
  set rmax [expr {$count($ent)+3}]
  
  if {$excelVersion >= 12} {
    set syntaxErr($ent) [lsort [lrmdups $syntaxErr($ent)]]
    for {set n 0} {$n < [llength $syntaxErr($ent)]} {incr n} {
      if {[catch {
        set err [lindex $syntaxErr($ent) $n]

# get row and column number
        set r [lindex $err 0]
        set c [lindex $err 1]

# row and column are integers
        if {[string is integer $c]} {
          if {$r <= $rmax} {[$worksheet($ent) Range [cellRange $r $c] [cellRange $r $c]] Style "Bad"}

# values are entity ID or row number (row) and attribute name (column)
        } else {
          #outputMsg "\n$ent / $r / $c / [string is integer $c]" red
          if {![info exists nc($c)]} { 
            for {set i 2} {$i < 30} {incr i} {
              set val [[$cells($ent) Item 3 $i] Value]
              if {$val == $c} {
                set nc($c) $i
                break
              }
            }
          }
          
          if {[info exists nc($c)]} {
            set c $nc($c)
          
# entity ID
            if {$r > 0} {
              for {set i $lastr} {$i <= $rmax} {incr i} {
                set val [[$cells($ent) Item $i 1] Value]
                if {$val == $r} {
                  set r $i
                  set lastr [expr {$r+1}]
                  [$worksheet($ent) Range [cellRange $r $c] [cellRange $r $c]] Style "Bad"
                  break
                }              
              }
            } else {
              set r [expr {abs($r)}]
              [$worksheet($ent) Range [cellRange $r $c] [cellRange $r $c]] Style "Bad"
            }
          }
        }
      } emsg]} {
        errorMsg "ERROR setting spreadsheet cell color: $emsg\n  $ent"
        catch {raise .}
      }
    }
  }
}

#-------------------------------------------------------------------------------
proc trimNum {num {prec 3} {checkcomma 0}} {
  global unq_num comma
  
  set numsav $num
  if {[info exists unq_num($numsav)]} {
    set num $unq_num($numsav)
  } else {
    if {[catch {
      set form "\%."
      append form $prec
      append form "f"
      set num [format $form $num]

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
    if {$checkcomma && $comma} {regsub -all {\.} $num "," num}
    set unq_num($numsav) $num
  }
  return $num
}

#-------------------------------------------------------------------------------
proc outputMsg {msg {color "black"}} {
  global outputWin

  if {[info exists outputWin]} {
    $outputWin issue "$msg " $color
    update 
  } else {
    puts $msg
  }
}

#-------------------------------------------------------------------------------
proc errorMsg {msg {color ""}} {
  global errmsg outputWin stepAP

  if {![info exists errmsg]} {set errmsg ""}
  
  if {[string first $msg $errmsg] == -1} {
    set errmsg "$msg\n$errmsg"
    
# this fix is necessary to handle messages related to inverses
    set c1 [string first "DELETETHIS" $msg]
    if {$c1 != -1} {set msg [string range $msg 0 $c1-1]}
    
    puts $msg
    if {[info exists outputWin]} {
      if {$color == ""} {
        if {[string first "syntax error" [string tolower $msg]] != -1} {
          if {$stepAP != ""} {$outputWin issue "$msg " syntax}
        } else {
          set ilevel ""
          catch {set ilevel "  \[[lindex [info level [expr {[info level]-1}]] 0]\]"}
          if {$ilevel == "  \[errorMsg\]"} {set ilevel ""}
          $outputWin issue "$msg$ilevel " error
        }
      } else {
        $outputWin issue "$msg " $color
      }
      update 
    }
    return 1
  } else {
    return 0
  }
}

# -------------------------------------------------------------------------------------------------
proc truncFileName {fname {compact 0}} {
  global mydocs myhome mydesk

  if {[string first $mydocs $fname] == 0} {
    set nname "[string range $fname 0 2]...[string range $fname [string length $mydocs] end]"
  } elseif {[string first $mydesk $fname] == 0 && $mydesk != $fname} {
    set nname "[string range $fname 0 2]...[string range $fname [string length $mydesk] end]"
  #} elseif {[string first $myhome $fname] == 0 && $myhome != $fname} {
  #  set nname "[string range $fname 0 2]...[string range $fname [string length $myhome] end]"
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
# copy schema rose files that are in the Tcl Virtual File System (VFS) to the IFCsvr dll directory
# this only works with Tcl 8.5.15 and lower
proc copyRoseFiles {} {
  global programfiles wdir mytemp developer env ifcsvrdir nistVersion
  
  if {[file exists $ifcsvrdir]} {

# rose files in SFA distribution
    if {[llength [glob -nocomplain -directory [file join $wdir schemas] *.rose]] > 0} {
      set ok 1
      foreach fn [glob -nocomplain -directory [file join $wdir schemas] *.rose] {
          
        set fn1 [file tail $fn]
        set f2 [file join $ifcsvrdir $fn1]
        set okcopy 0
        if {![file exists $f2]} {
          set okcopy 1
        } elseif {[file mtime $fn] > [file mtime $f2]} {
          set okcopy 1
        }

# developer schemas
        set devfile 0
        if {[string first "am_ap242" $fn] != -1} {set devfile 1}
        if {$developer == 0 && $devfile} {
          set okcopy 0
          if {[file exists $f2]} {catch {file delete -force $f2}}
        }

# copy files
        if {$okcopy} {
          if {[catch {
            file copy -force $fn $f2
            if {$developer} {outputMsg "Copying ROSE file: $fn1" red}
          } emsg]} {
            errorMsg "ERROR copying STEP schema files (*.rose) to $ifcsvrdir"
            .tnb select .tnb.status
          }
          if {![file exists [file join $ifcsvrdir $fn1]]} {
            set ok 0
            if {[catch {
              file copy -force $fn [file join $mytemp $fn1]
            } emsg1]} {
              #errorMsg "ERROR: $emsg1"
            }
          }
        }
      }

# error copying files
      if {!$ok} {
        errorMsg "STEP schema files (*.rose) could not be copied to IFCsvr/dll directory"
        errorMsg "Check if any STEP APs are supported at Help > Supported STEP APs"
        errorMsg "If none are supported, then before continuing,\n copy the *.rose file(s) in $mytemp\n to $ifcsvrdir\nThis might require administrator privileges."
        .tnb select .tnb.status
      }
    }
  } else {
    #errorMsg "ERROR: IFCsvr Toolkit needs to be installed before copying STEP schema files (*.rose) to\n $ifcsvrdir"
  }
}

#-------------------------------------------------------------------------------
# install IFCsvr
proc installIFCsvr {} {
  global wdir mydocs mytemp ifcsvrdir nistVersion

  set ifcsvr     "ifcsvrr300_setup_1008_en.msi"
  set ifcsvrinst [file join $wdir schemas $ifcsvr]

# install if not already installed
  #outputMsg "installIFCsvr [file exists $ifcsvrdir] $ifcsvrdir" red
  if {![file exists $ifcsvrdir]} {
    .tnb select .tnb.status
    set msg "The IFCsvr Toolkit needs to be installed to read and process STEP files."
    outputMsg $msg red
    if {[file exists $ifcsvrinst]} {
      set msg "The IFCsvr Toolkit needs to be installed to read and process STEP files."
      append msg "\n\nAfter clicking OK the IFCsvr Toolkit installation will start.\nUse the default installation folder for IFCsvr.\nPlease wait for the installation process to complete before generating a spreadsheet."
      append msg "\n\nSee Help > Supported STEP APs to see which type of STEP files are supported."
      set choice [tk_messageBox -type ok -message $msg -icon info -title "Install IFCsvr"]
      set msg "\nPlease wait for the installation process to complete before generating a spreadsheet.\n"
      outputMsg $msg red
    }

# try copying installation file to several locations
    set ifcsvrmsi [file join $mytemp $ifcsvr]
    if {[file exists $ifcsvrinst]} {
      if {[catch {
        file copy -force $ifcsvrinst $ifcsvrmsi
      } emsg1]} {
        set ifcsvrmsi [file join $mydocs $ifcsvr]
        if {[catch {
          file copy -force $ifcsvrinst $ifcsvrmsi
        } emsg2]} {
          set ifcsvrmsi [file join [pwd] $ifcsvr]
          if {[catch {
            file copy -force $ifcsvrinst $ifcsvrmsi
          } emsg3]} {
            errorMsg "ERROR copying the IFCsvr Toolkit installation file to a directory."
            outputMsg " $emsg1\n $emsg2\n $emsg3"
          }
        }
      }
    }

# ready or not to install
    if {[file exists $ifcsvrmsi]} {
      exec {*}[auto_execok start] "" $ifcsvrmsi
    } else {
      if {[file exists $ifcsvrinst]} {errorMsg "IFCsvr Toolkit cannot be automatically installed."}
      outputMsg " "
      if {!$nistVersion} {
        errorMsg "To install the IFCsvr Toolkit you must install the NIST version of the STEP File Analyzer."
        outputMsg " 1 - Go to https://go.usa.gov/yccx"
        outputMsg " 2 - Click on Download STEP File Analyzer"
        outputMsg " 3 - Fill out the form, submit it, and follow the instructions"
        outputMsg " 4 - IFCsvr Toolkit will be installed when the NIST STEP File Analyzer is run"
        outputMsg " 5 - Generate a spreadsheet for at least one STEP file"
        after 1000
        openURL https://go.usa.gov/yccx
      } else {
        errorMsg "To manually install IFCsvr:"
        outputMsg " 1 - Join the IFCsvr ActiveX Component Group (you will need a Yahoo account)"
        outputMsg "     https://groups.yahoo.com/neo/groups/ifcsvr-users/info"
        outputMsg " 2 - Download the installer (In the Yahoo group: Files > IFCsvrR300 > ifcsvrr300_setup_1008_en.zip)"
        outputMsg " 3 - Extract the installer  ifcsvrr300_setup_1008.en.msi  from the zip file"
        outputMsg " 4 - Run the installer and follow the instructions.  Use the default installation folder for IFCsvr."
        outputMsg " 5 - Rerun this software."
        outputMsg "\nIf there are still problems with the IFCsvr installation, email the Contact (Help > About)"
        after 1000
        openURL https://groups.yahoo.com/neo/groups/ifcsvr-users/info
      }
    }

# delete the installation program if it is already installed
  } else {
    #catch {file delete -force [file join $mytemp $ifcsvr]}
    #catch {file delete -force [file join $mydocs $ifcsvr]}
    #catch {file delete -force [file join [pwd]   $ifcsvr]}
  }
}

#-------------------------------------------------------------------------------
# shortcuts
proc setShortcuts {} {
  global mydesk mymenu mytemp nistVersion wdir
  
  set progname [info nameofexecutable]
  if {[string first "AppData/Local/Temp" $progname] != -1 || [string first ".zip" $progname] != -1} {
    errorMsg "For the STEP File Analyzer to run properly, it is recommended that you first\n extract all of the files from the ZIP file and run the extracted executable."
    return
  }

  set progstr "STEP File Analyzer"
  if {!$nistVersion} {set progstr "SFA"}
  
  if {[info exists mydesk] || [info exists mymenu]} {
    set ok 1
    set app "STEP_Excel"
    foreach scut [list "Shortcut to $app.exe.lnk" "$app.exe.lnk" "$app.lnk"] {
      catch {if {[file exists [file join $mydesk $scut]]} {set ok 0; break}}
    }
    if {[file exists [file join $mydesk [file tail [info nameofexecutable]]]]} {set ok 0}

    if {$ok} {
      set choice [tk_messageBox -type yesno -icon question -title "Shortcuts" \
        -message "Do you want to create or overwrite a shortcut to the $progstr (v[getVersion]) in the Start Menu and an icon on the Desktop?"]
    } else {
      set choice [tk_messageBox -type yesno -icon question -title "Shortcuts" \
        -message "Do you want to create or overwrite a shortcut to the $progstr (v[getVersion]) in the Start Menu"]
    }
    if {$choice == "yes"} {
      outputMsg " "
      if {$nistVersion} {catch {[file copy -force [file join $wdir images NIST.ico] [file join $mytemp NIST.ico]]}}
      catch {
        if {[info exists mymenu]} {
          if {[file exists [file join $mymenu "$progstr.lnk"]]} {outputMsg "Existing Start Menu shortcut will be overwritten" red}
          if {$nistVersion} {
            create_shortcut [file join $mymenu "$progstr.lnk"] Description $progstr \
              TargetPath [info nameofexecutable] IconLocation [file join $mytemp NIST.ico]
          } else {
            create_shortcut [file join $mymenu "$progstr.lnk"] Description $progstr TargetPath [info nameofexecutable]
          }
          outputMsg "Shortcut to the $progstr (v[getVersion]) created in Start Menu to [truncFileName [file nativename [info nameofexecutable]]]"
        }
      }

      if {$ok} {
        catch {
          if {[info exists mydesk]} {
            if {[file exists [file join $mydesk "$progstr.lnk"]]} {outputMsg "Existing Desktop icon will be overwritten" red}
            if {$nistVersion} {
              create_shortcut [file join $mydesk "$progstr.lnk"] Description $progstr \
                TargetPath [info nameofexecutable] IconLocation [file join $mytemp NIST.ico]
            } else {
              create_shortcut [file join $mydesk "$progstr.lnk"] Description $progstr TargetPath [info nameofexecutable]
            }
            outputMsg "Icon for the $progstr (v[getVersion]) created on Desktop to [truncFileName [file nativename [info nameofexecutable]]]"
          }
        }
      }
    }
  }
}

#-------------------------------------------------------------------------------
# set home, docs, desktop, menu directories
proc setHomeDir {} {
  global env tcl_platform drive myhome mydocs mydesk mymenu mytemp

  set drive "C:/"
  if {[info exists env(SystemDrive)]} {
    set drive $env(SystemDrive)
    append drive "/"
  }
  set myhome $drive

# set based on USERPROFILE and registry entries
  if {[info exists env(USERPROFILE)]} {
    set myhome $env(USERPROFILE)
    catch {
      set reg_personal [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Personal}]
      if {[string first "%USERPROFILE%" $reg_personal] == 0} {regsub "%USERPROFILE%" $reg_personal $env(USERPROFILE) mydocs}
    }
    catch {
      set reg_desktop  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]
      if {[string first "%USERPROFILE%" $reg_desktop]  == 0} {regsub "%USERPROFILE%" $reg_desktop  $env(USERPROFILE) mydesk}
    }
    catch {
      set reg_menu [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Programs}]
      if {[string first "%USERPROFILE%" $reg_menu] == 0} {regsub "%USERPROFILE%" $reg_menu $env(USERPROFILE) mymenu}
    }
    
# windows 7 or greater    
    if {$tcl_platform(osVersion) >= 6.0} {
      set reg_temp [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local AppData}]
      if {[string first "%USERPROFILE%" $reg_menu] == 0} {regsub "%USERPROFILE%" $reg_temp $env(USERPROFILE) mytemp}
      set mytemp [file join $mytemp Temp]
      if {[string first $env(USERNAME) $mytemp] == -1} {
        unset mytemp
      } else {
        if {[file exists [file join $mytemp NIST]]} {catch {file delete -force [file join $mytemp NIST]}}
        set mytemp [file nativename [file join $mytemp SFA]]
        if {![file exists $mytemp]} {file mkdir $mytemp}
      }

# windows xp
    } else {
      catch {
        set reg_temp [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local Settings}]
        if {[string first "%USERPROFILE%" $reg_menu] == 0} {regsub "%USERPROFILE%" $reg_temp $env(USERPROFILE) mytemp}
        set mytemp [file join $mytemp Temp]
        if {[string first $env(USERNAME) $mytemp] == -1} {
          unset mytemp
        } else {
          if {[file exists [file join $mytemp NIST]]} {catch {file delete -force [file join $mytemp NIST]}}
          set mytemp [file join $mytemp SFA]
          if {![file exists $mytemp]} {file mkdir $mytemp}
        }
      }
    }
  }

# construct directories from drive and env(USERNAME)
  if {[info exists env(USERNAME)] && $myhome == $drive} {
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
    set temp [file join AppData Local Temp]
    if {$tcl_platform(osVersion) < 6.0} {set temp [file join "Local Settings" Temp]}
    set temp [file join $mytemp $temp]
    if {[file exists $temp]} {if {[file isdirectory $temp]} {set mytemp $temp}}
  }

  set myhome [file nativename $myhome]
  set mydocs [file nativename $mydocs]
  set mydesk [file nativename $mydesk]
  set mytemp [file nativename $mytemp]
  set drive [string range $myhome 0 2]
}

# -------------------------------------------------------------------------------
# http://wiki.tcl.tk/1844
proc get_shortcut_filename {file} {
  set dir [file nativename [file dirname $file]]
  set tail [file nativename [file tail $file]]

  if {![string match ".lnk" [string tolower [file extension $file]]]} {
    return -code error "$file is not a valid shortcut name"
  }

  if {[string match "windows" $::tcl_platform(platform)]} {

# Get Shortcut file as an object
    set oShell [tcom::ref createobject "Shell.Application"]
    set oFolder [$oShell NameSpace $dir]
    set oFolderItem [$oFolder ParseName $tail]
    
# If its a shortcut, do modify
    if {[$oFolderItem IsLink]} {
      set oShellLink [$oFolderItem GetLink]
      set path [$oShellLink Path]
      regsub -all {\\} $path "/" path
      return $path
    } else {
      if {![catch {file readlink $file} new]} {
        set new
      } else {
        set file
      }
    }
  } else {
    if {![catch {file readlink $file} new]} {
      set new
    } else {
      set file
    }
  }
}

proc create_shortcut {file args} {
  if {![string match ".lnk" [string tolower [file extension $file]]]} {
    append file ".lnk"
  }

  if {[string match "windows" $::tcl_platform(platform)]} {
# Make sure filenames are in nativename format.
    array set opts $args
    foreach item [list IconLocation Path WorkingDirectory] {
      if {[info exists opts($item)]} {
        set opts($item) [file nativename $opts($item)]
      }
    }

    set oShell [tcom::ref createobject "WScript.Shell"]
    set oShellLink [$oShell CreateShortcut [file nativename $file]]
    foreach {opt val} [array get opts] {
      if {[catch {$oShellLink $opt $val} result]} {
        return -code error "Invalid shortcut option $opt or value $value: $result"
      }
    }
    $oShellLink Save
    return 1
  }
  return 0
}
 
#-------------------------------------------------------------------------------
proc memusage {{str ""}} {
  global anapid lastmem
  
  if {[info exists anapid]} {
    if {![info exists lastmem]} {set lastmem 0}
    set mem [lindex [twapi::get_process_info $anapid -workingset] 1]
    set dmem [expr {$mem-$lastmem}]
    outputMsg "  $str  dmem [expr {$dmem/1000}]  mem [expr {$mem/1000}]" red
    set lastmem $mem
  }
}

#-------------------------------------------------------------------------------
proc getTiming {{str ""}} {
  global tlast

  set t [clock clicks -milliseconds]
  if {[info exists tlast]} {outputMsg "Timing: [expr {($t-$tlast)}]  $str" red}
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
# From http://wiki.tcl.tk/3070
proc stringSimilarity {a b} {
  set totalLength [max [string length $a] [string length $b]]
  return [string range [max [expr {double($totalLength-[levenshteinDistance $a $b])/$totalLength}] 0.0] 0 4]
}

proc levenshteinDistance {s t} {
  if {![set n [string length $t]]} {
    return [string length $s]
  } elseif {![set m [string length $s]]} {
    return $n
  }
  for {set i 0} {$i <= $m} {incr i} {
    lappend d 0
    lappend p $i
  }
  for {set j 0} {$j < $n} {} {
    set tj [string index $t $j]
    lset d 0 [incr j]
    for {set i 0} {$i < $m} {} {
      set a [expr {[lindex $d $i]+1}]
      set b [expr {[lindex $p $i]+([string index $s $i] ne $tj)}]
      set c [expr {[lindex $p [incr i]]+1}]
      lset d $i [expr {$a<$b ? $c<$a ? $c : $a : $c<$b ? $c : $b}]
    }
    set nd $p; set p $d; set d $nd
  }
  return [lindex $p end]
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
  outputMsg "Unique to L1 ([llength [lindex $l3 0]])\n  [lrange [lindex $l3 0] 0 500]"
  outputMsg "Common to both ([llength [lindex $l3 1]])\n  [lrange [lindex $l3 1] 0 500]"
  outputMsg "Unique to L2 ([llength [lindex $l3 2]])\n  [lrange [lindex $l3 2] 0 600]"
}
