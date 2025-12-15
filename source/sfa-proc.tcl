# -------------------------------------------------------------------------------------------------
# set the min/max xyz coordinates
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
    if {[info exists entColorIndex($i)]} {
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
      if {[info exists entColorIndex($i)]} {
        set ent1 [string range $ent 0 $c1-1]
        if {[lsearch $entCategory($i) $ent1] != -1} {set tc1 $entColorIndex($i)}
        if {$c2 == $c1} {
          set ent2 [string range $ent $c1+5 end]
          if {[lsearch $entCategory($i) $ent2] != -1} {set tc2 $entColorIndex($i)}
        } elseif {$c2 != $c1} {
          set ent2 [string range $ent $c1+5 $c2-1]
          if {[lsearch $entCategory($i) $ent2] != -1} {set tc2 $entColorIndex($i)}
          set ent3 [string range $ent $c2+5 end]
          if {[lsearch $entCategory($i) $ent3] != -1} {set tc3 $entColorIndex($i)}
        }
      }
    }
    set tc [expr {min($tc1,$tc2,$tc3)}]

# exception for STEP measures (except *_11 composites entities)
    if {($tc1 == $entColorIndex(stepQUAN) || $tc2 == $entColorIndex(stepQUAN) || $tc3 == $entColorIndex(stepQUAN)) && [string first "_11" $ent] == -1} {
      set tc $entColorIndex(stepQUAN)
    }

# fix some AP209 entities with '_and_'
    if {[string first "AP209" $stepAP] != -1} {foreach str $andEntAP209 {if {[string first $str $ent] != -1} {set tc 19}}}

    if {$tc < 1000} {return $tc}
  }

# entity not in any category, color by AP
  foreach ap {AP209 AP210 AP224 AP235 AP238 AP239 ISO13584 CUTTING_TOOL} {if {[string first $ap $stepAP] != -1} {return 19}}

# entity from other APs (no color)
  return -2
}

#-------------------------------------------------------------------------------
# open a URL in whatever program is associated for the file type
proc openURL {url} {

  if {[catch {
    exec {*}[auto_execok start] "" $url
  } emsg]} {

# error message depends on the file and error type
    catch {.tnb select .tnb.status}
    if {[string first "is not recognized" $emsg] == -1} {
      if {[string first "UNC" $emsg] != -1} {set emsg [fixErrorMsg $emsg]}
      if {$emsg != ""} {

# error opening STEP file in app
        if {[string first ".stp" $url] != -1} {
          errorMsg "No app is associated with STEP files.  See Websites > STEP > STEP File Viewers"

# error opening viewer file
        } elseif {[string first "-sfa.html" $url] != -1} {
          errorMsg "Error opening Viewer file: $emsg\n Try manually opening the file in a web browser"

# error opening spreadsheet
        } elseif {[string first ".xlsx" $url] != -1} {
          if {[string first "The process cannot access the file" $emsg] != -1} {
            outputMsg " The Spreadsheet might already be opened." red
          } else {
            outputMsg " Error opening the Spreadsheet: $emsg" red
          }
        } else {
          errorMsg "Error opening the file: $emsg"
        }
      }
    } elseif {[string first "&" $url] != -1} {
      errorMsg "File cannot be opened because of the '&' in the file name.  Try manually opening the file."
    } else {
      errorMsg "Error opening the file: $emsg"
    }
  }
}

#-------------------------------------------------------------------------------
# file open dialog
proc openFile {{openName ""}} {
  global buttons drive editorCmd fileDir gen localName localNameList

  if {$openName == ""} {

# file types for file select dialog
    set typelist [list {"STEP " {".stp" ".step" ".p21" ".stpZ" ".stpA"}}]
    lappend typelist {"STL " {".stl"}}
    lappend typelist {"IFC " {".ifc"}}

# file open dialog
    set localNameList [tk_getOpenFile -title "Open File(s)" -filetypes $typelist -initialdir $fileDir -multiple true]
    if {[llength $localNameList] <= 1} {
      set localName [lindex $localNameList 0]
      if {$localName == ""} {return}
    } else {
      foreach name $localNameList {
        if {[string tolower [file extension $name]] == ".stpa"} {
          catch {.tnb select .tnb.status}
          errorMsg "Compressed archive files (.stpA) can only be selected one at a time.  Try again."
          set localNameList {}
          set localName ""
          return
        }
      }
    }

# file name passed in as openName
  } else {
    set localName $openName
    set localNameList [list $localName]
  }

# STL file
  set ok 0
  if {[llength $localNameList] > 1} {
    if {[string tolower [file extension [lindex $localNameList 0]]] == ".stl"} {set ok 1}
  } elseif {[file exists $localName]} {
    if {[string tolower [file extension $localName]] == ".stl"} {set ok 1}
  }
  if {$ok} {
    set opt(viewPart) 0
    set opt(partOnly) 0
    set gen(View) 1
    set allNone -1
    checkValues
    addFileToMenu
  }

# check for single zipped archive with possible multiple root names
  if {[info exists localName]} {
    if {[string first ".stpa" [string tolower $localName]] != -1} {
      unzipFile
      if {[llength [lindex $localName 0]] == 1} {
        set localName [file join [lindex $localName 1] [join [lindex $localName 0]]]
      } elseif {[llength [lindex $localName 0]] > 1} {
        set localNameList {}
        foreach item [lindex $localName 0] {lappend localNameList [file join [lindex $localName 1] $item]}
        set localName [lindex $localNameList 0]
      }
    }
  }

# multiple files selected
  if {[llength $localNameList] > 1} {
    set fileDir [file dirname [lindex $localNameList 0]]
    set str "STEP"
    if {$ok} {set str "STL"}

    outputMsg "\nReady to process [llength $localNameList] $str files" green
    if {[info exists buttons]} {
      $buttons(generate) configure -state normal
      if {[info exists buttons(appOpen)]} {$buttons(appOpen) configure -state normal}
      focus $buttons(generate)
    }

# single file selected
  } elseif {[file exists $localName]} {
    catch {pack forget $buttons(progressBarMulti)}

# check for zipped file
    set fext ""
    catch {set fext [string tolower [file extension $localName]]}
    if {$fext == ".stpz"} {unzipFile}

    set fileDir [file dirname $localName]
    if {[string first "z" [string tolower [file extension $localName]]] == -1} {
      outputMsg "\nReady to process: [file tail $localName]  ([fileSize $localName]  [fileTime $localName])" green
      checkFileSize

# check file extension
      set fext ""
      catch {set fext [string tolower [file extension $localName]]}

      if {$fileDir == $drive} {outputMsg "There might be problems processing the STEP file directly in the $fileDir directory." red}
      if {[info exists buttons]} {
        $buttons(generate) configure -state normal
        if {[info exists buttons(appOpen)]} {$buttons(appOpen) configure -state normal}
        focus $buttons(generate)
        if {$editorCmd != ""} {
          bind . <Key-F5> {
            if {[file exists $localName]} {
              outputMsg "\nOpening STEP file: [file tail $localName]"
              exec $editorCmd [file nativename $localName] &
            }
          }
          bind . <Shift-F5> {
            if {[file exists $localName]} {
              set dir [file nativename [file dirname $localName]]
              outputMsg "\nOpening STEP file directory: [truncFileName $dir]"
              catch {exec C:/Windows/explorer.exe $dir &}
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

# -------------------------------------------------------------------------------------------------
# get first file from file menu
proc getFirstFile {} {
  global buttons editorCmd openFileList

  set localName [lindex $openFileList 0]
  if {$localName != ""} {
    outputMsg "\nReady to process: [file tail $localName]  ([fileSize $localName]  [fileTime $localName])" green

    if {[info exists buttons(appOpen)]} {
      .tnb select .tnb.status
      $buttons(appOpen) configure -state normal
      if {$editorCmd != ""} {
        bind . <Key-F5> {
          if {[file exists $localName]} {
            outputMsg "\nOpening STEP file: [file tail $localName]"
            exec $editorCmd [file nativename $localName] &
          }
        }
        bind . <Shift-F5> {
          if {[file exists $localName]} {
            set dir [file nativename [file dirname $localName]]
            outputMsg "\nOpening STEP file directory: [truncFileName $dir]"
            catch {exec C:/Windows/explorer.exe $dir &}
          }
        }
      }
    }
  }
  return $localName
}

#-------------------------------------------------------------------------------
# add to the file menu
proc addFileToMenu {} {
  global buttons File localName openFileList stlFile

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
  if {$ifile != 0 && ($fext == ".stp" || $fext == ".stpa" || $fext == ".step" || $fext == ".p21" || $fext == ".ifc" || $fext == ".stl")} {
    if {$fext == ".stl"} {set stlFile 1}
    if {![info exists stlFile] || $fext == ".stl"} {
      set openFileList [linsert $openFileList 0 $localName]
      $File insert $filemenuinc command -label [truncFileName [file nativename $localName] 1] -command [list openFile $localName] -accelerator "F1"
      catch {$File entryconfigure 5 -accelerator {}}
    }
    if {[info exists stlFile] && $fext == ".stp"} {unset stlFile}
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
  }

# save the state so that if the program crashes the file list will be already saved
  saveState
  return
}

#-------------------------------------------------------------------------------
# unzip a STEP file (.stpZ .stpA)
proc unzipFile {} {
  global localName mytemp wdir

  if {[catch {
    outputMsg "\nUnzipping: [file tail $localName] ([fileSize $localName])" green
    set fzip {}

# unzip with vfs (works for .stpA, maybe not .stpZ)
    if {[catch {
      vfs::zip::Mount $localName stpzip
      set fzip [glob -nocomplain stpzip/*]
      set ztyp "vfs"
    } emsg]} {

# vfs error
      if {$emsg == "no header found"} {
        set ok 0

# unzip with 7-Zip command line (7z)
        set zcmd [file nativename "C:/Program Files/7-Zip/7z.exe"]
        if {[file exists $zcmd]} {
          if {[catch {
            set ztyp "7z"
            set fstp [string range $localName 0 end-1]
            outputMsg " Extracting ($ztyp): [file tail $fstp]"
            exec $zcmd e $localName -o[file dirname $localName] -aoa
            catch {file delete -force -- $fstp}
            file rename -force -- [file rootname $localName] $fstp
            set localName $fstp
            set ok 1
          } emsg2]} {
            errorMsg "Error extracting file with 7z"
          }
        }

# unzip with gunzip.exe packaged with SFA (original method)
        if {!$ok} {
          set ztyp "gunzip"
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
            outputMsg " Extracting ($ztyp): [file tail $ftmp]"
            exec $gunzip -Nf $fzip

# copy to new stp file
            set fstp [file join [file dirname $localName] [file tail $ftmp]]
            file copy -force -- $ftmp $fstp
            set localName $fstp
          }
        }
      } else {
        errorMsg "Unexpected error when unzipping: $emsg"
      }
    }

# vfs - .stpZ - one file
    if {$ztyp == "vfs"} {
      if {[llength $fzip] == 1} {
        set ftmp [string range [join $fzip] 7 end]
        outputMsg " Extracting ($ztyp): $ftmp"
        set fstp [file join [file dirname $localName] $ftmp]
        file copy -force -- $fzip $fstp
        set localName $fstp

# .stpA - multiple fzip
      } elseif {[llength $fzip] > 1} {

# copy fzip to directory based on file name
        set dirname [file join [file dirname $localName] [file tail [file rootname $localName]]]
        outputMsg " Extracting files to directory: [file tail $dirname]"
        catch {file delete -force -- $dirname}
        catch {file mkdir $dirname}
        foreach file $fzip {file copy -force -- $file $dirname}

# read ini file
        set iname [file join $dirname "stpa.ini"]
        if {[file exists $iname]} {
          set fini [open $iname r]
          while {[gets $fini line] >= 0} {

# get root file name
            if {[string first "Names=" $line] != -1} {
              set rnames {}
              set lline [split $line "\""]
              foreach item $lline {
                set litem [string tolower $item]
                if {[string first ".stp" $litem] != -1} {lappend rnames $item}
              }
              if {[llength $rnames] > 1} {outputMsg " Multiple root names: $rnames" red}
              set localName [list $rnames $dirname]
            }
          }
          close $fini
        } else {
          errorMsg "Syntax Error: Missing stpa.ini (STEP File Compression, Sec. 3.4.2)"
        }
      }
    }
  } emsg1]} {
    errorMsg "Error unzipping file: $emsg1"
  }
}

#-------------------------------------------------------------------------------
# file size
proc fileSize {fn} {
  set fs [expr {[file size $fn]/1024}]
  if {$fs < 10000} {
    return "$fs KB"
  } else {
    set fs [expr {round(double($fs)/1024.)}]
    if {$fs < 1024} {
      return "$fs MB"
    } else {
      set fs [trimNum [expr {double($fs)/1024.}] 2]
      return "$fs GB"
    }
  }
}

#-------------------------------------------------------------------------------
# file time
proc fileTime {fn} {
  set ftime [file mtime $fn]
  return "[clock format $ftime -format %d]-[clock format $ftime -format %b]-[clock format $ftime -format %Y]  [clock format $ftime -format %H:%M:%S]"
}

#-------------------------------------------------------------------------------
# check file size
proc checkFileSize {} {
  global localName

  if {[file size $localName] > 429000000} {
    set str "might be"
    if {[file size $localName] > 463000000} {set str "is"}
    outputMsg " The file $str too large to generate a Spreadsheet.  The limit is about 430 MB.\n For the Viewer, use Part Only on the Generate tab." red
    if {[file size $localName] > 1530000000} {outputMsg " The Viewer has not been tested with such a large STEP file." red}
  }
}

#-------------------------------------------------------------------------------
# save the state of variables to STEP-File-Analyzer-options.dat
proc saveState {{ok 1}} {
  global buttons commaSeparator dispCmd dispCmds fileDir fileDir1 filesProcessed gen lastX3DOM lastXLS lastXLS1 mydocs openFileList
  global opt optionsFile sfaVersion statusFont upgradeIFCsvr userEntityFile userWriteDir

# ok = 0 only after installing IFCsvr from the command-line version
  if {![info exists buttons] && $ok} {return}

  if {[catch {
    if {![file exists $optionsFile]} {outputMsg "\nCreating options file: [file nativename $optionsFile]"}
    set fileOptions [open $optionsFile w]
    puts $fileOptions "# Options file for the NIST STEP File Analyzer and Viewer [getVersion] ([string trim [clock format [clock seconds]]])"
    puts $fileOptions "# Do not edit or delete this file from the home directory $mydocs  Doing so might corrupt the current settings or cause errors.\n"

# opt variables
    foreach idx [lsort -nocase [array names opt]] {
      if {[string first "DEBUG" [string toupper $idx]] == -1 && [string first "indent" $idx] == -1 && \
          $idx != "PMISEMDIM" && $idx != "PMISEMDT" && $idx != "partNoGroup" && $idx != "tessPartMesh" && $idx != "viewTessPart"} {
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

    puts $fileOptions "set gen(View) $gen(View)"
    puts $fileOptions "\n# The lines below can be deleted for a command-line version (sfa-cl.exe) custom options file.\n"

# window position
    set winpos "+300+200"
    catch {
      set wg [winfo geometry .]
      set winpos [string range $wg [string first "+" $wg] end]
      set wingeo [string range $wg 0 [expr {[string first "+" $wg]-1}]]
    }
    catch {puts $fileOptions "set wingeo \"$wingeo\""}
    catch {puts $fileOptions "set winpos \"$winpos\""}

# variables in varlist, handle variables with [ or ]
    set varlist(1) [list statusFont upgradeIFCsvr sfaVersion filesProcessed commaSeparator]
    set varlist(2) [list fileDir fileDir1 userWriteDir userEntityFile lastXLS lastXLS1 lastX3DOM]
    set varlist(3) [list openFileList dispCmd dispCmds]
    foreach idx {1 2 3} {
      foreach var $varlist($idx) {
        if {[info exists $var]} {
          set vartmp [set $var]
          if {[string first "/" $vartmp] != -1 || [string first "\\" $vartmp] != -1 || [string first " " $vartmp] != -1} {
            regsub -all {\\} $vartmp "/" vartmp
            regsub -all {\[} $vartmp "\\\[" vartmp
            regsub -all {\]} $vartmp "\\\]" vartmp
            if {$var != "dispCmds" && $var != "openFileList"} {
              puts $fileOptions "set $var \"$vartmp\""
            } else {
              for {set i 0} {$i < [llength $vartmp]} {incr i} {
                if {$i == 0} {
                  if {[llength $vartmp] > 1} {
                    puts $fileOptions "set $var \"\{[lindex $vartmp $i]\} \\"
                  } else {
                    puts $fileOptions "set $var \"\{[lindex $vartmp $i]\}\""
                  }
                } elseif {$i == [expr {[llength $vartmp]-1}]} {
                  puts $fileOptions "  \{[lindex $vartmp $i]\}\""
                } else {
                  puts $fileOptions "  \{[lindex $vartmp $i]\} \\"
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
        if {$var == "openFileList"} {puts $fileOptions " "}
      }
      if {$idx < 3} {puts $fileOptions " "}
    }
    close $fileOptions

  } emsg]} {
    errorMsg "Error writing to options file: $emsg"
    catch {raise .}
  }
}

#-------------------------------------------------------------------------------
# open a STEP file in an app
proc runOpenProgram {} {
  global appName dispCmd File localName

  set dispFile $localName
  set idisp [file rootname [file tail $dispCmd]]
  if {[info exists appName]} {if {$appName != ""} {set idisp $appName}}

  outputMsg "\nOpening STEP file in: $idisp"

# open file
  if {[string first "Tree View" $idisp] == -1 && [string first "Default" $idisp] == -1} {

# start up with a file
    if {[catch {
      exec $dispCmd [file nativename $dispFile] &
    } emsg]} {
      errorMsg $emsg
    }

# default viewer associated with file extension
  } elseif {[string first "Default" $idisp] == 0} {
    openURL [file nativename $dispFile]

# file tree view
  } elseif {[string first "Tree View" $idisp] != -1} {
    .tnb select .tnb.status
    indentFile $dispFile

# all others
  } else {
    .tnb select .tnb.status
    outputMsg "You have to manually import the STEP file to $idisp." red
    exec $dispCmd &
  }

# add file to menu
  addFileToMenu
}

#-------------------------------------------------------------------------------
# open a spreadsheet
proc openXLS {filename {check 0} {multiFile 0}} {
  global buttons

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
      errorMsg "Error starting Excel: $emsg"
    }

# open spreadsheet in Excel, works even if Excel not already started above although slower
    outputMsg "\nOpening Spreadsheet: [file tail $filename]"
    openURL [file nativename $filename]

  } else {
    if {[file tail $filename] != ""} {
      outputMsg "\nOpening Spreadsheet: [file tail $filename]"
      outputMsg " Spreadsheet not found: [truncFileName [file nativename $filename]]" red
    }
    set filename ""
  }
  return $filename
}

# -------------------------------------------------------------------------------------------------
# open x3dom file
proc openX3DOM {{fn ""} {numFile 0}} {
  global buttons guiSFA lastX3DOM multiFile opt x3dMsgColor x3dFileName viz

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
      if {$opt(xlFormat) == "None"} {errorMsg "There is nothing in the STEP file for the Viewer to show based on the selections on the Generate tab."}
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
  } elseif {$numFile == 0 && $guiSFA != -1 && [info exists buttons]} {
    outputMsg " Use F3 to open the Viewer file" red
  }
}

#-------------------------------------------------------------------------------
# save the log file of messages from the Status tab
proc saveLogFile {lfile} {
  global buttons currLogFile editorCmd logFile multiFile

  outputMsg "\nSaving Log file to:"
  outputMsg " [truncFileName [file nativename $lfile]]" blue
  close $logFile
  if {!$multiFile && [info exists buttons]} {
    set currLogFile $lfile
    bind . <Key-F4> {
      if {[file exists $currLogFile]} {
        outputMsg "\nOpening Log file: [file tail $currLogFile]"
        exec $editorCmd [file nativename $currLogFile] &
      }
    }
  }
  unset logFile
}

#-------------------------------------------------------------------------------
# check if there are instances of Excel already open
proc checkForExcel {{multFile 0}} {
  global buttons lastXLS localName opt

  set pid1 [twapi::get_process_ids -name "EXCEL.EXE"]
  if {[llength $pid1] > 0 && $opt(xlFormat) != "None"} {
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
            foreach pid $pid1 {catch {twapi::end_process $pid -force}}
            set pid1 [twapi::get_process_ids -name "EXCEL.EXE"]
            if {[llength $pid1] == 0} {break}
          }
        }
      }

# stop excel for command-line version
    } else {
      foreach pid $pid1 {catch {twapi::end_process $pid -force}}
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
# format a complex entity name, e.g., first_and_second > (first)(second)
proc formatComplexEnt {str {space 0}} {
  global andEntAP209 entCategory opt stepAP

# check for an attribute name, i.e., a_and_b.attr
  set attr ""
  set c1 [string first "." $str]
  if {$c1 != -1} {
    set attr [string range $str $c1 end]
    set str  [string range $str 0 $c1-1]
  }

# possibly format for _and_
  set str1 $str
  if {[string first "_and_" $str1] != -1} {

# check if _and_ is part of the entity name
    set ok 1
    foreach cat {stepAP242 stepCOMM stepTOLR stepPRES stepKINE stepCOMP stepOTHR} {
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

# add back attribute
  if {$attr != ""} {append str1 $attr}
  return $str1
}

#-------------------------------------------------------------------------------
# generate cell, e.g., 2 4 > D2
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
# add comment to cell
proc addCellComment {ent r c comment {multi 0}} {
  global recPracNames worksheet worksheet1

  if {$multi == 0} {
    if {![info exists worksheet($ent)] || [string length $comment] < 2} {return}
  } else {
    if {![info exists worksheet1($ent)] || [string length $comment] < 2} {return}
  }

# modify comment
  if {[catch {
    while {[string first "  " $comment] != -1} {regsub -all "  " $comment " " comment}
    if {[string first "Syntax" $comment] == 0} {set comment "[string range $comment 14 end]"}
    if {[string first "GISU" $comment]  != -1} {regsub "GISU" $comment "geometric_item_specific_usage"  comment}
    if {[string first "IIRU" $comment]  != -1} {regsub "IIRU" $comment "item_identified_representation_usage" comment}

    foreach idx [array names recPracNames] {
      if {[string first $recPracNames($idx) $comment] != -1} {
        append comment "  See Websites > CAx Recommended Practices"
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

# add comment, delete existing
    if {$multi == 0} {
      catch {[[$worksheet($ent) Range [cellRange $r $c]] ClearComments]}
      set comm [[$worksheet($ent) Range [cellRange $r $c]] AddComment]
    } else {
      catch {[[$worksheet1($ent) Range [cellRange $r $c]] ClearComments]}
      set comm [[$worksheet1($ent) Range [cellRange $r $c]] AddComment]
    }
    $comm Text $ncomment
    catch {[[$comm Shape] TextFrame] AutoSize [expr 1]}

# error
  } emsg]} {
    if {[string first "Unknown error" $emsg] == -1} {errorMsg "Error adding Cell Comment: $emsg\n  $ent"}
  }
}

#-------------------------------------------------------------------------------
# color bad cells red, add cell comment with message
proc colorBadCells {ent} {
  global buttons cells count entsWithErrors idRow legendColor stepAP syntaxErr worksheet

  if {$stepAP == ""} {return}

# color red for syntax errors
  set rmax [expr {$count($ent)+3}]
  set okcomment 0

  set star " "
  if {![info exists buttons]} {set star " * "}
  outputMsg "$star[formatComplexEnt $ent]" red

  set syntaxErr($ent) [lsort -integer -index 0 [lrmdups $syntaxErr($ent)]]
  foreach err $syntaxErr($ent) {
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
        errorMsg "Error setting cell color (red) or comment: $emsg\n  $ent"
        catch {raise .}
      }
    }
  }
}

#-------------------------------------------------------------------------------
# trim the precision of a number
proc trimNum {num {prec 3}} {

# format number with 'prec'
  set tnum $num
  if {[catch {
    set form "\%.$prec"
    append form "f"
    set tnum [format $form $tnum]

# remove trailing zeros
    if {[string first "." $tnum] != -1 && [string index $tnum end] == 0} {
      for {set i 0} {$i < $prec} {incr i} {set tnum [string trimright $tnum "0"]}
      if {$tnum == "-0"} {set tnum 0.}
    }
  } errmsg]} {
    set tnum $num
  }
  return $tnum
}

#-------------------------------------------------------------------------------
# write message to the Status tab with the correct color
proc outputMsg {msg {color "black"}} {
  global logFile opt outputWin

  if {[info exists outputWin]} {
    $outputWin issue "$msg " $color
    update idletasks
  } else {
    puts $msg
  }
  if {$opt(logFile) && [info exists logFile]} {puts $logFile $msg}
}

#-------------------------------------------------------------------------------
# write error message to the Status tab with the correct color
proc errorMsg {msg {color ""}} {
  global errmsg logFile opt outputWin

  set oklog 0
  if {$opt(logFile) && [info exists logFile]} {set oklog 1}

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
        set stars "***"
        set logmsg "$stars $msg"
        if {[info exists outputWin]} {
          $outputWin issue "$msg " syntax
        } else {
          puts $logmsg
        }

# regular error message, ilevel is the procedure the error was generated in
      } else {
        set ilevel ""
        catch {set ilevel "  \[[lindex [info level [expr {[info level]-1}]] 0]\]"}
        if {$ilevel == "  \[errorMsg\]"} {set ilevel ""}

        set stars " **"
        set logmsg "$stars $msg$ilevel"
        if {[info exists outputWin]} {
          $outputWin issue "$msg$ilevel " error
        } else {
          puts $logmsg
        }
      }

# error message with color
    } else {
      set stars " **"
      set logmsg "$stars $msg"
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
        foreach str $newmsg {append logmsg "\n$stars $str"}
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
# fix the error message when it contains "UNC"
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
# truncate the file name to be shorter
proc truncFileName {fname {compact 0}} {
  global mydesk mydocs

# if file is in Documents, Desktop, Downloads, or OneDrive, then shorten name
  catch {
    if {[string first $mydocs $fname] == 0} {
      set nname "[string range $fname 0 2]...[string range $fname [string length $mydocs] end]"
    } elseif {[string first $mydesk $fname] == 0 && $mydesk != $fname} {
      set nname "[string range $fname 0 2]...[string range $fname [expr {[string length $mydesk]-8}] end]"
    } elseif {[string first "Downloads" $fname] != -1} {
      set indices [regexp -inline -all -indices {\\} $mydocs]
      set mydown "[string range $mydocs 0 [lindex [lindex $indices 2] 0]]Downloads"
      if {[string first $mydown $fname] == 0} {
        set nname "[string range $fname 0 2]...[string range $fname [expr {[string length $mydown]-10}] end]"
      }
    } elseif {[string first "\\OneDrive" $fname] != -1 || [string first "/OneDrive" $fname] != -1} {
      set c1 [string first "OneDrive" $fname]
      set nname "[string range $fname 0 2]...[string range $fname $c1-1 end]"
    }
  }
  if {[info exists nname]} {if {$nname != "C:\\..."} {set fname $nname}}

# compact name if longer than 80 characters
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
    set fn "[string range $fn 0 $c1-1]-$i$fext"
    catch {[file delete -force -- $fn]}
    if {![file exists $fn]} {break}
  }
  return $fn
}

#-------------------------------------------------------------------------------
# check file name for bad characters
proc checkFileName {fn} {
  global mydocs

  set fnt [file tail $fn]
  set fnd [file dirname $fn]
  if {[string first "\[" $fnd] != -1 || [string first "\]" $fnd] != -1} {
    set fn [file nativename [file join $mydocs $fnt]]
    errorMsg "Saving Spreadsheet to the home directory because of \[\] in the directory name." red
  }
  if {[string first "\[" $fnt] != -1 || [string first "\]" $fnt] != -1} {
    regsub -all {\[} $fn "(" fn
    regsub -all {\]} $fn ")" fn
    errorMsg "For the Spreadsheet file name \[\] are replaced by ()" red
  }
  return $fn
}

#-------------------------------------------------------------------------------
# install IFCsvr (or remove to reinstall)
proc installIFCsvr {{exit 0}} {
  global buttons ifcsvrVer mydocs mytemp nistVersion upgradeIFCsvr wdir

# IFCsvr version depends on string entered when IFCsvr is repackaged for new STEP schemas
  set versionIFCsvr 20251215

# if IFCsvr is alreadly installed, get version from registry, decide to reinstall newer version
  if {[catch {

# check IFCsvr CLSID and get version registry value "yyyy.mm.dd" or old format "1.0.0 (NIST Update yyyy-mm-dd)"
# if either fails, then install or reinstall
    set verIFCsvr [registry get $ifcsvrVer {DisplayVersion}]

# remove extra characters to format version to be yyyymmdd to compare with versionIFCsvr above
    set c1 [string first "20" $verIFCsvr]
    if {$c1 != -1} {
      set verIFCsvr [string range $verIFCsvr $c1 end]
      if {[string index $verIFCsvr end] == ")"} {set verIFCsvr [string range $verIFCsvr 0 end-1]}
      regsub -all {\-} $verIFCsvr "" verIFCsvr
      regsub -all {\.} $verIFCsvr "" verIFCsvr
    } else {
      set verIFCsvr 0
    }

# old version, reinstall
    if {$verIFCsvr < $versionIFCsvr} {
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
- If the IFCsvr toolkit cannot be installed or if you choose to Cancel the installation,
  you will still be able to use the Viewer for Part Geometry.  Open a STEP file, check
  View and Part Only, and click Generate View.
- If SFA crashes the first time you run it, see SFA-README-FIRST.pdf"

    if {[file exists $ifcsvrInst] && [info exists buttons]} {
      set msg "The IFCsvr toolkit must be installed to read and process STEP files (User Guide section 2.2.1).  After clicking OK the IFCsvr toolkit installation will start."
      append msg "\n\nYou might need administrator privileges (Run as administrator) to install the toolkit.  Antivirus software might respond that there is a security issue with the toolkit.  The toolkit is safe to install.  Use the default installation folder for the toolkit."
      append msg "\n\nIf the IFCsvr toolkit cannot be installed or if you choose to Cancel the installation, you will still be able to use the Viewer for Part Geometry.  Open a STEP file, check View and Part Only, and click Generate View."
      append msg "\n\nIf SFA crashes the first time you run it, see SFA-README-FIRST.pdf"
      set choice [tk_messageBox -type ok -message $msg -icon info -title "Install IFCsvr"]
      outputMsg "\nWait for the installation to finish before processing a STEP file." red
    } elseif {![info exists buttons]} {
      outputMsg "\nRerun this software after the installation has finished to process a STEP file."
    }

# reinstall
  } else {
    errorMsg "The IFCsvr toolkit must be reinstalled to update the STEP schemas."
    outputMsg "- First REMOVE the current installation of the IFCsvr toolkit."
    outputMsg "    In the IFCsvr Setup Wizard select 'REMOVE IFCsvrR300 ActiveX Component' and Finish" red
    outputMsg "    If the REMOVE was not successful, then manually uninstall the 'IFCsvrR300 ActiveX Component'"
    if {[info exists buttons]} {
      outputMsg "- Then restart this software or process a STEP file to install the updated IFCsvr toolkit.  You might have to reinstall the toolkit as Administrator."
    } else {
      outputMsg "- Then run this software again to install the updated IFCsvr toolkit.  You might have to reinstall the toolkit as Administrator."
    }
    outputMsg "- If you have to reinstall the toolkit every time you start the software, then REMOVE the IFCsvr\n  toolkit and download a new copy of the STEP File Analyzer and Viewer and start over."

    if {[file exists $ifcsvrInst] && [info exists buttons]} {
      set msg "The IFCsvr toolkit must be reinstalled to update the STEP schemas."
      append msg "\n\nFirst REMOVE the current installation of the IFCsvr toolkit."
      append msg "\n\nIn the IFCsvr Setup Wizard (after clicking OK) select 'REMOVE IFCsvrR300 ActiveX Component' and Finish.  If the REMOVE was not successful, then manually uninstall the 'IFCsvrR300 ActiveX Component'"
      append msg "\n\nThen restart this software or process a STEP file to install the updated IFCsvr toolkit.  You might have to reinstall the toolkit as Administrator."
      append msg "\n\nIf you have to reinstall the toolkit every time you start the software, then REMOVE the IFCsvr toolkit and download a new copy of the STEP File Analyzer and Viewer and start over."
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
          errorMsg "Error copying the IFCsvr toolkit installation file to a directory."
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
      if {$emsg != ""} {errorMsg "Error installing IFCsvr toolkit: $emsg"}
    }

# cannot find the toolkit
  } else {
    if {[file exists $ifcsvrInst]} {errorMsg "The IFCsvr toolkit cannot be automatically installed."}
    catch {.tnb select .tnb.status}
    update idletasks

# manual install instructions
    if {$nistVersion} {
      outputMsg " "
      errorMsg "To manually install the IFCsvr toolkit:"
      outputMsg "- The installation file ifcsvrr300_setup_1008_en-update.msi can be found in $mytemp
- Run the installer and follow the instructions.  Use the default installation folder for IFCsvr.
  You might need administrator privileges (Run as administrator) to install the toolkit.\n"
      after 1000
      errorMsg "Opening folder: $mytemp"
      catch {exec C:/Windows/explorer.exe [file nativename $dir] &}
      if {$exit} {exit}
    } else {
      outputMsg " "
      errorMsg "To install the IFCsvr toolkit you must first run the NIST version of the STEP File Analyzer and Viewer."
      outputMsg "- Download the zip file on the software web page.
- Follow the instructions to run the software.
- The IFCsvr toolkit will be installed when the NIST STEP File Analyzer and Viewer is run."
      after 1000
      openURL https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer
    }
  }
}

#-------------------------------------------------------------------------------
# shortcuts on desktop and start menu
proc setShortcuts {} {
  global mydesk myhome mymenu mytemp wdir

  set progname [info nameofexecutable]
  if {[string first "AppData/Local/Temp" $progname] != -1 || [string first ".zip" $progname] != -1} {
    errorMsg "You should first extract all of the files from the ZIP file and run the extracted software."
    return
  }

  set progstr "STEP File Analyzer and Viewer"
  if {[info exists mydesk] || [info exists mymenu]} {
    set msg "Do you want to create or overwrite shortcuts to the $progstr [getVersion]"
    if {[info exists mydesk]} {
      append msg " on the Desktop"
      if {[info exists mymenu]} {append msg " and"}
    }
    if {[info exists mymenu]} {append msg " in the Start Menu"}
    append msg "?"

    if {[info exists mydesk] || [info exists mymenu]} {
      set choice [tk_messageBox -type yesno -icon question -title "Shortcuts" -message $msg]
      if {$choice == "yes"} {
        set temp [string range $mytemp 0 end-4]
        catch {[file copy -force -- [file join $wdir images NIST.ico] [file join $temp NIST.ico]]}
        catch {[file copy -force -- [file join $wdir images NIST.ico] [file join $mytemp NIST.ico]]}
        catch {if {[info exists mymenu]} {twapi::write_shortcut [file join $mymenu "$progstr.lnk"] -path [info nameofexecutable] -desc $progstr -iconpath [file join $temp NIST.ico]}}
        catch {if {[info exists mydesk]} {twapi::write_shortcut [file join $mydesk "$progstr.lnk"] -path [info nameofexecutable] -desc $progstr -iconpath [file join $temp NIST.ico]}}
      }
    }
  }
  catch {file delete -force [file join $myhome $progstr.lnk]}
}

#-------------------------------------------------------------------------------
# set home, docs, desktop, menu directories
proc setHomeDir {} {
  global drive env mydesk mydocs myhome mymenu mytemp wdir

# C drive
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
      if {[string first "%USERPROFILE%" $reg_personal] == 0} {
        set mydocs "$env(USERPROFILE)\\[string range $reg_personal 14 end]"
      } else {
        set mydocs $reg_personal
      }
    }
    catch {
      set reg_desktop  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]
      if {[string first "%USERPROFILE%" $reg_desktop] == 0} {
        set mydesk "$env(USERPROFILE)\\[string range $reg_desktop 14 end]"
      } else {
        set mydesk $reg_desktop
      }
    }
    catch {
      set reg_menu [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Programs}]
      if {[string first "%USERPROFILE%" $reg_menu] == 0} {set mymenu "$env(USERPROFILE)\\[string range $reg_menu 14 end]"}
    }

# set mytemp
    catch {
      set reg_temp [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local AppData}]
      if {[string first "%USERPROFILE%" $reg_temp] == 0} {set mytemp "$env(USERPROFILE)\\[string range $reg_temp 14 end]"}
      set mytemp [file join $mytemp Temp]

# make mytemp dir
      set mytemp [file nativename [file join $mytemp SFA]]
      checkTempDir
    }

# create myhome if USERPROFILE does not exist
  } elseif {[info exists env(USERNAME)]} {
    set myhome [file join $drive Users $env(USERNAME)]
  }

# directories if not created
  if {![info exists mydocs]} {
    set mydocs $myhome
    set docs [file join $mydocs "Documents"]
    if {[file exists $docs]} {if {[file isdirectory $docs]} {set mydocs $docs}}
  }

  if {![info exists mydesk]} {
    set mydesk $myhome
    set desk [file join $mydesk "Desktop"]
    if {[file exists $desk]} {if {[file isdirectory $desk]} {set mydesk $desk}}
  }

  if {![info exists mytemp]} {
    set mytemp $myhome
    set temp [file join AppData Local Temp SFA]
    set mytemp [file join $mytemp $temp]
    checkTempDir
  }

  set myhome [file nativename $myhome]
  set mydocs [file nativename $mydocs]
  set mydesk [file nativename $mydesk]
  set mytemp [file nativename $mytemp]

  set temp [string range $mytemp 0 end-4]
  catch {[file copy -force -- [file join $wdir images NIST.ico] [file join $temp NIST.ico]]}
  catch {[file copy -force -- [file join $wdir images NIST.ico] [file join $mytemp NIST.ico]]}
  update idletasks
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
# reformat timestamp in the HEADER section
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
# get timing for the code, getTiming is always commented out, uncomment to debug
proc getTiming {{str ""}} {
  global tlast

  set t [clock clicks -milliseconds]
  if {[info exists tlast]} {outputMsg "Timing: [expr {$t-$tlast}]  $str" red}
  set tlast $t
}

#-------------------------------------------------------------------------------
# From https://wiki.tcl-lang.org/page/Custom+sorting
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
# Based on http://www.cawt.tcl3d.org/
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
  if {$col <= 0} {errorMsg "Column number $col is bad."}
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
proc vectrim {v1 {precision 4}} {
  foreach c1 $v1 {
    set prec $precision
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
  set angle [trimNum [expr {acos([vecdot $v1 $v2] / ([veclen $v1]*[veclen $v2]))}] 5]
  return $angle
}
