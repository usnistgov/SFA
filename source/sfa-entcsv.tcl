proc getEntityCSV {objEntity} {
  global thisEntType count
  global fixent fixprm localName
  global roseLogical
  global entCount badAttributes
  global csvfile csvdirnam csvstr fcsv

# get entity type
  set thisEntType [$objEntity Type]

# -------------------------------------------------------------------------------------------------
# csv file for each entity if it does not already exist
  if {![info exists csvfile($thisEntType)]} {
    set msg "[formatComplexEnt $thisEntType] ($entCount($thisEntType))"
    outputMsg $msg
    
    if {$thisEntType == "datum_reference_modifier_with_value"} {errorMsg " Datum reference value modifiers are ignored in the DRF"}
    update idletasks

# open csv file
    set csvfile($thisEntType) 1
    set csvfname [file join $csvdirnam $thisEntType.csv]
    set fcsv [open $csvfname w]
    puts $fcsv "[formatComplexEnt $thisEntType] ($entCount($thisEntType))"
    #outputMsg $fcsv red

# headings in first row
    set csvstr "ID"
    ::tcom::foreach objAttribute [$objEntity Attributes] {append csvstr ",[$objAttribute Name]"}
    puts $fcsv $csvstr
    unset csvstr

    set count($thisEntType) 0

# file of entities not to process
    set cfile [file rootname $localName]
    append cfile "_fix.dat"
    if {[catch {
      set fixfile [open $cfile w]
      foreach item $fixent {if {[lsearch $fixprm $item] == -1} {puts $fixfile $item}}
      if {[lsearch $fixent $thisEntType] == -1 && [lsearch $fixprm $thisEntType] == -1} {puts $fixfile $thisEntType}
      close $fixfile
    } emsg]} {
      errorMsg "ERROR processing 'fix' file: $emsg"
    }
    update idletasks
  }

# -------------------------------------------------------------------------------------------------
# start filling in the cells
  incr count($thisEntType)
  
# show progress with > 50000 entities
  if {$entCount($thisEntType) >= 50000} {
    set c1 [expr {$count($thisEntType)%20000}]
    if {$c1 == 0} {
      outputMsg " $count($thisEntType) of $entCount($thisEntType) processed"
      update idletasks
    }
  }

# entity ID
  set p21id [$objEntity P21ID]

# -------------------------------------------------------------------------------------------------
# for all attributes of the entity
  set nattr 0
  set csvstr $p21id
  set objAttributes [$objEntity Attributes]
  ::tcom::foreach objAttribute $objAttributes {
    set attrName [$objAttribute Name]
    #outputMsg "$p21id $attrName [$objAttribute NodeType]" red

    if {[catch {
      if {![info exists badAttributes($thisEntType)]} {
        set objValue [$objAttribute Value]

# look for bad attributes that cause a crash
      } else {
        set ok 1
        foreach ba $badAttributes($thisEntType) {if {$ba == $attrName} {set ok 0}}
        if {$ok} {
          set objValue [$objAttribute Value]
        } else {
          set objValue "???"
          errorMsg " Skipping '$attrName' attribute on $thisEntType" red
        }
      }

# error getting attribute value
    } emsgv]} {
      set msg "ERROR processing #[$objEntity P21ID]=[$objEntity Type] '$attrName' attribute: $emsgv"
      if {[string first "datum_reference_compartment 'modifiers' attribute" $msg] != -1 || \
          [string first "datum_reference_element 'modifiers' attribute" $msg] != -1 || \
          [string first "annotation_plane 'elements' attribute" $msg] != -1} {
        errorMsg "Syntax Error: On '[$objEntity Type]' entities change the '$attrName' attribute with\n '()' to '$' where applicable.  The attribute is an OPTIONAL SET\[1:?\] and '()' is not valid."
      }
      errorMsg $msg
      set objValue ""
      catch {raise .}
    }

    incr nattr

# -------------------------------------------------------------------------------------------------
# not a handle, just a single value
    if {[string first "handle" $objValue] == -1} {
      set ov $objValue
  
# if value is a boolean, substitute string roseLogical
      if {([$objAttribute Type] == "RoseBoolean" || [$objAttribute Type] == "RoseLogical") && [info exists roseLogical($ov)]} {set ov $roseLogical($ov)}

# check if displaying numbers without rounding
      append csvstr ",$ov"

# -------------------------------------------------------------------------------------------------
# if attribute is reference to another entity
    } else {
      
# node type 18=ENTITY, 19=SELECT TYPE  (node type is 20 for SET or LIST is processed below)
      if {[$objAttribute NodeType] == 18 || [$objAttribute NodeType] == 19} {
        set refEntity [$objAttribute Value]

# get refType, however, sometimes this is not a single reference, but rather a list
#  which causes an error and it has to be processed like a list below
        if {[catch {
          set refType [$refEntity Type]
          set valnotlist 1
        } emsg2]} {

# process like a list which is very unusual
          catch {foreach idx [array names cellval] {unset cellval($idx)}}
          ::tcom::foreach val $refEntity {
            append cellval([$val Type]) "[$val P21ID] "
          }
          set str ""
          set size 0
          catch {set size [array size cellval]}

          if {$size > 0} {
            foreach idx [lsort [array names cellval]] {
              set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
              if {$ncell > 1 || $size > 1} {
                if {$ncell < 30} {
                  append str "($ncell) [formatComplexEnt $idx 1] $cellval($idx)  "
                } else {
                  append str "($ncell) [formatComplexEnt $idx 1]  "
                }
              } else {
                append str "(1) [formatComplexEnt $idx 1] $cellval($idx)  "
              }
            }
          }
          append csvstr ",$str"
          set valnotlist 0
        }

# value is not a list which is the most common
        if {$valnotlist} {
          set str "[formatComplexEnt $refType 1] [$refEntity P21ID]"

# for length measure (and other measures), add the actual measure value
          if {[string first "measure_with_unit" $refType] != -1} {
            ::tcom::foreach refAttribute [$refEntity Attributes] {
              if {[$refAttribute Name] == "value_component"} {set str "[$refAttribute Value] ($str)"}
            }
          }
          append csvstr ",$str"
        }

# -------------------------------------------------------------------------------------------------
# node type 20=AGGREGATE (ENTITIES), usually SET or LIST, try as a tcom list or regular list (SELECT type)
      } elseif {[$objAttribute NodeType] == 20} {
        catch {foreach idx [array names cellval]     {unset cellval($idx)}}

        if {[catch {
          ::tcom::foreach val [$objAttribute Value] {

# collect the reference id's (P21ID) for the Type of entity in the SET or LIST
            append cellval([$val Type]) "[$val P21ID] "
          }

        } emsg]} {
          foreach val [$objAttribute Value] {
            append cellval([$val Type]) "[$val P21ID] "
          }
        }

# -------------------------------------------------------------------------------------------------
# format cell values for the SET or LIST
        set str ""
        set size 0
        catch {set size [array size cellval]}

        if {$size > 0} {
          foreach idx [lsort [array names cellval]] {
            set ncell [expr {[llength [split $cellval($idx) " "]] - 1}]
            if {$ncell > 1 || $size > 1} {
              if {$ncell < 30} {
                append str "($ncell) [formatComplexEnt $idx 1] $cellval($idx)  "
              } else {
                append str "($ncell) [formatComplexEnt $idx 1]  "
              }
            } else {
              append str "(1) [formatComplexEnt $idx 1] $cellval($idx)  "
            }
          }
        }
        append csvstr ",[string trim $str]"
      }
    }
  }
  #outputMsg "$fcsv $csvstr"

# write to CSV file
  if {[catch {
    puts $fcsv $csvstr
  } emsg]} {
    errorMsg "Error writing to CSV file for: $thisEntType"
  }

# -------------------------------------------------------------------------------------------------
# clean up variables to hopefully release some memory
  foreach var {objAttributes attrName refEntity refType} {if {[info exists $var]} {unset $var}}
  update idletasks
  return 1
}
