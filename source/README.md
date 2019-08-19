# STEP File Analyzer and Viewer source code

## Tcl files

Changes to the source code are listed in the [changelog](https://s3.amazonaws.com/nist-el/mfg_digitalthread/STEP-File-Analyzer-changelog.xlsx).

- sfa.tcl - main program
- sfa-cl.tcl - main program for command-line version
- sfa-coverage.tcl - PMI coverage analysis
- sfa-data.tcl - set lots of variables
- sfa-dimtol.tcl - process dimensional tolerances
- sfa-ent.tcl - write STEP entity to worksheet or CSV file
- sfa-fea.tcl - process AP209 files
- sfa-gen.tcl - generate a spreadsheet
- sfa-geom.tcl- process associated geometry
- sfa-geotol.tcl- process geometric tolerances
- sfa-grafpmi.tcl - process graphical PMI and tessellated geometry
- sfa-grafx3d.tcl - generate X3DOM graphics for visualization
- sfa-gui.tcl - generate user interface
- sfa-hole.tcl - process counterbore/sink/drill and spotface holes
- sfa-indent.tcl - indents STEP file for tree view
- sfa-inv.tcl - process inverse relationships
- sfa-multi.tcl - process multiple STEP files
- sfa-nist.tcl - process expected PMI for NIST models
- sfa-proc.tcl - utility procedures
- sfa-step.tcl - STEP utility procedures
- sfa-tess.tcl - process tessellated geometry
- sfa-valprop.tcl - process validation properties

- tclIndex - required Tcl code that lists all procedures in each Tcl file

- sfa-files.txt - freewrap input file that lists all of the above files

## Contact

[Robert Lipman](https://www.nist.gov/people/robert-r-lipman), <robert.lipman@nist.gov>, 301-975-3829

## Disclaimers

[NIST Disclaimer](http://www.nist.gov/public_affairs/disclaimer.cfm)
