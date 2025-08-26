# STEP File Analyzer and Viewer source code

## Tcl files

See the Release Notes in the Release directory.

Each Tcl file contains multiple procedures.

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
- sfa-grafpmi.tcl - process graphic PMI and tessellated geometry
- sfa-grafx3d.tcl - generate viewer X3D graphics
- sfa-gui.tcl - generate user interface
- sfa-hole.tcl - process counterbore/sink/drill and spotface holes
- sfa-help.tcl - help menu
- sfa-indent.tcl - indents STEP file for tree view
- sfa-inv.tcl - process inverse relationships
- sfa-multi.tcl - process multiple STEP files
- sfa-nist.tcl - process expected PMI for NIST models
- sfa-part.tcl - generate b-rep part geometry
- sfa-pmi.tcl - generate graphic PMI and tessellated part geometry
- sfa-proc.tcl - utility procedures
- sfa-step.tcl - STEP utility procedures
- sfa-supp.tcl - generate supplemental geometry graphics
- sfa-tess.tcl - process tessellated geometry
- sfa-uuid.tcl - process UUIDs
- sfa-valprop.tcl - process validation properties

- tclIndex - required Tcl code that lists all procedures in each Tcl file

- sfa-files.txt - freewrap input file that lists all of the above files

- teapot.zip - additional Tcl packages

## Disclaimers

[NIST Disclaimer](https://www.nist.gov/disclaimer)
