## Build SFA from the source code

Microsoft Excel is required to generate spreadsheets.  CSV (comma-separated values) files will be generated if Excel is not 
installed.  SFA is written in [Tcl](https://wiki.tcl-lang.org/) with some of the Tcl code based 
on [CAWT](https://www.tcl3d.org/cawt/).

To build SFA, first download the SFA code with the green 'Code' button on the previous page to a directory on your computer.  The name of the 
directory is not important.

**freewrap** wraps the SFA Tcl code to create an executable.

- Download [freewrap651.zip](https://sourceforge.net/projects/freewrap/files/freewrap%206/freeWrap%206.51/).  More recent versions of freewrap will **not** work with wrapping SFA.
- Extract freewrap.exe and put it in the same directory as the SFA files that were downloaded from the 'source' directory.

Several Tcl packages not included in freewrap also need to be installed.

- teapot.zip from this directory contains the additional Tcl packages
- Create a directory C:/Tcl/lib
- Unzip teapot.zip to the 'lib' directory to create C:/Tcl/lib/teapot

## Build the STEP File Analyzer and Viewer

- Open a command prompt window and change to the directory with the SFA Tcl files and freewrap.
- To generate the executable **sfa.exe**, enter the command: **freewrap -f sfa-files.txt**

Optionally build the command-line version:

- Download [freewrapTCLSH.zip](https://sourceforge.net/projects/freewrap/files/freewrap%206/freeWrap%206.51/)
- Extract freewrapTCLSH.exe to the directory with the SFA Tcl files
- Edit sfa-files.txt and change the first line 'sfa.tcl' to 'sfa-cl.tcl'
- To generate **sfa-cl.exe**, enter the command: **freewrapTCLSH -f sfa-files.txt**

## Run the User-built version

You must first install and run the NIST version of the SFA before running your own version.  SFA depends 
on two software packages (IFCsvr, stp2x3d) that are included with the Download version.  Some features are not available in the 
User-built version including tooltips, unzipping compressed STEP files, and those related to the NIST CAD models.

Click on Release directory above.  Extract STEP-File-Analyzer.exe from SFA-5.nn.zip, run it and process a STEP file to 
install other software.  Now you can run your own version of SFA.

## Tcl files

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
- sfa-unknown.tcl - process unknown entity types
- sfa-uuid.tcl - process UUIDs
- sfa-valprop.tcl - process validation properties
- tclIndex - required Tcl code that lists all procedures in each Tcl file
- sfa-files.txt - freewrap input file that lists all of the above files
- teapot.zip - additional Tcl packages

## Disclaimers

[NIST Disclaimer](https://www.nist.gov/pao/nist-disclaimer-statement)
