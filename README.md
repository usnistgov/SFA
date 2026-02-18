# STEP File Analyzer and Viewer

The free [STEP File Analyzer and Viewer](https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer) (SFA) 
generates a spreadsheet and visualization from an ISO 10303 Part 21 STEP file.  The STEP File Viewer supports parts, assemblies, dimensions, 
tolerances, and more.  The Analyzer generates a spreadsheet of all entity and attribute information; reports and analyzes any semantic PMI, 
graphic PMI, and validation properties for conformance to recommended practices; and checks for basic STEP file format errors.  STEP AP242, AP203, 
AP214, AP209, and other APs and EXPRESS schemas are supported.

## Viewer and Spreadsheet Examples

[Part with graphic PMI for GD&T](https://pages.nist.gov/CAD-PMI-Testing/graphical-pmi-viewer.html), [Box assembly](https://pages.nist.gov/CAD-PMI-Testing/step-file-viewer.html),
[Bracket assembly](https://pages.nist.gov/CAD-PMI-Testing/bracket.html), [Section view clipping planes](https://pages.nist.gov/CAD-PMI-Testing/section-views.html), 
[AP209 finite element analysis models](https://pages.nist.gov/CAD-PMI-Testing/ap209-viewer.html), [Spreadsheet](https://www.nist.gov/document/sfa-semantic-pmi-spreadsheet) 
with reports for semantic PMI, graphic PMI, and validation properties.

## Download or Build

**Download** the NIST version of SFA in the Release directory above.  Click on the zip file (SFA-5.nn.zip) in the Release directory and then the 
download icon to the right.  Read the README file.

**Build** your own version of SFA from the source code with the instructions below.

## Build Prerequisites

Microsoft Excel is required to generate spreadsheets.  CSV (comma-separated values) files will be generated if Excel is not 
installed.  SFA is written in [Tcl](https://wiki.tcl-lang.org/) with some of the Tcl code based 
on [CAWT](https://www.tcl3d.org/cawt/).

To build SFA, first download the Tcl and other files from the 'source' directory above to a directory on your computer.  The name of the 
directory is not important.

**freewrap** wraps the SFA Tcl code to create an executable.

- Download [freewrap651.zip](https://sourceforge.net/projects/freewrap/files/freewrap%206/freeWrap%206.51/).  More recent versions of freewrap will **not** work with wrapping SFA.
- Extract freewrap.exe and put it in the same directory as the SFA files that were downloaded from the 'source' directory.

Several Tcl packages not included in freewrap also need to be installed.

- teapot.zip from the 'source' directory contains the additional Tcl packages
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

## Disclaimers

[NIST Disclaimer](https://www.nist.gov/pao/nist-disclaimer-statement)
