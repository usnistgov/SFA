# NIST STEP File Analyzer and Viewer

The [NIST STEP File Analyzer and Viewer](https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer) (SFA) 
generates a spreadsheet and visualization from an ISO 10303 Part 21 STEP file.  More information, sample spreadsheets and visualizations, and documentation about SFA is available on the website 
including the [STEP File Analyzer and Viewer User Guide](https://www.nist.gov/publications/step-file-analyzer-and-viewer-user-guide-update-7).

The [NIST STEP to X3D Translator](https://www.nist.gov/services-resources/software/step-x3d-translator) is used by the SFA Viewer to convert STEP b-rep part geometry to X3D and has its own source code and executable.

## Download or Build

**Download** a pre-built Windows version of SFA in the Release directory above. 

**Build** your own version of SFA from the source code with the instructions below.  

## Build Prerequisites

Microsoft Excel is required to generate spreadsheets.  CSV (comma-separated values) files will be generated if Excel is not 
installed.  SFA is written in [Tcl](https://wiki.tcl-lang.org/) with some of the Tcl code based 
on [CAWT](https://www.tcl3d.org/cawt/).

To build SFA, first download the Tcl and other files from the GitHub 'source' directory to a directory on your computer.  The name of the 
directory is not important.

freewrap wraps the SFA Tcl code to create an executable.

- Download [freewrap651.zip](https://sourceforge.net/projects/freewrap/files/freewrap%206/freeWrap%206.51/).  More recent versions of freewrap will **not** work with wrapping SFA.
- Extract freewrap.exe and put it in the same directory as the SFA files that were downloaded from the 'source' directory.

Several Tcl packages not included in freewrap also need to be installed.

- teapot.zip from the 'source' directory contains the additional Tcl packages
- Create a directory C:/Tcl/lib
- Unzip teapot.zip to the 'lib' directory to create C:/Tcl/lib/teapot

## Build the STEP File Analyzer and Viewer

- Open a command prompt window and change to the directory with the SFA Tcl files and freewrap.
- To generate the executable **sfa.exe**, enter the command: freewrap -f sfa-files.txt

Optionally build the command-line version:

- Download [freewrapTCLSH.zip](https://sourceforge.net/projects/freewrap/files/freewrap%206/freeWrap%206.51/)
- Extract freewrapTCLSH.exe to the directory with the SFA Tcl files
- Edit sfa-files.txt and change the first line 'sfa.tcl' to 'sfa-cl.tcl'
- Edit sfa-cl.tcl similar to sfa.tcl above
- To generate **sfa-cl.exe**, enter the command: freewrapTCLSH -f sfa-files.txt

## Running the Software

**You must first install and run the NIST version of the STEP File Analyzer and Viewer before running your own version.**
- Click on Release above or to the right and download the zip file.
- Extract STEP-File-Analyzer.exe from the zip file, run it and process a STEP file to install other software.
- Some features are not available in the user-built version including tooltips, unzipping compressed STEP files, and those related to the NIST CAD models.
- Internally at NIST, SFA is built with [ActiveTcl 8.5.18 32-bit](https://www.activestate.com/products/tcl/) and the [Tcl Dev Kit](https://www.activestate.com/blog/tcl-dev-kit-now-open-source/) which is now an open source project.

## Disclaimers

[NIST Disclaimer](https://www.nist.gov/disclaimer)
