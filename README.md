# NIST STEP File Analyzer and Viewer

The [NIST STEP File Analyzer and Viewer](https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer) (SFA) 
generates a spreadsheet and visualization from an ISO 10303 Part 21 STEP file.  More information, sample spreadsheets and visualizations, and documentation about SFA is available on the website 
including the [STEP File Analyzer and Viewer User Guide](https://www.nist.gov/publications/step-file-analyzer-and-viewer-user-guide-update-7).

The [NIST STEP to X3D Translator](https://www.nist.gov/services-resources/software/step-x3d-translator) is used by the SFA Viewer to convert STEP b-rep part geometry to X3D and has its own source code and executable.

Download a pre-built Windows version of SFA with the Release link (zip file) to the right. 
Follow the instructions below to build your own version of SFA from the source code.  

## Prerequisites

Microsoft Excel is required to generate spreadsheets.  CSV (comma-separated values) files will be generated if Excel is not installed.

Download the SFA files from the GitHub 'source' directory to a directory on your computer.

- The name of the directory is not important
- The STEP File Analyzer and Viewer is written in [Tcl](https://wiki.tcl-lang.org/)
- Some of the Tcl code is based on [CAWT](https://www.tcl3d.org/cawt/)

freewrap wraps the SFA Tcl code to create an executable.

- Download freewrap651.zip from <https://sourceforge.net/projects/freewrap/files/freewrap/freeWrap%206.51/>  More recent versions of freewrap will **not** work with wrapping SFA.
- Extract freewrap.exe and put it in the same directory as the SFA files that were downloaded from the 'source' directory.

Several Tcl packages not included in freewrap also need to be installed.

- teapot.zip from the 'source' directory contains the additional Tcl packages
- Create a directory C:/Tcl/lib
- Unzip teapot.zip to the 'lib' directory to create C:/Tcl/lib/teapot

## Build the STEP File Analyzer and Viewer

- Open a command prompt window and change to the directory with the SFA Tcl files and freewrap.
- To generate the executable **sfa.exe**, enter the command: freewrap -f sfa-files.txt

## Running the Software

**You must first install and run the NIST version of the STEP File Analyzer and Viewer before running your own version.**
- Click on Release to the right and download the zip file.
- Extract STEP-File-Analyzer.exe from the zip file, run it and process a STEP file to install other software.
- Some features are not available in the user-built version including tooltips, unzipping compressed STEP files, and those related to the NIST CAD models.
- Internally at NIST, SFA is built with [ActiveTcl 8.5.18 32-bit](https://www.activestate.com/products/tcl/) and the [Tcl Dev Kit](https://www.activestate.com/blog/tcl-dev-kit-now-open-source/) which is now an open source project.

## Disclaimers

[NIST Disclaimer](https://www.nist.gov/disclaimer)
