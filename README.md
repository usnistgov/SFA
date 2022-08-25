# NIST STEP File Analyzer and Viewer

With these instructions you can build the [NIST STEP File Analyzer and Viewer](https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer) (SFA) from the source code.  SFA generates a spreadsheet and visualization from an ISO 10303 Part 21 STEP file.  More information, sample spreadsheets and visualizations, and documentation about SFA is available on the website including the [STEP File Analyzer and Viewer User Guide](https://www.nist.gov/publications/step-file-analyzer-and-viewer-user-guide-update-7).  The [NIST STEP to X3D Translator](https://www.nist.gov/services-resources/software/step-x3d-translator) is used by SFA to convert STEP b-rep part geometry to X3D and has its own source code.

## Prerequisites

The STEP File Analyzer and Viewer can only be built and run on Windows computers.  This is due to a dependence on the IFCsvr toolkit that is used to read and parse STEP files.  IFCsvr only runs on Windows.

Microsoft Excel is required to generate spreadsheets.  CSV (comma-separated values) files will be generated if Excel is not installed.  

**You must first install and run the NIST version of the STEP File Analyzer and Viewer before running your own version.**

- Go to the [STEP File Analyzer and Viewer](https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer) to download the software
- Extract STEP-File-Analyzer.exe from the zip file and run it.  This will install the IFCsvr toolkit that is used to read STEP files.

Download the SFA files from the GitHub 'source' directory to a directory on your computer.

- The name of the directory is not important
- The STEP File Analyzer and Viewer is written in [Tcl](https://wiki.tcl-lang.org/)
- Some of the Tcl code is based on [CAWT](http://www.cawt.tcl3d.org/)

freeWrap wraps the SFA Tcl code to create an executable.

- Download freewrap651.zip from <https://sourceforge.net/projects/freewrap/files/freewrap/freeWrap%206.51/>.  More recent versions of freeWrap will **not** work with wrapping SFA.
- Extract freewrap.exe and put it in the same directory as the SFA files that were downloaded from the 'source' directory.

Install the ActiveTcl **8.5 32-bit** version of Tcl.

- Download the ActiveTcl installer from <https://www.activestate.com/products/tcl/>.  You will have to create an ActiveState account.
- **Tcl 8.5 32-bit might only be available as a paid legacy version of Tcl.**
- The Windows installer file name is: ActiveTcl-8.5.18.0.nnnnnn-win32-ix86-threaded.exe
- SFA can be built only with ActiveTcl 8.5.18 (32-bit).  ActiveTcl 8.6.n and 64-bit versions are not supported.
- Run the installer and use the default installation folders

Several Tcl packages from ActiveTcl also need to be installed.  Open a command prompt window, change to C:\\Tcl\\bin, or wherever Tcl was installed, and enter the following commands:

```
teacup install tcom
teacup install tdom
teacup install twapi
teacup install Iwidgets
```

## Build the STEP File Analyzer and Viewer

If Tcl is installed in a different directory than 'C:/Tcl', then edit the source code file sfa.tcl with the lines that start with 'lappend auto_path C:/Tcl/lib/teapot/package/...'

Then open a command prompt window and change to the directory with the SFA Tcl files and freewrap.  To create the executable sfa.exe, enter the command:

```
freewrap -f sfa-files.txt
```

**Optionally, build the STEP File Analyzer and Viewer command-line version**

- Download freewrapTCLSH.zip from <https://sourceforge.net/projects/freewrap/files/freewrap/freeWrap%206.51/>
- Extract freewrapTCLSH.exe to the directory with the SFA Tcl files
- Edit sfa-files.txt and change the first line 'sfa.tcl' to 'sfa-cl.tcl'
- Edit sfa-cl.tcl similar to sfa.tcl above
- To create sfa-cl.exe, enter the command: freewrapTCLSH -f sfa-files.txt

## Differences from the NIST-built version of STEP File Analyzer and Viewer

Some features are not available in the user-built version including: tooltips, unzipping compressed STEP files, automated PMI checking for the [NIST CAD models](<https://www.nist.gov/el/systems-integration-division-73400/mbe-pmi-validation-and-conformance-testing>), and inserting images of the NIST test cases in the spreadsheets.  Some of the features are restored if the NIST-built version is run first.

## Disclaimers

[NIST Disclaimer](https://www.nist.gov/disclaimer)
