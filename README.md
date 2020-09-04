# NIST STEP File Analyzer and Viewer

With these instructions you can build the NIST STEP File Analyzer and Viewer (SFA) from the source code.  SFA generates a spreadsheet and visualization from an ISO 10303 Part 21 STEP file.  A pre-built Windows version of NIST STEP File Analyzer and Viewer is available [here](https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer).  More information, sample spreadsheets and visualizations, and documentation about SFA is available on the website including the [STEP File Analyzer and Viewer User Guide](https://www.nist.gov/publications/step-file-analyzer-and-viewer-user-guide-update-6).

STEP files are used to represent Product and Manufacturing Information (PMI) for data exchange and interoperability between Computer-Aided Design (CAD), Manufacturing (CAM), Analysis (CAE), and Inspection (CMM) software.  STEP is also used for the long-term archiving and retrieval product data (LOTAR).

## Prerequisites

The STEP File Analyzer and Viewer can only be built and run on Windows computers.  [Microsoft Excel](https://products.office.com/excel) is required to generate spreadsheets.  CSV (comma-separated values) files will be generated if Excel is not installed.  

**You must first install and run the NIST version of the STEP File Analyzer and Viewer before running your own version.**

- Go to the [STEP File Analyzer and Viewer](https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer) and click on the link to Download the STEP File Analyzer and Viewer
- Submit the form to get a link to download sfa-n.nn.zip
- Extract STEP-File-Analyzer.exe from the zip file and run it.  This will install the IFCsvr toolkit that is used to read STEP files.

Download the SFA files from the GitHub 'source' directory to a directory on your computer.

- The name of the directory is not important
- The STEP File Analyzer and Viewer is written in [Tcl](https://www.tcl.tk/)

freeWrap wraps the SFA Tcl code to create an executable.

- Download freewrap651.zip from <https://sourceforge.net/projects/freewrap/files/freewrap/freeWrap%206.51/>.  More recent versions of freeWrap will **not** work with wrapping SFA.
- Extract freewrap.exe and put it in the same directory as the SFA files that were downloaded from the 'source' directory.

Install the ActiveTcl **8.5 32-bit** version of Tcl.

- Download the ActiveTcl installer from <https://www.activestate.com/products/activetcl/downloads/>.  You will have to create an ActiveState account.
- The Windows installer file name is: ActiveTcl-8.5.18.0.nnnnnn-win32-ix86-threaded.exe
- SFA can be built only with ActiveTcl 8.5.18 (32-bit).  ActiveTcl 8.6.n and 64-bit versions are not supported.
- Run the installer and use the default installation folders

Several Tcl packages from ActiveTcl also need to be installed.  Open a command prompt window, change to C:\\Tcl\\bin, or wherever Tcl was installed, and enter the following three commands:

```
teacup install tcom
teacup install twapi
teacup install Iwidgets
```

## Build the STEP File Analyzer and Viewer

First, edit the source code file sfa.tcl and uncomment the lines at the top of the file that start with 'lappend auto_path C:/Tcl/lib/teapot/package/...'  Change 'C:/Tcl' if Tcl is installed in a different directory.

Then, open a command prompt window and change to the directory with the SFA Tcl files and freewrap.  To create the executable sfa.exe, enter the command:

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

## Suggested improvements

Replace the Tcl package [tcom](http://wiki.tcl.tk/1821) (COM) with the COM features in [twapi](http://twapi.magicsplat.com/).  This will allow upgrading to Tcl 8.6.

Replace the home-grown code to interact with Excel spreadsheets with [CAWT](http://www.cawt.tcl3d.org/).

## Contact

[Robert Lipman](https://www.nist.gov/people/robert-r-lipman), <robert.lipman@nist.gov>, 301-975-3829

## Disclaimers

[NIST Disclaimer](https://www.nist.gov/public_affairs/disclaimer.cfm)
