# NIST STEP File Analyzer and Viewer

The free [STEP File Analyzer and Viewer](https://www.nist.gov/services-resources/software/step-file-analyzer-and-viewer) (SFA) 
generates a spreadsheet and visualization from an [ISO 10303 Part 21](https://www.loc.gov/preservation/digital/formats/fdd/fdd000448.shtml) STEP file.  STEP [AP242](https://www.ap242.org/), 
[AP203](https://www.iso.org/standard/44305.html), [AP214](https://www.iso.org/standard/43669.html), 
[AP209](https://www.ap209.org/), [AP238](https://ap238.org/), and other EXPRESS schemas are supported.

The STEP File Viewer supports parts, assemblies, graphic PMI for dimensions and tolerances, datum targets, sketch geometry, supplemental geometry, 
viewpoints, clipping planes, point clouds, composite rosettes, hole features, AP242 tessellated part geometry and polyhedral B-rep geometry, and 
AP209 finite element models and results.  AP242 XML files (.stpx) are also supported in the Viewer.

Viewer Examples: [Part with graphic PMI for GD&T](https://pages.nist.gov/CAD-PMI-Testing/graphical-pmi-viewer.html), 
[Box assembly](https://pages.nist.gov/CAD-PMI-Testing/step-file-viewer.html), 
[Bracket assembly](https://pages.nist.gov/CAD-PMI-Testing/bracket.html), 
[Section view clipping planes](https://pages.nist.gov/CAD-PMI-Testing/section-views.html), 
[AP209 finite element analysis models](https://pages.nist.gov/CAD-PMI-Testing/ap209-viewer.html)

The Analyzer generates a spreadsheet of all entity and attribute information; reports and analyzes any semantic PMI, 
graphic PMI, and validation properties for conformance to [CAx-IF recommended practices](https://www.mbx-if.org/home/cax/recpractices/); 
and checks for basic STEP file format errors.

Spreadsheet Example: [Spreadsheet](https://www.nist.gov/document/sfa-semantic-pmi-spreadsheet) with reports for semantic PMI, graphic PMI, and 
validation properties.

## Download or Build

**Download** the most recent NIST version of SFA (SFA-5.nn.zip) in the Release directory above.  Also in the Release directory are the README file, Release Notes,
User Guide, sample STEP files, and old versions of SFA.

**Build** your own version of SFA from the source code and instructions in the 'source' directory above.  There is no need to build your own version
unless you are modifying the source code.

## Disclaimers

[NIST Disclaimer](https://www.nist.gov/copyrights-disclaimers)
