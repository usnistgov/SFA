# help menu, examples menu at the end
proc guiHelpMenu {} {
  global developer Examples filesProcessed Help ifcsvrDir ifcsvrVer mytemp opt scriptName stepAPs

  $Help add command -label "User Guide" -command {openUserGuide}
  $Help add command -label "Release Notes" -command {openURL https://www.nist.gov/document/sfa-release-notes}

  $Help add separator
  $Help add command -label "Overview" -command {
outputMsg "\nOverview ------------------------------------------------------------------------------------------" blue
outputMsg "The STEP File Analyzer and Viewer (SFA) opens a STEP file (ISO 10303 - STandard for Exchange of
Product model data) Part 21 file (.stp or .step or .p21 file extension) and

1 - generates an Excel spreadsheet or CSV files of all entity and attribute information,
2 - creates a visualization (view) of part geometry, graphic PMI, and other features that is
    displayed in a web browser,
3 - reports and analyzes validation properties, semantic PMI, and graphic PMI, and checks them for
    conformance to recommended practices, and
4 - checks for basic syntax errors.

Compressed STEP files (.stpZ) and STEP archive files (.stpA) are supported.
AP238 STEP-NC files (.stpnc) are supported by renaming the file extension to .stp
AP242 Domain Model XML files (.stpx) are not supported.

Help is available in this menu, in the User Guide, and in tooltip help.  New features are listed in
the Release Notes and described in some Help.  Help in the menu, tooltips, and spreadsheet comments
are more up-to-date than the User Guide."
    .tnb select .tnb.status
  }

# options help
  $Help add command -label "Options" -command {
outputMsg "\nOptions -------------------------------------------------------------------------------------------" blue
outputMsg "See Help > User Guide (sections 3.4, 3.5, 4, and 6)

Generate: Generate Excel spreadsheets, CSV files, and/or Views.  If Excel is not installed, CSV
files are automatically generated.  Some options are not supported with CSV files.  The Syntax
Checker can also be run when processing a STEP file.

All text in the Status tab can be written to a Log File when a STEP file is processed.  The log
file is written to myfile-sfa.log.  Syntax errors, warnings, and other messages are highlighted by
asterisks *.  Use F4 to open the log file.

Entity Types: Select which types of entities are processed from AP203, AP214, and AP242 for the
Spreadsheet.  All entities specific to other APs are always written to the Spreadsheet such as
AP209, AP210, and AP238.  The categories are used to group and color-code entities on the Summary
worksheet.  The tooltip help lists all the entities associated with that type.

Analyzer options report PMI and check for conformance to recommended practices.
- Semantic Representation PMI: Dimensional tolerances, geometric tolerances, and datum features are
  reported on various entities indicated by Semantic PMI on the Summary worksheet.
- Graphic Presentation PMI: Geometric entities used for Graphic PMI annotations are reported.
  Associated Saved Views, Validation Properties, and Geometry are also reported.
- Validation Properties: Geometric, assembly, PMI, annotation, attribute, and tessellated
  validation properties are reported.
- Inverse Relationships: For some entities, Inverse relationships and backwards references (Used In)
  are shown on the worksheets.

Viewer: Part geometry, graphic PMI annotations, tessellated part geometry in AP242 files, and AP209
finite element models can be shown in a web browser.

More tab: Spreadsheet formatting and other Analyzer and Viewer options."
    .tnb select .tnb.status
  }

  $Help add cascade -label "Viewer" -menu $Help.1
  set helpView [menu $Help.1 -tearoff 1]

# general viewer help
  $helpView add command -label "Overview" -command {
outputMsg "\nViewer Overview -----------------------------------------------------------------------------------" blue
outputMsg "The viewer generates an HTML file 'myfile-sfa.html' that is shown in the default web browser.  An
Internet connection is required.  The HTML file is self-contained and can be shared with other
users including those on non-Windows systems.  The viewer does not support measurements.

The viewer can be used without generating a spreadsheet.  See Generate on the Options tab.  The
Part Only option is useful when no other Viewer features are needed and for large STEP files.

The viewer supports boundary representation (b-rep) exact geometry.  For AP242, tessellated and
polyhedral b-rep geometry are supported.  Color, transparency, part edges, sketch geometry, and
assemblies are supported.  Part geometry viewer features:

- Part edges are shown in black.  Use the transparency slider to show only edges.  Some parts might
  not be affected by the transparency slider.  If a part is completely transparent and edges are
  not selected, then the part will not be visible.  In some cases transparency might look wrong for
  assemblies with many parts.

- Sketch geometry is supplemental lines created when generating a CAD model.  Sketch geometry is
  also known as construction, auxiliary, support, or reference geometry.  To show only sketch
  geometry, turn off edges and make the part completely transparent.  Sometimes processing sketch
  geometry will affect the behavior of the transparency slider.  Sketch geometry is not same as
  supplemental geometry.  See Help > Viewer > Supplemental Geometry

- Normals improve the default smooth shading by explicitly computing surface normals to improve the
  appearance of curved surfaces.

- Quality controls the number of facets used for curved surfaces.  Higher quality uses more facets
  around the circumference of a cylinder.  Using High Quality and the Normals options results in
  the best appearance for part geometry.  See the new feature below for the Alternative B-rep
  Geometry Processing.

- AP242 tessellated part geometry is typically written to a STEP file in addition to or instead of
  b-rep part geometry.  A wireframe mesh, outlining the facets of the tessellated surfaces can be
  shown with the Edges option.  In some cases, parts in an assembly might have the wrong position and
  orientation or be missing.  Quality and normals do not apply to tessellated part geometry.
  See Websites > CAx Recommended Practices (Tessellated 3D Geometry)

- The bounding box min and max XYZ coordinates are based on the faceted geometry being shown and
  not the exact geometry in the STEP file.  There might be a variation in the coordinates depending
  on the Quality option.  The bounding box also accounts for any sketch geometry if it is displayed
  but not graphic PMI and supplemental geometry.  The bounding box can be shown to confirm that the
  min and max coordinates are correct.  If the part is too large to rotate smoothly, turn off the
  part and rotate the bounding box.

- The origin of the model at '0 0 0' is shown with a small XYZ coordinate axis that can be switched
  off.  The background color can be changed between white, blue, gray, and black.

- See Help > Text Strings and Numbers for how non-English characters are supported.

For very large STEP files it might take several minutes to process the STEP part geometry.  To
speed up the process, select View and Part Only on the Generate tab.  The resulting HTML file might
also take several minutes to display in the web browser.  Select 'Wait' if the web browser prompts
that it is running slowly when opening the HTML file.

The viewer generates an X3D file that is embedded in the HTML file that is displayed in the default
web browser.  Select 'Save X3D ...' on the More tab to save the X3D file so that it can be shown in
an X3D viewer or imported to other software.  Part geometry including tessellated geometry and
graphic PMI is supported.  Use this option if an Internet connection is not available for the
Viewer.

The Viewer might indicate that there are possible syntax errors and to run the Syntax Checker.

See Help > User Guide (section 4)
See Help > Viewer for other topics

The viewer for part geometry is based on the NIST STEP to X3D Translator and only runs on 64-bit
computers.  It runs a separate program stp2x3d-part.exe from [file nativename $mytemp]
See Websites > STEP

Other STEP file viewers are available.  See Websites > STEP > STEP File Viewers.  Some of the
viewers are faster and have better features for viewing and measuring part geometry.  This viewer
supports many features that other viewers do not, including: graphic PMI, sketch geometry,
supplemental geometry, datum targets, viewpoints, clipping planes, point clouds, composite rosettes,
hole features, AP242 tessellated part geometry, and AP209 finite element models and results.  Try
the Open STEP Viewer with AP242 Domain Model XML files (.stpx)."
    .tnb select .tnb.status
  }

  $helpView add command -label "New Features" -command {
outputMsg "\nNew Features --------------------------------------------------------------------------------------" blue
outputMsg "These Viewer features are not documented in the User Guide (Update 7).

1 - Hidden buttons and sliders

Some checkboxes and sliders on the right side of the viewer might be hidden.  If hidden, they can
be shown by clicking on the buttons for More Options, Saved View Graphic PMI, and others.

2 - Cloud of points and point clouds

The cloud of points (COPS) geometric validation property are sampling points generated by the CAD
system on the surfaces and edges of a part.  The points are used to check the deviation of surfaces
from those points in an importing system.  The report for Validation Properties must be generated
to show the COPS.  See Websites > CAx Recommended Practices (Geometric and Assembly Validation Properties)

3D scanning point clouds are supported in AP242.  Point cloud colors, intensities, and normals are
not supported in the Viewer.

Points are shown with a blue dot.  In both cases, the exact points might not appear on part
surfaces because part geometry in the viewer is only a faceted approximation.  For parts in an
assembly, the COPS might have the wrong position and orientation.

3 - Section view clipping planes

Part geometry can be clipped by section view clipping planes defined in the STEP file.  Part Only
must not be selected on the Generate tab.  The planes are shown with a black square that might not
be centered on the model.  Checkboxes show the names of each clipping plane.  If there are
duplicate clipping plane names, a number in parentheses is appended to the name.  You have to
manually select the clipping plane that is associated with a viewpoint or saved view graphic PMI.

Use the option on the More tab to generate capped surfaces in the plane of a clipping plane.
Capped surfaces, in the plane of the black square, are generated when there is only one clipping
plane per section view.  Capped surfaces for parts in an assembly might be in the wrong position.
Sometimes capped surfaces are not generated.  Switching off parts in an assembly does not turn off
their capped surfaces.

4 - Parallel projection viewpoints

Use the option on the More tab to use parallel projection for saved view viewpoints defined in the
STEP file, instead of the default perspective projection.  See Help > Viewer > Viewpoints

5 - PMI placeholders

PMI placeholders provide information about the position, orientation, and organization of an
annotation without the graphic presentation of numeric values and symbols for geometric or
dimensional tolerances.  See Help > Viewer > PMI Placeholders

6 - B-rep part geometry processing

If curved surfaces for Part Geometry look wrong even with Quality set to High, select the
Alternative B-rep Geometry Processing method on the More tab.  It will take longer to process the
STEP file and the resulting Viewer file will be larger.

7 - Tessellated supplemental geometry

Supplemental geometry represented as tessellated geometry is supported.
See Help > Viewer > Supplemental Geometry

8 - Convert STL to AP242

STL files can be converted to STEP AP242 tessellated geometry that can be shown in the viewer.
In the Open File(s) dialog, change the 'Files of type' to 'STL (*.stl)'.  ASCII and binary STL
files are supported.  Tessellated geometry is not exact b-rep surfaces and may not be supported in
some CAD software.

9 - Composite rosettes defined by cartesian points and curves are shown in the viewer."
    .tnb select .tnb.status
  }

    $helpView add command -label "Viewpoints" -command {
outputMsg "\nViewpoints ----------------------------------------------------------------------------------------" blue
outputMsg "Use PageDown to switch between viewpoints in the viewer window.  Viewpoint names are shown in the
upper left corner of the viewer.  User-defined viewpoints are used with saved view graphic PMI.

If there are no user-defined viewpoints (saved views) in the STEP file, then front, side, top, and
isometric viewpoints are generated.  Since the default orientation of the part is not known, the
viewpoints might not correspond to the actual front, side, and top of the model.  The isometric
viewpoint might not be centered.  All of the viewpoints use perspective except for an additional
front parallel projection.

---------------------------------------------------------------------------------------------------
If there are user-defined viewpoints (saved views) in the STEP file, then in addition to the saved
views from the file, two additional front viewpoints named 'Front (SFA)' are generated, one
perspective and the other a parallel projection.  Pan and zoom might not work with parallel
projection.

If there are duplicate saved view names, then a number in parentheses is appended to the name.  For
example, two viewpoints named MBD_A will appear as MBD_A (1) and MBD_A (2) for the Viewpoint name
in the upper left corner of the viewer when cycling through the viewpoints with PageDown.

Saved view names with non-English characters (Unicode) are supported in the viewer if a spreadsheet
is also generated.

On the More tab, parallel projection viewpoints as defined in the STEP file can be used instead of
the default perspective.  Also, if the model has viewpoints with and without graphic PMI, then the
viewpoints without graphic PMI can also be shown.  Those viewpoints are usually top, front, and
side viewpoints.

If there is graphic PMI associated with saved views, then the PMI is automatically switched
on/off when using PageDown if 'Saved View Viewpoints' is checked on the Generate tab.  If there are
duplicate saved view names as described above, then the list of Saved View Graphic PMI will append,
after a slash, the saved view name to the PMI name.

Saved views always show all part geometry and ignore any view-specific geometry that is not visible
in a view.

---------------------------------------------------------------------------------------------------
In the viewer, use key 'a' to view all and 'r' to restore to the original view.  The function of
other keys is described in the link 'Use the mouse'.  Navigation uses the Examine Mode.

Sometimes a part is located far from the origin and not visible.  In this case, turn off the Origin
and Sketch Geometry and then use 'a' to view all.

Older implementations of saved views might not conform to current recommended practices.  The
resulting model orientation will look wrong.  Use the option on the More tab to correct for the
wrong orientation.  The position of the model might still be wrong.

The More tab option 'Show save view camera model viewpoints' can be used to debug the camera model
(view frustum) viewpoint geometry.

Saved views are ignored with Part Only.  Part visibility in saved views is not supported."
    .tnb select .tnb.status
  }

  $helpView add command -label "Assemblies" -command {
outputMsg "\nAssemblies ----------------------------------------------------------------------------------------" blue
outputMsg "Assemblies are related to part geometry, graphic PMI, and supplemental geometry.

Part Geometry for assemblies with b-rep or tessellated geometry is supported in the viewer.  Most
assemblies and parts can be switched on and off depending on the assembly structure.  An alphabetic
list of part and assembly names is shown on the right.

Parts with the same shape are usually grouped with the same checkbox.  Some names in the list might
have an underscore and number appended to their name.  Grouping parts and assemblies with the same
shape can be disabled with the option on the Spreadsheet tab.  In this case, parts with the same
shape will have an underscore and number appended to their name.  This might create a very long
list of part names.

Clicking on the model shows the part name in the upper left.  The part name shown may not be in the
list of assemblies and parts.  The part might be contained in a higher-level assembly that is in
the list.

Processing sketch geometry might also affect the list of names.  Some assemblies have no unique
names assigned to parts, therefore there is no list of part names.  See Help > Text Strings and
Numbers for how non-English characters in part names are supported.

Nested assemblies are also supported where one file contains the assembly structure with external
file references to individual assembly components that contain part geometry.

NOTE: Graphic PMI, supplemental geometry, capped surfaces for clipping planes, and cloud of points
on parts in an assembly is supported, however, it has not been thoroughly tested and might have the
wrong position and orientation.

Transparency for assemblies with AP242 tessellated geometry might look wrong.  In some rare cases,
parts in an assembly using tessellated geometry might have the wrong position and orientation or be
missing.

Assembly Structure is also supported by the AP242 Domain Model XML. See Websites > CAx Recommended Practices"
    .tnb select .tnb.status
  }

  $helpView add command -label "Supplemental Geometry" -command {
outputMsg "\nSupplemental Geometry -----------------------------------------------------------------------------" blue
outputMsg "Supplemental geometry is geometrical elements created in the CAD system that do not belong to the
manufactured part.  It is usually used to create other geometric shapes.  Supplemental geometry is
also known as construction, auxiliary, design, support, or reference geometry.

B-rep and tessellated geometry are supported for supplemental geometry. These types of supplemental
geometry and associated text are supported.  Colors defined in the STEP file override the default
colors below.

- Coordinate System: X axis red, Y axis green, Z axis blue
- Plane: blue transparent outlined surface (unbounded planes are with shown with a square surface)
- Cylinder: blue transparent cylinder
- Line/Circle/Ellipse: purple line/circle/ellipse (trimming with cartesian_point is not supported)
- Point: black dot
- Tessellated Surface: defined color

Supplemental geometry:
- can be optionally generated
- can be switched on and off
- is not associated with graphic PMI Saved Views
- in assemblies might have the wrong position and orientation (See Help > Assemblies)
- is counted on the PMI Coverage Analysis worksheet if a Viewer file is generated

See Websites > CAx Recommended Practices (Supplemental Geometry)"
    .tnb select .tnb.status
  }

  $helpView add command -label "Graphic PMI" -command {
outputMsg "\nGraphic PMI ---------------------------------------------------------------------------------------" blue
outputMsg "Graphic Presentation PMI annotations for geometric dimensioning and tolerancing composed of
polylines, lines, circles, and tessellated geometry are supported.  On the Generate tab, the color
of the annotations can be modified.  PMI associated with saved views can be switched on and off.

Some graphic PMI might not have equivalent or any semantic PMI in the STEP file.  Some STEP files
with semantic PMI might not have any graphic PMI.

Only graphic PMI defined in recommended practices is supported.  Older implementations of saved
view viewpoints might not conform to current recommended practices.
See Websites > CAx Recommended Practices
 (Representation and Presentation of PMI for AP242, PMI Polyline Presentation for AP203 and AP214)

Graphic PMI on parts in an assembly might have the wrong position and orientation.

See Help > User Guide (section 4.2)
See Help > Analyzer > Graphic Presentation PMI
See Examples > Viewer
See Examples > Sample STEP Files

---------------------------------------------------------------------------------------------------
Datum targets are shown only if a spreadsheet is generated with the Analyzer option for Semantic
PMI, and Part Geometry or Graphic PMI selected.

See Help > User Guide (section 4.2.2)
See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242, Sec. 6.6)"
    .tnb select .tnb.status
  }

  $helpView add command -label "PMI Placeholders" -command {
outputMsg "\nPMI Placeholders ----------------------------------------------------------------------------------" blue
outputMsg "PMI (annotation) placeholders provide information about the position, orientation, and organization
of an annotation without the graphic presentation of numeric values, symbols, and text for
geometric or dimensional tolerances.  Placeholders are associated with saved views if there is also
graphic PMI.  Placeholders are supported in AP242 editions >=2.  They are not documented in the
User Guide.

Placeholder coordinate systems are shown with an axes triad, gray sphere, and text label with the
name of the placeholder.  Leader lines and a rectangle for the annotation are shown with black
lines.  To identify which annotation a leader line is associated with, the first and last points of
a leader line have a text label.  Leader line symbols show their type and position with blue text.
Some implemetations of placeholders might not contain all of the graphic elements.

See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242, Sec. 7.2)"
    .tnb select .tnb.status
  }

  $helpView add command -label "Hole Features" -command {
outputMsg "\nHole Features -------------------------------------------------------------------------------------" blue
outputMsg "The position, orientation, dimensions and tolerances for drilled, counterbore, and countersink hole
features are supported.  Hole features in STEP are not the same as explicitly modeling holes in
part geometry.  Hole features can be used without explicitly modeling holes.

The Analyzer report for Semantic PMI must be generated to show hole features in the Viewer.
Cylindrical and conical surfaces are used to show the depth and diameter of the drilled hole,
counterbore, and countersink.  These surfaces are derived from the hole feature dimensions and not
the part geometry.  Countersink holes are cyan and other holes are green, and can be switched on
and off in the Viewer.  If no hole depth is specified (through hole), then only a short cylindrical
surface with the correct diameter is shown.  Flat and conical hole bottoms are supported if it is
not a through hole.  A text label is shown for each hole.  If the report for Semantic PMI is not
generated, then only the text label is shown.

In the Entity Types section on the Generate tab, Features is automatically selected when hole
feature entities are in the STEP file.  Semantic information related to holes is reported on
*_hole_definition and basic_round_hole worksheets.  Hole feature dimensions are not the same as
semantic PMI using dimensional_size and dimensional_location.  Hole features are supported in
AP242 editions > 1, but have not been widely implemented.

See Help > User Guide (section 4.2.3)"
    .tnb select .tnb.status
  }

  $helpView add command -label "AP209 Finite Element Model" -command {
outputMsg "\nAP209 Finite Element Model ------------------------------------------------------------------------" blue
outputMsg "All AP209 entities are always processed and written to a spreadsheet unless a User-defined list is
used.

The AP209 finite element model composed of nodes, mesh, elements, boundary conditions, loads, and
displacements are shown and can be toggled on and off in the viewer.  Internal faces for solid
elements are not supported.

Nodal loads and element surface pressures are shown.  Load vectors are colored by their magnitude.
The length of load vectors can be scaled by their magnitude.  Forces use a single-headed arrow.
Moments use a double-headed arrow.

Displacement vectors are colored by their magnitude.  The length of displacement vectors can be
scaled by their magnitude depending on if they have a tail.  The finite element mesh is not
deformed.

Boundary conditions for translation DOF are shown with a red, green, or blue line along the X, Y,
or Z axes depending on the constrained DOF.  Boundary conditions for rotation DOF are shown with a
red, green, or blue circle around the X, Y, or Z axes depending on the constrained DOF.  A gray box
is used when all six DOF are constrained.  A gray pyramid is used when all three translation DOF
are constrained.  A gray sphere is used when all three rotation DOF are constrained.

Stresses, strains, and multiple coordinate systems are not supported.

Setting Maximum Rows (More tab) does not affect the view.  For large AP209 files, there might be
insufficient memory to process all of the elements, loads, displacements, and boundary conditions.

See Help > User Guide (section 4.4)
See Examples > Viewer
See Websites > STEP > AP209 FEA"
    .tnb select .tnb.status
  }

  $Help add cascade -label "Analyzer" -menu $Help.0
  set helpAnalyze [menu $Help.0 -tearoff 1]

# analyzer overview
  $helpAnalyze add command -label "Overview" -command {
outputMsg "\nAnalyzer Overview ---------------------------------------------------------------------------------" blue
outputMsg "The Analyzer reports information related to validation properties, semantic PMI, and graphic PMI,
and checks them for conformance to recommended practices.  Syntax Errors are reported for
nonconformance.  Entities that report this information are highlighted on the File Summary
worksheet.

Inverse relationships and Backwards References show the relationship between some entities through
other entities.

PMI Coverage Analysis shows the distribution of specific semantic PMI elements related to geometric
dimensioning and tolerancing.

If a STEP AP242 file is processed that is generated from one of the NIST CAD models, the semantic
PMI Analyzer report is color-coded by the expected PMI.

See Help > Analyzer for other topics
See Help > User Guide (section 6)
See Websites > CAx Recommended Practices
See Examples > NIST CAD Models"
    .tnb select .tnb.status
  }

# validation properties, PMI, conformance checking help
  $helpAnalyze add command -label "Validation Properties" -command {
outputMsg "\nValidation Properties -----------------------------------------------------------------------------" blue
outputMsg "Geometric, assembly, PMI, annotation, attribute, tessellated, composite, and FEA validation
properties are reported.  The property values are reported in columns highlighted in yellow and
green on the property_definition worksheet.  The worksheet can also be sorted and filtered.  All
properties might not be shown depending on the Maximum Rows set on the More tab.

The name or description attribute of the entity referred to by the property_definition definition
attribute is included in brackets.

Validation properties are also reported on their associated annotation, dimension, geometric
tolerance, and shape aspect entities.  The report includes the validation property name and names
of the properties.  Some properties are reported only if the Semantic PMI Analyzer report is
selected.  Other properties and user defined attributes are also reported.

Another type of validation property is known as Semantic Text where explicit text strings in the
STEP file can be associated with part surfaces, similar to semantic PMI.  The semantic text will
appear in the spreadsheet on shape_aspect and other related entities.  The shape aspects can be
related to their corresponding dimensional or geometric tolerance entities.  A message will
indicate when semantic text is added to entities.

The PMI validation property Equivalent Unicode String is shown on worksheets for semantic and
graphic PMI with that validation property.  The sampling points for the Cloud of Points validation
property are shown in the viewer.  See Help > Viewer > New Features.  Neither of these features are
documented in the User Guide.

Syntax errors related to validation property attribute values are also reported in the Status tab
and the relevant worksheet cells.  Syntax errors are highlighted in red.  See Help > Analyzer > Syntax Errors

Clicking on the plus '+' symbols above the columns shows other columns that contain the entity ID
and attribute name of the validation property value.  All of the other columns can be shown or
hidden by clicking the '1' or '2' in the upper right corner of the spreadsheet.

The Summary worksheet indicates if properties are reported on property_definition and other
entities.

See Help > User Guide (section 6.3)
See Examples > Graphic PMI, Validation Properties

Validation properties must conform to recommended practices.
 See Websites > CAx Recommended Practices (Geometric and Assembly Validation Properties,
  User Defined Attributes, Representation and Presentation of PMI for AP242,
  Tessellated 3D Geometry)"
    .tnb select .tnb.status
  }

  $helpAnalyze add command -label "Semantic Representation PMI" -command {
outputMsg "\nSemantic Representation PMI -----------------------------------------------------------------------" blue
outputMsg "Semantic Representation PMI includes all information necessary to represent geometric and
dimensional tolerances (GD&T) without any graphic presentation elements.  Semantic PMI is
associated with CAD model geometry and is computer-interpretable to facilitate automated
consumption by downstream applications for manufacturing, measurement, inspection, and other
processes.  Semantic Representation PMI is mainly in AP242 files.

Worksheets for the Semantic Representation PMI Analyzer report show a visual recreation of the
representation for Dimensional Tolerances, Geometric Tolerances, and Datum Features.  The results
are in columns, highlighted in yellow and green, on the relevant worksheets.  The GD&T is recreated
as best as possible given the constraints of Excel.

All of the visual recreation of Datum Systems, Dimensional Tolerances, and Geometric Tolerances
that are reported on individual worksheets are collected on one Semantic PMI Summary worksheet.

If STEP files from the NIST CAD models (Examples > NIST CAD Models) are processed, then the
Semantic PMI Summary is color-coded by the expected PMI in each CAD model.
See Help > Analyzer > NIST CAD Models

Datum Features are reported on datum_* entities.  Datum_system will show the complete Datum
Reference Frame.  Datum Targets are reported on placed_datum_target_feature.

Dimensional Tolerances are reported on the dimensional_characteristic_representation worksheet.
The dimension name, representation name, length/angle, length/angle name, plus minus bounds, and
other associated information are reported.  The relevant section in the recommended practice is
shown in the column headings.

Geometric Tolerances are reported on *_tolerance entities by showing the complete Feature Control
Frame (FCF), and possible Dimensional Tolerance and Datum Feature.  The FCF should contain the
geometry tool, tolerance zone, datum reference frame, and associated modifiers.

---------------------------------------------------------------------------------------------------
If a Dimensional Tolerance refers to the same geometric element as a Geometric Tolerance, then it
will be shown above the FCF.  If a Datum Feature refers to the same geometric face as a Geometric
Tolerance, then it is shown below the FCF.  If an expected Dimensional Tolerance is not shown above
a Geometric Tolerance, then the tolerances do not reference the same geometric element.  For
example, the dimensional tolerance referencing the curved edge of a hole and the geometric
tolerance referencing the cylindrical surfaces of the same hole.

The association of the Datum Feature with a Geometric Tolerance is based on each referring to the
same geometric element.  However, the Graphic PMI might show the Geometric Tolerance and Datum
Feature as two separate annotations with leader lines attached to the same geometric element.

The number of decimal places for dimension and geometric tolerance values can be specified in the
STEP file.  By definition the value is always truncated, however, the values can be rounded instead.
For example with the value 0.5625, the qualifier 'NR2 1.3' will truncate it to 0.562  Rounding will
show 0.563  Rounding values might result in a better match to graphic PMI shown in the viewer or to
expected PMI in the NIST CAD models.  See More tab > Analyzer > Round ...

Some syntax errors that indicate non-conformance to a CAx-IF Recommended Practices related to PMI
Representation are also reported in the Status tab and the relevant worksheet cells.  Syntax errors
are highlighted in red.  See Help > Analyzer > Syntax Errors

A Semantic PMI Coverage Analysis worksheet is also generated.

See Help > User Guide (section 6.1)
See Help > Analyzer > PMI Coverage Analysis
See Examples > Spreadsheets - Semantic PMI
See Examples > Sample STEP Files

Semantic PMI must conform to recommended practices.
 See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242)"
    .tnb select .tnb.status
  }

  $helpAnalyze add command -label "Graphic Presentation PMI" -command {
outputMsg "\Graphic Presentation PMI ---------------------------------------------------------------------------" blue
outputMsg "Graphic Presentation PMI consists of geometric elements including lines and curves preserving the
exact appearance (color, shape, positioning) of the geometric and dimensional tolerance (GD&T)
annotations.  Graphic PMI is not intended to be computer-interpretable and does not have any
representation information, although it can be linked to its corresponding Semantic PMI.

The Analyzer report for Graphic PMI supports annotation_curve_occurrence, annotation_curve,
annotation_fill_area_occurrence, and tessellated_annotation_occurrence entities.  Geometric
entities used for Graphic PMI annotations are reported in columns, highlighted in yellow and green,
on those worksheets.  Presentation Style, Saved Views, Validation Properties, Annotation Plane,
Associated Geometry, and Associated Semantic PMI are also reported.

The Summary worksheet indicates on which worksheets Graphic PMI is reported.  Some syntax errors
related to Graphic PMI are also reported in the Status tab and the relevant worksheet
cells.  Syntax errors are highlighted in red.  See Help > Analyzer > Syntax Errors

A Graphic PMI Coverage Analysis worksheet is also generated.

See Help > User Guide (section 6.2)
See Help > Viewer > Graphic PMI
See Help > Analyzer > PMI Coverage Analysis
See Examples > Viewer
See Examples > Graphic PMI, Validation Properties
See Examples > Sample STEP Files

Graphic PMI must conform to recommended practices.
 See Websites > CAx Recommended Practices (Representation and Presentation of PMI for AP242,
  PMI Polyline Presentation for AP203 and AP214)"
    .tnb select .tnb.status
  }

# coverage analysis help
  $helpAnalyze add command -label "PMI Coverage Analysis" -command {
outputMsg "\nPMI Coverage Analysis -----------------------------------------------------------------------------" blue
outputMsg "PMI Coverage Analysis worksheets are generated when processing single or multiple files and when
reports for Semantic Representation PMI or Graphic Presentation PMI are selected.

Semantic PMI Coverage Analysis counts the number of PMI Elements in a STEP file for tolerances,
dimensions, datums, modifiers, and CAx-IF Recommended Practices for PMI Representation.  On the
Coverage Analysis worksheet, some PMI Elements show their associated symbol, while others show the
relevant section in the recommended practice.  PMI Elements without a section number do not have a
recommended practice for their implementation.  The PMI Elements are grouped by features related
tolerances, tolerance zones, dimensions, dimension modifiers, datums, datum targets, and other
modifiers.  The number of some modifiers, e.g., maximum material condition, does not differentiate
whether they appear in the tolerance zone definition or datum reference frame.  Rows with no count
of a PMI Element can be shown, see More tab.

Some PMI Elements might not be exported to a STEP file by your CAD system.  Some PMI Elements are
only in AP242 editions > 1.

If STEP files from the NIST CAD models (Examples > NIST CAD Models) are processed, then the
Semantic PMI Coverage Analysis worksheet is color-coded by the expected number of PMI elements in
each CAD model.  See Help > Analyzer > NIST CAD Models

The Graphic PMI Coverage Analysis counts the occurrences of the recommended name attribute defined
in the CAx-IF Recommended Practice for PMI Representation and Presentation of PMI (AP242).  The
name attribute is associated with the graphic elements used to draw a PMI annotation or placeholder.
There is no semantic PMI meaning to the name attributes.

See Help > Analyzer > Semantic Representation PMI
See Help > User Guide (sections 6.1.7 and 6.2.1)
See Examples > PMI Coverage Analysis"
    .tnb select .tnb.status
  }

  $helpAnalyze add command -label "Syntax Errors" -command {
outputMsg "\nSyntax Errors -------------------------------------------------------------------------------------" blue
outputMsg "Syntax Errors are generated when an Analyzer option related to Semantic PMI, Graphic PMI, and
Validation Properties is selected.  The errors refer to specific sections, figures, or tables in
the relevant CAx-IF Recommended Practice.  Errors should be fixed so that the STEP file can
interoperate with other CAx software and conform to recommended practices.
See Websites > CAx Recommended Practices

Syntax errors are highlighted in red in the Status tab.  Other informative warnings are highlighted
in yellow.  Syntax errors that refer to section, figure, and table numbers might use numbers that
are in a newer version of a recommended practice that has not been publicly released.

Some syntax errors use abbreviations for STEP entities:
 GISU - geometric_item_specific_usage
 IIRU - identified_item_representation_usage

On the Summary worksheet in column A, most entities that have syntax errors are colored gray.  A
comment indicating that there are errors is also shown with a small red triangle in the upper right
corner of a cell in column A.

On an entity worksheet, most syntax errors are highlighted in red and have a cell comment with the
text of the syntax error that was shown in the Status tab.  Syntax errors are highlighted by *** in
the log file.

The Inverse Relationships option in the Analyzer section might be useful to debug Syntax Errors.

NOTE - Syntax Errors related to CAx-IF Recommended Practices are unrelated to errors detected with
the Syntax Checker.  See Help > Syntax Checker

See Help > User Guide (section 6.5)"
    .tnb select .tnb.status
  }

# NIST CAD model help
  $helpAnalyze add command -label "NIST CAD Models" -command {
outputMsg "\nNIST CAD Models -----------------------------------------------------------------------------------" blue
outputMsg "If a STEP file from a NIST CAD model (CTC/FTC/STC) is processed, then the PMI in the STEP file is
automatically checked against the expected PMI in the corresponding NIST test case.  The Semantic
PMI Coverage and Summary worksheets are color-coded by the expected PMI in each NIST test case.

The color-coding only works if the STEP file name can be recognized as having been generated from
one of the NIST CAD models.  For example, nist_ctc_02-some-text.stp would recognize the STEP file
as being generated from CTC 2.  nist_ftc_06-more-text.stp would be for FTC 6 and
nist_stc_10-whatever.stp would be for STC 10.

---------------------------------------------------------------------------------------------------
* Semantic PMI Summary *
This worksheet is color-coded by the Expected PMI annotations in a test case drawing.
- Green is an Exact match to an expected PMI annotation in the test case drawing
- Green (lighter shade) is an Exact match with Exceptions
- Cyan is a Partial match
- Yellow is a Possible match
- Red is No match

The following Exceptions are ignored when considering an Exact match:
- repetitive dimensions 'nX'
- different, missing, or unexpected dimensional tolerances in a Feature Control Frame (FCF)
- some datum features associated with geometric tolerances
- some modifiers in an FCF
- all around symbol

Some causes of Partial and Possible matches are, missing or wrong:
- diameter and radius symbols
- numeric values for dimensions and tolerances
- datum features and datum reference frames
- modifiers for dimensions, tolerance zones, and datum reference frames
- composite tolerances

On the Summary worksheet the column for Similar PMI and Exceptions shows the most closely matching
Expected PMI for Partial and Possible matches and the reason for an Exact match with Exceptions.

Trailing and leading zeros are ignored when matching a PMI annotation.  Matches also only consider
the current capabilities of PMI annotations in STEP AP242 and CAx-IF Recommended Practices.  PMI
annotations for hole features including counterbore, countersink, and depth are also supported.

---------------------------------------------------------------------------------------------------
* Graphic PMI Coverage Analysis *
This worksheet is color-coded by the expected number of PMI elements in a test case drawing.  The
expected results were determined by manually counting the number of PMI elements in each drawing.
Counting of some modifiers, e.g., maximum material condition, does not differentiate whether they
appear in the tolerance zone definition or datum reference frame.
- A green cell is a match to the expected number of PMI elements. (3/3)
- Yellow, orange, and yellow-green means that less were found than expected. (2/3)
- Red means that no instances of an expected PMI element were found. (0/3)
- Cyan means that more were found than expected. (4/3)
- Magenta means that some PMI elements were found when none were expected. (3/0)

---------------------------------------------------------------------------------------------------
From the Semantic PMI Summary results, color-coded percentages of Exact, Partial, Possible and
Missing matches is shown in a table below the Semantic PMI Coverage Analysis.  Exceptions are
counted as an Exact match and do not affect the percentage, except one or two points are deducted
when the percentage would be 100.

The Total PMI on which the percentages are based on is also shown.  Coverage Analysis is only based
on individual PMI elements.  The Semantic PMI Summary is based on the entire PMI feature control
frame and provides a better understanding of the PMI.  The Coverage Analysis might show that there
is an exact match for all of the PMI elements, however, the Semantic PMI Summary might show less
than exact matches.

---------------------------------------------------------------------------------------------------
* Missing PMI *
Missing PMI refers to semantic PMI and does not imply that the corresponding graphic annotation is
missing.  Missing PMI annotations on the Summary worksheet or PMI elements on the Coverage
worksheet might mean that the CAD system or translator:
- PMI annotation defined in a NIST test case is not in the CAD model
- did not follow CAx-IF Recommended Practices for PMI (See Websites > CAx Recommended Practices)
- has not implemented exporting a PMI element to a STEP file
- mapped an internal PMI element to the wrong STEP PMI element

* User-defined Expected PMI *
A user-defined file of expected PMI can also be used.  The file must be named SFA-EPMI-yourstepfilename.xlsx
Contact the developer for more information.  This feature is not documented in the User Guide.

NOTE - Some of the NIST test cases have complex PMI annotations that are not commonly used.  There
might be ambiguities in counting the number of PMI elements.

See Help > User Guide (section 6.6)
See Examples > NIST CAD Models
See Examples > Spreadsheets - Semantic PMI"
    .tnb select .tnb.status
  }
  $Help add separator

  $Help add command -label "Syntax Checker" -command {
outputMsg "\nSyntax Checker ------------------------------------------------------------------------------------" blue
outputMsg "The Syntax Checker checks for basic syntax errors and warnings in the STEP file related to missing
or extra attributes, incompatible and unresolved entity references, select value types, illegal and
unexpected characters, and other problems with entity attributes.  Some errors might prevent this
software and others from processing a STEP file.  Characters that are identified as illegal or
unexpected might not be shown in a spreadsheet or in the viewer.  See Help > Text Strings and Numbers

Entities in the STEP file that are not in the STEP AP schema are reported as ignored entities. They
will not appear in the spreadsheet.  Attributes on other entities that refer to ignored entities
will be blank.  Analyzer reports and the Viewer might be affected.

If errors and warnings are reported, the number in parentheses is the line number in the STEP file
where the error or warning was detected.  There should not be any of these types of syntax errors
in a STEP file.  Errors should be fixed to ensure that the STEP file conforms to the STEP schema
and can interoperate with other software.

There are other validation rules defined by STEP schemas (where, uniqueness, and global rules,
inverses, derived attributes, and aggregates) that are NOT checked.  Conforming to the validation
rules is also important for interoperability with STEP files.  See Websites > STEP

---------------------------------------------------------------------------------------------------
The Syntax Checker can be run with function key F8 or when a Spreadsheet or View is generated.  All
entities are always checked and is not affected by the Maximum Rows setting.  The Status tab might
be grayed out when the Syntax Checker is running.

Syntax checker results appear in the Status tab.  If the Log File option is selected, the results
are also written to a log file (myfile-sfa-err.log).  The syntax checker errors and warnings are
not reported in the spreadsheet.

The Syntax Checker can also be run from the command-line version with the command-line argument
'syntax'.  For example: sfa-cl.exe myfile.stp syntax

The Syntax Checker works with any supported schema.  See Help > Supported STEP APs and
Help > Large STEP files

The Viewer might indicate that there are possible syntax errors and to run the Syntax Checker.

NOTE - Syntax Checker errors and warnings are unrelated to those detected when CAx-IF Recommended
Practices are checked with one of the Analyzer options.  See Help > Analyzer > Syntax Errors"
    .tnb select .tnb.status
  }

  $Help add command -label "Bill of Materials" -command {
outputMsg "\nBill of Materials ---------------------------------------------------------------------------------" blue
outputMsg "Select BOM on the Generate tab to generate a Bill of Materials.  The next_assembly_usage_occurrence
entity shows the assembly and component names for the relating and related products in an assembly.
If there are no next_assembly_usage_occurrence entities, then the Bill of Materials (BOM) cannot
be generated.

The BOM worksheet (third worksheet) lists the quantities of parts and assemblies in two tables.
Assemblies also show their components which can be parts or other assemblies.  A STEP file might
not contain all the necessary information to generate a complete BOM.  Parts do not have to be
contained in an assembly, therefore some BOMs will not have a list of assemblies and some parts
might not be listed as a component of an assembly.  See Examples > Bill of Materials

Generate the Analyzer report for Validation Properties to see possible properties associated with
Parts.

Bill of Materials are not documented in the User Guide.  See Examples > Bill of Materials"
    .tnb select .tnb.status
  }
  $Help add separator

# open Function Keys help
  $Help add command -label "Function Keys" -command {
outputMsg "\nFunction Keys -------------------------------------------------------------------------------------" blue
outputMsg "Function keys can be used as shortcuts for several commands:

F1 - Generate Spreadsheet and/or run the Viewer with the current or last STEP file
F2 - Open current or last Spreadsheet
F3 - Open current or last Viewer file in web browser
F4 - Open Log file
F5 - Open STEP file in a text editor  (See Help > Open STEP File in App)
Shift-F5 - Open STEP file directory

F6 - Generate Spreadsheets and/or run the Viewer with the current or last set of multiple STEP files
F7 - Open current or last File Summary Spreadsheet generated from a set of multiple STEP files

F8 - Run the Syntax Checker (See Help > Syntax Checker)

F9  - Decrease this font size (also ctrl -)
F10 - Increase this font size (also ctrl +)

F12 - Open Viewer file in text editor

For F1, F2, F3, F6, and F7 the last STEP file, Spreadsheet, and Viewer file are remembered between
sessions.  In other words, F1 can process the last STEP file from a previous session without having
to select the file.  F2 and F3 function similarly for Spreadsheets and the Viewer."
    .tnb select .tnb.status
  }

  $Help add command -label "Supported STEP APs" -command {
outputMsg "\nSupported STEP APs --------------------------------------------------------------------------------" blue
outputMsg "These STEP Application Protocols (AP) and other schemas are supported for generating spreadsheets.
The Viewer works with STEP AP242, AP203, AP214, AP238, and AP209.  See Websites > STEP
AP238 STEP-NC files (.stpnc) are supported by renaming the file extension to '.stp'.

The name of the AP is on the FILE_SCHEMA entity in the HEADER section of a STEP file.  The 'e1'
notation below, after an AP number, refers to an older Edition of that AP.  Some APs have multiple
editions with the same name.  AP242 editions 1-4 were released in 2014, 2020, 2022, and 2025.
See Websites > AP242\n"

    set schemas {}
    set ifcschemas {}
    foreach match [lsort [glob -nocomplain -directory $ifcsvrDir *.rose]] {
      set schema [string toupper [file rootname [file tail $match]]]
      if {[string first "HEADER_SECTION" $schema] == -1 && [string first "KEYSTONE" $schema] == -1 && \
          [string first "ENGINEERING_MIM_LF-OLD" $schema] == -1 && [string range $schema end-2 end] != "MIM"} {
        if {[info exists stepAPs($schema)] && $schema != "STRUCTURAL_FRAME_SCHEMA"} {
          if {[string first "CONFIGURATION" $schema] != 0} {
            set str $stepAPs($schema)
            if {[string first "e1" $str] == -1} {append str "  "}
            lappend schemas "$str - $schema"
          } else {
            lappend schemas $schema
          }
        } elseif {[string first "AP2" $schema] == 0} {
          lappend schemas "[string range $schema 0 4]   - $schema"
        } elseif {[string first "IFC" $schema] == -1} {
          lappend schemas $schema
        } elseif {$schema == "IFC2X3" || [string first "IFC4" $schema] == 0 || [string first "IFC5" $schema] == 0} {
          lappend ifcschemas [string range $schema 3 end]
        }
      }
    }
    if {[llength $ifcschemas] > 0} {lappend schemas "IFC ($ifcschemas)"}

    if {[llength $schemas] <= 1} {
      errorMsg "No Supported STEP APs were found."
      if {[llength $schemas] == 1} {errorMsg "- Manually uninstall the existing IFCsvrR300 ActiveX Component 'App'."}
      errorMsg "- Restart this software to install the new IFCsvr toolkit."
    }

    set n 0
    foreach item [lsort $schemas] {
      set c1 [string first "-" $item]
      if {$c1 == -1} {
        if {$n == 0} {
          incr n
          outputMsg "\nOther Schemas"
        }
        set txt [string toupper $item]
        if {$txt == "CUTTING_TOOL_SCHEMA_ARM"} {append txt " (ISO 13399)"}
        if {[string first "ISO13584_25" $txt] == 0} {append txt " (Supplier library)"}
        if {[string first "ISO13584_42" $txt] == 0} {append txt " (Parts library)"}
        if {$txt == "STRUCTURAL_FRAME_SCHEMA"} {append txt " (CIS/2)"}
        outputMsg "  $txt"
      } else {
        set txt "[string range $item 0 $c1][string toupper [string range $item $c1+1 end]]"
        outputMsg "  $txt"
      }
    }
    .tnb select .tnb.status
  }

  $Help add command -label "Text Strings and Numbers" -command {
outputMsg "\nText Strings and Numbers --------------------------------------------------------------------------" blue
outputMsg "Text strings in STEP files might use non-English characters or symbols.  Some examples are accented
characters in European languages (for example é), and Asian languages that use different characters
sets such as Japanese or Cyrillic.  Text strings with non-English characters or symbols are usually
on descriptive measure or product related entities with name, description, or id attributes.

According to ISO 10303 Part 21 section 6.4.3, Unicode can be used for non-English characters and
symbols with the control directives \\X2\\ and \\X0\\.  For example, \\X2\\00E9\\X0\\ is used for the
accented character é.  Some CAD software does not support these control directives when exporting
or importing a STEP file.

---------------------------------------------------------------------------------------------------
Spreadsheet - Use the option on the More tab to support non-English characters using the \\X2\\
control directive.  In some cases the option will be automatically selected based on the file size
or schema.  There is a warning message if \\X2\\ is detected in the STEP file and the option is not
selected.  In this case the \\X2\\ characters are ignored and will be missing in the spreadsheet.
Non-English characters that do not use the control directives might be missing in the spreadsheet.
Control directives are supported only if Excel is installed.

Unicode characters for GD&T symbols are used by Equivalent Unicode Strings reported on the
descriptive_representation_item worksheet and worksheets for semantic and graphic PMI where there
is an associated PMI validation property.  Equivalent Unicode Strings are not documented in the
User Guide.  See the Recommended Practice for PMI Unicode String Specification.

Unicode characters using \\X2\\ control directives are not processed for attribute strings on
Geometry entities and some Presentation entities.  The \\X\\ and \\S\\ control directives are supported
by default although they are no longer implemented in STEP files.

---------------------------------------------------------------------------------------------------
Viewer - All control directives are supported for part and assembly names.  Non-English characters
that do not use the control directives might be shown with the wrong characters.

Support for non-English characters that do not use the control directives can be improved by
converting the encoding of the STEP file to UTF-8 with the Notepad++ text editor or other software.
Regardless, some non-English characters might cause a crash or prevent the viewer from running.
See Help > Crash Recovery

The Syntax Checker identifies non-English characters as 'illegal characters'.  You should test your
CAD software to see if it supports non-English characters or control directives.

---------------------------------------------------------------------------------------------------
Numbers in a STEP file use a period '.' as the decimal separator.  Some non-English language
versions of Excel use a comma ',' as a decimal separator.  This might cause some real numbers to be
formatted as a date in a spreadsheet.  For example, the number 1.5 might appear as 1-Mai.

To check if the formatting is a problem, process the STEP file nist_ctc_05.stp included with the
SFA zip file and select the Geometry Entity Type category.  Check the 'radius' attribute on the
resulting 'circle' worksheet.

To change the formatting in Excel, go to the Excel File menu > Options > Advanced.  Uncheck
'Use system separators' and change 'Decimal separator' to a period . and 'Thousands separator' to a
comma ,

This change applies to ALL Excel spreadsheets on your computer.  Change the separators back to
their original values when finished.  You can always check the STEP file to see the actual value of
the number.

See Help > User Guide (section 5.5.2)"
    .tnb select .tnb.status
  }

# open STEP files help
  $Help add command -label "Open STEP File in App" -command {
outputMsg "\nOpen STEP File in App -----------------------------------------------------------------------------" blue
outputMsg "STEP files can be opened in other apps.  If apps are installed in their default directory, then the
pull-down menu on the Generate tab will contain apps including STEP viewers and browsers.

The 'Tree View (for debugging)' option rearranges and indents the entities to show the hierarchy of
information in a STEP file.  The 'tree view' file (myfile-sfa.txt) is written to the same directory
as the STEP file or to the same user-defined directory specified in the More tab.  It is useful for
debugging STEP files but is not recommended for large STEP files.

The 'Default STEP Viewer' option opens the STEP file in whatever app is associated with STEP files.
A text editor always appear in the menu.  Use F5 to open the STEP file in the text editor.

See Help > User Guide (section 3.4.5)"
    .tnb select .tnb.status
  }

# multiple files help
  $Help add command -label "Multiple STEP Files" -command {
outputMsg "\nMultiple STEP Files -------------------------------------------------------------------------------" blue
outputMsg "Multiple STEP files can be selected in the Open File(s) dialog by holding down the control or shift
key when selecting files or an entire directory of STEP files can be selected with 'Open Multiple
STEP Files in a Directory'.  Files in subdirectories of the selected directory can also be
processed.

When processing multiple STEP files, a File Summary spreadsheet is generated in addition to
individual spreadsheets for each file.  The File Summary spreadsheet shows the entity count and
totals for all STEP files. The File Summary spreadsheet also links to the individual spreadsheets
and STEP files.

If only the File Summary spreadsheet is needed, it can be generated faster by deselecting most
Entity Types and options on the Generate tab.

If the reports for Semantic or Graphic PMI are selected, then Coverage Analysis worksheets are also
generated.

In some rare cases an error will be reported with an entity when processing multiple files that is
not an error when processing it as a single file.  Reporting the error is a bug.

See Help > User Guide (section 8)
See Examples > PMI Coverage Analysis"
    .tnb select .tnb.status
  }

# large files help
  $Help add command -label "Large STEP Files" -command {
outputMsg "\nLarge STEP Files ----------------------------------------------------------------------------------" blue
outputMsg "The largest STEP file that can be processed for a Spreadsheet or the Syntax Checker is
approximately 430 MB.  Processing a larger STEP file might cause a crash.  A popup dialog might
appear that says 'unable to realloc xxx bytes'.  See Help > Crash Recovery

Some workarounds are available to (1) prevent a crash with a smaller file that can be processed,
(2) to reduce the processing time, or (3) to reduce the size of the spreadsheet:
- Deselect Entity Types that might not need to be processed such as Geometry and Coordinates
- Use a User-Defined List of only the required entity types
- Use a smaller value for the Maximum Rows
- Deselect Analyzer options and Inverse Relationships

The Status tab might be grayed out when a large STEP file is being read.

---------------------------------------------------------------------------------------------------
To use only the Viewer with a large STEP file, select View and Part Only.  A 1.5 GB STEP file has
been successfully tested this way in the Viewer."
    .tnb select .tnb.status
  }

  $Help add command -label "Crash Recovery" -command {
outputMsg "\nCrash Recovery ------------------------------------------------------------------------------------" blue
outputMsg "Sometimes this software crashes after a STEP file has been successfully opened and the processing
of entities has started.  Popup dialogs might appear that say 'Runtime Error!' or 'ActiveState
Basekit has stopped working'.

A crash is most likely due to syntax errors in the STEP file, a very large STEP file, the options
selected for the software, or due to limitations of the toolkit used to read STEP files.  Run the
Syntax Checker with function key F8 or the option on the Generate tab to check for errors with
entities that might have caused the crash.  See Help > Syntax Checker

The software keeps track of the last entity processed when it crashed.  The list of bad entity
types is stored in myfile-skip.dat.  Simply restart this software and use F1 to process the last
STEP file or use F6 if processing multiple files.  The entities listed in myfile-skip.dat that
caused the crash will be skipped.

When the STEP file is processed again, the list of specific entities that are not processed is
reported.  If syntax errors related to the bad entities are corrected, then delete or edit the
*-skip.dat file so that the corrected entities are processed.

There are two other workarounds.  On the Generate tab deselect the Entity Type that caused the
error.  However, this will prevent processing of other entities that do not cause a crash.  A
User-Defined List can be used to process only the required entities for a spreadsheet.

For errors similar to 'unable to realloc xxx bytes', see Help > Large STEP Files

See Help > User Guide (section 2.4)

---------------------------------------------------------------------------------------------------
If the software crashes the first time you run it, there might be a problem with the installation
of the IFCsvr toolkit.  First uninstall the IFCsvr toolkit.  Then run SFA as Administrator and when
prompted, install the IFCsvr toolkit for Everyone, not Just Me.  Subsequently, SFA does not have to
be run as Administrator.

If that does not work, then an environment variable might need to be set.  From the Windows menu,
search for ‘Edit the system environment variables’.  On the Advanced tab, select Environment
Variables.  Then create a new System Variable and set ROSE_SCHEMAS to C:\\Program Files (x86)\\IFCsvrR300\\dll
You might need administrator privileges."
    .tnb select .tnb.status
  }

  $Help add separator
  $Help add command -label "Disclaimers" -command {
outputMsg "\nDisclaimers ---------------------------------------------------------------------------------------" blue
outputMsg "Please see Help > NIST Disclaimer for the Software Disclaimer.

The Examples menu provides links to several sources of STEP files.  This software and others might
indicate that there are errors in some of the STEP files.  NIST assumes no responsibility
whatsoever for the use of the STEP files by other parties, and makes no guarantees, expressed or
implied, about their quality, reliability, or any other characteristic.

Any mention of commercial products or references to web pages is for information purposes only; it
does not imply recommendation or endorsement by NIST.  For any of the web links, NIST does not
necessarily endorse the views expressed, or concur with the facts presented on those web sites.

This software uses IFCsvr, Microsoft Excel, and software based on Open Cascade that are covered by
their own Software License Agreements.

If you are using this software in your own application, please explicitly acknowledge NIST as the
source of the software."
  .tnb select .tnb.status
  }

  $Help add command -label "NIST Disclaimer" -command {openURL https://www.nist.gov/disclaimer}
  $Help add command -label "About" -command {
    outputMsg "\nSTEP File Analyzer and Viewer ---------------------------------------------------------------------" blue
    set sysvar "Version: [getVersion] ([string trim [clock format $progtime -format "%e %b %Y"]])"
    catch {append sysvar ", IFCsvr [registry get $ifcsvrVer {DisplayVersion}]"}
    catch {append sysvar ", stp2x3d ([string trim [clock format [file mtime [file join $mytemp stp2x3d-part.exe]] -format "%e %b %Y"]])"}
    append sysvar "\nFiles processed: $filesProcessed"
    outputMsg $sysvar
    outputMsg "\nThis software was first released in April 2012 and developed in the NIST Engineering Laboratory.

Credits
- Reading and parsing STEP files
   IFCsvr ActiveX Component, Copyright \u00A9 1999, 2005 SECOM Co., Ltd. All Rights Reserved
   IFCsvr has been modified by NIST to include STEP schemas
   The license agreement is in C:\\Program Files (x86)\\IFCsvrR300\\doc
- Viewer for b-rep and tessellated part geometry
   STEP to X3D Translator (stp2x3d) developed by Soonjo Kwon, former NIST Associate
   See Websites > STEP
- Some Tcl code is based on CAWT https://www.tcl3d.org/cawt/

See Help > Disclaimers and NIST Disclaimer"

    set winver ""
    if {[catch {
      set winver [registry get {HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion} {ProductName}]
    } emsg]} {
      set winver "Windows [expr {int($tcl_platform(osVersion))}]"
    }
    catch {
      set build [string range [lindex [split [::twapi::get_os_description] " "] end] 0 end-1]
      if {[string is integer $build] && $build >=22000} {regsub "10" $winver "11" winver}
    }
    if {[string first "Server" $winver] != -1 || $tcl_platform(osVersion) < 6.1} {
      outputMsg "\n$winver is not supported." red
    } elseif {$tcl_platform(osVersion) < 10.0} {
      outputMsg "\nThis software is no longer tested in $winver." red
    }

# debug
    if {$opt(xlMaxRows) == 100003} {
      outputMsg " "
      outputMsg "SFA variables" red
      catch {outputMsg " Drive [file nativename $drive]"}
      catch {outputMsg " Home  $myhome"}
      catch {outputMsg " Docs  $mydocs"}
      catch {outputMsg " Desk  $mydesk"}
      catch {outputMsg " Menu  $mymenu"}
      catch {outputMsg " Temp  $mytemp"}
      outputMsg " pf32  $pf32"
      if {$pf64 != ""} {outputMsg " pf64  $pf64"}
      outputMsg " S [winfo screenwidth  .]x[winfo screenheight  .], M [winfo reqwidth .]x[expr {int([winfo reqheight .]*1.05)}]"
      outputMsg " $winver"
      catch {outputMsg " scriptName [file nativename $scriptName]"}

      outputMsg "Registry values" red
      catch {outputMsg " Personal  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Personal}]"}
      catch {outputMsg " Desktop   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Desktop}]"}
      catch {outputMsg " Programs  [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Programs}]"}
      catch {outputMsg " AppData   [registry get {HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders} {Local AppData}]"}

      outputMsg "Environment variables" red
      foreach id [lsort [array names env]] {
        foreach id1 [list HOME USER APP ROSE] {if {[string first $id1 $id] == 0} {catch {outputMsg " $id  $env($id)"; break}}}
      }
    }
    .tnb select .tnb.status
  }
}
