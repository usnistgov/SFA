proc initData {} {

global entCategory entColorIndex badAttributes roseLogical
global aoEntTypes gpmiTypes spmiEntTypes dimSizeNames tolNames tzfNames dimModNames pmiModifiers pmiModifiersRP pmiUnicode
global spmiTypes recPracNames modelPictures schemaLinks modelURLs legendColor

initDataEntities

set roseLogical(0) "FALSE"
set roseLogical(1) "TRUE"
set roseLogical(2) "UNKNOWN"

# -----------------------------------------------------------------------------------------------------

set recPracNames(pmi242)   "Representation and Presentation of PMI (AP242)"
set recPracNames(pmi203)   "PMI Polyline Presentation (AP203/AP214)"
set recPracNames(valprop)  "Geometric and Assembly Validation Properties"
set recPracNames(tessgeom) "3D Tessellated Geometry"
set recPracNames(uda)      "User Defined Attributes"

set schemaLinks(AP203) "http://www.steptools.com/support/stdev_docs/express/ap203e2/html/index.html"
set schemaLinks(AP209) "http://www.steptools.com/support/stdev_docs/express/ap209/html/index.html"
set schemaLinks(AP210) "http://www.steptools.com/support/stdev_docs/express/ap210/html/index.html"
set schemaLinks(AP214) "http://www.steptools.com/support/stdev_docs/express/ap214/html/index.html"
set schemaLinks(AP238) "http://www.steptools.com/support/stdev_docs/express/ap238/html/index.html"
set schemaLinks(AP242) "http://www.steptools.com/support/stdev_docs/express/ap242/html/index.html"

# list of annotation occurrence entities, order is important
set aoEntTypes [list \
  tessellated_annotation_occurrence \
  annotation_fill_area_occurrence \
  annotation_curve_occurrence \
  annotation_occurrence \
  draughting_annotation_occurrence \
]

# list of semantic PMI entities, order is important, not including tolerances
set spmiEntTypes [list \
  datum_reference_element \
  datum_reference_compartment \
  datum_system \
  datum_reference \
  referenced_modified_datum \
  placed_datum_target_feature \
  datum_target \
  dimensional_characteristic_representation \
]

# -----------------------------------------------------------------------------------------------------
# dimensional_size names (Section 5.1.5, Table 4)

set dimSizeNames [list \
  "curve length" "diameter" "thickness" "spherical diameter" "radius" "spherical radius" \
  "toroidal minor diameter" "toroidal major diameter" "toroidal minor radius" "toroidal major radius" \
  "toroidal high major diameter" "toroidal high minor diameter" "toroidal high major radius" "toroidal high minor radius"]

# add controlled radius and square (not in Table 4)
#lappend dimSizeNames "controlled radius"
#lappend dimSizeNames "square"
                   
# dimension modifiers (Section 5.3, Table 8)
set dimModNames [list \
  "any cross section" "any part of the feature" "area diameter calculated size" "average rank order size" \
  "circumference diameter calculated size" "common tolerance" "continuous feature" "controlled radius" \
  "free state condition" "least squares association criteria" "local size defined by a sphere" \
  "maximum inscribed association criteria" "maximum rank order size" "median rank order size" \
  "mid range rank order size" "minimum circumscribed association criteria" "minimum rank order size" \
  "range rank order size" "specific fixed cross section" "square" "statistical" \
  "two point size" "volume diameter calculated size"]

# -----------------------------------------------------------------------------------------------------
# tolerance entity names (Section 6.8, Table 10)

set tolNames [list \
  angularity_tolerance circular_runout_tolerance coaxiality_tolerance concentricity_tolerance cylindricity_tolerance \
  flatness_tolerance line_profile_tolerance parallelism_tolerance perpendicularity_tolerance position_tolerance \
  roundness_tolerance straightness_tolerance surface_profile_tolerance symmetry_tolerance total_runout_tolerance]
                   
# tolerance zone form names (Section 6.9.2, Tables 11, 12)
set tzfNames [list \
  "cylindrical or circular" "spherical" "within a circle" "between two concentric circles" "between two equidistant curves" \
  "within a cylinder" "between two coaxial cylinders" "between two equidistant surfaces" "non uniform" "unknown"]

# -----------------------------------------------------------------------------------------------------
# *Graphical PMI* names (Section 8.2, Table 13)

set gpmiTypes [list \
  "angularity" "circular runout" "circularity" "coaxiality" "concentricity" "cylindricity" \
  "flatness" "parallelism" "perpendicularity" "position" "profile of line" "profile of surface" \
  "roundness" "straightness" "symmetry" "total runout" "general tolerance" "linear dimension" \
  "radial dimension" "diameter dimension" "angular dimension" "ordinate dimension" "curve dimension" \
  "general dimension" "datum" "datum target" "note" "label" "surface roughness" "weld symbol"]

# -----------------------------------------------------------------------------------------------------
# Semantic PMI types for coverage analysis, order is important

set spmiTypes $tolNames

foreach item [list \
  "composite tolerance (6.9.9)" "dimensional location (5.1.1)" "dimensional size (5.1.5)" "angular location (5.1.2)" "angular size (5.1.6)" \
  "plusminus - equal (5.2.3)" "plusminus - unequal (5.2.3)" "value range (5.2.4)" "diameter \u2205 (5.1.5)" \
  "radius R (5.1.5)" "spherical diameter S\u2205 (5.1.5)" "spherical radius SR (5.1.5)" "controlled radius CR (5.3)" "square \u25A1 (5.3)" \
  "basic dimension (5.3)" "reference dimension (5.3)" "type qualifier (5.2.2)" "tolerance class (5.2.5)" \
  "oriented dimensional location (5.1.3)" "derived shapes dimensional location (5.1.4)" "location with path (5.1.7)" "decimal places (5.4)" \
  "datum (6.5)" "multiple datum features (6.9.8)" "datum with axis system (6.9.7)" "datum with modifiers (6.9.7)" \
  "point datum target (6.6)" "circle datum target (6.6)" "rectangle datum target (6.6)" "line datum target (6.6)" \
  "area datum target (6.6)" "curve datum target type" "moveable datum target (6.6.3)" "placed datum target feature (6.6.2)" \
  "tolerance zone diameter (6.9.2)" "tolerance zone spherical diameter (6.9.2)" "affected plane tolerance zone (6.9.2.1)" \
  "non-uniform tolerance zone (6.9.2.3)" "tolerance with max value (6.9.5)" "unit-basis tolerance (6.9.6)"] {lappend spmiTypes $item}

  #"stacked tolerance" "feature count 'nX'" "curve length" "placed datum target (6.6)" "dimension modifier (5.3)" "thickness"

# -----------------------------------------------------------------------------------------------------
# pmiModifiers are the symbols associated with many strings such as dimModNames and others

set pmiModifiersArray(all_around,6.4.2)                     "\u232E"
set pmiModifiersArray(all_over,6.3)                         "ALL OVER"
set pmiModifiersArray(any_cross_section,5.3)                "ACS"
set pmiModifiersArray(any_longitudinal_section,6.9.7)       "ALS"
set pmiModifiersArray(any_part_of_the_feature,5.3)          "/Length"
set pmiModifiersArray(arc_length)                           "\u2322"
set pmiModifiersArray(area_diameter_calculated_size,5.3)    "(CA)"
set pmiModifiersArray(average_rank_order_size,5.3)          "(SA)"
set pmiModifiersArray(basic,6.9.7)                          "\[BASIC\]"
set pmiModifiersArray(between,6.4.3)                        "\u2194"
set pmiModifiersArray(circumference_diameter_calculated_size,5.3) "(CC)"
set pmiModifiersArray(common_zone,6.9.3)                    "CZ"
set pmiModifiersArray(conical_taper)                        "\u2332"
set pmiModifiersArray(contacting_feature,6.9.7)             "CF"
set pmiModifiersArray(continuous_feature,5.3)               "<CF>"
set pmiModifiersArray(controlled_radius,5.3)                "CR"
set pmiModifiersArray(counterbore)                          "\u2334"
set pmiModifiersArray(countersink)                          "\u2335"
set pmiModifiersArray(degree_of_freedom_constraint_u,6.9.7) "u"
set pmiModifiersArray(degree_of_freedom_constraint_v,6.9.7) "v"
set pmiModifiersArray(degree_of_freedom_constraint_w,6.9.7) "w"
set pmiModifiersArray(degree_of_freedom_constraint_x,6.9.7) "x"
set pmiModifiersArray(degree_of_freedom_constraint_y,6.9.7) "y"
set pmiModifiersArray(degree_of_freedom_constraint_z,6.9.7) "z"
set pmiModifiersArray(depth)                                "\u21A7"
set pmiModifiersArray(dimension_origin)                     "\u2331"
set pmiModifiersArray(distance_variable,6.9.7)              "DV"
set pmiModifiersArray(each_radial_element,6.9.3)            "ERE"
set pmiModifiersArray(free_state,5.3-6.9.3)                 "\u24BB"
set pmiModifiersArray(hole_thread)                          ""
set pmiModifiersArray(independency,5.2.1)                   "\u24BE"
set pmiModifiersArray(least_material_condition)             "\u24C1"
set pmiModifiersArray(least_material_requirement,6.9.3-6.9.7)   "\u24C1"
set pmiModifiersArray(least_square_association_criteria,5.3)  "(GG)"
set pmiModifiersArray(line)                                 "SL"
set pmiModifiersArray(line_element,6.9.3)                   "LE"
set pmiModifiersArray(local_size_defined_by_a_sphere,5.3)   "(LS)"
set pmiModifiersArray(major_diameter,6.9.3)                 "MD"
set pmiModifiersArray(maximum_inscribed_association_criteria,5.3) "(GX)"
set pmiModifiersArray(maximum_material_condition)           "\u24C2"
set pmiModifiersArray(maximum_material_requirement,6.9.3-6.9.7) "\u24C2"
set pmiModifiersArray(maximum_rank_order_size,5.3)          "(SX)"
set pmiModifiersArray(median_rank_order_size,5.3)           "(SM)"
set pmiModifiersArray(mid_range_rank_order_size,5.3)        "(SD)"
set pmiModifiersArray(minimum_inscribed_association_criteria,5.3) "(GN)"
set pmiModifiersArray(minimum_rank_order_size,5.3)          "(SN)"
set pmiModifiersArray(minor_diameter,6.9.3)                 "LD"
set pmiModifiersArray(not_convex,6.9.3)                     "NC"
set pmiModifiersArray(orientation,6.9.7)                    "\u003E\u003C"
set pmiModifiersArray(pitch_diameter,6.9.7)                 "PD"
set pmiModifiersArray(plane,6.9.7)                          "PL"
set pmiModifiersArray(point,6.9.7)                          "PT"
set pmiModifiersArray(projected,6.9.2.2)                    "\u24C5"
set pmiModifiersArray(range_rank_order_size,5.3)            "(SR)"
set pmiModifiersArray(reciprocity_requirement,6.9.3)        "\u24C7"
set pmiModifiersArray(regardless_of_feature_size)           "\u24C8"
set pmiModifiersArray(separate_requirement,6.9.3)           "SEP REQT"
set pmiModifiersArray(simultaneous_requirement)             "SIM REQT"
set pmiModifiersArray(slope)                                "\u2333"
set pmiModifiersArray(specific_fixed_cross_section,5.3)     "SCS"
set pmiModifiersArray(spotface)                             "SF"
set pmiModifiersArray(square,5.3)                           "\u25A1"
set pmiModifiersArray(statistical,5.3)                      "<ST>"
set pmiModifiersArray(statistical_tolerance,6.9.3)          "<ST>"
set pmiModifiersArray(tangent_plane,6.9.3)                  "\u24C9"
set pmiModifiersArray(translation,6.9.7)                    "\u25B7"
set pmiModifiersArray(two_point_size,5.3)                   "(LP)"
set pmiModifiersArray(unequally_disposed,6.9.4)             "\u24CA"
set pmiModifiersArray(volume_diameter_calculated_size,5.3)  "(CV)"

foreach item [array names pmiModifiersArray] {
  set ids [split $item ","]
  set pmiModifiers([lindex $ids 0]) $pmiModifiersArray($item)
  if {[llength $ids] > 1} {set pmiModifiersRP([lindex $ids 0]) [lindex $ids 1]}
}

# pmfirst are things in the NIST models
set pmfirst [list maximum_material_requirement least_material_requirement unequally_disposed projected free_state tangent_plane \
             statistical statistical_tolerance \
             all_around separate_requirement simultaneous_requirement dimension_origin between \
             counterbore depth hole_thread countersink \
             slope conical_taper arc_length]

foreach pmf $pmfirst {             
  foreach item [lsort [array names pmiModifiers]] {
    set idx [lindex [split $item ","] 0]
    if {$pmf == $idx} {lappend spmiTypes $item}
  }
}
foreach item [lsort [array names pmiModifiers]] {
  set idx [lindex [split $item ","] 0]
  if {[lsearch $pmfirst $idx] == -1} {lappend spmiTypes $item}
}
#foreach item [array names pmiModifiers] {lappend spmiTypes $item}

# -----------------------------------------------------------------------------------------------------
# pmiUnicode are the symbols associated with tolerances and a few others

set idx "cylindrical or circular"
set pmiUnicode($idx)             "\u2205"
set pmiUnicode(angularity)       "\u2220"
set pmiUnicode(circular)         "\u2205"
set pmiUnicode(circular_runout)  "\u2197"
set pmiUnicode(coaxiality)       "\u25CE"
set pmiUnicode(concentricity)    "\u25CE"
set pmiUnicode(cylindrical)      "\u2205"
set pmiUnicode(cylindricity)     "\u232D"
set pmiUnicode(degree)           "\u00B0"
set pmiUnicode(diameter)         "\u2205"
set pmiUnicode(flatness)         "\u25B1"
set pmiUnicode(line_profile)     "\u2312"
set pmiUnicode(parallelism)      "\u2215\u2215"
set pmiUnicode(perpendicularity) "\u23CA"
set pmiUnicode(plusminus)        "\u00B1"
set pmiUnicode(position)         "\u2295"
set pmiUnicode(radius)           "R"
set pmiUnicode(roundness)        "\u25EF"
set idx "spherical diameter"
set pmiUnicode($idx)             "S\u2205"
set idx "spherical radius"
set pmiUnicode($idx)             "SR"
set pmiUnicode(straightness)     "-"
set pmiUnicode(square)           "\u25A1"
set pmiUnicode(surface_profile)  "\u2313"
set pmiUnicode(symmetry)         "\u232F"
set pmiUnicode(thickness)        "\u2346\u2345"
set pmiUnicode(total_runout)     "\u2330"

# -----------------------------------------------------------------------------------------------------
# colors, the number determines the order that the group of entities is processed
# do not use numbers less than 10  (dmcritchie.mvps.org/excel/colors.htm)

set entColorIndex(PR_STEP_AP203) -1			
set entColorIndex(PR_STEP_AP214) 15			

set entColorIndex(PR_STEP_AP209) 19			
set entColorIndex(PR_STEP_AP210) 19			
set entColorIndex(PR_STEP_AP238) 19

set entColorIndex(PR_STEP_AP242) 20	
set entColorIndex(PR_STEP_AP242_QUAL) 24	
set entColorIndex(PR_STEP_AP242_CONS) 33	
set entColorIndex(PR_STEP_AP242_MATH) 34	
set entColorIndex(PR_STEP_AP242_KINE) 35	
set entColorIndex(PR_STEP_AP242_GEOM) 36	

set entColorIndex(PR_STEP_TOLR) 37			
set entColorIndex(PR_STEP_PRES) 38			
set entColorIndex(PR_STEP_REP) 39			
set entColorIndex(PR_STEP_ASPECT) 40			
set entColorIndex(PR_STEP_OTHER) 42			
set entColorIndex(PR_STEP_GEO) 43			
set entColorIndex(PR_STEP_CPNT) 43			
set entColorIndex(PR_STEP_QUAN) 44

# PMI coverage colors
# [expr {int ($b) << 16 | int ($g) << 8 | int($r)}]

set legendColor(green)   [expr {int (128) << 16 | int (255) << 8 | int(128)}]
set legendColor(yellow)  [expr {int (128) << 16 | int (255) << 8 | int(255)}]
set legendColor(red)     [expr {int (128) << 16 | int (128) << 8 | int(255)}]
set legendColor(cyan)    [expr {int (255) << 16 | int (255) << 8 | int(128)}]
set legendColor(magenta) [expr {int (255) << 16 | int (128) << 8 | int(255)}]
set legendColor(gray)    [expr {int (208) << 16 | int (208) << 8 | int(208)}]
#set legendColor(green) [expr 8454016]
#set legendColor(yellow) [expr 8454143]
#set legendColor(red) [expr 8421631]
#set legendColor(gray) [expr 12632256]

# -----------------------------------------------------------------------------------------------------
# entity attributes that cause a crash, LIST of LIST

set badAttributes(b_spline_surface_with_knots) {control_points_list}
set badAttributes(b_spline_surface_with_knots_and_rational_b_spline_surface) {control_points_list weights_data}
set badAttributes(bezier_surface) {control_points_list}
set badAttributes(bezier_surface_and_rational_b_spline_surface) {control_points_list weights_data}
set badAttributes(cc_design_approval) {items}
set badAttributes(complex_triangulated_face) {normals triangle_fans triangle_strips}
set badAttributes(complex_triangulated_surface_set) {normals triangle_fans triangle_strips}
set badAttributes(coordinates_list) {position_coords}
set badAttributes(curve_3d_element_descriptor) {purpose}
set badAttributes(finite_function) {pairs}
set badAttributes(quasi_uniform_surface) {control_points_list}
set badAttributes(quasi_uniform_surface_and_rational_b_spline_surface) {control_points_list weights_data}
set badAttributes(rational_b_spline_surface) {weights_data}
set badAttributes(rational_b_spline_surface_and_uniform_surface) {control_points_list weights_data}
set badAttributes(rectangular_composite_surface) {segments}
set badAttributes(solid_with_incomplete_rectangular_pattern) {omitted_instances}
set badAttributes(surface_3d_element_descriptor) {purpose}
set badAttributes(tessellated_curve_set) {line_strips}
set badAttributes(tessellated_face) {normals}
set badAttributes(tessellated_surface_set) {normals}
set badAttributes(triangulated_face) {normals triangles}
set badAttributes(triangulated_surface_set) {triangles}

# -----------------------------------------------------------------------------------------------------
# pictures that are embedded in a spreadsheet based on STEP file name
set modelPictures  {{sp3-1101  caxif-1101.jpg  E3 0} \
                    {sp3-16792 caxif-16792.jpg E3 0} \
                    {sp3-box   caxif-boxy1.jpg E3 0} \
                    {sp3-box   caxif-boxy2.jpg Q3 16} \
                    {sp3-box   caxif-boxy3.jpg AC3 28} \
                    {STEP-File-Analyzer nist_ctc_01.jpg E3 0} \
                    {nist_ctc_01 nist_ctc_01.jpg E4 0} \
                    {nist_ctc_02 nist_ctc_02a.jpg E4 0} \
                    {nist_ctc_02 nist_ctc_02b.jpg U4 20} \
                    {nist_ctc_02 nist_ctc_02c.jpg AK4 36} \
                    {nist_ctc_03 nist_ctc_03.jpg E4 0} \
                    {nist_ctc_04 nist_ctc_04.jpg E4 0} \
                    {nist_ctc_05 nist_ctc_05a.jpg E4 0} \
                    {nist_ctc_05 nist_ctc_05b.jpg U4 20} \
                    {nist_ftc_06 nist_ftc_06a.jpg E4 0} \
                    {nist_ftc_06 nist_ftc_06b.jpg U4 20} \
                    {nist_ftc_06 nist_ftc_06c.jpg AK4 36} \
                    {nist_ftc_07 nist_ftc_07a.jpg E4 0} \
                    {nist_ftc_07 nist_ftc_07b.jpg U4 20} \
                    {nist_ftc_07 nist_ftc_07c.jpg AK4 36} \
                    {nist_ftc_07 nist_ftc_07d.jpg BA4 52} \
                    {nist_ftc_08 nist_ftc_08a.jpg E4 0} \
                    {nist_ftc_08 nist_ftc_08b.jpg U4 20} \
                    {nist_ftc_08 nist_ftc_08c.jpg AK4 36} \
                    {nist_ftc_08 nist_ftc_08d.jpg BA4 52} \
                    {nist_ftc_09 nist_ftc_09a.jpg E4 0} \
                    {nist_ftc_09 nist_ftc_09b.jpg U4 20} \
                    {nist_ftc_09 nist_ftc_09c.jpg AK4 36} \
                    {nist_ftc_09 nist_ftc_09d.jpg BA4 52} \
                    {nist_ftc_10 nist_ftc_10a.jpg E4 0} \
                    {nist_ftc_10 nist_ftc_10b.jpg U4 20} \
                    {nist_ftc_10 nist_ftc_10c.jpg AK4 36} \
                    {nist_ftc_10 nist_ftc_10d.jpg BA4 52} \
                    {nist_ftc_10 nist_ftc_10e.jpg BQ4 68} \
                    {nist_ftc_11 nist_ftc_11a.jpg E4 0} \
                    {nist_ftc_11 nist_ftc_11b.jpg U4 20}}

set modelURLs [list nist_ctc_01_asme1_rd.pdf \
                    nist_ctc_02_asme1_rc.pdf \
                    nist_ctc_03_asme1_rc.pdf \
                    nist_ctc_04_asme1_rd.pdf \
                    nist_ctc_05_asme1_rd.pdf \
                    nist_ftc_06_asme1_rd.pdf \
                    nist_ftc_07_asme1_rc.pdf \
                    nist_ftc_08_asme1_rc.pdf \
                    nist_ftc_09_asme1_rd.pdf \
                    nist_ftc_10_asme1_rb.pdf \
                    nist_ftc_11_asme1_rb.pdf]

# -----------------------------------------------------------------------------------------------------
# STEP AP203

set entCategory(PR_STEP_AP203) [lsort [list \
abstract_variable action_method_assignment action_method_role angle_direction_reference applied_action_method_assignment applied_attribute_classification_assignment applied_usage_right assigned_requirement atomic_formula attribute_assertion \
back_chaining_rule back_chaining_rule_body beveled_sheet_representation breakdown_context breakdown_element_group_assignment breakdown_element_realization breakdown_element_usage breakdown_of \
cc_design_approval cc_design_certification cc_design_contract cc_design_date_and_time_assignment cc_design_person_and_organization_assignment cc_design_security_classification cc_design_specification_reference change change_request character_glyph_font_usage character_glyph_style_outline character_glyph_style_stroke character_glyph_symbol_outline character_glyph_symbol_stroke characteristic_data_column_header characteristic_data_column_header_link characteristic_data_table_header characteristic_data_table_header_decomposition characteristic_type class_by_extension class_by_intension complex_clause complex_conjunctive_clause complex_disjunctive_clause complex_shelled_solid composite_assembly_definition composite_assembly_sequence_definition composite_assembly_table composite_material_designation composite_sheet_representation composite_text_with_delineation configuration_item_hierarchical_relationship configuration_item_relationship configuration_item_revision_sequence conical_stepped_hole_transition contract_relationship currency curve_style_font_and_scaling \
definitional_representation_relationship definitional_representation_relationship_with_same_context design_make_from_relationship dimension_curve_terminator_to_projection_curve_associativity document_identifier document_identifier_assignment double_offset_shelled_solid draped_defined_transformation drawing_sheet_revision_sequence \
edge_blended_solid elementary_brep_shape_representation entity_assertion enum_reference_prefix evaluated_characteristic evaluation_product_definition event_occurrence_relationship expanded_uncertainty explicit_procedural_geometric_representation_item_relationship explicit_procedural_representation_item_relationship explicit_procedural_representation_relationship explicit_procedural_shape_representation_relationship extent external_class_library external_source_relationship externally_defined_colour externally_defined_context_dependent_unit externally_defined_conversion_based_unit externally_defined_currency externally_defined_marker externally_defined_picture_representation_item externally_defined_representation_item externally_defined_string externally_defined_terminator_symbol externally_defined_tile extruded_face_solid_with_draft_angle extruded_face_solid_with_multiple_draft_angles extruded_face_solid_with_trim_conditions \
fact_type fill_area_style_tile_coloured_region fill_area_style_tile_curve_with_style flat_pattern_ply_representation_relationship forward_chaining_rule forward_chaining_rule_premise func functional_breakdown_context functional_element_usage \
geometric_model_element_relationship global_assignment ground_fact \
included_text_block indirectly_selected_elements indirectly_selected_shape_elements information_right information_usage_right instance_usage_context_assignment iso4217_currency \
laid_defined_transformation laminate_table literal_conjunction literal_disjunction logical_literal \
mechanical_design_presentation_representation_with_draughting mechanical_design_shaded_presentation_area mechanical_design_shaded_presentation_representation min_and_major_ply_orientation_basis modified_solid modified_solid_with_placed_configuration \
ordinal_date \
part_laminate_table partial_document_with_structured_text_representation_assignment percentage_laminate_definition percentage_laminate_table percentage_ply_definition physical_breakdown_context physical_element_usage picture_representation ply_laminate_definition ply_laminate_sequence_definition ply_laminate_table point_and_vector point_path polar_complex_number_literal positioned_sketch pre_defined_surface_side_style pre_defined_tile procedural_representation procedural_representation_sequence procedural_shape_representation procedural_shape_representation_sequence product_definition_element_relationship product_definition_group_assignment product_material_composition_relationship \
range_characteristic representation_item_relationship requirement_assigned_object requirement_assignment requirement_source requirement_view_definition_relationship revolved_face_solid_with_trim_conditions right_to_usage_association row_value row_variable rule_action rule_condition rule_definition rule_set rule_set_group rule_software_definition rule_superseded_assignment rule_supersedence \
satisfied_requirement satisfies_requirement satisfying_item scalar_variable scattering_parameter sculptured_solid shape_feature_definition shell_based_wireframe_model shell_based_wireframe_shape_representation shelled_solid simple_clause smeared_material_definition solid_curve_font solid_with_angle_based_chamfer solid_with_chamfered_edges solid_with_circular_pattern solid_with_circular_pocket solid_with_circular_protrusion solid_with_conical_bottom_round_hole solid_with_constant_radius_edge_blend solid_with_curved_slot solid_with_depression solid_with_double_offset_chamfer solid_with_flat_bottom_round_hole solid_with_general_pocket solid_with_general_protrusion solid_with_groove solid_with_hole solid_with_incomplete_circular_pattern solid_with_incomplete_rectangular_pattern solid_with_pocket solid_with_protrusion solid_with_rectangular_pattern solid_with_rectangular_pocket solid_with_rectangular_protrusion solid_with_shape_element_pattern solid_with_single_offset_chamfer solid_with_slot solid_with_spherical_bottom_round_hole solid_with_stepped_round_hole solid_with_stepped_round_hole_and_conical_transitions solid_with_straight_slot solid_with_tee_section_slot solid_with_through_depression solid_with_trapezoidal_section_slot solid_with_variable_radius_edge_blend source_for_requirement sourced_requirement specification_definition start_request start_work structured_text_composition structured_text_representation supplied_part_relationship surfaced_open_shell symbol \
tagged_text_format tagged_text_item text_font text_font_family text_font_in_family thickened_face_solid thickness_laminate_definition thickness_laminate_table time_interval_relationship track_blended_solid track_blended_solid_with_end_conditions transformation_with_derived_angle \
uniform_resource_identifier usage_association user_defined_curve_font user_defined_marker user_defined_terminator_symbol user_selected_elements user_selected_shape_elements \
vertex_shell \
week_of_year_and_day_date wire_shell \
year_month \
zone_structural_makeup]]

# -----------------------------------------------------------------------------------------------------
# STEP AP214

set entCategory(PR_STEP_AP214) [lsort [list \
abs_function acos_function action action_assignment action_directive action_method action_method_relationship action_property action_property_representation action_relationship action_request_assignment action_request_solution action_request_status action_resource action_resource_requirement action_resource_type action_status and_expression application_context_relationship applied_area applied_ineffectivity_assignment approximation_tolerance approximation_tolerance_deviation approximation_tolerance_parameter asin_function atan_function \
barring_hole bead bead_end binary_boolean_expression binary_function_call boolean_defined_function boolean_variable boss boss_top \
camera_image_2d_with_scale camera_model_d2 chamfer chamfer_offset circular_closed_profile circular_pattern closed_path_profile comparison_equal comparison_greater comparison_greater_equal comparison_less comparison_less_equal comparison_not_equal composite_hole compound_feature concat_expression configuration_definition configuration_interpolation cos_function cylindrical_pair cylindrical_pair_range cylindrical_pair_value \
defined_character_glyph defined_function directed_angle direction_shape_representation div_expression draughting_specification_reference drawing_sheet_layout \
edge_round element_delivery equals_expression event_occurrence_context_assignment event_occurrence_context_role exp_function externally_defined_character_glyph externally_defined_feature_definition externally_defined_style \
face_shape_representation feature_component_definition feature_component_relationship feature_definition feature_in_panel feature_pattern featured_shape fillet format_function founded_kinematic_path fully_constrained_pair \
gear_pair gear_pair_range gear_pair_value general_feature \
hole_bottom hole_in_panel homokinetic_pair \
index_expression initial_state int_numeric_variable int_value_function integer_defined_function interpolated_configuration_sequence \
joggle joggle_termination \
kinematic_analysis_consistency kinematic_analysis_result kinematic_control kinematic_frame_background_representation kinematic_frame_background_representation_association kinematic_frame_based_transformation kinematic_ground_representation kinematic_joint kinematic_link kinematic_link_representation kinematic_link_representation_association kinematic_link_representation_relation kinematic_pair kinematic_path kinematic_property_definition kinematic_property_representation_relation kinematic_structure \
language_assignment length_function like_expression location_shape_representation locator log10_function log2_function log_function \
maximum_function mechanism mechanism_base_placement minimum_function minus_expression minus_function mod_expression modified_pattern motion_link_relationship mult_expression multiple_arity_function_call \
ngon_closed_profile not_expression numeric_defined_function numeric_variable \
odd_function open_path_profile or_expression \
pair_actuator pair_value partial_circular_profile path_feature_component path_shape_representation pattern_offset_membership pattern_omit_membership physically_modelled_product_definition planar_curve_pair planar_curve_pair_range planar_pair planar_pair_range planar_pair_value planar_shape_representation plus_expression pocket pocket_bottom point_on_planar_curve_pair point_on_planar_curve_pair_range point_on_planar_curve_pair_value point_on_surface_pair point_on_surface_pair_range point_on_surface_pair_value power_expression pre_defined_presentation_style prismatic_pair prismatic_pair_range prismatic_pair_value process_operation process_plan process_product_association process_property_association product_definition_process product_definition_resource product_process_plan property_process \
rack_and_pinion_pair rack_and_pinion_pair_range rack_and_pinion_pair_value real_defined_function real_numeric_variable rectangular_closed_profile rectangular_pattern replicate_feature requirement_for_action_resource resource_property resource_property_representation resource_requirement_type resulting_path retention revolute_pair revolute_pair_range revolute_pair_value rib rolling_curve_pair rolling_curve_pair_value rolling_surface_pair rolling_surface_pair_value rotation_about_direction round_hole rounded_u_profile \
screw_pair screw_pair_range screw_pair_value seam_edge shape_defining_relationship simple_pair_range simple_string_expression sin_function sliding_curve_pair sliding_curve_pair_value sliding_surface_pair sliding_surface_pair_value slot slot_end spherical_pair spherical_pair_range spherical_pair_value sql_mappable_defined_function square_root_function square_u_profile string_defined_function string_expression string_literal string_variable substring_expression surface_pair surface_pair_range \
tan_function taper tee_profile thread transition_feature \
unary_boolean_expression unary_function_call unconstrained_pair unconstrained_pair_value universal_pair universal_pair_range universal_pair_value \
value_function variable vee_profile versioned_action_request_relationship \
xor_expression]]

# -----------------------------------------------------------------------------------------------------
# STEP AP242 geometry

set entCategory(PR_STEP_AP242_GEOM) [lsort [list \
area_with_outer_boundary \
b_spline_curve_knot_locator b_spline_curve_segment b_spline_surface_knot_locator b_spline_surface_patch b_spline_surface_strip boolean_result_2d boundary_curve_of_b_spline_or_rectangular_composite_surface \
circular_area complex_triangulated_face complex_triangulated_surface_set coordinates_list csg_2d_shape_representation csg_primitive_solid_2d csg_solid_2d curve_segment_set \
elliptic_area \
flat_face \
half_space_2d \
implicit_intersection_curve implicit_model_intersection_curve implicit_planar_curve implicit_planar_intersection_point implicit_planar_projection_point implicit_point_on_plane implicit_projected_curve implicit_silhouette_curve \
plane_angle_and_length_pair plane_angle_and_ratio_pair point_on_edge_curve point_on_face_surface polygonal_area primitive_2d primitive_2d_with_inner_boundary \
rectangular_area repositioned_neutral_sketch repositioned_tessellated_item \
single_area_csg_2d_shape_representation single_boundary_csg_2d_shape_representation surface_patch_set \
tessellated_connecting_edge tessellated_curve_set tessellated_edge tessellated_face tessellated_geometric_set tessellated_item tessellated_point_set tessellated_shell tessellated_solid tessellated_structured_item tessellated_surface_set tessellated_vertex tessellated_wire triangulated_face triangulated_surface_set \
volume \
cyclide_segment_solid eccentric_cone ellipsoid faceted_primitive tetrahedron convex_hexahedron rectangular_pyramid \
]]

# -----------------------------------------------------------------------------------------------------
# STEP AP242

set entCategory(PR_STEP_AP242) [lsort [list \
SQL_mappable_defined_function \
abstracted_expression_function \
add_element \
array_placement_group \
assembly_bond_definition \
assembly_component \
assembly_group_component \
assembly_group_component_definition_placement_link \
assembly_joint \
atom_based_literal \
basic_sparse_matrix \
binary_literal \
bound_parameter_environment \
bound_variational_parameter \
chain_based_geometric_item_specific_usage \
chain_based_item_identified_representation_usage \
change_composition_relationship \
change_element \
change_element_sequence \
change_group \
change_group_assignment \
characterized_chain_based_item_within_representation \
closed_curve_style_parameters \
complex_area \
complex_number_literal \
complex_number_literal_polar \
component_definition \
component_feature \
component_feature_joint \
component_feature_relationship \
component_terminal \
composite_curve_transition_locator \
connection_zone_based_assembly_joint \
connection_zone_interface_plane_relationship \
contacting_feature \
current_change_element_assignment \
curve_style_parameters_representation \
curve_style_parameters_with_ends \
definite_integral_expression \
delete_element \
detailed_report_request \
detailed_report_request_with_number_of_data \
evaluated_characteristic_of_product_as_individual_test_result \
expression_extension_numeric \
expression_extension_string \
expression_extension_to_select \
externally_defined_item_with_multiple_references \
externally_defined_representation \
face_shape_representation_relationship \
feature_definition_with_connection_area \
fixed_instance_attribute_set \
free_form_assignment \
free_form_relation \
frozen_assignment \
function_application \
gear \
general_datum_reference \
generated_finite_numeric_space \
generic_product_definition_reference \
geometric_representation_context_with_parameter \
implicit_explicit_positioned_sketch_relationship \
instance_attribute_reference \
instance_report_item_with_extreme_instances \
integer_tuple_literal \
interfaced_group_component \
linear_array_component_definition_link \
linear_array_placement_group_component \
linear_profile \
listed_data \
location_in_aggregate_representation_item \
make_from_feature_relationship \
marking \
mated_part_relationship \
modify_element \
multi_level_reference_designator \
near_point_relationship \
neutral_sketch_representation \
oriented_joint \
outer_round \
outside_profile \
path_area_with_parameters \
path_parameter_representation \
path_parameter_representation_context \
physical_component \
physical_component_feature \
physical_component_terminal \
point_placement_shape_representation \
pre_defined_character_glyph \
previous_change_element_assignment \
product_as_planned \
product_definition_reference \
product_definition_reference_with_local_representation \
product_design_to_individual \
product_design_version_to_individual \
product_planned_to_realized \
product_relationship \
profile_floor \
protrusion \
quantifier_expression \
real_tuple_literal \
rectangular_array_placement_group_component \
rectangular_composite_surface_transition_locator \
removal_volume \
representation_proxy_item \
representative_shape_representation \
requirement_view_definition_relationship \
revolved_profile \
rib_top \
rib_top_floor \
rigid_subsketch \
rounded_end \
shape_criteria_representation_with_accuracy \
shape_inspection_result_accuracy_association \
shape_inspection_result_representation_with_accuracy \
shape_measurement_accuracy \
shape_summary_request_with_representative_value \
single_property_is_definition \
spherical_cap \
step \
su_parameters \
subsketch \
summary_report_request \
thermal_component \
thread_runout \
turned_knurl \
unbound_parameter_environment \
unbound_variational_parameter \
unbound_variational_parameter_semantics \
variable_expression \
variational_current_representation_relationship \
variational_parameter \
variational_representation \
]]

# -----------------------------------------------------------------------------------------------------
# STEP geometry

set entCategory(PR_STEP_GEO) [lsort [list \
advanced_face axis1_placement axis2_placement_2d axis2_placement_3d \
b_spline_curve b_spline_curve_with_knots b_spline_surface b_spline_surface_with_knots bezier_curve bezier_surface block boolean_result boundary_curve bounded_curve bounded_pcurve bounded_surface bounded_surface_curve box_domain boxed_half_space brep_with_voids \
cartesian_transformation_operator cartesian_transformation_operator_2d cartesian_transformation_operator_3d circle closed_shell composite_curve composite_curve_on_surface composite_curve_segment conic conical_surface connected_edge_set connected_face_set connected_face_sub_set csg_solid curve curve_bounded_surface curve_replica cylindrical_surface \
degenerate_pcurve degenerate_toroidal_surface direction \
edge edge_based_wireframe_model edge_curve edge_loop elementary_surface ellipse evaluated_degenerate_pcurve extruded_area_solid extruded_face_solid \
face face_based_surface_model face_bound face_outer_bound face_surface faceted_brep \
geometric_curve_set geometric_set \
half_space_solid hyperbola \
intersection_curve \
line loop \
manifold_solid_brep \
offset_curve_2d offset_curve_3d offset_surface open_shell oriented_closed_shell oriented_edge oriented_face oriented_open_shell oriented_path oriented_surface outer_boundary_curve \
parabola path pcurve placement planar_box planar_extent plane point point_on_curve point_on_surface point_replica poly_loop polyline \
quasi_uniform_curve quasi_uniform_surface \
rational_b_spline_curve rational_b_spline_surface rectangular_composite_surface rectangular_trimmed_surface reparametrised_composite_curve_segment revolved_area_solid revolved_face_solid right_angular_wedge right_circular_cone right_circular_cylinder ruled_surface_swept_area_solid \
seam_curve shell_based_surface_model solid_model solid_replica sphere spherical_surface subedge subface surface surface_curve surface_curve_swept_area_solid surface_of_linear_extrusion surface_of_revolution surface_patch surface_replica swept_area_solid swept_disk_solid swept_face_solid swept_surface \
toroidal_surface torus trimmed_curve \
uniform_curve uniform_surface \
vector vertex vertex_loop vertex_point \
]]

# STEP cartesian point

set entCategory(PR_STEP_CPNT) [list cartesian_point ]

# -----------------------------------------------------------------------------------------------------
# STEP other

set entCategory(PR_STEP_OTHER) [lsort [list \
address alternate_product_relationship \
application_context application_context_element application_protocol_definition \
applied_action_assignment applied_action_request_assignment applied_approval_assignment applied_certification_assignment applied_classification_assignment applied_contract_assignment applied_date_and_time_assignment applied_date_assignment applied_document_reference applied_document_usage_constraint_assignment applied_effectivity_assignment applied_event_occurrence_assignment applied_external_identification_assignment applied_group_assignment applied_identification_assignment applied_name_assignment applied_organization_assignment applied_organizational_project_assignment applied_person_and_organization_assignment applied_presented_item applied_security_classification_assignment applied_time_interval_assignment \
approval approval_assignment approval_date_time approval_person_organization approval_relationship approval_role approval_status \
assembly_component_usage assembly_component_usage_substitute \
attribute_classification_assignment attribute_language_assignment attribute_value_assignment attribute_value_role \
binary_generic_expression binary_numeric_expression \
boolean_expression boolean_literal \
calendar_date \
certification certification_assignment certification_type \
characterized_class characterized_object characterized_representation \
class class_system class_usage_effectivity_context_assignment \
classification_assignment classification_role \
comparison_expression \
concept_feature_operator concept_feature_relationship concept_feature_relationship_with_condition \
conditional_concept_feature \
configurable_item configuration_design configuration_effectivity configuration_item configured_effectivity_assignment configured_effectivity_context_assignment \
contact_ratio_representation \
contract contract_assignment contract_type \
coordinated_universal_time_offset \
data_environment \
date date_and_time date_and_time_assignment date_assignment date_role date_time_role dated_effectivity \
description_attribute design_context directed_action \
document document_file document_product_association document_product_equivalence document_reference document_relationship document_type document_usage_constraint document_usage_constraint_assignment document_usage_role \
effectivity effectivity_assignment effectivity_context_assignment effectivity_context_role effectivity_relationship \
environment \
event_occurrence event_occurrence_assignment event_occurrence_role \
exclusive_product_concept_feature_category \
executed_action expression expression_conversion_based_unit extension external_identification_assignment external_source \
founded_item \
functionally_defined_transformation \
general_material_property general_property general_property_association general_property_relationship \
generic_expression generic_variable \
group group_assignment group_relationship \
id_attribute \
identification_assignment identification_role \
inclusion_product_concept_feature instanced_feature int_literal interval_expression \
known_source \
language \
light_source light_source_ambient light_source_directional light_source_positional light_source_spot \
literal_number local_time lot_effectivity \
make_from_usage_option \
mapped_item \
material_designation material_designation_characterization material_property \
mechanical_context mechanical_design_geometric_presentation_area \
multi_language_attribute_assignment multiple_arity_boolean_expression multiple_arity_generic_expression multiple_arity_numeric_expression \
name_assignment name_attribute next_assembly_usage_occurrence numeric_expression \
object_role \
one_direction_repeat_factor \
organization organization_assignment organization_relationship organization_role organizational_address organizational_project organizational_project_assignment organizational_project_relationship organizational_project_role \
package_product_concept_feature \
person person_and_organization person_and_organization_address person_and_organization_assignment person_and_organization_role personal_address \
presented_item \
product product_category product_category_relationship product_class product_concept product_concept_context product_concept_feature product_concept_feature_association product_concept_feature_category product_concept_feature_category_usage product_concept_relationship product_context product_definition product_definition_context product_definition_context_association product_definition_context_role product_definition_effectivity product_definition_formation product_definition_formation_relationship product_definition_formation_with_specified_source product_definition_occurrence_relationship product_definition_relationship product_definition_shape product_definition_substitute product_definition_usage product_definition_with_associated_documents product_identification product_related_product_category product_specification \
promissory_usage_occurrence \
qualitative_uncertainty \
quantified_assembly_component_usage \
real_literal \
relative_event_occurrence \
rep_item_group \
role_association \
security_classification security_classification_assignment security_classification_level \
serial_numbered_effectivity \
simple_boolean_expression simple_generic_expression simple_numeric_expression \
slash_expression \
specified_higher_usage_occurrence \
standard_uncertainty \
time_interval time_interval_assignment time_interval_based_effectivity time_interval_role time_interval_with_bounds \
two_direction_repeat_factor \
unary_generic_expression \
unary_numeric_expression \
uncertainty_qualifier \
variable_semantics \
versioned_action_request \
]]

# -----------------------------------------------------------------------------------------------------
# STEP shape aspect

set entCategory(PR_STEP_ASPECT) [lsort [list \
all_around_shape_aspect \
apex \
between_shape_aspect \
centre_of_symmetry \
component_path_shape_aspect \
composite_group_shape_aspect \
composite_shape_aspect \
composite_unit_shape_aspect \
continuous_shape_aspect \
derived_shape_aspect \
geometric_alignment \
geometric_contact \
geometric_intersection \
geometric_item_specific_usage \
parallel_offset \
perpendicular_to \
shape_aspect \
shape_aspect_associativity \
shape_aspect_deriving_relationship \
shape_aspect_relationship \
shape_aspect_relationship_representation_association \
shape_aspect_transition \
symmetric_shape_aspect \
tangent \
]]

# -----------------------------------------------------------------------------------------------------
# STEP presentation, annotation

set entCategory(PR_STEP_PRES) [lsort [list \
angular_dimension \
annotation_curve_occurrence \
annotation_fill_area \
annotation_fill_area_occurrence \
annotation_occurrence \
annotation_occurrence_associativity \
annotation_occurrence_relationship \
annotation_plane \
annotation_subfigure_occurrence \
annotation_symbol \
annotation_symbol_occurrence \
annotation_text \
annotation_text_character \
annotation_text_occurrence \
area_in_set \
background_colour \
camera_image \
camera_image_3d_with_scale \
camera_model \
camera_model_d3 \
camera_model_d3_multi_clipping \
camera_model_d3_multi_clipping_intersection \
camera_model_d3_multi_clipping_union \
camera_model_d3_with_hlhsr \
camera_model_with_light_sources \
camera_usage \
character_glyph_symbol \
colour \
colour_rgb \
colour_specification \
composite_text \
composite_text_with_associated_curves \
composite_text_with_blanking_box \
composite_text_with_extent \
context_dependent_invisibility \
context_dependent_over_riding_styled_item \
curve_dimension \
curve_style \
curve_style_font \
curve_style_font_pattern \
curve_style_rendering \
datum_feature_callout \
datum_target_callout \
defined_symbol \
diameter_dimension \
dimension_callout \
dimension_callout_component_relationship \
dimension_callout_relationship \
dimension_curve \
dimension_curve_directed_callout \
dimension_curve_terminator \
dimension_pair \
dimension_related_tolerance_zone_element \
dimension_text_associativity \
draughting_annotation_occurrence \
draughting_callout \
draughting_callout_relationship \
draughting_elements \
draughting_model \
draughting_model_item_association \
draughting_pre_defined_colour \
draughting_pre_defined_curve_font \
draughting_pre_defined_text_font \
draughting_subfigure_representation \
draughting_symbol_representation \
draughting_text_literal_with_delineation \
draughting_title \
drawing_definition \
drawing_revision \
drawing_revision_sequence \
drawing_sheet_revision \
drawing_sheet_revision_usage \
externally_defined_class \
externally_defined_curve_font \
externally_defined_general_property \
externally_defined_hatch_style \
externally_defined_item \
externally_defined_item_relationship \
externally_defined_symbol \
externally_defined_text_font \
externally_defined_tile_style \
fill_area_style \
fill_area_style_colour \
fill_area_style_hatching \
fill_area_style_tile_symbol_with_style \
fill_area_style_tiles \
generic_character_glyph_symbol \
generic_literal \
geometrical_tolerance_callout \
hidden_element_over_riding_styled_item \
invisibility \
leader_curve \
leader_directed_callout \
leader_directed_dimension \
leader_terminator \
linear_dimension \
mechanical_design_and_draughting_relationship \
ordinate_dimension \
over_riding_styled_item \
point_style \
pre_defined_colour \
pre_defined_curve_font \
pre_defined_dimension_symbol \
pre_defined_geometrical_tolerance_symbol \
pre_defined_item \
pre_defined_marker \
pre_defined_point_marker_symbol \
pre_defined_surface_condition_symbol \
pre_defined_symbol \
pre_defined_terminator_symbol \
pre_defined_text_font \
presentation_area \
presentation_layer_assignment \
presentation_representation \
presentation_set \
presentation_size \
presentation_style_assignment \
presentation_style_by_context \
presentation_view \
projection_curve \
projection_directed_callout \
radius_dimension \
structured_dimension_callout \
styled_item \
surface_condition_callout \
surface_rendering_properties \
surface_side_style \
surface_style_boundary \
surface_style_control_grid \
surface_style_fill_area \
surface_style_parameter_line \
surface_style_reflectance_ambient \
surface_style_reflectance_ambient_diffuse \
surface_style_reflectance_ambient_diffuse_specular \
surface_style_rendering \
surface_style_rendering_with_properties \
surface_style_segmentation_curve \
surface_style_silhouette \
surface_style_transparent \
surface_style_usage \
surface_texture_representation \
symbol_colour \
symbol_representation \
symbol_representation_map \
symbol_style \
symbol_target \
terminator_symbol \
tessellated_annotation_occurrence \
text_literal \
text_literal_with_associated_curves \
text_literal_with_blanking_box \
text_literal_with_delineation \
text_literal_with_extent \
text_string_representation \
text_style \
text_style_for_defined_font \
text_style_with_box_characteristics \
text_style_with_mirror \
text_style_with_spacing \
vector_style \
view_volume \
]]

# -----------------------------------------------------------------------------------------------------
# STEP GD&T common

set entCategory(PR_STEP_TOLR) [lsort [list \
angular_location \
angular_size \
angularity_tolerance \
circular_runout_tolerance \
coaxiality_tolerance \
common_datum \
concentricity_tolerance \
cylindricity_tolerance \
datum \
datum_feature \
datum_reference \
datum_reference_compartment \
datum_reference_element \
datum_reference_modifier_with_value \
datum_system \
datum_target \
default_tolerance_table \
default_tolerance_table_cell \
dimensional_characteristic_representation \
dimensional_location \
dimensional_location_with_datum_feature \
dimensional_location_with_path \
dimensional_size \
dimensional_size_with_datum_feature \
dimensional_size_with_path \
directed_dimensional_location \
externally_defined_dimension_definition \
feature_for_datum_target_relationship \
flatness_tolerance \
geometric_tolerance \
geometric_tolerance_relationship \
geometric_tolerance_with_datum_reference \
geometric_tolerance_with_defined_area_unit \
geometric_tolerance_with_defined_unit \
geometric_tolerance_with_maximum_tolerance \
geometric_tolerance_with_modifiers \
limits_and_fits \
line_profile_tolerance \
modified_geometric_tolerance \
non_uniform_zone_definition \
parallelism_tolerance \
perpendicularity_tolerance \
placed_datum_target_feature \
placed_feature \
plus_minus_tolerance \
position_tolerance \
projected_zone_definition \
projected_zone_definition_with_offset \
referenced_modified_datum \
roundness_tolerance \
runout_zone_definition \
runout_zone_orientation \
runout_zone_orientation_reference_direction \
shape_dimension_representation \
straightness_tolerance \
surface_profile_tolerance \
symmetry_tolerance \
tolerance_value \
tolerance_zone \
tolerance_zone_definition \
tolerance_zone_form \
total_runout_tolerance \
type_qualifier \
unequally_disposed_geometric_tolerance \
value_format_type_qualifier \
]]

# -----------------------------------------------------------------------------------------------------
# STEP AP242 data quality

set entCategory(PR_STEP_AP242_QUAL) [lsort [list \
abrupt_change_of_surface_normal \
curve_with_excessive_segments curve_with_small_curvature_radius \
data_quality_assessment_measurement_association data_quality_assessment_specification data_quality_criteria_representation data_quality_criterion data_quality_criterion_assessment_association data_quality_criterion_measurement_association data_quality_definition data_quality_definition_relationship data_quality_definition_representation_relationship data_quality_inspection_criterion_report data_quality_inspection_criterion_report_item data_quality_inspection_instance_report data_quality_inspection_instance_report_item data_quality_inspection_report data_quality_inspection_result data_quality_inspection_result_representation data_quality_inspection_result_with_judgement data_quality_measurement_requirement data_quality_report_measurement_association data_quality_report_request \
disallowed_assembly_relationship_usage disconnected_face_set discontinuous_geometry \
edge_with_excessive_segments \
entirely_narrow_face entirely_narrow_solid entirely_narrow_surface \
erroneous_b_spline_curve_definition erroneous_b_spline_surface_definition erroneous_data erroneous_geometry erroneous_manifold_solid_brep erroneous_topology erroneous_topology_and_geometry_relationship \
excessive_use_of_groups excessive_use_of_layers excessively_high_degree_curve excessively_high_degree_surface \
externally_conditioned_data_quality_criteria_representation externally_conditioned_data_quality_criterion externally_conditioned_data_quality_inspection_instance_report_item externally_conditioned_data_quality_inspection_result externally_conditioned_data_quality_inspection_result_representation \
extreme_instance extreme_patch_width_variation \
face_surface_with_excessive_patches_in_one_direction free_edge \
g1_discontinuity_between_adjacent_faces g1_discontinuous_curve g1_discontinuous_surface g2_discontinuity_between_adjacent_faces g2_discontinuous_curve g2_discontinuous_surface \
gap_between_adjacent_edges_in_loop gap_between_edge_and_base_surface gap_between_faces_related_to_an_edge gap_between_pcurves_related_to_an_edge gap_between_vertex_and_base_surface gap_between_vertex_and_edge \
geometric_gap_in_topology geometry_with_local_irregularity geometry_with_local_near_degeneracy \
high_degree_axi_symmetric_surface high_degree_conic high_degree_linear_curve high_degree_planar_surface \
inappropriate_element_visibility inappropriate_use_of_layer \
inapt_data inapt_geometry inapt_manifold_solid_brep inapt_topology inapt_topology_and_geometry_relationship \
inconsistent_adjacent_face_normals inconsistent_curve_transition_code inconsistent_edge_and_curve_directions inconsistent_element_reference inconsistent_face_and_closed_shell_normals inconsistent_face_and_surface_normals inconsistent_surface_transition_code \
indistinct_curve_knots indistinct_surface_knots \
intersecting_connected_face_sets intersecting_loops_in_face intersecting_shells_in_solid \
multiply_defined_cartesian_points multiply_defined_curves multiply_defined_directions multiply_defined_edges multiply_defined_faces multiply_defined_geometry multiply_defined_placements multiply_defined_solids multiply_defined_surfaces multiply_defined_vertices \
narrow_surface_patch \
nearly_degenerate_geometry nearly_degenerate_surface_boundary nearly_degenerate_surface_patch \
non_agreed_accuracy_parameter_usage non_agreed_scale_usage non_agreed_unit_usage non_referenced_coordinate_system non_manifold_at_edge non_manifold_at_vertex non_smooth_geometry_transition_across_edge \
open_closed_shell open_edge_loop over_used_vertex overcomplex_geometry overcomplex_topology_and_geometry_relationship overlapping_geometry \
partly_overlapping_curves partly_overlapping_edges partly_overlapping_faces partly_overlapping_solids partly_overlapping_surfaces \
product_data_and_data_quality_relationship \
self_intersecting_curve self_intersecting_geometry self_intersecting_loop self_intersecting_shell self_intersecting_surface \
shape_data_quality_assessment_by_logical_test shape_data_quality_assessment_by_numerical_test shape_data_quality_criteria_representation shape_data_quality_criterion shape_data_quality_criterion_and_accuracy_association shape_data_quality_inspected_shape_and_result_relationship shape_data_quality_inspection_criterion_report shape_data_quality_inspection_instance_report shape_data_quality_inspection_instance_report_item shape_data_quality_inspection_result shape_data_quality_inspection_result_representation shape_data_quality_lower_value_limit shape_data_quality_upper_value_limit shape_data_quality_value_limit shape_data_quality_value_range \
short_length_curve short_length_curve_segment short_length_edge \
small_area_face small_area_surface small_area_surface_patch small_volume_solid \
software_for_data_quality_check \
solid_with_excessive_number_of_voids solid_with_wrong_number_of_voids \
steep_angle_between_adjacent_edges \
steep_angle_between_adjacent_faces steep_geometry_transition_across_edge \
surface_with_excessive_patches_in_one_direction surface_with_small_curvature_radius \
topology_related_to_multiply_defined_geometry topology_related_to_nearly_degenerate_geometry topology_related_to_overlapping_geometry topology_related_to_self_intersecting_geometry \
unused_patches unused_shape_element \
wrong_element_name wrongly_oriented_void wrongly_placed_loop wrongly_placed_void \
zero_surface_normal \
]]

# -----------------------------------------------------------------------------------------------------
# STEP AP242 constraint

set entCategory(PR_STEP_AP242_CONS) [lsort [list \
agc_with_dimension angle_assembly_constraint_with_dimension angle_geometric_constraint assembly_geometric_constraint \
binary_assembly_constraint \
cdgc_with_dimension clgc_with_dimension coaxial_assembly_constraint coaxial_geometric_constraint component_mating_constraint_condition curve_distance_geometric_constraint curve_length_geometric_constraint curve_smoothness_geometric_constraint \
defined_constraint \
equal_parameter_constraint explicit_constraint explicit_geometric_constraint \
fixed_constituent_assembly_constraint fixed_element_geometric_constraint free_form_constraint \
incidence_assembly_constraint incidence_geometric_constraint \
parallel_assembly_constraint parallel_assembly_constraint_with_dimension parallel_geometric_constraint parallel_offset_geometric_constraint pdgc_with_dimension perpendicular_assembly_constraint perpendicular_geometric_constraint pgc_with_dimension pogc_with_dimension point_distance_geometric_constraint \
radius_geometric_constraint rgc_with_dimension \
sdgc_with_dimension simultaneous_constraint_group skew_line_distance_geometric_constraint surface_distance_assembly_constraint_with_dimension surface_distance_geometric_constraint surface_smoothness_geometric_constraint swept_curve_surface_geometric_constraint swept_point_curve_geometric_constraint symmetry_geometric_constraint \
tangent_assembly_constraint tangent_geometric_constraint \
]]

# -----------------------------------------------------------------------------------------------------
# STEP AP242 kinematics

set entCategory(PR_STEP_AP242_KINE) [lsort [list \
actuated_kinematic_pair \
circular_path constrained_kinematic_motion_representation context_dependent_kinematic_link_representation curve_based_path curve_based_path_with_orientation curve_based_path_with_orientation_and_parameters cylindrical_pair_with_range \
free_kinematic_motion_representation \
gear_pair_with_range \
high_order_kinematic_pair \
interpolated_configuration_representation interpolated_configuration_segment item_link_motion_relationship \
kinematic_loop kinematic_path_defined_by_curves kinematic_path_defined_by_nodes kinematic_path_segment kinematic_property_definition_representation kinematic_property_mechanism_representation kinematic_property_topology_representation kinematic_topology_directed_structure kinematic_topology_network_structure kinematic_topology_structure kinematic_topology_substructure kinematic_topology_tree_structure \
linear_flexible_and_pinion_pair linear_flexible_and_planar_curve_pair linear_flexible_link_representation linear_path link_motion_relationship link_motion_representation_along_path link_motion_transformation low_order_kinematic_pair low_order_kinematic_pair_value low_order_kinematic_pair_with_motion_coupling low_order_kinematic_pair_with_range \
mechanism_representation mechanism_state_representation \
pair_representation_relationship path_node planar_pair_with_range point_on_planar_curve_pair_with_range point_on_planar_curve_pair_with_range point_on_surface_pair_with_range point_to_point_path prescribed_path prismatic_pair_with_range product_definition_kinematics product_definition_relationship_kinematics \
rack_and_pinion_pair_with_range revolute_pair_with_range rigid_link_representation \
screw_pair_with_range spherical_pair_with_pin spherical_pair_with_pin_and_range spherical_pair_with_range surface_pair_with_range \
universal_pair_with_range \
]]

# -----------------------------------------------------------------------------------------------------
# STEP AP242 math

set entCategory(PR_STEP_AP242_MATH) [lsort [list \
application_defined_function b_spline_basis b_spline_function cartesian_complex_number_region constant_function definite_integral_function elementary_function elementary_space explicit_table_function expression_denoted_function extended_tuple_space externally_listed_data finite_function finite_integer_interval finite_real_interval finite_space function_space general_linear_function homogeneous_linear_function \
imported_curve_function imported_point_function imported_surface_function imported_volume_function integer_interval_from_min integer_interval_to_max linearized_table_function listed_product_space maths_enum_literal maths_function maths_space maths_tuple_literal maths_variable parallel_composed_function partial_derivative_expression partial_derivative_function polar_complex_number_region rationalize_function real_interval_from_min real_interval_to_max reindexed_array_function repackaging_function restriction_function selector_function series_composed_function uniform_product_space \
]]

# -----------------------------------------------------------------------------------------------------
# STEP measure and unit (quantity)

set entCategory(PR_STEP_QUAN) [lsort [list \
absorbed_dose_measure_with_unit absorbed_dose_unit acceleration_measure_with_unit acceleration_unit amount_of_substance_measure_with_unit amount_of_substance_unit area_measure_with_unit area_unit \
capacitance_measure_with_unit capacitance_unit celsius_temperature_measure_with_unit conductance_measure_with_unit conductance_unit context_dependent_unit conversion_based_unit currency_measure_with_unit \
derived_unit derived_unit_element derived_unit_variable dielectric_constant_measure_with_unit dimensional_exponents dose_equivalent_measure_with_unit dose_equivalent_unit \
electric_charge_measure_with_unit electric_charge_unit electric_current_measure_with_unit electric_current_unit electric_potential_measure_with_unit electric_potential_unit energy_measure_with_unit energy_unit \
force_measure_with_unit force_unit frequency_measure_with_unit frequency_unit \
global_uncertainty_assigned_context global_unit_assigned_context \
illuminance_measure_with_unit illuminance_unit inductance_measure_with_unit inductance_unit \
length_measure_with_unit length_unit loss_tangent_measure_with_unit luminous_flux_measure_with_unit luminous_flux_unit luminous_intensity_measure_with_unit luminous_intensity_unit \
magnetic_flux_density_measure_with_unit magnetic_flux_density_unit magnetic_flux_measure_with_unit magnetic_flux_unit mass_measure_with_unit mass_unit measure_qualification measure_representation_item measure_with_unit \
named_unit named_unit_variable \
plane_angle_measure_with_unit plane_angle_unit power_measure_with_unit power_unit precision_qualifier pressure_measure_with_unit pressure_unit \
qualified_representation_item \
radioactivity_measure_with_unit radioactivity_unit ratio_measure_with_unit ratio_unit resistance_measure_with_unit resistance_unit \
si_absorbed_dose_unit \
si_capacitance_unit si_conductance_unit si_dose_equivalent_unit si_electric_charge_unit si_electric_potential_unit si_energy_unit si_force_unit si_frequency_unit si_illuminance_unit si_inductance_unit si_magnetic_flux_density_unit si_magnetic_flux_unit si_power_unit si_pressure_unit si_radioactivity_unit si_resistance_unit si_unit solid_angle_measure_with_unit solid_angle_unit \
thermal_resistance_measure_with_unit thermal_resistance_unit thermodynamic_temperature_measure_with_unit thermodynamic_temperature_unit time_measure_with_unit time_unit \
uncertainty_measure_with_unit \
velocity_measure_with_unit velocity_unit volume_measure_with_unit volume_unit \
binary_representation_item boolean_representation_item bytes_representation_item \
date_representation_item date_time_representation_item \
integer_representation_item descriptive_representation_item logical_representation_item real_representation_item value_representation_item \
]]

# -----------------------------------------------------------------------------------------------------
# STEP representation

set entCategory(PR_STEP_REP) [lsort [list \
advanced_brep_shape_representation auxiliary_geometric_representation_item \
characterized_item_within_representation \
compound_representation_item compound_shape_representation constructive_geometry_representation constructive_geometry_representation_relationship context_dependent_shape_representation csg_shape_representation curve_swept_solid_shape_representation \
definitional_representation document_representation_type \
edge_based_wireframe_shape_representation \
faceted_brep_shape_representation \
geometric_representation_context geometric_representation_item geometrically_bounded_2d_wireframe_representation geometrically_bounded_surface_shape_representation geometrically_bounded_wireframe_shape_representation \
hardness_representation \
item_defined_transformation item_identified_representation_usage \
manifold_subsurface_shape_representation manifold_surface_shape_representation material_property_representation mechanical_design_geometric_presentation_representation moments_of_inertia_representation \
non_manifold_surface_shape_representation null_representation_item \
parametric_representation_context picture_representation_item predefined_picture_representation_item presented_item_representation property_definition property_definition_relationship property_definition_representation \
rational_representation_item representation representation_context representation_item representation_map representation_relationship representation_relationship_with_transformation row_representation_item \
shape_definition_representation shape_representation shape_representation_relationship shape_representation_with_parameters \
table_representation_item tactile_appearance_representation tessellated_shape_representation topological_representation_item \
uncertainty_assigned_representation \
value_range variational_representation_item visual_appearance_representation \
]]

# -----------------------------------------------------------------------------------------------------
# STEP AP238
set entCategory(PR_STEP_AP238) [lsort [list \
action_method_with_associated_documents action_resource_relationship action_resource_requirement_relationship back_boring_operation block_shape_representation boring_operation bottom_and_side_milling_operation concurrent_action_method contouring_turning_operation cylindrical_shape_representation drilling_operation drilling_type_operation drilling_type_strategy expression_representation_item externally_defined_representation_with_parameters facing_turning_operation freeform_milling_operation freeform_milling_strategy freeform_milling_tolerance_representation grooving_turning_operation knurling_turning_operation machining_adaptive_control_relationship machining_approach_retract_strategy machining_cutting_component machining_cutting_corner_representation machining_dwell_time_representation machining_execution_resource machining_feature_process machining_feature_relationship machining_feature_sequence_relationship machining_feed_speed_representation machining_final_feature_relationship machining_functions machining_functions_relationship machining_nc_function machining_offset_vector_representation machining_operation machining_operation_relationship machining_operator_instruction machining_operator_instruction_relationship machining_process_body_relationship machining_process_branch_relationship machining_process_concurrent_relationship machining_process_executable machining_process_model machining_process_model_relationship machining_process_sequence_relationship machining_project machining_project_workpiece_relationship machining_rapid_movement machining_setup machining_setup_workpiece_relationship machining_spindle_speed_representation machining_strategy machining_strategy_relationship machining_technology machining_technology_relationship machining_tool machining_tool_body_representation machining_tool_direction_representation machining_tool_usage machining_toolpath machining_toolpath_sequence_relationship machining_toolpath_speed_profile_representation machining_touch_probing machining_workingstep machining_workplan milling_type_operation milling_type_strategy ngon_shape_representation plane_milling_operation sequential_method serial_action_method side_milling_operation tapping_operation threading_turning_operation turning_type_operation turning_type_strategy \
]]

# -----------------------------------------------------------------------------------------------------
# STEP AP209
set entCategory(PR_STEP_AP209) [lsort [list \
aligned_axis_tolerance aligned_curve_3d_element_coordinate_system aligned_surface_2d_element_coordinate_system aligned_surface_3d_element_coordinate_system analysis_item_within_representation analysis_message analysis_step arbitrary_volume_2d_element_coordinate_system arbitrary_volume_3d_element_coordinate_system axisymmetric_2d_element_property axisymmetric_curve_2d_element_descriptor axisymmetric_curve_2d_element_representation axisymmetric_surface_2d_element_descriptor axisymmetric_surface_2d_element_representation axisymmetric_volume_2d_element_descriptor axisymmetric_volume_2d_element_representation calculated_state constant_surface_3d_element_coordinate_system constraint_element control control_analysis_step control_linear_modes_and_frequencies_analysis_step control_linear_modes_and_frequencies_process control_linear_static_analysis_step control_linear_static_analysis_step_with_harmonic control_linear_static_load_increment_process control_process control_result_relationship curve_2d_element_basis curve_2d_element_constant_specified_variable_value curve_2d_element_constant_specified_volume_variable_value curve_2d_element_coordinate_system curve_2d_element_field_variable_definition curve_2d_element_group curve_2d_element_integrated_matrix curve_2d_element_integrated_matrix_with_definition curve_2d_element_integration curve_2d_element_location_point_variable_values curve_2d_element_location_point_volume_variable_values curve_2d_element_property curve_2d_element_value_and_location curve_2d_element_value_and_volume_location curve_2d_node_field_aggregated_variable_values curve_2d_node_field_section_variable_values curve_2d_node_field_variable_definition curve_2d_substructure_element_reference curve_2d_whole_element_variable_value curve_3d_element_basis curve_3d_element_constant_specified_variable_value curve_3d_element_constant_specified_volume_variable_value curve_3d_element_descriptor curve_3d_element_field_variable_definition curve_3d_element_group curve_3d_element_integrated_matrix curve_3d_element_integrated_matrix_with_definition curve_3d_element_integration curve_3d_element_length_integration_explicit curve_3d_element_length_integration_rule curve_3d_element_location_point_variable_values curve_3d_element_location_point_volume_variable_values curve_3d_element_nodal_specified_variable_values curve_3d_element_position_weight curve_3d_element_property curve_3d_element_representation curve_3d_element_value_and_location curve_3d_element_value_and_volume_location curve_3d_node_field_aggregated_variable_values curve_3d_node_field_section_variable_values curve_3d_node_field_variable_definition curve_3d_substructure_element_reference curve_3d_whole_element_variable_value curve_constraint curve_element_end_offset curve_element_end_release curve_element_end_release_packet curve_element_interval curve_element_interval_constant curve_element_interval_linearly_varying curve_element_location curve_element_section_definition curve_element_section_derived_definitions curve_freedom_action_definition curve_freedom_and_value_definition curve_freedom_values curve_section_element_location curve_section_integration_explicit curve_volume_element_location cylindrical_point cylindrical_symmetry_control data_environment_relationship direction_node directionally_explicit_element_coefficient directionally_explicit_element_coordinate_system_aligned directionally_explicit_element_coordinate_system_arbitrary directionally_explicit_element_representation document_with_class dummy_node element_analysis_message element_definition element_descriptor element_geometric_relationship element_group element_group_analysis_message element_material element_nodal_freedom_actions element_nodal_freedom_terms element_representation element_sequence euler_angles explicit_element_matrix explicit_element_representation fea_area_density fea_axis2_placement_2d fea_axis2_placement_3d fea_curve_section_geometric_relationship fea_group fea_group_relation fea_linear_elasticity fea_mass_density fea_material_property_geometric_relationship fea_material_property_representation fea_material_property_representation_item fea_model fea_model_2d fea_model_3d fea_model_definition fea_moisture_absorption fea_parametric_point fea_representation_item fea_secant_coefficient_of_linear_thermal_expansion fea_shell_bending_stiffness fea_shell_membrane_bending_coupling_stiffness fea_shell_membrane_stiffness fea_shell_shear_stiffness fea_surface_section_geometric_relationship fea_tangential_coefficient_of_linear_thermal_expansion field_variable_definition field_variable_element_definition field_variable_element_group_value field_variable_node_definition field_variable_whole_model_value freedom_and_coefficient freedoms_list geometric_node grounded_damper grounded_spring linear_constraint_equation_element linear_constraint_equation_element_value linear_constraint_equation_nodal_term linearly_superimposed_state no_symmetry_control nodal_dof_reduction nodal_freedom_action_definition nodal_freedom_and_value_definition nodal_freedom_values node node_analysis_message node_definition node_geometric_relationship node_group node_representation node_sequence node_set node_with_solution_coordinate_system node_with_vector output_request_state parametric_curve_3d_element_coordinate_direction parametric_curve_3d_element_coordinate_system parametric_surface_2d_element_coordinate_system parametric_surface_3d_element_coordinate_system parametric_volume_2d_element_coordinate_system parametric_volume_3d_element_coordinate_system plane_2d_element_property plane_curve_2d_element_descriptor plane_curve_2d_element_representation plane_surface_2d_element_descriptor plane_surface_2d_element_representation plane_volume_2d_element_descriptor plane_volume_2d_element_representation point_constraint point_element_matrix point_element_representation point_freedom_action_definition point_freedom_and_value_definition point_freedom_values point_representation result result_analysis_step result_linear_modes_and_frequencies_analysis_sub_step result_linear_static_analysis_sub_step retention_assignment simple_plane_2d_element_property single_point_constraint_element single_point_constraint_element_values solid_constraint solid_freedom_action_definition solid_freedom_and_value_definition solid_freedom_values specified_state spherical_point state state_component state_definition state_relationship state_with_harmonic stationary_mass structural_response_property structural_response_property_definition_representation substructure_element_representation substructure_node_reference substructure_node_relationship surface_2d_element_basis surface_2d_element_boundary_constant_specified_surface_variable_value surface_2d_element_boundary_constant_specified_variable_value surface_2d_element_boundary_edge_constant_specified_surface_variable_value surface_2d_element_boundary_edge_constant_specified_variable_value surface_2d_element_boundary_edge_location_point_surface_variable_values surface_2d_element_boundary_edge_location_point_variable_values surface_2d_element_boundary_edge_nodal_specified_variable_values surface_2d_element_boundary_edge_whole_edge_variable_value surface_2d_element_boundary_location_point_surface_variable_values surface_2d_element_boundary_nodal_specified_variable_values surface_2d_element_boundary_whole_face_variable_value surface_2d_element_constant_specified_variable_value surface_2d_element_constant_specified_volume_variable_value surface_2d_element_field_variable_definition surface_2d_element_group surface_2d_element_integrated_matrix surface_2d_element_integrated_matrix_with_definition surface_2d_element_integration surface_2d_element_length_integration_explicit surface_2d_element_length_integration_rule surface_2d_element_location_point_variable_values surface_2d_element_location_point_volume_variable_values surface_2d_element_nodal_specified_variable_values surface_2d_element_value_and_location surface_2d_element_value_and_volume_location surface_2d_node_field_aggregated_variable_values surface_2d_node_field_section_variable_values surface_2d_node_field_variable_definition surface_2d_substructure_element_reference surface_2d_whole_element_variable_value surface_3d_element_basis surface_3d_element_boundary_constant_specified_surface_variable_value surface_3d_element_boundary_constant_specified_variable_value surface_3d_element_boundary_edge_constant_specified_surface_variable_value surface_3d_element_boundary_edge_constant_specified_variable_value surface_3d_element_boundary_edge_location_point_surface_variable_values surface_3d_element_boundary_edge_location_point_variable_values surface_3d_element_boundary_edge_nodal_specified_variable_values surface_3d_element_boundary_edge_whole_edge_variable_value surface_3d_element_boundary_location_point_surface_variable_values surface_3d_element_boundary_nodal_specified_variable_values surface_3d_element_boundary_whole_face_variable_value surface_3d_element_constant_specified_variable_value surface_3d_element_constant_specified_volume_variable_value surface_3d_element_descriptor surface_3d_element_field_integration_explicit surface_3d_element_field_integration_rule surface_3d_element_field_variable_definition surface_3d_element_group surface_3d_element_integrated_matrix surface_3d_element_integrated_matrix_with_definition surface_3d_element_integration surface_3d_element_location_point_variable_values surface_3d_element_location_point_volume_variable_values surface_3d_element_nodal_specified_variable_values surface_3d_element_representation surface_3d_element_value_and_location surface_3d_element_value_and_volume_location surface_3d_node_field_aggregated_variable_values surface_3d_node_field_section_variable_values surface_3d_node_field_variable_definition surface_3d_substructure_element_reference surface_3d_whole_element_variable_value surface_constraint surface_element_location surface_element_property surface_freedom_action_definition surface_freedom_and_value_definition surface_freedom_values surface_position_weight surface_section surface_section_element_location surface_section_element_location_absolute surface_section_element_location_dimensionless surface_section_field surface_section_field_constant surface_section_field_varying surface_section_integration_explicit surface_section_integration_rule surface_section_position_weight surface_volume_element_location symmetry_control system_and_freedom tensor_representation_item uniform_surface_section uniform_surface_section_layered volume_2d_element_basis volume_2d_element_boundary_constant_specified_variable_value volume_2d_element_boundary_edge_constant_specified_volume_variable_value volume_2d_element_boundary_edge_location_point_volume_variable_values volume_2d_element_boundary_edge_nodal_specified_variable_values volume_2d_element_boundary_edge_whole_edge_variable_value volume_2d_element_boundary_location_point_variable_values volume_2d_element_boundary_nodal_specified_variable_values volume_2d_element_boundary_whole_face_variable_value volume_2d_element_constant_specified_variable_value volume_2d_element_field_integration_explicit volume_2d_element_field_integration_rule volume_2d_element_field_variable_definition volume_2d_element_group volume_2d_element_integrated_matrix volume_2d_element_integrated_matrix_with_definition volume_2d_element_location_point_variable_values volume_2d_element_nodal_specified_variable_values volume_2d_element_value_and_location volume_2d_node_field_variable_definition volume_2d_substructure_element_reference volume_2d_whole_element_variable_value volume_3d_element_basis volume_3d_element_boundary_constant_specified_variable_value volume_3d_element_boundary_edge_constant_specified_volume_variable_value volume_3d_element_boundary_edge_location_point_volume_variable_values volume_3d_element_boundary_edge_nodal_specified_variable_values volume_3d_element_boundary_edge_whole_edge_variable_value volume_3d_element_boundary_location_point_variable_values volume_3d_element_boundary_nodal_specified_variable_values volume_3d_element_boundary_whole_face_variable_value volume_3d_element_constant_specified_variable_value volume_3d_element_descriptor volume_3d_element_field_integration_explicit volume_3d_element_field_integration_rule volume_3d_element_field_variable_definition volume_3d_element_group volume_3d_element_integrated_matrix volume_3d_element_integrated_matrix_with_definition volume_3d_element_location_point_variable_values volume_3d_element_nodal_specified_variable_values volume_3d_element_representation volume_3d_element_value_and_location volume_3d_node_field_variable_definition volume_3d_substructure_element_reference volume_3d_whole_element_variable_value volume_element_location volume_position_weight whole_model_analysis_message whole_model_modes_and_frequencies_analysis_message \
]]

# -----------------------------------------------------------------------------------------------------
# STEP AP210
set entCategory(PR_STEP_AP210) [lsort [list \
across_port_variable additive_laminate_text_component aggregate_connectivity_requirement allocated_passage_minimum_annular_ring altered_package_terminal analog_analytical_model_port analog_port_variable analytical_model_definition analytical_model_make_from_relationship analytical_model_parameter analytical_model_port analytical_model_port_assignment analytical_model_scalar_port analytical_model_vector_port analytical_representation annotation_point_occurrence annotation_table annotation_table_occurrence annotation_text_with_associated_curves annotation_text_with_blanking_box annotation_text_with_delineation annotation_text_with_extent area_component area_qualified_layout_spacing_requirement assembly_group_spacing_requirement assembly_item_number assembly_module_component assembly_module_design_view assembly_module_interface_terminal assembly_module_macro_component assembly_module_macro_component_join_terminal assembly_module_macro_terminal assembly_module_terminal assembly_module_usage_view assembly_shield_allocation assembly_spacing_requirement assembly_to_part_connectivity_structure_allocation auxiliary_characteristic_dimension_representation axis_placement_2d_3d_mapping bare_die bare_die_bottom_surface bare_die_component bare_die_edge_segment_surface bare_die_edge_surface bare_die_surface bare_die_template_terminal bare_die_terminal bare_die_top_surface basic_multi_stratum_printed_component basic_multi_stratum_printed_part_template blind_passage_template blind_via block_terminal_schematic_symbol_callout breakout_footprint_definition breakout_occurrence buried_via bus_element_link bus_structural_definition cable_component cable_terminal cable_usage_view category_model_parameter complex_passage_padstack_definition component_2d_location component_3d_location component_functional_terminal component_functional_unit component_material_relationship component_material_relationship_assignment component_mounting_feature component_part_2d_non_planar_geometric_representation_relationship component_termination_passage component_termination_passage_template composite_array_shape_aspect composite_array_shape_aspect_link composite_sequential_text_reference composite_signal_property_relationship conductive_interconnect_element conductive_interconnect_element_terminal_link conductive_interconnect_element_with_pre_defined_transitions connected_area_component connection_zone_based_fabrication_joint connection_zone_map_identification connector_based_interconnect_definition contact_size_dependent_land continuous_template coordinated_geometric_relationship_with_2d_3d_placement_transformation copy_stratum_technology_occurrence_relationship counterbore_passage_template countersunk_passage_template cutout cutout_edge_segment datum_difference datum_difference_based_characteristic datum_difference_based_model_parameter datum_difference_functional_unit_usage_view_terminal_assignment default_attachment_size_based_land_physical_template default_passage_based_land_physical_template default_plated_passage_based_land_physical_template default_trace_template default_unsupported_passage_based_land_physical_template default_value_property_definition_representation defined_table dependent_electrical_isolation_removal_component dependent_electrical_isolation_removal_template dependent_thermal_isolation_removal_component dependent_thermal_isolation_removal_template derived_laminate_assignment derived_stratum derived_stratum_technology_occurrence_relationship design_composition_path design_layer_stratum design_specific_stratum_technology_mapping_relationship design_stack_model device_terminal_map dielectric_crossover_area dielectric_material_passage digital_analytical_model_port digital_analytical_model_scalar_port digital_analytical_model_vector_port dimensional_size_property direct_stratum_component_join_implementation documentation_layer_stratum edge_segment_cross_section edge_segment_vertex electrical_isolation_laminate_component electrical_isolation_removal_template electrical_network electromagnetic_compatibility_requirement_allocation equivalent_stackup_model_definition equivalent_sub_stack_definition explicit_text_reference explicit_text_reference_occurrence externally_defined_physical_network_group externally_defined_physical_network_group_element_relationship fabrication_joint fiducial fiducial_part_feature fiducial_stratum_feature fill_area_template filled_via footprint_definition footprint_library_stratum_technology footprint_occurrence footprint_occurrence_product_definition_relationship functional_specification functional_specification_definition functional_terminal_group functional_unit functional_unit_terminal_definition generic_footprint_definition generic_laminate_text_component geometric_template geometric_tolerance_group group_product_definition group_shape_aspect guided_wave_terminal hatch_area_template hatch_line_element impedance_measurement_setup_requirement impedance_requirement implicit_text_reference indirect_stratum_component_join_implementation integral_shield inter_stratum_feature inter_stratum_feature_dependent_land inter_stratum_feature_edge_segment_template inter_stratum_feature_edge_segment_template_with_cross_section inter_stratum_feature_edge_template inter_stratum_feature_template interconnect_module_component interconnect_module_component_surface_feature interconnect_module_cutout_segment_surface interconnect_module_design_object_category interconnect_module_design_view interconnect_module_edge interconnect_module_edge_segment interconnect_module_edge_segment_surface interconnect_module_interface_terminal interconnect_module_macro_component interconnect_module_macro_component_join_terminal interconnect_module_macro_terminal interconnect_module_stratum_based_terminal interconnect_module_terminal interconnect_module_usage_view interconnect_shield_allocation interface_access_component_definition interface_access_material_removal_laminate_component interface_access_stratum_feature_template_component interface_component interface_mounted_join interface_plane interfacial_connection internal_probe_access_area inverse_copy_stratum_technology_occurrence_relationship item_restricted_requirement join_shape_aspect keepout_design_object_category laminate_component laminate_component_feature laminate_component_interface_terminal laminate_component_join_terminal laminate_group_component_make_from_relationship laminate_text_string_component land land_physical_template land_template_terminal land_with_join_terminal layer_connection_point layer_qualified_layout_spacing_requirement layer_stack_region layered_assembly_module_design_view layered_assembly_module_usage_view layered_assembly_panel_design_view layered_interconnect_module_design_view layered_interconnect_module_usage_view layered_interconnect_panel_design_view layout_junction layout_macro_component layout_macro_definition layout_macro_definition_terminal_to_usage_terminal_assignment layout_macro_floor_plan_template layout_spacing_contextual_area layout_spacing_requirement length_trimmed_terminal library_stack_model library_to_design_stack_model_mapping linear_composite_array_shape_aspect linear_composite_array_shape_aspect_link linear_profile_tolerance local_linear_stack make_from_connectivity_relationship make_from_functional_unit_terminal_definition_relationship make_from_model_port_relationship make_from_part_feature_relationship manifold_constraining_context_dependent_shape_representation material_designation_with_conductivity_classification material_electrical_conductivity_category material_removal_feature_template material_removal_laminate_component material_removal_laminate_text_component material_removal_structured_component material_removal_structured_template minimally_defined_bare_die_terminal minimally_defined_connector model_parameter model_parameter_with_unit mounting_restriction_area mounting_restriction_volume multi_layer_component_definition multi_layer_material_removal_laminate_component multi_layer_stratum_feature_template_component multi_stratum_printed_component multi_stratum_printed_part_template multi_stratum_special_symbol_component multi_stratum_special_symbol_template multi_stratum_structured_template network_node_definition next_assembly_usage_occurrence_relationship non_conductive_base_blind_via non_conductive_cross_section_template operational_requirement_relationship opposing_boundary_dimensional_size package package_body package_body_bottom_surface package_body_edge_segment_surface package_body_edge_surface package_body_surface package_body_top_surface package_footprint_relationship_definition package_terminal package_terminal_template_definition packaged_component packaged_connector packaged_connector_component packaged_connector_terminal_relationship packaged_part packaged_part_terminal padstack_definition padstack_occurrence padstack_occurrence_product_definition_relationship parameter_assignment parameter_assignment_override parametric_template part_connected_terminals_definition part_connected_terminals_definition_domain part_connected_terminals_element part_connected_terminals_layout_topology_requirement_assignment part_connected_terminals_structure_definition part_feature_template_definition part_interface_access_feature part_level_schematic_symbol_representation part_mating_feature part_mounting_feature part_string_template part_template_definition part_template_keepout_shape_allocation_to_stratum_stack part_terminal_external_reference part_terminal_schematic_symbol_callout part_text_template part_tooling_feature partially_plated_cutout partially_plated_interconnect_module_edge passage_deposition_material_identification passage_filling_material_identification passage_padstack_definition passage_technology passage_technology_allocation_to_stack_model passage_terminal_based_fabrication_joint_link physical_component_interface_terminal physical_connectivity_definition physical_connectivity_definition_domain physical_connectivity_element physical_connectivity_interrupting_cutout physical_connectivity_layout_topology_link physical_connectivity_layout_topology_node physical_connectivity_layout_topology_requirement physical_connectivity_layout_topology_requirement_assignment physical_network physical_network_group physical_network_group_element_relationship physical_network_supporting_inter_stratum_feature physical_node_branch_requirement_to_implementing_component_allocation physical_node_requirement_to_implementing_component_allocation physical_shield physical_unit physical_unit_datum_feature physical_unit_geometric_tolerance physical_unit_interconnect_definition physical_unit_keepout_shape_allocation_to_stratum_stack physical_unit_keepout_shape_allocation_to_stratum_technology physical_unit_network_definition planar_closed_path_shape_representation_with_parameters planar_path_shape_representation_with_parameters plated_conductive_base_blind_via plated_cutout plated_cutout_edge_segment plated_inter_stratum_feature plated_interconnect_module_edge plated_interconnect_module_edge_segment plated_passage plated_passage_dependent_land port_variable positional_boundary positional_boundary_member pre_defined_parallel_datum_axis_symbol_3d_2d_relationship pre_defined_perpendicular_datum_axis_symbol_3d_2d_relationship pre_defined_perpendicular_datum_plane_symbol_3d_2d_relationship pre_defined_physical_network_group pre_defined_physical_network_group_element_relationship predefined_requirement_view_definition primary_orientation_feature primary_reference_terminal primary_stratum_indicator_symbol printed_component printed_connector_component printed_connector_template printed_connector_template_terminal_relationship printed_part_cross_section_template printed_part_cross_section_template_terminal printed_part_template printed_part_template_connected_terminals_definition printed_part_template_material printed_part_template_material_link printed_part_template_terminal printed_part_template_terminal_connection_zone_category printed_tiebar_template probe_access_area product_specific_parameter_value_assignment protocol_physical_layer_definition protocol_physical_layer_definition_with_characterization protocol_requirement_allocation_to_part_terminal rectangular_composite_array_shape_aspect reference_composition_path reference_graphic_registration_mark region_based_derived_pattern routed_interconnect_component routed_physical_component routed_physical_shield routed_shield routed_transmission_line scalar_terminal_definition_link schema_based_model_parameter schematic_symbol_callout schematic_symbol_representation seating_plane secondary_orientation_feature sequential_laminate_passage_based_fabrication_joint sequential_laminate_stackup_component sequential_laminate_stackup_definition shape_definition_3d_intersection shape_formed_terminal signal signal_category single_stratum_continuous_template single_stratum_printed_component single_stratum_printed_part_template single_stratum_special_symbol_component single_stratum_special_symbol_template single_stratum_structured_template single_stratum_template snowball_template solid_character_glyph_2d_symbol special_symbol_template statistical_dimensional_location statistical_dimensional_size statistical_geometric_tolerance stratum stratum_feature stratum_feature_based_derived_pattern stratum_feature_conductive_join stratum_feature_template stratum_feature_template_component stratum_feature_template_component_with_stratum_feature stratum_specific_template_location stratum_stack_dependent_template stratum_stack_model stratum_sub_stack stratum_surface stratum_surface_technology stratum_technology stratum_technology_mapping_relationship stratum_technology_occurrence stratum_technology_occurrence_feature_constraint stratum_technology_occurrence_link stratum_technology_occurrence_relationship stratum_technology_occurrence_swap_relationship stratum_technology_swap_relationship structured_inter_stratum_feature_template structured_layout_component structured_layout_component_sub_assembly_relationship structured_layout_component_sub_assembly_relationship_with_component structured_printed_part_template structured_printed_part_template_terminal structured_template surface_prepped_terminal symbol_representation_relationship symbol_representation_with_blanking_box table_record_field_representation table_record_field_representation_with_clipping_box table_record_representation table_representation table_representation_relationship table_text_relationship teardrop_by_angle_template teardrop_by_length_template teardrop_template template_material_cross_section_boundary terminal_schematic_symbol_callout tertiary_orientation_feature test_method_based_parameter_assignment test_point_part_feature thermal_feature thermal_isolation_removal_component thermal_isolation_removal_template thermal_network thermal_network_node_definition thermal_requirement_allocation through_port_variable tiebar_printed_component tile_area_template tolerance_zone_boundary tolerance_zone_explicit_opposing_boundary_set tolerance_zone_implicit_opposing_boundary_set tool_registration_mark trace_template transform_port_variable unplated_cutout_edge_segment unplated_interconnect_module_edge_segment unrouted_conductive_interconnect_element unsupported_passage unsupported_passage_dependent_land unsupported_passage_template usage_concept_usage_relationship usage_view_connection_zone_terminal_shape_relationship valid_range_property_definition_representation via via_template viewing_plane visual_orientation_feature wire_terminal wire_terminal_template_definition \
]]
}
