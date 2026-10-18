# ZAMF_FIELD_AUTOMATION: SAP GUI Report (PERFORM and Class Versions)

Reads an Excel (.xlsx) file of SAP custom field information, applies two business rules, and writes the result as an Excel (.xlsx) file.

**Rule 1 (AMF Scope):** if the S/4HANA release is <= 2023 and `Data Type` is QUAN, CURR, CUKY or UNIT, set `AMF Scope` = `OUT`.

**Rule 2 (Proposed Append Name):**
- `APPEND_STRUCTURE`: `<Enhancement Name>_AMF` (ZST_MARA -> ZST_MARA_AMF)
- `INCLUDE_STRUCTURE`: `Z<Enhancement Name>_AMF` (CI_EKKODB -> ZCI_EKKODB_AMF)
- `CUSTOMIZING_INCLUDE`: no rule defined yet, so the name is left blank.

> **ASSUMED:** the input/output column layout (`ty_data`, header map, output headings) is a typical layout, because the real column list was not provided. Adjust these three places: the `ty_data` type, the header map, and the output headings.

---

## Common setup (both versions)

### 1. Dictionary objects (SE11)

| Object | Name | Details |
|---|---|---|
| Domain | `ZAMF_S4_REL` | Data type `NUMC`, length 4. Value Range tab, fixed values: 1511, 1610, 1709, 1809, 1909, 2020, 2021, 2022, 2023, 2024, 2025 (with short texts) |
| Data element | `ZAMF_S4_REL` | Domain `ZAMF_S4_REL`, field labels "S/4 Release" |

NUMC4 allows a numeric comparison with `2023` for all releases.

### 2. Program and includes (SE38)

Create executable program `ZAMF_FIELD_AUTOMATION`, then create the three includes by double-clicking each INCLUDE statement:

- `ZAMF_FIELD_AUTOMATION_TOP`
- `ZAMF_FIELD_AUTOMATION_SEL`
- `ZAMF_FIELD_AUTOMATION_PROC`

### 3. Text elements

- `TEXT-001`: "S/4HANA Release"
- `TEXT-002`: "Files"
- Selection texts: `P_REL` = "S/4HANA release version", `P_FILE` = "Input Excel file (.xlsx)", `P_OUT` = "Output Excel file (.xlsx)"

### Include `ZAMF_FIELD_AUTOMATION_SEL` (identical in both versions)

```abap
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-001.
  PARAMETERS p_rel TYPE zamf_s4_rel AS LISTBOX VISIBLE LENGTH 12
                   OBLIGATORY DEFAULT '2023'.
SELECTION-SCREEN END OF BLOCK b1.

SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-002.
  PARAMETERS: p_file TYPE rlgrap-filename OBLIGATORY,   " input .xlsx
              p_out  TYPE rlgrap-filename OBLIGATORY.   " output .xlsx
SELECTION-SCREEN END OF BLOCK b2.
```

---

# PART A: PERFORM (procedural) version

### Main program: `ZAMF_FIELD_AUTOMATION`

```abap
*&---------------------------------------------------------------------*
*& Report ZAMF_FIELD_AUTOMATION
*& Reads custom-field Excel, sets AMF Scope / Proposed Append Name,
*& writes result as Excel
*&---------------------------------------------------------------------*
REPORT zamf_field_automation.

INCLUDE zamf_field_automation_top.
INCLUDE zamf_field_automation_sel.
INCLUDE zamf_field_automation_proc.

INITIALIZATION.
  PERFORM init_defaults.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  PERFORM f4_input_file.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_out.
  PERFORM f4_output_file.

AT SELECTION-SCREEN.
  PERFORM validate_input.

START-OF-SELECTION.
  PERFORM read_excel.
  PERFORM process_data.
  PERFORM download_excel.
```

### Include `ZAMF_FIELD_AUTOMATION_TOP`

```abap
*&---------------------------------------------------------------------*
*& Include ZAMF_FIELD_AUTOMATION_TOP
*&---------------------------------------------------------------------*
TYPES: BEGIN OF ty_data,                 " ASSUMED layout - adjust to yours
         object_name  TYPE c LENGTH 30,  " Table / structure name
         field_name   TYPE c LENGTH 30,
         enh_type     TYPE c LENGTH 30,  " Enhancement Type
         enh_name     TYPE c LENGTH 30,  " Enhancement Name
         data_element TYPE c LENGTH 30,
         domain       TYPE c LENGTH 30,
         data_type    TYPE c LENGTH 10,  " Data Type
         length       TYPE c LENGTH 10,
         decimals     TYPE c LENGTH 10,
         bus_context  TYPE c LENGTH 255, " Business Context
         " ---- output-only columns ----
         amf_scope    TYPE c LENGTH 10,  " AMF Scope
         prop_append  TYPE c LENGTH 60,  " Proposed Append Name
       END OF ty_data.

" Excel header text (spaces/underscores removed, upper case) -> field
TYPES: BEGIN OF ty_map,
         header    TYPE string,
         fieldname TYPE fieldname,
       END OF ty_map.

" Excel column index -> field
TYPES: BEGIN OF ty_colidx,
         idx       TYPE i,
         fieldname TYPE fieldname,
       END OF ty_colidx.

DATA: gt_data   TYPE STANDARD TABLE OF ty_data,
      gt_map    TYPE STANDARD TABLE OF ty_map,
      gt_colidx TYPE STANDARD TABLE OF ty_colidx,
      gv_warn   TYPE i.                  " names longer than 30 chars

" Quantity / currency data types
DATA gr_qty_curr TYPE RANGE OF char10.

CONSTANTS: gc_cutoff_rel TYPE zamf_s4_rel VALUE '2023',
           gc_out        TYPE c LENGTH 3  VALUE 'OUT',
           gc_append     TYPE string      VALUE 'APPEND_STRUCTURE',
           gc_include    TYPE string      VALUE 'INCLUDE_STRUCTURE',
           gc_custom_inc TYPE string      VALUE 'CUSTOMIZING_INCLUDE',
           gc_suffix     TYPE string      VALUE '_AMF',
           gc_prefix     TYPE string      VALUE 'Z'.
```

### Include `ZAMF_FIELD_AUTOMATION_PROC`

```abap
*&---------------------------------------------------------------------*
*& Include ZAMF_FIELD_AUTOMATION_PROC
*&---------------------------------------------------------------------*

*&---------------------------------------------------------------------*
FORM init_defaults.
  " Header text (normalised) -> structure field. ADAPT to your columns.
  gt_map = VALUE #(
    ( header = 'OBJECTNAME'         fieldname = 'OBJECT_NAME'  )
    ( header = 'FIELDNAME'          fieldname = 'FIELD_NAME'   )
    ( header = 'ENHANCEMENTTYPE'    fieldname = 'ENH_TYPE'     )
    ( header = 'ENHANCEMENTNAME'    fieldname = 'ENH_NAME'     )
    ( header = 'DATAELEMENT'        fieldname = 'DATA_ELEMENT' )
    ( header = 'DOMAIN'             fieldname = 'DOMAIN'       )
    ( header = 'DATATYPE'           fieldname = 'DATA_TYPE'    )
    ( header = 'LENGTH'             fieldname = 'LENGTH'       )
    ( header = 'DECIMALS'           fieldname = 'DECIMALS'     )
    ( header = 'BUSINESSCONTEXT'    fieldname = 'BUS_CONTEXT'  ) ).

  gr_qty_curr = VALUE #( sign = 'I' option = 'EQ'
                         ( low = 'QUAN' ) ( low = 'CURR' )
                         ( low = 'CUKY' ) ( low = 'UNIT' ) ).
ENDFORM.

*&---------------------------------------------------------------------*
FORM f4_input_file.
  DATA: lt_files TYPE filetable,
        lv_rc    TYPE i,
        lv_act   TYPE i.

  cl_gui_frontend_services=>file_open_dialog(
    EXPORTING window_title = 'Select input Excel file'
              default_extension = 'xlsx'
              file_filter = 'Excel (*.xlsx)|*.xlsx'
    CHANGING  file_table = lt_files
              rc         = lv_rc
              user_action = lv_act
    EXCEPTIONS OTHERS = 1 ).
  IF sy-subrc = 0 AND lv_act = cl_gui_frontend_services=>action_ok
     AND lines( lt_files ) > 0.
    p_file = lt_files[ 1 ]-filename.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
FORM f4_output_file.
  DATA: lv_name TYPE string,
        lv_path TYPE string,
        lv_full TYPE string,
        lv_act  TYPE i.

  cl_gui_frontend_services=>file_save_dialog(
    EXPORTING window_title = 'Save output Excel file'
              default_extension = 'xlsx'
              default_file_name = 'AMF_Output.xlsx'
              file_filter = 'Excel (*.xlsx)|*.xlsx'
    CHANGING  filename = lv_name
              path     = lv_path
              fullpath = lv_full
              user_action = lv_act
    EXCEPTIONS OTHERS = 1 ).
  IF sy-subrc = 0 AND lv_act = cl_gui_frontend_services=>action_ok.
    p_out = lv_full.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
FORM validate_input.
  IF to_lower( p_file ) NP '*.xlsx' OR to_lower( p_out ) NP '*.xlsx'.
    MESSAGE 'Input and output files must be .xlsx files' TYPE 'E'.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Normalisers
*&---------------------------------------------------------------------*
FORM norm_header USING iv_in TYPE any CHANGING cv_out TYPE string.
  cv_out = to_upper( |{ iv_in }| ).
  REPLACE ALL OCCURRENCES OF REGEX '[ _]' IN cv_out WITH ''.
ENDFORM.

FORM norm_enh_type USING iv_in TYPE any CHANGING cv_out TYPE string.
  cv_out = to_upper( condense( |{ iv_in }| ) ).
  REPLACE ALL OCCURRENCES OF ` ` IN cv_out WITH '_'.   " 'APPEND STRUCTURE'
ENDFORM.

*&---------------------------------------------------------------------*
*& Read Excel (.xlsx) without OLE / Excel installed
*&---------------------------------------------------------------------*
FORM read_excel.
  DATA: lv_filename TYPE string,
        lt_bin      TYPE solix_tab,
        lv_len      TYPE i,
        lv_xstr     TYPE xstring,
        lo_excel    TYPE REF TO cl_fdt_xl_spreadsheet,
        lt_sheets   TYPE if_fdt_doc_spreadsheet=>t_worksheet_names,
        lv_sheet    TYPE string,
        lr_data     TYPE REF TO data,
        lv_header   TYPE string,
        lv_row      TYPE i,
        lv_col      TYPE i,
        ls_data     TYPE ty_data.

  FIELD-SYMBOLS: <lt_data>   TYPE STANDARD TABLE,
                 <ls_row>    TYPE any,
                 <lv_cell>   TYPE any,
                 <lv_target> TYPE any.

  lv_filename = p_file.

  " 1. Upload binary file from presentation server
  cl_gui_frontend_services=>gui_upload(
    EXPORTING filename = lv_filename
              filetype = 'BIN'
    IMPORTING filelength = lv_len
    CHANGING  data_tab = lt_bin
    EXCEPTIONS OTHERS = 1 ).
  IF sy-subrc <> 0.
    MESSAGE 'Unable to upload the input file' TYPE 'E'.
  ENDIF.

  CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
    EXPORTING input_length = lv_len
    IMPORTING buffer       = lv_xstr
    TABLES    binary_tab   = lt_bin.

  " 2. Parse workbook (first worksheet)
  TRY.
      lo_excel = NEW cl_fdt_xl_spreadsheet( document_name = lv_filename
                                            xdocument     = lv_xstr ).
      lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
        IMPORTING worksheet_names = lt_sheets ).
      READ TABLE lt_sheets INTO lv_sheet INDEX 1.
      lr_data = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet( lv_sheet ).
    CATCH cx_fdt_excel_core INTO DATA(lx_excel).
      MESSAGE lx_excel->get_text( ) TYPE 'E'.
  ENDTRY.

  ASSIGN lr_data->* TO <lt_data>.
  IF <lt_data> IS NOT ASSIGNED OR lines( <lt_data> ) < 2.
    MESSAGE 'No data found in the Excel file' TYPE 'E'.
  ENDIF.

  " 3. Row 1 = headers -> build column index map
  LOOP AT <lt_data> ASSIGNING <ls_row>.
    lv_row = sy-tabix.
    IF lv_row = 1.
      lv_col = 1.
      DO.
        ASSIGN COMPONENT lv_col OF STRUCTURE <ls_row> TO <lv_cell>.
        IF sy-subrc <> 0. EXIT. ENDIF.
        PERFORM norm_header USING <lv_cell> CHANGING lv_header.
        READ TABLE gt_map INTO DATA(ls_map) WITH KEY header = lv_header.
        IF sy-subrc = 0.
          APPEND VALUE #( idx = lv_col fieldname = ls_map-fieldname ) TO gt_colidx.
        ENDIF.
        lv_col = lv_col + 1.
      ENDDO.
      CONTINUE.
    ENDIF.

    " 4. Data rows
    CLEAR ls_data.
    LOOP AT gt_colidx INTO DATA(ls_ci).
      ASSIGN COMPONENT ls_ci-idx OF STRUCTURE <ls_row> TO <lv_cell>.
      ASSIGN COMPONENT ls_ci-fieldname OF STRUCTURE ls_data TO <lv_target>.
      IF <lv_cell> IS ASSIGNED AND <lv_target> IS ASSIGNED.
        <lv_target> = <lv_cell>.
      ENDIF.
    ENDLOOP.
    IF ls_data IS NOT INITIAL.
      APPEND ls_data TO gt_data.
    ENDIF.
  ENDLOOP.

  IF gt_data IS INITIAL.
    MESSAGE 'No usable data rows found' TYPE 'E'.
  ENDIF.
ENDFORM.

*&---------------------------------------------------------------------*
*& Business rules
*&---------------------------------------------------------------------*
FORM process_data.
  FIELD-SYMBOLS <ls_data> TYPE ty_data.
  LOOP AT gt_data ASSIGNING <ls_data>.
    PERFORM apply_amf_scope  CHANGING <ls_data>.   " Requirement 1
    PERFORM set_append_name  CHANGING <ls_data>.   " Requirement 2
  ENDLOOP.
ENDFORM.

*& Requirement 1: only when release <= 2023
FORM apply_amf_scope CHANGING cs_data TYPE ty_data.
  IF p_rel <= gc_cutoff_rel.
    DATA(lv_dtype) = to_upper( condense( |{ cs_data-data_type }| ) ).
    IF lv_dtype IN gr_qty_curr.
      cs_data-amf_scope = gc_out.
    ENDIF.
  ENDIF.
ENDFORM.

*& Requirement 2: Proposed Append Name
FORM set_append_name CHANGING cs_data TYPE ty_data.
  DATA: lv_type TYPE string,
        lv_name TYPE string,
        lv_new  TYPE string.

  PERFORM norm_enh_type USING cs_data-enh_type CHANGING lv_type.
  lv_name = condense( |{ cs_data-enh_name }| ).
  IF lv_name IS INITIAL.
    RETURN.
  ENDIF.

  CASE lv_type.
    WHEN gc_append.
      lv_new = |{ lv_name }{ gc_suffix }|.                 " ZST_MARA -> ZST_MARA_AMF
    WHEN gc_include.
      lv_new = |{ gc_prefix }{ lv_name }{ gc_suffix }|.    " CI_EKKODB -> ZCI_EKKODB_AMF
    WHEN gc_custom_inc.
      " Rule not defined yet -> left blank. Change here if CIs should
      " follow the INCLUDE_STRUCTURE rule.
    WHEN OTHERS.
  ENDCASE.

  IF strlen( lv_new ) > 30.      " DDIC object names max 30 chars
    gv_warn = gv_warn + 1.
  ENDIF.
  cs_data-prop_append = lv_new.
ENDFORM.

*&---------------------------------------------------------------------*
*& Build output .xlsx via SALV XML export and save locally
*&---------------------------------------------------------------------*
FORM set_header USING io_cols TYPE REF TO cl_salv_columns_table
                      iv_field TYPE lvc_fname
                      iv_text  TYPE string.
  DATA lo_col TYPE REF TO cl_salv_column.
  TRY.
      lo_col = io_cols->get_column( iv_field ).
      lo_col->set_long_text(   CONV #( iv_text ) ).
      lo_col->set_medium_text( CONV #( substring( val = iv_text off = 0
                               len = nmin( val1 = 20 val2 = strlen( iv_text ) ) ) ) ).
      lo_col->set_short_text(  CONV #( substring( val = iv_text off = 0
                               len = nmin( val1 = 10 val2 = strlen( iv_text ) ) ) ) ).
    CATCH cx_salv_not_found.
  ENDTRY.
ENDFORM.

FORM download_excel.
  DATA: lo_alv   TYPE REF TO cl_salv_table,
        lo_cols  TYPE REF TO cl_salv_columns_table,
        lv_xlsx  TYPE xstring,
        lt_solix TYPE solix_tab,
        lv_size  TYPE i,
        lv_file  TYPE string.

  TRY.
      cl_salv_table=>factory( IMPORTING r_salv_table = lo_alv
                              CHANGING  t_table      = gt_data ).
      lo_cols = lo_alv->get_columns( ).

      " Column headings exactly as business expects (ADAPT)
      PERFORM set_header USING lo_cols 'OBJECT_NAME'  'Object Name'.
      PERFORM set_header USING lo_cols 'FIELD_NAME'   'Field Name'.
      PERFORM set_header USING lo_cols 'ENH_TYPE'     'Enhancement Type'.
      PERFORM set_header USING lo_cols 'ENH_NAME'     'Enhancement Name'.
      PERFORM set_header USING lo_cols 'DATA_ELEMENT' 'Data Element'.
      PERFORM set_header USING lo_cols 'DOMAIN'       'Domain'.
      PERFORM set_header USING lo_cols 'DATA_TYPE'    'Data Type'.
      PERFORM set_header USING lo_cols 'LENGTH'       'Length'.
      PERFORM set_header USING lo_cols 'DECIMALS'     'Decimals'.
      PERFORM set_header USING lo_cols 'BUS_CONTEXT'  'Business Context'.
      PERFORM set_header USING lo_cols 'AMF_SCOPE'    'AMF Scope'.
      PERFORM set_header USING lo_cols 'PROP_APPEND'  'Proposed Append Name'.

      lv_xlsx = lo_alv->to_xml( if_salv_bs_xml=>c_type_xlsx ).
    CATCH cx_salv_msg cx_salv_not_found INTO DATA(lx_salv).
      MESSAGE lx_salv->get_text( ) TYPE 'E'.
  ENDTRY.

  lv_size  = xstrlen( lv_xlsx ).
  lt_solix = cl_bcs_convert=>xstring_to_solix( iv_xstring = lv_xlsx ).
  lv_file  = p_out.

  cl_gui_frontend_services=>gui_download(
    EXPORTING bin_filesize = lv_size
              filename     = lv_file
              filetype     = 'BIN'
    CHANGING  data_tab     = lt_solix
    EXCEPTIONS OTHERS      = 1 ).
  IF sy-subrc <> 0.
    MESSAGE 'Unable to save the output file' TYPE 'E'.
  ENDIF.

  MESSAGE |Output saved: { lines( gt_data ) } rows processed.| TYPE 'S'.
  IF gv_warn > 0.
    MESSAGE |{ gv_warn } proposed name(s) exceed 30 characters - please review.| TYPE 'S'
            DISPLAY LIKE 'W'.
  ENDIF.
ENDFORM.
```

---

# PART B: Local classes (object-oriented) version

| Include | Contents |
|---|---|
| `_TOP` | Types, exception class, and all class definitions |
| `_SEL` | Selection screen (same as above) |
| `_PROC` | All class implementations |

| Class | Responsibility |
|---|---|
| `lcx_amf` | Exception class for error handling |
| `lcl_excel_reader` | Upload the .xlsx, parse it, map columns to a typed table |
| `lcl_amf_processor` | Business rules 1 and 2 (testable on its own) |
| `lcl_excel_writer` | Build the output .xlsx and save it locally |
| `lcl_app` | Controller: F4 helps, validation, orchestration |

### Main program: `ZAMF_FIELD_AUTOMATION`

```abap
REPORT zamf_field_automation.

INCLUDE zamf_field_automation_top.
INCLUDE zamf_field_automation_sel.
INCLUDE zamf_field_automation_proc.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_file.
  lcl_app=>f4_input_file( CHANGING cv_file = p_file ).

AT SELECTION-SCREEN ON VALUE-REQUEST FOR p_out.
  lcl_app=>f4_output_file( CHANGING cv_file = p_out ).

AT SELECTION-SCREEN.
  lcl_app=>validate_input( iv_input = p_file iv_output = p_out ).

START-OF-SELECTION.
  DATA(go_app) = NEW lcl_app( iv_release = p_rel
                              iv_input   = p_file
                              iv_output  = p_out ).
  go_app->run( ).
```

### Include `ZAMF_FIELD_AUTOMATION_TOP`

```abap
*&---------------------------------------------------------------------*
*& Types
*&---------------------------------------------------------------------*
TYPES: BEGIN OF ty_data,                 " ASSUMED layout - adjust to yours
         object_name  TYPE c LENGTH 30,
         field_name   TYPE c LENGTH 30,
         enh_type     TYPE c LENGTH 30,  " Enhancement Type
         enh_name     TYPE c LENGTH 30,  " Enhancement Name
         data_element TYPE c LENGTH 30,
         domain       TYPE c LENGTH 30,
         data_type    TYPE c LENGTH 10,  " Data Type
         length       TYPE c LENGTH 10,
         decimals     TYPE c LENGTH 10,
         bus_context  TYPE c LENGTH 255, " Business Context
         " ---- output-only columns ----
         amf_scope    TYPE c LENGTH 10,  " AMF Scope
         prop_append  TYPE c LENGTH 60,  " Proposed Append Name
       END OF ty_data,
       tt_data TYPE STANDARD TABLE OF ty_data WITH DEFAULT KEY,

       BEGIN OF ty_map,                  " Excel header -> structure field
         header    TYPE string,
         fieldname TYPE fieldname,
       END OF ty_map,
       tt_map TYPE STANDARD TABLE OF ty_map WITH DEFAULT KEY.

*&---------------------------------------------------------------------*
*& Exception class
*&---------------------------------------------------------------------*
CLASS lcx_amf DEFINITION FINAL INHERITING FROM cx_static_check.
  PUBLIC SECTION.
    METHODS constructor IMPORTING iv_text TYPE string.
    METHODS get_text REDEFINITION.
  PRIVATE SECTION.
    DATA mv_text TYPE string.
ENDCLASS.

*&---------------------------------------------------------------------*
*& Excel reader
*&---------------------------------------------------------------------*
CLASS lcl_excel_reader DEFINITION FINAL.
  PUBLIC SECTION.
    METHODS constructor IMPORTING iv_file TYPE string.
    METHODS get_data RETURNING VALUE(rt_data) TYPE tt_data
                     RAISING   lcx_amf.

  PRIVATE SECTION.
    TYPES: BEGIN OF ty_colidx,
             idx       TYPE i,
             fieldname TYPE fieldname,
           END OF ty_colidx,
           tt_colidx TYPE STANDARD TABLE OF ty_colidx WITH DEFAULT KEY.

    DATA: mv_file TYPE string,
          mt_map  TYPE tt_map.

    METHODS init_map.
    METHODS upload_file RETURNING VALUE(rv_xstring) TYPE xstring
                        RAISING   lcx_amf.
    METHODS get_worksheet_data IMPORTING iv_xstring     TYPE xstring
                               RETURNING VALUE(rr_data) TYPE REF TO data
                               RAISING   lcx_amf.
    METHODS build_column_index IMPORTING is_header_row    TYPE any
                               RETURNING VALUE(rt_colidx) TYPE tt_colidx.
    METHODS check_required_columns IMPORTING it_colidx TYPE tt_colidx
                                   RAISING   lcx_amf.
    METHODS map_row IMPORTING is_row         TYPE any
                              it_colidx      TYPE tt_colidx
                    RETURNING VALUE(rs_data) TYPE ty_data.
    METHODS normalize_header IMPORTING iv_text          TYPE any
                             RETURNING VALUE(rv_header) TYPE string.
ENDCLASS.

*&---------------------------------------------------------------------*
*& Business rules
*&---------------------------------------------------------------------*
CLASS lcl_amf_processor DEFINITION FINAL.
  PUBLIC SECTION.
    CONSTANTS: c_cutoff_release TYPE zamf_s4_rel VALUE '2023',
               c_out            TYPE c LENGTH 3  VALUE 'OUT',
               c_append         TYPE string      VALUE 'APPEND_STRUCTURE',
               c_include        TYPE string      VALUE 'INCLUDE_STRUCTURE',
               c_prefix         TYPE string      VALUE 'Z',
               c_suffix         TYPE string      VALUE '_AMF',
               c_max_name_len   TYPE i           VALUE 30.

    METHODS constructor IMPORTING iv_release TYPE zamf_s4_rel.
    METHODS process CHANGING ct_data TYPE tt_data.
    METHODS get_long_name_count RETURNING VALUE(rv_count) TYPE i.

  PRIVATE SECTION.
    DATA: mv_release    TYPE zamf_s4_rel,
          mv_long_names TYPE i,
          mt_qty_curr   TYPE STANDARD TABLE OF char10 WITH DEFAULT KEY.

    METHODS is_scope_rule_active RETURNING VALUE(rv_active) TYPE abap_bool.
    METHODS is_qty_or_curr IMPORTING iv_data_type     TYPE any
                           RETURNING VALUE(rv_result) TYPE abap_bool.
    METHODS set_amf_scope CHANGING cs_data TYPE ty_data.
    METHODS build_append_name IMPORTING iv_enh_type    TYPE any
                                        iv_enh_name    TYPE any
                              RETURNING VALUE(rv_name) TYPE string.
    METHODS normalize_enh_type IMPORTING iv_text        TYPE any
                               RETURNING VALUE(rv_type) TYPE string.
ENDCLASS.

*&---------------------------------------------------------------------*
*& Excel writer
*&---------------------------------------------------------------------*
CLASS lcl_excel_writer DEFINITION FINAL.
  PUBLIC SECTION.
    METHODS constructor IMPORTING iv_file TYPE string.
    METHODS save_to_excel IMPORTING it_data TYPE tt_data
                          RAISING   lcx_amf.

  PRIVATE SECTION.
    DATA mv_file TYPE string.

    METHODS build_xlsx IMPORTING it_data         TYPE tt_data
                       RETURNING VALUE(rv_xlsx)  TYPE xstring
                       RAISING   lcx_amf.
    METHODS set_header IMPORTING io_columns TYPE REF TO cl_salv_columns_table
                                 iv_field   TYPE lvc_fname
                                 iv_text    TYPE string.
    METHODS download IMPORTING iv_xlsx TYPE xstring
                     RAISING   lcx_amf.
ENDCLASS.

*&---------------------------------------------------------------------*
*& Controller
*&---------------------------------------------------------------------*
CLASS lcl_app DEFINITION FINAL.
  PUBLIC SECTION.
    CLASS-METHODS f4_input_file  CHANGING cv_file TYPE rlgrap-filename.
    CLASS-METHODS f4_output_file CHANGING cv_file TYPE rlgrap-filename.
    CLASS-METHODS validate_input IMPORTING iv_input  TYPE rlgrap-filename
                                           iv_output TYPE rlgrap-filename.

    METHODS constructor IMPORTING iv_release TYPE zamf_s4_rel
                                  iv_input   TYPE rlgrap-filename
                                  iv_output  TYPE rlgrap-filename.
    METHODS run.

  PRIVATE SECTION.
    DATA: mv_release TYPE zamf_s4_rel,
          mv_input   TYPE string,
          mv_output  TYPE string.
ENDCLASS.
```

### Include `ZAMF_FIELD_AUTOMATION_PROC`

```abap
*&---------------------------------------------------------------------*
*& lcx_amf
*&---------------------------------------------------------------------*
CLASS lcx_amf IMPLEMENTATION.
  METHOD constructor.
    super->constructor( ).
    mv_text = iv_text.
  ENDMETHOD.

  METHOD get_text.
    result = mv_text.
  ENDMETHOD.
ENDCLASS.

*&---------------------------------------------------------------------*
*& lcl_excel_reader
*&---------------------------------------------------------------------*
CLASS lcl_excel_reader IMPLEMENTATION.

  METHOD constructor.
    mv_file = iv_file.
    init_map( ).
  ENDMETHOD.

  METHOD init_map.
    " Normalised header (upper case, no blanks/underscores) -> field.
    " ADAPT to your real column headings.
    mt_map = VALUE #(
      ( header = 'OBJECTNAME'      fieldname = 'OBJECT_NAME'  )
      ( header = 'FIELDNAME'       fieldname = 'FIELD_NAME'   )
      ( header = 'ENHANCEMENTTYPE' fieldname = 'ENH_TYPE'     )
      ( header = 'ENHANCEMENTNAME' fieldname = 'ENH_NAME'     )
      ( header = 'DATAELEMENT'     fieldname = 'DATA_ELEMENT' )
      ( header = 'DOMAIN'          fieldname = 'DOMAIN'       )
      ( header = 'DATATYPE'        fieldname = 'DATA_TYPE'    )
      ( header = 'LENGTH'          fieldname = 'LENGTH'       )
      ( header = 'DECIMALS'        fieldname = 'DECIMALS'     )
      ( header = 'BUSINESSCONTEXT' fieldname = 'BUS_CONTEXT'  ) ).
  ENDMETHOD.

  METHOD get_data.
    FIELD-SYMBOLS <lt_sheet> TYPE STANDARD TABLE.

    DATA(lr_data) = get_worksheet_data( upload_file( ) ).
    ASSIGN lr_data->* TO <lt_sheet>.
    IF <lt_sheet> IS NOT ASSIGNED OR lines( <lt_sheet> ) < 2.
      RAISE EXCEPTION NEW lcx_amf( iv_text = 'No data found in the Excel file' ).
    ENDIF.

    DATA lt_colidx TYPE tt_colidx.
    LOOP AT <lt_sheet> ASSIGNING FIELD-SYMBOL(<ls_row>).
      IF sy-tabix = 1.                        " header row
        lt_colidx = build_column_index( <ls_row> ).
        check_required_columns( lt_colidx ).
        CONTINUE.
      ENDIF.

      DATA(ls_data) = map_row( is_row = <ls_row> it_colidx = lt_colidx ).
      IF ls_data IS NOT INITIAL.
        APPEND ls_data TO rt_data.
      ENDIF.
    ENDLOOP.

    IF rt_data IS INITIAL.
      RAISE EXCEPTION NEW lcx_amf( iv_text = 'No usable data rows found' ).
    ENDIF.
  ENDMETHOD.

  METHOD upload_file.
    DATA: lt_bin TYPE solix_tab,
          lv_len TYPE i.

    cl_gui_frontend_services=>gui_upload(
      EXPORTING filename   = mv_file
                filetype   = 'BIN'
      IMPORTING filelength = lv_len
      CHANGING  data_tab   = lt_bin
      EXCEPTIONS OTHERS    = 1 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION NEW lcx_amf( iv_text = 'Unable to upload the input file' ).
    ENDIF.

    rv_xstring = cl_bcs_convert=>solix_to_xstring( it_solix = lt_bin
                                                   iv_size  = lv_len ).
  ENDMETHOD.

  METHOD get_worksheet_data.
    DATA lt_sheets TYPE if_fdt_doc_spreadsheet=>t_worksheet_names.

    TRY.
        DATA(lo_excel) = NEW cl_fdt_xl_spreadsheet( document_name = mv_file
                                                    xdocument     = iv_xstring ).
        lo_excel->if_fdt_doc_spreadsheet~get_worksheet_names(
          IMPORTING worksheet_names = lt_sheets ).
        IF lt_sheets IS INITIAL.
          RAISE EXCEPTION NEW lcx_amf( iv_text = 'The Excel file has no worksheet' ).
        ENDIF.
        rr_data = lo_excel->if_fdt_doc_spreadsheet~get_itab_from_worksheet(
                    lt_sheets[ 1 ] ).                " first worksheet
      CATCH cx_fdt_excel_core INTO DATA(lx_excel).
        RAISE EXCEPTION NEW lcx_amf( iv_text = lx_excel->get_text( ) ).
    ENDTRY.
  ENDMETHOD.

  METHOD build_column_index.
    DATA lv_col TYPE i VALUE 1.
    FIELD-SYMBOLS <lv_cell> TYPE any.

    DO.
      ASSIGN COMPONENT lv_col OF STRUCTURE is_header_row TO <lv_cell>.
      IF sy-subrc <> 0.
        EXIT.
      ENDIF.
      DATA(lv_header) = normalize_header( <lv_cell> ).
      ASSIGN mt_map[ header = lv_header ] TO FIELD-SYMBOL(<ls_map>).
      IF sy-subrc = 0.
        APPEND VALUE #( idx = lv_col fieldname = <ls_map>-fieldname ) TO rt_colidx.
      ENDIF.
      lv_col = lv_col + 1.
    ENDDO.
  ENDMETHOD.

  METHOD check_required_columns.
    LOOP AT VALUE string_table( ( `DATA_TYPE` ) ( `ENH_TYPE` ) ( `ENH_NAME` ) )
         INTO DATA(lv_required).
      IF NOT line_exists( it_colidx[ fieldname = lv_required ] ).
        RAISE EXCEPTION NEW lcx_amf(
          iv_text = |Required column missing in Excel header: { lv_required }| ).
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD map_row.
    FIELD-SYMBOLS: <lv_cell>   TYPE any,
                   <lv_target> TYPE any.

    LOOP AT it_colidx ASSIGNING FIELD-SYMBOL(<ls_ci>).
      ASSIGN COMPONENT <ls_ci>-idx OF STRUCTURE is_row TO <lv_cell>.
      IF sy-subrc <> 0.
        CONTINUE.
      ENDIF.
      ASSIGN COMPONENT <ls_ci>-fieldname OF STRUCTURE rs_data TO <lv_target>.
      IF sy-subrc = 0.
        <lv_target> = <lv_cell>.
      ENDIF.
    ENDLOOP.
  ENDMETHOD.

  METHOD normalize_header.
    rv_header = to_upper( |{ iv_text }| ).
    REPLACE ALL OCCURRENCES OF REGEX '[ _]' IN rv_header WITH ''.
  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& lcl_amf_processor
*&---------------------------------------------------------------------*
CLASS lcl_amf_processor IMPLEMENTATION.

  METHOD constructor.
    mv_release  = iv_release.
    mt_qty_curr = VALUE #( ( 'QUAN' ) ( 'CURR' ) ( 'CUKY' ) ( 'UNIT' ) ).
  ENDMETHOD.

  METHOD process.
    LOOP AT ct_data ASSIGNING FIELD-SYMBOL(<ls_data>).
      set_amf_scope( CHANGING cs_data = <ls_data> ).          " Requirement 1
      <ls_data>-prop_append =
        build_append_name( iv_enh_type = <ls_data>-enh_type   " Requirement 2
                           iv_enh_name = <ls_data>-enh_name ).
    ENDLOOP.
  ENDMETHOD.

  METHOD get_long_name_count.
    rv_count = mv_long_names.
  ENDMETHOD.

  METHOD is_scope_rule_active.
    rv_active = xsdbool( mv_release <= c_cutoff_release ).
  ENDMETHOD.

  METHOD is_qty_or_curr.
    DATA(lv_type) = to_upper( condense( |{ iv_data_type }| ) ).
    rv_result = xsdbool( line_exists( mt_qty_curr[ table_line = lv_type ] ) ).
  ENDMETHOD.

  METHOD set_amf_scope.
    IF is_scope_rule_active( ) = abap_true
       AND is_qty_or_curr( cs_data-data_type ) = abap_true.
      cs_data-amf_scope = c_out.
    ENDIF.
  ENDMETHOD.

  METHOD build_append_name.
    DATA(lv_type) = normalize_enh_type( iv_enh_type ).
    DATA(lv_name) = condense( |{ iv_enh_name }| ).
    IF lv_name IS INITIAL.
      RETURN.
    ENDIF.

    CASE lv_type.
      WHEN c_append.                                   " ZST_MARA -> ZST_MARA_AMF
        rv_name = |{ lv_name }{ c_suffix }|.
      WHEN c_include.                                  " CI_EKKODB -> ZCI_EKKODB_AMF
        rv_name = |{ c_prefix }{ lv_name }{ c_suffix }|.
      WHEN OTHERS.
        " CUSTOMIZING_INCLUDE: rule not defined yet -> left blank.
        " To treat it like INCLUDE_STRUCTURE use: WHEN c_include OR 'CUSTOMIZING_INCLUDE'.
    ENDCASE.

    IF strlen( rv_name ) > c_max_name_len.
      mv_long_names = mv_long_names + 1.
    ENDIF.
  ENDMETHOD.

  METHOD normalize_enh_type.
    rv_type = to_upper( condense( |{ iv_text }| ) ).
    REPLACE ALL OCCURRENCES OF ` ` IN rv_type WITH '_'.   " 'APPEND STRUCTURE'
  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& lcl_excel_writer
*&---------------------------------------------------------------------*
CLASS lcl_excel_writer IMPLEMENTATION.

  METHOD constructor.
    mv_file = iv_file.
  ENDMETHOD.

  METHOD save_to_excel.
    download( build_xlsx( it_data ) ).
  ENDMETHOD.

  METHOD build_xlsx.
    DATA(lt_data) = it_data.        " SALV needs a changeable table

    TRY.
        cl_salv_table=>factory( IMPORTING r_salv_table = DATA(lo_alv)
                                CHANGING  t_table      = lt_data ).
        DATA(lo_cols) = lo_alv->get_columns( ).

        " Column headings exactly as business expects (ADAPT)
        set_header( io_columns = lo_cols iv_field = 'OBJECT_NAME'  iv_text = `Object Name` ).
        set_header( io_columns = lo_cols iv_field = 'FIELD_NAME'   iv_text = `Field Name` ).
        set_header( io_columns = lo_cols iv_field = 'ENH_TYPE'     iv_text = `Enhancement Type` ).
        set_header( io_columns = lo_cols iv_field = 'ENH_NAME'     iv_text = `Enhancement Name` ).
        set_header( io_columns = lo_cols iv_field = 'DATA_ELEMENT' iv_text = `Data Element` ).
        set_header( io_columns = lo_cols iv_field = 'DOMAIN'       iv_text = `Domain` ).
        set_header( io_columns = lo_cols iv_field = 'DATA_TYPE'    iv_text = `Data Type` ).
        set_header( io_columns = lo_cols iv_field = 'LENGTH'       iv_text = `Length` ).
        set_header( io_columns = lo_cols iv_field = 'DECIMALS'     iv_text = `Decimals` ).
        set_header( io_columns = lo_cols iv_field = 'BUS_CONTEXT'  iv_text = `Business Context` ).
        set_header( io_columns = lo_cols iv_field = 'AMF_SCOPE'    iv_text = `AMF Scope` ).
        set_header( io_columns = lo_cols iv_field = 'PROP_APPEND'  iv_text = `Proposed Append Name` ).

        rv_xlsx = lo_alv->to_xml( if_salv_bs_xml=>c_type_xlsx ).
      CATCH cx_salv_msg INTO DATA(lx_salv).
        RAISE EXCEPTION NEW lcx_amf( iv_text = lx_salv->get_text( ) ).
    ENDTRY.
  ENDMETHOD.

  METHOD set_header.
    TRY.
        DATA(lo_col) = io_columns->get_column( iv_field ).
        lo_col->set_long_text(   CONV #( iv_text ) ).
        lo_col->set_medium_text( CONV #( iv_text ) ).   " truncated to 20 chars
        lo_col->set_short_text(  CONV #( iv_text ) ).   " truncated to 10 chars
      CATCH cx_salv_not_found.
    ENDTRY.
  ENDMETHOD.

  METHOD download.
    DATA(lt_solix) = cl_bcs_convert=>xstring_to_solix( iv_xstring = iv_xlsx ).

    cl_gui_frontend_services=>gui_download(
      EXPORTING bin_filesize = xstrlen( iv_xlsx )
                filename     = mv_file
                filetype     = 'BIN'
      CHANGING  data_tab     = lt_solix
      EXCEPTIONS OTHERS      = 1 ).
    IF sy-subrc <> 0.
      RAISE EXCEPTION NEW lcx_amf( iv_text = 'Unable to save the output file' ).
    ENDIF.
  ENDMETHOD.

ENDCLASS.

*&---------------------------------------------------------------------*
*& lcl_app
*&---------------------------------------------------------------------*
CLASS lcl_app IMPLEMENTATION.

  METHOD f4_input_file.
    DATA: lt_files TYPE filetable,
          lv_rc    TYPE i,
          lv_act   TYPE i.

    cl_gui_frontend_services=>file_open_dialog(
      EXPORTING window_title      = 'Select input Excel file'
                default_extension = 'xlsx'
                file_filter       = 'Excel (*.xlsx)|*.xlsx'
      CHANGING  file_table        = lt_files
                rc                = lv_rc
                user_action       = lv_act
      EXCEPTIONS OTHERS           = 1 ).
    IF sy-subrc = 0
       AND lv_act = cl_gui_frontend_services=>action_ok
       AND lines( lt_files ) > 0.
      cv_file = lt_files[ 1 ]-filename.
    ENDIF.
  ENDMETHOD.

  METHOD f4_output_file.
    DATA: lv_name TYPE string,
          lv_path TYPE string,
          lv_full TYPE string,
          lv_act  TYPE i.

    cl_gui_frontend_services=>file_save_dialog(
      EXPORTING window_title      = 'Save output Excel file'
                default_extension = 'xlsx'
                default_file_name = 'AMF_Output.xlsx'
                file_filter       = 'Excel (*.xlsx)|*.xlsx'
      CHANGING  filename          = lv_name
                path              = lv_path
                fullpath          = lv_full
                user_action       = lv_act
      EXCEPTIONS OTHERS           = 1 ).
    IF sy-subrc = 0 AND lv_act = cl_gui_frontend_services=>action_ok.
      cv_file = lv_full.
    ENDIF.
  ENDMETHOD.

  METHOD validate_input.
    IF NOT matches( val = to_lower( iv_input )  regex = '.*\.xlsx' )
    OR NOT matches( val = to_lower( iv_output ) regex = '.*\.xlsx' ).
      MESSAGE 'Input and output files must be .xlsx files' TYPE 'E'.
    ENDIF.
  ENDMETHOD.

  METHOD constructor.
    mv_release = iv_release.
    mv_input   = iv_input.
    mv_output  = iv_output.
  ENDMETHOD.

  METHOD run.
    TRY.
        DATA(lo_reader)    = NEW lcl_excel_reader( mv_input ).
        DATA(lo_processor) = NEW lcl_amf_processor( mv_release ).
        DATA(lo_writer)    = NEW lcl_excel_writer( mv_output ).

        DATA(lt_data) = lo_reader->get_data( ).
        lo_processor->process( CHANGING ct_data = lt_data ).
        lo_writer->save_to_excel( lt_data ).

        MESSAGE |Output saved: { lines( lt_data ) } rows processed.| TYPE 'S'.

        DATA(lv_long) = lo_processor->get_long_name_count( ).
        IF lv_long > 0.
          MESSAGE |{ lv_long } proposed name(s) exceed 30 characters - please review.|
                  TYPE 'S' DISPLAY LIKE 'W'.
        ENDIF.

      CATCH lcx_amf INTO DATA(lx_amf).
        MESSAGE lx_amf->get_text( ) TYPE 'S' DISPLAY LIKE 'E'.
    ENDTRY.
  ENDMETHOD.

ENDCLASS.
```

---

## Notes and open points

- **Output columns:** replace the assumed layout (`ty_data`, header map, output headings) with the real column lists.
- **CUSTOMIZING_INCLUDE:** the proposed name stays blank because no rule was given. To treat it like INCLUDE_STRUCTURE, use `WHEN gc_include OR gc_custom_inc.` (PERFORM version) or `WHEN c_include OR 'CUSTOMIZING_INCLUDE'.` (class version).
- **Name length:** DDIC names are limited to 30 characters. Names that exceed this after adding `Z` and `_AMF` are counted and a warning is shown.
- **Data Type matching:** the match is exact (e.g. `QUAN`). If the file holds values like `QUAN(13,3)`, switch to a prefix match.
- **SALV headers:** the SALV export may use the short or medium heading in some releases. Check the output in your system.
- **Limitations:** only .xlsx is supported (not .xls). The report must run in the foreground because it uses frontend upload and download.
- **Release requirement:** the code uses inline declarations, `VALUE`, `NEW` and `xsdbool`, so it needs NetWeaver 7.40 SP08 or higher.
- **Testing the class version:** `lcl_amf_processor` has no GUI or file dependencies, so you can add an ABAP Unit test class (`FOR TESTING`) with a `tt_data` table and assert results for releases 2022, 2023 and 2024.
