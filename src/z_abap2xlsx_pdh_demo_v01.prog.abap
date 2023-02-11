*&---------------------------------------------------------------------*
*& Report Z_ABAP2XLSX_PDH_DEMO_V01
*&---------------------------------------------------------------------*
*& See (Insert BLOG URL)
*&---------------------------------------------------------------------*
REPORT z_abap2xlsx_pdh_demo_v01.
*--------------------------------------------------------------------*
* Data Definitions
*--------------------------------------------------------------------*
TYPES: BEGIN OF g_typ_selections,
         carrid TYPE s_carr_id,
         connid TYPE s_conn_id,
         fldate TYPE s_date,
       END OF   g_typ_selections.

TYPES: g_tt_alv_output TYPE TABLE OF sflight WITH DEFAULT KEY.

DATA: gs_selection_screen TYPE g_typ_selections ##needed."To avoid the need for TABLES

*--------------------------------------------------------------------*
* Class Definitions
*--------------------------------------------------------------------*
CLASS lcl_selections DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    CLASS-DATA: s_carrid TYPE RANGE OF s_carr_id,
                s_connid TYPE RANGE OF s_conn_id,
                s_fldate TYPE RANGE OF s_date,
                p_vari   TYPE disvariant-variant,
                p_send   TYPE char01,
                p_email  TYPE ad_smtpadr.

    CLASS-METHODS: set_data IMPORTING
                              is_carrid LIKE s_carrid
                              is_connid LIKE s_connid
                              is_fldate LIKE s_fldate
                              ip_vari   TYPE disvariant-variant
                              ip_send   TYPE abap_bool
                              ip_email  TYPE ad_smtpadr.

ENDCLASS.

CLASS lcl_pers_layer DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    METHODS: derive_data RETURNING VALUE(rt_output_data) TYPE g_tt_alv_output.
ENDCLASS.

CLASS lcl_view DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    DATA: mo_alv TYPE REF TO cl_salv_table.

    METHODS: initialise CHANGING ct_output_data TYPE g_tt_alv_output,
      display.

ENDCLASS.

CLASS lcl_xlsx_view DEFINITION ##CLASS_FINAL.

  PUBLIC SECTION.
    METHODS: constructor        IMPORTING io_alv   TYPE REF TO cl_salv_table,
      create_spreadsheet IMPORTING it_table TYPE STANDARD TABLE,
      email_spreadsheet.

  PRIVATE SECTION.
    DATA: mo_alv   TYPE REF TO cl_salv_table,
          mo_excel TYPE REF TO zcl_excel.

ENDCLASS.

CLASS lcl_model DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    DATA: mt_output_data TYPE g_tt_alv_output.

    METHODS: derive_data.

  PRIVATE SECTION.
    DATA: mo_pers TYPE REF TO lcl_pers_layer.

ENDCLASS.

CLASS lcl_controller DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    METHODS: constructor IMPORTING io_model TYPE REF TO lcl_model
                                   io_view  TYPE REF TO lcl_view.

  PRIVATE SECTION.
    DATA: mo_model    TYPE REF TO lcl_model ##NEEDED,
          mo_alv_view TYPE REF TO lcl_view ##NEEDED.
ENDCLASS.

CLASS lcl_application DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    CLASS-METHODS: main.
  PRIVATE SECTION.
    CLASS-DATA: mo_alv_view   TYPE REF TO lcl_view,
                mo_xlsx_view  TYPE REF TO lcl_xlsx_view,
                mo_controller TYPE REF TO lcl_controller,
                mo_model      TYPE REF TO lcl_model.
ENDCLASS.

CLASS ltd_pers_layer DEFINITION ##NEEDED ##CLASS_FINAL.
  PUBLIC SECTION.

ENDCLASS.

CLASS ltc_model DEFINITION ##NEEDED ##CLASS_FINAL.
  PUBLIC SECTION.
ENDCLASS.

*--------------------------------------------------------------------*
* Selection Screen
*--------------------------------------------------------------------*
* Organisational Based Selections
SELECTION-SCREEN BEGIN OF BLOCK blk1 WITH FRAME TITLE TEXT-001.

SELECT-OPTIONS: s_carrid FOR gs_selection_screen-carrid,
                s_connid FOR gs_selection_screen-connid.

SELECTION-SCREEN END OF BLOCK blk1.

* Time Based Selections
SELECTION-SCREEN BEGIN OF BLOCK blk2 WITH FRAME TITLE TEXT-002.

SELECT-OPTIONS: s_fldate FOR gs_selection_screen-fldate.

SELECTION-SCREEN END OF BLOCK blk2.

* Display Options
SELECTION-SCREEN BEGIN OF BLOCK blk3 WITH FRAME TITLE TEXT-003.
PARAMETERS: p_vari LIKE disvariant-variant.
SELECTION-SCREEN END OF BLOCK blk3.

* Background Execution Options
SELECTION-SCREEN BEGIN OF BLOCK blk4 WITH FRAME TITLE TEXT-004.
PARAMETERS: p_send  AS CHECKBOX.
PARAMETERS: p_email TYPE ad_smtpadr.
SELECTION-SCREEN END OF BLOCK blk4.
*--------------------------------------------------------------------*
* Initialisation
*--------------------------------------------------------------------*
INITIALIZATION.

*--------------------------------------------------------------------*
* Start-of-Selection
*--------------------------------------------------------------------*
START-OF-SELECTION.
  lcl_selections=>set_data(
    is_carrid = s_carrid[]
    is_connid = s_connid[]
    is_fldate = s_fldate[]
    ip_vari   = p_vari
    ip_send   = p_send
    ip_email  = p_email ).

  lcl_application=>main( ).

*--------------------------------------------------------------------*
* Class Implementations
*--------------------------------------------------------------------*
CLASS lcl_selections IMPLEMENTATION.

  METHOD set_data.

    s_carrid[] = is_carrid[].
    s_connid[] = is_connid[].
    s_fldate[] = is_fldate[].
    p_vari     = ip_vari.
    p_send     = ip_send.
    p_email    = ip_email.

  ENDMETHOD.

ENDCLASS.

CLASS lcl_pers_layer IMPLEMENTATION.

  METHOD derive_data.

    SELECT *
      FROM sflight
      INTO CORRESPONDING FIELDS OF TABLE rt_output_data
      WHERE carrid IN lcl_selections=>s_carrid[]
      AND   connid IN lcl_selections=>s_connid[]
      AND   fldate IN lcl_selections=>s_fldate[]
      ORDER BY PRIMARY KEY.

    IF sy-subrc NE 0.
      RETURN.
    ENDIF.

  ENDMETHOD.

ENDCLASS.

CLASS lcl_view IMPLEMENTATION.

  METHOD initialise.

    TRY.
        cl_salv_table=>factory(
          IMPORTING r_salv_table = mo_alv
          CHANGING  t_table      = ct_output_data[]  ).
      CATCH cx_salv_msg INTO DATA(lx_salv_msg).
        MESSAGE lx_salv_msg->get_text( ) TYPE 'I'.
    ENDTRY.

  ENDMETHOD.

  METHOD display.

    mo_alv->display( ).

  ENDMETHOD.

ENDCLASS.

CLASS lcl_xlsx_view IMPLEMENTATION.

  METHOD constructor.
    mo_alv = io_alv.
  ENDMETHOD.

  METHOD create_spreadsheet.
    "Convert SALV object into excel
    DATA: ld_sheet_title TYPE zexcel_sheet_title,
          l_ws           TYPE c LENGTH 10 VALUE 'ITS'.

    "If we do this (call the convertor) in the foreground we get a dump
    "unless we do a dirty trick
    IF sy-batch EQ abap_false.
      EXPORT l_ws = l_ws TO MEMORY ID 'WWW_ALV_ITS'.
    ENDIF.

    DATA(lo_converter) = NEW zcl_excel_converter( ).

    TRY.
        IF mo_alv IS BOUND.
          lo_converter->convert(
            EXPORTING
              io_alv        = mo_alv
              it_table      = it_table[]
              i_table       = abap_true "Create as Table
              i_style_table = zcl_excel_table=>builtinstyle_medium2
            CHANGING
              co_excel      = mo_excel ).
        ELSE.
          RETURN.
        ENDIF.

        FREE MEMORY ID 'WWW_ALV_ITS'.

        DATA(lo_worksheet) = mo_excel->get_active_worksheet( ).

        ld_sheet_title = 'SFLIGHT'.
        lo_worksheet->set_title( ld_sheet_title ).

      CATCH zcx_excel INTO DATA(lo_exception).
        DATA(ld_message) = lo_exception->get_text( ).
        MESSAGE ld_message TYPE 'I'.
    ENDTRY.

  ENDMETHOD.

  METHOD email_spreadsheet.
*--------------------------------------------------------------------------------------------------------------------*
* Some of the code here may look VERY familiar, because instead of re-inventing the wheel I copied it straight off
* the internet about 10+ years ago. It worked straightaway so I never bothered changing it aprt from adding inline
* declarations
*--------------------------------------------------------------------------------------------------------------------*
    DATA: lo_excel_writer TYPE REF TO zif_excel_writer,
          lo_recipient    TYPE REF TO if_recipient_bcs,
          ld_bytecount    TYPE i,
          ld_maxbytecount TYPE i,
          ld_filelen      TYPE so_obj_len.

    "Preconditions
    CHECK lcl_selections=>p_send  EQ abap_true.
    CHECK lcl_selections=>p_email IS NOT INITIAL.

    TRY.
        lo_excel_writer   = NEW zcl_excel_writer_2007( ).
        DATA(ld_xml_file) = lo_excel_writer->write_file( mo_excel ).

      CATCH zcx_excel INTO DATA(lo_exception).
        DATA(ld_message) = lo_exception->get_text( ).
        MESSAGE ld_message TYPE 'I'.
    ENDTRY.

    TRY.
        "Convert to binary
        DATA(lt_file_tab) = cl_bcs_convert=>xstring_to_solix( iv_xstring  = ld_xml_file ).
        ld_bytecount      = xstrlen( ld_xml_file ).
        ld_maxbytecount   = 10000000 ##NUMBER_OK.

        "Create persistent send request
        DATA(lo_send_request) = cl_bcs=>create_persistent( ).

        "Create document object from internal table with text
        DATA(lt_main_text) = VALUE bcsy_text( ( line = 'This is the text in the body of the email'(009) ) ).
        IF ld_bytecount >= ld_maxbytecount.
          APPEND INITIAL LINE TO lt_main_text ASSIGNING FIELD-SYMBOL(<ls_main_text>).
          <ls_main_text>-line =
          |{ 'The excel extract of the report cannot be sent because the resulting file is beyond the'(008) } | &&
          |{ ld_maxbytecount / 1000000 DECIMALS = 0 }{ 'MB limit'(007) }|.
        ENDIF.

        DATA(lo_document) = cl_document_bcs=>create_document( i_type    = 'RAW'
                                                              i_text    = lt_main_text
                                                              i_subject = 'Email Subject'(006) ).

        IF ld_bytecount < ld_maxbytecount.

          "Add the spreadsheet as attachment to document object
          ld_filelen = ld_bytecount.

          DATA: lt_att_head  TYPE soli_tab,
                lv_text_line TYPE soli.
          CONCATENATE '&SO_FILENAME=' 'Attachment Subject.XLSX'(005) INTO lv_text_line."Must end in XLSX or XLSM
          APPEND lv_text_line TO lt_att_head.

          lo_document->add_attachment( i_attachment_type    = 'EXT'
                                       i_attachment_subject = 'Attachment Subject.XLSX'(005)
                                       i_attachment_size    = ld_filelen
                                       i_att_content_hex    = lt_file_tab
                                       i_attachment_header  = lt_att_head ) ##NO_TEXT.

        ENDIF.

        "Add document object to send request
        lo_send_request->set_document( lo_document ).

        "Create recipient object
        lo_recipient = cl_cam_address_bcs=>create_internet_address( lcl_selections=>p_email ).

        "Add recipient object to send request
        lo_send_request->add_recipient( lo_recipient ).

        lo_send_request->set_status_attributes( 'N' )."Never

        "Send Document
        DATA(lf_sent_to_all) = lo_send_request->send( 'X' )."With Error Screen

        COMMIT WORK.

        IF lf_sent_to_all IS INITIAL.
          MESSAGE i500(sbcoms) WITH lcl_selections=>p_email."Document not sent to &1
        ELSE.
          MESSAGE s022(so)."Document sent
          "Kick off the send job so the email goes out immediately
          WAIT UP TO 2 SECONDS.     "ensure the mail has been queued
          SUBMIT rsconn01
            WITH mode   = '*'       "process everything you find.
            WITH output = ' '
          AND RETURN.                                    "#EC CI_SUBMIT
        ENDIF.

      CATCH cx_bcs INTO DATA(lo_bcs_exception).
        "Error occurred during transmission - return code: <&>
        MESSAGE i865(so) WITH lo_bcs_exception->error_type.
    ENDTRY.

  ENDMETHOD.

ENDCLASS.

CLASS lcl_model IMPLEMENTATION.

  METHOD derive_data.

    mo_pers = NEW #( ).

    mt_output_data = mo_pers->derive_data( ).

  ENDMETHOD.

ENDCLASS.

CLASS lcl_controller IMPLEMENTATION.

  METHOD constructor.
    mo_model = io_model.
    mo_alv_view  = io_view.
  ENDMETHOD.

ENDCLASS.

CLASS lcl_application IMPLEMENTATION.

  METHOD main.

    mo_model      = NEW #( ).
    mo_alv_view   = NEW #( ).
    mo_controller = NEW #( io_model = mo_model
                           io_view  = mo_alv_view ) ##NEEDED."For User Commands

    mo_model->derive_data( ).
    mo_alv_view->initialise( CHANGING ct_output_data = mo_model->mt_output_data[] ).

    IF lcl_selections=>p_send EQ abap_true.
      mo_xlsx_view = NEW #( mo_alv_view->mo_alv ).
      mo_xlsx_view->create_spreadsheet( mo_model->mt_output_data[] ).
      mo_xlsx_view->email_spreadsheet( ).
    ENDIF.

    mo_alv_view->display( ).

  ENDMETHOD.

ENDCLASS.

CLASS ltd_pers_layer IMPLEMENTATION.

ENDCLASS.

CLASS ltc_model IMPLEMENTATION.

ENDCLASS.
