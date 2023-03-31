*&---------------------------------------------------------------------*
*& Report Z_ABAP2XLSX_PDH_DEMO_V02
*&---------------------------------------------------------------------*
*& See https://blogs.sap.com/.......
*&---------------------------------------------------------------------*
REPORT z_abap2xlsx_pdh_demo_v02.
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

CLASS lcl_alv_view DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    DATA: mo_alv TYPE REF TO cl_salv_table.

    METHODS:
      initialise CHANGING ct_output_data TYPE g_tt_alv_output,
      application_specific_changes,
      display.

ENDCLASS.

CLASS lcl_xlsx_view DEFINITION ##CLASS_FINAL.

  PUBLIC SECTION.
    METHODS:
      constructor        IMPORTING io_alv   TYPE REF TO cl_salv_table,
      create_spreadsheet IMPORTING it_table TYPE STANDARD TABLE,
      application_specific_changes,
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
                                   io_view  TYPE REF TO lcl_alv_view.

  PRIVATE SECTION.
    DATA: mo_model    TYPE REF TO lcl_model ##NEEDED,
          mo_alv_view TYPE REF TO lcl_alv_view ##NEEDED.
ENDCLASS.

CLASS lcl_application DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    CLASS-METHODS: main.
  PRIVATE SECTION.
    CLASS-DATA: mo_alv_view   TYPE REF TO lcl_alv_view,
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

CLASS lcl_alv_view IMPLEMENTATION.

  METHOD initialise.

    TRY.
        cl_salv_table=>factory(
          IMPORTING r_salv_table = mo_alv
          CHANGING  t_table      = ct_output_data[]  ).
      CATCH cx_salv_msg INTO DATA(salv_exception).
        MESSAGE salv_exception->get_text( ) TYPE 'I'.
    ENDTRY.

  ENDMETHOD.

  METHOD application_specific_changes.

    DATA: lo_column TYPE REF TO cl_salv_column_table.

    DATA(lo_columns) = mo_alv->get_columns( ).

    TRY.
        lo_column ?= lo_columns->get_column( 'MANDT' ).
        lo_column->set_technical( abap_true ).

      CATCH cx_salv_not_found INTO DATA(not_found).
        "Raise a Fatal Exception
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
    DATA: l_ws TYPE c LENGTH 10 VALUE 'ITS'.

    "If we do this (call the convertor) in the foreground we get a dump unless we do a dirty trick
    IF sy-batch EQ abap_false.
      EXPORT l_ws = l_ws TO MEMORY ID 'WWW_ALV_ITS'.
    ENDIF.

    DATA(converter) = NEW zcl_excel_converter( ).

    TRY.
        IF mo_alv IS BOUND.
          converter->convert(
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

      CATCH zcx_excel INTO DATA(exception).
        DATA(message) = exception->get_text( ).
        MESSAGE message TYPE 'I'.
    ENDTRY.

  ENDMETHOD.

  METHOD application_specific_changes.

    DATA(worksheet)   = mo_excel->get_active_worksheet( ).
    DATA(sheet_title) = VALUE zexcel_sheet_title( ).

    TRY.
        "Every Excel spreadsheet has a title at the bottom left which defaults to "Sheet1"
        "Here I am hardcoding the value but you can set the value using whatever logic you want
        sheet_title = 'SFLIGHT'.
        worksheet->set_title( sheet_title ).

        "Make sure vital values are always visibnle when user scrolls in spreadsheet
        worksheet->freeze_panes( ip_num_columns = 1
                                 ip_num_rows    = 1 ).

      CATCH zcx_excel INTO DATA(exception).
        DATA(message) = exception->get_text( ).
        MESSAGE message TYPE 'I'.
    ENDTRY.

  ENDMETHOD.

  METHOD email_spreadsheet.
*--------------------------------------------------------------------------------------------------------------------*
* Some of the code here may look VERY familiar, because instead of re-inventing the wheel I copied it straight off
* the internet about 10+ years ago. It worked straightaway so I never bothered changing it aprt from adding inline
* declarations
*--------------------------------------------------------------------------------------------------------------------*
    DATA: excel_writer TYPE REF TO zif_excel_writer,
          recipient    TYPE REF TO if_recipient_bcs,
          bytecount    TYPE i,
          maxbytecount TYPE i,
          filelen      TYPE so_obj_len.

    "Preconditions
    CHECK lcl_selections=>p_send  EQ abap_true.
    CHECK lcl_selections=>p_email IS NOT INITIAL.

    TRY.
        excel_writer   = NEW zcl_excel_writer_2007( ).
        DATA(xml_file) = excel_writer->write_file( mo_excel ).

      CATCH zcx_excel INTO DATA(exception).
        DATA(message) = exception->get_text( ).
        MESSAGE message TYPE 'I'.
    ENDTRY.

    TRY.
        "Convert to binary
        DATA(file_tab) = cl_bcs_convert=>xstring_to_solix( iv_xstring  = xml_file ).
        bytecount      = xstrlen( xml_file ).
        maxbytecount   = 10000000 ##NUMBER_OK.

        "Create persistent send request
        DATA(send_request) = cl_bcs=>create_persistent( ).

        "Create document object from internal table with text
        DATA(main_text_table) = VALUE bcsy_text( ( line = 'This is the text in the body of the email'(009) ) ).
        IF bytecount >= maxbytecount.
          APPEND INITIAL LINE TO main_text_table ASSIGNING FIELD-SYMBOL(<main_text_line>).
          <main_text_line>-line =
          |{ 'The excel extract of the report cannot be sent because the resulting file is beyond the'(008) } | &&
          |{ maxbytecount / 1000000 DECIMALS = 0 }{ 'MB limit'(007) }|.
        ENDIF.

        DATA(document) = cl_document_bcs=>create_document( i_type    = 'RAW'
                                                           i_text    = main_text_table
                                                           i_subject = 'Email Subject'(006) ).

        IF bytecount < maxbytecount.

          "Add the spreadsheet as attachment to document object
          filelen = bytecount.

          DATA: att_head_table TYPE soli_tab,
                att_text_line  LIKE LINE OF att_head_table.

          CONCATENATE '&SO_FILENAME=' 'Attachment Subject.XLSX'(005) INTO att_text_line."Must end in XLSX or XLSM
          INSERT att_text_line INTO TABLE att_head_table.

          document->add_attachment( i_attachment_type    = 'EXT'
                                    i_attachment_subject = 'Attachment Subject.XLSX'(005)
                                    i_attachment_size    = filelen
                                    i_att_content_hex    = file_tab
                                    i_attachment_header  = att_head_table ) ##NO_TEXT.

        ENDIF.

        "Add document object to send request
        send_request->set_document( document ).

        "Create recipient object
        recipient = cl_cam_address_bcs=>create_internet_address( lcl_selections=>p_email ).

        "Add recipient object to send request
        send_request->add_recipient( recipient ).

        send_request->set_status_attributes( 'N' )."Never

        send_request->set_send_immediately( 'X' ).

        "Send Document
        DATA(sent_flag) = send_request->send( 'X' )."With Error Screen

        COMMIT WORK.

        IF sent_flag EQ abap_false.
          MESSAGE i500(sbcoms) WITH lcl_selections=>p_email."Document not sent to &1
        ELSE.
          MESSAGE s022(so)."Document sent
        ENDIF.

      CATCH cx_bcs INTO DATA(bcs_exception).
        "Error occurred during transmission - return code: <&>
        MESSAGE i865(so) WITH bcs_exception->error_type.
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
    mo_model     = io_model.
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
    mo_alv_view->application_specific_changes( ).

    IF lcl_selections=>p_send EQ abap_true.
      mo_xlsx_view = NEW #( mo_alv_view->mo_alv ).
      mo_xlsx_view->create_spreadsheet( mo_model->mt_output_data[] ).
      mo_xlsx_view->application_specific_changes( ).
      mo_xlsx_view->email_spreadsheet( ).
    ENDIF.

    mo_alv_view->display( ).

  ENDMETHOD.

ENDCLASS.

CLASS ltd_pers_layer IMPLEMENTATION.

ENDCLASS.

CLASS ltc_model IMPLEMENTATION.

ENDCLASS.
