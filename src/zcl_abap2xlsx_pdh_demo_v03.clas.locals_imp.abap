*"* use this source file for the definition and implementation of
*"* local helper classes, interface definitions and type
*"* declarations
*--------------------------------------------------------------------*
* Defintions
*--------------------------------------------------------------------*
CLASS lcl_pers_layer DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    METHODS:
      constructor IMPORTING is_selections TYPE zcl_abap2xlsx_pdh_demo_v03=>m_typ_selections,
      derive_data RETURNING VALUE(rt_output_data) TYPE zcl_abap2xlsx_pdh_demo_v03=>m_tt_alv_output.

  PRIVATE SECTION.
    DATA: ms_selections TYPE zcl_abap2xlsx_pdh_demo_v03=>m_typ_selections.

ENDCLASS.

CLASS lcl_alv_view DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    DATA: mo_alv TYPE REF TO cl_salv_table.

    METHODS:
      initialise CHANGING ct_output_data TYPE zcl_abap2xlsx_pdh_demo_v03=>m_tt_alv_output,
      application_specific_changes,
      display.

ENDCLASS.

CLASS lcl_xlsx_view DEFINITION ##CLASS_FINAL.

  PUBLIC SECTION.
    METHODS:
      constructor        IMPORTING is_selections TYPE zcl_abap2xlsx_pdh_demo_v03=>m_typ_selections,
      set_alv            IMPORTING io_alv        TYPE REF TO cl_salv_table,
      create_spreadsheet IMPORTING it_table      TYPE STANDARD TABLE,
      application_specific_changes,
      email_spreadsheet.

  PRIVATE SECTION.
    DATA: mo_alv        TYPE REF TO cl_salv_table,
          mo_excel      TYPE REF TO zcl_excel,
          ms_selections TYPE zcl_abap2xlsx_pdh_demo_v03=>m_typ_selections.

ENDCLASS.

CLASS lcl_model DEFINITION ##CLASS_FINAL.
  PUBLIC SECTION.
    DATA: mt_output_data TYPE zcl_abap2xlsx_pdh_demo_v03=>m_tt_alv_output,
          ms_selections  TYPE zcl_abap2xlsx_pdh_demo_v03=>m_typ_selections.

    METHODS:
      constructor IMPORTING is_selections TYPE zcl_abap2xlsx_pdh_demo_v03=>m_typ_selections,
      derive_data.

  PRIVATE SECTION.
    DATA: mo_pers TYPE REF TO lcl_pers_layer.

ENDCLASS.

CLASS lcl_controller DEFINITION CREATE PRIVATE ##CLASS_FINAL.
  PUBLIC SECTION.
    CLASS-METHODS: get_instance IMPORTING io_model             TYPE REF TO lcl_model
                                          io_alv_view          TYPE REF TO lcl_alv_view
                                          io_xlsx_view         TYPE REF TO lcl_xlsx_view
                                RETURNING VALUE(ro_controller) TYPE REF TO lcl_controller.

    METHODS: constructor IMPORTING io_model     TYPE REF TO lcl_model
                                   io_alv_view  TYPE REF TO lcl_alv_view
                                   io_xlsx_view TYPE REF TO lcl_xlsx_view,
      main.

  PRIVATE SECTION.
    DATA: mo_model     TYPE REF TO lcl_model ##NEEDED,
          mo_alv_view  TYPE REF TO lcl_alv_view ##NEEDED,
          mo_xlsx_view TYPE REF TO lcl_xlsx_view ##NEEDED.
ENDCLASS.

CLASS ltd_pers_layer DEFINITION ##NEEDED ##CLASS_FINAL.
  PUBLIC SECTION.

ENDCLASS.

*--------------------------------------------------------------------*
* Implementations
*--------------------------------------------------------------------*
CLASS lcl_pers_layer IMPLEMENTATION.

  METHOD constructor.

    ms_selections = is_selections.

  ENDMETHOD.

  METHOD derive_data.

    SELECT *
      FROM sflight
      INTO CORRESPONDING FIELDS OF TABLE rt_output_data
      WHERE carrid IN ms_selections-r_carrid[]
      AND   connid IN ms_selections-r_connid[]
      AND   fldate IN ms_selections-r_fldate[]
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
        "Object = Column
        "Key    = Field Name e.g. VBELN
        "Object &OBJECT& &KEY& not found (class &CLASS& method &METHOD&)
        MESSAGE |{ not_found->get_text( ) }| TYPE 'E'.
    ENDTRY.

  ENDMETHOD.

  METHOD display.

    mo_alv->display( ).

  ENDMETHOD.

ENDCLASS.

CLASS lcl_xlsx_view IMPLEMENTATION.

  METHOD constructor.
    ms_selections = is_selections.
  ENDMETHOD.

  METHOD set_alv.
    mo_alv = io_alv.
  ENDMETHOD.

  METHOD create_spreadsheet.

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

        "Make sure vital values are always visible when user scrolls in spreadsheet
        worksheet->freeze_panes( ip_num_columns = 1
                                 ip_num_rows    = 1 ).

      "Page printing settings
      "Margins are to be set to the values for "narrow". I just copy
      "the values in the "narrow" option on the print preview
      worksheet->sheet_setup->set_page_margins( ip_top    = '1.91'
                                                ip_bottom = '1.91'
                                                ip_left   = '0.64'
                                                ip_right  = '0.64'
                                                ip_header = '0.76'
                                                ip_footer = '0.76'
                                                ip_unit   = 'cm' ).

      "No point wasting money on coloured printouts
      worksheet->sheet_setup->black_and_white = 'X'.

      "Requirement is landscape mode plus fit all columns on one sheet
      worksheet->sheet_setup->orientation         = zcl_excel_sheet_setup=>c_orientation_landscape.
      worksheet->sheet_setup->fit_to_page         = 'X'.
      worksheet->sheet_setup->fit_to_width        = 1.  " used only if ip_fit_to_page = 'X'
      worksheet->sheet_setup->page_order          = zcl_excel_sheet_setup=>c_ord_downthenover.
      worksheet->sheet_setup->paper_size          = zcl_excel_sheet_setup=>c_papersize_a4.
      worksheet->sheet_setup->scale               = 80. " used only if ip_fit_to_page = SPACE
      worksheet->sheet_setup->horizontal_centered = abap_true.

      DATA: ls_header TYPE zexcel_s_worksheet_head_foot,
            ls_footer TYPE zexcel_s_worksheet_head_foot.

      "Put Tab Name in Header Centre
      ls_header-center_value     = worksheet->get_title( ).
      ls_header-center_font-size = 8.
      ls_header-center_font-name = zcl_excel_style_font=>c_name_arial.

      "Put date on footer left
      ls_footer-left_value = '&D'.
      ls_footer-left_font  = ls_header-center_font.

      "Put page X of Y on Footer Right
      ls_footer-right_value = 'Page &P of &N' ##no_text.   "page x of y
      ls_footer-right_font  = ls_header-center_font.

      worksheet->sheet_setup->set_header_footer( ip_odd_header = ls_header
                                                 ip_odd_footer = ls_footer ).

      "When printing, repeat the header row
      worksheet->zif_excel_sheet_printsettings~set_print_repeat_rows(
            iv_rows_from = 1
            iv_rows_to   = 1 ).

      CATCH zcx_excel INTO DATA(exception).
        DATA(message) = exception->get_text( ).
        MESSAGE message TYPE 'I'.
    ENDTRY.

  ENDMETHOD.

  METHOD email_spreadsheet.
*--------------------------------------------------------------------------------------------------------------------*
* Some of the code here may look VERY familiar, because instead of re-inventing the wheel I copied it straight off
* the internet about 10+ years ago. It worked straightaway so I never bothered changing it apart from adding inline
* declarations
*--------------------------------------------------------------------------------------------------------------------*
    DATA: excel_writer TYPE REF TO zif_excel_writer,
          recipient    TYPE REF TO if_recipient_bcs,
          bytecount    TYPE i,
          maxbytecount TYPE i,
          filelen      TYPE so_obj_len.

    "Preconditions
    CHECK ms_selections-send_email    EQ abap_true.
    CHECK ms_selections-email_address IS NOT INITIAL.

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
        recipient = cl_cam_address_bcs=>create_internet_address( ms_selections-email_address ).

        "Add recipient object to send request
        send_request->add_recipient( recipient ).

        send_request->set_status_attributes( 'N' )."Never

        send_request->set_send_immediately( abap_true ).

        "Send Document
        DATA(sent_flag) = send_request->send( abap_true )."With Error Screen

        COMMIT WORK.

        IF sent_flag EQ abap_false.
          MESSAGE i500(sbcoms) WITH ms_selections-email_address."Document not sent to &1
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

  METHOD constructor.

    ms_selections = is_selections.

  ENDMETHOD.

  METHOD derive_data.

    mo_pers = NEW #( ms_selections ).

    mt_output_data = mo_pers->derive_data( ).

  ENDMETHOD.

ENDCLASS.

CLASS lcl_controller IMPLEMENTATION.

  METHOD get_instance.

    ro_controller = NEW lcl_controller( io_model     = io_model
                                        io_alv_view  = io_alv_view
                                        io_xlsx_view = io_xlsx_view ).

  ENDMETHOD.

  METHOD constructor.
    mo_model     = io_model.
    mo_alv_view  = io_alv_view.
    mo_xlsx_view = io_xlsx_view.
  ENDMETHOD.

  METHOD main.

    mo_model->derive_data( ).

    mo_alv_view->initialise( CHANGING ct_output_data = mo_model->mt_output_data[] ).
    mo_alv_view->application_specific_changes( ).

    "Have to send the email before showing the ALV on screen.
    "Sending the mail only when the user exits the screen
    "means you are snding it at a random time in some senses,
    "and that might confuse people
    IF mo_model->ms_selections-send_email EQ abap_true.
      mo_xlsx_view->set_alv( mo_alv_view->mo_alv ).
      mo_xlsx_view->create_spreadsheet( mo_model->mt_output_data[] ).
      mo_xlsx_view->application_specific_changes( ).
      mo_xlsx_view->email_spreadsheet( ).
    ENDIF.

    "This will bring up the screen and then wait till the user chooses a command
    mo_alv_view->display( ).

  ENDMETHOD.

ENDCLASS.

CLASS ltd_pers_layer IMPLEMENTATION.

ENDCLASS.
