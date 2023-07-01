*&---------------------------------------------------------------------*
*& Report Z_ABAP2XLSX_PDH_DEMO_V03
*&---------------------------------------------------------------------*
*& See https://blogs.sap.com/.......
*&---------------------------------------------------------------------*
* The main difference between version 2 and version 3 is that now the
* executable program has one job only, which is to get the users
* choices from the selection screen and then pass them to a global class
* which does all the processing
*----------------------------------------------------------------------*
REPORT z_abap2xlsx_pdh_demo_v03.
*--------------------------------------------------------------------*
* Data Definitions
*--------------------------------------------------------------------*
TYPES: BEGIN OF g_typ_selections,
         carrid TYPE s_carr_id,
         connid TYPE s_conn_id,
         fldate TYPE s_date,
       END OF   g_typ_selections.

DATA: gs_selection_screen TYPE g_typ_selections ##needed."To avoid the need for a TABLES statement

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
PARAMETERS: p_vari TYPE disvariant-variant.
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
  PERFORM main.

FORM main.

  zcl_abap2xlsx_pdh_demo_v03=>get_instance(
                             VALUE #( variant       = p_vari
                                      send_email    = p_send
                                      email_address = p_email
                                      r_carrid      = s_carrid[]
                                      r_connid      = s_connid[]
                                      r_fldate      = s_fldate[] ) )->main( ).

ENDFORM.
