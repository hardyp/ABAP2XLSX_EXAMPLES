class ZCL_ABAP2XLSX_PDH_DEMO_V03 definition
  public
  create private .

public section.

  interfaces ZIF_ABAP2XLSX_PDH_DEMO_V03 .

  aliases MAIN
    for ZIF_ABAP2XLSX_PDH_DEMO_V03~MAIN .

  types:
    BEGIN OF m_typ_selections,
           variant       TYPE disvariant-variant,
           send_email    TYPE abap_bool,
           email_address TYPE ad_smtpadr,
           r_carrid      TYPE typ_r_carrid,
           r_connid      TYPE typ_r_connid,
           r_fldate      TYPE typ_r_fldate,
         END OF   m_typ_selections .
  types:
    M_TT_ALV_OUTPUT TYPE STANDARD TABLE OF sflight WITH EMPTY KEY .

  methods CONSTRUCTOR
    importing
      !IS_SELECTIONS type M_TYP_SELECTIONS .
  class-methods GET_INSTANCE
    importing
      !IS_SELECTIONS type M_TYP_SELECTIONS
    returning
      value(RO_INSTANCE) type ref to ZCL_ABAP2XLSX_PDH_DEMO_V03 .
protected section.
private section.

  data MS_SELECTIONS type M_TYP_SELECTIONS .
ENDCLASS.



CLASS ZCL_ABAP2XLSX_PDH_DEMO_V03 IMPLEMENTATION.


  METHOD constructor.

    ms_selections = is_selections.

  ENDMETHOD.


  METHOD get_instance.

    ro_instance = NEW zcl_abap2xlsx_pdh_demo_v03( is_selections ).

  ENDMETHOD.


  METHOD zif_abap2xlsx_pdh_demo_v03~main.

     lcl_controller=>get_instance(
       io_model     = NEW lcl_model( ms_selections )
       io_alv_view  = NEW lcl_alv_view( )
       io_xlsx_view = NEW lcl_xlsx_view( ms_selections ) )->main( ).

  ENDMETHOD.
ENDCLASS.
