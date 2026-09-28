*&---------------------------------------------------------------------*
*& Report  ZPP024
*&---------------------------------------------------------------------*
* Program ID/Name: ZPP024
* Author's name: ABAP_LLP
* Program title: 生产订单批量取消TECO\关闭(结算)
* Project Name : PP
* Version: 1.0
*----------------------------------------------------------------------*
* Change History
*----------------------------------------------------------------------*
*     Date   |   Programmer   |   Corr. #   |   Description
* 2020.02.17 |   ABAP_LLP     |   DE1K906065|   Initial
************************************************************************
REPORT zppc009 MESSAGE-ID zpp01.
*----------------------------------------------------------------------*
*             Table
*----------------------------------------------------------------------*
TABLES: caufv,jest,sscrfields.
*----------------------------------------------------------------------*
*             TYPE-POOLS
*----------------------------------------------------------------------*
TYPE-POOLS: slis,abap,icon.
*----------------------------------------------------------------------*
*             Global Constants
*----------------------------------------------------------------------*
CONSTANTS: con_x    TYPE c VALUE 'X',
           con_a    TYPE c VALUE 'A',
           con_w    TYPE c VALUE 'W',
           con_e    TYPE c VALUE 'E',
           con_s    TYPE c VALUE 'S',
           con_st45 TYPE jcds-stat VALUE 'I0045', "TECO
           con_st46 TYPE jcds-stat VALUE 'I0046'. "结算	已结算
*----------------------------------------------------------------------*
*             Global Variables
*----------------------------------------------------------------------*
DATA: g_tabix  TYPE sy-tabix,
      g_filter TYPE string,
      g_seqno  TYPE i,
      g_uflag  TYPE c.
*----------------------------------------------------------------------*
*             Types
*----------------------------------------------------------------------*
*----------------------------------------------------------------------*
*             Types
*----------------------------------------------------------------------*
TYPES: BEGIN OF ty_upld,
         aufnr TYPE caufv-aufnr,    "生产订单
       END OF ty_upld.

TYPES: BEGIN OF ty_itab,
         seqno  TYPE i,
         aufnr  TYPE caufv-aufnr,    "生产订单
         micon  TYPE char04,         "状态
         msgtyp TYPE char01,         "类型
         msgtxt TYPE string,         "消息
         objnr  TYPE caufv-objnr,
       END OF ty_itab.
*----------------------------------------------------------------------*
*             Internal Table & Work Area
*----------------------------------------------------------------------*
DATA: gt_upld TYPE TABLE OF ty_upld,
      gs_upld TYPE ty_upld,
      gt_itab TYPE TABLE OF ty_itab,
      gs_itab TYPE ty_itab.
*----------------------------------------------------------------------*
*             BDC
*----------------------------------------------------------------------*
DATA: gt_bdcd TYPE TABLE OF bdcdata,
      gs_bdcd TYPE bdcdata,
      gt_bmsg TYPE TABLE OF bdcmsgcoll,
      gs_bmsg TYPE bdcmsgcoll.
DATA: gs_ctu_params TYPE ctu_params,
      gv_msgtx      TYPE string.
*----------------------------------------------------------------------*
*             ALV Define
*----------------------------------------------------------------------*
DATA: gt_fcat TYPE slis_t_fieldcat_alv WITH HEADER LINE,
      gt_sort TYPE slis_t_sortinfo_alv WITH HEADER LINE.
DATA: gs_layout  TYPE slis_layout_alv,
      gs_variant TYPE disvariant.
DATA: g_repid TYPE sy-repid,
      g_title TYPE lvc_title.
DATA: g_pos TYPE i.
DEFINE macro_fieldcat. "用于显示
  g_pos = g_pos + 1.
  gt_fcat-col_pos       = g_pos.
  gt_fcat-fieldname     = &1.
  gt_fcat-seltext_m     = &2.
  gt_fcat-seltext_s     = &2.
  gt_fcat-seltext_l     = &2.
  gt_fcat-outputlen     = &3.
  gt_fcat-key           = &4.
  gt_fcat-ref_tabname   = &5.
  gt_fcat-ref_fieldname = &6.
  gt_fcat-hotspot       = &7.
  gt_fcat-do_sum        = &8.
  APPEND gt_fcat.
  CLEAR  gt_fcat.
END-OF-DEFINITION.
*----------------------------------------------------------------------*
*             SELECT-SCREEN
*----------------------------------------------------------------------*
"SELECTION-SCREEN FUNCTION KEY: 1 .
SELECTION-SCREEN BEGIN OF BLOCK b1 WITH FRAME TITLE TEXT-s01.
  PARAMETERS:p_path LIKE rlgrap-filename DEFAULT 'C:\*.xls' NO-DISPLAY.
  SELECT-OPTIONS: s_aufnr FOR caufv-aufnr NO INTERVALS.   "生产订单
SELECTION-SCREEN END OF BLOCK b1.
SELECTION-SCREEN BEGIN OF BLOCK b2 WITH FRAME TITLE TEXT-s02.
  PARAMETERS: g_r1 RADIOBUTTON GROUP rb USER-COMMAND uc DEFAULT 'X',
              g_r2 RADIOBUTTON GROUP rb.
SELECTION-SCREEN END OF BLOCK b2.
*----------------------------------------------------------------------*
*             INITIALZATION
*----------------------------------------------------------------------*
INITIALIZATION. "初始化事件
  PERFORM frm_init.
*----------------------------------------------------------------------*
*             AT SELECTION-SCREEN
*----------------------------------------------------------------------*
AT SELECTION-SCREEN.
  IF sscrfields-ucomm = 'FC01'.
    PERFORM frm_download_tmp.
  ENDIF.

*AT SELECTION-SCREEN ON VALUE-REQUEST FOR P_PATH.
*  PERFORM FRM_SELECT_FILE USING G_FILTER P_PATH.

*----------------------------------------------------------------------*
*             START-OF-SELECTION
*----------------------------------------------------------------------*
START-OF-SELECTION. "选择后时间
  "  PERFORM FRM_UPLOAD_DATA.
  DELETE s_aufnr WHERE low  = '*'.
  IF s_aufnr[] IS INITIAL.
    MESSAGE i000 WITH '请输入生产订单号'.
    STOP.
  ENDIF.
  PERFORM frm_check_data.
  IF gt_itab[] IS INITIAL.
    MESSAGE s002 DISPLAY LIKE 'E'.
  ELSE.
    PERFORM frm_out_data.
  ENDIF.
*&---------------------------------------------------------------------*
*& Form FRM_OUT_DATA
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM frm_out_data .
  g_repid = sy-repid.
*>1.显示字段
  PERFORM frm_fieldcat_init USING ''.
*>2.Layout
  " GS_LAYOUT-COLWIDTH_OPTIMIZE = CON_X.
  gs_layout-zebra = con_x.

*>3.VARIANT
  CLEAR:gs_variant.
  gs_variant-report = g_repid.
  gs_variant-handle = sy-cprog && sy-datum && sy-uzeit.
  IF g_r1 = con_x. "取消TECO
    g_title = '取消TECO'(p01).
  ELSE.
    g_title = '取消关闭(结算)'(p02).
  ENDIF.
*>4.ALV输出
  CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
    EXPORTING
      i_callback_program       = g_repid
      i_callback_pf_status_set = 'FRM_PF_STATUS_SET'
      i_callback_user_command  = 'FRM_USER_COMMAND'
      it_fieldcat              = gt_fcat[]
      is_layout                = gs_layout
      is_variant               = gs_variant
      i_grid_title             = g_title
      i_save                   = con_a
    TABLES
      t_outtab                 = gt_itab[]
    EXCEPTIONS
      program_error            = 1
      OTHERS                   = 2.
  IF sy-subrc <> 0.
    MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
          WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*&      Form  fm_fieldcat_init
*&---------------------------------------------------------------------*
*       设定要显示的字段
*----------------------------------------------------------------------*
FORM frm_fieldcat_init USING p_ucomm.
  REFRESH gt_fcat.
  CLEAR: g_pos.
  CASE p_ucomm.
    WHEN ''. "所有
      macro_fieldcat:
        'SEQNO'  '序号'           6 'X' '' '' '' '',
        'AUFNR'  '生产订单'      12 'X' 'AUFK' 'AUFNR' 'X' '',
        'MICON'  '状态'           4 ' ' '' '' '' '',
        'MSGTYP' '类型'           4 ' ' '' '' '' '',
        'MSGTXT' '消息'         100 ' ' '' '' '' ''.
    WHEN OTHERS.
      ...
  ENDCASE.
ENDFORM.                    "fm_fieldcat_init
*&---------------------------------------------------------------------*
*&      Form  FRM_PF_STATUS_SET
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*  -->  p1        text
*  <--  p2        text
*----------------------------------------------------------------------*
FORM frm_pf_status_set USING rt_extab TYPE slis_t_extab.
  DATA:ls_extab LIKE LINE OF rt_extab.
  REFRESH:rt_extab.
  IF g_uflag = con_x.
    APPEND 'UPDATE' TO rt_extab.
  ENDIF.
  SET PF-STATUS 'PF_STATUS' EXCLUDING rt_extab.
ENDFORM.
*&--------------------------------------------------------------------*
*&      Form FRM_USER_COMMAND
*&--------------------------------------------------------------------*
*       text
*---------------------------------------------------------------------*
*      -->R_UCOMM    text
*      -->SELFIELD   text
*---------------------------------------------------------------------*
FORM frm_user_command USING r_ucomm TYPE sy-ucomm
                        rs_selfield TYPE slis_selfield.
  DATA: lr_grid TYPE REF TO cl_gui_alv_grid,
        l_aufnr TYPE aufnr.
  CALL FUNCTION 'GET_GLOBALS_FROM_SLVC_FULLSCR'
    IMPORTING
      e_grid = lr_grid.
  CALL METHOD lr_grid->check_changed_data.
  CASE r_ucomm.
    WHEN '&IC1'.
      CLEAR: gs_itab.
      READ TABLE gt_itab INTO gs_itab INDEX rs_selfield-tabindex.
      CASE rs_selfield-sel_tab_field.
        WHEN '1-AUFNR'.
          CHECK gs_itab-aufnr IS NOT INITIAL.
          SET PARAMETER ID 'ANR' FIELD gs_itab-aufnr .
          CALL TRANSACTION 'CO03' AND SKIP FIRST SCREEN.
        WHEN OTHERS.
          ...
      ENDCASE.
    WHEN 'UPDATE'. "批量更改
      READ TABLE gt_itab INTO gs_itab WITH KEY msgtyp = con_e.
      IF sy-subrc = 0.
        MESSAGE e000 WITH '数据存在错误,不能批量更改!'.
        EXIT.
      ELSE.
        PERFORM frm_batch_update.
      ENDIF.
    WHEN OTHERS.
      ...
  ENDCASE.

**自动刷新
  rs_selfield-refresh    = con_x.
  rs_selfield-col_stable = con_x.
  rs_selfield-row_stable = con_x.
ENDFORM.                    " FRM_USER_COMMAND
*&---------------------------------------------------------------------*
*& Form FRM_INIT
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM frm_init.
  DATA: wa_functxt TYPE smp_dyntxt.
  g_filter = 'EXCEL-Files (*.XLSX)|*.XLSX|(*.XLS)|*.XLS|All Files(*.*)|*.*|'.

  wa_functxt-icon_id    = icon_export.
  wa_functxt-quickinfo  = '下载模板文件'.
  wa_functxt-icon_text  = '下载模板文件'.
  sscrfields-functxt_01 = wa_functxt.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form FRM_BATCH_UPDATE
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM frm_batch_update .
  DATA: l_projn(24).
  CHECK gt_itab[] IS NOT INITIAL.
  CLEAR: g_uflag.
  g_uflag = con_x.
**=============================================
**BDC CO02 批量取消关闭生产订单
**=============================================
  CLEAR: gs_ctu_params.
  "GS_CTU_PARAMS-DISMODE = 'A'.
  gs_ctu_params-dismode = 'N'.
  gs_ctu_params-updmode = 'S'.
  gs_ctu_params-defsize = 'X'.
**
  LOOP AT gt_itab ASSIGNING FIELD-SYMBOL(<fs_itab>).
**BDC
    CLEAR: gs_bdcd,gt_bdcd[].
** 初始屏幕
    PERFORM frm_dynpro USING 'SAPLCOKO1' '0110' 'X'.
    PERFORM frm_dynpro USING ' ' 'BDC_CURSOR'  'CAUFVD-AUFNR'.
    PERFORM frm_dynpro USING ' ' 'BDC_OKCODE'  '=ENTK'.
    PERFORM frm_dynpro USING ' ' 'CAUFVD-AUFNR' <fs_itab>-aufnr.
    PERFORM frm_dynpro USING ' ' 'R62CLORD-FLG_OVIEW' 'X'.

    PERFORM frm_dynpro USING 'SAPLCOKO1' '0115' 'X'.
    PERFORM frm_dynpro USING ' ' 'BDC_CURSOR'  'CAUFVD-GAMNG'.
**
    IF g_r1 = con_x. "取消TECO
      PERFORM frm_dynpro USING ' ' 'BDC_OKCODE'  '=TABR'. "撤消技术性完成
    ELSE.
      PERFORM frm_dynpro USING ' ' 'BDC_OKCODE'  '=RABK'. "取消关闭/完成
    ENDIF.

    PERFORM frm_dynpro USING 'SAPLCOKO1' '0115' 'X'.
    PERFORM frm_dynpro USING ' ' 'BDC_CURSOR'  'CAUFVD-GAMNG'.
    PERFORM frm_dynpro USING ' ' 'BDC_OKCODE'  '=VERW'.

    PERFORM frm_dynpro USING 'SAPLCOKO1' '0115' 'X'.
    PERFORM frm_dynpro USING ' ' 'BDC_CURSOR'  'CAUFVD-CH_PROC'.
    PERFORM frm_dynpro USING ' ' 'BDC_OKCODE'  '=BU'.

    CLEAR: gs_bmsg,gt_bmsg[].
    CALL TRANSACTION 'CO02' USING gt_bdcd
                     OPTIONS FROM gs_ctu_params
                    MESSAGES INTO gt_bmsg.
    " BREAK ABAP_LLP.
    CLEAR: gv_msgtx.
    READ TABLE gt_bmsg INTO gs_bmsg WITH KEY msgtyp = 'S'
                                             msgid  = 'CO'
                                             msgnr  = '100'.
    IF sy-subrc = 0 AND gs_bmsg-msgv1 IS NOT INITIAL.
      <fs_itab>-msgtyp = con_s.
      <fs_itab>-micon  = icon_green_light.
      <fs_itab>-msgtxt = TEXT-m03.
    ELSE.
      <fs_itab>-msgtyp = con_e.
      <fs_itab>-micon  = icon_red_light.
      CLEAR: <fs_itab>-msgtxt,gv_msgtx.
      LOOP AT gt_bmsg INTO gs_bmsg.
        CALL FUNCTION 'MESSAGE_TEXT_BUILD'
          EXPORTING
            msgid               = gs_bmsg-msgid
            msgnr               = gs_bmsg-msgnr
            msgv1               = gs_bmsg-msgv1
            msgv2               = gs_bmsg-msgv2
            msgv3               = gs_bmsg-msgv3
            msgv4               = gs_bmsg-msgv4
          IMPORTING
            message_text_output = gv_msgtx.
        CONDENSE gv_msgtx NO-GAPS.
        IF <fs_itab>-msgtxt IS INITIAL.
          <fs_itab>-msgtxt = gv_msgtx.
        ELSE.
          CONCATENATE <fs_itab>-msgtxt ';' gv_msgtx INTO <fs_itab>-msgtxt.
        ENDIF.
      ENDLOOP.
    ENDIF.
  ENDLOOP.
ENDFORM.

*&---------------------------------------------------------------------*
*&      Form  FRM_DYNPRO
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      -->P_A   text
*      -->P_B   text
*      -->P_C   text
*----------------------------------------------------------------------*
FORM frm_dynpro USING  p_a p_b p_c.
  CLEAR gs_bdcd.
  IF p_a NE space.
    MOVE: p_a TO gs_bdcd-program,
          p_b TO gs_bdcd-dynpro,
          p_c TO gs_bdcd-dynbegin.
  ELSE.
    MOVE: p_b TO gs_bdcd-fnam,
          p_c TO gs_bdcd-fval.
  ENDIF.
  APPEND gs_bdcd TO gt_bdcd.
ENDFORM.                    " FRM_DYNPRO
*&---------------------------------------------------------------------*
*&      Form  FM_SELECT_FILE
*&---------------------------------------------------------------------*
*       选择文件
*----------------------------------------------------------------------*
FORM frm_select_file USING p_filter
                           p_filename.

  DATA: l_filename TYPE string.
  DATA: lt_file_table TYPE filetable,
        l_rc          TYPE i.

  REFRESH: lt_file_table.

  CALL METHOD cl_gui_frontend_services=>file_open_dialog
    EXPORTING
      file_filter             = p_filter
    CHANGING
      file_table              = lt_file_table
      rc                      = l_rc
    EXCEPTIONS
      file_open_dialog_failed = 1
      cntl_error              = 2
      error_no_gui            = 3
      not_supported_by_gui    = 4
      OTHERS                  = 5.
  IF sy-subrc = 0.
    READ TABLE lt_file_table INTO l_filename INDEX 1.
    p_filename = l_filename.
  ENDIF.

ENDFORM. " FM_SELECT_FILE
*&---------------------------------------------------------------------*
*& Form FRM_UPLOAD_DATA
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM frm_upload_data .
  DATA: lt_xlstmp TYPE STANDARD TABLE OF alsmex_tabline,
        ls_xlstmp LIKE LINE OF lt_xlstmp.
  DATA: l_filename LIKE rlgrap-filename.
  DATA: l_row(4),
        lv_no_sh.
  DATA: lv_tabix TYPE sy-tabix.
  DATA: lv_flag TYPE c.


  IF p_path IS INITIAL.
    MESSAGE s000 WITH '上载路径不能为空！' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  l_filename = p_path.
  CLEAR ls_xlstmp.
  REFRESH lt_xlstmp[].
  CALL FUNCTION 'ALSM_EXCEL_TO_INTERNAL_TABLE'
    EXPORTING
      filename                = l_filename
      i_begin_col             = 1
      i_begin_row             = 2
      i_end_col               = 7
      i_end_row               = 60000
    TABLES
      intern                  = lt_xlstmp
    EXCEPTIONS
      inconsistent_parameters = 1
      upload_ole              = 2
      OTHERS                  = 3.
  IF sy-subrc NE 0 .
    MESSAGE e000 WITH p_path '地址有误'.
  ENDIF.

  IF lt_xlstmp[] IS INITIAL.
    MESSAGE s000 WITH '上载数据为空!' DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  CLEAR: gs_upld,gt_upld.
  LOOP AT lt_xlstmp INTO ls_xlstmp.
    CASE ls_xlstmp-col.
      WHEN '1'.
        gs_upld-aufnr = ls_xlstmp-value.
      WHEN OTHERS.
        ...
    ENDCASE.
**
    AT END OF row.
      TRANSLATE gs_upld TO UPPER CASE.
      APPEND gs_upld TO gt_upld.
      CLEAR gs_upld.
    ENDAT.
  ENDLOOP.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form FRM_CHECK_DATA
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM frm_check_data .
  DATA: lv_aufnr TYPE aufnr.
  "LOOP AT GT_UPLD INTO GS_UPLD.
  LOOP AT s_aufnr.
    CLEAR: gs_itab.
    g_seqno = g_seqno + 1.
    gs_itab-seqno = g_seqno.
    gs_itab-aufnr = s_aufnr-low.
**
    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
      EXPORTING
        input  = gs_itab-aufnr
      IMPORTING
        output = lv_aufnr.

    SELECT SINGLE objnr INTO @DATA(lv_objnr)
        FROM caufv
      WHERE aufnr = @lv_aufnr.
    IF sy-subrc <> 0.
      gs_itab-msgtxt = '生产订单号不存在'.
    ELSE.
      IF g_r1 = con_x. "取消TECO
        SELECT SINGLE objnr INTO lv_objnr
            FROM jest
          WHERE objnr = lv_objnr
            AND stat = con_st45 "已TECO
            AND inact = ''.
        IF sy-subrc <> 0.
          gs_itab-msgtxt = '生产订单状态不能取消TECO'.
        ENDIF.
**已结算生产订单
        SELECT SINGLE objnr INTO lv_objnr
            FROM jest
          WHERE objnr = lv_objnr
            AND stat = con_st46 "已结算
            AND inact = ''.
        IF sy-subrc = 0.
          gs_itab-msgtxt = gs_itab-aufnr && '已结算，不允许修改REL'.
        ENDIF.

      ELSE.
**已结算生产订单
        SELECT SINGLE objnr INTO lv_objnr
            FROM jest
          WHERE objnr = lv_objnr
            AND stat = con_st46 "已结算
            AND inact = ''.
        IF sy-subrc <> 0.
          gs_itab-msgtxt = '生产订单状态不能取消关闭'.
        ENDIF.
      ENDIF.
    ENDIF.

    IF gs_itab-msgtxt IS NOT INITIAL.
      gs_itab-micon  = icon_red_light.
      gs_itab-msgtyp = con_e.
    ENDIF.

    APPEND gs_itab TO gt_itab.
    CLEAR: gs_itab.
  ENDLOOP.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form FRM_DOWNLOAD_TMP
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM frm_download_tmp .
  DATA:lt_map_tab   TYPE ztpp_output_table_style.
  DATA:lv_save_ok   TYPE boolean.
  DATA:lt_template  TYPE STANDARD TABLE OF  ty_upld.
  DATA:lv_file      TYPE string.

  FIELD-SYMBOLS:<fs_template> TYPE ty_upld.
  FIELD-SYMBOLS:<fs_map_tab>  TYPE zspp_output_table_style.


  DEFINE mac_display_col.
    APPEND INITIAL LINE TO lt_map_tab ASSIGNING <fs_map_tab>.
    <fs_map_tab>-fieldname   = &1.
    <fs_map_tab>-headertext  = &2.
  END-OF-DEFINITION.

  PERFORM frm_save_file_name  CHANGING lv_file.

  IF lv_file = space.
    EXIT.
  ELSE.

*&  设置列标题
    mac_display_col      'AUFNR'  '生产订单号'  .

*& 设置样例数据
    APPEND INITIAL LINE TO lt_template ASSIGNING <fs_template>.
    <fs_template>-aufnr = '2100031826'  .

    lv_save_ok = zcl_interface_utils=>generate_excel_file( EXPORTING it_data          = lt_template
                                                                     it_mapping_table = lt_map_tab
                                                                     iv_filename      = lv_file ).
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*&      Form  FRM_SAVE_FILE_NAME
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
*      <--P_LV_SEL_FILE  text
*----------------------------------------------------------------------*
FORM frm_save_file_name  CHANGING cv_sel_file.

  DATA: l_fieldname(40) TYPE c.
  DATA: lv_filename     TYPE string.
  DATA: lv_fullpath     TYPE string.
  DATA: lv_path         TYPE string.

  FIELD-SYMBOLS <fs>.

  l_fieldname = 'FNAME'.
  CALL METHOD cl_gui_frontend_services=>file_save_dialog
    EXPORTING
      file_filter = '*.xls|*.xlsx'
    CHANGING
      filename    = lv_filename
      fullpath    = lv_fullpath
      path        = lv_path.
  IF sy-subrc = 0.
*    ASSIGN (L_FIELDNAME) TO <FS>.
*    <FS> = VL_FULLPATH.
    cv_sel_file = lv_fullpath.
  ELSE.
    cv_sel_file = space.
  ENDIF.

ENDFORM.
