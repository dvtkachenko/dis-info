�
 TFINDENTERPRISEFORM 0�	  TPF0TFindEnterpriseFormFindEnterpriseFormLeft� Top� BorderStylebsDialogCaption����� �����������ClientHeight� ClientWidth�Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	
OnActivateFormActivatePixelsPerInch`
TextHeight TPanelEnterprisePanelLeft Top Width�Height� TabOrder  TLabel	FindLabelLeftTopWidth� HeightCaption������� ������ ��� ������:Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style ParentColor
ParentFont  TEditEnterpriseEditLeft� TopWidthHeightTabOrder OnEnterEnterpriseEditEnter  TDBGridEnterpriseDBGridLeftTop Width�Height� 
DataSourceFindEnterpriseDataSourceTabOrderTitleFont.CharsetDEFAULT_CHARSETTitleFont.ColorclWindowTextTitleFont.Height�TitleFont.NameMS Sans SerifTitleFont.Style ColumnsExpanded	FieldNameOBJECT_NAMETitle.Caption������������ �����������Width�Visible	    TBitBtn
FindBitBtnLeft�TopWidth)HeightCaption�����Default	TabOrderOnClickFindBitBtnClick   TBitBtnChooseBitBtnLeftTop� Width� Height!Caption�������Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style ModalResult
ParentFontTabOrder
Glyph.Data
�  �  BM�      v   (   $            h                      �  �   �� �   � � ��  ��� ���   �  �   �� �   � � ��  ��� 333333333333333333  333333333333�33333  334C33333338�33333  33B$3333333�8�3333  34""C33333833�3333  3B""$33333�338�333  4"*""C3338�8�3�333  2"��"C3338�3�333  :*3:"$3338�38�8�33  3�33�"C333�33�3�33  3333:"$3333338�8�3  33333�"C333333�3�3  33333:"$3333338�8�  333333�"C333333�3�  333333:"C3333338�  3333333�#3333333��  3333333:3333333383  333333333333333333  	NumGlyphs  TBitBtnCancelBitBtnLeft Top� Width� Height!Caption��������Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontTabOrderKindbkCancel  TQueryFindEnterpriseQueryDatabaseNamemy_disSessionNameDefaultSQL.Strings<select e.enterprise_name object_name, e.enterpr_id object_idfrom enterpr eCwhere upper(e.enterprise_name collate pxw_cyrl) like upper(:Param1)order by e.enterpr_id Left8Top	ParamDataDataTypeftStringNameParam1	ParamType	ptUnknown    TDataSourceFindEnterpriseDataSourceDataSetFindEnterpriseQueryLeft8Top8   