�
 TCHOOSECONTRACTFORM 09  TPF0TChooseContractFormChooseContractFormLeft� TopyBorderIconsbiSystemMenu BorderStylebsDialogCaption����� ��������ClientHeight� ClientWidthaColor	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	PixelsPerInch`
TextHeight TPanelContractPanelLeft Top WidthaHeight� TabOrder  TDBGridContractDBGridLeftTopWidthQHeight� 
DataSourceChooseContractDataSourceTabOrder TitleFont.CharsetDEFAULT_CHARSETTitleFont.ColorclWindowTextTitleFont.Height�TitleFont.NameMS Sans SerifTitleFont.Style ColumnsExpanded	FieldNameCONTRACT_NOTitle.Caption�������Width� Visible	 Expanded	FieldNameSIGNING_DATETitle.Caption����Visible	 Expanded	FieldName	ROLE_NAMETitle.Caption����Width_Visible	     TBitBtnChooseBitBtnLeftTop� Width� Height!Caption�������Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontTabOrderKindbkOK  TBitBtnCacelBitBtnLeft� Top� Width� Height!Caption��������Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFontTabOrderKindbkCancel  TDataSourceChooseContractDataSourceDataSetChooseContractQueryLefthTopH  TQueryChooseContractQueryDatabaseNamemy_disSQL.Stringshselect distinct c.contract_id, cs.contract_no contract_no, c.signing_date, r.role_name, ct.contract_type>from contract c, roles r, contract_types ct, contract_sides cswhere cs.enterpr_id = :ent_id"and cs.contract_no = c.contract_no,and c.contract_type_id = ct.contract_type_idand cs.role_id = r.role_idorder by c.signing_date Left@TopH	ParamDataDataType	ftUnknownNameent_id	ParamType	ptUnknown     