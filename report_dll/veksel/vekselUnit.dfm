�
 TVEKSELEXPORTFORM 0E  TPF0TVekselExportFormVekselExportFormLeftTop� BorderIconsbiSystemMenu BorderStylebsDialogCaption������� ��������ClientHeight� ClientWidth�Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	PositionpoScreenCenterOnShowFormShowPixelsPerInch`
TextHeight TPageControlVekselPageControlLeft Top(Width�Height� 
ActivePageforAllTabSheetTabOrder  	TTabSheetforAllTabSheetCaption���
ImageIndex TLabelLabel1LeftTopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel2Left� TopWidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont   	TTabSheetforEnterprTabSheetCaption�� ������������
ImageIndex TLabelLabel4LeftTopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel6Left� TopWidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont   	TTabSheetforSaldoPayVekselTabSheetCaption������� ��������
ImageIndexOnHideforSaldoPayVekselTabSheetHideOnShowforSaldoPayVekselTabSheetShow TLabelLabel5LeftTTopWidthAHeightCaption	������ ��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont   	TTabSheetforSaldoSaleVekselTabSheetCaption������� ��������
ImageIndexOnHideforSaldoSaleVekselTabSheetHideOnShowforSaldoSaleVekselTabSheetShow TLabelLabel10LeftTTopWidthAHeightCaption	������ ��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont   	TTabSheetchangeVekselTabSheetCaption��������� �� ��������
ImageIndex TLabelLabel3LeftTopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel7Left� TopWidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel8Left� TopWidthzHeightCaption(�������� ������)Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont  TLabelLabel9Left� TopDWidth� HeightCaption(��������� ������� � ...)Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont  	TMaskEditJournalDateMaskEditLeft(Top@WidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrder Text
  .  .        	TMaskEditVekselBeginMaskEditLeft+TopXWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      	TMaskEditVekselEndMaskEditLeft� TopXWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      TCoolBarmainCoolBarLeft Top Width�Height$BandsControlmainToolBar
ImageIndex�	MinHeightWidth�  Ctl3D TToolBarmainToolBarLeft	Top Width�HeightCaptionmainToolBar	EdgeInner	esLoweredTabOrder  TSpeedButtonsbReportToExcelLeft TopWidthHeightHint������������ �����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 330     337wwwww330�����?����3�� 0  � �w�wws�w7�������ws3?��3����  ��w�3wws�7࿿�����w�3?��3����  ��w�3wws�7࿿�����w�?���3���   ��w�wwws77 ������ws��333730 �����37w���?�330� �  337�w3ww330����337���7330���337wwws330��� 3337���w3330   3337wwws3	NumGlyphsParentShowHintShowHint	OnClicksbReportToExcelClick  TToolButtonToolButton1LeftTopWidthCaptionToolButton1Style
tbsDivider  TSpeedButtonExitSpeedButtonLeft$TopWidthHeightHint�����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 3     33wwwww33333333333333333333333333333333333333333333333?33�33333s3333333333333���33��337ww�33��337���33��337ww3333333333333����33     33wwwwws3	NumGlyphsParentShowHintShowHint	OnClickExitSpeedButtonClick    TQueryallVekselQueryDatabaseNamemy_disLeftTop� 	ParamData   TQueryallVekselInContractQueryDatabaseNamemy_disSQL.Strings:select * from all_veksel_saldo_in('01.01.1996', :end_date)where (debit-credit <> 0)order by enterprise_name Left8Top� 	ParamDataDataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryGetContractDateQueryDatabaseNamemy_disSQL.Stringsselect signing_date from contract where contract_no = :contract_no Left� Top� 	ParamDataDataType	ftUnknownNamecontract_no	ParamType	ptUnknown    TQuerychangeVekselQueryDatabaseNamemy_disSQL.StringsKselect * from get_change_in_operations(:journal_date,:begin_date,:end_date)order by JOURNAL_DATE, debitor Left� Top� 	ParamDataDataType	ftUnknownNamejournal_date	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryallVekselOutContractQueryDatabaseNamemy_disSQL.Strings;select * from all_veksel_saldo_out('01.01.1996', :end_date)where (debit-credit <> 0)order by enterprise_name LeftXTop� 	ParamDataDataType	ftUnknownNameend_date	ParamType	ptUnknown     