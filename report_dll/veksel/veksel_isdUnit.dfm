�
 TVEKSEL_ISDEXPORTFORM 0M  TPF0TVeksel_isdExportFormVeksel_isdExportFormLeft�Top� BorderIconsbiSystemMenu BorderStylebsDialogCaption+������ � �������� �������� � ���������� ���ClientHeight� ClientWidth�Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	PositionpoScreenCenterOnShowFormShowPixelsPerInch`
TextHeight TPageControlVeksel_isdPageControlLeft Top(Width�Height� 
ActivePageforAllTabSheetTabOrder  	TTabSheetforAllTabSheetCaption���
ImageIndex TLabelLabel1LeftTopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel2Left� TopWidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont   	TTabSheetforEnterprTabSheetCaption�� ������������Enabled
ImageIndex TLabelLabel4LeftTopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel6Left� TopWidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont   	TTabSheetforSaldoPayVekselTabSheetCaption������� ��������Enabled
ImageIndex TLabelLabel5LeftTTopWidthAHeightCaption	������ ��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont   	TTabSheetchangeVekselTabSheetCaption��������� �� ��������Enabled
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
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 3     33wwwww33333333333333333333333333333333333333333333333?33�33333s3333333333333���33��337ww�33��337���33��337ww3333333333333����33     33wwwwws3	NumGlyphsParentShowHintShowHint	OnClickExitSpeedButtonClick    TQueryallisdVekselQueryDatabaseName
my_ora_isdSQL.Stringsselect
v.bill_id,[v.bill_no, v.emission_date, v.sight_date, v.emission_place, v.amount*v.rate nominal_amount,4bm.object_name bill_maker, bp.object_name bill_payer5from admin.v_bill v, admin.object bm, admin.object bpwherev.bill_maker = bm.object_id (+)#and v.bill_payer = bp.object_id (+)
and exists*(select * from document d, doc_relation drwheredr.doc_1=v.bill_idand dr.doc_2=d.document_idand d.bk_received = 'Y'+and (d.debitor_id = 0 or d.creditor_id = 0)and d.wire_date >= :begin_dateand d.wire_date <= :end_date) LeftTop� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryeventInForVekselQueryDatabaseName
my_ora_isdSQL.StringsselectAdt.type_name, d.wire_date, d.doc_date , (d.amount*d.rate) amount,3crd.object_name creditor, cntr.document_no contractfrom=admin.document d, admin.doc_relation dr, admin.document cntr,#admin.object crd, admin.doc_type dtwheredr.doc_1 = :veksel_idand dr.doc_2 = d.document_idand d.bk_received = 'Y'"and d.doc_type_id = dt.doc_type_id!and d.creditor_id = crd.object_id(and d.contract_id = cntr.document_id (+)and d.debitor_id = 0and d.wire_date >= :begin_dateand d.wire_date <= :end_date Left,Top� 	ParamDataDataType	ftUnknownName	veksel_id	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryeventOutForVekselQueryDatabaseName
my_ora_isdSQL.StringsselectAdt.type_name, d.wire_date, d.doc_date , (d.amount*d.rate) amount,3dbtr.object_name debitor, cntr.document_no contractfrom=admin.document d, admin.doc_relation dr, admin.document cntr,$admin.object dbtr, admin.doc_type dtwheredr.doc_1 = :veksel_idand dr.doc_2=d.document_idand d.bk_received = 'Y'"and d.doc_type_id = dt.doc_type_id!and d.debitor_id = dbtr.object_id(and d.contract_id = cntr.document_id (+)and d.creditor_id = 0and d.wire_date >= :begin_dateand d.wire_date <= :end_date LeftLTop� 	ParamDataDataType	ftUnknownName	veksel_id	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown     