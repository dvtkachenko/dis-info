�
 TNDSREPORTFORM 0�  TPF0TndsReportFormndsReportFormLeft�Top� BorderIconsbiSystemMenu BorderStylebsDialogCaption������������ ������ �� ���ClientHeight� ClientWidth�Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	PositionpoScreenCenterOnShowFormShowPixelsPerInch`
TextHeight TPageControlInvPageControlLeft Top(Width�Height� 
ActivePageforndsGeneralTabSheetTabOrder  	TTabSheetforndsGeneralTabSheetCaption����� �� ���
ImageIndex TLabelLabel3Left� TopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel4LeftTop
Width	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel6Left� Top
WidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  	TCheckBoxacceptCheckBoxLeftTop8Width� HeightCaption������������� �����Checked	EnabledState	cbCheckedTabOrder    	TTabSheetprotocol_ndsTabSheetCaption��������� ����. ������
ImageIndex TLabelLabel1LeftTop
Width	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel2Left� Top
WidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont    	TMaskEditndsBeginMaskEditLeft+TopHWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      	TMaskEditndsEndMaskEditLeft� TopHWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      TCoolBarmainCoolBarLeft Top Width�Height$BandsControlmainToolBar
ImageIndex�	MinHeightWidth�  Ctl3D TToolBarmainToolBarLeft	Top Width�HeightCaptionmainToolBar	EdgeInner	esLoweredTabOrder  TSpeedButtonsbReportToExcelLeft TopWidthHeightHint������������ �����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 330     337wwwww330�����?����3�� 0  � �w�wws�w7�������ws3?��3����  ��w�3wws�7࿿�����w�3?��3����  ��w�3wws�7࿿�����w�?���3���   ��w�wwws77 ������ws��333730 �����37w���?�330� �  337�w3ww330����337���7330���337wwws330��� 3337���w3330   3337wwws3	NumGlyphsParentShowHintShowHint	OnClicksbReportToExcelClick  TToolButtonToolButton1LeftTopWidthCaptionToolButton1Style
tbsDivider  TSpeedButtonExitSpeedButtonLeft$TopWidthHeightHint�����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 3     33wwwww33333333333333333333333333333333333333333333333?33�33333s3333333333333���33��337ww�33��337���33��337ww3333333333333����33     33wwwwws3	NumGlyphsParentShowHintShowHint	OnClickExitSpeedButtonClick    TQueryallInvQueryDatabaseNamemy_disLeftTop� 	ParamData   TQueryallDeptInQueryDatabaseNamemy_disSQL.Strings#select distinct dept_id , dept_name7from balans_report_all_invoices(:begin_date, :end_date)where payer_id = 0 LeftDTop� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryallDeptOutQueryDatabaseNamemy_disSQL.Strings#select distinct dept_id , dept_name7from balans_report_all_invoices(:begin_date, :end_date)where sender_id = 0 LeftdTop� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryTestInQueryDatabaseNamemy_disSQL.StringsDselect sum(amount) amount, sum(nds) nds from balans_invoices_list(1)where payer_id = 0  andpay_date >= :begin_date andpay_date <= :end_dateand is_in_oper = 'Y' LeftDTop� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryTestOutQueryDatabaseNamemy_disSQL.StringsDselect sum(amount) amount, sum(nds) nds from balans_invoices_list(1)where sender_id = 0  andpay_date >= :begin_date andpay_date <= :end_dateand is_in_oper = 'Y' LeftdTop� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryallServInvQueryDatabaseNamemy_disSQL.Strings  Left� Top� 	ParamData   TQuerypayCoalQueryDatabaseNamemy_disSQL.Strings+select d.enterprise_name, st.type_name, o.*-from operations o, source_types st, enterpr dwhere o.type_id = st.type_idand o.debitor_id = d.enterpr_idand o.pay_date >= :begin_dateand o.pay_date <= :end_dateand o.creditor_id = 0
and exists0(select * from b_is_coal_contract(o.contract_no) where is_coal = 'Y')order by o.debitor_id LeftTop� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryallProtocolZchtQueryDatabaseNamemy_disSQL.Strings1SELECT o.operation_id,e.enterprise_name creditor,'e1.enterprise_name debitor, O.PAY_DATE,CO.AMOUNTHRIVN,O.AMOUNT_USD, s.type_name, O.COMMENTS , o.contract_no9FROM OPERATIONS O, source_types s,  enterpr e, enterpr e1WHERE s.type_id = o.type_id"AND (o.creditor_id = e.enterpr_id)"AND (o.debitor_id = e1.enterpr_id)AND (o.type_id = 7)AND o.pay_date >= :begin_dateAND o.pay_date <= :end_date<ORDER BY o.creditor_id, O.PAY_DATE,O.AMOUNTHRIVN, O.COMMENTS Left<Top� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown     