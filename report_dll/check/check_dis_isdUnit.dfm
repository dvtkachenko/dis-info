�
 TCOMP_DIS_ISDFORM 0^  TPF0Tcomp_dis_isdFormcomp_dis_isdFormLeft�Top� BorderIconsbiSystemMenu BorderStylebsDialogCaption0������ ������������ ������ � �� ���98 � ��� 2000ClientHeight ClientWidth�Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	PositionpoScreenCenterOnShowFormShowPixelsPerInch`
TextHeight TPageControlcomp_dis_isdPageControlLeft Top(Width�Height� 
ActivePagecomp_isd_disTabSheetTabOrder  	TTabSheetcomp_isd_disTabSheetCaption������
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
ParentFont    	TMaskEditCompBeginMaskEditLeft+TopHWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      	TMaskEditCompEndMaskEditLeft� TopHWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      	TGroupBoxruleGroupBoxLeftTop� Width�Height1Caption�������������� ������� EnabledTabOrder 	TCheckBoxchainCheckBoxLeftTopWidth� HeightCaption�������� ��������� �����Checked	EnabledState	cbCheckedTabOrder   	TCheckBoxSkidkiPriplCheckBoxLeft� TopWidth� HeightCaption�������� �������� ������EnabledTabOrder   TCoolBarmainCoolBarLeft Top Width�Height$BandsControlmainToolBar
ImageIndex�	MinHeightWidth�  Ctl3D TToolBarmainToolBarLeft	Top Width�HeightCaptionmainToolBar	EdgeInner	esLoweredTabOrder  TSpeedButtonsbReportToExcelLeft TopWidthHeightHint������������ �����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 330     337wwwww330�����?����3�� 0  � �w�wws�w7�������ws3?��3����  ��w�3wws�7࿿�����w�3?��3����  ��w�3wws�7࿿�����w�?���3���   ��w�wwws77 ������ws��333730 �����37w���?�330� �  337�w3ww330����337���7330���337wwws330��� 3337���w3330   3337wwws3	NumGlyphsParentShowHintShowHint	OnClicksbReportToExcelClick  TToolButtonToolButton1LeftTopWidthCaptionToolButton1Style
tbsDivider  TSpeedButtonExitSpeedButtonLeft$TopWidthHeightHint�����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 3     33wwwww33333333333333333333333333333333333333333333333?33�33333s3333333333333���33��337ww�33��337���33��337ww3333333333333����33     33wwwwws3	NumGlyphsParentShowHintShowHint	OnClickExitSpeedButtonClick    TQuerysaldo_contract_isdQueryDatabaseName
my_ora_isdSQL.StringsOselect nvl(sum(decode(debitor_id,0,amount*rate,-amount*rate)),0) saldo_contractfrom admin.document dwhere,(d.debitor_id= 3493 or d.creditor_id = 3493)and d.wire_date <= :saldo_dateand d.accepted='Y'and d.bk_received='Y'and d.in_balance='Y' and d.contract_id = :contract_id LeftTop� 	ParamDataDataType	ftUnknownName
saldo_date	ParamType	ptUnknown DataType	ftUnknownNamecontract_id	ParamType	ptUnknown    TQueryall_contr_oper_isdQueryDatabaseName
my_ora_isdSQL.Strings$select cntr.document_id contract_id,cntr.document_no contract_no,cntr.doc_date signing_date,dt.type_name,;nvl(sum(decode(d.debitor_id,0,d.amount*d.rate,0)),0) debit,<nvl(sum(decode(d.creditor_id,0,d.amount*d.rate,0)),0) credit?from admin.document d , admin.doc_type dt , admin.document cntr3where (d.creditor_id = 3493 or d.debitor_id = 3493)"and d.doc_type_id = dt.doc_type_id(and d.contract_id = cntr.document_id (+)<and  d.wire_date >= :begin_date and d.wire_date <= :end_dateand d.bk_received = 'Y'and d.in_balance = 'Y'and d.accepted = 'Y'Fgroup by cntr.document_no,cntr.doc_date,cntr.document_id, dt.type_name  LeftTop� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryall_contr_oper_disQueryDatabaseNamemy_disSQL.Strings4select o.contract_no, c.signing_date , st.type_name,sum(o.debit_hrv) debit,sum(o.credit_hrv) credit3from operation_list2 o, source_types st, contract cwhere (o.type_id = st.type_id)1and (o.debitor_id = 3493 or o.creditor_id = 3493)and (o.pay_date >= :begin_date)and (o.pay_date <= :end_date)#and (o.contract_no = c.contract_no)4group by o.contract_no, c.signing_date, st.type_name LeftDTop� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQuerysaldo_contract_disQueryDatabaseNamemy_disSQL.Strings/select sum(debit_hrv-credit_hrv) saldo_contractfrom operation_list(3493) where contract_no = :contract_idand pay_date <= :saldo_date LeftDTop� 	ParamDataDataType	ftUnknownNamecontract_id	ParamType	ptUnknown DataType	ftUnknownName
saldo_date	ParamType	ptUnknown     