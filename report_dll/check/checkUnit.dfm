�
 TCHECKDATAFORM 0�  TPF0TcheckDataFormcheckDataFormLeft�Top� BorderIconsbiSystemMenu BorderStylebsDialogCaption�������� ������������ ������ClientHeight ClientWidth�Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	PositionpoScreenCenterOnShowFormShowPixelsPerInch`
TextHeight TPageControlcheckDataPageControlLeft Top(Width�Height� 
ActivePageno_contract_relTabSheetTabOrder  	TTabSheetcontract_relTabSheetCaption�������� �������� � ���������
ImageIndex TLabelLabel3Left� TopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel4LeftTop
Width	HeightCaption�EnabledFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel6Left� Top
WidthHeightCaption��EnabledFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont   	TTabSheetno_contract_relTabSheetCaption�������� "��� ��������"
ImageIndexOnHideno_contract_relTabSheetHideOnShowno_contract_relTabSheetShow TLabelLabel1LeftTop
Width	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel2Left� Top
WidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont    	TMaskEditcrBeginMaskEditLeft+TopHWidthIHeightEnabledEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      	TMaskEditcrEndMaskEditLeft� TopHWidthIHeightEnabledEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      TCoolBarmainCoolBarLeft Top Width�Height$BandsControlmainToolBar
ImageIndex�	MinHeightWidth�  Ctl3D TToolBarmainToolBarLeft	Top Width�HeightCaptionmainToolBar	EdgeInner	esLoweredTabOrder  TSpeedButtonsbReportToExcelLeft TopWidthHeightHint������������ �����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 330     337wwwww330�����?����3�� 0  � �w�wws�w7�������ws3?��3����  ��w�3wws�7࿿�����w�3?��3����  ��w�3wws�7࿿�����w�?���3���   ��w�wwws77 ������ws��333730 �����37w���?�330� �  337�w3ww330����337���7330���337wwws330��� 3337���w3330   3337wwws3	NumGlyphsParentShowHintShowHint	OnClicksbReportToExcelClick  TToolButtonToolButton1LeftTopWidthCaptionToolButton1Style
tbsDivider  TSpeedButtonExitSpeedButtonLeft$TopWidthHeightHint�����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 3     33wwwww33333333333333333333333333333333333333333333333?33�33333s3333333333333���33��337ww�33��337���33��337ww3333333333333����33     33wwwwws3	NumGlyphsParentShowHintShowHint	OnClickExitSpeedButtonClick    	TGroupBoxruleGroupBoxLeftTop� Width�Height1Caption�������������� ������� TabOrder  TQuerycheckContract_relQueryDatabaseNamemy_disSQL.Strings<select d.enterprise_name debitor,c.enterprise_name creditor,=st.type_name, o.amounthrivn amount, o.pay_date, o.contract_no8FROM OPERATIONs O, enterpr d, enterpr c, source_types st#where (o.debitor_id = d.enterpr_id)"and (o.creditor_id = c.enterpr_id)and (o.type_id = st.type_id)(and (o.contract_no = '��� �������� !!!')order by o.pay_date LeftTop� 	ParamData   TQuerycheck_no_Contract_relQueryDatabaseNamemy_disSQL.Strings<select d.enterprise_name debitor,c.enterprise_name creditor,=st.type_name, o.amounthrivn amount, o.pay_date, o.contract_no8FROM OPERATIONs O, enterpr d, enterpr c, source_types st#where (o.debitor_id = d.enterpr_id)"and (o.creditor_id = c.enterpr_id)and (o.type_id = st.type_id)and (o.pay_date >= :begin_date)and (o.pay_date <= :end_date)$and (o.contract_no = '��� ��������')order by o.pay_date LeftTTop� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown     