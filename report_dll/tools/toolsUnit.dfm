�
 TTOOLSFORM 0�  TPF0
TtoolsForm	toolsFormLeft�Top� BorderIconsbiSystemMenu BorderStylebsDialogCaption������ClientHeight ClientWidth�Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	PositionpoScreenCenterOnShowFormShowPixelsPerInch`
TextHeight TPageControltoolsPageControlLeft Top(Width�Height� 
ActivePageavg_rateTabSheetTabOrder  	TTabSheetavg_rateTabSheetCaption���� ������
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
ParentFont  	TMaskEditarBeginMaskEditLeft#TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrder Text
  .  .      	TMaskEditarEndMaskEditLeft� TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      TRadioGroupcurrencyRadioGroupLeftTop(Width� HeightYCaption����� ������	ItemIndex Items.Strings��������������� TabOrder   	TTabSheetchangeOperationsTabSheetCaption��������� �� ���� ���������
ImageIndex TLabelLabel1Left� Top
WidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel2LeftTop
Width	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel8Left� TopWidthzHeightCaption(�������� ������)Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont  TLabelLabel9Left� TopDWidth� HeightCaption(��������� ������� � ...)Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont  	TMaskEditchopBeginMaskEditLeft#TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrder Text
  .  .      	TMaskEditchopEndMaskEditLeft� TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      	TMaskEditJournalDateMaskEditLeft(Top@WidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .        TCoolBarmainCoolBarLeft Top Width�Height$BandsControlmainToolBar
ImageIndex�	MinHeightWidth�  Ctl3D TToolBarmainToolBarLeft	Top Width�HeightCaptionmainToolBar	EdgeInner	esLoweredTabOrder  TSpeedButtonsbReportToExcelLeft TopWidthHeightHint������������ �����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 330     337wwwww330�����?����3�� 0  � �w�wws�w7�������ws3?��3����  ��w�3wws�7࿿�����w�3?��3����  ��w�3wws�7࿿�����w�?���3���   ��w�wwws77 ������ws��333730 �����37w���?�330� �  337�w3ww330����337���7330���337wwws330��� 3337���w3330   3337wwws3	NumGlyphsParentShowHintShowHint	OnClicksbReportToExcelClick  TToolButtonToolButton1LeftTopWidthCaptionToolButton1Style
tbsDivider  TSpeedButtonExitSpeedButtonLeft$TopWidthHeightHint�����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 3     33wwwww33333333333333333333333333333333333333333333333?33�33333s3333333333333���33��337ww�33��337���33��337ww3333333333333����33     33wwwwws3	NumGlyphsParentShowHintShowHint	OnClickExitSpeedButtonClick    	TGroupBoxruleGroupBoxLeftTop� Width�Height1Caption�������������� ������� TabOrder  TQueryavg_rate_disQueryDatabaseNamemy_disSQL.Strings$select rate_date, rate from cur_ratewhere currency_id = :cur_idand rate_date >= :begin_dateand rate_date <= :end_dateorder by rate_date Left� Top� 	ParamDataDataType	ftUnknownNamecur_id	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryavg_rate_isdQueryDatabaseName
my_ora_isdSQL.Strings*select rate_date, rate from admin.cur_ratewhere currency_id = :cur_idand rate_date >= :begin_dateand rate_date <= :end_dateorder by rate_date Left� Top� 	ParamDataDataType	ftUnknownNamecur_id	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryget_changeOperationsQueryDatabaseNamemy_disSQL.StringsRselect * from get_change_in_operations_modify(:journal_date,:begin_date,:end_date)$order by OJ_JOURNAL_DATE, OJ_debitor Left4Top� 	ParamDataDataType	ftUnknownNamejournal_date	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown     