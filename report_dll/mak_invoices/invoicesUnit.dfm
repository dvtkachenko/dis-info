�
 TMAKINVEXPORTFORM 0  TPF0TmakInvExportFormmakInvExportFormLeft�Top� BorderIconsbiSystemMenu BorderStylebsDialogCaption/������� ������-������ �� ����� ��� ������������ClientHeight� ClientWidth�Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	PositionpoScreenCenterOnShowFormShowPixelsPerInch`
TextHeight TPageControlInvPageControlLeft Top(Width�Height� 
ActivePageforEnterprTabSheetTabOrder  	TTabSheetforEnterprTabSheetCaption�� ������������
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
ParentFont  	TCheckBoxacceptCheckBoxLeftTop8Width� HeightCaption������������� �����Checked	EnabledState	cbCheckedTabOrder     	TMaskEditInvBeginMaskEditLeft+TopHWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      	TMaskEditInvEndMaskEditLeft� TopHWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      TCoolBarmainCoolBarLeft Top Width�Height$BandsControlmainToolBar
ImageIndex�	MinHeightWidth�  Ctl3D TToolBarmainToolBarLeft	Top Width�HeightCaptionmainToolBar	EdgeInner	esLoweredTabOrder  TSpeedButtonsbReportToExcelLeft TopWidthHeightHint������������ �����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 330     337wwwww330�����?����3�� 0  � �w�wws�w7�������ws3?��3����  ��w�3wws�7࿿�����w�3?��3����  ��w�3wws�7࿿�����w�?���3���   ��w�wwws77 ������ws��333730 �����37w���?�330� �  337�w3ww330����337���7330���337wwws330��� 3337���w3330   3337wwws3	NumGlyphsParentShowHintShowHint	OnClicksbReportToExcelClick  TToolButtonToolButton1LeftTopWidthCaptionToolButton1Style
tbsDivider  TSpeedButtonExitSpeedButtonLeft$TopWidthHeightHint�����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 3     33wwwww33333333333333333333333333333333333333333333333?33�33333s3333333333333���33��337ww�33��337���33��337ww3333333333333����33     33wwwwws3	NumGlyphsParentShowHintShowHint	OnClickExitSpeedButtonClick    TQueryallInvQueryDatabaseNamemy_disLeftTop� 	ParamData   TQueryInvoiceItemsQueryDatabaseNamemy_disSQL.Strings4select i.item_id, s.trade_mark, p.dimention, i.qnty,"i.price_without_nds, i.full_price,i.summ_without_nds, i.full_summ*from invoice_items i, supply s, products p"where (i.invoice_id = :invoice_id)and (i.supply_id = s.supply_id)and (s.prod_id = p.prod_id)and (s.prod_id <> 1320)order by i.item_id Left,Top� 	ParamDataDataType	ftUnknownName
invoice_id	ParamType	ptUnknown    TQueryExtraInvoiceItemsQueryDatabaseNamemy_disSQL.Strings;select e.extra_item_name, i.price_without_nds, i.full_price/from extra_items_guide e, invoice_extra_items i"where (i.invoice_id = :invoice_id)and (i.extra_id = e.extra_id)order by i.extra_item_id LeftTTop� 	ParamDataDataType	ftUnknownName
invoice_id	ParamType	ptUnknown    TQuerySkidkiPriplQueryDatabaseNamemy_disSQL.Strings,select  sum(i.summ_without_nds) skidki_priplfrom invoice_items i, supply s"where (i.invoice_id = :invoice_id)and (i.supply_id = s.supply_id)and (s.prod_id = 1320) Left|Top� 	ParamDataDataType	ftUnknownName
invoice_id	ParamType	ptUnknown     