�
 TCONTRACTFORM 0�!  TPF0TContractFormContractFormLeftjTop� BorderIconsbiSystemMenu BorderStylebsDialogCaption��������ClientHeightClientWidthColor	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	PositionpoScreenCenterOnShowFormShowPixelsPerInch`
TextHeight TPageControlContractPageControlLeft Top(WidthHeight� 
ActivePagecontractTabSheetTabOrder  	TTabSheetcontractTabSheetCaption�������� �� ������
ImageIndex TLabelLabel1LeftTopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel2Left� TopWidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  	TMaskEditcontractBeginMaskEditLeft+TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrder Text
  .  .      	TMaskEditcontractEndMaskEditLeft� TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .       	TTabSheetchangeTabSheetCaption������������ ���������
ImageIndex TLabelLabel3LeftTopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
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
  .  .      	TMaskEditchangeBeginMaskEditLeft+TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      	TMaskEditchangeEndMaskEditLeft� TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .       	TTabSheetall_dolgiContractTabSheetCaption�������� �������������
ImageIndex TLabelLabel5Left� TopWidthHeightCaption%������ �� ��������� �� ��������� ����Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 
ParentFont  	TMaskEditall_dolgiContractEndMaskEditLeft+TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrder Text
  .  .        TCoolBarmainCoolBarLeft Top WidthHeight$BandsControlmainToolBar
ImageIndex�	MinHeightWidth�  Ctl3D TToolBarmainToolBarLeft	Top Width�HeightCaptionmainToolBar	EdgeInner	esLoweredTabOrder  TSpeedButtonsbReportToExcelLeft TopWidthHeightHint������������ �����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 330     337wwwww330�����?����3�� 0  � �w�wws�w7�������ws3?��3����  ��w�3wws�7࿿�����w�3?��3����  ��w�3wws�7࿿�����w�?���3���   ��w�wwws77 ������ws��333730 �����37w���?�330� �  337�w3ww330����337���7330���337wwws330��� 3337���w3330   3337wwws3	NumGlyphsParentShowHintShowHint	OnClicksbReportToExcelClick  TToolButtonToolButton1LeftTopWidthCaptionToolButton1Style
tbsDivider  TSpeedButtonExitSpeedButtonLeft$TopWidthHeightHint�����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 3     33wwwww33333333333333333333333333333333333333333333333?33�33333s3333333333333���33��337ww�33��337���33��337ww3333333333333����33     33wwwwws3	NumGlyphsParentShowHintShowHint	OnClickExitSpeedButtonClick    TQueryallContractOldQueryDatabaseNamemy_disSQL.StringsRselect contragent.enterpr_id contragent_id, contragent.enterprise_name contragent,?storona.enterpr_id storona_id ,storona.enterprise_name storona,.c.contract_no, c.signing_date, c.contract_sum,ct.contract_type2from (((contract c left join enterpr contragent on>contragent.enterpr_id=c.enterpr_id)left join contract_sides cs#on  c.contract_no = cs.contract_no)Aleft join enterpr storona on cs.enterpr_id = storona.enterpr_id),contract_types ct#where c.signing_date >= :begin_dateand c.signing_date <= :end_date,and c.contract_type_id = ct.contract_type_id?order by c.signing_date, c.contract_no, storona.enterprise_name LeftTop� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryallContractQueryDatabaseNamemy_disSQL.Strings6select * from get_contract_info(:begin_date,:end_date)3order by ENTERPRISE_NAME, contract_no, signing_date Left@Top� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryall_dolgiContractQueryDatabaseNamemy_disSQL.Strings+select * from ALL_CONTRACT_SALDO(:pay_date)=where ((debit - credit)) > 0.01 or ((debit - credit) < -0.01)$order by enterprise_id, signing_date Left� Top� 	ParamDataDataType	ftUnknownNamepay_date	ParamType	ptUnknown    TQuerydolgi_prihodQueryDatabaseNamemy_disSQL.Stringsselect    o.creditor_ID enterpr_id,    o.TYPE_ID,    st.type_name,    o.SOURCE_ID,    o.AMOUNTHRIVN amount,    o.PAY_DATE,    o.CONTRACT_NO,    o.COMMENTS,from (operations o left join source_types ston o.type_id = st.type_id)where o.creditor_id = :ent_id and o.contract_no = :contract_noand o.pay_date <= :pay_dateorder by o.PAY_DATE Left� Top� 	ParamDataDataType	ftUnknownNameent_id	ParamType	ptUnknown DataType	ftUnknownNamecontract_no	ParamType	ptUnknown DataType	ftUnknownNamepay_date	ParamType	ptUnknown    TQuerydolgi_rashodQueryDatabaseNamemy_disSQL.Stringsselect    o.debitor_ID enterpr_id,    o.TYPE_ID,    st.type_name,    o.SOURCE_ID,    o.AMOUNTHRIVN amount,    o.PAY_DATE,    o.CONTRACT_NO,    o.COMMENTS,from (operations o left join source_types ston o.type_id = st.type_id)where o.debitor_id = :ent_id and o.contract_no = :contract_noand o.pay_date <= :pay_dateorder by o.PAY_DATE Left� Top� 	ParamDataDataType	ftUnknownNameent_id	ParamType	ptUnknown DataType	ftUnknownNamecontract_no	ParamType	ptUnknown DataType	ftUnknownNamepay_date	ParamType	ptUnknown    TQuerycount_dolgi_prihodQueryDatabaseNamemy_disSQL.Stringsselect count(*) count_rec,from (operations o left join source_types ston o.type_id = st.type_id)where o.creditor_id = :ent_id and o.contract_no = :contract_noand o.pay_date <= :pay_dateorder by o.PAY_DATE Left� Top� 	ParamDataDataType	ftUnknownNameent_id	ParamType	ptUnknown DataType	ftUnknownNamecontract_no	ParamType	ptUnknown DataType	ftUnknownNamepay_date	ParamType	ptUnknown    TQuerycount_dolgi_rashodQueryDatabaseNamemy_disSQL.Stringsselect count(*) count_rec,from (operations o left join source_types ston o.type_id = st.type_id)where o.debitor_id = :ent_id and o.contract_no = :contract_noand o.pay_date <= :pay_dateorder by o.PAY_DATE Left� Top� 	ParamDataDataType	ftUnknownNameent_id	ParamType	ptUnknown DataType	ftUnknownNamecontract_no	ParamType	ptUnknown DataType	ftUnknownNamepay_date	ParamType	ptUnknown     