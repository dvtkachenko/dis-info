�
 TSTATISTICREPORTFORM 0G8  TPF0TStatisticReportFormStatisticReportFormLeft� Top� BorderStylebsSingleCaption
����������ClientHeight� ClientWidth�Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	PositionpoScreenCenterPixelsPerInch`
TextHeight TCoolBarmainCoolBarLeft Top Width�Height$BandsControlmainToolBar
ImageIndex�	MinHeightWidth�  Ctl3D TToolBarmainToolBarLeft	Top Width�HeightCaptionmainToolBar	EdgeInner	esLoweredTabOrder  TSpeedButtonsbReportToExcelLeft TopWidthHeightHint��������� �����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 330     337wwwww330�����?����3�� 0  � �w�wws�w7�������ws3?��3����  ��w�3wws�7࿿�����w�3?��3����  ��w�3wws�7࿿�����w�?���3���   ��w�wwws77 ������ws��333730 �����37w���?�330� �  337�w3ww330����337���7330���337wwws330��� 3337���w3330   3337wwws3	NumGlyphsParentShowHintShowHint	OnClicksbReportToExcelClick  TToolButtonToolButton1LeftTopWidthCaptionToolButton1Style
tbsDivider  TSpeedButtonSpeedButton3Left$TopWidthHeightHint�����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 3     33wwwww33333333333333333333333333333333333333333333333?33�33333s3333333333333���33��337ww�33��337���33��337ww3333333333333����33     33wwwwws3	NumGlyphsParentShowHintShowHint	    TPageControlStatisticPageControlLeft Top(Width�Height� 
ActivePageStatisticTabSheet	MultiLine	TabOrder 	TTabSheetStatisticTabSheetCaption�� ����������� TLabelLabel1LeftTopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel2Left� TopWidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  	TMaskEditStatBeginMaskEditLeft(TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrder Text
  .  .      	TMaskEditStatEndMaskEditLeft� TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      	TCheckBoxStatisticCheckBoxLeftTop`Width1HeightCaption.�� ���� ����������� ���� �� ����������� ������TabOrder  TProgressBarProgressBar1LeftTop� Width�HeightMin MaxdTabOrder   	TTabSheetCoalSenderTabSheetCaption�� ���������� ����
ImageIndex TLabelChangeBeginLabelLeftTopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelChangeEndLabelLeft� TopWidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  	TMaskEditBeginMaskEditLeft(TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrder Text
  .  .      	TMaskEditEndMaskEditLeft� TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .      	TCheckBoxAllCoalCheckBoxLeftTop`Width1HeightCaption.�� ���� ����������� ���� �� ����������� ������TabOrder  TProgressBarCoalStatisticProgressBarLeftTop� Width�HeightMin MaxdTabOrder   	TTabSheetallPlategiTabSheetCaption������� �������
ImageIndex TLabelLabel3LeftTopWidth	HeightCaption�Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  TLabelLabel4Left� TopWidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  	TMaskEditPlatBeginMaskEditLeft(TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrder Text
  .  .      	TMaskEditPlatEndMaskEditLeft� TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrderText
  .  .        TQuery	CoalQueryDatabaseNamemy_disSessionNameDefaultSQL.StringsGselect invoice_no, pay_date, invoice_date, trade_mark, dimention, qnty,Kamount, nds, cargo_sender, cargo_receiver, cargo_date, is_in_oper, contract9from balans_report_input_coal_all(:begin_date, :end_date)where sender_id = :ent_idand amount <> 0order by  pay_date, invoice_no LeftTop`	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown DataType	ftUnknownNameent_id	ParamType	ptUnknown    TQuerydisInvoiceOutQueryDatabaseNamemy_disSessionNameDefaultSQL.Strings[SELECT distinct PAY_DATE, IS_IN_OPER, INVOICE_DATE, AMOUNT, AMOUNT_USD,  NDS, INVOICE_NO , 1short_trade_mark, cargo_date, contract, dept_nameDFROM  balans_report_output_part_ent(:ent_id, :begin_date, :end_date)&ORDER BY  PAY_DATE, INVOICE_NO, AMOUNT Left4Top`	ParamDataDataTypeftFloatNameent_id	ParamType	ptUnknownValue      @�	@ DataType
ftDateTimeName
begin_date	ParamType	ptUnknownValue    ���@ DataType
ftDateTimeNameend_date	ParamType	ptUnknownValue    ���@    TQuerydisInvoiceInQueryDatabaseNamemy_disSessionNameDefaultSQL.Strings@SELECT distinct IS_IN_OPER, PAY_DATE, INVOICE_DATE, AMOUNT, NDS,?INVOICE_NO , invoice_id, short_trade_mark, cargo_date, contractGFROM  balans_report_input_part_ent(:ent_id, :begin_date, :end_date) I, &invoice_items I1, supply s, products p$where (I1.INVOICE_ID = I.INVOICE_ID)! AND (S.SUPPLY_ID = I1.SUPPLY_ID) AND (P.PROD_ID = S.PROD_ID) AND (P.PROD_GROUP_ID <> 12.0)*ORDER BY  INVOICE_DATE, INVOICE_NO, AMOUNT Left\Toph	ParamDataDataTypeftFloatNameent_id	ParamType	ptUnknownValue      @�	@ DataType
ftDateTimeName
begin_date	ParamType	ptUnknownValue    ���@ DataType
ftDateTimeNameend_date	ParamType	ptUnknownValue    ���@    TQuerydisStatisticQueryCreditorDatabaseNamemy_dis_cyrrSessionNameDefaultSQL.Strings!SELECT O.PAY_DATE, O.AMOUNTHRIVN,DS.TYPE_NAME, CONTRACT_NO contract FROM OPERATIONS O, SOURCE_TYPES S WHERE (S.TYPE_ID = O.TYPE_ID)  AND (O.DEBITOR_ID = :debitor_id)AND (O.TYPE_ID <> 12)=AND (O.PAY_DATE >= :begin_date) AND (O.PAY_DATE <= :end_date)0ORDER BY  O.PAY_DATE, S.TYPE_NAME, O.AMOUNTHRIVN Left� Toph	ParamDataDataType	ftUnknownName
debitor_id	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQuerydisStatisticQueryDebitorDatabaseNamemy_dis_cyrrSessionNameDefaultSQL.Strings!SELECT O.PAY_DATE, O.AMOUNTHRIVN,DS.TYPE_NAME, CONTRACT_NO contract FROM OPERATIONS O, SOURCE_TYPES S WHERE (S.TYPE_ID = O.TYPE_ID) "AND (O.CREDITOR_ID = :creditor_id)AND (O.TYPE_ID <> 12)=AND (O.PAY_DATE >= :begin_date) AND (O.PAY_DATE <= :end_date)0ORDER BY  O.PAY_DATE, S.TYPE_NAME, O.AMOUNTHRIVN Left� Toph	ParamDataDataType	ftUnknownNamecreditor_id	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryallCoalSenderQueryDatabaseNamemy_disSessionNameDefaultSQL.Strings(select distinct sender_id, enterpr_name 9from balans_report_input_coal_all(:begin_date, :end_date)order by enterpr_name Left� Toph	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryallSaldoQueryDatabaseNamemy_disSessionNameDefaultSQL.StringsFSELECT sum(debit_hrv-credit_hrv) allSaldo FROM operation_list(:ent_id)where pay_date <= :saldo_date LeftToph	ParamDataDataType	ftUnknownNameent_id	ParamType	ptUnknown DataType	ftUnknownName
saldo_date	ParamType	ptUnknown    TQuerydisPlategiQueryDebitorDatabaseNamemy_dis_cyrrSQL.StringsGSELECT  B.DOC_DATE, S.TYPE_NAME, B.PAY_ORDER, B.DEBIT AMOUNT, B.CREDIT,5B.DEBIT/O.USD_RATE AMOUNT_USD, B.CREDIT/O.USD_RATE , $o.contract_no ,B.DESCRIPTION comment-FROM BDOC_ITM B, OPERATIONS O, source_types sWHERE (O.creditor_id = :ent_id)  AND (B.DOC_DATE >= :begin_date) AND (B.DOC_DATE <= :end_date) AND (O.TYPE_ID = 2.0)! AND (O.SOURCE_ID = B.BD_ITEM_ID) AND (s.type_id = o.type_id)5ORDER BY B.DOC_DATE, B.DEBIT, B.CREDIT, B.DESCRIPTION LeftTToph	ParamDataDataType
ftSmallintNameent_id	ParamType	ptUnknownValue� DataType
ftDateTimeName
begin_date	ParamType	ptUnknownValue    ���@ DataType
ftDateTimeNameend_date	ParamType	ptUnknownValue    ���@    TQuerydisPlategiQueryCreditorDatabaseNamemy_dis_cyrrSQL.StringsGSELECT  B.DOC_DATE,S.TYPE_NAME, B.PAY_ORDER, B.DEBIT , B.CREDIT AMOUNT,5B.DEBIT/O.USD_RATE , B.CREDIT/O.USD_RATE AMOUNT_USD, $o.contract_no ,B.DESCRIPTION comment-FROM BDOC_ITM B, OPERATIONS O, source_types sWHERE (O.debitor_id = :ent_id)  AND (B.DOC_DATE >= :begin_date) AND (B.DOC_DATE <= :end_date) AND (O.TYPE_ID = 2.0)! AND (O.SOURCE_ID = B.BD_ITEM_ID) AND (s.type_id = o.type_id)5ORDER BY B.DOC_DATE, B.DEBIT, B.CREDIT, B.DESCRIPTION  Left|Toph	ParamDataDataType	ftUnknownNameent_id	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQuerydisAnyQueryCreditorDatabaseNamemy_dis_cyrrSQL.Strings@SELECT O.PAY_DATE,s.type_name, I.ACT_NO, O.AMOUNTHRIVN  AMOUNT, ?O.AMOUNT_USD, o.contract_no, O.COMMENTS comment, I.SIGNING_DATE.FROM OPERATIONS O, INTERPAYM I, source_types sWHERE(I.ACT_ID = O.SOURCE_ID) and s.type_id = o.type_id# AND (I.SOURCE_TYPE_ID = O.TYPE_ID) AND (o.debitor_id = :ent_id) AND (o.type_id <> 2) AND (o.type_id <> 12) AND o.pay_date >= :begin_date AND o.pay_date <= :end_date:ORDER BY O.PAY_DATE, s.type_name, I.ACT_NO, O.AMOUNTHRIVN, o.contract_no, O.COMMENTS Left|Top� 	ParamDataDataType	ftUnknownNameent_id	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQuerydisAnyQueryDebitorDatabaseNamemy_dis_cyrrSQL.Strings@SELECT O.PAY_DATE,s.type_name, I.ACT_NO, O.AMOUNTHRIVN  AMOUNT, ?O.AMOUNT_USD, o.contract_no, O.COMMENTS comment, I.SIGNING_DATE.FROM OPERATIONS O, INTERPAYM I, source_types sWHERE(I.ACT_ID = O.SOURCE_ID) and s.type_id = o.type_id# AND (I.SOURCE_TYPE_ID = O.TYPE_ID) AND (o.creditor_id = :ent_id) AND (o.type_id <> 2) AND (o.type_id <> 12) AND o.pay_date >= :begin_date AND o.pay_date <= :end_date:ORDER BY O.PAY_DATE, s.type_name, I.ACT_NO, O.AMOUNTHRIVN, o.contract_no, O.COMMENTS LeftTTop� 	ParamDataDataType	ftUnknownNameent_id	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryallContragentQueryDatabaseNamemy_disSQL.Strings/select distinct e.enterpr_id, e.enterprise_namefrom operations o, enterpr e!where (o.pay_date >= :begin_date)and (o.pay_date <= :end_date)!and (o.debitor_id = e.enterpr_id) union /select distinct e.enterpr_id, e.enterprise_namefrom operations o, enterpr e!where (o.pay_date >= :begin_date)and (o.pay_date <= :end_date)"and (o.creditor_id = e.enterpr_id) union /select distinct e.enterpr_id, e.enterprise_nameDfrom balans_report_all_invoices(:begin_date, :end_date) b, enterpr e7where is_in_oper = 'N' and (b.sender_id = e.enterpr_id) union /select distinct e.enterpr_id, e.enterprise_nameDfrom balans_report_all_invoices(:begin_date, :end_date) b, enterpr e6where is_in_oper = 'N' and (b.payer_id = e.enterpr_id) Left,Top� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown DataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown    TQueryallPlategiQueryDatabaseNamemy_disSQL.Strings=SELECT DISTINCT B.BD_ITEM_ID, S.ACCOUNT_NAME , B.ACCOUNT_ID, 8   E.ENTERPRISE_NAME, T.BANK_NAME, B.DOC_DATE, B.DEBIT, -   B.CREDIT,  B.DEBIT/O.USD_RATE  DEBIT_USD, <   B.CREDIT/O.USD_RATE  CREDIT_USD, O.GAS_PAY, B.DESCRIPTIONGFROM ENTERPR E, PAYMDATA P, BDOC_ITM B, OPERATIONS O,ACCOUNTS S,BANKS TWHERE(P.ENTERPR_ID = E.ENTERPR_ID) AND  B.ACCOUNT_ID=S.ACCOUNT_ID AND B.MFO=T.MFO  AND (B.DOC_DATE >= :begin_date) AND (B.DOC_DATE <= :end_date) AND (B.MFO = P.MFO)" AND (B.ACCOUNT_NO = P.ACCOUNT_NO) AND (O.TYPE_ID = 2.0)! AND (O.SOURCE_ID = B.BD_ITEM_ID)8ORDER BY  B.DOC_DATE, S.ACCOUNT_NAME,  E.ENTERPRISE_NAME Left� Top� 	ParamDataDataType
ftDateTimeName
begin_date	ParamType	ptUnknownValue    ���@ DataType
ftDateTimeNameend_date	ParamType	ptUnknownValue    ���@    TQueryallCoalTestQueryDatabaseNamemy_disSQL.Strings2select sum(amount)  testAmount, sum(qnty) testQnty: from balans_report_input_coal_all(:begin_date, :end_date)where sender_id <> 0and qnty <> 0 LeftTop� 	ParamDataDataType	ftUnknownName
begin_date	ParamType	ptUnknown DataType	ftUnknownNameend_date	ParamType	ptUnknown     