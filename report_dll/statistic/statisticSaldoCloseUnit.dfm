�
 TSTATISTICSALDOCLOSEFORM 0  TPF0TstatisticSaldoCloseFormstatisticSaldoCloseFormLeftpTop� BorderIconsbiSystemMenu BorderStylebsSingleCaption
����������ClientHeight� ClientWidth�Color	clBtnFaceFont.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style OldCreateOrder	PositionpoScreenCenterOnShowFormShowPixelsPerInch`
TextHeight TCoolBarmainCoolBarLeft Top Width�Height$BandsControlmainToolBar
ImageIndex�	MinHeightWidth�  Ctl3D TToolBarmainToolBarLeft	Top Width�HeightCaptionmainToolBar	EdgeInner	esLoweredTabOrder  TSpeedButtonsbReportToExcelLeft TopWidthHeightHint������������ �����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 330     337wwwww330�����?����3�� 0  � �w�wws�w7�������ws3?��3����  ��w�3wws�7࿿�����w�3?��3����  ��w�3wws�7࿿�����w�?���3���   ��w�wwws77 ������ws��333730 �����37w���?�330� �  337�w3ww330����337���7330���337wwws330��� 3337���w3330   3337wwws3	NumGlyphsParentShowHintShowHint	OnClicksbReportToExcelClick  TToolButtonToolButton1LeftTopWidthCaptionToolButton1Style
tbsDivider  TSpeedButtonExitSpeedButtonLeft$TopWidthHeightHint�����Flat	
Glyph.Data
z  v  BMv      v   (                                       �  �   �� �   � � ��   ���   �  �   �� �   � � ��  ��� 3     33wwwww33333333333333333333333333333333333333333333333?33�33333s3333333333333���33��337ww�33��337���33��337ww3333333333333����33     33wwwwws3	NumGlyphsParentShowHintShowHint	OnClickExitSpeedButtonClick    TPageControlStatisticPageControlLeft Top(Width�Height� 
ActivePagesaldoCloseTabSheetTabOrder 	TTabSheetsaldoCloseTabSheetCaption�������� ������������� TLabelLabel2LeftTopWidthHeightCaption��Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.StylefsBold 
ParentFont  	TMaskEditsaldoCloseMaskEditLeft(TopWidthIHeightEditMask!99/99/0000;1;_Font.CharsetDEFAULT_CHARSET
Font.ColorclWindowTextFont.Height�	Font.NameMS Sans Serif
Font.Style 	MaxLength

ParentFontTabOrder Text
  .  .        TQueryallCoalSenderQueryDatabaseNamemy_disLeft,Top� 	ParamData   	TDatabasedis_ibdbDatabaseDatabaseNamemy_dis_ibdb_cyrr
DriverNameINTRBASELoginPromptParams.Strings/SERVER NAME=dtkachenko:e:\dis_ibdb\dis_ibdb.gdbUSER NAME=SYSDBAOPEN MODE=READ/WRITELANGDRIVER=ancyrrSCHEMA CACHE SIZE=32SQLQRYMODE="SQLPASSTHRU MODE=SHARED AUTOCOMMITSCHEMA CACHE TIME=-1MAX ROWS=-1BATCH COUNT=200ENABLE SCHEMA CACHE=TRUEENABLE BCD=FALSEBLOBS TO CACHE=64BLOB SIZE=32PASSWORD=masterkey ReadOnly	SessionNameDefaultLeftTopH  TQueryblackListEntQueryDatabaseNamemy_dis_ibdb_cyrrSQL.Strings%select enterpr_id from balans_enterprwhere black_list = 'Y' Left\Top� 	ParamData   TQueryallOtherEntContractQueryDatabaseNamemy_disSQL.StringsGselect c.* from all_enterpr_contract_saldo_date(:ent_id, :saldo_date) cwhere exists-(select * from b_is_coal_contract(c.contract) where is_coal = 'N') Left� Top� 	ParamDataDataType	ftUnknownNameent_id	ParamType	ptUnknown DataType	ftUnknownName
saldo_date	ParamType	ptUnknown    TQueryallCoalEntContractQueryDatabaseNamemy_disSQL.StringsGselect c.* from all_enterpr_contract_saldo_date(:ent_id, :saldo_date) cwhere exists-(select * from b_is_coal_contract(c.contract) where is_coal = 'Y') Left� Top� 	ParamDataDataType	ftUnknownNameent_id	ParamType	ptUnknown DataType	ftUnknownName
saldo_date	ParamType	ptUnknown     