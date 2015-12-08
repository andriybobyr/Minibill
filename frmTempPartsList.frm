VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Begin VB.Form frmTempPartsList 
   Caption         =   "Temporary Parts List"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   487
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   491
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDown 
      Caption         =   "Down"
      Height          =   615
      Left            =   4800
      Picture         =   "frmTempPartsList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up "
      Height          =   615
      Left            =   4800
      Picture         =   "frmTempPartsList.frx":024A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.ListBox lstAvailableSequence 
      Height          =   3180
      ItemData        =   "frmTempPartsList.frx":0494
      Left            =   960
      List            =   "frmTempPartsList.frx":049B
      TabIndex        =   5
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Display Report"
      Default         =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   1185
      TabIndex        =   4
      Top             =   5760
      Width           =   1665
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Save To Excel"
      Height          =   495
      Index           =   1
      Left            =   3090
      TabIndex        =   1
      Top             =   5760
      Width           =   1635
   End
   Begin VB.CommandButton cmdExitReport 
      Cancel          =   -1  'True
      Caption         =   "&Exit Report"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4965
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
   Begin TrueDBReports60Ctl.TDBReports TDBDailyScheduleSheet 
      Height          =   570
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   1005
      Caption         =   "PartCategoryReport"
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ErrorMsgCaption =   ""
      Filtered        =   0   'False
      DataMode        =   1
      DataMember      =   ""
      LinkSequence    =   1
      LinkOrder       =   0
      NameSubstitute  =   ""
      ConnectionString=   ""
      ConnectStringType=   1
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      CursorLocation  =   2
      ConnectionTimeout=   15
      CommandTimeout  =   30
      RecordSource    =   ""
      CursorType      =   3
      CommandType     =   8
      MaxRecords      =   0
      LinkType        =   0
      Master          =   ""
      CallDataRead    =   0   'False
      ConvertNullToEmpty=   -1  'True
      DesignConnection=   -1  'True
      DesignTimeout   =   5
      UnitsOfMeasurement=   0
      Vedit_ShowGrid  =   -1  'True
      Vedit_SnapToGrid=   0   'False
      Vedit_GridUnitWidth=   8
      Vedit_GridUnitHeight=   8
      Vedit_ShowCellExpressions=   -1  'True
      Norm_rect_left  =   0
      Norm_rect_top   =   0
      Norm_rect_right =   0
      Norm_rect_bottom=   0
      Virgin          =   0   'False
      Parameters.Count=   13
      Parameters(0).Name=   "division_name"
      Parameters(1).Name=   "sort_parameters"
      Parameters(2).Name=   "rec_count"
      Parameters(2).Type=   2
      Parameters(3).Name=   "Total_Pages"
      Parameters(4).Name=   "HeaderName_0"
      Parameters(5).Name=   "HeaderName_1"
      Parameters(6).Name=   "HeaderName_2"
      Parameters(7).Name=   "HeaderName_3"
      Parameters(8).Name=   "HeaderName_4"
      Parameters(9).Name=   "HeaderName_5"
      Parameters(10).Name=   "HeaderName_6"
      Parameters(11).Name=   "HeaderName_7"
      Parameters(12).Name=   "HeaderName_8"
      Fields.Count    =   9
      Fields(0).Name  =   "FIELD_0"
      Fields(0).DisplayName=   "FIELD_0"
      Fields(1).Name  =   "FIELD_1"
      Fields(1).DisplayName=   "FIELD_1"
      Fields(2).Name  =   "FIELD_2"
      Fields(2).DisplayName=   "FIELD_2"
      Fields(3).Name  =   "FIELD_3"
      Fields(3).DisplayName=   "FIELD_3"
      Fields(4).Name  =   "FIELD_4"
      Fields(4).DisplayName=   "FIELD_4"
      Fields(5).Name  =   "FIELD_5"
      Fields(5).DisplayName=   "FIELD_5"
      Fields(6).Name  =   "FIELD_6"
      Fields(6).DisplayName=   "FIELD_6"
      Fields(7).Name  =   "FIELD_7"
      Fields(7).DisplayName=   "FIELD_7"
      Fields(8).Name  =   "FIELD_8"
      Fields(8).DisplayName=   "FIELD_8"
      Sections.Count  =   6
      Sections(0).Name=   "SECTION_3"
      Sections(0).Type=   1
      Sections(0).StyleExp=   "'tdb_RepHeader_RJ'"
      Sections(0).Cells.Count=   1
      Sections(0).Cells(0).Name=   "CELL_0"
      Sections(0).Cells(0).Exp=   "division_name & "" - Temporary Parts List"""
      Sections(1).Name=   "SECTION_4"
      Sections(1).Type=   1
      Sections(1).StyleExp=   "'tdb_GroupFooterBase'"
      Sections(1).dtopts=   2
      Sections(1).Cells.Count=   3
      Sections(1).Cells(0).Name=   "CELL_0"
      Sections(1).Cells(0).Exp=   """Sort By: ""  & sort_parameters"
      Sections(1).Cells(0).Width=   38
      Sections(1).Cells(1).Name=   "CELL_1"
      Sections(1).Cells(1).Exp=   "Space(10)"
      Sections(1).Cells(1).Width=   29
      Sections(1).Cells(2).Name=   "CELL_2"
      Sections(1).Cells(2).Exp=   """Print Date: "" & format(now,""mm/dd/yyyy hh:mm:ss"") & chr(10) & chr(13)"
      Sections(2).Name=   "SECTION_1"
      Sections(2).Condition=   "((RecNo() Mod rec_count) = 0) and RecNo() > 0"
      Sections(2).StyleExp=   "'tdb_Base'"
      Sections(2).KeepWithPrev=   1
      Sections(2).Cells.Count=   1
      Sections(2).Cells(0).Name=   "CELL_0"
      Sections(2).Cells(0).Exp=   "chr(10) & chr(13)"
      Sections(3).Name=   "SECTION_5"
      Sections(3).Type=   3
      Sections(3).Condition=   "IsTopOfPage()"
      Sections(3).StyleExp=   "'tdb_PageHeader'"
      Sections(3).Cells.Count=   9
      Sections(3).Cells(0).Name=   "CELL_0"
      Sections(3).Cells(0).Exp=   "HeaderName_0"
      Sections(3).Cells(0).WidthInPercent=   0   'False
      Sections(3).Cells(1).Name=   "CELL_1"
      Sections(3).Cells(1).Exp=   "HeaderName_1"
      Sections(3).Cells(1).WidthInPercent=   0   'False
      Sections(3).Cells(2).Name=   "CELL_2"
      Sections(3).Cells(2).Exp=   "HeaderName_2"
      Sections(3).Cells(2).WidthInPercent=   0   'False
      Sections(3).Cells(3).Name=   "CELL_3"
      Sections(3).Cells(3).Exp=   "HeaderName_3"
      Sections(3).Cells(3).WidthInPercent=   0   'False
      Sections(3).Cells(4).Name=   "CELL_4"
      Sections(3).Cells(4).Exp=   "HeaderName_4"
      Sections(3).Cells(4).WidthInPercent=   0   'False
      Sections(3).Cells(5).Name=   "CELL_5"
      Sections(3).Cells(5).Exp=   "HeaderName_5"
      Sections(3).Cells(5).WidthInPercent=   0   'False
      Sections(3).Cells(6).Name=   "CELL_6"
      Sections(3).Cells(6).Exp=   "HeaderName_6"
      Sections(3).Cells(6).WidthInPercent=   0   'False
      Sections(3).Cells(7).Name=   "CELL_7"
      Sections(3).Cells(7).Exp=   "HeaderName_7"
      Sections(3).Cells(7).WidthInPercent=   0   'False
      Sections(3).Cells(8).Name=   "CELL_8"
      Sections(3).Cells(8).Exp=   "HeaderName_8"
      Sections(3).Cells(8).WidthInPercent=   0   'False
      Sections(4).Name=   "SECTION_6"
      Sections(4).Type=   2
      Sections(4).SpacingBefore=   0.5
      Sections(4).Cells.Count=   1
      Sections(4).Cells(0).Name=   "PageNumber"
      Sections(4).Cells(0).Exp=   """Page  "" & cstr(pageno())"
      Sections(5).Name=   "SECTION_2"
      Sections(5).Type=   4
      Sections(5).StyleExp=   "'tdb_Detail_LJ'"
      Sections(5).KeepWithPrev=   1
      Sections(5).Cells.Count=   9
      Sections(5).Cells(0).Name=   "CELL_0"
      Sections(5).Cells(0).Exp=   "FIELD_0"
      Sections(5).Cells(0).WidthInPercent=   0   'False
      Sections(5).Cells(1).Name=   "CELL_1"
      Sections(5).Cells(1).Exp=   "FIELD_1"
      Sections(5).Cells(1).WidthInPercent=   0   'False
      Sections(5).Cells(2).Name=   "CELL_2"
      Sections(5).Cells(2).Exp=   "FIELD_2"
      Sections(5).Cells(2).WidthInPercent=   0   'False
      Sections(5).Cells(3).Name=   "CELL_3"
      Sections(5).Cells(3).Exp=   "FIELD_3"
      Sections(5).Cells(3).WidthInPercent=   0   'False
      Sections(5).Cells(4).Name=   "CELL_4"
      Sections(5).Cells(4).Exp=   "FIELD_4"
      Sections(5).Cells(4).WidthInPercent=   0   'False
      Sections(5).Cells(5).Name=   "CELL_5"
      Sections(5).Cells(5).Exp=   "FIELD_5"
      Sections(5).Cells(5).WidthInPercent=   0   'False
      Sections(5).Cells(6).Name=   "CELL_6"
      Sections(5).Cells(6).Exp=   "FIELD_6"
      Sections(5).Cells(6).WidthInPercent=   0   'False
      Sections(5).Cells(7).Name=   "CELL_7"
      Sections(5).Cells(7).Exp=   "FIELD_7"
      Sections(5).Cells(7).WidthInPercent=   0   'False
      Sections(5).Cells(8).Name=   "CELL_8"
      Sections(5).Cells(8).Exp=   "FIELD_8"
      Sections(5).Cells(8).WidthInPercent=   0   'False
      Styles.Count    =   32
      Styles(0).Name  =   "tdb_Base"
      Styles(0).ParentName=   ""
      Styles(0).Font_Name=   "Arial"
      Styles(0).Font_Size=   9.75
      Styles(0).Font_Charset=   0
      Styles(0).NoClipping=   -1  'True
      Styles(1).Name  =   "ShadedRow"
      Styles(1).ParentName=   "tdb_TableOddRow"
      Styles(1).Font_Name=   "Arial"
      Styles(1).Font_Size=   9.75
      Styles(1).Font_Charset=   0
      Styles(1).TextAlign=   0
      Styles(1).BackColor=   65535
      Styles(1).NoFill=   0   'False
      Styles(1).BorderHT=   "tdb_ThinBlack"
      Styles(1).BorderHI=   "tdb_ThinBlack"
      Styles(1).NoClipping=   -1  'True
      Styles(1).fprops=   98353
      Styles(2).Name  =   "tdb_TableBase"
      Styles(2).ParentName=   "tdb_Base"
      Styles(2).Font_Name=   "Arial"
      Styles(2).Font_Size=   9.75
      Styles(2).Font_Charset=   0
      Styles(2).BorderHT=   "tdb_ThinBlack"
      Styles(2).BorderHI=   "tdb_Invisible"
      Styles(2).BorderHB=   "tdb_ThinBlack"
      Styles(2).BorderVL=   "tdb_ThinBlack"
      Styles(2).BorderVI=   "tdb_ThinGray"
      Styles(2).BorderVR=   "tdb_ThinBlack"
      Styles(2).NoClipping=   -1  'True
      Styles(2).fprops=   4161536
      Styles(3).Name  =   "tdb_8ptFont"
      Styles(3).ParentName=   "tdb_TableHeader"
      Styles(3).Font_Name=   "Arial"
      Styles(3).Font_Size=   7.5
      Styles(3).Font_Charset=   0
      Styles(3).TextAlign=   1
      Styles(3).NoFill=   0   'False
      Styles(3).MarginLeft=   2
      Styles(3).MarginTop=   2
      Styles(3).MarginRight=   2
      Styles(3).MarginBottom=   2
      Styles(3).BorderHI=   "tdb_ThinBlack"
      Styles(3).BorderHB=   "tdb_ThinBlack"
      Styles(3).BorderVI=   "tdb_ThinBlack"
      Styles(3).NoClipping=   -1  'True
      Styles(3).fprops=   4945977
      Styles(4).Name  =   "tdb_TableHeader"
      Styles(4).ParentName=   "tdb_TableBase"
      Styles(4).Font_Name=   "Arial"
      Styles(4).Font_Size=   9.75
      Styles(4).Font_Charset=   0
      Styles(4).ForeColor=   8388608
      Styles(4).BackColor=   15132390
      Styles(4).NoFill=   0   'False
      Styles(4).BorderHI=   "tdb_ThinGray"
      Styles(4).NoClipping=   -1  'True
      Styles(4).fprops=   155254840
      Styles(5).Name  =   "tdb_TableOddRow"
      Styles(5).ParentName=   "tdb_TableBase"
      Styles(5).Font_Name=   "Arial"
      Styles(5).Font_Size=   9
      Styles(5).Font_Charset=   0
      Styles(5).TextAlign=   1
      Styles(5).BorderHI=   "tdb_ThinBlack"
      Styles(5).BorderHB=   "tdb_ThinBlack"
      Styles(5).BorderVI=   "tdb_ThinBlack"
      Styles(5).BorderVR=   "tdb_ThinBlack"
      Styles(5).NoClipping=   -1  'True
      Styles(5).fprops=   5963777
      Styles(6).Name  =   "tdb_TableEvenRow"
      Styles(6).ParentName=   "tdb_TableOddRow"
      Styles(6).Font_Name=   "Arial"
      Styles(6).Font_Size=   9.75
      Styles(6).Font_Charset=   0
      Styles(6).BackColor=   15132390
      Styles(6).NoFill=   0   'False
      Styles(6).NoClipping=   -1  'True
      Styles(6).fprops=   48
      Styles(7).Name  =   "tdb_TableOddAlt"
      Styles(7).ParentName=   "tdb_TableOddRow"
      Styles(7).Font_Name=   "Arial"
      Styles(7).Font_Size=   9.75
      Styles(7).Font_Charset=   0
      Styles(7).NoClipping=   -1  'True
      Styles(7).fprops=   0
      Styles(8).Name  =   "tdb_TableEvenAlt"
      Styles(8).ParentName=   "tdb_TableEvenRow"
      Styles(8).Font_Name=   "Arial"
      Styles(8).Font_Size=   9.75
      Styles(8).Font_Charset=   0
      Styles(8).NoClipping=   -1  'True
      Styles(8).fprops=   0
      Styles(9).Name  =   "tdb_TableHighlight"
      Styles(9).ParentName=   "tdb_TableOddRow"
      Styles(9).Font_Name=   "Arial"
      Styles(9).Font_Size=   9.75
      Styles(9).Font_Charset=   0
      Styles(9).BackColor=   16777088
      Styles(9).NoFill=   0   'False
      Styles(9).BorderHT=   "tdb_ThickRed"
      Styles(9).BorderHI=   "tdb_ThickRed"
      Styles(9).BorderHB=   "tdb_ThickRed"
      Styles(9).BorderVL=   "tdb_ThickRed"
      Styles(9).BorderVI=   "tdb_ThickRed"
      Styles(9).BorderVR=   "tdb_ThickRed"
      Styles(9).NoClipping=   -1  'True
      Styles(9).fprops=   2064432
      Styles(10).Name =   "tdb_TableFiller"
      Styles(10).ParentName=   "tdb_TableOddRow"
      Styles(10).Font_Name=   "Arial"
      Styles(10).Font_Size=   9.75
      Styles(10).Font_Charset=   0
      Styles(10).MarginTop=   0
      Styles(10).MarginBottom=   0
      Styles(10).NoClipping=   -1  'True
      Styles(10).fprops=   20480
      Styles(11).Name =   "tdb_TableFooter"
      Styles(11).ParentName=   "tdb_TableBase"
      Styles(11).Font_Name=   "Arial"
      Styles(11).Font_Size=   9.75
      Styles(11).Font_Charset=   0
      Styles(11).ForeColor=   8388608
      Styles(11).BackColor=   15132390
      Styles(11).NoFill=   0   'False
      Styles(11).BorderHI=   "tdb_ThinGray"
      Styles(11).NoClipping=   -1  'True
      Styles(11).fprops=   65592
      Styles(12).Name =   "tdb_Bullet"
      Styles(12).ParentName=   "tdb_Base"
      Styles(12).Font_Name=   "Arial"
      Styles(12).Font_Size=   9.75
      Styles(12).Font_Charset=   0
      Styles(12).NoClipping=   -1  'True
      Styles(12).fprops=   536871424
      Styles(13).Name =   "tdb_BulletTriangle"
      Styles(13).ParentName=   "tdb_Base"
      Styles(13).Font_Name=   "Arial"
      Styles(13).Font_Size=   9.75
      Styles(13).Font_Charset=   0
      Styles(13).NoClipping=   -1  'True
      Styles(13).fprops=   536871424
      Styles(14).Name =   "tdb_BulletHollow"
      Styles(14).ParentName=   "tdb_Base"
      Styles(14).Font_Name=   "Arial"
      Styles(14).Font_Size=   9.75
      Styles(14).Font_Charset=   0
      Styles(14).NoClipping=   -1  'True
      Styles(14).fprops=   536871424
      Styles(15).Name =   "tdb_PageHeader"
      Styles(15).ParentName=   "tdb_Base"
      Styles(15).Font_Name=   "Arial"
      Styles(15).Font_Size=   11.25
      Styles(15).Font_Bold=   -1  'True
      Styles(15).Font_Charset=   0
      Styles(15).TextAlign=   0
      Styles(15).NoClipping=   -1  'True
      Styles(15).fprops=   23068673
      Styles(16).Name =   "tdb_PageFooter"
      Styles(16).ParentName=   "tdb_PageHeader"
      Styles(16).Font_Name=   "Arial"
      Styles(16).Font_Size=   9.75
      Styles(16).Font_Charset=   0
      Styles(16).NoClipping=   -1  'True
      Styles(16).fprops=   0
      Styles(17).Name =   "tdb_RepHeader"
      Styles(17).ParentName=   "tdb_Base"
      Styles(17).Font_Name=   "Arial"
      Styles(17).Font_Size=   14.25
      Styles(17).Font_Bold=   -1  'True
      Styles(17).Font_Italic=   -1  'True
      Styles(17).Font_Charset=   0
      Styles(17).TextAlign=   1
      Styles(17).NoClipping=   -1  'True
      Styles(17).fprops=   56623105
      Styles(18).Name =   "tdb_RepHeader_RJ"
      Styles(18).ParentName=   "tdb_RepHeader"
      Styles(18).Font_Name=   "Arial"
      Styles(18).Font_Size=   9.75
      Styles(18).Font_Charset=   0
      Styles(18).TextAlign=   2
      Styles(18).ForePicFile=   "\\Tul-ares\vol1\USER\FRANCDE\CALQuality Unit Disp\WHRLOGO3.bmp"
      Styles(18).NoClipping=   -1  'True
      Styles(18).fprops=   536870913
      Styles(19).Name =   "tdb_RepFooter"
      Styles(19).ParentName=   "tdb_Base"
      Styles(19).Font_Name=   "Arial"
      Styles(19).Font_Size=   14.25
      Styles(19).Font_Bold=   -1  'True
      Styles(19).Font_Charset=   0
      Styles(19).TextAlign=   2
      Styles(19).NoClipping=   -1  'True
      Styles(19).fprops=   23068673
      Styles(20).Name =   "tdb_GroupHeaderBase"
      Styles(20).ParentName=   "tdb_Base"
      Styles(20).Font_Name=   "Arial"
      Styles(20).Font_Size=   9.75
      Styles(20).Font_Charset=   0
      Styles(20).NoClipping=   -1  'True
      Styles(20).fprops=   2097152
      Styles(21).Name =   "tdb_GroupFooterBase"
      Styles(21).ParentName=   "tdb_Base"
      Styles(21).Font_Name=   "Arial"
      Styles(21).Font_Size=   9.75
      Styles(21).Font_Charset=   0
      Styles(21).TextAlign=   2
      Styles(21).NoClipping=   -1  'True
      Styles(21).fprops=   2097153
      Styles(22).Name =   "tdb_GroupHeader1"
      Styles(22).ParentName=   "tdb_GroupHeaderBase"
      Styles(22).Font_Name=   "Arial"
      Styles(22).Font_Size=   14.25
      Styles(22).Font_Bold=   -1  'True
      Styles(22).Font_Charset=   0
      Styles(22).NoClipping=   -1  'True
      Styles(22).fprops=   20971520
      Styles(23).Name =   "tdb_GroupFooter1"
      Styles(23).ParentName=   "tdb_GroupFooterBase"
      Styles(23).Font_Name=   "Arial"
      Styles(23).Font_Size=   14.25
      Styles(23).Font_Bold=   -1  'True
      Styles(23).Font_Charset=   0
      Styles(23).NoClipping=   -1  'True
      Styles(23).fprops=   20971520
      Styles(24).Name =   "tdb_GroupHeader2"
      Styles(24).ParentName=   "tdb_GroupHeaderBase"
      Styles(24).Font_Name=   "Arial"
      Styles(24).Font_Size=   14.25
      Styles(24).Font_Charset=   0
      Styles(24).NoClipping=   -1  'True
      Styles(24).fprops=   4194304
      Styles(25).Name =   "tdb_GroupFooter2"
      Styles(25).ParentName=   "tdb_GroupFooterBase"
      Styles(25).Font_Name=   "Arial"
      Styles(25).Font_Size=   14
      Styles(25).Font_Charset=   0
      Styles(25).NoClipping=   -1  'True
      Styles(25).fprops=   4194304
      Styles(26).Name =   "tdb_GroupHeader3"
      Styles(26).ParentName=   "tdb_GroupHeaderBase"
      Styles(26).Font_Name=   "Arial"
      Styles(26).Font_Size=   12
      Styles(26).Font_Bold=   -1  'True
      Styles(26).Font_Charset=   0
      Styles(26).NoClipping=   -1  'True
      Styles(26).fprops=   20971520
      Styles(27).Name =   "tdb_GroupFooter3"
      Styles(27).ParentName=   "tdb_GroupFooterBase"
      Styles(27).Font_Name=   "Arial"
      Styles(27).Font_Size=   12
      Styles(27).Font_Bold=   -1  'True
      Styles(27).Font_Charset=   0
      Styles(27).NoClipping=   -1  'True
      Styles(27).fprops=   20971520
      Styles(28).Name =   "tdb_GroupHeader4"
      Styles(28).ParentName=   "tdb_GroupHeaderBase"
      Styles(28).Font_Name=   "Arial"
      Styles(28).Font_Size=   12
      Styles(28).Font_Charset=   0
      Styles(28).NoClipping=   -1  'True
      Styles(28).fprops=   4194304
      Styles(29).Name =   "tdb_GroupFooter4"
      Styles(29).ParentName=   "tdb_GroupFooterBase"
      Styles(29).Font_Name=   "Arial"
      Styles(29).Font_Size=   12
      Styles(29).Font_Charset=   0
      Styles(29).NoClipping=   -1  'True
      Styles(29).fprops=   4194304
      Styles(30).Name =   "tdb_Detail"
      Styles(30).ParentName=   "tdb_Base"
      Styles(30).Font_Name=   "Arial"
      Styles(30).Font_Size=   9
      Styles(30).Font_Charset=   0
      Styles(30).TextAlign=   1
      Styles(30).MarginLeft=   2
      Styles(30).MarginTop=   2
      Styles(30).MarginRight=   2
      Styles(30).MarginBottom=   2
      Styles(30).BorderHT=   "tdb_ThinBlack"
      Styles(30).BorderHI=   "tdb_ThinBlack"
      Styles(30).BorderHB=   "tdb_ThinBlack"
      Styles(30).BorderVL=   "tdb_ThinBlack"
      Styles(30).BorderVI=   "tdb_ThinBlack"
      Styles(30).BorderVR=   "tdb_ThinBlack"
      Styles(30).NoClipping=   -1  'True
      Styles(30).fprops=   8386561
      Styles(31).Name =   "tdb_Detail_LJ"
      Styles(31).ParentName=   "tdb_Detail"
      Styles(31).Font_Name=   "Arial"
      Styles(31).Font_Size=   9.75
      Styles(31).Font_Charset=   0
      Styles(31).TextAlign=   0
      Styles(31).NoFill=   0   'False
      Styles(31).NoClipping=   -1  'True
      Styles(31).fprops=   33
      Mappings.Count  =   5
      Mappings(0).Name=   "tdb_CheckboxV"
      Mappings(0).ValueItems.Count=   4
      Mappings(0).ValueItems(0).Key=   "False"
      Mappings(0).ValueItems(1).Key=   "True"
      Mappings(0).ValueItems(1).Default=   -1  'True
      Mappings(0).ValueItems(2).Key=   ""
      Mappings(0).ValueItems(2).LinkedKey=   "False"
      Mappings(0).ValueItems(3).Key=   "0"
      Mappings(0).ValueItems(3).LinkedKey=   "False"
      Mappings(1).Name=   "tdb_CheckboxVBoxed"
      Mappings(1).ValueItems.Count=   4
      Mappings(1).ValueItems(0).Key=   "False"
      Mappings(1).ValueItems(1).Key=   "True"
      Mappings(1).ValueItems(1).Default=   -1  'True
      Mappings(1).ValueItems(2).Key=   ""
      Mappings(1).ValueItems(2).LinkedKey=   "False"
      Mappings(1).ValueItems(3).Key=   "0"
      Mappings(1).ValueItems(3).LinkedKey=   "False"
      Mappings(2).Name=   "tdb_CheckboxX"
      Mappings(2).ValueItems.Count=   4
      Mappings(2).ValueItems(0).Key=   "False"
      Mappings(2).ValueItems(1).Key=   "True"
      Mappings(2).ValueItems(1).Default=   -1  'True
      Mappings(2).ValueItems(2).Key=   ""
      Mappings(2).ValueItems(2).LinkedKey=   "False"
      Mappings(2).ValueItems(3).Key=   "0"
      Mappings(2).ValueItems(3).LinkedKey=   "False"
      Mappings(3).Name=   "tdb_CheckboxXBoxed"
      Mappings(3).ValueItems.Count=   4
      Mappings(3).ValueItems(0).Key=   "False"
      Mappings(3).ValueItems(1).Key=   "True"
      Mappings(3).ValueItems(1).Default=   -1  'True
      Mappings(3).ValueItems(2).Key=   ""
      Mappings(3).ValueItems(2).LinkedKey=   "False"
      Mappings(3).ValueItems(3).Key=   "0"
      Mappings(3).ValueItems(3).LinkedKey=   "False"
      Mappings(4).Name=   "tdb_CheckboxCircle"
      Mappings(4).ValueItems.Count=   4
      Mappings(4).ValueItems(0).Key=   "False"
      Mappings(4).ValueItems(1).Key=   "True"
      Mappings(4).ValueItems(1).Default=   -1  'True
      Mappings(4).ValueItems(2).Key=   ""
      Mappings(4).ValueItems(2).LinkedKey=   "False"
      Mappings(4).ValueItems(3).Key=   "0"
      Mappings(4).ValueItems(3).LinkedKey=   "False"
      Lines.Count     =   14
      Lines(0).Name   =   "tdb_Invisible"
      Lines(1).Name   =   "tdb_ThinBlack"
      Lines(1).Thickness=   2
      Lines(2).Name   =   "tdb_MediumBlack"
      Lines(2).Thickness=   5
      Lines(3).Name   =   "tdb_ThickBlack"
      Lines(3).Thickness=   7
      Lines(4).Name   =   "tdb_ThinGray"
      Lines(4).Thickness=   2
      Lines(4).Color  =   8421504
      Lines(5).Name   =   "tdb_MediumGray"
      Lines(5).Thickness=   5
      Lines(5).Color  =   8421504
      Lines(6).Name   =   "tdb_ThickGray"
      Lines(6).Thickness=   7
      Lines(6).Color  =   8421504
      Lines(7).Name   =   "tdb_ThinRed"
      Lines(7).Thickness=   2
      Lines(7).Color  =   255
      Lines(8).Name   =   "tdb_MediumRed"
      Lines(8).Thickness=   5
      Lines(8).Color  =   255
      Lines(9).Name   =   "tdb_ThickRed"
      Lines(9).Thickness=   7
      Lines(9).Color  =   255
      Lines(10).Name  =   "tdb_ThinOrange"
      Lines(10).Thickness=   2
      Lines(10).Color =   33023
      Lines(11).Name  =   "tdb_MediumWhite"
      Lines(11).Thickness=   5
      Lines(11).Color =   16777215
      Lines(12).Name  =   "tdb_ThinBlue"
      Lines(12).Thickness=   2
      Lines(12).Color =   8404992
      Lines(13).Name  =   "tdb_MediumBlue"
      Lines(13).Thickness=   5
      Lines(13).Color =   8404992
      Profiles.Count  =   1
      Profiles(0).Name=   "PROFILE_0"
      Profiles(0).Active=   -1  'True
      Profiles(0).Collate=   -1  'True
      Profiles(0).PreviewMaximized=   -1  'True
      Profiles(0).PreviewInitialZoom=   75
      Profiles(0).PrinterMarginLeft=   20
      Profiles(0).PrinterMarginTop=   10
      Profiles(0).PrinterMarginRight=   22
      Profiles(0).PrinterMarginBottom=   20
      Profiles(0).PrinterMargins_set=   -1  'True
      Profiles(0).PrinterPaperSize_set=   -1  'True
      Profiles(0).PrinterPaperUserSize_set=   -1  'True
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"frmTempPartsList.frx":04B5
      Height          =   975
      Left            =   840
      TabIndex        =   8
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Temporary Part Listing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmTempPartsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private marrstrLine() As String
Private marrstrLocationId() As String
Private mrsTempParts As New ADODB.Recordset
Private mxarrReportData As XArrayDB         ' Report array
Private mxarrTemp As XArrayDB               ' Temporary array
Private mxarrModelNotes As XArrayDB         ' Array containing model/notes
Private mintMaxCategories As Integer        ' Maximum number of categories allowed
Private mintMaxModels As Integer            ' Maximum number of sort fields
Private mintMaxCols As Integer              ' Maximum number of sort fields
Private mintModelCount As Integer           ' Counter for current number of models
Private mintRows As Integer                 ' Counter for current number of rows
Private mintCols As Integer                 ' Counter for current number of columns
Private mintModelIndex As Integer           ' Current model index
Private mintSeqNumber As Integer            ' Current sequence number value
Private mintLocationIndex As String
Public mintNumberOfCopies As Integer        ' Number of print copies
Public mblnCancelPrint As Boolean           ' Cancel Print flag
Public strCategorySequence As String        ' Sequence of Category in Stocking Location


Public strModelLength As String             'Length of model number in strModel
Public strModelWOQuantity As String         'String of Model without the Quantity

Public strOrderBy As String                 'Order based on screen inputs

Public varDisplay As Variant

Private mlngCurrentRow As Long
Private mobjXL As Excel.Application

Private strDisplay As String

Private Function PrintReport(strOrderBy As String, blnPrintPreview As Boolean, index As Integer) As Integer
    ' Purpose:  Retrieve data and put it into an array used for the report object's data
    ' source.  Report will reflect sorts made in selection screen.
          
                        
    ' Increase the timeout for the query so it can complete
    gconDatabase.CommandTimeout = 400
    
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strDate As String

    PrintReport = 0
    intRow = 0
     
         
    ' Execute the query to create the recordset
    Set mrsTempParts = New ADODB.Recordset
    With mrsTempParts
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "SELECT " & strOrderBy & ", Quantity FROM V_MNB_Temp_Model_Line_Part_Mast_Detail pm " & _
            "JOIN v_prod_stocking_location sl ON pm.stocking_location_id = sl.stocking_location_id " & _
            "ORDER BY " & strOrderBy
                 
        .Open
            
        mintRows = .RecordCount
        
        If .RecordCount = 0 Then
            MsgBox "No data found for report"
            GoTo PROC_EXIT
        End If
               
        Set mxarrTemp = New XArrayDB
        mxarrTemp.ReDim 0, mintRows, 0, mintMaxCols
        
        intRow = 0
        
        .MoveFirst
        Do While Not .EOF
           
          With mxarrTemp
            
            For intCol = 0 To mintMaxCols
                If Trim(lstAvailableSequence.List(intCol)) = "Temp_Part_Activity_Code" Then
                    If mrsTempParts!temp_part_activity_code = "A" Then
                        mxarrTemp(intRow, intCol) = "Add Part"
                    End If
                    If mrsTempParts!temp_part_activity_code = "I" Then
                        mxarrTemp(intRow, intCol) = "Inactive Part"
                    End If
                    If mrsTempParts!temp_part_activity_code = "R" Then
                        mxarrTemp(intRow, intCol) = "Replace Part"
                    End If
                End If
            
            
                'Quantity is not one of the sorted fields.  By default, it will always be the last field
                'on the report.
                If Trim(lstAvailableSequence.List(intCol)) = "" Then
                    mxarrTemp(intRow, intCol) = Trim(mrsTempParts.Fields(intCol))
                End If
            
                If Trim(lstAvailableSequence.List(intCol)) = "Quit_ECN_Number" Then
                    mxarrTemp(intRow, intCol) = Trim(mrsTempParts.Fields(intCol))
                End If
            
                If Trim(lstAvailableSequence.List(intCol)) = "Part_Id_Replaces" Then
                    mxarrTemp(intRow, intCol) = Trim(mrsTempParts.Fields(intCol))
                End If
            
                If Trim(lstAvailableSequence.List(intCol)) = "Line_Id" Then
                    mxarrTemp(intRow, intCol) = Trim(mrsTempParts.Fields(intCol))
                End If
            
                If Trim(lstAvailableSequence.List(intCol)) = "Part_Id" Then
                    mxarrTemp(intRow, intCol) = Trim(mrsTempParts.Fields(intCol))
                End If
            
                If Trim(lstAvailableSequence.List(intCol)) = "Model_Number" Then
                    mxarrTemp(intRow, intCol) = Trim(mrsTempParts.Fields(intCol))
                End If
                        
                If Trim(lstAvailableSequence.List(intCol)) = "Quit_ECN_Date" Then
                    If Len(Trim(mrsTempParts.Fields(intCol))) > 0 Then
                        strDate = Mid(mrsTempParts.Fields(intCol), 5, 2) & "/" & _
                            Mid(mrsTempParts.Fields(intCol), 7, 2) & "/" & Mid(mrsTempParts.Fields(intCol), 1, 4)
                            mxarrTemp(intRow, intCol) = strDate
                    Else
                        strDate = ""
                    End If
                End If
                       
                If Trim(lstAvailableSequence.List(intCol)) = "Stocking_Location_Description" Then
                    mxarrTemp(intRow, intCol) = Trim(mrsTempParts.Fields(intCol))
                End If
            
            Next
            
          End With
                   
          intRow = intRow + 1
            
            .MoveNext
        Loop
    End With
    
                
    Dim intX As Integer
    Dim intY As Integer

    Dim intSections As Integer
    Dim intS As Integer
    Dim intModelLower As Integer
    Dim intModelUpper As Integer
    Dim intComponentLower As Integer
    Dim intComponentUpper As Integer
    Dim intReportModelCtr As Integer
    Dim intModelCtrOnLastSection As Integer
    
    With TDBDailyScheduleSheet
        
        If index = 1 Then
            ' Dump data into Excel
            Dim xlWorkBook As Excel.Workbook

            Set mobjXL = New Excel.Application
            Set xlWorkBook = mobjXL.Workbooks.Add

            mobjXL.Rows(mxarrTemp.UpperBound(1)).Insert
                        
            intX = 0
            
            For intCol = 0 To mxarrTemp.UpperBound(2) Step 1
                intX = intX + 1
                If (lstAvailableSequence.List(intCol)) = "" Then
                    mobjXL.Cells(1, intX) = "Quantity"
                Else
                    mobjXL.Cells(1, intX) = (lstAvailableSequence.List(intCol))
                End If
            Next
            
            intX = 1
            intY = 1
            
            For intRow = 0 To mxarrTemp.UpperBound(1)
                intY = intY + 1
                intX = 0
                For intCol = 0 To mxarrTemp.UpperBound(2) Step 1
                    intX = intX + 1
                    mobjXL.Cells(intY, intX) = mxarrTemp(intRow, intCol)
                Next
            Next
                
            mobjXL.Rows(1).Insert
            mobjXL.Rows(1).Insert

            mobjXL.Cells(1, 1) = "Sort By:"
            mobjXL.Cells(1, 2) = strOrderBy
            mobjXL.Cells(1, 5) = "Printed Date:"
            mobjXL.Cells(1, 6) = Now()
            
            mobjXL.Range("A4:A4").Columns.AutoFit
            mobjXL.Range("B3:Z3").Columns.AutoFit

            mobjXL.Visible = True
            Set xlWorkBook = Nothing
            Set mobjXL = Nothing
        Else
            ' Print the report
            If blnPrintPreview Then
                .PrintPreview
            Else
                .Profiles.Item(0).NumberOfCopies = mintNumberOfCopies
                .PrintData
            End If
        End If
    End With
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    Call ShowError(Me.Name, "PrintReport", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Function

Sub FillSequenceSelections()

    ' Purpose:   Read the database and load the line and date combo boxes.
    '    The descriptions for the line and locations are loaded to the
    'dropdowns, while the codes are loaded to an array at the same time.
    'The codes are the key fields to gather sql from the data bases.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
      
    Dim intIndex As Integer
    Dim arrFields(0 To 7) As String
    
    arrFields(0) = "Model_Number"
    arrFields(1) = "Line_Id"
    arrFields(2) = "Part_Id"
    arrFields(3) = "Temp_Part_Activity_Code"
    arrFields(4) = "Stocking_Location_Description"
    arrFields(5) = "Part_Id_Replaces"
    arrFields(6) = "Quit_ECN_Number"
    arrFields(7) = "Quit_ECN_Date"
   
           
    lstAvailableSequence.Clear

    
    For intIndex = 0 To 7
        lstAvailableSequence.AddItem arrFields(intIndex)
    Next intIndex
    
    mintMaxCols = intIndex
    
   
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FillComboBoxes", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdExitReport_Click()
    Unload Me
End Sub





Private Sub cmdReport_Click(index As Integer)
    ' Print report for selected locations by calling PrintReport function
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
   
    Dim intIndex As Integer
    Dim strLocation As String
    Dim strLine As String
    Dim strToDate As String
    Dim intCount As Integer
    Dim intTotal As Integer
    
    ' Set the variable values from the screen to the query.
    
    intIndex = 0

    
        strOrderBy = ""
        For intIndex = 0 To mintMaxCols - 1
            strOrderBy = strOrderBy & Trim(lstAvailableSequence.List(intIndex))
            If intIndex < mintMaxCols - 1 Then
                strOrderBy = strOrderBy & ", "
            End If
        Next intIndex
        
        ' Prompt user to input Number of Copies to Print
        If index = 0 Then
'            frmCopies.Show vbModal, Me
            If mblnCancelPrint Then Exit Sub
        End If
            
        ' Print report for each location directly to the printer
        Dim intI As Integer
        intTotal = 0
'        For intI = 1 To UBound(marrstrLocationId)
'            strLocation = marrstrLocationId(intI)
'            mintLocationIndex = intI
            intCount = PrintReport(strOrderBy, True, index)
'            intTotal = intTotal + intCount
'        Next
'        If intTotal = 0 Then
'            MsgBox "No data found for report"
'        End If
'    Else
        'Create report for selected location, in printpreview mode
'        mintLocationIndex = cboLocation.ListIndex
'        intCount = PrintReport(strLine, strLocation, strToDate, True, index)
'        If intCount = 0 Then
'            MsgBox "No data found for report"
'        End If
'    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdReport_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdUp_Click()
    If lstAvailableSequence.ListIndex = -1 Then
        MsgBox "No Category was selected."
        Exit Sub
    ElseIf lstAvailableSequence.ListIndex = 0 Then
        Beep
        Exit Sub
    End If
    
    Dim strSaveLine As String
    strSaveLine = lstAvailableSequence.List(lstAvailableSequence.ListIndex - 1)
    lstAvailableSequence.List(lstAvailableSequence.ListIndex - 1) = lstAvailableSequence.List(lstAvailableSequence.ListIndex)
    lstAvailableSequence.List(lstAvailableSequence.ListIndex) = strSaveLine
 '   mblnRecChanged = True
    lstAvailableSequence.ListIndex = lstAvailableSequence.ListIndex - 1
End Sub

Private Sub cmdDown_Click()
    If lstAvailableSequence.ListIndex = -1 Then
        MsgBox "No Category was selected."
        Exit Sub
    ElseIf lstAvailableSequence.ListIndex = lstAvailableSequence.ListCount - 1 Then
        Beep
        Exit Sub
    End If
    
    Dim strSaveLine As String
    strSaveLine = lstAvailableSequence.List(lstAvailableSequence.ListIndex + 1)
    lstAvailableSequence.List(lstAvailableSequence.ListIndex + 1) = lstAvailableSequence.List(lstAvailableSequence.ListIndex)
    lstAvailableSequence.List(lstAvailableSequence.ListIndex) = strSaveLine
'    mblnRecChanged = True
    lstAvailableSequence.ListIndex = lstAvailableSequence.ListIndex + 1
End Sub

Private Sub Form_Load()
    ' Purpose:  Load the form
        
'This application will print a report showing Temporary Parts active at that site.
'  One of the issues with this report was determining a sort that would make any division
'  satisfied.  One set sorting sequence did not make sense, so made a decision to allow
'  all fields in this query to be sorted on.  (Quantity will always show at the end because
'  it is NOT sorted.)
'If a new field is added to this report, please take care to add the appropriate fields and
'  parameters for it to show up throughout the code.  The report is flexible to the Business, but does cause the
'  program to have some constraints.
        
        
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    If gconDatabase Is Nothing Then
        Set gconDatabase = gclsSQLServer.Connect( _
            gclsMESApplication.ApplicationRole, _
            gclsMESApplication.ApplicationPassword)
    
        If gconDatabase.State <> adStateOpen Then
            Err.Raise gconDatabase.Errors(0).NativeError, _
                "Form_Load", gconDatabase.Errors(0).Description
        End If
    End If
    
    Call FillSequenceSelections
       
    mintNumberOfCopies = 1
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gconDatabase.CommandTimeout = 30
End Sub

Private Sub TDBDailyScheduleSheet_OpenData()
'  The List from the main screen is the sort order for the report.  This subroutine is reading
'  that list and creating the headings for the report from it.  Date and division parameters
'  are loaded here also.

    On Error GoTo PROC_ERR

    Dim intRow As Integer
    Dim intSavePartIndex As Integer
    
    'Looping through the screen List of sort variables
    With TDBDailyScheduleSheet
            
        For intRow = 0 To mintMaxCols
            If Trim(lstAvailableSequence.List(intRow)) = "Temp_Part_Activity_Code" Then
                .Parameters("HeaderName_" & intRow) = "Action Code"
            End If
            
            If Trim(lstAvailableSequence.List(intRow)) = "" Then
                .Parameters("HeaderName_" & intRow) = "Quantity"
            End If
            
            If Trim(lstAvailableSequence.List(intRow)) = "Quit_ECN_Number" Then
                .Parameters("HeaderName_" & intRow) = "Quit ECN Number"
            End If
            
            If Trim(lstAvailableSequence.List(intRow)) = "Part_Id_Replaces" Then
                .Parameters("HeaderName_" & intRow) = "Replaced Part"
            End If
            
            If Trim(lstAvailableSequence.List(intRow)) = "Line_Id" Then
                .Parameters("HeaderName_" & intRow) = "Line"
            End If
            
            If Trim(lstAvailableSequence.List(intRow)) = "Part_Id" Then
                .Parameters("HeaderName_" & intRow) = "Part Number"
            End If
            
            If Trim(lstAvailableSequence.List(intRow)) = "Model_Number" Then
                .Parameters("HeaderName_" & intRow) = "Model Number"
            End If
            
            If Trim(lstAvailableSequence.List(intRow)) = "Quit_ECN_Date" Then
                .Parameters("HeaderName_" & intRow) = "Quit ECN Date"
            End If
                       
            If Trim(lstAvailableSequence.List(intRow)) = "Stocking_Location_Description" Then
                .Parameters("HeaderName_" & intRow) = "Stocking Location"
            End If
 
        Next

        .Parameters("sort_parameters") = strOrderBy
        
        .Parameters("rec_count") = mintRows
                
        .Parameters("division_name") = gclsMESApplication.Division
        
        Set .Array = mxarrTemp
    
    End With
            
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "TDBDailyScheduleSheet_OpenData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub TDBDailyScheduleSheet_WillOpenData()
Dim intColCtr As Integer
intColCtr = 0
'  One width is 1/72 of an inch.  Increase width in increments of 12.
    With TDBDailyScheduleSheet
        For intColCtr = 0 To mintMaxCols
            If Trim(lstAvailableSequence.List(intColCtr)) = "Line_Id" Then
                .Sections.Item("Section_5").Cells.Item("Cell_" & intColCtr).Width = 36
                .Sections.Item("Section_2").Cells.Item("Cell_" & intColCtr).Width = 36
            End If
            
            If Trim(lstAvailableSequence.List(intColCtr)) = "Model_Number" Then
                .Sections.Item("Section_5").Cells.Item("Cell_" & intColCtr).Width = 96
                .Sections.Item("Section_2").Cells.Item("Cell_" & intColCtr).Width = 96
            End If
        Next
    End With
    
End Sub
