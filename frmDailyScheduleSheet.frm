VERSION 5.00
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Begin VB.Form frmDailyScheduleSheet 
   Caption         =   "Daily Schedule/Process Sheet"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
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
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Display Report"
      Default         =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   1185
      TabIndex        =   11
      Top             =   3840
      Width           =   1665
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Save To Excel"
      Height          =   495
      Index           =   1
      Left            =   3090
      TabIndex        =   3
      Top             =   3840
      Width           =   1635
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selection Criteria"
      Height          =   2835
      Left            =   435
      TabIndex        =   1
      Top             =   750
      Width           =   6675
      Begin VB.ComboBox cboLine 
         Height          =   360
         Left            =   2655
         TabIndex        =   7
         Top             =   585
         Width           =   2595
      End
      Begin VB.ComboBox cboLocation 
         Height          =   360
         Left            =   2655
         TabIndex        =   6
         Top             =   1260
         Width           =   2775
      End
      Begin VB.ComboBox cboDate 
         Height          =   360
         ItemData        =   "frmDailyScheduleSheet.frx":0000
         Left            =   2655
         List            =   "frmDailyScheduleSheet.frx":0002
         TabIndex        =   5
         Top             =   1935
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Line:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   1965
         TabIndex        =   2
         Top             =   645
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   1515
         TabIndex        =   9
         Top             =   1335
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1905
         TabIndex        =   8
         Top             =   1995
         Width           =   570
      End
   End
   Begin VB.CommandButton cmdExitReport 
      Cancel          =   -1  'True
      Caption         =   "&Exit Report"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4965
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin TrueDBReports60Ctl.TDBReports TDBDailyScheduleSheet 
      Height          =   570
      Left            =   0
      TabIndex        =   10
      Top             =   0
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
      Parameters.Count=   5
      Parameters(0).Name=   "line_id"
      Parameters(1).Name=   "division_name"
      Parameters(2).Name=   "location_id"
      Parameters(3).Name=   "schedule_date"
      Parameters(3).Type=   7
      Parameters(4).Name=   "rec_count"
      Parameters(4).Type=   2
      Fields.Count    =   17
      Fields(0).Name  =   "category"
      Fields(0).DisplayName=   "category"
      Fields(0).MaxLength=   30
      Fields(1).Name  =   "model_1"
      Fields(1).DisplayName=   "model_1"
      Fields(1).MaxLength=   20
      Fields(2).Name  =   "part_1"
      Fields(2).DisplayName=   "part_1"
      Fields(2).MaxLength=   20
      Fields(3).Name  =   "model_2"
      Fields(3).DisplayName=   "model_2"
      Fields(3).MaxLength=   20
      Fields(4).Name  =   "part_2"
      Fields(4).DisplayName=   "part_2"
      Fields(4).MaxLength=   20
      Fields(5).Name  =   "model_3"
      Fields(5).DisplayName=   "model_3"
      Fields(5).MaxLength=   20
      Fields(6).Name  =   "part_3"
      Fields(6).DisplayName=   "part_3"
      Fields(6).MaxLength=   20
      Fields(7).Name  =   "model_4"
      Fields(7).DisplayName=   "model_4"
      Fields(7).MaxLength=   20
      Fields(8).Name  =   "part_4"
      Fields(8).DisplayName=   "part_4"
      Fields(8).MaxLength=   20
      Fields(9).Name  =   "model_5"
      Fields(9).DisplayName=   "model_5"
      Fields(9).MaxLength=   20
      Fields(10).Name =   "part_5"
      Fields(10).DisplayName=   "part_5"
      Fields(10).MaxLength=   20
      Fields(11).Name =   "model_6"
      Fields(11).DisplayName=   "model_6"
      Fields(11).MaxLength=   20
      Fields(12).Name =   "part_6"
      Fields(12).DisplayName=   "part_6"
      Fields(12).MaxLength=   20
      Fields(13).Name =   "model_7"
      Fields(13).DisplayName=   "model_7"
      Fields(13).MaxLength=   20
      Fields(14).Name =   "part_7"
      Fields(14).DisplayName=   "part_7"
      Fields(14).MaxLength=   20
      Fields(15).Name =   "model_8"
      Fields(15).DisplayName=   "model_8"
      Fields(15).MaxLength=   20
      Fields(16).Name =   "part_8"
      Fields(16).DisplayName=   "part_8"
      Fields(16).MaxLength=   20
      Sections.Count  =   5
      Sections(0).Name=   "SECTION_3"
      Sections(0).Type=   1
      Sections(0).StyleExp=   "'tdb_RepHeader_RJ'"
      Sections(0).Cells.Count=   1
      Sections(0).Cells(0).Name=   "CELL_0"
      Sections(0).Cells(0).Exp=   "division_name & chr(10) & chr(13) & ""Daily Schedule/Process Sheet"""
      Sections(1).Name=   "SECTION_4"
      Sections(1).Type=   1
      Sections(1).StyleExp=   "'tdb_PageHeader'"
      Sections(1).Cells.Count=   3
      Sections(1).Cells(0).Name=   "CELL_0"
      Sections(1).Cells(0).Exp=   """Line: "" & trim(line_id)"
      Sections(1).Cells(1).Name=   "CELL_1"
      Sections(1).Cells(1).Exp=   """Location: "" & trim(location_id)"
      Sections(1).Cells(1).PrivateStyle=   -1  'True
      Sections(1).Cells(1).Style.Name=   "<private>"
      Sections(1).Cells(1).Style.ParentName=   "tdb_PageHeader"
      Sections(1).Cells(1).Style.Font_Name=   "Arial"
      Sections(1).Cells(1).Style.Font_Size=   11.25
      Sections(1).Cells(1).Style.Font_Bold=   -1  'True
      Sections(1).Cells(1).Style.Font_Italic=   0   'False
      Sections(1).Cells(1).Style.Font_Underline=   0   'False
      Sections(1).Cells(1).Style.Font_Strikeout=   0   'False
      Sections(1).Cells(1).Style.Font_Charset=   0
      Sections(1).Cells(1).Style.TextAlign=   1
      Sections(1).Cells(1).Style.TextVAlign=   0
      Sections(1).Cells(1).Style.TextWrap=   -1  'True
      Sections(1).Cells(1).Style.ForeColor=   0
      Sections(1).Cells(1).Style.BackColor=   16777215
      Sections(1).Cells(1).Style.NoFill=   -1  'True
      Sections(1).Cells(1).Style.BackPicFile=   ""
      Sections(1).Cells(1).Style.ForePicFile=   ""
      Sections(1).Cells(1).Style.BackPicVertPlacement=   0
      Sections(1).Cells(1).Style.BackPicHorzPlacement=   0
      Sections(1).Cells(1).Style.ForePicPlacement=   0
      Sections(1).Cells(1).Style.ForePicDrawMode=   0
      Sections(1).Cells(1).Style.MarginLeft=   6
      Sections(1).Cells(1).Style.MarginTop=   6
      Sections(1).Cells(1).Style.MarginRight=   6
      Sections(1).Cells(1).Style.MarginBottom=   6
      Sections(1).Cells(1).Style.HasBorders=   -1  'True
      Sections(1).Cells(1).Style.BorderHT=   ""
      Sections(1).Cells(1).Style.BorderHI=   ""
      Sections(1).Cells(1).Style.BorderHB=   ""
      Sections(1).Cells(1).Style.BorderVL=   ""
      Sections(1).Cells(1).Style.BorderVI=   ""
      Sections(1).Cells(1).Style.BorderVR=   ""
      Sections(1).Cells(1).Style.NoClipping=   -1  'True
      Sections(1).Cells(1).Style.RTF=   0   'False
      Sections(1).Cells(1).Style.fprops=   1
      Sections(1).Cells(2).Name=   "CELL_2"
      Sections(1).Cells(2).Exp=   """Date: "" & format(schedule_date,""mm/dd/yyyy"")"
      Sections(1).Cells(2).PrivateStyle=   -1  'True
      Sections(1).Cells(2).Style.Name=   "<private>"
      Sections(1).Cells(2).Style.ParentName=   "tdb_PageHeader"
      Sections(1).Cells(2).Style.Font_Name=   "Arial"
      Sections(1).Cells(2).Style.Font_Size=   11.25
      Sections(1).Cells(2).Style.Font_Bold=   -1  'True
      Sections(1).Cells(2).Style.Font_Italic=   0   'False
      Sections(1).Cells(2).Style.Font_Underline=   0   'False
      Sections(1).Cells(2).Style.Font_Strikeout=   0   'False
      Sections(1).Cells(2).Style.Font_Charset=   0
      Sections(1).Cells(2).Style.TextAlign=   2
      Sections(1).Cells(2).Style.TextVAlign=   0
      Sections(1).Cells(2).Style.TextWrap=   -1  'True
      Sections(1).Cells(2).Style.ForeColor=   0
      Sections(1).Cells(2).Style.BackColor=   16777215
      Sections(1).Cells(2).Style.NoFill=   -1  'True
      Sections(1).Cells(2).Style.BackPicFile=   ""
      Sections(1).Cells(2).Style.ForePicFile=   ""
      Sections(1).Cells(2).Style.BackPicVertPlacement=   0
      Sections(1).Cells(2).Style.BackPicHorzPlacement=   0
      Sections(1).Cells(2).Style.ForePicPlacement=   0
      Sections(1).Cells(2).Style.ForePicDrawMode=   0
      Sections(1).Cells(2).Style.MarginLeft=   6
      Sections(1).Cells(2).Style.MarginTop=   6
      Sections(1).Cells(2).Style.MarginRight=   6
      Sections(1).Cells(2).Style.MarginBottom=   6
      Sections(1).Cells(2).Style.HasBorders=   -1  'True
      Sections(1).Cells(2).Style.BorderHT=   ""
      Sections(1).Cells(2).Style.BorderHI=   ""
      Sections(1).Cells(2).Style.BorderHB=   ""
      Sections(1).Cells(2).Style.BorderVL=   ""
      Sections(1).Cells(2).Style.BorderVI=   ""
      Sections(1).Cells(2).Style.BorderVR=   ""
      Sections(1).Cells(2).Style.NoClipping=   -1  'True
      Sections(1).Cells(2).Style.RTF=   0   'False
      Sections(1).Cells(2).Style.fprops=   1
      Sections(2).Name=   "SECTION_1"
      Sections(2).Condition=   "((RecNo() Mod rec_count) = 0) and RecNo() > 0"
      Sections(2).StyleExp=   "'tdb_Base'"
      Sections(2).KeepWithPrev=   1
      Sections(2).Cells.Count=   1
      Sections(2).Cells(0).Name=   "CELL_0"
      Sections(2).Cells(0).Exp=   "chr(10) & chr(13)"
      Sections(3).Name=   "SECTION_5"
      Sections(3).Type=   3
      Sections(3).Condition=   "(RecNo() Mod rec_count) = 0"
      Sections(3).Cells.Count=   9
      Sections(3).Cells(0).Name=   "CELL_0"
      Sections(3).Cells(0).Exp=   """Category"""
      Sections(3).Cells(0).StyleExp=   "IIF(category = """", ""tdb_Base"", ""tdb_Detail"")"
      Sections(3).Cells(1).Name=   "CELL_1"
      Sections(3).Cells(1).Exp=   "model_1"
      Sections(3).Cells(1).StyleExp=   "IIF(instr(1,model_1,""("",1) > 13,  ""tdb_8ptFont"",IIF(instr(1,model_1,""("",1) = 0,""tdb_Base"",""tdb_Detail""))"
      Sections(3).Cells(2).Name=   "CELL_2"
      Sections(3).Cells(2).Exp=   "model_2"
      Sections(3).Cells(2).StyleExp=   "IIF(instr(1,model_2,""("",1) > 13,  ""tdb_8ptFont"",IIF(instr(1,model_2,""("",1) = 0,""tdb_Base"",""tdb_Detail""))"
      Sections(3).Cells(3).Name=   "CELL_3"
      Sections(3).Cells(3).Exp=   "model_3"
      Sections(3).Cells(3).StyleExp=   "IIF(instr(1,model_3,""("",1) > 13,  ""tdb_8ptFont"",IIF(instr(1,model_3,""("",1) = 0,""tdb_Base"",""tdb_Detail""))"
      Sections(3).Cells(4).Name=   "CELL_4"
      Sections(3).Cells(4).Exp=   "model_4"
      Sections(3).Cells(4).StyleExp=   "IIF(instr(1,model_4,""("",1) > 13,  ""tdb_8ptFont"",IIF(instr(1,model_4,""("",1) = 0,""tdb_Base"",""tdb_Detail""))"
      Sections(3).Cells(5).Name=   "CELL_5"
      Sections(3).Cells(5).Exp=   "model_5"
      Sections(3).Cells(5).StyleExp=   "IIF(instr(1,model_5,""("",1) > 13,  ""tdb_8ptFont"",IIF(instr(1,model_5,""("",1) = 0,""tdb_Base"",""tdb_Detail""))"
      Sections(3).Cells(6).Name=   "CELL_6"
      Sections(3).Cells(6).Exp=   "model_6"
      Sections(3).Cells(6).StyleExp=   "IIF(instr(1,model_6,""("",1) > 13,  ""tdb_8ptFont"",IIF(instr(1,model_6,""("",1) = 0,""tdb_Base"",""tdb_Detail""))"
      Sections(3).Cells(7).Name=   "CELL_7"
      Sections(3).Cells(7).Exp=   "model_7"
      Sections(3).Cells(7).StyleExp=   "IIF(instr(1,model_7,""("",1) > 13,  ""tdb_8ptFont"",IIF(instr(1,model_7,""("",1) = 0,""tdb_Base"",""tdb_Detail""))"
      Sections(3).Cells(8).Name=   "CELL_8"
      Sections(3).Cells(8).Exp=   "model_8"
      Sections(3).Cells(8).StyleExp=   "IIF(instr(1,model_8,""("",1) > 13,  ""tdb_8ptFont"",IIF(instr(1,model_8,""("",1) = 0,""tdb_Base"",""tdb_Detail""))"
      Sections(4).Name=   "SECTION_2"
      Sections(4).Type=   4
      Sections(4).StyleExp=   "'tdb_Detail'"
      Sections(4).KeepWithPrev=   1
      Sections(4).Cells.Count=   9
      Sections(4).Cells(0).Name=   "CELL_0"
      Sections(4).Cells(0).Exp=   "category"
      Sections(4).Cells(0).StyleExp=   "IIF(category = """", ""tdb_Base"", ""tdb_Detail_LJ"")"
      Sections(4).Cells(1).Name=   "CELL_1"
      Sections(4).Cells(1).Exp=   "IIF(model_1 = """", """", IIF(part_1 = """", ""---"", part_1))"
      Sections(4).Cells(1).StyleExp=   "IIF(model_1 = """", ""tdb_Base"", ""tdb_Detail"")"
      Sections(4).Cells(2).Name=   "CELL_2"
      Sections(4).Cells(2).Exp=   "IIF(model_2 = """", """", IIF(part_2 = """", ""---"", part_2))"
      Sections(4).Cells(2).StyleExp=   "IIF(model_2 = """", ""tdb_Base"", ""tdb_Detail"")"
      Sections(4).Cells(3).Name=   "CELL_3"
      Sections(4).Cells(3).Exp=   "IIF(model_3 = """", """", IIF(part_3 = """", ""---"", part_3))"
      Sections(4).Cells(3).StyleExp=   "IIF(model_3 = """", ""tdb_Base"", ""tdb_Detail"")"
      Sections(4).Cells(4).Name=   "CELL_4"
      Sections(4).Cells(4).Exp=   "IIF(model_4 = """", """", IIF(part_4 = """", ""---"", part_4))"
      Sections(4).Cells(4).StyleExp=   "IIF(model_4 = """", ""tdb_Base"", ""tdb_Detail"")"
      Sections(4).Cells(5).Name=   "CELL_5"
      Sections(4).Cells(5).Exp=   "IIF(model_5 = """", """", IIF(part_5 = """", ""---"", part_5))"
      Sections(4).Cells(5).StyleExp=   "IIF(model_5 = """", ""tdb_Base"", ""tdb_Detail"")"
      Sections(4).Cells(6).Name=   "CELL_6"
      Sections(4).Cells(6).Exp=   "IIF(model_6 = """", """", IIF(part_6 = """", ""---"", part_6))"
      Sections(4).Cells(6).StyleExp=   "IIF(model_6 = """", ""tdb_Base"", ""tdb_Detail"")"
      Sections(4).Cells(7).Name=   "CELL_7"
      Sections(4).Cells(7).Exp=   "IIF(model_7 = """", """", IIF(part_7 = """", ""---"", part_7))"
      Sections(4).Cells(7).StyleExp=   "IIF(model_7 = """", ""tdb_Base"", ""tdb_Detail"")"
      Sections(4).Cells(8).Name=   "CELL_8"
      Sections(4).Cells(8).Exp=   "IIF(model_8 = """", """", IIF(part_8 = """", ""---"", part_8))"
      Sections(4).Cells(8).StyleExp=   "IIF(model_8 = """", ""tdb_Base"", ""tdb_Detail"")"
      Styles.Count    =   31
      Styles(0).Name  =   "tdb_Base"
      Styles(0).ParentName=   ""
      Styles(0).Font_Name=   "Arial"
      Styles(0).Font_Size=   9.75
      Styles(0).Font_Charset=   0
      Styles(0).NoClipping=   -1  'True
      Styles(1).Name  =   "tdb_TableBase"
      Styles(1).ParentName=   "tdb_Base"
      Styles(1).Font_Name=   "Arial"
      Styles(1).Font_Size=   9.75
      Styles(1).Font_Charset=   0
      Styles(1).BorderHT=   "tdb_ThinBlack"
      Styles(1).BorderHI=   "tdb_Invisible"
      Styles(1).BorderHB=   "tdb_ThinBlack"
      Styles(1).BorderVL=   "tdb_ThinBlack"
      Styles(1).BorderVI=   "tdb_ThinGray"
      Styles(1).BorderVR=   "tdb_ThinBlack"
      Styles(1).NoClipping=   -1  'True
      Styles(1).fprops=   4161536
      Styles(2).Name  =   "tdb_TableHeader"
      Styles(2).ParentName=   "tdb_TableBase"
      Styles(2).Font_Name=   "Arial"
      Styles(2).Font_Size=   9.75
      Styles(2).Font_Charset=   0
      Styles(2).ForeColor=   8388608
      Styles(2).BackColor=   15132390
      Styles(2).NoFill=   0   'False
      Styles(2).BorderHI=   "tdb_ThinGray"
      Styles(2).NoClipping=   -1  'True
      Styles(2).fprops=   65592
      Styles(3).Name  =   "tdb_8ptFont"
      Styles(3).ParentName=   "tdb_TableHeader"
      Styles(3).Font_Name=   "Arial"
      Styles(3).Font_Size=   8.25
      Styles(3).Font_Charset=   0
      Styles(3).TextAlign=   1
      Styles(3).MarginLeft=   2
      Styles(3).MarginTop=   2
      Styles(3).MarginRight=   2
      Styles(3).MarginBottom=   2
      Styles(3).BorderHT=   "tdb_ThinBlack"
      Styles(3).BorderHI=   "tdb_ThinBlack"
      Styles(3).BorderVI=   "tdb_ThinBlack"
      Styles(3).NoClipping=   -1  'True
      Styles(3).fprops=   4847641
      Styles(4).Name  =   "tdb_TableOddRow"
      Styles(4).ParentName=   "tdb_TableBase"
      Styles(4).Font_Name=   "Arial"
      Styles(4).Font_Size=   9.75
      Styles(4).Font_Charset=   0
      Styles(4).NoClipping=   -1  'True
      Styles(4).fprops=   0
      Styles(5).Name  =   "tdb_TableEvenRow"
      Styles(5).ParentName=   "tdb_TableOddRow"
      Styles(5).Font_Name=   "Arial"
      Styles(5).Font_Size=   9.75
      Styles(5).Font_Charset=   0
      Styles(5).BackColor=   8454143
      Styles(5).NoFill=   0   'False
      Styles(5).NoClipping=   -1  'True
      Styles(5).fprops=   48
      Styles(6).Name  =   "tdb_TableOddAlt"
      Styles(6).ParentName=   "tdb_TableOddRow"
      Styles(6).Font_Name=   "Arial"
      Styles(6).Font_Size=   9.75
      Styles(6).Font_Charset=   0
      Styles(6).NoClipping=   -1  'True
      Styles(6).fprops=   0
      Styles(7).Name  =   "tdb_TableEvenAlt"
      Styles(7).ParentName=   "tdb_TableEvenRow"
      Styles(7).Font_Name=   "Arial"
      Styles(7).Font_Size=   9.75
      Styles(7).Font_Charset=   0
      Styles(7).NoClipping=   -1  'True
      Styles(7).fprops=   0
      Styles(8).Name  =   "tdb_TableHighlight"
      Styles(8).ParentName=   "tdb_TableOddRow"
      Styles(8).Font_Name=   "Arial"
      Styles(8).Font_Size=   9.75
      Styles(8).Font_Charset=   0
      Styles(8).BackColor=   16777088
      Styles(8).NoFill=   0   'False
      Styles(8).BorderHT=   "tdb_ThickRed"
      Styles(8).BorderHI=   "tdb_ThickRed"
      Styles(8).BorderHB=   "tdb_ThickRed"
      Styles(8).BorderVL=   "tdb_ThickRed"
      Styles(8).BorderVI=   "tdb_ThickRed"
      Styles(8).BorderVR=   "tdb_ThickRed"
      Styles(8).NoClipping=   -1  'True
      Styles(8).fprops=   2064432
      Styles(9).Name  =   "tdb_TableFiller"
      Styles(9).ParentName=   "tdb_TableOddRow"
      Styles(9).Font_Name=   "Arial"
      Styles(9).Font_Size=   9.75
      Styles(9).Font_Charset=   0
      Styles(9).MarginTop=   0
      Styles(9).MarginBottom=   0
      Styles(9).NoClipping=   -1  'True
      Styles(9).fprops=   20480
      Styles(10).Name =   "tdb_TableFooter"
      Styles(10).ParentName=   "tdb_TableBase"
      Styles(10).Font_Name=   "Arial"
      Styles(10).Font_Size=   9.75
      Styles(10).Font_Charset=   0
      Styles(10).ForeColor=   8388608
      Styles(10).BackColor=   15132390
      Styles(10).NoFill=   0   'False
      Styles(10).BorderHI=   "tdb_ThinGray"
      Styles(10).NoClipping=   -1  'True
      Styles(10).fprops=   65592
      Styles(11).Name =   "tdb_Bullet"
      Styles(11).ParentName=   "tdb_Base"
      Styles(11).Font_Name=   "Arial"
      Styles(11).Font_Size=   9.75
      Styles(11).Font_Charset=   0
      Styles(11).ForePic=   "frmDailyScheduleSheet.frx":0004
      Styles(11).NoClipping=   -1  'True
      Styles(11).fprops=   536871424
      Styles(12).Name =   "tdb_BulletTriangle"
      Styles(12).ParentName=   "tdb_Base"
      Styles(12).Font_Name=   "Arial"
      Styles(12).Font_Size=   9.75
      Styles(12).Font_Charset=   0
      Styles(12).ForePic=   "frmDailyScheduleSheet.frx":02A6
      Styles(12).NoClipping=   -1  'True
      Styles(12).fprops=   536871424
      Styles(13).Name =   "tdb_BulletHollow"
      Styles(13).ParentName=   "tdb_Base"
      Styles(13).Font_Name=   "Arial"
      Styles(13).Font_Size=   9.75
      Styles(13).Font_Charset=   0
      Styles(13).ForePic=   "frmDailyScheduleSheet.frx":0548
      Styles(13).NoClipping=   -1  'True
      Styles(13).fprops=   536871424
      Styles(14).Name =   "tdb_PageHeader"
      Styles(14).ParentName=   "tdb_Base"
      Styles(14).Font_Name=   "Arial"
      Styles(14).Font_Size=   11.25
      Styles(14).Font_Bold=   -1  'True
      Styles(14).Font_Charset=   0
      Styles(14).TextAlign=   0
      Styles(14).NoClipping=   -1  'True
      Styles(14).fprops=   23068673
      Styles(15).Name =   "tdb_PageFooter"
      Styles(15).ParentName=   "tdb_PageHeader"
      Styles(15).Font_Name=   "Arial"
      Styles(15).Font_Size=   9.75
      Styles(15).Font_Charset=   0
      Styles(15).NoClipping=   -1  'True
      Styles(15).fprops=   0
      Styles(16).Name =   "tdb_RepHeader"
      Styles(16).ParentName=   "tdb_Base"
      Styles(16).Font_Name=   "Arial"
      Styles(16).Font_Size=   14.25
      Styles(16).Font_Bold=   -1  'True
      Styles(16).Font_Italic=   -1  'True
      Styles(16).Font_Charset=   0
      Styles(16).TextAlign=   1
      Styles(16).NoClipping=   -1  'True
      Styles(16).fprops=   56623105
      Styles(17).Name =   "tdb_RepHeader_RJ"
      Styles(17).ParentName=   "tdb_RepHeader"
      Styles(17).Font_Name=   "Arial"
      Styles(17).Font_Size=   9.75
      Styles(17).Font_Charset=   0
      Styles(17).TextAlign=   2
      Styles(17).ForePicFile=   "\\Tul-ares\vol1\USER\FRANCDE\CALQuality Unit Disp\WHRLOGO3.bmp"
      Styles(17).NoClipping=   -1  'True
      Styles(17).fprops=   536870913
      Styles(18).Name =   "tdb_RepFooter"
      Styles(18).ParentName=   "tdb_Base"
      Styles(18).Font_Name=   "Arial"
      Styles(18).Font_Size=   14
      Styles(18).Font_Bold=   -1  'True
      Styles(18).Font_Charset=   0
      Styles(18).TextAlign=   2
      Styles(18).NoClipping=   -1  'True
      Styles(18).fprops=   23068673
      Styles(19).Name =   "tdb_GroupHeaderBase"
      Styles(19).ParentName=   "tdb_Base"
      Styles(19).Font_Name=   "Arial"
      Styles(19).Font_Size=   9.75
      Styles(19).Font_Charset=   0
      Styles(19).NoClipping=   -1  'True
      Styles(19).fprops=   2097152
      Styles(20).Name =   "tdb_GroupFooterBase"
      Styles(20).ParentName=   "tdb_Base"
      Styles(20).Font_Name=   "Arial"
      Styles(20).Font_Size=   9.75
      Styles(20).Font_Charset=   0
      Styles(20).TextAlign=   2
      Styles(20).NoClipping=   -1  'True
      Styles(20).fprops=   2097153
      Styles(21).Name =   "tdb_GroupHeader1"
      Styles(21).ParentName=   "tdb_GroupHeaderBase"
      Styles(21).Font_Name=   "Arial"
      Styles(21).Font_Size=   14
      Styles(21).Font_Bold=   -1  'True
      Styles(21).Font_Charset=   0
      Styles(21).NoClipping=   -1  'True
      Styles(21).fprops=   20971520
      Styles(22).Name =   "tdb_GroupFooter1"
      Styles(22).ParentName=   "tdb_GroupFooterBase"
      Styles(22).Font_Name=   "Arial"
      Styles(22).Font_Size=   14
      Styles(22).Font_Bold=   -1  'True
      Styles(22).Font_Charset=   0
      Styles(22).NoClipping=   -1  'True
      Styles(22).fprops=   20971520
      Styles(23).Name =   "tdb_GroupHeader2"
      Styles(23).ParentName=   "tdb_GroupHeaderBase"
      Styles(23).Font_Name=   "Arial"
      Styles(23).Font_Size=   14
      Styles(23).Font_Charset=   0
      Styles(23).NoClipping=   -1  'True
      Styles(23).fprops=   4194304
      Styles(24).Name =   "tdb_GroupFooter2"
      Styles(24).ParentName=   "tdb_GroupFooterBase"
      Styles(24).Font_Name=   "Arial"
      Styles(24).Font_Size=   14
      Styles(24).Font_Charset=   0
      Styles(24).NoClipping=   -1  'True
      Styles(24).fprops=   4194304
      Styles(25).Name =   "tdb_GroupHeader3"
      Styles(25).ParentName=   "tdb_GroupHeaderBase"
      Styles(25).Font_Name=   "Arial"
      Styles(25).Font_Size=   12
      Styles(25).Font_Bold=   -1  'True
      Styles(25).Font_Charset=   0
      Styles(25).NoClipping=   -1  'True
      Styles(25).fprops=   20971520
      Styles(26).Name =   "tdb_GroupFooter3"
      Styles(26).ParentName=   "tdb_GroupFooterBase"
      Styles(26).Font_Name=   "Arial"
      Styles(26).Font_Size=   12
      Styles(26).Font_Bold=   -1  'True
      Styles(26).Font_Charset=   0
      Styles(26).NoClipping=   -1  'True
      Styles(26).fprops=   20971520
      Styles(27).Name =   "tdb_GroupHeader4"
      Styles(27).ParentName=   "tdb_GroupHeaderBase"
      Styles(27).Font_Name=   "Arial"
      Styles(27).Font_Size=   12
      Styles(27).Font_Charset=   0
      Styles(27).NoClipping=   -1  'True
      Styles(27).fprops=   4194304
      Styles(28).Name =   "tdb_GroupFooter4"
      Styles(28).ParentName=   "tdb_GroupFooterBase"
      Styles(28).Font_Name=   "Arial"
      Styles(28).Font_Size=   12
      Styles(28).Font_Charset=   0
      Styles(28).NoClipping=   -1  'True
      Styles(28).fprops=   4194304
      Styles(29).Name =   "tdb_Detail"
      Styles(29).ParentName=   "tdb_Base"
      Styles(29).Font_Name=   "Arial"
      Styles(29).Font_Size=   9
      Styles(29).Font_Charset=   0
      Styles(29).TextAlign=   1
      Styles(29).MarginLeft=   2
      Styles(29).MarginTop=   2
      Styles(29).MarginRight=   2
      Styles(29).MarginBottom=   2
      Styles(29).BorderHT=   "tdb_ThinBlack"
      Styles(29).BorderHI=   "tdb_ThinBlack"
      Styles(29).BorderHB=   "tdb_ThinBlack"
      Styles(29).BorderVL=   "tdb_ThinBlack"
      Styles(29).BorderVI=   "tdb_ThinBlack"
      Styles(29).BorderVR=   "tdb_ThinBlack"
      Styles(29).NoClipping=   -1  'True
      Styles(29).fprops=   8386561
      Styles(30).Name =   "tdb_Detail_LJ"
      Styles(30).ParentName=   "tdb_Detail"
      Styles(30).Font_Name=   "Arial"
      Styles(30).Font_Size=   9.75
      Styles(30).Font_Charset=   0
      Styles(30).TextAlign=   0
      Styles(30).NoClipping=   -1  'True
      Styles(30).fprops=   1
      Mappings.Count  =   5
      Mappings(0).Name=   "tdb_CheckboxV"
      Mappings(0).ValueItems.Count=   4
      Mappings(0).ValueItems(0).Key=   "False"
      Mappings(0).ValueItems(0).Picture=   "frmDailyScheduleSheet.frx":07EA
      Mappings(0).ValueItems(1).Key=   "True"
      Mappings(0).ValueItems(1).Default=   -1  'True
      Mappings(0).ValueItems(1).Picture=   "frmDailyScheduleSheet.frx":0884
      Mappings(0).ValueItems(2).Key=   ""
      Mappings(0).ValueItems(2).LinkedKey=   "False"
      Mappings(0).ValueItems(3).Key=   "0"
      Mappings(0).ValueItems(3).LinkedKey=   "False"
      Mappings(1).Name=   "tdb_CheckboxVBoxed"
      Mappings(1).ValueItems.Count=   4
      Mappings(1).ValueItems(0).Key=   "False"
      Mappings(1).ValueItems(0).Picture=   "frmDailyScheduleSheet.frx":091E
      Mappings(1).ValueItems(1).Key=   "True"
      Mappings(1).ValueItems(1).Default=   -1  'True
      Mappings(1).ValueItems(1).Picture=   "frmDailyScheduleSheet.frx":09B8
      Mappings(1).ValueItems(2).Key=   ""
      Mappings(1).ValueItems(2).LinkedKey=   "False"
      Mappings(1).ValueItems(3).Key=   "0"
      Mappings(1).ValueItems(3).LinkedKey=   "False"
      Mappings(2).Name=   "tdb_CheckboxX"
      Mappings(2).ValueItems.Count=   4
      Mappings(2).ValueItems(0).Key=   "False"
      Mappings(2).ValueItems(0).Picture=   "frmDailyScheduleSheet.frx":0A52
      Mappings(2).ValueItems(1).Key=   "True"
      Mappings(2).ValueItems(1).Default=   -1  'True
      Mappings(2).ValueItems(1).Picture=   "frmDailyScheduleSheet.frx":0AEC
      Mappings(2).ValueItems(2).Key=   ""
      Mappings(2).ValueItems(2).LinkedKey=   "False"
      Mappings(2).ValueItems(3).Key=   "0"
      Mappings(2).ValueItems(3).LinkedKey=   "False"
      Mappings(3).Name=   "tdb_CheckboxXBoxed"
      Mappings(3).ValueItems.Count=   4
      Mappings(3).ValueItems(0).Key=   "False"
      Mappings(3).ValueItems(0).Picture=   "frmDailyScheduleSheet.frx":0B86
      Mappings(3).ValueItems(1).Key=   "True"
      Mappings(3).ValueItems(1).Default=   -1  'True
      Mappings(3).ValueItems(1).Picture=   "frmDailyScheduleSheet.frx":0C20
      Mappings(3).ValueItems(2).Key=   ""
      Mappings(3).ValueItems(2).LinkedKey=   "False"
      Mappings(3).ValueItems(3).Key=   "0"
      Mappings(3).ValueItems(3).LinkedKey=   "False"
      Mappings(4).Name=   "tdb_CheckboxCircle"
      Mappings(4).ValueItems.Count=   4
      Mappings(4).ValueItems(0).Key=   "False"
      Mappings(4).ValueItems(0).Picture=   "frmDailyScheduleSheet.frx":0CBA
      Mappings(4).ValueItems(1).Key=   "True"
      Mappings(4).ValueItems(1).Default=   -1  'True
      Mappings(4).ValueItems(1).Picture=   "frmDailyScheduleSheet.frx":0D54
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
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Daily Schedule/Process Sheet"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmDailyScheduleSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private marrstrLine() As String
Private marrstrLocationId() As String
Private mrsPartCat As New ADODB.Recordset
Private mxarrReportData As XArrayDB         ' Report array
Private mxarrTemp As XArrayDB               ' Temporary array
Private mintMaxCategories As Integer        ' Maximum number of categories allowed
Private mintMaxModels As Integer            ' Maximum number of models allowed
Private mintMaxCols As Integer              ' Maximum number of array columns allowed
Private mintModelCount As Integer           ' Counter for current number of models
Private mintRows As Integer                 ' Counter for current number of rows
Private mintCols As Integer                 ' Counter for current number of columns
Private mintModelIndex As Integer           ' Current model index
Private mintSeqNumber As Integer            ' Current sequence number value
Private mintLocationIndex As String
Public mintNumberOfCopies As Integer        ' Number of print copies
Public mblnCancelPrint As Boolean           ' Cancel Print flag

Private mlngCurrentRow As Long
Private mobjXL As Excel.Application

Private Function PrintReport(strLine As String, strLocation As String, strToDate As String, _
                        blnPrintPreview As Boolean, index As Integer) As Integer
    ' Purpose:  Retrieve data and put it into an array used for the report object's data
    ' source.  Will retrieve data for selected date and beyond, so that data from 2nd
    ' shift can be included if 1st shift has < 16 models.
    '
    ' This the layout of the columns in the array:
    '
    ' col 1  - Category
    ' col 2  - Model Number (1)
    ' col 3  - Part Number (1)
    ' col 4  - Model Number (2)
    ' col 5  - Part Number (2)
    ' col 6  - Model Number (3)
    ' col 7  - Part Number (3)
    ' col 8  - Model Number (4)
    ' col 9  - Part Number (4)
    ' col 10 - Model Number (5)
    ' col 11 - Part Number (5)
    ' col 12 - Model Number (6)
    ' col 13 - Part Number (6)
    ' col 14 - Model Number (7)
    ' col 15 - Part Number (7)
    ' col 16 - Model Number (8)
    ' col 17 - Part Number (8)
                        
    ' Increase the timeout for the query so it can complete
    gconDatabase.CommandTimeout = 350

    PrintReport = 0
    
    ' Execute the query to create the recordset
    Set mrsPartCat = New ADODB.Recordset
    With mrsPartCat
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select count(distinct v_prod_schedule.model_number) as model_count, " & _
            "count(distinct v_mnb_category.category_description) as category_count " & _
            "from v_prod_schedule Join v_mnb_model_part_active " & _
            "on v_prod_schedule.line_id = v_mnb_model_part_active.line_id " & _
            "and v_prod_schedule.model_number = v_mnb_model_part_active.model_number " & _
            "Join v_mnb_category_part on v_mnb_model_part_active.part_id = v_mnb_category_part.part_id " & _
            "Join v_mnb_category on v_mnb_category_part.category_id = v_mnb_category.category_id " & _
            "Join v_mnb_model_part_stocking_location on v_mnb_model_part_active.line_id = v_mnb_model_part_stocking_location.line_id " & _
            "and v_mnb_model_part_active.model_number = v_mnb_model_part_stocking_location.model_number " & _
            "and v_mnb_model_part_active.part_id = v_mnb_model_part_stocking_location.part_id " & _
            "and v_mnb_model_part_active.part_sequence_number = v_mnb_model_part_stocking_location.part_sequence_number " & _
            "where v_prod_schedule.line_id like '" & strLine & "' " & _
            "and v_mnb_model_part_stocking_location.stocking_location_id = '" & strLocation & "' " & _
            "and start_date >= '" & strToDate & "' and minibill_only_flag = 1 " & _
            "and common_parts_category_flag = 0 and balance_to_start > 0"
            
        .Open
        
        If !model_count = 0 And !category_count = 0 Then
'            MsgBox "No data found for report"
            GoTo PROC_EXIT
        End If
        
        PrintReport = !model_count
        ' Assign maximum number of categories
        If !model_count < 9 Then
            mintMaxCategories = !category_count
        Else
            mintMaxCategories = 25
        End If
        .Close
        
        Dim strSelect As String
        strSelect = "select case when sub_assembly_id is null then rtrim(v_prod_schedule.model_number) else " & _
            "rtrim(v_prod_schedule.model_number) + '   [' + rtrim(sub_assembly_id) + '] ' end + '  (' + " & _
            "min(rtrim(cast(V_PROD_Schedule.balance_to_start as varchar(5)))) + ') ' as Model, " & _
            "min(case when v_mnb_model_part_active.quantity > 1 " & _
            "then rtrim(v_mnb_model_part_active.part_id) " & _
            "+ ' (' + rtrim(cast(v_mnb_model_part_active.quantity as char(5))) + ')' " & _
            "else rtrim(v_mnb_model_part_active.part_id) end) as Part, " & _
            "rtrim(v_mnb_category.category_description) as Category, " & _
            "min(category_sequence_number) as category_sequence_number, " & _
            "v_prod_schedule.sequence_number As sequence_number "

        
        .Source = strSelect & "From v_prod_schedule Join v_mnb_model_part_active " & _
            "on v_prod_schedule.line_id = v_mnb_model_part_active.line_id " & _
            "and v_prod_schedule.model_number = v_mnb_model_part_active.model_number " & _
            "Join v_mnb_category_part on v_mnb_model_part_active.part_id = v_mnb_category_part.part_id " & _
            "Join v_mnb_category on v_mnb_category_part.category_id = v_mnb_category.category_id " & _
            "Join v_mnb_model_part_stocking_location on v_mnb_model_part_active.line_id = v_mnb_model_part_stocking_location.line_id " & _
            "and v_mnb_model_part_active.model_number = v_mnb_model_part_stocking_location.model_number " & _
            "and v_mnb_model_part_active.part_id = v_mnb_model_part_stocking_location.part_id " & _
            "and v_mnb_model_part_active.part_sequence_number = v_mnb_model_part_stocking_location.part_sequence_number " & _
            "left outer join v_mnb_model_location_sub_assembly sub on " & _
            "v_mnb_model_part_stocking_location.model_number = sub.model_number and " & _
            "v_mnb_model_part_stocking_location.line_id = sub.line_id and " & _
            "v_mnb_model_part_stocking_location.stocking_location_id = sub.stocking_location_id " & _
            "where v_prod_schedule.line_id like '" & strLine & "' " & _
            "and v_mnb_model_part_stocking_location.stocking_location_id = '" & strLocation & "' " & _
            "and start_date >= '" & strToDate & "' and minibill_only_flag = 1 " & _
            "and common_parts_category_flag = 0 and balance_to_start > 0 " & _
            "group by sequence_number, v_prod_schedule.model_number, category_description, sub_assembly_id " & _
            "order by sequence_number, category_sequence_number, v_mnb_category.category_description, sub_assembly_id"
        .Open
    
        Dim intCatIndex As Integer
        Dim intPartIndex As Integer
        Dim intSeqNumber As Integer
        Dim strModel As String

        mintMaxCols = 34
        mintMaxModels = 16
        mintRows = 0
        mintModelCount = 0
        mintCols = 1
        mintModelIndex = 0
        mintSeqNumber = -1
    
        ' Populate a temporary array with data returned to the recordset.  This
        ' array will be broken up later into 2 sections if the number of models in the
        ' results set is > 8.
        
        Set mxarrTemp = New XArrayDB
        
        Do While Not .EOF
            
            ' Find category's row index in the temporary array
            intCatIndex = InsertCategory(.Fields("category").Value)
            
            ' Exit if no row index returned, meaning maximum number of categories
            ' has been reached
            If intCatIndex = 0 Then
'                .Close
'                Exit Do
                GoTo READ_NEXT
            End If
            
            strModel = .Fields("model").Value
            intSeqNumber = .Fields("sequence_number").Value
            
            ' Find column index for part number based on the value of the model
            ' and sequence number
            intPartIndex = InsertModel(strModel, intSeqNumber)
            
            ' Exit if no column index returned, meaning maximum number of models has
            ' been reached
            If intPartIndex = 0 Then
                .Close
                Exit Do
            End If
            
            ' Insert part number at location of category and model (next to model)
            mxarrTemp(intCatIndex, intPartIndex) = .Fields("part").Value
            
READ_NEXT:
            .MoveNext
        Loop
    End With
    
    Set mrsPartCat = Nothing
                
    Dim intX As Integer
    Dim intY As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intSections As Integer
    
    ' Copy model values to each row in the array.   This is to make the report display
    ' properly.  Some of the code in the report uses the model to determine
    ' what style to use when printing the part number.
    For intRow = 2 To mintRows
        For intCol = 2 To mintCols Step 2
            mxarrTemp(intRow, intCol) = mxarrTemp(1, intCol)
        Next
    Next

    ' Create array that will be used by the report
    Set mxarrReportData = New XArrayDB
    
    ' Determine the number of report sections based on the number of models in the
    ' temporary array
    If mintModelCount > 8 Then
        mxarrReportData.ReDim 1, mintRows * 2, 1, 17
        intSections = 2
    Else
        mxarrReportData.ReDim 1, mintRows, 1, 17
        intSections = 1
    End If

    intRow = 0
    intCol = 0

    ' Copy data to first section of report array
    For intY = 1 To mintRows
        For intX = 1 To 17
            mxarrReportData(intY, intX) = mxarrTemp(intY, intX)
        Next
    Next
    
    ' Copy data to second section of report array
    If intSections > 1 Then
        intRow = 0
        For intY = mintRows + 1 To mintRows * 2
            intRow = intRow + 1
            intCol = 17
            For intX = 1 To 17
                If intX = 1 Then
                    mxarrReportData(intY, 1) = mxarrTemp(intRow, 1)
                Else
                    intCol = intCol + 1
                    mxarrReportData(intY, intX) = mxarrTemp(intRow, intCol)
                End If
            Next
        Next
    End If
    
    With TDBDailyScheduleSheet

        If index = 1 Then
            ' Dump data into Excel
            Dim xlWorkBook As Excel.Workbook

            Set mobjXL = New Excel.Application
            Set xlWorkBook = mobjXL.Workbooks.Add

            mobjXL.Rows(mxarrTemp.UpperBound(1)).Insert
                        
            intX = 1
            mobjXL.Cells(1, intX) = "Categories"
            
            For intCol = 2 To mxarrTemp.UpperBound(2) Step 2
                intX = intX + 1
                mobjXL.Cells(1, intX) = mxarrTemp(1, intCol)
            Next
            
            intX = 0
            intY = 1
            
            For intRow = 1 To mxarrTemp.UpperBound(1)
                intY = intY + 1
                intX = 0
                For intCol = 1 To mxarrTemp.UpperBound(2) Step 2
                    intX = intX + 1
                    mobjXL.Cells(intY, intX) = mxarrTemp(intRow, intCol)
                Next
            Next
                
            mobjXL.Rows(1).Insert
            mobjXL.Rows(1).Insert

            mobjXL.Cells(1, 1) = "Line:"
            mobjXL.Cells(1, 2) = cboLine.Text
            mobjXL.Cells(1, 3) = "Location:"
            mobjXL.Cells(1, 4) = cboLocation.Text
            mobjXL.Cells(1, 5) = "Date:"
            mobjXL.Cells(1, 6) = cboDate.Text
            
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
Private Sub FillLocation()
' Whenever the line changes, the locations need refreshed
    
   ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare ADO variables
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    Dim strDisplay As String
    Dim strLine As String
    
    ' Instantiate variables
    Set rsList = New ADODB.Recordset
    
    ' Set up recordset fields
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .CursorType = adOpenForwardOnly
    
        If cboLine.ListIndex = -1 Then
            strLine = marrstrLine(0)
            cboLine.ListIndex = 0
        Else
            strLine = marrstrLine(cboLine.ListIndex)
        End If
       
        .Source = "select min(stocking_location_description) as stocking_location_description, " & _
                "v_prod_stocking_location.stocking_location_id " & _
                "From v_prod_stocking_location " & _
                "join v_prod_line_stocking_location on " & _
                "v_prod_stocking_location.stocking_location_id = v_prod_line_stocking_location.stocking_location_id " & _
                "where line_id like '" & strLine & _
                "' group by v_prod_stocking_location.stocking_location_id " & _
                "order by stocking_location_description"
        .Open
        
        Erase marrstrLocationId
        intIndex = .RecordCount + 1
        ReDim marrstrLocationId(intIndex)
        
        cboLocation.Clear
        cboLocation.AddItem "- - - - ALL LOCATIONS - - - -"
        marrstrLocationId(0) = "ALL"
        
        intIndex = 1
        
        Do While Not .EOF
            cboLocation.AddItem RTrim$(!stocking_location_description)
            marrstrLocationId(intIndex) = !stocking_location_id
            intIndex = intIndex + 1
            .MoveNext
        Loop

        .Close
        
        'Moves selection to the top of the location combo box
        If cboLocation.ListIndex = -1 Then
            If cboLocation.ListCount > 0 Then
                cboLocation.ListIndex = 0
            End If
        End If
         
 End With
 
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FillLocation", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
 
End Sub

Private Sub cboLine_GotFocus()
    cboLine.SelStart = 0
    cboLine.SelLength = Len(cboLine.Text)
End Sub

Private Sub cboLine_Validate(Cancel As Boolean)
    ' Purpose:  Validate the line entered
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' If there is no data in the field, display a message,
    ' set cancel to true and exit.
    If Len(Trim(cboLine.Text)) = 0 Then
        Cancel = False
        GoTo PROC_EXIT
    
    Else
        ' If the length of the field is one, add a
        ' leading zero.
        'If Len(cboLine.Text) = 1 Then
        '    cboLine.Text = "0" & cboLine.Text
        'End If
        
        ' If the line id has changed, look up the new
        ' line in the listbox.
        cboFindFirst cboLine
            
        ' If the line was not found in the list,
        ' display a message, set cancel to true and
        ' exit.
        If cboLine.ListIndex = -1 Then
            MsgBox cboLine.Text & " is not valid"
            Cancel = True
            GoTo PROC_EXIT
        End If
            
        Cancel = False
    End If
    
    Call FillLocation
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboLineID_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cboLocation_Change()
    cboFindFirst cboLocation
End Sub

Private Sub cboLocation_GotFocus()
    cboLocation.SelStart = 0
    cboLocation.SelLength = Len(cboLocation.Text)
End Sub

Private Sub cboLocation_Validate(Cancel As Boolean)
    ' Purpose:  Validate the location
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    If Len(Trim(cboLocation.Text)) = 0 Then
        Cancel = True
        MsgBox "Must make a Location Selection" & vbCrLf & "      Or Select Different Line"
        Call cboLine_GotFocus
        GoTo PROC_EXIT
   
    Else
        cboFindFirst cboLocation
        If cboLocation.ListIndex = -1 Then
            MsgBox "Location " & cboLocation.Text & " is not valid."
            Cancel = True
        Else
            Cancel = False
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cboLocation_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Sub FillComboBoxes()

    ' Purpose:   Read the database and load the line and date combo boxes.
    '    The descriptions for the line and locations are loaded to the
    'dropdowns, while the codes are loaded to an array at the same time.
    'The codes are the key fields to gather sql from the data bases.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare ADO variables
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    Dim strDisplay As String
    
    ' Instantiate variables
    Set rsList = New ADODB.Recordset
    
    ' Set up recordset fields
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .CursorType = adOpenForwardOnly
        
         .Source = "V_PROD_Line"
            .Open
        
        intIndex = 0
            Do While Not .EOF
                ReDim Preserve marrstrLine(intIndex)
                cboLine.AddItem RTrim$(!line_description)
                marrstrLine(intIndex) = !line_id
                intIndex = intIndex + 1
                .MoveNext
            Loop
        
           .Close
                      
        .Source = "select distinct convert(char(10), start_date, 101) " & _
                  "as start_date, datepart(year, start_date) as year_part, " & _
                  "datepart(month, start_date) as month_part, " & _
                  "datepart(day, start_date) as day_part " & _
                  "from v_prod_schedule " & _
                  "order by year_part, month_part, day_part"
        .Open

        Do While Not .EOF
            cboDate.AddItem !start_date
            .MoveNext
        Loop

        .Close
        
        If cboDate.ListIndex = -1 Then
            cboDate.ListIndex = 0
        End If
    End With
    
    Set rsList = Nothing
    
'    ReDim marrstrLocationId(0)
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FillComboBoxes", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cboDate_Change()
    cboFindFirst cboDate
End Sub

Private Sub cboDate_GotFocus()
    cboDate.SelStart = 0
    cboDate.SelLength = Len(cboDate.Text)
End Sub

Private Sub cboDate_KeyPress(KeyAscii As Integer)
    cboKeyPress cboDate, KeyAscii
End Sub

Private Sub cmdExitReport_Click()
    Unload Me
End Sub

Private Function InsertCategory(strCategory As String) As Integer
    ' Purpose:   Find row number of category if it exists in the temporary array.
    ' Otherwise, insert the category in the array, keeping the array sorted in category
    ' sequence.

    On Error GoTo PROC_ERR
                        
    Dim intRow As Integer
    Dim intCol As Integer
    Dim intReturn As Integer
                  
    ' Initialize first row of array if none exists yet
    If mxarrTemp.Count(1) = 0 Then
        mxarrTemp.ReDim 1, 1, 1, mintMaxCols
        mxarrTemp(1, 1) = strCategory
        mintRows = 1
        InsertCategory = 1
        GoTo PROC_EXIT
    End If
        
    InsertCategory = 0
                   
    ' Search for category in the array.  Exit if category is found, or if right location
    ' for a new row is found
    For intRow = 1 To mxarrTemp.UpperBound(1)
        ' Category found, assign row number to return value and exit loop
        If mxarrTemp(intRow, 1) = strCategory Then
            InsertCategory = intRow
            Exit For
        End If
        
        ' Location for new category found, exit loop
        If mxarrTemp(intRow, 1) > strCategory Then
            Exit For
        End If
    Next
        
    If InsertCategory = 0 Then
        ' Since return value is still 0, category was not found in the array.  Add a new
        ' row if the maximum number has not been reached yet.
        If mintRows < mintMaxCategories Then
            ' Append new row if loop index is > the number of rows in the array
            If intRow > mxarrTemp.Count(1) Then
                mxarrTemp.AppendRows (1)
            ' Otherwise, insert new row before row identified by loop index
            Else
                intReturn = mxarrTemp.InsertRows(intRow, 1)
                
                ' If inserting new first row, copy model values from previous first row
                If intRow = 1 Then
                    For intCol = 2 To mintCols Step 2
                        mxarrTemp(1, intCol) = mxarrTemp(2, intCol)
                    Next
                End If
            End If
            
            mxarrTemp(intRow, 1) = strCategory          'Assign category value
            InsertCategory = intRow                     'Assign return value
            mintRows = mintRows + 1                     'Increment row count
        End If
    End If
        
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    Call ShowError(Me.Name, "InsertCategory", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
    
End Function

Private Function InsertModel(strModel As String, seqNumber As Integer) As Integer
    ' Purpose:   Determine column index of model in the array, based on the sequence number.
    ' Insert model value into array if current sequence number has changed.  Returns
    ' column index for the part number.
    
    On Error GoTo PROC_ERR
            
    InsertModel = 0
            
    ' Start a new model column if sequence number changes.  This enables a model to
    ' display more than once if it appears multiple times in the schedule
    If seqNumber <> mintSeqNumber Then
    
        ' Start new model if the model limit has not yet been reached
        If mintModelCount < mintMaxModels Then
            mintModelCount = mintModelCount + 1     ' Number of models added
            mintModelIndex = mintModelIndex + 2     ' Array index of current model
            mintCols = mintCols + 2                 ' Number of total array columns
            mxarrTemp(1, mintModelIndex) = strModel ' Assign new model value
            InsertModel = mintModelIndex + 1        ' Calculate index for part number
            mintSeqNumber = seqNumber               ' Save current sequence number
        End If
        
    ' Sequence number matches, use current sequence number
    Else
        InsertModel = mintModelIndex + 1            ' Calculate index for part number
    End If
            
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    Call ShowError(Me.Name, "InsertModel", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
    
End Function

Private Sub cmdReport_Click(index As Integer)
    ' Print report for selected locations by calling PrintReport function
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim strLocation As String
    Dim strLine As String
    Dim strToDate As String
    Dim intCount As Integer
    Dim intTotal As Integer
    
    ' Set the variable values from the screen to the query.
    strLocation = marrstrLocationId(cboLocation.ListIndex)
    strLine = marrstrLine(cboLine.ListIndex)
    strToDate = Format(cboDate.Text, "mm/dd/yyyy")
     
    If strLocation = "ALL" Then
        ' Print a report for each location of selected line
        
        ' Prompt user to input Number of Copies to Print
        If index = 0 Then
            frmCopies.Show vbModal, Me
            If mblnCancelPrint Then Exit Sub
        End If
            
        ' Print report for each location directly to the printer
        Dim intI As Integer
        intTotal = 0
        For intI = 1 To UBound(marrstrLocationId)
            strLocation = marrstrLocationId(intI)
            mintLocationIndex = intI
            intCount = PrintReport(strLine, strLocation, strToDate, False, index)
            intTotal = intTotal + intCount
        Next
        If intTotal = 0 Then
            MsgBox "No data found for report"
        End If
    Else
        'Create report for selected location, in printpreview mode
        mintLocationIndex = cboLocation.ListIndex
        intCount = PrintReport(strLine, strLocation, strToDate, True, index)
        If intCount = 0 Then
            MsgBox "No data found for report"
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdDisplayReport_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub Form_Load()
    ' Purpose:  Load the form
    '******************************************************
    '*  Do not make changes to this report until Tulsa is
    '*  Contacted.  There were specific criteria required to
    '*  make this report come out like it does.
    '******************************************************
        
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
    
    Call FillComboBoxes
    Call FillLocation
    
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
    
    cboLine.Clear
    cboLocation.Clear
    cboDate.Clear
    
End Sub

Private Sub TDBDailyScheduleSheet_CellExpression(ByVal Section As Integer, ByVal Cell As Integer, Value As Variant)
    Static varPreviousBookmark As Variant
    Static blnSkipRecord As Boolean

    If Section = 1 And mlngCurrentRow > 1 Then
        Exit Sub
    ElseIf Section = 1 Then
        blnSkipRecord = False
        varPreviousBookmark = -1
    ElseIf Cell = 0 Then
        If varPreviousBookmark >= TDBDailyScheduleSheet.Bookmark Then
            blnSkipRecord = True
            Exit Sub
        End If
        blnSkipRecord = False
        varPreviousBookmark = TDBDailyScheduleSheet.Bookmark
    ElseIf blnSkipRecord Then
        Exit Sub
    End If

    If Cell = 0 Then
        mlngCurrentRow = mlngCurrentRow + 1
    End If
    If Not IsNull(Value) Then
        Value = Replace(Value, Chr(10), " ")
        Value = Replace(Value, Chr(13), " ")
    End If
    mobjXL.Cells(mlngCurrentRow, Cell + 1) = Value

End Sub

Private Sub TDBDailyScheduleSheet_OpenData()

    On Error GoTo PROC_ERR

    With TDBDailyScheduleSheet
            
        If cboLine.ListIndex > -1 Then
            .Parameters("line_id") = cboLine.Text
        End If
        
        If cboLocation.ListIndex > -1 Then
            .Parameters("location_id") = cboLocation.List(mintLocationIndex)
        End If
        
        If cboDate.ListIndex > -1 Then
            .Parameters("schedule_date") = cboDate.Text
        End If
        
        .Parameters("rec_count") = mintRows
        
        .Parameters("division_name") = _
            gclsMESApplication.Division
            
        Set .Array = mxarrReportData
    End With
            
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "TDBDailyScheduleSheet_OpenData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

