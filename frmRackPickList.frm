VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D6235E7-DBA2-11D1-B5DF-0060976089D0}#1.0#0"; "tdbr6.ocx"
Begin VB.Form frmRackPickList 
   Caption         =   "MiniBill - Rack Pick List Report Request..."
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
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
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdScheduleDateCalendar 
      CausesValidation=   0   'False
      Height          =   450
      Left            =   3480
      Picture         =   "frmRackPickList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1140
      Width           =   450
   End
   Begin VB.ComboBox cboCategory 
      Height          =   360
      Left            =   2100
      TabIndex        =   0
      Top             =   420
      Width           =   2955
   End
   Begin VB.ComboBox cboLine 
      Height          =   360
      Left            =   2100
      TabIndex        =   2
      Top             =   1860
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Display Report"
      Default         =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin TrueDBReports60Ctl.TDBReports TDBReport 
      Height          =   570
      Left            =   5820
      TabIndex        =   6
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1005
      Caption         =   "Model Counts"
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   1
      ErrorMsgCaption =   ""
      Filtered        =   0   'False
      DataMode        =   0
      DataMember      =   ""
      LinkSequence    =   1
      LinkOrder       =   0
      NameSubstitute  =   ""
      ConnectionString=   $"frmRackPickList.frx":018A
      ConnectStringType=   1
      OLEDBString     =   $"frmRackPickList.frx":0211
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      CursorLocation  =   3
      ConnectionTimeout=   15
      CommandTimeout  =   30
      RecordSource    =   $"frmRackPickList.frx":0298
      CursorType      =   3
      CommandType     =   1
      MaxRecords      =   0
      LinkType        =   0
      Master          =   ""
      CallDataRead    =   0   'False
      ConvertNullToEmpty=   -1  'True
      DesignConnection=   -1  'True
      DesignTimeout   =   5
      UnitsOfMeasurement=   2
      Vedit_ShowGrid  =   -1  'True
      Vedit_SnapToGrid=   0   'False
      Vedit_GridUnitWidth=   0.111111111111111
      Vedit_GridUnitHeight=   0.111111111111111
      Vedit_ShowCellExpressions=   -1  'True
      Norm_rect_left  =   0
      Norm_rect_top   =   0
      Norm_rect_right =   0
      Norm_rect_bottom=   0
      Virgin          =   0   'False
      Parameters.Count=   4
      Parameters(0).Name=   "report_date"
      Parameters(0).Type=   7
      Parameters(1).Name=   "lineid"
      Parameters(2).Name=   "CategoryName"
      Parameters(3).Name=   "Division_Name"
      Fields.Count    =   8
      Fields(0).Name  =   "category_description"
      Fields(0).DisplayName=   "category_description"
      Fields(0).FieldKind=   0
      Fields(0).DataSourceField=   "category_description"
      Fields(0).MaxLength=   30
      Fields(1).Name  =   "line_id"
      Fields(1).DisplayName=   "line_id"
      Fields(1).FieldKind=   0
      Fields(1).DataSourceField=   "line_id"
      Fields(1).MaxLength=   2
      Fields(2).Name  =   "model_number"
      Fields(2).DisplayName=   "model_number"
      Fields(2).FieldKind=   0
      Fields(2).DataSourceField=   "model_number"
      Fields(2).MaxLength=   20
      Fields(3).Name  =   "part_id"
      Fields(3).DisplayName=   "part_id"
      Fields(3).FieldKind=   0
      Fields(3).DataSourceField=   "part_id"
      Fields(3).MaxLength=   20
      Fields(4).Name  =   "part_description"
      Fields(4).DisplayName=   "part_description"
      Fields(4).FieldKind=   0
      Fields(4).DataSourceField=   "part_description"
      Fields(4).MaxLength=   50
      Fields(5).Name  =   "quantity"
      Fields(5).DisplayName=   "quantity"
      Fields(5).Type  =   3
      Fields(5).FieldKind=   0
      Fields(5).DataSourceField=   "quantity"
      Fields(5).MaxLength=   4
      Fields(6).Name  =   "level_number"
      Fields(6).DisplayName=   "level_number"
      Fields(6).FieldKind=   0
      Fields(6).DataSourceField=   "level_number"
      Fields(6).MaxLength=   2
      Fields(7).Name  =   "parts_required"
      Fields(7).DisplayName=   "parts_required"
      Fields(7).Type  =   3
      Fields(7).FieldKind=   0
      Fields(7).DataSourceField=   "parts_required"
      Fields(7).MaxLength=   4
      Sections.Count  =   7
      Sections(0).Name=   "PageHeader"
      Sections(0).Type=   1
      Sections(0).StyleExp=   "'tab_RepHeader_RJ'"
      Sections(0).Cells.Count=   1
      Sections(0).Cells(0).Name=   "PageHeader"
      Sections(0).Cells(0).Exp=   $"frmRackPickList.frx":0607
      Sections(0).Cells(0).Placement=   2
      Sections(0).Cells(0).CallBeforePrint=   -1  'True
      Sections(1).Name=   "ColumnHeading1"
      Sections(1).Type=   1
      Sections(1).Condition=   "GroupName="""""
      Sections(1).StyleExp=   "'tdb_TableHeaderBold'"
      Sections(1).dtopts=   2
      Sections(1).Cells.Count=   1
      Sections(1).Cells(0).Name=   "category_heading"
      Sections(1).Cells(0).Exp=   """Category"""
      Sections(1).Cells(0).Width=   100
      Sections(1).Cells(0).PrivateStyle=   -1  'True
      Sections(1).Cells(0).Style.Name=   "<private>"
      Sections(1).Cells(0).Style.ParentName=   "tdb_TableHeaderBold"
      Sections(1).Cells(0).Style.Font_Name=   "Arial"
      Sections(1).Cells(0).Style.Font_Size=   9
      Sections(1).Cells(0).Style.Font_Bold=   -1  'True
      Sections(1).Cells(0).Style.Font_Italic=   0   'False
      Sections(1).Cells(0).Style.Font_Underline=   0   'False
      Sections(1).Cells(0).Style.Font_Strikeout=   0   'False
      Sections(1).Cells(0).Style.Font_Charset=   0
      Sections(1).Cells(0).Style.TextAlign=   0
      Sections(1).Cells(0).Style.TextVAlign=   0
      Sections(1).Cells(0).Style.TextWrap=   -1  'True
      Sections(1).Cells(0).Style.ForeColor=   8388608
      Sections(1).Cells(0).Style.BackColor=   16777088
      Sections(1).Cells(0).Style.NoFill=   0   'False
      Sections(1).Cells(0).Style.BackPicFile=   ""
      Sections(1).Cells(0).Style.ForePicFile=   ""
      Sections(1).Cells(0).Style.BackPicVertPlacement=   0
      Sections(1).Cells(0).Style.BackPicHorzPlacement=   0
      Sections(1).Cells(0).Style.ForePicPlacement=   0
      Sections(1).Cells(0).Style.ForePicDrawMode=   0
      Sections(1).Cells(0).Style.MarginLeft=   6
      Sections(1).Cells(0).Style.MarginTop=   6
      Sections(1).Cells(0).Style.MarginRight=   6
      Sections(1).Cells(0).Style.MarginBottom=   6
      Sections(1).Cells(0).Style.HasBorders=   -1  'True
      Sections(1).Cells(0).Style.BorderHT=   "tdb_ThinBlack"
      Sections(1).Cells(0).Style.BorderHI=   "tdb_ThinGray"
      Sections(1).Cells(0).Style.BorderHB=   "tdb_ThinBlack"
      Sections(1).Cells(0).Style.BorderVL=   "tdb_ThinBlack"
      Sections(1).Cells(0).Style.BorderVI=   "tdb_ThinGray"
      Sections(1).Cells(0).Style.BorderVR=   "tdb_ThinBlack"
      Sections(1).Cells(0).Style.NoClipping=   -1  'True
      Sections(1).Cells(0).Style.RTF=   0   'False
      Sections(1).Cells(0).Style.fprops=   1
      Sections(2).Name=   "ColumnHeading2"
      Sections(2).Type=   1
      Sections(2).StyleExp=   "'tab_TableHeader_LJ'"
      Sections(2).Tabulator=   "RackDetail"
      Sections(2).Cells.Count=   8
      Sections(2).Cells(0).Name=   "DummyHeader"
      Sections(2).Cells(1).Name=   "LineHeader"
      Sections(2).Cells(1).Exp=   """Line"""
      Sections(2).Cells(2).Name=   "Model"
      Sections(2).Cells(2).Exp=   """Model Number"""
      Sections(2).Cells(3).Name=   "ScheduleQtyHeader"
      Sections(2).Cells(3).Exp=   """Schedule Qty"""
      Sections(2).Cells(4).Name=   "PartHeader"
      Sections(2).Cells(4).Exp=   """Part Number"""
      Sections(2).Cells(5).Name=   "PartDescHeader"
      Sections(2).Cells(5).Exp=   """Part Description"""
      Sections(2).Cells(6).Name=   "LevelHeader"
      Sections(2).Cells(6).Exp=   """Level"""
      Sections(2).Cells(7).Name=   "PartsRequiredHeading"
      Sections(2).Cells(7).Exp=   """Parts Required"""
      Sections(3).Name=   "CategoryContinuedHeading"
      Sections(3).Type=   1
      Sections(3).Condition=   "IsTopOfPage() and GroupName = """""
      Sections(3).StyleExp=   "'tab_TableEvenRow_LJ'"
      Sections(3).Cells.Count=   1
      Sections(3).Cells(0).Name=   "CategoryContinued"
      Sections(3).Cells(0).Exp=   "trim(category_description)"
      Sections(4).Name=   "CategoryHeader"
      Sections(4).Condition=   "HasChanged(category_description) and GroupName = """" and not IsTopOfPage()"
      Sections(4).StyleExp=   "'tab_TableEvenRow_LJ'"
      Sections(4).Cells.Count=   1
      Sections(4).Cells(0).Name=   "Category"
      Sections(4).Cells(0).Exp=   "category_description"
      Sections(4).Cells(0).Width=   100
      Sections(5).Name=   "RackDetail"
      Sections(5).Type=   4
      Sections(5).StyleExp=   "'tdb_TableEvenRow'"
      Sections(5).Cells.Count=   8
      Sections(5).Cells(0).Name=   "Dumm1"
      Sections(5).Cells(0).Width=   7
      Sections(5).Cells(1).Name=   "Line"
      Sections(5).Cells(1).Exp=   "line_id"
      Sections(5).Cells(1).Width=   6
      Sections(5).Cells(2).Name=   "Model"
      Sections(5).Cells(2).Exp=   "rtrim(model_number)"
      Sections(5).Cells(2).Width=   20
      Sections(5).Cells(3).Name=   "ScheduleQuantity"
      Sections(5).Cells(3).Exp=   "cstr(quantity)"
      Sections(5).Cells(3).Width=   12
      Sections(5).Cells(4).Name=   "Part"
      Sections(5).Cells(4).Exp=   "rtrim(part_id)"
      Sections(5).Cells(4).Width=   15
      Sections(5).Cells(5).Name=   "PartDescription"
      Sections(5).Cells(5).Exp=   "rtrim(part_description)"
      Sections(5).Cells(5).Width=   30
      Sections(5).Cells(6).Name=   "Level"
      Sections(5).Cells(6).Exp=   "level_number"
      Sections(5).Cells(6).Width=   7
      Sections(5).Cells(7).Name=   "PartsRequired"
      Sections(5).Cells(7).Exp=   "cstr(parts_required)"
      Sections(5).Cells(7).Width=   13
      Sections(6).Name=   "ReportFooter"
      Sections(6).Type=   2
      Sections(6).Cells.Count=   2
      Sections(6).Cells(0).Name=   "PageNumber"
      Sections(6).Cells(0).Exp=   """Page "" + cstr(PageNo())"
      Sections(6).Cells(0).StyleExp=   "'tdb_RepHeader'"
      Sections(6).Cells(1).Name=   "DateTime"
      Sections(6).Cells(1).Exp=   "format(now, ""mm/dd/yyyy hh:nn AM/PM"")"
      Sections(6).Cells(1).StyleExp=   "'tdb_RepHeader_RJ_NoPicture'"
      Styles.Count    =   14
      Styles(0).Name  =   "tdb_Base"
      Styles(0).ParentName=   ""
      Styles(0).Font_Size=   9.75
      Styles(0).Font_Charset=   0
      Styles(0).TextAlign=   0
      Styles(0).NoClipping=   -1  'True
      Styles(1).Name  =   "tdb_PageHeader"
      Styles(1).ParentName=   "tdb_Base"
      Styles(1).Font_Size=   9
      Styles(1).Font_Bold=   -1  'True
      Styles(1).Font_Italic=   -1  'True
      Styles(1).Font_Charset=   0
      Styles(1).TextAlign=   1
      Styles(1).BorderHT=   "tdb_ThinBlack"
      Styles(1).BorderHI=   "tdb_ThinBlack"
      Styles(1).BorderHB=   "tdb_ThinBlack"
      Styles(1).BorderVL=   "tdb_ThinBlack"
      Styles(1).BorderVI=   "tdb_ThinBlack"
      Styles(1).BorderVR=   "tdb_ThinBlack"
      Styles(1).NoClipping=   -1  'True
      Styles(1).fprops=   56590385
      Styles(2).Name  =   "tab_pageHeader_LJ"
      Styles(2).ParentName=   "tdb_PageHeader"
      Styles(2).Font_Size=   9.75
      Styles(2).Font_Charset=   0
      Styles(2).TextAlign=   0
      Styles(2).NoClipping=   -1  'True
      Styles(2).fprops=   1
      Styles(3).Name  =   "tdb_TableBase"
      Styles(3).ParentName=   "tdb_Base"
      Styles(3).Font_Name=   "Arial"
      Styles(3).Font_Size=   9.75
      Styles(3).Font_Charset=   0
      Styles(3).TextAlign=   0
      Styles(3).BorderHT=   "tdb_ThinBlack"
      Styles(3).BorderHI=   "tdb_Invisible"
      Styles(3).BorderHB=   "tdb_ThinBlack"
      Styles(3).BorderVL=   "tdb_ThinBlack"
      Styles(3).BorderVI=   "tdb_ThinGray"
      Styles(3).BorderVR=   "tdb_ThinBlack"
      Styles(3).NoClipping=   -1  'True
      Styles(3).fprops=   4161536
      Styles(4).Name  =   "tdb_TableOddRow"
      Styles(4).ParentName=   "tdb_PageHeader"
      Styles(4).Font_Size=   9
      Styles(4).Font_Charset=   0
      Styles(4).TextAlign=   0
      Styles(4).BackColor=   16777088
      Styles(4).NoClipping=   -1  'True
      Styles(4).fprops=   4194353
      Styles(5).Name  =   "tdb_TableEvenRow"
      Styles(5).ParentName=   "tdb_TableOddRow"
      Styles(5).Font_Size=   9.75
      Styles(5).Font_Charset=   0
      Styles(5).TextAlign=   0
      Styles(5).BackColor=   8454143
      Styles(5).NoClipping=   -1  'True
      Styles(5).fprops=   50331696
      Styles(6).Name  =   "tab_TableEvenRow_LJ"
      Styles(6).ParentName=   "tdb_TableEvenRow"
      Styles(6).Font_Size=   9.75
      Styles(6).Font_Charset=   0
      Styles(6).TextAlign=   0
      Styles(6).NoClipping=   -1  'True
      Styles(6).fprops=   50331649
      Styles(7).Name  =   "tdb_RepHeader"
      Styles(7).ParentName=   "tdb_Base"
      Styles(7).Font_Name=   "Arial"
      Styles(7).Font_Size=   14.25
      Styles(7).Font_Bold=   -1  'True
      Styles(7).Font_Italic=   -1  'True
      Styles(7).Font_Charset=   0
      Styles(7).TextAlign=   0
      Styles(7).NoClipping=   -1  'True
      Styles(7).fprops=   56623105
      Styles(8).Name  =   "tdb_RepHeader_RJ_NoPicture"
      Styles(8).ParentName=   "tdb_RepHeader"
      Styles(8).Font_Size=   9.75
      Styles(8).Font_Charset=   0
      Styles(8).TextAlign=   2
      Styles(8).NoClipping=   -1  'True
      Styles(8).fprops=   1
      Styles(9).Name  =   "tab_RepHeader_RJ"
      Styles(9).ParentName=   "tdb_RepHeader"
      Styles(9).Font_Size=   9.75
      Styles(9).Font_Charset=   0
      Styles(9).TextAlign=   2
      Styles(9).ForePic=   "frmRackPickList.frx":0723
      Styles(9).NoClipping=   -1  'True
      Styles(9).fprops=   541065281
      Styles(10).Name =   "tdb_PageFooter"
      Styles(10).ParentName=   "tdb_PageHeader"
      Styles(10).Font_Size=   9.75
      Styles(10).Font_Charset=   0
      Styles(10).TextAlign=   0
      Styles(10).NoClipping=   -1  'True
      Styles(10).fprops=   0
      Styles(11).Name =   "tdb_TableHeader"
      Styles(11).ParentName=   "tdb_TableBase"
      Styles(11).Font_Size=   9
      Styles(11).Font_Charset=   0
      Styles(11).TextAlign=   1
      Styles(11).ForeColor=   8388608
      Styles(11).BackColor=   16777088
      Styles(11).NoFill=   0   'False
      Styles(11).BorderHI=   "tdb_ThinGray"
      Styles(11).NoClipping=   -1  'True
      Styles(11).fprops=   21037113
      Styles(12).Name =   "tdb_TableHeaderBold"
      Styles(12).ParentName=   "tdb_TableHeader"
      Styles(12).Font_Size=   9.75
      Styles(12).Font_Bold=   -1  'True
      Styles(12).Font_Charset=   0
      Styles(12).TextAlign=   0
      Styles(12).NoClipping=   -1  'True
      Styles(12).fprops=   16777217
      Styles(13).Name =   "tab_TableHeader_LJ"
      Styles(13).ParentName=   "tdb_TableHeader"
      Styles(13).Font_Size=   9.75
      Styles(13).Font_Charset=   0
      Styles(13).TextAlign=   0
      Styles(13).NoClipping=   -1  'True
      Styles(13).fprops=   1
      Lines.Count     =   3
      Lines(0).Name   =   "tdb_Invisible"
      Lines(1).Name   =   "tdb_ThinBlack"
      Lines(1).Thickness=   2
      Lines(2).Name   =   "tdb_ThinGray"
      Lines(2).Thickness=   2
      Lines(2).Color  =   8421504
      Profiles.Count  =   1
      Profiles(0).Name=   "PROFILE_0"
      Profiles(0).Active=   -1  'True
      Profiles(0).Draft=   -1  'True
      Profiles(0).PreviewModal=   0   'False
      Profiles(0).PreviewMaximized=   -1  'True
      Profiles(0).PreviewInitialZoom=   75
   End
   Begin MSMask.MaskEdBox mskScheduleDate 
      Height          =   360
      Left            =   2100
      TabIndex        =   1
      Tag             =   "Obsolete Date:"
      Top             =   1140
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   635
      _Version        =   393216
      ClipMode        =   1
      MaxLength       =   10
      Format          =   "mm/dd/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Schedule Date:"
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
      Left            =   255
      TabIndex        =   9
      Top             =   1200
      Width           =   1605
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
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
      Index           =   0
      Left            =   930
      TabIndex        =   8
      Top             =   480
      Width           =   1020
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
      Index           =   6
      Left            =   1380
      TabIndex        =   7
      Top             =   1920
      Width           =   510
   End
End
Attribute VB_Name = "frmRackPickList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngCurrentRow As Long
Private marrstrCategory() As String


Private Sub cboCategory_Change()
    cboFindFirst cboCategory
End Sub

Private Sub cboCategory_GotFocus()
    cboCategory.SelStart = 0
    cboCategory.SelLength = Len(cboCategory.Text)
End Sub

Private Sub cboCategory_KeyPress(KeyAscii As Integer)
    cboKeyPress cboCategory, KeyAscii
End Sub



Private Sub cboCategory_Validate(Cancel As Boolean)
    If Len(Trim(cboCategory.Text)) = 0 Then
        Cancel = False
        Exit Sub
    End If
    cboFindFirst cboCategory
    If cboCategory.ListIndex = -1 Then
        MsgBox "Category is invalid."
        Cancel = True
        Exit Sub
    Else
        Cancel = False
    End If
End Sub

Private Sub cboLine_Change()
    If Len(cboLine.Text) = 2 Then
        cboFindFirst cboLine
    End If
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
        If Len(cboLine.Text) = 1 Then
            cboLine.Text = "0" & cboLine.Text
        End If
        
        ' If the line id has changed, look up the new
        ' line in the listbox.
        cboFindFirst cboLine
            
        ' If the line was not found in the list,
        ' display a message, set cancel to true and
        ' exit.
        If cboLine.ListIndex = -1 Then
            MsgBox "Line " & cboLine.Text & " is not valid"
            Cancel = True
            GoTo PROC_EXIT
        End If
            
        Cancel = False
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboLineID_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Sub FillComboBoxes()
    ' Purpose:   Read the databasse and load the line and location combo boxes
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare ADO variables
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    
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
        
        Do While Not .EOF
            cboLine.AddItem !line_id
            .MoveNext
        Loop
        
        .Close
        
        .Source = "select * from v_mnb_Category " & _
            "where minibill_only_flag = 0 " & _
            "order by Category_description asc"
        .LockType = adLockReadOnly
        .Open
    
        If gconDatabase.Errors.Count > 0 Then
            Err.Raise gconDatabase.Errors(0).NativeError, _
                "RetrieveCategoryData", _
                gconDatabase.Errors(0).Description
        End If

        ' if no records were retrieved, add a new record to the
        ' recordset and reset fields to their original value.
        If .EOF Then
            MsgBox ("No records were retrieved from Category table")
            GoTo PROC_EXIT
        End If
    
        ' Go to the first record in the recordset and set the
        ' line ID
        ' Loop through the file
        cboCategory.Clear
        
        ReDim marrstrCategory(.RecordCount - 1)
        Do While Not .EOF
            cboCategory.AddItem Trim(!Category_description)
            marrstrCategory(cboCategory.NewIndex) = !Category_id
            .MoveNext
        Loop
        
        .Close
    End With
    
    Set rsList = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveLineAndLocation", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub



Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdReport_Click(index As Integer)
    ' Purpose:  Build data for report and display the
    '           report.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare cancel boolean variable
    Dim blnCancel As Boolean
    Dim strWhere As String
    Dim strSQL As String
    Dim intOrderByPos As Integer
    Dim intWherePos As Integer
    Dim intIndex As Integer
    Dim lngRecords As Long
    
    ' Validate Category
    Call cboCategory_Validate(blnCancel)
    If blnCancel Then
        cboCategory.SetFocus
        GoTo PROC_EXIT
    End If
    
    ' Validate Schedule Date
    Call mskScheduleDate_Validate(blnCancel)
    If blnCancel Then
        mskScheduleDate.SetFocus
        GoTo PROC_EXIT
    End If
    
    ' Validate the Location Type
    Call cboLine_Validate(blnCancel)
    If blnCancel Then
        cboLine.SetFocus
        GoTo PROC_EXIT
    End If
    
    strWhere = "where start_date = '" & mskScheduleDate.Text & "' and minibill_only_flag = 0 " & _
        "and v_prod_schedule.balance_to_start > 0 "
    If cboCategory.ListIndex > -1 Then
        strWhere = strWhere & " and v_mnb_Category.Category_id = '" & marrstrCategory(cboCategory.ListIndex) & "' "
    End If
    If cboLine.ListIndex > -1 Then
        strWhere = strWhere & " and v_prod_schedule.line_id = '" & cboLine.Text & "' "
    End If
    
    
    With Me.TDBReport
        intOrderByPos = InStr(1, UCase(.RecordSource), "ORDER BY", vbTextCompare)
        intWherePos = InStr(1, UCase(.RecordSource), "WHERE", vbTextCompare)
        strSQL = .RecordSource
        If intWherePos > 0 Then
            strSQL = Left(strSQL, intWherePos - 1) & _
                strWhere & Mid(strSQL, intOrderByPos, Len(strSQL) - intOrderByPos + 1)
        Else
            strSQL = Left(strSQL, intOrderByPos - 1) & _
                strWhere & Mid(strSQL, intOrderByPos, Len(strSQL) - intOrderByPos + 1)
        End If
        .RecordSource = strSQL
        gconDatabase.CursorLocation = adUseClient
        gconDatabase.Execute strSQL, lngRecords
        If lngRecords = 0 Then
            MsgBox "No information is available for the current selections."
            GoTo PROC_EXIT
        End If
        
        .PrintPreview

    End With
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdDisplayReport_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub Form_Load()
    ' Purpose:  Load the form
        
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
    mskScheduleDate.Text = Format(Now, "mm/dd/yyyy")
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub




Private Sub TDBReport_OpenData()
    On Error GoTo PROC_ERR
        
    With TDBReport
        gconDatabase.CursorLocation = adUseClient
        Set .SourceRecordset = gconDatabase.Execute(.RecordSource)
        .Parameters("report_date") = mskScheduleDate.Text
        If cboLine.ListIndex > -1 Then
            .Parameters("lineid") = cboLine.Text
        Else
            .Parameters("lineid") = Null
        End If
        
        If cboCategory.ListIndex > -1 Then
            .Parameters("CategoryName") = cboCategory.Text
        Else
            .Parameters("CategoryName") = vbNullString
        End If
        .Parameters("Division_Name") = _
            gclsMESApplication.Division
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "TDBDataReport_OpenData", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mskScheduleDate_GotFocus()
    ' Purpose:  Select the field for easy update
    
    mskScheduleDate.SelStart = 0
    mskScheduleDate.SelLength = 10
End Sub

Private Sub mskScheduleDate_Validate(Cancel As Boolean)
    ' Purpose:  Validate the field to make sure that it is
    '           either a valid date or empty.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    If mskScheduleDate.Text = "__/__/____" Then
        MsgBox "Schedule date must be entered."
        Cancel = True
        GoTo PROC_EXIT
    ElseIf IsDate(mskScheduleDate.Text) Then
        Cancel = False
        GoTo PROC_EXIT
    Else
        Cancel = True
        MsgBox "Invalid Schedule Date Entered!", _
            vbExclamation + vbOKOnly, _
            "Schedule Date Validateion Error"
        mskScheduleDate.SelStart = 0
        mskScheduleDate.SelLength = 10
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mskScheduleDate_Validate", _
        Err.Number, Err.Description)
End Sub

Private Sub cmdScheduleDateCalendar_Click()
    dlgCalendar.mdteSelectedDate = CDate(mskScheduleDate.Text)
    dlgCalendar.Show vbModal
    If Not IsNull(dlgCalendar.mdteSelectedDate) Then
        mskScheduleDate.Text = Format( _
            dlgCalendar.mdteSelectedDate, "mm/dd/yyyy")
    End If
End Sub



