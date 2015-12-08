VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmSplitPartMaintenance 
   Caption         =   "MiniBill - Split Part Maintenance"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
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
   ScaleHeight     =   558
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   1  'CenterOwner
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   8370
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   8415
      _cx             =   14843
      _cy             =   14764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   12640511
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   -1  'True
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin VB.CommandButton cmdSelectPart 
         Caption         =   "Select &Part"
         Height          =   435
         Left            =   195
         TabIndex        =   8
         Top             =   7800
         Width           =   1755
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Refresh Display"
         Height          =   435
         Left            =   6060
         TabIndex        =   3
         Top             =   480
         Width           =   2085
      End
      Begin VB.ComboBox cboLine 
         Height          =   360
         Left            =   705
         TabIndex        =   0
         Top             =   600
         Width           =   675
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   6540
         TabIndex        =   6
         Top             =   7740
         Width           =   1635
      End
      Begin VB.ComboBox cboModel 
         Height          =   360
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGMiniBill 
         Height          =   6435
         Left            =   195
         TabIndex        =   4
         Top             =   1140
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   11351
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Part"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Orig Sequence"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Description"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Quantity"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "# of Splits"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   -1  'True
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3387"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3307"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=6191"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6112"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=1296"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1217"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(24)=   "Column(3).AllowFocus=0"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(26)=   "Column(4).Width=2196"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2117"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=2"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         TabAction       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=3,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bold=0,.fontsize=750,.italic=0,.underline=0"
         _StyleDefs(25)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(26)  =   ":id=13,.fontname=Arial"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=675,.italic=0"
         _StyleDefs(29)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(30)  =   ":id=14,.fontname=Small Fonts"
         _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(34)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=82,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=79,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=80,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=81,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(60)  =   "Named:id=33:Normal"
         _StyleDefs(61)  =   ":id=33,.parent=0,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
         _StyleDefs(62)  =   ":id=33,.charset=0"
         _StyleDefs(63)  =   ":id=33,.fontname=Arial"
         _StyleDefs(64)  =   "Named:id=34:Heading"
         _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(66)  =   ":id=34,.wraptext=-1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(67)  =   ":id=34,.strikethrough=0,.charset=0"
         _StyleDefs(68)  =   ":id=34,.fontname=Arial"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
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
         Left            =   120
         TabIndex        =   7
         Top             =   660
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1485
         TabIndex        =   5
         Top             =   660
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmSplitPartMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsModelPartLocation As ADODB.Recordset

Private mblnRecChanged As Boolean

Private mxarrLocations As XArrayDB
Private mxarrMiniBill As XArrayDB

Private mblnBuildingGrid As Boolean

Private Sub cboLine_Change()
    If Len(cboLine.Text) = 2 Then
        cboFindFirst cboLine
    End If
End Sub


Private Sub cboLine_Click()
    RetrieveModelData
    LoadLocations
    If cboModel.ListCount > 0 And mxarrLocations.UpperBound(1) > 0 And Screen.ActiveForm Is Me Then
        cboModel.SetFocus
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
        MsgBox "Line is required."
        Cancel = True
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
        
        If mxarrLocations.UpperBound(1) = 0 Then
            Cancel = True
            cboLine_GotFocus
        Else
            Cancel = False
        End If
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboLineID_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub




Private Sub cboModel_Change()
    cboFindFirst cboModel
End Sub


Private Sub cboModel_GotFocus()
    cboModel.SelStart = 0
    cboModel.SelLength = Len(cboModel.Text)
End Sub

Private Sub cboModel_KeyPress(KeyAscii As Integer)
    cboKeyPress cboModel, KeyAscii
End Sub

Private Sub cboModel_Validate(Cancel As Boolean)
    ' Purpose:  Check to see that a model was selected.
    
    If Len(cboModel.Text) = 0 Then
        MsgBox "Model is required."
        Cancel = True
        Exit Sub
    End If
    
    cboFindFirst cboModel
    If cboModel.ListIndex = -1 Then
        MsgBox "Invalid Model selected."
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdDisplay_Click()
    ' Purpose:  Build the model/part display
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    Dim intRetCode As Integer
    Dim strOrderBy As String
    Dim strWhere As String
    
    mblnBuildingGrid = True
    
    Set mxarrMiniBill = New XArrayDB
    
    strOrderBy = "order by original_sequence_number"
    
    Set rsList = New ADODB.Recordset
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select v_mnb_model_part.part_id, v_mnb_model_part.original_sequence_number, " & _
            "min(part_description) as part_description, sum(quantity) as quantity, " & _
            "count(quantity) as split_count " & _
            "from v_mnb_model_part " & _
            "join v_mnb_model_line on v_mnb_model_line.model_number = v_mnb_model_part.model_number and " & _
            "v_mnb_model_line.line_id = v_mnb_model_part.line_id " & _
            "join v_prod_part on v_mnb_model_part.part_id = v_prod_part.part_id " & _
            "where v_mnb_model_part.model_number = '" & cboModel.Text & "' and " & _
            "v_mnb_model_part.line_id = '" & cboLine.Text & "' " & _
            "group by v_mnb_model_part.part_id, original_sequence_number " & _
            strOrderBy
        .Open
        .Filter = "split_count > 1"
        If .RecordCount = 0 Then
            MsgBox "No Split Parts were found for this model."
            .Close
            Exit Sub
        Else
            mxarrMiniBill.LoadRows rsList.GetRows, True
        End If
        
        .Close
    End With
    'start testing here
    Dim data1 As String
    rsList.Open
    With rsList
        Do While Not .EOF
            data1 = rsList!part_id
            data1 = rsList!original_sequence_number
            data1 = rsList!part_description
            data1 = rsList!quantity
            data1 = rsList!split_count
            .MoveNext
        Loop
    End With
here:
    'end testing here
    Set rsList = Nothing
    LoadDataGrid
    
    With mrsModelPartLocation
        If .State = adStateOpen Then
            .Close
        End If
        .Source = "select * from v_mnb_model_part_stocking_location " & _
            "where model_number = '" & cboModel.Text & "' " & _
            "order by part_id"
        .Open
    End With
    mblnRecChanged = False
    TDBGMiniBill.Row = 0
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdDisplay_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdSelectPart_Click()
    ' Purpose:  To display a form allowing the user to split a part among multiple locations.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim intRetCode As Integer
    Dim rsList As ADODB.Recordset
    Dim lngTotalQuantity As Long
    Dim lngRow As Long
    
    
    If TDBGMiniBill.Row < 0 Then
        MsgBox "Select a part to be split before clicking this button."
        Exit Sub
    End If
    
    
    With frmSplitPart
        .txtLine.Text = cboLine.Text
        .txtModel.Text = cboModel.Text
        .txtPart.Text = TDBGMiniBill.Columns(0).Value
        .txtPartDescription.Text = TDBGMiniBill.Columns(2).Value
        .txtQuantity.Text = TDBGMiniBill.Columns(3).Value
        .mlngPartSequence = TDBGMiniBill.Columns(1).Value
        Set .mxarrLocations = mxarrLocations
        Set .mrsModelPartLocation = mrsModelPartLocation
        .ProcessPart
    End With
    cmdDisplay_Click
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdSplitPart_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub Form_Activate()
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    ' Purpose:  Load form data
    
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
    
    LoadLines
    cboModel.Enabled = False
    mblnRecChanged = False
    
    Set mrsModelPartLocation = New ADODB.Recordset
    
    With mrsModelPartLocation
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub LoadDataGrid()
    ' Purpose:  Load the grid with data from the recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim intIndex As Integer
    With Me.TDBGMiniBill
        .Array = mxarrMiniBill
        .ReBind
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "LoadDataGrid", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Sub RetrieveModelData()
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    
    cboModel.Clear
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select distinct model_number from v_mnb_model_line " & _
            "where line_id = '" & cboLine.Text & "' " & _
            "order by model_number asc"
        .Open
        
        Do While Not .EOF
            cboModel.AddItem Trim(!model_number)
            .MoveNext
        Loop
        .Close
    End With
    Set rsList = Nothing
    If cboModel.ListCount = 0 Then
        MsgBox "No models were found for the line selected."
        cboModel.Enabled = False
        cboLine.SetFocus
    Else
        cboModel.Enabled = True
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FindDescription", _
        Err.Number, Err.Description)
End Sub

Private Sub LoadLocations()
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    
    Set rsList = New ADODB.Recordset
    
    Set mxarrLocations = New XArrayDB
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select v_prod_stocking_location.stocking_location_id, " & _
            "stocking_location_description from v_prod_stocking_location " & _
            "join v_prod_line_stocking_location on " & _
            "v_prod_stocking_location.stocking_location_id = " & _
            "v_prod_line_stocking_location.stocking_location_id " & _
            "where line_id = '" & cboLine.Text & "' " & _
            "order by stocking_location_description"
        .Open
        
        mxarrLocations.ReDim 1, .RecordCount, 0, 1
        
        intIndex = 1
        Do While Not .EOF
            mxarrLocations(intIndex, 0) = .Fields("stocking_location_id").Value
            mxarrLocations(intIndex, 1) = Trim(!stocking_location_description)
            intIndex = intIndex + 1
            .MoveNext
        Loop
        .Close
        If mxarrLocations.UpperBound(1) = 0 Then
            MsgBox "No Locations found for the line selected."
            Set rsList = Nothing
            If Screen.ActiveForm Is Me Then
                cboLine.SetFocus
                cboLine_GotFocus
            End If
            GoTo PROC_EXIT
        End If
    End With
    Set rsList = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FindDescription", _
        Err.Number, Err.Description)
End Sub



Private Sub Form_Resize()
    Me.ElasticOne1.Refresh
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mxarrMiniBill = Nothing
    If gconDatabase.State = adStateOpen Then
        If Not mrsModelPartLocation Is Nothing Then
            If mrsModelPartLocation.State = adStateOpen Then
                mrsModelPartLocation.Close
            End If
            Set mrsModelPartLocation = Nothing
        End If
    End If
End Sub


Private Sub LoadLines()
    ' Purpose:  Load lines
    
    ' Set up error handling.
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    
    cboLine.Clear
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
'        .Source = "select line_id from v_prod_line order by line_id"
        .Source = "SELECT line_id From V_MNB_Model_Line GROUP BY line_id " & _
                "ORDER BY line_id"
        .Open
        
        Do While Not .EOF
            cboLine.AddItem !line_id
            .MoveNext
        Loop
        .Close
    End With
    
    Set rsList = Nothing
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "LoadLines", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub














Private Sub TDBGMiniBill_DblClick()
    cmdSelectPart_Click
End Sub
