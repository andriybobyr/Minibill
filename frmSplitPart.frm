VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmSplitPart 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MiniBill - Split Part To Multiple Locations"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
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
   ScaleHeight     =   422
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   701
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPartDescription 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Width           =   2715
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      CausesValidation=   0   'False
      Height          =   435
      Left            =   1860
      TabIndex        =   2
      Top             =   5700
      Width           =   1395
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      CausesValidation=   0   'False
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   5700
      Width           =   1395
   End
   Begin TDBNumber6Ctl.TDBNumber tdbQuantity 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   2580
      Visible         =   0   'False
      Width           =   1035
      _Version        =   65536
      _ExtentX        =   1826
      _ExtentY        =   556
      Calculator      =   "frmSplitPart.frx":0000
      Caption         =   "frmSplitPart.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmSplitPart.frx":008C
      Keys            =   "frmSplitPart.frx":00AA
      Spin            =   "frmSplitPart.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1179653
      Value           =   0
      MaxValueVT      =   826081285
      MinValueVT      =   1380253701
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGPartInfo 
      Height          =   3555
      Left            =   2460
      TabIndex        =   0
      Top             =   1740
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   6271
      _LayoutType     =   4
      _RowHeight      =   18
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Quantity"
      Columns(0).DataField=   ""
      Columns(0).ExternalEditor=   "tdbQuantity"
      Columns(0).ExternalEditor.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   1
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Location"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3069"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2963"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=2"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5847"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5741"
      Splits(0)._ColumnProps(9)=   "Column(1).Button=1"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3810"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3704"
      Splits(0)._ColumnProps(14)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      TabAction       =   2
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
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
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Named:id=33:Normal"
      _StyleDefs(49)  =   ":id=33,.parent=0"
      _StyleDefs(50)  =   "Named:id=34:Heading"
      _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(52)  =   ":id=34,.wraptext=-1,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(53)  =   ":id=34,.strikethrough=0,.charset=0"
      _StyleDefs(54)  =   ":id=34,.fontname=MS Sans Serif"
      _StyleDefs(55)  =   "Named:id=35:Footing"
      _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=36:Selected"
      _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=37:Caption"
      _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(61)  =   "Named:id=38:HighlightRow"
      _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=39:EvenRow"
      _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(65)  =   "Named:id=40:OddRow"
      _StyleDefs(66)  =   ":id=40,.parent=33"
      _StyleDefs(67)  =   "Named:id=41:RecordSelector"
      _StyleDefs(68)  =   ":id=41,.parent=34"
      _StyleDefs(69)  =   "Named:id=42:FilterBar"
      _StyleDefs(70)  =   ":id=42,.parent=33"
   End
   Begin VB.TextBox txtQuantity 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   780
      Width           =   1155
   End
   Begin VB.TextBox txtPart 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   720
      Width           =   1515
   End
   Begin VB.TextBox txtModel 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   180
      Width           =   2055
   End
   Begin VB.TextBox txtLine 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   375
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   435
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   435
      Left            =   8160
      TabIndex        =   4
      Top             =   5700
      Width           =   2055
   End
   Begin VB.CommandButton cmdSaveClose 
      Caption         =   "&Save And Close"
      Height          =   435
      Left            =   5760
      TabIndex        =   3
      Top             =   5700
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity:"
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
      Index           =   3
      Left            =   5820
      TabIndex        =   8
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Part:"
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
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   780
      Width           =   1155
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
      Index           =   1
      Left            =   5820
      TabIndex        =   6
      Top             =   300
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "frmSplitPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsModelPart As ADODB.Recordset
Public mrsModelPartLocation As ADODB.Recordset
Public mlngPartSequence As Long
Public mxarrLocations As XArrayDB

Private mxarrPartInfo As XArrayDB

Private Sub cmdAdd_Click()
    ' Purpose:  Add a new record
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim lngPartSeq As Long
    Dim strComments As String
    Dim strStartDate As String
    Dim strStartECN As String
    Dim strQuitEcn As String
    Dim strQuitDate As String
    Dim strLevel As String
    
    TDBGPartInfo.Update
    
    If Len(Trim(TDBGPartInfo.Columns(1))) = 0 Then
        MsgBox "Please select a location for the current record before attempting to add a new record."
        GoTo PROC_EXIT
    End If
    
    
    With mrsModelPart
       .MoveLast
       lngPartSeq = !part_sequence_number + 1
       If Not IsNull(!Comments) Then
            strComments = !Comments
        End If
        If Not IsNull(!start_ecn_date) Then
            strStartDate = !start_ecn_date
        End If
        If Not IsNull(!start_ecn_number) Then
               strStartECN = !start_ecn_number
        End If
        If Not IsNull(!quit_ecn_date) Then
            strQuitDate = !quit_ecn_date
        End If
        If Not IsNull(!quit_ecn_number) Then
               strQuitEcn = !quit_ecn_number
        End If
       strLevel = !level_number
       .AddNew
       !line_id = txtLine.Text
       !model_number = txtModel.Text
       !part_id = txtPart.Text
       !original_sequence_number = mlngPartSequence
       !quantity = 0
       !start_ecn_number = strStartECN
       !start_ecn_date = strStartDate
       !quit_ecn_number = strQuitEcn
       !quit_ecn_date = strQuitDate
       !level_number = strLevel
       !Comments = strComments
       !part_sequence_number = lngPartSeq
    End With
    
    mxarrPartInfo.AppendRows
    mxarrPartInfo(mxarrPartInfo.UpperBound(1), 2) = mrsModelPart.Fields("part_sequence_number").Value
    TDBGPartInfo.ReBind
    TDBGPartInfo.Row = mxarrPartInfo.UpperBound(1)
    TDBGPartInfo.Col = 0
    MoveToRecord
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdAdd_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdCancel_Click()
    mrsModelPart.CancelBatch
    mrsModelPartLocation.CancelBatch
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If TDBGPartInfo.Row < 0 Then
        MsgBox "Select a row to be deleted before clicking the the Delete button."
        Exit Sub
    End If
    
    If mxarrPartInfo.UpperBound(1) = 0 Then
        MsgBox "At least one row must be present."
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion + vbDefaultButton2) = _
            vbYes Then
6        mrsModelPart.Delete
        mrsModelPartLocation.Delete
        mxarrPartInfo.DeleteRows (TDBGPartInfo.Row)
        TDBGPartInfo.ReBind
        mrsModelPart.MoveFirst
        mrsModelPartLocation.MoveFirst
        TDBGPartInfo.Row = 0
        TDBGPartInfo.Col = 0
    End If
    MoveToRecord
End Sub

Private Sub cmdSaveClose_Click()
    ' Purpose:  Save data and close.
    
    ' Set up error handling.
    On Error GoTo PROC_ERR
    
    Dim intIndex As Integer
    Dim intRetCode As Integer
    Dim lngQuantity As Long
    
    TDBGPartInfo.Update
    
    lngQuantity = 0
    For intIndex = 0 To mxarrPartInfo.UpperBound(1)
        If mxarrPartInfo(intIndex, 0) = 0 Then
            MsgBox "Quantiy must be greater than 0"
            TDBGPartInfo.Row = intIndex
            TDBGPartInfo.Col = 0
            TDBGPartInfo.SetFocus
            GoTo PROC_EXIT
        Else
            lngQuantity = lngQuantity + mxarrPartInfo(intIndex, 0)
        End If
        
        If Len(Trim(mxarrPartInfo(intIndex, 1))) = 0 Then
            MsgBox "Location is required."
            TDBGPartInfo.Row = intIndex
            TDBGPartInfo.Col = 1
            TDBGPartInfo.SetFocus
            GoTo PROC_EXIT
        End If
    Next intIndex
    
    If lngQuantity <> Val(txtQuantity.Text) Then
        MsgBox "The sum of the quantities entered is " & lngQuantity & _
            ", but it should be " & txtQuantity.Text
        GoTo PROC_EXIT
        TDBGPartInfo.Row = intIndex
        TDBGPartInfo.Col = 0
        TDBGPartInfo.SetFocus
    End If
    
    mrsModelPart.UpdateBatch
    mrsModelPartLocation.UpdateBatch
    
    Unload Me
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdSaveClose_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Public Sub ProcessPart()
    ' Purpose:  Load data for the selected part.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Retrieve model/part data
    Set mrsModelPart = New ADODB.Recordset
    With mrsModelPart
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Source = "select * from v_mnb_model_part where line_id = '" & txtLine.Text & _
            "' and model_number = '" & txtModel.Text & "' order by part_sequence_number asc"
        .Open
    End With
    
    ' Load data into array for grid
    Set mxarrPartInfo = New XArrayDB
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select quantity, case when stocking_location_id is null " & _
            "then '' else stocking_location_id end as Location, " & _
            "v_mnb_model_part.part_sequence_number " & _
            "from v_mnb_model_part " & _
            "left outer join v_mnb_model_part_stocking_location " & _
            "on v_mnb_model_part.model_number = v_mnb_model_part_stocking_location.model_number and " & _
            "v_mnb_model_part.line_id = v_mnb_model_part_stocking_location.line_id and " & _
            "v_mnb_model_part.part_id = v_mnb_model_part_stocking_location.part_id and " & _
            "v_mnb_model_part.part_sequence_number = v_mnb_model_part_stocking_location.part_sequence_number " & _
            "where v_mnb_model_part.model_number = '" & txtModel.Text & _
            "' and v_mnb_model_part.line_id = '" & txtLine.Text & _
            "' and original_sequence_number = " & mlngPartSequence & _
            " and v_mnb_model_part.part_id = '" & txtPart.Text & "' " & _
            "order by v_mnb_model_part.part_sequence_number asc"
        .Open
        mxarrPartInfo.LoadRows rsList.GetRows, True
        .Close
    End With
    Set rsList = Nothing
    LoadDataGrid
    Me.Show vbModal

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "ProcessPart", Err.Number, Err.Description)
    Unload Me
    GoTo PROC_EXIT
End Sub

Private Sub LoadDataGrid()
    ' Purpose:  Load the grid with data from the recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim intIndex As Integer
    Dim Item As TrueOleDBGrid70.ValueItem
    Set Item = New TrueOleDBGrid70.ValueItem
    
    With Me.TDBGPartInfo
        .Array = mxarrPartInfo
        .ReBind
        
        With .Columns(1)
            If .ValueItems.Count = 0 Then
                .AutoCompletion = True
                .ButtonAlways = True
                With .ValueItems
                    For intIndex = 1 To mxarrLocations.UpperBound(1)
                        Item.Value = mxarrLocations.Value(intIndex, 0)
                        Item.DisplayValue = mxarrLocations.Value(intIndex, 1)
                        .Add Item
                    Next intIndex
                    .Translate = True
                End With
            End If
        End With
        
    End With
    If Screen.ActiveForm Is Me Then
        MoveToRecord
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "LoadDataGrid", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub




Private Sub Form_Unload(Cancel As Integer)
    If mrsModelPart.State = adStateOpen Then
        mrsModelPart.Close
    End If
    Set mrsModelPart = Nothing
    
End Sub

Private Sub TDBGPartInfo_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
        mrsModelPart!quantity = TDBGPartInfo.Columns(ColIndex).Value
    ElseIf ColIndex = 1 Then
        mrsModelPartLocation!stocking_location_id = TDBGPartInfo.Columns(ColIndex).Value
    End If
    
End Sub

Private Sub TDBGPartInfo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    MoveToRecord
End Sub


Sub MoveToRecord()
    ' Purpose:  Find records to go with row to which we are being moved.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    With mrsModelPart
        .MoveFirst
        .Find "part_sequence_number = " & mxarrPartInfo(TDBGPartInfo.RowBookmark(TDBGPartInfo.Row), 2)
        Do While Not .EOF
            If Trim(!part_id) = Trim(txtPart.Text) Then
                Exit Do
            Else
                .MoveNext
            End If
            .Find "part_sequence_number = " & mxarrPartInfo(TDBGPartInfo.RowBookmark(TDBGPartInfo.Row), 2)
        Loop
    End With
    
    
    With mrsModelPartLocation
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "part_sequence_number = " & mxarrPartInfo(TDBGPartInfo.RowBookmark(TDBGPartInfo.Row), 2)
            Do While Not .EOF
                If Trim(!part_id) = Trim(txtPart.Text) Then
                    Exit Do
                Else
                    .MoveNext
                End If
            Loop
        End If
        
        If .EOF Or .RecordCount = 0 Then
            gconDatabase.Errors.Clear
            .AddNew
            !line_id = mrsModelPart!line_id
            !model_number = mrsModelPart!model_number
            !part_id = mrsModelPart!part_id
            !part_sequence_number = mrsModelPart!part_sequence_number
        End If
    End With

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "MoveToRecord", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub
