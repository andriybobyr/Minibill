VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmModelsByPart 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Minbill - Models by Part..."
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All Models"
      Height          =   615
      Left            =   8640
      TabIndex        =   11
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   435
      Left            =   6720
      TabIndex        =   7
      Top             =   6300
      Width           =   1635
   End
   Begin VB.ComboBox cboLocation 
      Height          =   315
      Left            =   5280
      TabIndex        =   5
      Top             =   720
      Width           =   3015
   End
   Begin VB.ComboBox cboLine 
      Height          =   315
      Left            =   1620
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.ComboBox cboPart 
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Top             =   180
      Width           =   1755
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   8880
      TabIndex        =   9
      Top             =   6300
      Width           =   1335
   End
   Begin TrueOleDBGrid70.TDBGrid TDBgModelInfo 
      Height          =   4515
      Left            =   240
      TabIndex        =   6
      Top             =   1500
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7964
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Model Number"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Description"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Line"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Quantity"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Schedule Date"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Short Date"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Location"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   4
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Use Selected Loc"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2196"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2117"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=4260"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4180"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=767"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=688"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=1931"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=1852"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(25)=   "Column(4).Width=2064"
      Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=1984"
      Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(29)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(31)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=1"
      Splits(0)._ColumnProps(42)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
      _StyleDefs(64)  =   "Named:id=33:Normal"
      _StyleDefs(65)  =   ":id=33,.parent=0"
      _StyleDefs(66)  =   "Named:id=34:Heading"
      _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(68)  =   ":id=34,.wraptext=-1"
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
   Begin VB.Label lblPartDescription 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      Top             =   180
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3780
      TabIndex        =   10
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Index           =   1
      Left            =   4020
      TabIndex        =   8
      Top             =   780
      Width           =   1095
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
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Top             =   780
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Part Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "frmModelsByPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mxarrModelInfo As XArrayDB
Private mxarrPartInfo As XArrayDB
Private mxarrLocations As XArrayDB

Public mstrPart As String
Public mstrPartDescription As String

Private mrsModelPartLocation As ADODB.Recordset
Private rsExistingStockLoc As ADODB.Recordset

Private mblnFromNewParts As Boolean
Private mblnRecChanged As Boolean
Private mblnSaved As Boolean
Private mblnSelectionsOK As Boolean

Private Sub cboLine_Change()
    If Len(cboLine.Text) = 2 Then
        cboFindFirst cboLine
    End If
End Sub
Private Sub cboLine_Click()
    Call ReQueryModelsForShiftChange
    Call LoadLocations
End Sub

Private Sub cboLine_GotFocus()
    cboLine.SelStart = 0
    cboLine.SelLength = Len(cboLine.Text)
End Sub
Private Sub cboLine_KeyPress(KeyAscii As Integer)
    cboLine.SelStart = 0
    cboLine.SelLength = Len(cboLine.Text)
End Sub
Private Sub cboLine_Validate(Cancel As Boolean)
    ' Purpose:  Validate the line entered
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim intRetCode As Integer
    
    If mblnRecChanged Then
        intRetCode = MsgBox("Save Changes?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Data Changed")
        If intRetCode = vbYes Then
            cmdSave_Click
        ElseIf intRetCode = vbNo Then
            mrsModelPartLocation.CancelBatch
        Else
            cboLine.Text = mrsModelPartLocation!line_id
            Cancel = True
            Exit Sub
        End If
    End If
    
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
        
        If mblnFromNewParts Then
            If mxarrLocations.UpperBound(1) = 0 Then
                Cancel = True
'                cboLine_GotFocus
            Else
                Cancel = False
            End If
        End If
    End If
    
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

Private Sub cboLocation_KeyPress(KeyAscii As Integer)
    cboKeyPress cboLocation, KeyAscii
End Sub

Private Sub cboPart_Change()
    cboFindFirst cboPart
    lblPartDescription = " "
End Sub

Private Sub cboPart_GotFocus()
    cboPart.SelStart = 0
    cboPart.SelLength = Len(cboPart.Text)
End Sub
Private Sub cboPart_KeyPress(KeyAscii As Integer)
    cboKeyPress cboPart, KeyAscii
End Sub

Private Sub cboPart_Validate(Cancel As Boolean)
    Dim intRetCode As Integer
    
    'Check for changes from old Part on dropdown
    If mblnRecChanged Then
        intRetCode = MsgBox("Save Changes?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Data Changed")
        If intRetCode = vbYes Then
            cmdSave_Click
        ElseIf intRetCode = vbNo Then
            mrsModelPartLocation.CancelBatch
        Else
            cboPart.Text = mrsModelPartLocation!part_id
            Cancel = True
            Exit Sub
        End If
    End If
    
    frmModelsByPart.MousePointer = vbHourglass
    
    'Check for part validation
    If Len(Trim(cboPart.Text)) = 0 Then
        MsgBox "Part is required."
        Cancel = True
        Exit Sub
    End If
    
    cboFindFirst cboPart
    If cboPart.ListIndex = -1 Then
        MsgBox "Invalid part entered."
        Cancel = True
        Exit Sub
    End If
    Cancel = False
    
    'Found part, so fill the description
    lblPartDescription.Caption = mxarrPartInfo(cboPart.ListIndex, 1)
    
    'Part dropdown has changed, update the fields on the screen
    Call DisplayModelsForPart

    frmModelsByPart.MousePointer = vbDefault
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    frmModelsByPart.MousePointer = vbDefault
    Call ShowError(Me.Name, "cboPart_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    ' Purpose:  Save current changes.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim Row As Long
    mblnSaved = False
    
    With gconDatabase
        mrsModelPartLocation.UpdateBatch
        If .Errors.Count > 0 Then
            Err.Raise .Errors(0).NativeError, "cmdSave_click", .Errors(0).Description
        End If
    End With
    
    MsgBox "Records successfully saved."
    cmdSelectAll.Caption = "Select All Models"
    mblnRecChanged = False
    mblnSaved = True

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdSave_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdSelectAll_Click()

    ' Purpose:  Mass change all Model/parts to have this Stocking Location.
Dim Row As Long
Dim intRetCode As Integer

    ' Set up error handling
    On Error GoTo PROC_ERR
    
'Using button as a toggle to either select all the models or deselect all the models.
    mblnSelectionsOK = True
    
' Puts cursor at the top of the grid to change all data.  If not, the grid will be
'   updated at the point where the cursor is.
    Row = 0
    TDBgModelInfo.Bookmark = 0
    TDBgModelInfo.MoveFirst


    If cmdSelectAll.Caption = "Select All Models" Then
        'Update array checkbox for all items on the screen.
        '  Then, Refresh array into grid, and it will appear on the screen.
        intRetCode = MsgBox("This will change existing Stocking Locations" & vbCrLf & _
            "Click OK to continue, or Cancel to exit", vbQuestion + vbOKCancel + vbDefaultButton1, _
             "Change all Locations? ...")
        If intRetCode = vbCancel Then
            Exit Sub
        Else
        With TDBgModelInfo
            With .Array
                For Row = .LowerBound(1) To .UpperBound(1)
                    .Value(Row, 6) = 1
                    TDBgModelInfo.Columns(6).Value = -1
                    Call TDBgModelInfo_AfterColUpdate(6)
                    If Not mblnSelectionsOK Then
                        Exit Sub
                    End If
                    TDBgModelInfo.MoveNext
                Next Row
            End With
            cmdSelectAll.Caption = "Deselect All " & vbCrLf & "    Models"
            .Refresh
            .SetFocus
        End With
        End If
    Else
        'This will take off the selections, and cancel changes.  If the user then hits the
        ' Cancel button, then everything will be restored back to the beginning.
        'Once the SAVE button is pressed, old stocking locations are gone.
        intRetCode = MsgBox("This will delete all Stocking Locations" & vbCrLf & _
            "Click OK to continue, or Cancel to exit", vbQuestion + vbOKCancel + vbDefaultButton1, _
             "Cancel Location Change? ...")
        If intRetCode = vbCancel Then
            Exit Sub
        Else
        With TDBgModelInfo
            With .Array
                For Row = .LowerBound(1) To .UpperBound(1)
                    .Value(Row, 6) = 0
                    TDBgModelInfo.Columns(6).Value = 0
                    Call TDBgModelInfo_AfterColUpdate(6)
                    If Not mblnSelectionsOK Then
                        Exit Sub
                    End If
                    TDBgModelInfo.MoveNext
                Next Row
            End With
            cmdSelectAll.Caption = "Select All Models"
            .Refresh
            .SetFocus
        End With
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdSelectAll_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub Form_Load()
    ' Purpose:  Complete necessary elements to load the form
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set up connection
    If gconDatabase Is Nothing Then
        Set gconDatabase = gclsSQLServer.Connect( _
            gclsMESApplication.ApplicationRole, _
            gclsMESApplication.ApplicationPassword)
        
        If gconDatabase.State <> adStateOpen Then
            Err.Raise gconDatabase.Errors(0).NativeError, "Form_Main", _
                gconDatabase.Errors(0).Description
        End If
    End If
    
    gconDatabase.CommandTimeout = 350
    
    If Screen.ActiveForm Is Me Then
        frmProcessing.Label1 = "Gathering Models... Please Wait"
        frmProcessing.Label2 = "Get Configuration Models List at " & Time
        frmProcessing.Show
    End If
    
    DoEvents
    
    mblnRecChanged = False
    
    ' If this form was not loaded by the New Parts Review form, then load the part combo
    ' box.
    If Not Screen.ActiveForm Is frmNewPartsReview Then
        Call LoadParts
        mblnFromNewParts = False
    Else
        mblnFromNewParts = True
    End If
    
    If Screen.ActiveForm Is Me Then
        frmProcessing.Hide
    End If
    
    DoEvents
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    frmModelsByPart.MousePointer = vbDefault
    Call ShowError(Me.Name, "Form_Load", Err.Number, Err.Description)
    Unload Me
    GoTo PROC_EXIT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim intRetCode As Integer
    
    If mblnRecChanged Then
        intRetCode = MsgBox("Save Changes?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Data Changed")
        If intRetCode = vbYes Then
            cmdSave_Click
        ElseIf intRetCode = vbNo Then
            mrsModelPartLocation.CancelBatch
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mxarrModelInfo = Nothing
    mstrPart = vbNullString
    
    If gconDatabase.State = adStateOpen Then
        If Not mrsModelPartLocation Is Nothing Then
            If mrsModelPartLocation.State = adStateOpen Then
                mrsModelPartLocation.Close
            End If
            Set mrsModelPartLocation = Nothing
        End If
        
    End If
End Sub

Public Sub DisplayModelsForPart()
    ' Purpose:  This procedure takes a part that is sent from another form and builds the grid
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim intIndex As Integer
    
    ' Set part fields
    If Screen.ActiveForm Is frmNewPartsReview Then
        cboPart.Clear
        cboPart.AddItem mstrPart
        cboPart.Text = mstrPart
        cboPart.ListIndex = 0
        cboPart.Enabled = False
        lblPartDescription.Caption = mstrPartDescription
    End If
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        
        If mblnFromNewParts Then
'        Temporarily changing the query to include all new parts
            .Source = "select distinct line_id " & _
                "from v_mnb_model_part " & _
                "where part_id = '" & cboPart.Text & "' and part_create_date is not null and part_reviewed_flag = 0 " & _
                "and line_id = '" & gblnLine & "' order by line_id"
         Else
            .Source = "select distinct line_id " & _
                "from v_mnb_model_part " & _
                "where part_id = '" & cboPart.Text & "' " & _
                "order by line_id"
        End If
        .Open

        cboLine.Clear
        Do While Not .EOF
            cboLine.AddItem !line_id
            .MoveNext
        Loop
        If cboLine.ListCount > 0 Then
            cboLine.ListIndex = 0
        End If
        .Close
      
    End With

    mblnRecChanged = False
    
PROC_EXIT:
            
    Exit Sub
    
PROC_ERR:

    Call ShowError(Me.Name, "DisplayModelsForPart", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub LoadLocations()
    ' Purpose:  Fill the location combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    
    Set rsList = New ADODB.Recordset
    
    Set mxarrLocations = New XArrayDB
    cboLocation.Clear
    
    'Create Stocking Locations List Dropdown for a particular Line
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
            cboLocation.AddItem mxarrLocations(intIndex, 1)
            intIndex = intIndex + 1
            .MoveNext
        Loop
        .Close
        If mxarrLocations.UpperBound(1) = 0 Then
            Set rsList = Nothing
            If Screen.ActiveForm Is Me Then
                MsgBox "No Locations found for the line selected."
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


Sub LoadParts()
    ' Purpose:  Load all default parts to the dropdown list.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Source = "select v_mnb_model_part.part_id, min(rtrim(part_description)) as part_description " & _
            "from v_mnb_model_part " & _
            "join v_prod_part on v_mnb_model_part.part_id = v_prod_part.part_id " & _
            "where v_mnb_model_part.inactive_part_flag = '0' " & _
            "or v_mnb_model_part.start_ecn_date >= '" & gstrCurrentDate & "'" & _
            " group by v_mnb_model_part.part_id " & _
            "order by v_mnb_model_part.part_id "
        .Open
        
        Set mxarrPartInfo = New XArrayDB
        mxarrPartInfo.LoadRows .GetRows, True
        .Close
        cboPart.Clear
                
        Dim lngIndex As Integer
        For lngIndex = 0 To mxarrPartInfo.UpperBound(1)
            cboPart.AddItem mxarrPartInfo(lngIndex, 0)
        Next lngIndex
    End With
    Set rsList = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "LoadParts", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub TDBgModelInfo_AfterColUpdate(ByVal ColIndex As Integer)
    
    If TDBgModelInfo.Columns(2).Value <> cboLine.Text Then
        MsgBox "The Line on this record does not match the Line selected above."
        mblnSelectionsOK = False
        If TDBgModelInfo.Columns(6).Value <> 0 Then
            TDBgModelInfo.Columns(6).Value = 0
        End If
        Exit Sub
    End If
    
    If cboLocation.ListIndex = -1 Then
        MsgBox "Location must be specified before changes may be made."
        mblnSelectionsOK = False
        If TDBgModelInfo.Columns(6).Value <> 0 Then
            TDBgModelInfo.Columns(6).Value = 0
        End If
        cboLocation.SetFocus
        Exit Sub
    End If
    '*******************************************************************
           'penny start :  Why dont the updates work here???
    '********************************************************************
        'Sometimes inactive Stocking Locations are still in the system.  Need to check if the inactive
        'stocking location is the one being added now.
'        Dim rsExistingStockLoc As ADODB.Recordset
        Set rsExistingStockLoc = New ADODB.Recordset
        With rsExistingStockLoc
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Source = "SELECT model_number,line_id,part_id,part_sequence_number,stocking_location_id," & _
                   "step_number,model_part_stocking_location_obsolete_flag, " & _
                   "model_part_stocking_location_obsolete_date From V_MNB_Model_Part_Stocking_Location_All " & _
                   "where part_id = '" & cboPart.Text & "' and model_number = '" & _
                    TDBgModelInfo.Columns(0).Value & "' and line_id= '" & cboLine.Text & _
                    "' and stocking_location_id = '" & mxarrLocations(cboLocation.ListIndex + 1, 0) & _
                    "'and model_part_stocking_location_obsolete_flag = 1"
            .Open
        
                If .RecordCount > 0 Then
                    !model_part_stocking_location_obsolete_flag = 0
                    !model_part_stocking_location_obsolete_date = vbNull
                    !part_sequence_number = mxarrModelInfo(TDBgModelInfo.Bookmark, 7)
                    .Update
  '                  .Delete
                
                End If
             .Close
            
        End With
            
           
  '      Set rsExistingStockLoc = Nothing
        
        
        'penny end
    
    
    
    
    'Check Model number to see if an active location already exists for that part.
    mblnRecChanged = True
    With mrsModelPartLocation
        .Requery
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                .Find "model_number = '" & TDBgModelInfo.Columns(0).Value & "'"
                If Not .EOF Then
                    If !part_sequence_number = mxarrModelInfo(TDBgModelInfo.Bookmark, 7) And _
                            !line_id = cboLine.Text Then
                        Exit Do
                    End If
                End If
                If Not .EOF Then
                    .MoveNext
                End If
            Loop
        End If
        
        'If user changes mind and takes the check off, delete from the recordset.
        If TDBgModelInfo.Columns(6).Value = 0 Then
            .Delete
            TDBgModelInfo.Columns(5).Value = vbNullString
            Exit Sub
        End If
        
        'penny start
        'Sometimes inactive Stocking Locations are still in the system.  Need to check if the inactive
        'stocking location is the one being added now.
'        Dim rsExistingStockLoc As ADODB.Recordset
'        Set rsExistingStockLoc = New ADODB.Recordset
'        With rsExistingStockLoc
'            Set .ActiveConnection = gconDatabase
'            .CursorLocation = adUseClient
'            .CursorType = adOpenStatic
'            .LockType = adLockBatchOptimistic
'            .Source = "SELECT model_number,line_id,part_id,part_sequence_number,stocking_location_id," & _
'                   "step_number,model_part_stocking_location_obsolete_flag, " & _
'                   "model_part_stocking_location_obsolete_date From V_MNB_Model_Part_Stocking_Location_All " & _
'                   "where part_id = '" & cboPart.Text & "' and model_number = '" & _
'                    TDBgModelInfo.Columns(0).Value & "' and line_id= '" & cboLine.Text & _
'                    "' and model_part_stocking_location_obsolete_flag = 1"
'            .Open
'
'                If .RecordCount > 0 Then
'                    !model_part_stocking_location_obsolete_flag = 0
'                    !model_part_stocking_location_obsolete_date = vbNull
'                    !part_sequence_number = mxarrModelInfo(TDBgModelInfo.Bookmark, 7)
'  '                  .Delete
'
'                Else
'
'
'   '         End With
'
'   '         Set rsExistingStockLoc = Nothing
'
'
'        'penny end
'
'
        'If that Model does not have a Stocking Location for the part, add it to the list.
                If .EOF Or .RecordCount = 0 Then
                    .AddNew
                    !model_number = TDBgModelInfo.Columns(0).Value
                    !line_id = TDBgModelInfo.Columns(2).Value
                    !part_id = cboPart.Text
                    !part_sequence_number = mxarrModelInfo(TDBgModelInfo.Bookmark, 7)
                End If
        
                !stocking_location_id = mxarrLocations(cboLocation.ListIndex + 1, 0)
                TDBgModelInfo.Columns(5).Value = cboLocation.Text
        
'            End With
'            Set rsExistingStockLoc = Nothing
        
        
    End With
    
End Sub

Sub ReQueryModelsForShiftChange()
' Purpose:  Requery model grid when shift changes.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
 
 Dim intRetCode As Integer
      
 'Check to save changes before getting a different shift
 If mblnRecChanged Then
        intRetCode = MsgBox("Save Changes?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Data Changed")
        If intRetCode = vbYes Then
            cmdSave_Click
        ElseIf intRetCode = vbNo Then
            mrsModelPartLocation.CancelBatch
        Else
            cboLine.Text = mrsModelPartLocation!line_id
'            Cancel = True
            Exit Sub
        End If
    End If
  
 Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
  
  If mblnFromNewParts Then
'        Temporarily changing the query to include all new parts
            .Source = "select v_mnb_model_line.model_number, " & _
                "min(rtrim(v_prod_model_all.sales_model_description)) as Description, " & _
                "v_mnb_model_part.line_id, min(v_mnb_model_part.quantity), min(v_prod_schedule.start_date), " & _
                "min(rtrim(v_prod_stocking_location.stocking_location_description)) as Location, " & _
                "0 as set_from_location, v_mnb_model_part.part_sequence_number " & _
                "from v_mnb_model_part " & _
                "join v_mnb_model_line on v_mnb_model_part.model_number = v_mnb_model_line.model_number " & _
                "and v_mnb_model_part.line_id = v_mnb_model_line.line_id " & _
                "left outer join v_prod_model_line_all on v_mnb_model_part.model_number = v_prod_model_line_all.engineering_model_number " & _
                "and v_mnb_model_part.line_id = v_prod_model_line_all.line_id " & _
                "left outer join v_prod_model_all on v_prod_model_line_all.sales_model_number = v_prod_model_all.sales_model_number " & _
                "left outer join v_prod_schedule on v_mnb_model_part.model_number = v_prod_schedule.model_number and v_mnb_model_part.line_id = v_prod_schedule.line_id " & _
                "and v_mnb_model_part.line_id = v_prod_schedule.line_id " & _
                "left outer join v_mnb_model_part_stocking_location on " & _
                "v_mnb_model_part.line_id = v_mnb_model_part_stocking_location.line_id and " & _
                "v_mnb_model_part.model_number = v_mnb_model_part_stocking_location.model_number  and " & _
                "v_mnb_model_part.part_id = v_mnb_model_part_stocking_location.part_id and " & _
                "v_mnb_model_part.part_sequence_number = v_mnb_model_part_stocking_location.part_sequence_number " & _
                "left outer join v_prod_stocking_location on v_mnb_model_part_stocking_location.stocking_location_id = " & _
                "v_prod_stocking_location.stocking_location_id " & _
                "where v_mnb_model_part.line_id = '" & cboLine.Text & "' and " & _
                "v_mnb_model_part.part_id = '" & cboPart.Text & "' and part_create_date is not null " & _
                "group by v_mnb_model_line.model_number, v_mnb_model_part.line_id, v_mnb_model_part.part_sequence_number " & _
                "order by v_mnb_model_part.line_id, v_mnb_model_line.model_number"
        Else
            .Source = "select v_mnb_model_line.model_number, " & _
                "min(rtrim(v_prod_model.sales_model_description)) as Description, " & _
                "v_mnb_model_part.line_id, min(v_mnb_model_part.quantity), min(v_prod_schedule.start_date), " & _
                "min(rtrim(v_prod_stocking_location.stocking_location_description)) as Location, " & _
                "0 as set_from_location, v_mnb_model_part.part_sequence_number " & _
                "from v_mnb_model_part " & _
                "left outer join v_prod_model_line on v_mnb_model_part.model_number = v_prod_model_line.engineering_model_number " & _
                "and v_mnb_model_part.line_id = v_prod_model_line.line_id " & _
                "left outer join v_prod_model on v_prod_model_line.sales_model_number = v_prod_model.sales_model_number " & _
                "left outer join v_mnb_model_line on v_mnb_model_part.model_number = v_mnb_model_line.model_number and v_mnb_model_part.line_id = v_mnb_model_line.line_id " & _
                "left outer join v_prod_schedule on v_mnb_model_part.model_number = v_prod_schedule.model_number and v_mnb_model_part.line_id = v_prod_schedule.line_id " & _
                "and v_mnb_model_part.line_id = v_prod_schedule.line_id " & _
                "left outer join v_mnb_model_part_stocking_location on " & _
                "v_mnb_model_part.line_id = v_mnb_model_part_stocking_location.line_id and " & _
                "v_mnb_model_part.model_number = v_mnb_model_part_stocking_location.model_number  and " & _
                "v_mnb_model_part.part_id = v_mnb_model_part_stocking_location.part_id and " & _
                "v_mnb_model_part.part_sequence_number = v_mnb_model_part_stocking_location.part_sequence_number " & _
                "left outer join v_prod_stocking_location on v_mnb_model_part_stocking_location.stocking_location_id = " & _
                "v_prod_stocking_location.stocking_location_id " & _
                "where v_mnb_model_part.part_id = '" & cboPart.Text & "' and " & _
                "v_mnb_model_part.line_id = '" & cboLine.Text & "' " & _
                "and v_mnb_model_line.model_line_inactive_flag = 0 " & _
                "group by v_mnb_model_line.model_number, v_mnb_model_part.line_id, v_mnb_model_part.part_sequence_number " & _
                "order by v_mnb_model_part.line_id, v_mnb_model_line.model_number"
        End If
        .Open
        
        'Notify user when no models are returned
        If .RecordCount = 0 Then
            Set mxarrModelInfo = New XArrayDB
            mxarrModelInfo.Clear
            TDBgModelInfo.Array = mxarrModelInfo
            TDBgModelInfo.ReBind
            MsgBox ("No Active models on this line " & vbCrLf & " for part " & _
                cboPart.Text)
            Exit Sub
        End If
        
        'Load Xarray used by Grid
        Set mxarrModelInfo = New XArrayDB
        mxarrModelInfo.LoadRows rsList.GetRows, True
        TDBgModelInfo.Array = mxarrModelInfo
        TDBgModelInfo.ReBind
        TDBgModelInfo.Scroll 0, TDBgModelInfo.Bookmark * -1
        TDBgModelInfo.Row = 0
        TDBgModelInfo.Columns(6).AllowFocus = True
        .Close
   End With
   
   'Get all existing default models setup for this part.
   Set mrsModelPartLocation = New ADODB.Recordset
   With mrsModelPartLocation
       Set .ActiveConnection = gconDatabase
       .CursorLocation = adUseClient
       .CursorType = adOpenStatic
       .LockType = adLockBatchOptimistic
       .Source = "select * from v_mnb_model_part_stocking_location " & _
           "where part_id = '" & cboPart.Text & "' " & _
           "order by model_number, line_id, part_sequence_number "
       .Open
    End With
    
    mblnRecChanged = False

PROC_EXIT:
            
    Exit Sub
    
PROC_ERR:

    Call ShowError(Me.Name, "ReQueryModelsForShiftChange", Err.Number, Err.Description)
    GoTo PROC_EXIT
   
End Sub

