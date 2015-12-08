VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTemporaryNewPart 
   Caption         =   "MiniBill - Temporary New Part"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10935
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
   ScaleHeight     =   614
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   729
   StartUpPosition =   1  'CenterOwner
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   9210
      Left            =   0
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   0
      Width           =   10935
      _cx             =   19288
      _cy             =   16245
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   0
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
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin VB.CommandButton cmdList 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   3000
         Picture         =   "frmTemporaryNewPart.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Show Listing"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   840
         Picture         =   "frmTemporaryNewPart.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Delete This Record"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdAddNew 
         CausesValidation=   0   'False
         Height          =   285
         Left            =   120
         Picture         =   "frmTemporaryNewPart.frx":0314
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Add new entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.TextBox txtPart 
         Height          =   360
         Left            =   420
         MaxLength       =   15
         TabIndex        =   13
         Top             =   4380
         Width           =   2850
      End
      Begin VB.CommandButton cmdFirst 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   1380
         Picture         =   "frmTemporaryNewPart.frx":0936
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Display First Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdPrevious 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   1740
         Picture         =   "frmTemporaryNewPart.frx":0F10
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Display Previous Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdNext 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   2100
         Picture         =   "frmTemporaryNewPart.frx":14EA
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Display Next Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdLast 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   2460
         Picture         =   "frmTemporaryNewPart.frx":1AC4
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Display Last Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdClose 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   3720
         Picture         =   "frmTemporaryNewPart.frx":209E
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Close This Form"
         Top             =   180
         Width           =   300
      End
      Begin VB.ListBox lstAvailable 
         Height          =   3420
         Left            =   420
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   4935
         Width           =   3780
      End
      Begin VB.ListBox lstSelected 
         Height          =   3420
         Left            =   6360
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   4920
         Width           =   4140
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add >"
         Height          =   360
         Left            =   4560
         TabIndex        =   15
         Top             =   5550
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "< &Remove"
         Height          =   360
         Left            =   4560
         TabIndex        =   16
         Top             =   6150
         Width           =   1455
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "<< &Remove All"
         Height          =   360
         Left            =   4560
         TabIndex        =   17
         Top             =   6825
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   480
         Picture         =   "frmTemporaryNewPart.frx":2800
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Save Changes"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdHelp 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   3360
         Picture         =   "frmTemporaryNewPart.frx":2DAA
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Help"
         Top             =   180
         Width           =   300
      End
      Begin MSComctlLib.StatusBar staDBMaint 
         Height          =   390
         Left            =   480
         TabIndex        =   27
         Top             =   8640
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               Alignment       =   1
               TextSave        =   "11/2/2007"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               Alignment       =   1
               TextSave        =   "11:03 AM"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   10292
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.Frame fraSelection 
         Caption         =   "Pick an Option from the Three Choices:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   10275
         Begin VB.TextBox txtStepNumber 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   9240
            MaxLength       =   3
            TabIndex        =   9
            Top             =   1920
            Width           =   735
         End
         Begin VB.ComboBox cboStockingLocation 
            Height          =   360
            Left            =   1920
            TabIndex        =   10
            Top             =   2520
            Width           =   3135
         End
         Begin VB.CommandButton cmdTempEndDate 
            CausesValidation=   0   'False
            Height          =   435
            Left            =   7440
            Picture         =   "frmTemporaryNewPart.frx":2EAC
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1920
            Width           =   450
         End
         Begin VB.TextBox txtPartQuantity 
            Height          =   375
            Left            =   1080
            MaxLength       =   5
            TabIndex        =   5
            Top             =   1920
            Width           =   735
         End
         Begin VB.CommandButton cmdRefreshModels 
            Caption         =   "Refresh Models"
            Height          =   375
            Left            =   6960
            TabIndex        =   11
            ToolTipText     =   "Show Models on Stocking Location Line"
            Top             =   2520
            Width           =   1575
         End
         Begin MSMask.MaskEdBox mskECNNumber 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   375
            Left            =   3480
            TabIndex        =   6
            Top             =   1920
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   "_"
         End
         Begin VB.OptionButton optReplacePart 
            Caption         =   "Replace a Part"
            Height          =   255
            Left            =   5760
            TabIndex        =   2
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton optInactive 
            Caption         =   "Make Part Inactive"
            Height          =   255
            Left            =   3300
            TabIndex        =   1
            Top             =   480
            Width           =   2175
         End
         Begin VB.OptionButton optAdd 
            Caption         =   "Add New Part"
            Height          =   255
            Left            =   840
            TabIndex        =   0
            Top             =   480
            Width           =   2175
         End
         Begin VB.ComboBox cboNewPart 
            Height          =   360
            Left            =   2160
            TabIndex        =   3
            Top             =   840
            Width           =   3015
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print"
            Height          =   375
            Left            =   8880
            TabIndex        =   12
            ToolTipText     =   "Print Models Affected by this Part"
            Top             =   2520
            Width           =   735
         End
         Begin VB.ComboBox cboReplacePart 
            Height          =   360
            Left            =   2160
            TabIndex        =   4
            Top             =   1320
            Width           =   3015
         End
         Begin MSMask.MaskEdBox mskTempEndDate 
            Height          =   375
            Left            =   6120
            TabIndex        =   7
            Top             =   1920
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Format          =   "mm/dd/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblStepNumber 
            Caption         =   "Step Number:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   8160
            TabIndex        =   43
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label lblStockingLocation 
            Caption         =   "Stocking Location:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   480
            TabIndex        =   40
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Temporary End Date:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4800
            TabIndex        =   39
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Part Quantity:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   37
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "ECN Number:"
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
            Left            =   2040
            TabIndex        =   36
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label lblReplacePartDesc 
            Height          =   375
            Left            =   5400
            TabIndex        =   35
            Top             =   1320
            Width           =   4455
         End
         Begin VB.Label lblNewPartDesc 
            Height          =   375
            Left            =   5400
            TabIndex        =   34
            Top             =   840
            Width           =   4455
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Replaces Part:"
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
            Left            =   240
            TabIndex        =   33
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Part Selection:"
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
            Left            =   240
            TabIndex        =   32
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Available Models:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   30
         Top             =   4080
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Models"
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
         Left            =   6480
         TabIndex        =   29
         Top             =   4380
         Width           =   1995
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   1
         X1              =   480
         X2              =   10440
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   0
         X1              =   480
         X2              =   10440
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Whirlpool MES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6240
         TabIndex        =   28
         Top             =   120
         Width           =   2475
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewFirst 
         Caption         =   "&First Entry"
      End
      Begin VB.Menu mnuViewPrevious 
         Caption         =   "&Previous Entry"
      End
      Begin VB.Menu mnuViewNext 
         Caption         =   "&Next Entry"
      End
      Begin VB.Menu mnuViewLast 
         Caption         =   "&Last Entry"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuViewList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "&Index..."
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmTemporaryNewPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents mrsDatabase As ADODB.Recordset
Attribute mrsDatabase.VB_VarHelpID = -1
Public WithEvents mrsTempModelPartsMaster As ADODB.Recordset
Attribute mrsTempModelPartsMaster.VB_VarHelpID = -1
Public WithEvents mrsTempModelPartsDetail As ADODB.Recordset
Attribute mrsTempModelPartsDetail.VB_VarHelpID = -1
Public WithEvents mrsModelPartStockLoc As ADODB.Recordset
Attribute mrsModelPartStockLoc.VB_VarHelpID = -1

Private blnCancel As Boolean
Private mblnRecChanged As Boolean
Private mblnAddTempRec As Boolean
Private mblnLineChanged As Boolean
Private mblnRetrieveModels As Boolean
Private blnSaveMasterInfo As Boolean
Private blnDeleteRecord As Boolean
Private blnEmptyTempFile As Boolean
Private marrstrActiveParts() As String
Private marrstrAllParts() As String
Private marrstrStockingLocationId() As String
Private marrstrStockingLocationDesc() As String
Private mblnLoadedAllParts As Boolean
Private marrstrpartDescription() As String
Private arrstrCategory() As String
Private strPartId As String
Private strModelNumber As String
Private strPartDescription As String
Private strStockingLocationID As String
Private strStockingLocationDesc As String
Private mstrCategory As String
Private strActivityCode As String
Private strLineId As String
Private intArrayCount As Integer
Private intNewPartIndex As Integer
Private varBookMarkKeep As Variant

Dim dteToday As Date
Dim dteWeekFuture As Date
Dim dteWeekTwoFuture As Date
Dim dtePicked As Date

Private strDisplay As String
Private Sub cboNewPart_Change()
    cboFindFirst cboNewPart
End Sub
Private Sub cboNewPart_Click()
 ' Set error handling
    On Error GoTo PROC_ERR
   
    Dim varBookmark As Variant
    
    'Bypasses this edit when there are no Temporary records in the data base.
    With mrsTempModelPartsMaster
        If .RecordCount = 0 Then
            GoTo PROC_EXIT
        End If
    End With

' If the Part has changed, find it in the recordset.  Verify there isn't already
' a setup for the part.  If none found, keep current record showing.  If found,
' make the part found the current record.
    If cboNewPart.Visible Then
        With mrsTempModelPartsMaster
            If RTrim(!part_id) <> cboNewPart.Text Then
                varBookmark = .Bookmark
                .MoveFirst
                .Find "part_id = '" & cboNewPart.Text & "'", , adSearchForward, 0
                If .EOF Then
                    strDisplay = .RecordCount
                    .Bookmark = varBookmark
                Else
                    Call CheckActivityCode
                    Call MoveTempModelPartsMaster_MoveComplete
                    Call cmdRefreshModels_Click
                End If
            End If
        End With
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboNewPart_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub
Private Sub cboNewPart_GotFocus()
  cboNewPart.SelStart = 0
  cboNewPart.SelLength = Len(cboNewPart.Text)
End Sub

Private Sub cboNewPart_KeyPress(KeyAscii As Integer)
    cboKeyPress cboNewPart, KeyAscii
End Sub
    
Private Sub CheckActivityCode()
'This subroutine is comparing the Activity Code entered on the screen to the one found in the data
'base.  Right now, a record has been found in the mnb_temp_model_line_part_Master that the user has
'selected to add in as New.  If the Activity Code entered is different from the one already in the
'data base, unpredictable results happen.  Master and detail records get disconnected from each
'other in the data base.

' Set error handling
    On Error GoTo PROC_ERR
        
            
    With mrsTempModelPartsMaster
    'Activity code in data base is different than what is selected on screen
        If strActivityCode <> !temp_part_activity_code Then
            'Change the Screen Activity to what is in data base
            If !temp_part_activity_code = "A" Then
                Call optAdd_Click
                mblnAddTempRec = False
            End If
            
            If !temp_part_activity_code = "I" Then
                Call optInactive_Click
            End If
            
            If !temp_part_activity_code = "R" Then
                Call optReplacePart_Click
            End If
            
        End If
        
    End With
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "CheckActivityCode", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub FindDescription()
    'Use this routine to find and display the Part Description once the
    'parts have been selected for either of the combo's.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsPartDesc As ADODB.Recordset
    Set rsPartDesc = New ADODB.Recordset
    

    ' Set values of fields
    With rsPartDesc
'        'tells the recordset where to get its data from
'        'i.e. gcondatabase="Provider=SQLOLEDB.1;PASSWORD=090400;USERID=SCHIMC
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        ' Change the literal below to the name of your view
                .Source = "select part_id, part_description From v_prod_part " & _
            "where part_id = '" & strPartId & "'"
        .LockType = adLockReadOnly
        .Open

        ' Check for errors returned from the recordset
        If gconDatabase.Errors.Count > 0 Then
            Err.Raise gconDatabase.Errors(0).NativeError, _
                "RetrievepartData", _
                gconDatabase.Errors(0).Description
        End If
    
        If .RecordCount > 0 Then
            strPartDescription = Mid(!part_description, 1, 25)
        Else
            strPartDescription = "New part-No Description"
        End If
            
        .Close
    End With
    
    Set rsPartDesc = Nothing
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "FindDescription", _
        Err.Number, Err.Description)
    
End Sub
Sub FindStockingLocationDesc()
'Use this routine to find and display the Stocking Location in the dropdown once it has
    'been found in the data base record.
    
Dim intLoopIndex As Integer

    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim blnFoundDesc As Boolean
    
    blnFoundDesc = False
    
    For intLoopIndex = 0 To UBound(marrstrStockingLocationId)
        If marrstrStockingLocationId(intLoopIndex) = strStockingLocationID Then
            blnFoundDesc = True
            strStockingLocationDesc = marrstrStockingLocationDesc(intLoopIndex)
            intLoopIndex = UBound(marrstrStockingLocationId)
        End If
    Next intLoopIndex
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "FindStockingLocationDesc", _
        Err.Number, Err.Description)
End Sub


Private Sub cboNewPart_Validate(Cancel As Boolean)
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' If there is no data in the field, display a message,
    ' set cancel to true and exit.
    If Len(Trim(cboNewPart.Text)) = 0 Then
        Cancel = True
        MsgBox "Part Is Required!", _
            vbExclamation + vbOKOnly, _
            "Part Validation"
        GoTo PROC_EXIT
     End If
            
     'Find the part in the list
     cboFindFirst cboNewPart
     
    ' If the part id was not found in the list,
    ' display a message, set cancel to true and
    ' exit.
        If cboNewPart.ListIndex = -1 Then
            MsgBox "Part " & cboNewPart.Text & " is not valid"
            Cancel = True
            GoTo PROC_EXIT
        End If
            
    'Find the description and put to the right of the dropdown
    strPartId = cboNewPart.Text
    Call FindDescription
    lblNewPartDesc.Caption = Left(strPartDescription, 25)
    
    Cancel = False
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboNewPart_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
    
End Sub

Private Sub cboReplacePart_Change()
    cboFindFirst cboReplacePart
End Sub

Private Sub cboReplacePart_KeyPress(KeyAscii As Integer)
    cboKeyPress cboNewPart, KeyAscii
End Sub

Private Sub cboReplacePart_Validate(Cancel As Boolean)
     ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' If there is no data in the field, display a message,
    ' set cancel to true and exit.
    If Len(Trim(cboReplacePart.Text)) = 0 Then
        Cancel = True
        MsgBox "Part to Replace with Is Required!", _
            vbExclamation + vbOKOnly, _
            "Replace Part Validation"
        GoTo PROC_EXIT
    End If
        
    ' If the replace part id has changed, look up the new
    ' part id in the listbox.
    If Not mrsTempModelPartsMaster.BOF Then
        If mrsTempModelPartsMaster!part_id_replaces <> cboReplacePart.Text Then
            mblnRecChanged = True
            cboFindFirst cboReplacePart
            
        ' If the type id was not found in the list,
        ' display a message, set cancel to true and
        ' exit.
            If cboReplacePart.ListIndex = -1 Then
                MsgBox "Line " & cboReplacePart.Text & " is not valid"
                Cancel = True
                GoTo PROC_EXIT
            End If
        End If
    End If
    
    'Find the description and put to the right of the dropdown
    strPartId = cboReplacePart.Text
    Call FindDescription
    lblReplacePartDesc.Caption = Left(strPartDescription, 30)
    
    Cancel = False
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboReplacePart_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cboStockingLocation_Change()
    cboFindFirst cboStockingLocation
End Sub

Private Sub cboStockingLocation_Click()
'Find the current combo record and get the Stocking location id.
' Set error handling
    On Error GoTo PROC_ERR

    'Do not process this event when getting a new record or changing to a differen
    '   record.  Only when the dropdown has been selected.
    If mblnRetrieveModels Then
        Exit Sub
    End If
    
' If the Stocking Location has changed, find the Stocking Location ID from the array.
    If Trim(strStockingLocationDesc) <> Trim(cboStockingLocation.Text) Then
        strStockingLocationID = marrstrStockingLocationId(cboStockingLocation.ListIndex)
        mblnLineChanged = True
    End If
    
    If strLineId = "" Then
        strLineId = Mid(cboStockingLocation.Text, 1, 2)
    End If
    
    If strLineId <> Mid(cboStockingLocation.Text, 1, 2) Then
        mblnLineChanged = True
    End If
    
    strLineId = Mid(cboStockingLocation.Text, 1, 2)
    
    'If the Stocking Location dropdown has changed, then need to refresh the models because the line may
    '  have changed.  However, because of how this screen processes, the cmdRefreshModels_Click would
    '  be called within this process.  The mblnRetrieveModels flag will stop it when that is the circumstance.
    If Not mblnRetrieveModels Then
        If mblnLineChanged Then
            cmdRefreshModels_Click
        End If
    End If
    
    mblnLineChanged = False

PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboStockingLocation_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cboStockingLocation_GotFocus()
    cboStockingLocation.SelStart = 0
    cboStockingLocation.SelLength = Len(cboStockingLocation.Text)
End Sub

Private Sub cboStockingLocation_KeyPress(KeyAscii As Integer)
    cboKeyPress cboStockingLocation, KeyAscii
End Sub

Private Sub cmdAddNew_Click()
'Button for New Data base record has been clicked.  Clear screen, and refresh dropdowns.
  ' Set up error handling
    On Error GoTo PROC_ERR
  
  'Clear screen components to prepare for new entry.
  blnEmptyTempFile = True
  mblnRetrieveModels = True
  Call ClearTempPartScreen
  
    If mblnRecChanged Then
        Call SaveChanges
    End If
  
  optAdd.Enabled = True
  optAdd.Value = False
  optInactive.Enabled = True
  optInactive.Value = False
  optReplacePart.Enabled = True
  optReplacePart.Value = False
  
  cboNewPart.Enabled = False
  cboReplacePart.Enabled = False
  txtPartQuantity.Enabled = False
  txtStepNumber.Enabled = False
  mskECNNumber.Enabled = False
  cboStockingLocation.Enabled = False
  cmdRefreshModels.Enabled = False
  cmdAdd.Enabled = False
  cmdClearAll.Enabled = False
  cmdRemove.Enabled = False
  
  cmdFirst.Enabled = False
  cmdLast.Enabled = False
  cmdNext.Enabled = False
  cmdPrevious.Enabled = False
  cmdList.Enabled = False
  
  strLineId = ""
  strStockingLocationID = ""
  
  mblnAddTempRec = True
  mblnRetrieveModels = False
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cmdAddNew_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdDelete_Click()
    Call mnuFileDelete_Click
End Sub

Private Sub cmdClose_Click()
    ' Purpose:  Close the form by the user's request.
    Call mnuFileClose_Click
End Sub

Private Sub cmdFirst_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewFirst menu item.
    Call mnuViewFirst_Click
End Sub

Private Sub cmdHelp_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the HelpContents menu item.
    Call mnuHelpContents_Click
End Sub

Private Sub cmdLast_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewLast menu item.
    Call mnuViewLast_Click
End Sub

Private Sub cmdList_Click()
     Call mnuViewList_Click
End Sub

Private Sub cmdNext_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewNext menu item.
    Call mnuViewNext_Click
End Sub

Private Sub cmdPrevious_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewPrevious menu item.
    Call mnuViewPrevious_Click
End Sub

Private Sub cmdRefreshModels_Click()
    ' Retrieve model data.  The models showing change based on the radio button
    '    option selected, and the different part selected.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Verify one of the option buttons have been selected
    If optAdd.Value = False And optInactive.Value = False And optReplacePart = False Then
        MsgBox "Select an Option, Add Part, Inactive Part, or Replace Part", vbOKOnly, _
        "Select Options..."
        Exit Sub
    End If

    'Verify a Part is selected from the combo list.  Necessary before the Model
    '    Query can run.
    If Len(cboNewPart.Text) > 0 Then
        cboFindFirst cboNewPart
    End If
       
    If cboNewPart.Visible Then
        If cboNewPart.ListIndex = -1 Then
            MsgBox "Please select a Part ", vbOKOnly, "Select a Part..."
            Exit Sub
        End If
'    End If
        
    'Verify a Replacement part selected from combo list if the Replace Part option
    '   was selected.
    If optReplacePart.Value = True Then
        If Len(cboReplacePart.Text) > 0 Then
                cboFindFirst cboReplacePart
        End If
        If cboReplacePart.ListIndex = -1 Then
            MsgBox "Please select a Part to Replace ", vbOKOnly, "Select a Replace Part..."
            Exit Sub
        End If
    End If
    
    End If
      
       
    'Running the ValidEntries subroutine only if a New Record has been added to guarantee
    'All the fields are entered.
    If Screen.ActiveForm Is Me Then
        If mblnAddTempRec Then
            Call ValidEntries(blnCancel)
            If blnCancel Then
                GoTo PROC_EXIT
            End If
        End If
    End If
    
    If Screen.ActiveForm Is Me And optReplacePart.Value = True Then
        If cboReplacePart.ListIndex = -1 Then
            MsgBox "Please select a Replace Part ", vbOKOnly, "Select a Replace with Part..."
        Exit Sub
        End If
    End If
    
    'Check to see that the New Part added is different from the Replacement Part
    If Screen.ActiveForm Is Me And optReplacePart.Value = True Then
        If cboReplacePart.Text = cboNewPart.Text Then
            MsgBox "The Part Selections are the same -- " & vbCrLf & _
                "Please change one of them.", vbOKOnly, "Select a Replace with Part..."
        Exit Sub
        End If
    End If
    
 ' Clear the available Models.
    lstAvailable.Clear
    lstSelected.Clear

 ' Get the models based on the parts requested.
    Call RetrieveModelData

PROC_EXIT:
    blnCancel = True
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdRefreshModels_Click", Err.Number, _
        Err.Description)
    Unload Me
End Sub

Private Sub cmdSave_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the FileSave menu item.
    
    Call mnuFileSave_Click
End Sub

Private Sub cmdAdd_Click()
    
    ' Purpose:  Add the selected record in the available column to the selected list and to the
    '           recordset.  Add records to mnb_model_part, mnb_model_part_stocking_location,
    '           and mnb_temp_model_line_part_master.
       
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsModifyModelPart As New ADODB.Recordset
    Dim rsModifyModelPartStockingLocation As New ADODB.Recordset
    
    Dim rsGetHighestOriginalSequence As ADODB.Recordset
    Set rsGetHighestOriginalSequence = New ADODB.Recordset
    
    ' If no model was selected send an error and exit
    ' the sub.
    If lstAvailable.ListIndex = -1 Then
        MsgBox "Please select an entry from the left."
        GoTo PROC_EXIT
    End If
  
    If cboStockingLocation.ListIndex = -1 Then
            cboStockingLocation.SetFocus
            MsgBox "Please select a Stocking Location."
            GoTo PROC_EXIT
    End If
    
    ' Set the part code
    strModelNumber = Mid(lstAvailable.Text, 1, 15)
    strLineId = Right(lstAvailable.Text, 2)
    
        With rsGetHighestOriginalSequence
                       Set .ActiveConnection = gconDatabase
                      .CursorLocation = adUseClient
                      .CursorType = adOpenKeyset
                      .LockType = adLockReadOnly
                      .Source = "select model_number,line_Id,max(original_sequence_number) as high_original_sequence , " & _
                         "max(part_sequence_number) As high_part_sequence_number From v_mnb_model_part " & _
                         "where model_number = '" & strModelNumber & "' and line_id = '" & strLineId & _
                         "' group by model_number, line_id"
                    .Open
        End With
    
    

 '***  Activity Code = A or R:  Add the part to the mnb_model_part and mnb_model_part_stocking_location
        If Not optInactive.Value = True Then
        'Determining sequence numbers by finding the last ones used for that model/line.
            
            Dim rsPartSequence As ADODB.Recordset
            Set rsPartSequence = New ADODB.Recordset
                                
            If optAdd.Value = True Then
                With rsPartSequence
                    Set .ActiveConnection = gconDatabase
                      .CursorLocation = adUseClient
                      .CursorType = adOpenKeyset
                      .LockType = adLockReadOnly
                      .Source = "select model_number,line_Id,max(original_sequence_number) as high_original_sequence , " & _
                         "max(part_sequence_number) As high_part_sequence_number From v_mnb_model_part " & _
                         "where model_number = '" & strModelNumber & "' and line_id = '" & strLineId & _
                         "' group by model_number, line_id"
                    .Open
                End With
                    
                    '  Check to see if there are mnb_model_part records before adding a new one
                Set rsModifyModelPart = New ADODB.Recordset
                With rsModifyModelPart                                                             '3
                    Set .ActiveConnection = gconDatabase
                        .CursorLocation = adUseClient
                        .CursorType = adOpenKeyset
                        .LockType = adLockBatchOptimistic
                        .Source = "SELECT * From V_MNB_Model_part where model_number = '" & _
                               strModelNumber & _
                               "' and line_id = '" & strLineId & _
                                "' and part_id = '" & cboNewPart.Text & "'"
                        .Open
                      
                  'There is no mnb_model_part
                    If .RecordCount = 0 Then
                        .AddNew
                        !model_number = Trim(strModelNumber)
                        !line_id = strLineId
                        !part_id = Trim(cboNewPart.Text)
                        !original_sequence_number = rsPartSequence!high_original_sequence + 1
                        !quantity = txtPartQuantity.Text
                        !start_ecn_number = ""
                        !start_ecn_date = ""
                        !start_ecn_flag = "I"
                        !quit_ecn_number = mskECNNumber.ClipText
                        !quit_ecn_date = Mid(mskTempEndDate.ClipText, 5, 4) & _
                                        Mid(mskTempEndDate.ClipText, 1, 2) & _
                                        Mid(mskTempEndDate.ClipText, 3, 2)
                        !quit_ecn_flag = "Q"
                        !level_number = "02"
                        !Comments = "Minibill App New"
                        !part_create_date = Now
                        !part_reviewed_flag = True
                        !parent_part_number = strModelNumber
                        !inactive_Part_flag = False
                        !part_sequence_number = rsPartSequence!high_part_sequence_number + 1
                        .UpdateBatch
                    End If
                End With
   
 '               Create the mnb_Model_part_stocking_location record if new part to model.
                 Set rsModifyModelPartStockingLocation = New ADODB.Recordset
                 With rsModifyModelPartStockingLocation
                    Set .ActiveConnection = gconDatabase
                        .CursorLocation = adUseClient
                        .CursorType = adOpenKeyset
                        .LockType = adLockBatchOptimistic
                        .Source = "SELECT * From V_MNB_Model_Part_stocking_location_all where  " & _
                            "model_number = '" & strModelNumber & _
                            "' and line_id = '" & strLineId & _
                            "' and part_id = '" & cboNewPart.Text & _
                            "' and stocking_location_id = '" & _
                            marrstrStockingLocationId(cboStockingLocation.ListIndex) & "' and part_sequence_number = '" & _
                            rsPartSequence!high_part_sequence_number + 1 & "'"
                     
                        .Open
                        .AddNew
                        !model_number = strModelNumber
                        !line_id = strLineId
                        !part_id = cboNewPart.Text
                        !part_sequence_number = rsPartSequence!high_part_sequence_number + 1
                        !stocking_location_id = marrstrStockingLocationId(cboStockingLocation.ListIndex)
 '                         !step_number = txtStepNumber.Text
                        !model_part_stocking_location_obsolete_flag = False
                        !model_part_stocking_location_obsolete_date = Null
                        .UpdateBatch
                            
                               ' Check for errors
                        If gconDatabase.Errors.Count > 0 Then
                            Err.Raise gconDatabase.Errors(0).NativeError, _
                            "cmdAdd_Click", gconDatabase.Errors(0).Description
                        End If
            
                 End With
            End If
            'End logic for Adding parts to models.
                       
            Set rsModifyModelPartStockingLocation = Nothing
            Set rsModifyModelPart = Nothing
'            Set rsPartSequence = Nothing

              'Looking for all parts to be replaced in the stocking location requested from the entry page. Find this first
              ' to get sequence number to get correct mnb_model_part record.
            If optReplacePart.Value = True Then
                With rsPartSequence
                    Set .ActiveConnection = gconDatabase
                        .CursorLocation = adUseClient
                        .CursorType = adOpenKeyset
                        .LockType = adLockBatchOptimistic
                        .Source = "select * From v_mnb_model_part_stocking_location " & _
                            "where model_number = '" & strModelNumber & "' and line_id = '" & strLineId & _
                            "' and part_id = '" & cboReplacePart.Text & "' and stocking_location_id = '" & _
                            marrstrStockingLocationId(cboStockingLocation.ListIndex) & "' order by part_sequence_number"
                         .Open
                    .MoveFirst
                    Do While Not .EOF
 
     'Need to see if the part exists in the mnb_model_part_stocking_location.  If it does,
     'then do not add again.  Prevents part from being duplicated when there are updates
     'coming from the text file.  Only add if it does not exist.
                
     '  Check to see if there are mnb_model_part records before adding a new one
                        Set rsModifyModelPart = New ADODB.Recordset
                        With rsModifyModelPart
                        Set .ActiveConnection = gconDatabase
                            .CursorLocation = adUseClient
                            .CursorType = adOpenKeyset
                            .LockType = adLockBatchOptimistic
                            .Source = "SELECT * From V_MNB_Model_part where model_number = '" & _
                               strModelNumber & _
                               "' and line_id = '" & strLineId & _
                                "' and part_id = '" & cboNewPart.Text & "' and original_sequence_number = '" & _
                                rsPartSequence!part_sequence_number & "'"
                            .Open
                                Do While Not .EOF
               
                                    Set rsModifyModelPartStockingLocation = New ADODB.Recordset
                
                                    With rsModifyModelPartStockingLocation
                                        Set .ActiveConnection = gconDatabase
                                            .CursorLocation = adUseClient
                                            .CursorType = adOpenKeyset
                                            .LockType = adLockBatchOptimistic
                                            .Source = "SELECT * From V_MNB_Model_Part_stocking_location_all where  " & _
                                                "model_number = '" & strModelNumber & _
                                                "' and line_id = '" & strLineId & _
                                                "' and part_id = '" & cboNewPart.Text & _
                                                "' and stocking_location_id = '" & _
                                                marrstrStockingLocationId(cboStockingLocation.ListIndex) & "' and part_sequence_number = '" & _
                                                rsModifyModelPart!part_sequence_number & "'"
                     
                                            .Open

                                            If .RecordCount > 0 Then
                                                !model_part_stocking_location_obsolete_flag = "0"
                                                !model_part_stocking_location_obsolete_date = Null
                                                .Update
                                            End If
                                    End With
                                    .MoveNext
                                Loop
  
                  'There is no mnb_model_part
                                If .RecordCount = 0 Then
                                    .AddNew
                                    !model_number = strModelNumber
                                    !line_id = strLineId
                                    !part_id = cboNewPart.Text
                                    !original_sequence_number = rsGetHighestOriginalSequence!high_original_sequence + 1
                                    !quantity = txtPartQuantity.Text
                                    !start_ecn_number = ""
                                    !start_ecn_date = ""
                                    !start_ecn_flag = "I"
                                    !quit_ecn_number = mskECNNumber.ClipText
                                    !quit_ecn_date = Mid(mskTempEndDate.ClipText, 5, 4) & _
                                            Mid(mskTempEndDate.ClipText, 1, 2) & _
                                            Mid(mskTempEndDate.ClipText, 3, 2)
                                    !quit_ecn_flag = "Q"
                                    !level_number = "02"
                                    !Comments = "Minibill App New"
                                    !part_create_date = Now
                                    !part_reviewed_flag = True
                                    !parent_part_number = strModelNumber
                                    !inactive_Part_flag = False
                                    !part_sequence_number = rsPartSequence!part_sequence_number
                                    .UpdateBatch
 
 '               Create the mnb_Model_part_stocking_location record if new part to model.
                                    Set rsModifyModelPartStockingLocation = New ADODB.Recordset
                                    With rsModifyModelPartStockingLocation
                                        Set .ActiveConnection = gconDatabase
                                            .CursorLocation = adUseClient
                                            .CursorType = adOpenKeyset
                                            .LockType = adLockBatchOptimistic
                                            .Source = "SELECT * From V_MNB_Model_Part_stocking_location_all where  " & _
                                                "model_number = '" & strModelNumber & _
                                                "' and line_id = '" & strLineId & _
                                                "' and part_id = '" & cboNewPart.Text & _
                                                "' and stocking_location_id = '" & _
                                                marrstrStockingLocationId(cboStockingLocation.ListIndex) & "' and part_sequence_number = '" & _
                                                rsPartSequence!part_sequence_number & "'"
                     
                                        .Open
                                        .AddNew
                                        !model_number = strModelNumber
                                        !line_id = strLineId
                                        !part_id = cboNewPart.Text
                                        !part_sequence_number = rsPartSequence!part_sequence_number
                                        !stocking_location_id = marrstrStockingLocationId(cboStockingLocation.ListIndex)
 '                                       !step_number = txtStepNumber.Text
                                        !model_part_stocking_location_obsolete_flag = False
                                        !model_part_stocking_location_obsolete_date = Null
                                        .UpdateBatch
                            
                               ' Check for errors
                                        If gconDatabase.Errors.Count > 0 Then
                                            Err.Raise gconDatabase.Errors(0).NativeError, _
                                                "cmdAdd_Click", gconDatabase.Errors(0).Description
                                        End If
            
                                    End With
                                End If
                        End With
      
      '  This logic makes the parts inactive that are being replaced by the new part from the entry screen.
      '  Find the configured stocking location record first to get the sequence code for the records to be made inactive
      '   in the mnb_model_part record.  Reads all matching parts in that stocking location.
        Set rsModifyModelPartStockingLocation = New ADODB.Recordset
            With rsModifyModelPartStockingLocation
            Set .ActiveConnection = gconDatabase
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
                .LockType = adLockBatchOptimistic

                .Source = "SELECT * From V_MNB_Model_Part_Stocking_Location_all where  " & _
                    "model_number = '" & strModelNumber & "' and line_id = '" & _
                    strLineId & "' and part_id = '" & _
                    cboReplacePart.Text & "' and stocking_location_id = '" & _
                    marrstrStockingLocationId(cboStockingLocation.ListIndex) & _
                    "' order by part_sequence_number"
                    
                .Open
                    'Find the matching model_part from the model_part_stocking_location for the replaced part in model.

                 Do While Not .EOF
                    rsModifyModelPartStockingLocation!model_part_stocking_location_obsolete_flag = True
                    rsModifyModelPartStockingLocation!model_part_stocking_location_obsolete_date = Now
                    .UpdateBatch
                           
                    Set rsModifyModelPart = New ADODB.Recordset
                    With rsModifyModelPart
                        Set .ActiveConnection = gconDatabase
                            .CursorLocation = adUseClient
                            .CursorType = adOpenKeyset
                            .LockType = adLockBatchOptimistic
                            .Source = "SELECT * From V_MNB_Model_Part where  " & _
                                "model_number = '" & strModelNumber & _
                                "' and line_id = '" & strLineId & _
                                "' and part_id = '" & cboReplacePart.Text & _
                                "' and part_sequence_number = '" & rsModifyModelPartStockingLocation!part_sequence_number & "'"
                            .Open
                        
                            If .RecordCount > 0 Then
                                Do While Not .EOF
                                    !quit_ecn_number = mskECNNumber.ClipText
                                    !quit_ecn_date = Mid(mskTempEndDate.ClipText, 5, 4) & _
                                                Mid(mskTempEndDate.ClipText, 1, 2) & _
                                                Mid(mskTempEndDate.ClipText, 3, 2)
                                    !quit_ecn_flag = "Q"
                                    !Comments = "Minibill App-Replace"
                                    !inactive_Part_flag = "1"
                                    .UpdateBatch
                                    .MoveNext
                                Loop
                            End If
                    End With
                    .MoveNext
                Loop
            End With
                    
                        Set rsModifyModelPartStockingLocation = Nothing
                        Set rsModifyModelPart = Nothing
                        
'                    Else
'                        MsgBox "Cannot use this screen to configure." & vbCrLf & "Part already exists in Model"
'                        GoTo PROC_EXIT
                    .MoveNext
                    Loop
                   
                   End With
                   Set rsGetHighestOriginalSequence = Nothing
               End If
            End If

        Set rsPartSequence = Nothing

    ' Check for errors
            If gconDatabase.Errors.Count > 0 Then
                Err.Raise gconDatabase.Errors(0).NativeError, _
                "cmdAdd_Click", gconDatabase.Errors(0).Description
            End If
            
        
            '****************If  Activity code is I, make the part inactive
            If optInactive.Value = True Then
              
             'Read through all model_part to find all model_part_stocking_location for a part in model.
            
                Set rsModifyModelPartStockingLocation = New ADODB.Recordset
                With rsModifyModelPartStockingLocation
                    Set .ActiveConnection = gconDatabase
                        .CursorLocation = adUseClient
                        .CursorType = adOpenKeyset
                        .LockType = adLockBatchOptimistic

                        .Source = "SELECT * From V_MNB_Model_Part_Stocking_Location_all where  " & _
                           "model_number = '" & strModelNumber & "' and line_id = '" & _
                            strLineId & "' and part_id = '" & _
                            cboNewPart.Text & "' and stocking_location_id = '" & _
                            marrstrStockingLocationId(cboStockingLocation.ListIndex) & _
                            "' order by part_sequence_number"
                    
                    .Open
                    Do While Not .EOF
                        
                        If .RecordCount > 0 Then
                            Do While Not .EOF
                                !model_part_stocking_location_obsolete_flag = True
                                !model_part_stocking_location_obsolete_date = Now
                                .UpdateBatch
                                                
                            Set rsModifyModelPart = New ADODB.Recordset
                            With rsModifyModelPart
                                Set .ActiveConnection = gconDatabase
                                    .CursorLocation = adUseClient
                                    .CursorType = adOpenKeyset
                                    .LockType = adLockBatchOptimistic
                                    .Source = "SELECT * From V_MNB_Model_Part where  " & _
                                        "model_number = '" & strModelNumber & _
                                        "' and line_id = '" & strLineId & _
                                        "' and part_id = '" & cboNewPart.Text & _
                                        "' and part_sequence_number = '" & rsModifyModelPartStockingLocation!part_sequence_number & "'"

                                    .Open
                                If .RecordCount > 0 Then
                                    !quit_ecn_number = mskECNNumber.ClipText
                                    !quit_ecn_date = Mid(mskTempEndDate.ClipText, 5, 4) & _
                                                    Mid(mskTempEndDate.ClipText, 1, 2) & _
                                                    Mid(mskTempEndDate.ClipText, 3, 2)
                                    !quit_ecn_flag = "Q"
                                    !Comments = "APP - Inactive Temp. Parts"
                                    !inactive_Part_flag = "1"
                                    .UpdateBatch
                                End If
                            End With
                            .MoveNext
                            Loop
                        End If
                    Loop
                End With
        End If

    
'   Create the Temp Part Record
    Call CreateNewTempPartRecord
    
    ' Add the item to the selected listbox
    lstSelected.AddItem lstAvailable.Text & "  " & Mid(cboStockingLocation.Text, 3, 25)
    
    ' Remove from the available listbox
    lstAvailable.RemoveItem _
    lstAvailable.ListIndex
    
    ' Set the listindex of the selected listbox to the
    ' index just added.
    lstSelected.ListIndex = lstSelected.NewIndex
    
    Set mrsModelPartStockLoc = Nothing
  '  Set rsGetHighestOriginalSequence = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdAdd_Click", Err.Number, _
        Err.Description)
    Unload Me

End Sub

Private Sub cmdRemove_Click()
    'Purpose:  To delete the changes made to a model for a particular part.
          'Change mnb_model_part table for the following:
            '1. Make the part inactive for that model if it was added.
            '2. Make the part active if it was made inactive.
            '3. Make the new part inactive and the old part active if the part was being replaced.
            '4. Delete the mnb_temp_model_line_part_detail for a model/line
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim intLoopIndex As Integer
    
    ' If no item was selected in the Selected listbox,
    ' display an error and leave the sub.
    If lstSelected.ListIndex = -1 Then
        MsgBox "There is no entry selected"
        GoTo PROC_EXIT
    End If
    
    If cboStockingLocation.ListIndex = -1 Then
        MsgBox "Stocking Location not Selected"
        GoTo PROC_EXIT
    End If
    
 'Find the stocking location id from the stocking location description of the record to be removed.
    For intLoopIndex = 0 To UBound(marrstrStockingLocationDesc)
        If Trim(marrstrStockingLocationDesc(intLoopIndex)) = Trim(Mid(lstSelected.Text, 19, 30)) Then
            strStockingLocationDesc = Trim(marrstrStockingLocationDesc(intLoopIndex))
            strStockingLocationID = marrstrStockingLocationId(intLoopIndex)
            intLoopIndex = UBound(marrstrStockingLocationId)
        End If
    Next intLoopIndex
    
    ' Set the model and line code
    strModelNumber = Mid(lstSelected.Text, 1, 15)
    strLineId = Mid(lstSelected.Text, 19, 2)
    
    'Find the Part records in the selected models.  Change the inactive_part_flag = 1 to make
    '   part inactive for Add and Replace Options.
    
    If Not optInactive Then                  'Reference Not optInactive*****
        Dim rsList As ADODB.Recordset
        Set rsList = New ADODB.Recordset

        With rsList
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Source = "select * from v_mnb_model_part " & _
                "where model_number = '" & strModelNumber & "' and line_id = '" & strLineId & _
                "' and part_id = '" & cboNewPart.Text & "'"
            .Open
        
            If .RecordCount > 0 Then
                Do While Not .EOF
                
                Dim rsInactivateModelPartStockLoc As ADODB.Recordset
                Set rsInactivateModelPartStockLoc = New ADODB.Recordset

                With rsInactivateModelPartStockLoc
                    Set .ActiveConnection = gconDatabase
                    .CursorLocation = adUseClient
                    .CursorType = adOpenKeyset
                    .LockType = adLockOptimistic
                    .Source = "select * from v_mnb_model_part_stocking_location " & _
                        "where model_number = '" & strModelNumber & "' and line_id = '" & strLineId & _
                        "' and part_id = '" & cboNewPart.Text & "' and part_sequence_number = '" & _
                        rsList!part_sequence_number & "'"
                    .Open
        
                    If .RecordCount > 0 Then
                        Do While Not .EOF
                            .Delete
                            .MoveNext
                        Loop
                    End If

                End With
                    .Delete
                    .MoveNext
                Loop
            End If

        End With
    
    
    End If                                    'Reference Not optInactive*****
    
    ' Delete the model/line record selected in the list box from mnb_temp_model_line_part_detail
    'Create empty recordset for Temporary part changes
        Dim mrsTempModelParts As ADODB.Recordset
        Set mrsTempModelParts = New ADODB.Recordset
    
        With mrsTempModelParts
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Source = "select * From v_mnb_temp_model_line_part_detail where model_number = '" & _
            strModelNumber & " ' and Line_id = '" & strLineId & "' and part_id = '" & _
            cboNewPart.Text & "' and stocking_location_id = '" & strStockingLocationID & "'"
            .Open
            .Delete
        End With
  
 'Setting a value for strPartId.  Changes based on the Option selected.
 If optInactive Then
    strPartId = cboNewPart.Text
 End If
 
 If optReplacePart Then
    strPartId = cboReplacePart.Text
 End If
  
 'If Replace Part option is used, then reactivate the part that was replaced.
    If Not optAdd.Value = True Then
        
        Dim mrsActivateReplacePartStockingLoc As ADODB.Recordset
        Set mrsActivateReplacePartStockingLoc = New ADODB.Recordset
 
        With mrsActivateReplacePartStockingLoc
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Source = "select * From v_mnb_model_part_stocking_location_all " & _
                "where model_number = '" & strModelNumber & "' and line_id = '" & strLineId & _
                "' and part_id = '" & strPartId & "' and stocking_location_id = '" & strStockingLocationID & "'"
            .Open
    
            If .RecordCount > 0 Then
                Do While Not .EOF
                    !model_part_stocking_location_obsolete_flag = "0"
                    !model_part_stocking_location_obsolete_date = Null
                    .Update

                    Dim mrsActivateReplacePart As ADODB.Recordset
                    Set mrsActivateReplacePart = New ADODB.Recordset
    
                    With mrsActivateReplacePart
                        Set .ActiveConnection = gconDatabase
                        .CursorLocation = adUseClient
                        .CursorType = adOpenKeyset
                        .LockType = adLockOptimistic
                        .Source = "select * From v_mnb_model_part where model_number = '" & _
                            strModelNumber & " ' and Line_id = '" & strLineId & "' and part_id = '" & _
                            strPartId & "' and part_sequence_number = '" & mrsActivateReplacePartStockingLoc!part_sequence_number & "'"
                        .Open

                        If .RecordCount > 0 Then
                            Do While Not .EOF
                                !inactive_Part_flag = "0"
                                !quit_ecn_number = ""
                                !quit_ecn_date = ""
                                !quit_ecn_flag = ""
                                .Update
                                .MoveNext
                            Loop
                        End If
                    End With
        
                    .MoveNext
                Loop
            End If
        
        
        End With
 
        
        Set mrsActivateReplacePart = Nothing
        Set mrsActivateReplacePartStockingLoc = Nothing
        Set rsInactivateModelPartStockLoc = Nothing
    
    End If
    
    ' Add the item to the available listbox only if the lines match.
    If strLineId = Mid(cboStockingLocation.Text, 1, 2) Then
        lstAvailable.AddItem Mid(lstSelected.Text, 1, 20)
    End If
    
    ' Remove the item from the selected listbox
    lstSelected.RemoveItem lstSelected.ListIndex
    
    ' Set the listindex of the Available listbox to the
    ' index just added if the lines match.
    If strLineId = Mid(cboStockingLocation.Text, 1, 2) Then
        lstAvailable.ListIndex = lstAvailable.NewIndex
    End If
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdRemove_Click", Err.Number, _
        Err.Description)
    Unload Me

End Sub

Private Sub cmdClearAll_Click()
    ' Purpose:  Remove all models from the selected
    '           list and from the recordset.  Make parts inactive in the model.
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ' Set the variable for the looping index
    Dim intLoopIndex As Integer
    Dim strQueryPart As String
        
    Dim rsActivateModelPart As ADODB.Recordset
    Set rsActivateModelPart = New ADODB.Recordset
    
    Dim rsInactivateModelPartStockLoc As ADODB.Recordset
    Set rsInactivateModelPartStockLoc = New ADODB.Recordset
    
    ' Loop through the selected listbox to find the mnb_model_part to make record inactive.
    For intLoopIndex = 0 To lstSelected.ListCount - 1
        
        strModelNumber = Mid(lstSelected.List(intLoopIndex), 1, 15)
        strLineId = Mid(lstSelected.List(intLoopIndex), 19, 2)
        strQueryPart = cboNewPart.Text
        
        'Find the mnb_model_Part_stocking_location records in the selected models.  Change the flag inactive for Add and Replace Options.
        With rsInactivateModelPartStockLoc
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Source = "select * from v_mnb_model_part_stocking_location_all " & _
                "where model_number = '" & strModelNumber & "' and line_id = '" & strLineId & _
                "' and part_id = '" & strQueryPart & "'"
            .Open
'Delete any Adds/Replace, and add any Inactives back in for all occurrences of the part.
            
            If .RecordCount > 0 Then
                Do While Not .EOF
        
        'Find the mnb_model_Part_stocking_location records in the selected models.  Change the
        '  flag inactive for Add and Replace Options.

        Dim rsModifyModelPart As ADODB.Recordset
        Set rsModifyModelPart = New ADODB.Recordset

        With rsModifyModelPart
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
'            .LockType = adLockReadOnly
            .LockType = adLockOptimistic
            .Source = "select * from v_mnb_model_part " & _
                "where model_number = '" & strModelNumber & "' and line_id = '" & strLineId & _
                "' and part_id = '" & strQueryPart & "' and part_sequence_number = '" & rsInactivateModelPartStockLoc!part_sequence_number & "'"
            .Open
        
            If Not optInactive = True Then
                If .RecordCount > 0 Then
                    Do While Not .EOF
                        .Delete
                       .MoveNext
                    Loop
                End If
            End If
         End With
         
         'Find the mnb_model_Part records in the selected models.  Change the flag inactive for Replace Option.
         If optInactive = True Then
            With rsModifyModelPart
                If .RecordCount > 0 Then
                    Do While Not .EOF
                        !inactive_Part_flag = "0"
                        !quit_ecn_number = ""
                        !quit_ecn_date = ""
                        !quit_ecn_flag = ""
                        .Update
                        .MoveNext
                    Loop
                End If

            End With
        End If
            
         
        'If the Replace Option is used, then activate the inactive part.
        If optReplacePart.Value = True Then
        
            strQueryPart = cboReplacePart.Text
            Dim rsModifyModelPartForReplace As ADODB.Recordset
            Set rsModifyModelPartForReplace = New ADODB.Recordset
            
            Dim rsModelPartStockLocForReplace As ADODB.Recordset
            Set rsModelPartStockLocForReplace = New ADODB.Recordset
            
            With rsModelPartStockLocForReplace
                Set .ActiveConnection = gconDatabase
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
'                .LockType = adLockReadOnly
                .LockType = adLockOptimistic
                .Source = "select * from v_mnb_model_part_stocking_location_all " & _
                    "where model_number = '" & strModelNumber & "' and line_id = '" & strLineId & _
                    "' and part_id = '" & strQueryPart & "'"
                .Open
                
                If .RecordCount > 0 Then
                 
        'Make all occurrences of the part in the model active
            With rsModifyModelPartForReplace
                Set .ActiveConnection = gconDatabase
                .CursorLocation = adUseClient
                .CursorType = adOpenKeyset
'                .LockType = adLockReadOnly
                .LockType = adLockOptimistic
                .Source = "select * from v_mnb_model_part " & _
                    "where model_number = '" & strModelNumber & "' and line_id = '" & strLineId & _
                    "' and part_id = '" & strQueryPart & "' and part_sequence_number = '" & rsModelPartStockLocForReplace!part_sequence_number & "'"
                .Open
                    If .RecordCount > 0 Then
                        Do While Not .EOF
                            !inactive_Part_flag = "0"
                            !quit_ecn_number = ""
                            !quit_ecn_date = ""
                            !quit_ecn_flag = ""
                            .Update
                            .MoveNext
                        Loop
                    End If
                    .Close
                End With
                End If
                
                !model_part_stocking_location_obsolete_flag = "0"
                !model_part_stocking_location_obsolete_date = Null
                .Update
            End With
         End If
            

'  This section loops through all model_part for a model in case there are multiples of the same part.
                    If Not optInactive Then
                        .Delete
                    Else
                        !model_part_stocking_location_obsolete_flag = "0"
                        !model_part_stocking_location_obsolete_date = Null
                    End If
                    .MoveNext
                Loop
            End If
                
            .Close
        End With
        
        Set rsModifyModelPart = Nothing
'   End section for looping through model_part records.

     Next intLoopIndex
    
    Set rsModifyModelPart = Nothing
    Set rsActivateModelPart = Nothing
    Set rsInactivateModelPartStockLoc = Nothing
    
    'Delete the Temporary Detail Records
    Dim rsdeleteTempModelParts As ADODB.Recordset
    Set rsdeleteTempModelParts = New ADODB.Recordset
    
    Dim intDeleteLoop As Integer
        
    With rsdeleteTempModelParts
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
'         .LockType = adLockBatchOptimistic
        .LockType = adLockOptimistic
        .Source = "Delete from v_mnb_temp_model_line_part_detail where part_id = '" & _
            cboNewPart.Text & "'"
       .Open
    End With

    Set rsdeleteTempModelParts = Nothing
    
    'Do not refresh models if the record is going to be deleted.
    If Not blnDeleteRecord Then
        Call cmdRefreshModels_Click
    End If

PROC_EXIT:
    
    Exit Sub
    
PROC_ERR:
    
    Call ShowError(Me.Name, "cmdClearAll", Err.Number, _
        Err.Description)
    Unload Me
End Sub

Private Sub cmdTempEndDate_Click()
    'Determine the ending effective date
    
     ' Set up error handling
    On Error GoTo PROC_ERR

    'If the date from the screen is blank, fill with future week date.
    If IsNull(mskTempEndDate.Text) Then
        mskTempEndDate.Text = dteWeekFuture
        dlgCalendar.mdteSelectedDate = mskTempEndDate.Text
    Else
        dlgCalendar.mdteSelectedDate = mskTempEndDate.Text
    End If
    
    'Go to the Calendar, and show date from the screen.
    dlgCalendar.Show vbModal
    
    'Validate the date changes from the calendar
    If Format(dteToday, "yyyy/mm/dd") > Format(dlgCalendar.mdteSelectedDate, "yyyy/mm/dd") Then
        MsgBox "Date must be Future", vbOKOnly, "Verify Date"
        Exit Sub
    End If
    
'    If Format(dlgCalendar.mdteSelectedDate, "yyyy/mm/dd") > Format(dteWeekTwoFuture, "yyyy/mm/dd") Then
'        MsgBox "Date can be no more than 30 days" & vbCrLf & "in the future", _
'        vbOKOnly, "Verify Date"
'        Exit Sub
'    End If
   
    'Reformat the date and compare what was in the record
    With mrsTempModelPartsMaster
        If Not .BOF And .EOF Then
            If Not IsNull(dlgCalendar.mdteSelectedDate) Then
                mskTempEndDate.Text = Format( _
                    dlgCalendar.mdteSelectedDate, "mm/dd/yyyy")
                dtePicked = Format(Mid(!quit_ecn_date, 1, 4) & _
                    "/" & Mid(!quit_ecn_date, 5, 2) & _
                    "/" & Mid(!quit_ecn_date, 7, 2), "mm/dd/yyyy")
                    If Format(dtePicked, "yyyy/mm/dd") <> Format(dlgCalendar.mdteSelectedDate, "yyyy/dd/mm") Then
                        mblnRecChanged = True
                        !quit_ecn_date = Format(dlgCalendar.mdteSelectedDate, "yyyy") & _
                            Format(dlgCalendar.mdteSelectedDate, "mm") & Format(dlgCalendar.mdteSelectedDate, "dd")
                    End If
                Else
                    MsgBox "Must enter Temporary End Date ", vbOKOnly, "Date Entry Error..."
                End If
        Else
            mskTempEndDate.Text = Format( _
                    dlgCalendar.mdteSelectedDate, "mm/dd/yyyy")
        End If
        
    End With
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdTempEndDate_Click", Err.Number, _
        Err.Description)
    Unload Me

End Sub

Private Sub Form_Load()
    
    
    ' Purpose:  Show the form and login to the server
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    frmProcessing.Label2 = Time
    frmProcessing.Show
   
    DoEvents
    
    If gconDatabase Is Nothing Then
        Set gconDatabase = gclsSQLServer.Connect( _
            gclsMESApplication.ApplicationRole, _
            gclsMESApplication.ApplicationPassword)
    
        If gconDatabase.State <> adStateOpen Then
            Err.Raise gconDatabase.Errors(0).NativeError, _
                "Form_Load", gconDatabase.Errors(0).Description
        End If
    End If
    
 'Make items on screen unavailable until critical items are filled.
    cboNewPart.Visible = False
    cboReplacePart.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    cmdAdd.Enabled = False
    cmdClearAll.Enabled = False
    cmdRemove.Enabled = False
    cmdTempEndDate.Enabled = False
    mskECNNumber.Enabled = False
    txtPart.Enabled = False
    txtPartQuantity.Enabled = False
    txtStepNumber.Enabled = False
    mskTempEndDate.Enabled = False
'    cmdPrint.Enabled = False
    cmdPrint.Visible = False
    cmdRefreshModels.Enabled = False
    cboStockingLocation.Enabled = False
    
    'Create comparison date of today to verify the Temporary End Date
    '  Check for Week 1 and Week 2.
    
    dlgCalendar.mdteSelectedDate = DateAdd("d", 7, Now)
    mskTempEndDate.Text = Format( _
            dlgCalendar.mdteSelectedDate, "mm/dd/yyyy")
            
    dteWeekFuture = dlgCalendar.mdteSelectedDate
    dteWeekTwoFuture = DateAdd("d", 30, Now)
    dteToday = Now
    
    ' Disable fields if update is not allowed.
    If Not gblnUpdate Then
        cmdAdd.Enabled = False
        cmdRemove.Enabled = False
        cmdClearAll.Enabled = False
    End If
    
    gconDatabase.CommandTimeout = 350
      
    ' Set all option buttons to false at Initialization
    optAdd.Value = False
    optInactive.Value = False
    optReplacePart.Value = False

   'Create Recordsets and Arrays of All Parts and Active Parts
    Call RetrievePartData
    
   'Create Arrays of Stocking Location ID and Description
    Call RetrieveStockingLocationData
    
    frmProcessing.Cls
    
     ' Retrieve the Temporary Part data
     Set mrsTempModelPartsMaster = New ADODB.Recordset
    
     With mrsTempModelPartsMaster
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Source = "SELECT TempMast.part_id, part_description, temp_part_activity_code, " & _
           "quit_ecn_date,quit_ecn_number,quantity,part_id_replaces " & _
           "From V_MNB_Temp_Model_Line_Part_Master TempMast join v_prod_part Part on " & _
           "TempMast.part_id = Part.part_id order by TempMast.part_id"
        .Open
    End With
    
    'Show Listing of Temporary Part records
    If mrsTempModelPartsMaster.RecordCount > 1 Then
        blnEmptyTempFile = False
        frmTemporaryNewPartDisplay.Show vbModal
    Else
        blnEmptyTempFile = True
    End If
    
 'On the first time in, load the part information and models.
    If mrsTempModelPartsMaster.RecordCount > 0 Then
        Call MoveTempModelPartsMaster_MoveComplete
        Call RetrieveModelData
    End If
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, _
        Err.Description)
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Purpose:  Ask the user if he would like to save changes
    '           before closing the form.  If "Yes", call the
    '           mnuFileSave_Click procedure.  If "Cancel",
    '           set the Cancel flag to true and exit.  If no
    '           exit without setting the Cancel flag to
    '           true.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare Return Code variable
    Dim intRetCode As Integer
    
    ' Check to see if any changes have been made to the
    ' recordset.
    If mblnRecChanged Then
    
        ' Ask the user if he would like to save the changes.
        intRetCode = MsgBox("Save Changes?", _
            vbQuestion + vbYesNoCancel, "Closing")
        If intRetCode = vbYes Then
            If Not ValidEntries(blnCancel) Then
                Cancel = True
                GoTo PROC_EXIT
            End If
            
            Call mnuFileSave_Click
            Cancel = False
        ElseIf intRetCode = vbCancel Then
            Cancel = True
        Else
            Cancel = False
        End If
        
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_QueryUnload", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Purpose:  Close and de-reference objects used by this
    '           form
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    
    ' Check to see if the recordset exists.  If so, check to see
    ' if it is open.  If it is open, close it.  De-reference it
    ' by setting it to nothing.
    If gconDatabase.State = adStateOpen Then
        If Not mrsDatabase Is Nothing Then
            If mrsDatabase.State = adStateOpen Then
                mrsDatabase.Close
            End If
            Set mrsDatabase = Nothing
        End If
        If Not mrsTempModelPartsMaster Is Nothing Then
            If mrsTempModelPartsMaster.State = adStateOpen Then
                mrsTempModelPartsMaster.Close
            End If
            Set mrsTempModelPartsMaster = Nothing
        End If
        If Not mrsTempModelPartsDetail Is Nothing Then
            If mrsTempModelPartsDetail.State = adStateOpen Then
                mrsTempModelPartsDetail.Close
            End If
            Set mrsTempModelPartsDetail = Nothing
        End If
        If Not mrsModelPartStockLoc Is Nothing Then
            If mrsModelPartStockLoc.State = adStateOpen Then
                mrsModelPartStockLoc.Close
            End If
            Set mrsModelPartStockLoc = Nothing
        End If
    End If
    
    Unload frmProcessing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Resume Next

End Sub
Private Sub lstAvailable_DblClick()
    ' Purpose:  If an entry in the available column was
    '           selected, call the Add sub.
    
    If lstAvailable.ListIndex > -1 Then
        cmdAdd_Click
    End If
End Sub

Private Sub lstSelected_Click()
    ' Purpose:  When an item in the selected listbox was
    '           clicked, fill the data for that record
    '           and make the fields visible.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    ' Find the current selected Model and Line to be used for part verification.
    strModelNumber = Mid(lstSelected.Text, 1, 15)
    strLineId = Right(lstSelected.Text, 2)
  
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "lstSelected_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuFileCancel()
    ' Purpose:  This procedure will delete a newly added
    '           record or return an updated record to its
    '           original state.
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ' Set up a field object
    Dim fld As ADODB.Field
    
    ' If the record is newly added, delete it.
'    With mrsDatabase
'        If .EditMode = adEditAdd Then
'            .Delete
'            .MoveFirst
'            GoTo PROC_EXIT
'        End If
'
'        ' Loop through the existing record and reset
'        ' each field to it original value.
'        For Each fld In .Fields
'            fld.Value = fld.OriginalValue
'        Next fld
'        Set fld = Nothing
'        mrsDatabase_MoveComplete adRsnMove, Nothing, adStatusOK, _
'            Nothing
'    End With
    
    Call cmdClearAll_Click
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuFileCancel", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuFileClose_Click()
    '  Purpose: Close this form
    
    ' Set up error handling
    On Error GoTo PROC_ERR
   
    Unload Me
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuClose_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuFileDelete_Click()
 'To Delete Temp changes and put parts back to their original status
    
  ' Set up error handling
    On Error GoTo PROC_ERR
    
  'Flag used to determine a delete is in progress.
  blnDeleteRecord = True
    
  'Clears all the Models, and resets parts in mnb_model_part.  Delete the mnb_temp_model_line_part_detail
  '  Record.
    Call cmdClearAll_Click
    
    'Delete the Temporary Master Record using a different recordset.  An attempt to delete the row
    '  in the original recordset was unsuccessful because there is a field in the recordset from another
    '  view that it tries to delete that it should not.
     Dim rsDelTempModelPartsMaster As ADODB.Recordset
     Set rsDelTempModelPartsMaster = New ADODB.Recordset

     With rsDelTempModelPartsMaster
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
'        .LockType = adLockBatchOptimistic
        .LockType = adLockOptimistic
        .Source = "select * From v_mnb_temp_model_line_part_master where part_id = '" & _
            cboNewPart.Text & "'"
         .Open
        .Delete
        End With

        Set rsDelTempModelPartsMaster = Nothing
    
    'Requery the recordset because one was deleted.  Reestablishes bookmarks used.
    With mrsTempModelPartsMaster
        varBookMarkKeep = .Bookmark
        
        .Requery
        
        'If the last record was deleted, reset recordset to the beginning.
        'Otherwise, use the bookmark.
        If .RecordCount <> 0 Then
            If .RecordCount < varBookMarkKeep Then
                .MoveFirst
            Else
                .Bookmark = varBookMarkKeep
            End If
        Else
            blnEmptyTempFile = True
        End If
        
    End With
        
    cboStockingLocation.ListIndex = 0
        
    If Not blnEmptyTempFile Then
        Call MoveTempModelPartsMaster_MoveComplete

        Call cmdRefreshModels_Click
    Else
        Call ClearTempPartScreen
    End If
    
    blnDeleteRecord = False
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuFileDelete_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT

End Sub

Private Sub mnuFileExit_Click()
    ' Purpose:  Exit the application
    Unload frmMain
End Sub

Private Sub mnuFileSave_Click()
    ' Purpose:  Save the current changes to the database
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    'Make changes to the Master Temporary Record only, not Detail Temporary Record.
    blnSaveMasterInfo = True
    Call CreateNewTempPartRecord
    blnSaveMasterInfo = False
    
    ' Check for errors
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuSave_Click", gconDatabase.Errors(0).Description
    End If
    
       
    If mrsModelPartStockLoc Is Nothing Then
    'Make changes to mnb_model_part_stocking_location if Step_Number changed
        Set mrsModelPartStockLoc = New ADODB.Recordset

        With mrsModelPartStockLoc
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
'            .LockType = adLockReadOnly
            .LockType = adLockOptimistic
            .Source = "select * From v_mnb_model_part_stocking_location_all " & _
                "where model_number = '" & strModelNumber & "' and line_id = '" & strLineId & _
                "' and part_id = '" & cboNewPart.Text & "' and stocking_location_id = '" & _
                strStockingLocationID & "'"
            .Open
         End With
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuFileSave_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuHelpAbout_Click()
    ' Purpose:  Display the About form for the application
    
    ' Declare and instantiate Splash screen object.
    Dim clsSplash As MES.Splash
    Set clsSplash = New MES.Splash
    
    clsSplash.Show
    
End Sub

Private Sub mnuHelpContents_Click()
    ' Purpose:  Display Help Contents
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ShowHelpContents
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuHelpContents_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuHelpIndex_Click()
    ' Purpose:  Display Help Index
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ShowHelpIndex
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuHelpContents_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewFirst_Click()
    ' Purpose:  Move to the first record in the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' If any changes have been made, save them before
    ' changing record position.
    If mblnRecChanged Then
        Call SaveChanges
    End If
    
    mblnRetrieveModels = True
        
   ' Find the first Temporary Modified Part record.
    mrsTempModelPartsMaster.MoveFirst
    Call MoveTempModelPartsMaster_MoveComplete
    
    'Load the models for each screen viewing.
    lstAvailable.Clear
    lstSelected.Clear
    cboStockingLocation.ListIndex = 0
    Call RetrieveModelData
    
    mblnRetrieveModels = False
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewFirst_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub mnuViewPrevious_Click()
    ' Purpose:  Move to the previous record.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' If any changes have been made, save them before
    ' changing record position.
    If mblnRecChanged Then
        Call SaveChanges
    End If
    
    mblnRetrieveModels = True
    
    ' if not at the first record, move to the previous
    ' item.  If at the first item, move to the last item.
    With mrsTempModelPartsMaster
        .MovePrevious
        If .BOF Then
            .MoveLast
        End If
    End With
      
    Call MoveTempModelPartsMaster_MoveComplete
    'Load the models for each screen viewing.
    lstAvailable.Clear
    lstSelected.Clear
    cboStockingLocation.ListIndex = 0
    Call RetrieveModelData
    mblnRetrieveModels = False
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewPrevious_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewNext_Click()
      ' Purpose:  Move to the first record in the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
     ' Validate controls before moving
    Call ValidEntries(blnCancel)
    
    If blnCancel Then
        GoTo PROC_EXIT
    End If
    
    ' If any changes have been made, save them before
    ' changing record position.
    If mblnRecChanged Then
        Call SaveChanges
    End If
    
    mblnRetrieveModels = True
    
    ' If not at the last item, move to the next one,
    ' otherwise move to the first item.
    
    With mrsTempModelPartsMaster
        .MoveNext
        If .EOF Then
            .MoveFirst
        End If
    End With
     
    Call MoveTempModelPartsMaster_MoveComplete
    'Load the models for each screen viewing.
    lstAvailable.Clear
    lstSelected.Clear
    cboStockingLocation.ListIndex = 0
    Call RetrieveModelData
     
    mblnRetrieveModels = False
     
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewNext_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewLast_Click()
    ' Purpose:  Move to the last record.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' If any changes have been made, save them before
    ' changing record position.
    If mblnRecChanged Then
        Call SaveChanges
    End If
    
    mblnRetrieveModels = True
    
    ' Move to the last item.
    With mrsTempModelPartsMaster
        .MoveLast
    End With
       
    Call MoveTempModelPartsMaster_MoveComplete
    
    'Load the models for each screen viewing.
    lstAvailable.Clear
    lstSelected.Clear
    cboStockingLocation.ListIndex = 0
    Call RetrieveModelData
    
    mblnRetrieveModels = False
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewLast_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub
Private Sub mnuViewList_Click()
    ' Purpose:  Find a specific record in the database and move
    '            to that record.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
        
    ' Validate controls before moving
    Call ValidEntries(blnCancel)
    
    If blnCancel Then
        GoTo PROC_EXIT
    End If
    
    If mblnRecChanged = True Then
        Call SaveChanges
    End If

    mblnRetrieveModels = True
    
    mrsTempModelPartsMaster.Requery
    
    frmTemporaryNewPartDisplay.Show vbModal
    
    Call MoveTempModelPartsMaster_MoveComplete
        
    lstAvailable.Clear
    lstSelected.Clear
    cboStockingLocation.ListIndex = 0
    Call RetrieveModelData
    
    mblnRetrieveModels = False
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewList_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub
Private Sub RetrievePartData()
    ' Purpose:  Retrieve the part ID's from the table and
    '           build the array of all parts and active parts.
    '           The first Array being built is of all parts in the system, the second
    '           contains only active parts in the Minibill system.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare the recordset variable
    Dim rsPart As ADODB.Recordset
    
    ' Instantiate the recordset
    Set rsPart = New ADODB.Recordset
    
    intArrayCount = 0

    frmProcessing.Label2 = "** Retrieve All Parts List ** " & Time
    frmProcessing.Refresh

    ' Set values of fields
    With rsPart
        'tells the recordset where to get its data from
        'i.e. gcondatabase="Provider=SQLOLEDB.1;PASSWORD=090400;USERID=SCHIMC
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
                .Source = "select * from v_prod_part " & _
                    "order by v_prod_part.part_id asc"
'        '.LockType = adLockBatchOptimistic
        .LockType = adLockReadOnly
        .Open
    End With
'
    ' Check for errors returned from the recordset
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "RetrievepartData", _
            gconDatabase.Errors(0).Description
    End If
    
    ' Using the recordset, load the array.
    With rsPart
        ' Move to the first record.
        .MoveFirst

        ' Reset the size of the arrays.
        ReDim marrstrAllParts(.RecordCount - 1)
        
        ' Loop through the recordset.
        Do While Not .EOF
            marrstrAllParts(intArrayCount) = Trim(!part_id)
            
            ' Move to the next record.
            .MoveNext
            
            ' Increment the array counter
            intArrayCount = intArrayCount + 1
        Loop
        .Close
    End With
    
'    mblnLoadedAllParts = True
    Set rsPart = Nothing
    
    
    ' Declare the recordset variable.  Using the same recordset, but loading only active
    '   parts.
    
    ' Instantiate the recordset
    Set rsPart = New ADODB.Recordset

    frmProcessing.Label2 = "** Retrieve Active Parts List ** " & Time
    frmProcessing.Refresh

    ' Set values of fields
    With rsPart
        'tells the recordset where to get its data from
        'i.e. gcondatabase="Provider=SQLOLEDB.1;PASSWORD=090400;USERID=SCHIMC
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .CursorType = adOpenStatic
        ' Change the literal below to the name of your view
                .Source = "select distinct part_id From v_mnb_model_part " & _
                    "join v_PROD_Line_Stocking_Location on " & _
                    "v_prod_line_stocking_location.line_id = v_mnb_model_part.line_id " & _
                    "join v_mnb_model_line on " & _
                    "v_mnb_model_line.model_number = v_mnb_model_part.model_number " & _
                    "and v_mnb_model_line.line_id = v_mnb_model_part.line_id " & _
                    "order by part_Id"
        .Open
    End With
'
    ' Check for errors returned from the recordset
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "RetrievepartData - Active Parts", _
            gconDatabase.Errors(0).Description
    End If
    
    ' Using the recordset, load the array and the
    ' available listbox.
    With rsPart
        ' Move to the first record.
        .MoveFirst
        intArrayCount = 0

        ' Reset the size of the arrays.
        ReDim marrstrActiveParts(.RecordCount - 1)
        
        ' Loop through the recordset.
        Do While Not .EOF
            marrstrActiveParts(intArrayCount) = Trim(!part_id)
            
            ' Move to the next record.
            .MoveNext
            
            ' Increment the array counter
            intArrayCount = intArrayCount + 1
        Loop
        .Close
    End With
    
'    mblnLoadedAllParts = True
    Set rsPart = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrievePartData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Function ValidEntries(blnCancel As Boolean)
' Purpose:  Validate any fields which have validation routines.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Initialize the variable to hold cancel.
    blnCancel = False
    
    ' If the current record has been deleted, do not perform edits.
    If mrsTempModelPartsMaster.EditMode = adEditDelete Then
        blnCancel = False
        GoTo PROC_EXIT
    End If
    
    Call txtPartQuantity_Validate(blnCancel)
    If blnCancel Then
        GoTo PROC_EXIT
    End If
    
   
    Call mskECNNumber_Validate(blnCancel)
    If blnCancel Then
        GoTo PROC_EXIT
    End If
    
    Call mskTempEndDate_Validate(blnCancel)
    If blnCancel Then
        GoTo PROC_EXIT
    End If
    
    Call txtStepNumber_Validate(blnCancel)
    If blnCancel Then
        GoTo PROC_EXIT
    End If
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    Call ShowError(Me.Name, "ValidEntries", Err.Number, _
        Err.Description)
    blnCancel = True
    GoTo PROC_EXIT

End Function

Sub RetrieveModelData()
    
    ' Purpose:  Fill the Model Available and Selected List boxes.
    
    Dim strCompareListModelToSelected As String
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsSelected As ADODB.Recordset
    Set rsSelected = New ADODB.Recordset

    mblnRetrieveModels = True
    
    'Select Temp_model_line_part_detail records that match the part from the dropdown.
    With rsSelected
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
'        .LockType = adLockBatchOptimistic
        .Source = "select model_number, line_id, stocking_location_id From v_mnb_temp_model_line_part_detail " & _
          "where part_id = '" & cboNewPart.Text & _
          "' order by model_number, line_id,stocking_location_id"
       .Open
        'Find the Stocking Location id and the line.
        If .RecordCount > 0 Then
            optAdd.Enabled = False
            optInactive.Enabled = False
            optReplacePart.Enabled = False
            If cboStockingLocation.ListIndex > 0 Then
                If marrstrStockingLocationId(cboStockingLocation.ListIndex) <> !stocking_location_id Then
                    strStockingLocationID = marrstrStockingLocationId(cboStockingLocation.ListIndex)
                Else
                    strStockingLocationID = !stocking_location_id
                End If
            Else
                'If stocking location not selected from combo for current record
                strStockingLocationID = !stocking_location_id
                strLineId = !line_id
            End If
        Else
            If cboStockingLocation.ListIndex < 1 Then
                cboStockingLocation.Text = cboStockingLocation.List(1)
                strStockingLocationID = marrstrStockingLocationId(1)
                strLineId = Mid(cboStockingLocation.Text, 1, 2)
                If lstSelected.ListCount = 0 Then
                    optAdd.Enabled = True
                    optInactive.Enabled = True
                    optReplacePart.Enabled = True
'                    MsgBox "Please select a Stocking Location ", vbOKOnly, "Select a Stocking Location..."
'                    cboStockingLocation.SetFocus
'                    Exit Sub
                End If
            Else
                strStockingLocationID = marrstrStockingLocationId(cboStockingLocation.ListIndex)
            End If
        End If
    
    End With
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset

    'Set up models for the Available and Selected Lists.  Query is different because of radio button
    '  options.
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
'       .LockType = adLockReadOnly
        .LockType = adLockOptimistic
        
        If optInactive.Value = True Then
         .Source = "select distinct v_mnb_model_part_stocking_location.model_number, " & _
                   "v_mnb_model_part_stocking_location.line_id From v_mnb_model_part_stocking_location " & _
                    "join v_PROD_Line_Stocking_Location on " & _
                    "v_prod_line_stocking_location.line_id = v_mnb_model_part_stocking_location.line_id " & _
                    "join v_mnb_model_line on " & _
                    "v_mnb_model_line.model_number = v_mnb_model_part_stocking_location.model_number " & _
                    "and v_mnb_model_line.line_id = v_mnb_model_part_stocking_location.line_id " & _
                    "where v_mnb_model_part_stocking_location.part_id = '" & strPartId & _
                    "' and v_mnb_model_part_stocking_location.line_id = '" & strLineId & _
                    "' and v_mnb_model_Part_stocking_location.stocking_location_id = '" & _
                    strStockingLocationID & "' order by v_mnb_model_part_stocking_location.model_number, v_mnb_model_part_stocking_location.line_id"
        End If
       
        If optAdd.Value = True Then
        .Source = "select v_mnb_model_line.model_number, v_mnb_model_line.line_id " & _
                    "From v_mnb_model_line " & _
                    "join V_PROD_Line_Stocking_Location on V_PROD_Line_Stocking_Location.line_id = v_mnb_model_line.line_id " & _
                    "where v_mnb_model_line.line_id = '" & strLineId & "' group by v_mnb_model_line.model_number, v_mnb_model_line.line_id " & _
                    "order by v_mnb_model_line.model_number, v_mnb_model_line.line_id"
        End If
        
        If optReplacePart.Value = True Then
            .Source = "select distinct v_mnb_model_part_stocking_location.model_number, v_mnb_model_part_stocking_location.line_id " & _
                "From v_mnb_model_part_stocking_location join v_PROD_Line_Stocking_Location on " & _
                "v_prod_line_stocking_location.line_id = v_mnb_model_part_stocking_location.line_id " & _
                "join v_mnb_model_line on v_mnb_model_line.model_number = v_mnb_model_part_stocking_location.model_number and " & _
                "v_mnb_model_line.line_id = v_mnb_model_part_stocking_location.line_id " & _
                "where v_mnb_model_part_stocking_location.part_id = '" & cboReplacePart.Text & _
                "' and v_mnb_model_part_stocking_location.line_id = '" & strLineId & _
                "' and v_mnb_model_part_stocking_location.stocking_location_id = '" & strStockingLocationID & _
                "' order by v_mnb_model_part_stocking_location.model_number, v_mnb_model_part_stocking_location.line_id"

        End If
        
       .Open
       If Not optAdd.Value = True Then
        If cboNewPart.Visible Then
            If .RecordCount = 0 Then
                MsgBox "Part " & strPartId & " not found in any models" & vbCrLf & _
                "for Line " & strLineId & " and stocking location " & Trim(cboStockingLocation.Text) & "." & _
                vbCrLf & "Make another selection.", _
                    vbOKOnly, "Part Not in any Models..."
            End If
        End If
       End If
        
        
        'Go through the models found to fill the list boxes.  Determines which are available and which have
        'already been selected.
        Do While Not .EOF
            If Not rsSelected.EOF Then
            'Set up value for Case statement.  Use Case statement to determine the comparison
            'between the list of all models in that line to the list of models already having
            'a temporary setup.
                If Trim(!model_number) > Trim(rsSelected!model_number) Then
                    strCompareListModelToSelected = "Greater"
                End If
        
                If Trim(!model_number) = Trim(rsSelected!model_number) Then
                    strCompareListModelToSelected = "Equal"
                End If
        
                If Trim(!model_number) < Trim(rsSelected!model_number) Then
                    strCompareListModelToSelected = "Less"
                End If
            
                Select Case strCompareListModelToSelected
                    
                    Case Is = "Greater"
                        strStockingLocationID = rsSelected!stocking_location_id
                        Call FindStockingLocationDesc
                        lstSelected.AddItem Mid(rsSelected!model_number, 1, 15) & "   " & _
                            Mid(rsSelected!line_id, 1, 2) & "  " & Mid(strStockingLocationDesc, 5, 35)
                        rsSelected.MoveNext
 
                    Case Is = "Equal"
                        If Trim(!line_id) = Trim(rsSelected!line_id) Then
                            strStockingLocationID = rsSelected!stocking_location_id
                            Call FindStockingLocationDesc
                            lstSelected.AddItem Mid(!model_number, 1, 15) & "   " & _
                                Mid(!line_id, 1, 2) & "  " & Mid(strStockingLocationDesc, 5, 35)
                            .MoveNext
                            rsSelected.MoveNext
                        Else
                            lstAvailable.AddItem Mid(!model_number, 1, 15) & "   " & Mid(!line_id, 1, 2)
                            .MoveNext
                        End If
                    
                    Case Is = "Less"
                        lstAvailable.AddItem Mid(!model_number, 1, 15) & "   " & Mid(!line_id, 1, 2)
                        .MoveNext
                    
                    End Select
                    
            End If
            
        'This will put models in available list after Selected list has been processed.
            If rsSelected.EOF Then
                Do While Not .EOF
                    lstAvailable.AddItem Mid(!model_number, 1, 15) & "   " & Mid(!line_id, 1, 2)
                    .MoveNext
                Loop
            End If
            
        Loop
        .Close
        
     ' This will put models in Selected List after All Available models have been processed.
        With rsSelected
            Do While Not .EOF
                strStockingLocationID = rsSelected!stocking_location_id
                Call FindStockingLocationDesc
                lstSelected.AddItem Mid(!model_number, 1, 15) & "   " & strStockingLocationDesc
                .MoveNext
            Loop
        .Close
        End With
        
    End With
    
    Set rsList = Nothing
    Set rsSelected = Nothing

    cmdAdd.Enabled = True
    cmdRemove.Enabled = True
    cmdClearAll.Enabled = True
    txtPart.Enabled = True
'    cmdPrint.Enabled = True
    txtPartQuantity.Enabled = True
    txtStepNumber.Enabled = True
    
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    cmdList.Enabled = True
    
    mblnRetrieveModels = False
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveModelData", _
        Err.Number, Err.Description)
    
End Sub
Private Sub MoveTempModelPartsMaster_MoveComplete()

' Purpose:  Fill the fields after a move in the
    '           recordset is complete.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim strDate As String
 
    ' If the recordset is at a valid record, fill the
    ' controls.
    With mrsTempModelPartsMaster
      If Not .EOF And Not .BOF Then
        If Not .EditMode = adEditAdd Then
            ' Add code here to fill the controls on the form
            ' with the data from the Recordset.
            If !temp_part_activity_code = "A" Then
                optAdd.Value = True
            End If
            If !temp_part_activity_code = "R" Then
                optReplacePart.Value = True
            End If
            If !temp_part_activity_code = "I" Then
                optInactive.Value = True
            End If
            
            If !temp_part_activity_code = "R" Then
                If IsNull(!part_id_replaces) Then
                    MsgBox "Please Select Replaces Part", vbOKOnly, "Missing Replacement Part"
                End If
            End If
            
            lblNewPartDesc.Caption = !part_description
            strDate = Mid(!quit_ecn_date, 1, 4) & "/" & Mid(!quit_ecn_date, 5, 2) & _
                    "/" & Mid(!quit_ecn_date, 7, 2)
            mskTempEndDate.Text = Format(strDate, "mm/dd/yyyy")
            cboNewPart.Text = !part_id
            cboNewPart.SelStart = 0
            cboNewPart.SelLength = Len(!part_id)
            cboNewPart.SelText = RTrim(!part_id)

            If IsNull(!part_id_replaces) Then
                cboReplacePart.Text = ""
            Else
                cboReplacePart.Text = !part_id_replaces
                cboReplacePart.SelStart = 0
                cboReplacePart.SelLength = Len(cboReplacePart.Text)
                cboReplacePart.SelText = RTrim(cboReplacePart.Text)
                strPartId = !part_id_replaces
                Call FindDescription
                lblReplacePartDesc.Caption = strPartDescription
            End If
            txtPartQuantity.Text = !quantity
            mskECNNumber.Mask = "########"
            mskECNNumber.Text = "________"
            If Len(RTrim(!quit_ecn_number)) = 6 Then
                mskECNNumber.Text = RTrim(!quit_ecn_number) & "__"
            End If
            
            If Len(RTrim(!quit_ecn_number)) = 7 Then
                mskECNNumber.Text = RTrim(!quit_ecn_number) & "_"
            End If
'            If gblnUpdate Then
'                mskObsoleteDate.Enabled = !stocking_location_obsolete_flag
'                cmdObsoleteDateCalendar.Enabled = mskObsoleteDate.Enabled
'            End If
'            If .EditMode = adEditAdd Then
'                cmdCancel.ToolTipText = "Cancel Add"
'                txtLocationID.Enabled = True
'            Else
'                cmdCancel.ToolTipText = "Cancel Update of Current Entry"
'                txtLocationID.Enabled = False
'            End If
        
            ' Set focus to the color ID field
'            If Screen.ActiveForm Is Me And gblnUpdate Then
'            If Screen.ActiveForm Is Me Then
'                If .EditMode = adEditNone Then
'                    cboNewPart.SetFocus
'                Else
'                    txtLocationDescription.SetFocus
'                End If
'            End If
          End If
        End If
    
'Resets variable strPartId to run queries
        If Not .EOF And Not .BOF Then
            strPartId = RTrim(!part_id)
        End If

    End With
    
    cboNewPart.Enabled = False
    
    cboReplacePart.Enabled = False

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "MoveTempModelPartsMaster_MoveComplete", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
    
End Sub



Private Sub mskECNNumber_GotFocus()
    'Select the field for easy update
    mskECNNumber.SelStart = 0
    mskECNNumber.SelLength = 8
End Sub

Private Sub mskECNNumber_Validate(blnCancel As Boolean)
'Verify some valid ECN Number is filled in.  Cancel refers to the boolean in
'this subroutine, not a Cancel button.
 
 ' Set up error handling
    On Error GoTo PROC_ERR
    
    If Len(mskECNNumber.ClipText) < 6 Then
        blnCancel = True
        MsgBox "Must show 6-digit ECN part is set up for", vbOKOnly, "ECN Error"
        mskECNNumber.SetFocus
        Exit Sub
    End If
    
    If mskECNNumber.ClipText = 0 Then
        blnCancel = True
        MsgBox "ECN Number must be valid", vbOKOnly, "Quantity Error"
        mskECNNumber.SetFocus
        Exit Sub
    End If
    
    If Not blnEmptyTempFile Then
        If RTrim(mrsTempModelPartsMaster!quit_ecn_number) <> mskECNNumber.ClipText Then
            mblnRecChanged = True
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mskECNNumber_Validate", _
        Err.Number, Err.Description)
End Sub

Private Sub mskTempEndDate_GotFocus()
' Purpose:  Select the field for easy update
    mskTempEndDate.SelStart = 0
    mskTempEndDate.SelLength = 10
End Sub

Private Sub mskTempEndDate_Validate(blnCancel As Boolean)
 ' Purpose:  Validate the field to make sure that it is
    '           either a valid date or empty.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    If Len(mskTempEndDate.ClipText) = 0 Then
        blnCancel = True
        MsgBox "Invalid Temporary End Date Entered!", _
            vbExclamation + vbOKOnly, _
            "Temporary End Date Validation Error"
        mskTempEndDate.SelStart = 0
        mskTempEndDate.SelLength = 10
        cmdTempEndDate.SetFocus
        Exit Sub
    End If
    
    If Format(dteToday, "yyyymmdd") > Format(mskTempEndDate.Text, "yyyymmdd") Then
        blnCancel = True
         MsgBox "Date must be Future", vbOKOnly, "Verify Date"
        cmdTempEndDate.SetFocus
        Exit Sub
    End If
    
'    If Format(mskTempEndDate.Text, "yyyymmdd") > Format(dteWeekTwoFuture, "yyyymmdd") Then
'        blnCancel = True
'        MsgBox "Date can be no more than 1 month" & vbCrLf & "in the future", _
'        vbOKOnly, "Verify Date"
'        cmdTempEndDate.SetFocus
'        Exit Sub
'    End If

    If Not blnEmptyTempFile Then
        If RTrim(mrsTempModelPartsMaster!quit_ecn_date) <> Format(mskTempEndDate.Text, "yyyymmdd") Then
            mblnRecChanged = True
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mskTempEndDate_Validate", _
        Err.Number, Err.Description)
End Sub

Private Sub optAdd_Click()
'Load all parts array to the parts dropdown.

    ' Set up error handling
    On Error GoTo PROC_ERR
   
    If mblnRecChanged Then
        Call SaveChanges
    End If
           
    strActivityCode = "A"
    mblnLineChanged = False
    
    'Clear the available products if the Parts dropdown has been changed.
    Call ClearTempPartScreen
    
    If mblnAddTempRec = True Then
        cboNewPart.Clear
    End If

    'Load parts array to the dropdown
    For intArrayCount = LBound(marrstrAllParts) To UBound(marrstrAllParts)
        cboNewPart.AddItem marrstrAllParts(intArrayCount)
    Next
    
    cboReplacePart.Visible = False
    Label4.Visible = False
    lblReplacePartDesc.Caption = " "
    
    cboNewPart.Visible = True
    cboNewPart.Enabled = True
    cboStockingLocation.Enabled = True
    Label3.Visible = True
    cmdRefreshModels.Enabled = True
    txtPartQuantity.Enabled = True
    txtStepNumber.Enabled = True
    cmdTempEndDate.Enabled = True
    mskECNNumber.Enabled = True
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "optAdd_Click", _
        Err.Number, Err.Description)
        
End Sub

Private Sub optInactive_Click()
'Load all parts array to the parts dropdown.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
       
    strActivityCode = "I"
    mblnLineChanged = False
    
    '     Clear the available products when the Parts dropdown has been changed.
'    Call ClearTempPartScreen
    
    If mblnAddTempRec = True Then
        cboNewPart.Clear
    End If

    'Load Active Parts array to the dropdown
    For intArrayCount = LBound(marrstrActiveParts) To UBound(marrstrActiveParts)
        cboNewPart.AddItem marrstrActiveParts(intArrayCount)
    Next
    
    cboReplacePart.Visible = False
    Label4.Visible = False
    lblReplacePartDesc.Visible = False
    
    cboNewPart.Visible = True
    cboNewPart.Enabled = True
    cboStockingLocation.Visible = True
    cboStockingLocation.Enabled = True
    Label3.Visible = True
    cmdRefreshModels.Enabled = True
    txtPartQuantity.Enabled = True
    txtStepNumber.Enabled = True
    cmdTempEndDate.Enabled = True
    mskECNNumber.Enabled = True
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "optInactive_Click", _
        Err.Number, Err.Description)
        
End Sub

Private Sub optReplacePart_Click()
    On Error GoTo PROC_ERR
'    Call SaveChanges
    
    strActivityCode = "R"
    mblnLineChanged = False
'  Clear fields on current screen when the Parts dropdown has been changed.
    Call ClearTempPartScreen
    
    If mblnAddTempRec = True Then
        cboNewPart.Clear
        cboReplacePart.Clear
    End If
    
    'Load All Parts array to the New Part dropdown
    For intArrayCount = LBound(marrstrAllParts) To UBound(marrstrAllParts)
        cboNewPart.AddItem marrstrAllParts(intArrayCount)
    Next
    
    'Load Active Parts array to the Replace Parts dropdown
    For intArrayCount = LBound(marrstrActiveParts) To UBound(marrstrActiveParts)
        cboReplacePart.AddItem marrstrActiveParts(intArrayCount)
    Next
    
    'Make the combo and text boxes available for the Replace Parts Controls
    cboReplacePart.Visible = True
    cboReplacePart.Enabled = True
    Label4.Visible = True
    lblReplacePartDesc.Visible = True
    lblReplacePartDesc.Caption = " "
    
    cboNewPart.Visible = True
    cboNewPart.Enabled = True
    cboStockingLocation.Visible = True
    cboStockingLocation.Enabled = True
    Label3.Visible = True
    cmdRefreshModels.Enabled = True
    txtPartQuantity.Enabled = True
    txtStepNumber.Enabled = True
    cmdTempEndDate.Enabled = True
    mskECNNumber.Enabled = True
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "optReplacePart_Click", _
        Err.Number, Err.Description)
End Sub
Private Sub txtPart_Change()
    Dim lngIndex As Long
    
    For lngIndex = 0 To lstAvailable.ListCount - 1
        If Mid(lstAvailable.List(lngIndex), 1, Len(txtPart.Text)) >= Trim(txtPart.Text) Then
            lstAvailable.ListIndex = lngIndex
            lstAvailable.TopIndex = lngIndex
            Exit For
        End If
    Next lngIndex
End Sub

Private Sub txtPart_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub RetrieveStockingLocationData()
    ' Purpose:  Fill the Stocking Location Combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim RsStockingLocation As ADODB.Recordset
    Set RsStockingLocation = New ADODB.Recordset
    
    frmProcessing.Label2 = "** Retrieve Stocking Loc. List ** " & Time
    frmProcessing.Refresh
    
    With RsStockingLocation
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "SELECT linestock.line_id, stocking_location_description, " & _
           "linestock.stocking_location_id FROM V_PROD_Line_Stocking_Location LineStock " & _
           "join V_PROD_Stocking_Location Stock on " & _
           "LineStock.stocking_location_id = stock.stocking_location_id " & _
           "order by Line_id, stocking_location_description"
        .Open
        
        intArrayCount = 0
        ReDim marrstrStockingLocationId(.RecordCount)
        ReDim marrstrStockingLocationDesc(.RecordCount)
        
        cboStockingLocation.AddItem "--Select Stocking Location", 0
        marrstrStockingLocationDesc(intArrayCount) = "--Select Stocking Location"
        marrstrStockingLocationId(intArrayCount) = "--"
        intArrayCount = 1
        
        Do While Not .EOF
            cboStockingLocation.AddItem !line_id & "  " & !stocking_location_description
            marrstrStockingLocationDesc(intArrayCount) = !line_id & "   " & !stocking_location_description
            marrstrStockingLocationId(intArrayCount) = !stocking_location_id
            .MoveNext
            intArrayCount = intArrayCount + 1
        Loop
        .Close
    End With
    Set RsStockingLocation = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveStockingLocationData", _
        Err.Number, Err.Description)
    
End Sub

Private Sub txtPartQuantity_KeyPress(KeyAscii As Integer)
    'Using the ascii code for numbers 0-9 to verify only numbers have been entered.
    'Cause a backspace to erase the illegal character.
    If KeyAscii < 48 Or KeyAscii > 57 Then
        MsgBox "Quantity is any number greater than zero", vbOKOnly, "Quantity Error"
        Exit Sub
    End If
End Sub

Private Sub txtPartQuantity_Validate(blnCancel As Boolean)
'Verify some valid quantity is filled in.  Cancel refers to the boolean in this
'  subroutine, not a Cancel button.
    
Dim msgReturn As Integer
    
    ' Set up error handling
    On Error GoTo PROC_ERR

    If Len(txtPartQuantity) = 0 Then
        blnCancel = True
        MsgBox "Must fill in a quantity", vbOKOnly, "Quantity Error"
        txtPartQuantity.SetFocus
        Exit Sub
    End If
    
    If txtPartQuantity.Text = 0 Then
        blnCancel = True
        MsgBox "Part quantity must be greater than zero", vbOKOnly, "Quantity Error"
        txtPartQuantity.SetFocus
        Exit Sub
    End If

    If txtPartQuantity.Text > 15 Then
        msgReturn = MsgBox("Quantity of " & txtPartQuantity.Text & " requested. " & vbCrLf & _
        "Are you sure?", vbYesNo + vbExclamation + vbDefaultButton2)
        If msgReturn = vbNo Then
            txtPartQuantity.SetFocus
            Exit Sub
        End If
    End If

    If Not blnEmptyTempFile Then
        If mrsTempModelPartsMaster!quantity <> txtPartQuantity.Text Then
            mblnRecChanged = True
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "txtPartQuantity_Validate", _
        Err.Number, Err.Description)
    
End Sub

Private Sub CreateNewTempPartRecord()
    Dim mrsTempMasterUpdate As New ADODB.Recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
        
    ' Add the record to the recordset and load the fields.  For Temporary Part Changes
    If mrsTempModelPartsMaster Is Nothing Then
        Set mrsTempModelPartsMaster = New ADODB.Recordset
    End If
        
         ' Retrieve the Temporary Part data
     Set mrsTempMasterUpdate = New ADODB.Recordset
    
     With mrsTempMasterUpdate
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Source = "SELECT TempMast.part_id, temp_part_activity_code, quit_ecn_date , " & _
            "quit_ecn_number, quantity, part_id_replaces From V_MNB_Temp_Model_Line_Part_Master " & _
            "TempMast where part_id = '" & cboNewPart.Text & "'"
        .Open
    End With
        
    With mrsTempMasterUpdate
        'Allow an add to occur for the first record
        If .RecordCount = 0 Then
            mblnAddTempRec = True
            blnEmptyTempFile = False
        End If
        
        ' Add or update new master record
       If mblnAddTempRec Then
            .AddNew
            !part_id = cboNewPart.Text
            !quantity = txtPartQuantity.Text
            !quit_ecn_number = mskECNNumber.ClipText
            !quit_ecn_date = Mid(mskTempEndDate.ClipText, 5, 4) & _
                    Mid(mskTempEndDate.ClipText, 1, 2) & _
                    Mid(mskTempEndDate.ClipText, 3, 2)
            If optReplacePart.Value = True Then
                !part_id_replaces = cboReplacePart.Text
            Else
                !part_id_replaces = Null
            End If
            !temp_part_activity_code = strActivityCode
            .Update
        Else
            !quantity = txtPartQuantity.Text
            !quit_ecn_number = mskECNNumber.ClipText
            !quit_ecn_date = Mid(mskTempEndDate.ClipText, 5, 4) & _
                    Mid(mskTempEndDate.ClipText, 1, 2) & _
                    Mid(mskTempEndDate.ClipText, 3, 2)
            !temp_part_activity_code = strActivityCode
            If optReplacePart.Value = True Then
                !part_id_replaces = Trim(cboReplacePart.Text)
            Else
                !part_id_replaces = Null
            End If
            .Update
        End If
        
        ' Check for errors
        ' Only show the Confirming MsgBox if the cmdSave button was clicked.  The save
        '     also occurs when models are moved from the Available box to the Selected box.
        '     It is not necessary to show the confirming box then.
        If blnSaveMasterInfo Then
            If gconDatabase.Errors.Count > 0 Then
                Err.Raise gconDatabase.Errors(0).NativeError, _
                    "CreateNewTempPartRecord - Master", gconDatabase.Errors(0).Description
            Else: MsgBox "Changes completed Successfully!", vbExclamation, "Successful save - Top Box"
            End If
        Else
            If gconDatabase.Errors.Count > 0 Then
                Err.Raise gconDatabase.Errors(0).NativeError, _
                    "CreateNewTempPartRecord - Master", gconDatabase.Errors(0).Description
            End If
        End If

    End With
    
'Do not update the detail if the cmdSave button was clicked.
If Not blnSaveMasterInfo Then
    If mrsTempModelPartsDetail Is Nothing Then
        Set mrsTempModelPartsDetail = New ADODB.Recordset
    Else
        mrsTempModelPartsDetail.Close
    End If
    
        With mrsTempModelPartsDetail
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Source = "SELECT * FROM V_MNB_Temp_Model_Line_Part_Detail where part_id = '" & _
            cboNewPart.Text & "' and temp_Part_activity_code = '" & strActivityCode & _
            "' and model_number = '" & strModelNumber & "' and line_id = '" & strLineId & "'"
            .Open

    strStockingLocationID = marrstrStockingLocationId(cboStockingLocation.ListIndex)

        If .RecordCount = 0 Then
            .AddNew
            !part_id = cboNewPart.Text
            !temp_part_activity_code = strActivityCode
            !stocking_location_id = strStockingLocationID
            !model_number = strModelNumber
            !line_id = strLineId
            .Update
        End If
        
        ' Check for errors
        If gconDatabase.Errors.Count > 0 Then
            Err.Raise gconDatabase.Errors(0).NativeError, _
                "CreateNewTempPartRecord - Detail", gconDatabase.Errors(0).Description
        End If
        
        .Requery
        
    End With
    
End If
    
    optAdd.Enabled = False
    optInactive.Enabled = False
    optReplacePart.Enabled = False
    cboNewPart.Enabled = False
    cboReplacePart.Enabled = False
    mblnRecChanged = False
    mblnAddTempRec = False


PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "CreateNewTempPartRecord", _
        Err.Number, Err.Description)
End Sub

Private Sub SaveChanges()
'Check to see if changes need to be saved.  Only affects ECN Number, Quantity, and Effective Date.

' Set up error handling
 On Error GoTo PROC_ERR

Dim intRetCode As Integer
    
    
    ' Check to see if any changes have been made to the
    ' recordset.
    If mblnRecChanged Then
        ' Ask the user if he would like to save the changes.
        intRetCode = MsgBox("Save Changes?", _
            vbQuestion + vbYesNoCancel, "Change Quantity, ECN Number, Date")
        If intRetCode = vbYes Then
            Call mnuFileSave_Click
        ElseIf intRetCode = vbNo Then
            mrsTempModelPartsMaster.CancelBatch
            mblnRecChanged = False
            MsgBox "Quantity, ECN, and Date Changes not updated "
            GoTo PROC_EXIT
        ElseIf intRetCode = vbCancel Then
            GoTo PROC_EXIT
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "SaveChanges", _
        Err.Number, Err.Description)
End Sub

Sub ClearTempPartScreen()
'Clear out the fields for the new part

' Set up error handling
 On Error GoTo PROC_ERR

        cboNewPart.Text = ""
        cboReplacePart.Text = ""
        mskECNNumber.Text = "________"
    
        dlgCalendar.mdteSelectedDate = dteWeekFuture
        mskTempEndDate.Text = Format( _
            dlgCalendar.mdteSelectedDate, "mm/dd/yyyy")
        lstAvailable.Clear
        lstSelected.Clear
        txtPartQuantity.Text = ""
        txtPart.Text = ""
        txtStepNumber.Text = ""
        cboStockingLocation.Text = cboStockingLocation.List(0)
        lblNewPartDesc.Caption = ""
        lblReplacePartDesc.Caption = ""
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "ClearTempPartScreen", _
        Err.Number, Err.Description)
End Sub

Private Sub txtStepNumber_KeyPress(KeyAscii As Integer)
'Using the ascii code for numbers 0-9 to verify only numbers have been entered.
    'Cause a backspace to erase the illegal character.
    If Len(txtStepNumber) > 3 Then
        MsgBox "Only 3 Characters for Step Number", vbExclamation, "Step Number Error"
    End If
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
        MsgBox "The Step number is any number greater than zero", vbOKOnly, "Step Number Error"
        Exit Sub
    End If
End Sub

Private Sub txtStepNumber_Validate(Cancel As Boolean)
'Entry must contain leading zeros.  If not entered, then must put them in.
    
    If Len(txtStepNumber) = 1 Then
        txtStepNumber.Text = "00" & Mid(txtStepNumber.Text, 1, 1)
    End If
    
    If Len(txtStepNumber) = 2 Then
        txtStepNumber.Text = "0" & Mid(txtStepNumber.Text, 1, 2)
    End If
    
    If Len(txtStepNumber) > 3 Then
        MsgBox "Only 3 Characters for Step Number", vbExclamation, "Step Number Error"
    End If
    
End Sub
