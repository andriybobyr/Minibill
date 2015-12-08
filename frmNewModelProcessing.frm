VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmNewModelProcessing 
   Caption         =   "MiniBill - New Model Processing"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
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
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   787
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   9
      Top             =   5820
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      Caption         =   "Action To Be Taken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   5280
      TabIndex        =   1
      Top             =   1860
      Width           =   6315
      Begin VB.ComboBox cboLine 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdProcess 
         Caption         =   "&Process"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   7
         Top             =   2400
         Width           =   1755
      End
      Begin VB.OptionButton optMiniBillMaintenance 
         Caption         =   "Go To MiniBill Maintenance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   4395
      End
      Begin VB.OptionButton optCopyModel 
         Caption         =   "Copy From Line:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1515
      End
      Begin VB.OptionButton optUseDefaults 
         Caption         =   "Set Up Parts With Default Locations"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   4335
      End
      Begin VB.ComboBox cboModel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   840
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Model:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGNewModels 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   5953
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Model"
      Columns(0).DataField=   "model_number"
      Columns(0).DataWidth=   15
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Scheduled"
      Columns(1).DataField=   "start_date"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Line"
      Columns(2).DataField=   "line_id"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3810"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3704"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2170"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2064"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1402"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1296"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
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
      _StyleDefs(52)  =   ":id=34,.wraptext=-1"
      _StyleDefs(53)  =   "Named:id=35:Footing"
      _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(55)  =   "Named:id=36:Selected"
      _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=37:Caption"
      _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(59)  =   "Named:id=38:HighlightRow"
      _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(61)  =   "Named:id=39:EvenRow"
      _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(63)  =   "Named:id=40:OddRow"
      _StyleDefs(64)  =   ":id=40,.parent=33"
      _StyleDefs(65)  =   "Named:id=41:RecordSelector"
      _StyleDefs(66)  =   ":id=41,.parent=34"
      _StyleDefs(67)  =   "Named:id=42:FilterBar"
      _StyleDefs(68)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNewModelProcessing.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   540
      TabIndex        =   8
      Top             =   180
      Width           =   9735
   End
End
Attribute VB_Name = "frmNewModelProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsModels As ADODB.Recordset
Private mrsDatabase As ADODB.Recordset
Private mxarrModelsToCopy As XArrayDB

Private Sub cboLine_Click()

Dim strStartingRow As Integer
    
    If cboLine.ListIndex = -1 Then
        Exit Sub
    End If
    
    strStartingRow = mxarrModelsToCopy.Find(1, 2, cboLine.Text)
    cboModel.Clear
    
    Do While strStartingRow <= mxarrModelsToCopy.UpperBound(1)
        If mxarrModelsToCopy(strStartingRow, 2) <> cboLine.Text Then
            Exit Do
        End If
        cboModel.AddItem mxarrModelsToCopy(strStartingRow, 1)
        strStartingRow = strStartingRow + 1
    Loop
    cboModel.ListIndex = 0
    cboModel.Enabled = True
End Sub

Private Sub cboLine_GotFocus()
    cboLine.SelStart = 0
    cboLine.SelLength = Len(cboLine.Text)
End Sub

Private Sub cboLine_Validate(Cancel As Boolean)
    If Len(cboLine.Text) = 1 Then
        cboLine.Text = "0" & cboLine.Text
    End If
    cboFindFirst cboLine
    
    If Len(cboLine.Text) > 0 And cboLine.ListIndex = -1 Then
        MsgBox "A valid line must be selected."
        Cancel = True
    End If
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
cboFindFirst cboModel
If Len(cboModel.Text) > 0 And cboModel.ListIndex = -1 Then
        MsgBox "A valid Model must be selected."
        Cancel = True
End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdProcess_Click()
    ' Purpose:  Process requested model.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Me.MousePointer = vbHourglass
    
    ' If the user selected the radio button to Use default parts to set up the model,
    ' data will be copied from any parts set up as defaults.
    If optUseDefaults.Value Then
        UseDefaults
        
    ElseIf optCopyModel.Value Then
        CopyModel
        
    ' If the user selects the radio button to go to Minibill Maintenance, the screen will be
    ' displayed for the model / line selected.
    ElseIf optMiniBillMaintenance.Value Then
        With frmMiniBillMaintenance
            Me.MousePointer = vbDefault
            .cboLine.Text = Me.TDBGNewModels.Columns(2).Value
            ' if the line associated with the model does not have any locations set up,
            ' a message will be displayed (by the Minibill Maintenance) and no further
            ' processing will take place.  Otherwise, continue with displaying the form.
            If .mrsPartLocation.State = adStateOpen Then
                .cboModel.Text = Me.TDBGNewModels.Columns(0).Value
                .cmdDisplay_Click
                .Show vbModal
            End If
            If gblnMaintPassedUpdates Then
                ' Requery the models
                'RequeryModels
                Me.TDBGNewModels.Delete
                gblnMaintPassedUpdates = False
            End If
        End With
    Else
        MsgBox ("Please select an option ")
            Me.MousePointer = vbDefault
            Exit Sub
    End If
    
    Me.MousePointer = vbDefault
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Me.MousePointer = vbDefault
    Call ShowError(Me.Name, "cmdProcess_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub Form_Load()
    ' Purpose:  Build form with new models.
    
    Dim intCounter As Integer
    
    ' Set up error handling.
    On Error GoTo PROC_ERR
    
    ' If the connection to SQL has not yet been opened, open it.
    If gconDatabase Is Nothing Then
        Set gconDatabase = gclsSQLServer.Connect( _
            gclsMESApplication.ApplicationRole, _
            gclsMESApplication.ApplicationPassword)
            
        If gconDatabase.State <> adStateOpen Then
            Err.Raise gconDatabase.Errors(0).NativeError, "Form_Load", _
                gconDatabase.Errors(0).Description
        End If
    End If
    
    gconDatabase.CommandTimeout = 350
           
    frmProcessing.Label1 = "Gathering New Models... Please Wait"
    frmProcessing.Label2 = "Get Configuration Models List at " & Time
    frmProcessing.Show
           
    DoEvents
         
    ' Instantiate and open a recordset to display models which have not yet been set up with locations.
    Set mrsModels = New ADODB.Recordset
    
    With mrsModels
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Source = "select v_mnb_model_part.model_number, min(start_date) as start_date, " & _
            "v_mnb_model_part.line_id, " & _
            "sum(case when stocking_location_id is null then 0 else 1 end) as locations_set_up " & _
            "From v_mnb_model_part " & _
            "join v_mnb_model_line on v_mnb_model_part.model_number = v_mnb_model_line.model_number  " & _
            "and v_mnb_model_part.line_id = v_mnb_model_line.line_id " & _
            "left outer join v_prod_schedule on v_mnb_model_part.model_number = v_prod_schedule.model_number " & _
            "and v_mnb_model_part.line_id = v_prod_schedule.line_id " & _
            "left outer join v_mnb_model_part_stocking_location " & _
            "on v_mnb_model_part.model_number = v_mnb_model_part_stocking_location.model_number and " & _
            "v_mnb_model_part.part_id = v_mnb_model_part_stocking_location.part_id and " & _
            "v_mnb_model_part.line_id = v_mnb_model_part_stocking_location.line_id and " & _
            "v_mnb_model_part.part_sequence_number = v_mnb_model_part_stocking_location.part_sequence_number " & _
            "group by v_mnb_model_part.model_number, v_mnb_model_part.line_id " & _
            "order by v_mnb_model_part.line_id, v_mnb_model_part.model_number"
        .Open
                
        ' Clear the model and Line combo boxes.
        cboModel.Clear
        cboLine.Clear
        
        Dim intModels As Integer
        Dim strSaveModel As String
        Dim strSaveLine As String
        
        intModels = 0
        
        frmProcessing.Label2 = "Load Models and Shifts to Configure Grid at " & Time
        frmProcessing.Refresh
        
        Set mxarrModelsToCopy = New XArrayDB
        ' Load any models which have been set up into the combo box of models.  This
        ' combo box will be used for a selection list of models from which location info
        ' may be copied.
        Do While Not .EOF
            If !locations_set_up > 0 Then
                If intModels = 0 Then
                    mxarrModelsToCopy.ReDim 1, 1, 1, 2
                Else
                    mxarrModelsToCopy.AppendRows
                End If
                intModels = intModels + 1
                mxarrModelsToCopy(intModels, 1) = Trim(!model_number)
                mxarrModelsToCopy(intModels, 2) = .Fields("line_id").Value
                
                If mxarrModelsToCopy(intModels, 2) <> strSaveLine Then
                    intCounter = intCounter + 1
                    cboLine.AddItem mxarrModelsToCopy(intModels, 2)
                    cboLine.ItemData(cboLine.NewIndex) = intCounter
                    strSaveLine = mxarrModelsToCopy(intModels, 2)
                End If
            End If
            
            .MoveNext
        Loop
 
 ' Allow options to be selected as long as there are models with no parts configured.
        If intModels = 0 Then
            optCopyModel.Enabled = False
            cboModel.Enabled = False
            cboLine.Enabled = False
        Else
            optCopyModel.Enabled = True
            strSaveModel = vbNullString
            For intModels = 1 To mxarrModelsToCopy.UpperBound(1)
                If mxarrModelsToCopy(intModels, 1) <> strSaveModel Then
                    cboModel.AddItem mxarrModelsToCopy(intModels, 1)
                    cboModel.ItemData(cboModel.NewIndex) = intModels
                    strSaveModel = mxarrModelsToCopy(intModels, 1)
                End If
            Next intModels
        End If
        
        ' Filter out models which have been set up so that only new models will be displayed
        ' in the grid.
        .Filter = "locations_set_up = 0"
    End With
    
    ' Associate the grid with the model recordset and bind.
    Set TDBGNewModels.DataSource = mrsModels
    TDBGNewModels.ReBind

    cboLine.Enabled = False
    cboModel.Enabled = False

    frmProcessing.Hide
    DoEvents
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, Err.Description)
    Unload Me
    GoTo PROC_EXIT
End Sub


Private Sub UseDefaults()
    ' Purpose:  Set up parts for the model based on the defaults.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set up a variable to save the number of parts processed.
    Dim lngRecords As Long
    
    ' Set up and open a recordset with default parts that would apply to the model selected.
    Dim rsDefaults As ADODB.Recordset
    Set rsDefaults = New ADODB.Recordset
    
    With rsDefaults
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select v_mnb_part_line_stocking_location.part_id, " & _
            "v_mnb_part_line_stocking_location.line_id, " & _
            "v_mnb_model_part.part_sequence_number, " & _
            "v_mnb_part_line_stocking_location.stocking_location_id " & _
            "from v_mnb_part_line_stocking_location " & _
            "join v_mnb_model_part on v_mnb_part_line_stocking_location.part_id = v_mnb_model_part.part_id and " & _
            "v_mnb_part_line_stocking_location.line_id = v_mnb_model_part.line_id " & _
            "where v_mnb_model_part.model_number = '" & TDBGNewModels.Columns(0).Value & "' and " & _
            "v_mnb_model_part.line_id = '" & TDBGNewModels.Columns(2).Value & "'"
        .Open
    End With
    
    ' If the recordset has not be instaniated, instantiate it and initialize properties.
    If mrsDatabase Is Nothing Then
        Set mrsDatabase = New ADODB.Recordset
        With mrsDatabase
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
'            .CursorType = adOpenStatic
            .LockType = adLockPessimistic
'            .LockType = adLockBatchOptimistic
        End With
    
    ' If the recordset is open, close it.
    ElseIf mrsDatabase.State = adStateOpen Then
        mrsDatabase.Close
    End If
    
    ' Read any records for the selected model.  This should return an empty recordset
    ' since only models which have not yet been set up will be displayed on the list.
    With mrsDatabase
        .Source = "select * from v_mnb_model_part_stocking_location " & _
            "where model_number = '" & TDBGNewModels.Columns(0).Value & "' and " & _
            "line_id = '" & TDBGNewModels.Columns(2).Value & "'"
        .Open
        
        ' Read each record from the default part recordset.  Add a record to the
        ' model_part_stocking_location table and set fields.
        lngRecords = 0
        Do While Not rsDefaults.EOF
            .AddNew
            !model_number = TDBGNewModels.Columns(0).Value
            !part_id = rsDefaults!part_id
            !line_id = rsDefaults!line_id
            !stocking_location_id = rsDefaults!stocking_location_id
            !part_sequence_number = rsDefaults!part_sequence_number
            rsDefaults.MoveNext
            lngRecords = lngRecords + 1
        Loop
        
        ' Save the new records and check for errors.
        .UpdateBatch
        If gconDatabase.Errors.Count > 0 Then
            Err.Raise gconDatabase.Errors(0).NativeError, "UseDefaults", _
                gconDatabase.Errors(0).Description
        End If
        
        ' Close the recordset.
        .Close
    End With
    
        
    ' Display a record to confirm that parts were set up from the defaults.
    MsgBox Trim(CStr(lngRecords)) & " parts were set up from defaults for model " & _
        Me.TDBGNewModels.Columns(0).Value
    
    ' Close the default parts recordset and de-reference it.
    rsDefaults.Close
    Set rsDefaults = Nothing
    
    ' Requery the models to remove this model from the new model list and add to the set up list.
    'RequeryModels
    
    'Delete the model just copied from the grid.
    If lngRecords > 0 Then
        Me.TDBGNewModels.Delete
    End If

    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "UseDefaults", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub RequeryModels()
    ' Purpose:  Requery models to update lists after new models have been processed.
    '  Deactivated this subroutine 10/28/2004.  Because this query runs 1 minute, a
    '  decision was made not to requery, but just delete the model updated from the grid.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    With mrsModels
        ' Reset the filter to include all models.
        
        .Filter = ""
        
        ' Requery the models.
        .Requery
            
        ' Clear the model combo box
        cboModel.Clear
        cboLine.Clear
        
        Set mxarrModelsToCopy = New XArrayDB
        
        Dim intModels As Integer
        Dim strSaveModel As String
        intModels = 0
        
        ' Load any models which have been set up into the combo box of models.  This
        ' combo box will be used for a selection list of models from which location info
        ' may be copied.
        Do While Not .EOF
            If !locations_set_up > 0 Then
                If intModels = 0 Then
                    mxarrModelsToCopy.ReDim 1, 1, 1, 2
                Else
                    mxarrModelsToCopy.AppendRows
                End If
                intModels = intModels + 1
                mxarrModelsToCopy(intModels, 1) = Trim(!model_number)
                mxarrModelsToCopy(intModels, 2) = .Fields("line_id").Value
            End If
            .MoveNext
        Loop
        
        If intModels = 0 Then
            optCopyModel.Enabled = False
            cboModel.Enabled = False
            cboLine.Enabled = False
        Else
            optCopyModel.Enabled = True
            cboModel.Enabled = True
            cboLine.Enabled = True
            strSaveModel = vbNullString
            For intModels = 1 To mxarrModelsToCopy.UpperBound(1)
                If mxarrModelsToCopy(intModels, 1) <> strSaveModel Then
                    cboModel.AddItem mxarrModelsToCopy(intModels, 1)
                    cboModel.ItemData(cboModel.NewIndex) = intModels
                    strSaveModel = mxarrModelsToCopy(intModels, 1)
                End If
            Next intModels
        End If
        
        
        ' Set the filter for the recordset to exclude models which have been set up.
        .Filter = "locations_set_up = 0"
    End With
    
    ' Rebind the grid to the recordset to rebuild the values.
    TDBGNewModels.ReBind
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RequeryModels", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub CopyModel()
    ' Purpose:  Copy all matching parts from one model to another.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set up an index to be used for going through array.
    Dim lngIndex As Long
    Dim lngRecords As Long
    Dim lngRow As Long
    
    If Len(cboModel.Text) = 0 Then
        MsgBox "Model is required."
        cboModel.SetFocus
        Exit Sub
    End If
    cboFindFirst cboModel
    
    If cboModel.ListIndex = -1 Then
        MsgBox "Model must be selected."
        cboModel.SetFocus
        Exit Sub
    End If
    
    If Len(cboLine.Text) = 0 Then
        MsgBox "Line is required."
        cboLine.SetFocus
        Exit Sub
    End If
    
    If cboLine.ListIndex = -1 Then
        MsgBox "Line must be selected."
        Exit Sub
    End If
    
    ' Declare and instantiate an array to hold the info needed to load a new model from an old one.
    Dim xarrSourceModelInfo As XArrayDB
    Dim xarrDestModelInfo As XArrayDB
    Set xarrSourceModelInfo = New XArrayDB
    Set xarrDestModelInfo = New XArrayDB
    
    ' Set up recordset to hold source model info
    Dim rsModel As ADODB.Recordset
    Set rsModel = New ADODB.Recordset
    With rsModel
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select v_mnb_model_part.model_number, v_mnb_model_part.line_id, " & _
            "v_mnb_model_part.part_id, v_mnb_model_part.part_sequence_number, " & _
            "v_mnb_model_part.quantity, stocking_location_id, step_number " & _
            "from v_mnb_model_part_stocking_location " & _
            "join v_mnb_model_part " & _
            "on v_mnb_model_part_stocking_location.model_number = v_mnb_model_part.model_number and " & _
            "v_mnb_model_part_stocking_location.line_id = v_mnb_model_part.line_id and " & _
            "v_mnb_model_part_stocking_location.part_id = v_mnb_model_part.part_id and " & _
            "v_mnb_model_part_stocking_location.part_sequence_number = v_mnb_model_part.part_sequence_number " & _
            "where v_mnb_model_part.model_number = '" & cboModel.Text & "' and " & _
            "v_mnb_model_part.line_id = '" & cboLine.Text & "'"
        .Open
        xarrSourceModelInfo.LoadRows .GetRows, True
        .Close
        
        If cboLine.Text <> TDBGNewModels.Columns(2).Value Then
            frmLineLocationCheck.lblMessage = "Lines mismatch "
            frmLineLocationCheck.mstrModel = cboModel.Text
            frmLineLocationCheck.mstrFromLine = cboLine.Text
            frmLineLocationCheck.mstrToLine = TDBGNewModels.Columns(2).Value
            frmLineLocationCheck.CheckLocations
            If Not frmLineLocationCheck.mblnContinue Then
                MsgBox "Processing was cancelled."
                Exit Sub
            End If
        End If
        
        ' Retrieve data for destination model
        .Source = "select * from v_mnb_model_part " & _
            "where model_number = '" & TDBGNewModels.Columns(0).Value & "' and " & _
            "line_id = '" & TDBGNewModels.Columns(2).Value & "' " & _
            "order by part_id"
        .Open
        xarrDestModelInfo.LoadRows .GetRows, True
        .Close
    End With
    
    Set rsModel = Nothing
    
    ' Set up recordset to be used to set up locations for destination model
    ' If the recordset has not be instaniated, instantiate it and initialize properties.
    If mrsDatabase Is Nothing Then
        Set mrsDatabase = New ADODB.Recordset
        With mrsDatabase
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
'            .CursorType = adOpenStatic
            .LockType = adLockPessimistic
'            .LockType = adLockBatchOptimistic
        End With
    
    ' If the recordset is open, close it.
    ElseIf mrsDatabase.State = adStateOpen Then
        mrsDatabase.Close
    End If
    
    ' Read any records for the selected model.  This should return an empty recordset
    ' since only models which have not yet been set up will be displayed on the list.
    With mrsDatabase
        .Source = "select * from v_mnb_model_part_stocking_location " & _
            "where model_number = '" & TDBGNewModels.Columns(0).Value & "' and " & _
            "line_id = '" & TDBGNewModels.Columns(2).Value & "'"
        .Open
        
        
    'Read through the Source Model to match parts with the Model that needs locations loaded.
    '  Adds record(s) to v_mnb_model_part_stocking_location with location
        Dim blnProcessed() As Boolean
        ReDim blnProcessed(xarrDestModelInfo.UpperBound(1))
        lngRecords = 0
        For lngIndex = 0 To xarrSourceModelInfo.UpperBound(1)
            lngRow = xarrDestModelInfo.Find(0, 2, xarrSourceModelInfo(lngIndex, 2))
            If lngRow > -1 Then
                Do While lngRow <= xarrDestModelInfo.UpperBound(1) And blnProcessed(lngRow) And _
                        xarrDestModelInfo(lngRow, 2) = xarrSourceModelInfo(lngIndex, 2)
                            lngRow = lngRow + 1
                   If lngRow > xarrDestModelInfo.UpperBound(1) Then
                        Exit Do
                   End If
                Loop
                
                If lngRow <= xarrDestModelInfo.UpperBound(1) Then
                                         
                    If lngRow <= xarrDestModelInfo.UpperBound(1) And xarrDestModelInfo(lngRow, 2) = _
                            xarrSourceModelInfo(lngIndex, 2) Then
                        .AddNew
                        !model_number = xarrDestModelInfo(lngRow, 0)
                        !line_id = xarrDestModelInfo(lngRow, 1)
                        !part_id = xarrDestModelInfo(lngRow, 2)
                        !part_sequence_number = xarrDestModelInfo(lngRow, 3)
                        !stocking_location_id = xarrSourceModelInfo(lngIndex, 5)
                        !step_number = xarrSourceModelInfo(lngIndex, 6)
                        blnProcessed(lngRow) = True
                        lngRecords = lngRecords + 1
                    End If
                End If
            End If
        Next lngIndex
        
        ' Save the new records and check for errors.
        .UpdateBatch
        If gconDatabase.Errors.Count > 0 Then
            Err.Raise gconDatabase.Errors(0).NativeError, "UseDefaults", _
                gconDatabase.Errors(0).Description
        End If
        
        ' Close the recordset.
        .Close
    End With
    
    ' Display a record to confirm that parts were set up from the defaults.
    MsgBox Trim(CStr(lngRecords)) & " parts were copied from model " & cboModel.Text & " to model " & _
        Me.TDBGNewModels.Columns(0).Value
    
    'RequeryModels
    If lngRecords > 0 Then
        Me.TDBGNewModels.Delete
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Me.MousePointer = vbDefault
    Call ShowError(Me.Name, "CopyModel", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub optCopyModel_Click()
    cboLine.Enabled = True
End Sub

Private Sub optMiniBillMaintenance_Click()
    cboLine.Enabled = False
    cboModel.Enabled = False
End Sub

Private Sub optUseDefaults_Click()
    cboLine.Enabled = False
    cboModel.Enabled = False
End Sub
