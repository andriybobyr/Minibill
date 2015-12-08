VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLineLocationCheck 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Minibill - Check Locations For Lines"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
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
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2940
      TabIndex        =   5
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Action To Be Taken"
      Height          =   1575
      Left            =   180
      TabIndex        =   2
      Top             =   3720
      Width           =   6015
      Begin VB.OptionButton optAbort 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel processing"
         Height          =   435
         Left            =   240
         TabIndex        =   4
         Top             =   900
         Value           =   -1  'True
         Width           =   5595
      End
      Begin VB.OptionButton optAddToLine 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add these locations to this line and continue processing the selected model."
         Height          =   555
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   5595
      End
   End
   Begin MSDataGridLib.DataGrid dgrdLocations 
      Height          =   1875
      Left            =   240
      TabIndex        =   1
      Top             =   1620
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3307
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "stocking_location_id"
         Caption         =   "Location"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "stocking_location_description"
         Caption         =   "Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   5
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   264.983
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Height          =   915
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   6015
   End
End
Attribute VB_Name = "frmLineLocationCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mblnContinue As Boolean
Public mstrModel As String
Public mstrFromLine As String
Public mstrToLine As String

Private mrsLocations As ADODB.Recordset

Public Sub CheckLocations()
    ' Purpose:  Check the locations necessary for copying the model from one line to a second line.
    '           If the locations are set up properly for the line, the mblnContinue flag will be
    '           set to true, and the form will be unloaded.  If the locations are not set up, the
    '           form will be displayed allowing the user to choose to copy the locations or to abort
    '           the operation.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Instantiate and open recordset to retrieve line/location information.
    Set mrsLocations = New ADODB.Recordset
    With mrsLocations
        Set .ActiveConnection = gconDatabase
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Source = "select v_prod_stocking_location.stocking_location_id, " & _
            "min(stocking_location_description) as stocking_location_description, " & _
            "sum(case when v_prod_line_stocking_location.line_id = '" & _
            mstrToLine & "' then 1 else 0 end) as to_line_found " & _
            "From v_mnb_model_part_stocking_location " & _
            "join v_prod_stocking_location on " & _
            "v_mnb_model_part_stocking_location.stocking_location_id = v_prod_stocking_location.stocking_location_id " & _
            "join v_prod_line_stocking_location on " & _
            "v_prod_stocking_location.stocking_location_id = v_prod_line_stocking_location.stocking_location_id " & _
            "where model_number = '" & mstrModel & "' and " & _
            "v_mnb_model_part_stocking_location.line_id = '" & mstrFromLine & "' " & _
            "group by v_prod_stocking_location.stocking_location_id " & _
            "order by v_prod_stocking_location.stocking_location_id"
        .Open
        .Filter = "to_line_found = 0"
        If .RecordCount = 0 Then
            .Close
            mblnContinue = True
            Unload Me
            GoTo PROC_EXIT
        End If
    End With
    Set dgrdLocations.DataSource = mrsLocations
    dgrdLocations.ReBind
    Me.Show vbModal

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "CheckLocations", Err.Number, Err.Description)
    Unload Me
End Sub

Private Sub cmdOK_Click()
        
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    If Me.optAbort.Value Then
        Me.mblnContinue = False
        Unload Me
        GoTo PROC_EXIT
    End If
    
    Dim rsLineLocation As ADODB.Recordset
    Set rsLineLocation = New ADODB.Recordset
    
    With rsLineLocation
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Source = "select * from v_prod_line_stocking_location where line_id = '" & _
            mstrToLine & "'"
        .Open
        mrsLocations.MoveFirst
        Do While Not mrsLocations.EOF
            .AddNew
            !line_id = mstrToLine
            !stocking_location_id = mrsLocations!stocking_location_id
            mrsLocations.MoveNext
        Loop
        .UpdateBatch
        .Close
    End With
    Set rsLineLocation = Nothing
    
    Me.mblnContinue = True
    Unload Me
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdOK_Click", Err.Number, Err.Description)
    Unload Me
End Sub

