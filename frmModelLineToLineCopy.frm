VERSION 5.00
Begin VB.Form frmModelLineToLineCopy 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Minibill - Copy Model Location Information From Between Lines"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
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
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   446
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2460
      Width           =   1335
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      Default         =   -1  'True
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2460
      Width           =   1335
   End
   Begin VB.ComboBox cboToLine 
      Height          =   360
      Left            =   2055
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.ComboBox cboFromLine 
      Height          =   360
      Left            =   2055
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox cboModel 
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Line:"
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
      Left            =   1155
      TabIndex        =   5
      Top             =   1620
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Line:"
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
      Left            =   915
      TabIndex        =   3
      Top             =   1020
      Width           =   1095
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
      Left            =   1200
      TabIndex        =   1
      Top             =   420
      Width           =   795
   End
End
Attribute VB_Name = "frmModelLineToLineCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsDatabase As ADODB.Recordset
Private mxarrModelInfo As XArrayDB
Private mstrSaveModel As String

Private Sub cboFromLine_Change()
    If Len(cboFromLine.Text) = 2 Then
        cboFindFirst cboFromLine
    End If
End Sub

Private Sub cboFromLine_GotFocus()
    cboFromLine.SelStart = 0
    cboFromLine.SelLength = Len(cboFromLine)
End Sub

Private Sub cboFromLine_Validate(Cancel As Boolean)
    If Len(cboFromLine.Text) = 1 Then
        cboFromLine.Text = "0" & cboFromLine.Text
    End If
    cboFindFirst cboFromLine
    If cboFromLine.ListIndex = -1 Then
        MsgBox "Invalid line selected."
        Cancel = True
    End If
End Sub


Private Sub cboModel_Change()
    cboFindFirst cboModel
End Sub

Private Sub cboModel_Click()
    Dim lngIndex As Long
    
    
    If cboModel.ListIndex > -1 Then
        If mstrSaveModel = cboModel.Text Then
            Exit Sub
        End If
        cboFromLine.Clear
        cboToLine.Clear
        lngIndex = cboModel.ItemData(cboModel.ListIndex)
        Do While lngIndex <= mxarrModelInfo.UpperBound(1)
            If mxarrModelInfo(lngIndex, 1) <> cboModel.Text Then
                Exit Do
            End If
            If mxarrModelInfo(lngIndex, 3) Then
                cboFromLine.AddItem mxarrModelInfo(lngIndex, 2)
            Else
                cboToLine.AddItem mxarrModelInfo(lngIndex, 2)
            End If
            lngIndex = lngIndex + 1
        Loop
    End If
    mstrSaveModel = cboModel.Text
    
    If cboFromLine.ListCount > 0 And cboToLine.ListCount > 0 Then
        cboFromLine.ListIndex = 0
        cboToLine.ListIndex = 0
    End If
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
    If cboModel.ListIndex = -1 Then
        MsgBox "Invalid model selected."
        Cancel = True
    End If
    If cboFromLine.ListCount = 0 Then
        MsgBox "This model has not yet been set up with any locations."
        Cancel = True
    ElseIf cboToLine.ListCount = 0 Then
        MsgBox "This model has no lines which need to be set up."
        Cancel = True
    End If
End Sub

Private Sub cboToLine_Change()
    If Len(cboToLine.Text) = 2 Then
        cboFindFirst cboToLine
    End If
End Sub

Private Sub cboToLine_GotFocus()
    cboToLine.SelStart = 0
    cboToLine.SelLength = Len(cboToLine)
End Sub

Private Sub cboToLine_Validate(Cancel As Boolean)
    If Len(cboToLine.Text) = 1 Then
        cboToLine.Text = "0" & cboToLine.Text
    End If
    cboFindFirst cboToLine
    If cboToLine.ListIndex = -1 Then
        MsgBox "Invalid line selected."
        Cancel = True
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdProcess_Click()
    cboFindFirst cboModel
    If cboModel.ListIndex = -1 Then
        MsgBox "Model must be selected."
        Exit Sub
    End If
    
    cboFindFirst cboFromLine
    If cboFromLine.ListIndex = -1 Then
        MsgBox "From Line must be selected."
        Exit Sub
    End If
    
    cboFindFirst cboToLine
    If cboToLine.ListIndex = -1 Then
        MsgBox "To Line must be selected."
        Exit Sub
    End If
    
    With frmLineLocationCheck
        .lblMessage = "Lines mismatch "
        .mstrModel = cboModel.Text
        .mstrFromLine = cboFromLine.Text
        .mstrToLine = cboToLine.Text
        .CheckLocations
        If .mblnContinue Then
            CopySplits
            CopyModel
        End If
    End With
End Sub

Private Sub Form_Load()
    ' Purpose:  Load data
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim lngIndex As Long
    Dim strPrevModel As String
    Dim strPrevLine As String
    Dim blnLineUsed As Boolean
    Dim blnWritten As Boolean
    
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
    
    frmProcessing.Label1 = "Gathering Models... Please Wait"
    frmProcessing.Label2 = "Get Configuration Models List at " & Time
    frmProcessing.Show
    
    DoEvents
      
    Set mrsDatabase = New ADODB.Recordset
    With mrsDatabase
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select v_mnb_model_line.model_number, v_mnb_model_part.line_id, " & _
            "sum(case when stocking_location_id is null then 0 else 1 end) as Location_Count " & _
            "From v_mnb_model_part " & _
            "left outer join v_mnb_model_line on " & _
            "v_mnb_model_part.model_number = v_mnb_model_line.model_number and " & _
            "v_mnb_model_part.line_id = v_mnb_model_line.line_id " & _
            "left outer join v_mnb_model_part_stocking_location on " & _
            "v_mnb_model_part.model_number = v_mnb_model_part_stocking_location.model_number and " & _
            "v_mnb_model_part.line_id = v_mnb_model_part_stocking_location.line_id and " & _
            "v_mnb_model_part.part_id = v_mnb_model_part_stocking_location.part_id and " & _
            "v_mnb_model_part.part_sequence_number = v_mnb_model_part_stocking_location.part_sequence_number " & _
            "where v_mnb_model_line.model_number is not null " & _
            "group by v_mnb_model_line.model_number, v_mnb_model_part.line_id " & _
            "order by v_mnb_model_line.model_number, v_mnb_model_part.line_id"
        .Open
                
        If .RecordCount > 0 Then
            Set mxarrModelInfo = New XArrayDB
            lngIndex = 0
            Do While Not .EOF
                If !model_number = strPrevModel Then
                    If Not blnWritten Then
                        If lngIndex = 0 Then
                            mxarrModelInfo.ReDim 1, 1, 1, 3
                        Else
                            mxarrModelInfo.AppendRows
                        End If
                        lngIndex = lngIndex + 1
                        mxarrModelInfo(lngIndex, 1) = Trim(strPrevModel)
                        mxarrModelInfo(lngIndex, 2) = strPrevLine
                        mxarrModelInfo(lngIndex, 3) = blnLineUsed
                        blnWritten = True
                        cboModel.AddItem Trim(!model_number)
                        cboModel.ItemData(cboModel.NewIndex) = lngIndex
                    End If
                    mxarrModelInfo.AppendRows
                    lngIndex = lngIndex + 1
                    mxarrModelInfo(lngIndex, 1) = Trim(.Fields("model_number").Value)
                    mxarrModelInfo(lngIndex, 2) = .Fields("line_id").Value
                    If !location_count > 0 Then
                        blnLineUsed = True
                    Else
                        blnLineUsed = False
                    End If
                    mxarrModelInfo(lngIndex, 3) = blnLineUsed
                Else
                    strPrevModel = !model_number
                    strPrevLine = !line_id
                    If !location_count > 0 Then
                        blnLineUsed = True
                    Else
                        blnLineUsed = False
                    End If
                    blnWritten = False
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    Dim intCounter As Integer
    Dim lngFromIndex As Long
    Dim lngToIndex As Long
    Dim blnToLineFound As Boolean
    Dim blnFromLinefound As Boolean
    Dim arrblnDeleteModel() As Boolean
    
    ReDim arrblnDeleteModel(cboModel.ListCount - 1)
    If cboModel.ListCount > 0 Then
        For intCounter = cboModel.ListCount - 1 To 0 Step -1
            blnFromLinefound = False
            blnToLineFound = False
            lngFromIndex = cboModel.ItemData(intCounter)
            If intCounter < cboModel.ListCount - 1 Then
                lngToIndex = cboModel.ItemData(intCounter + 1) - 1
            Else
                lngToIndex = mxarrModelInfo.UpperBound(1)
            End If
            For lngIndex = lngFromIndex To lngToIndex
                If mxarrModelInfo(lngIndex, 3) Then
                    blnFromLinefound = True
                Else
                    blnToLineFound = True
                End If
            Next
            If blnFromLinefound = False Or blnToLineFound = False Then
                arrblnDeleteModel(intCounter) = True
            Else
                arrblnDeleteModel(intCounter) = False
            End If
        Next intCounter
    End If
    
    For intCounter = cboModel.ListCount - 1 To 0 Step -1
        If arrblnDeleteModel(intCounter) Then
            cboModel.RemoveItem intCounter
        End If
    Next intCounter
    
    If cboModel.ListCount = 0 Then
        MsgBox "No models exist on multiple lines."
        Unload Me
        GoTo PROC_EXIT
    End If
    
    frmProcessing.Hide
    
    DoEvents
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, Err.Description)
    Unload frmProcessing
    Unload Me
End Sub
Private Sub CopySplits()
    ' Purpose:  Copy all mnb_model_part records from one model/line to another.
    ' This code was added to catch Model/Line records with split parts
    ' The destination mnb_model_Part records will be deleted and new records
    ' will be added from the source mnb_model_part file
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set up an index to be used for going through array.
    Dim lngIndex As Long
    Dim lngRecords As Long
    Dim lngRow As Long
    
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
        .Source = "delete from v_mnb_model_part " & _
            "where v_mnb_model_part.model_number = '" & cboModel.Text & "' and " & _
            "v_mnb_model_part.line_id = '" & cboToLine.Text & "'"
        .Open
'        xarrSourceModelInfo.LoadRows .GetRows, True
'        .Close
        
        ' Retrieve data for destination model
        .Source = "select * from v_mnb_model_part " & _
            "where model_number = '" & cboModel.Text & "' and " & _
            "line_id = '" & cboFromLine.Text & "' " & _
            "order by part_id"
        .Open
'        xarrDestModelInfo.LoadRows .GetRows, True
'        .Close
    End With
    
'    Set rsModel = Nothing
    
    ' Set up recordset to be used to add mnb_model_part records for destination model
    '
    If mrsDatabase Is Nothing Then
        Set mrsDatabase = New ADODB.Recordset
        With mrsDatabase
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
        End With
    
    ' If the recordset is open, close it.
    Else
        If mrsDatabase.State = adStateOpen Then
            mrsDatabase.Close
        End If
        mrsDatabase.CursorType = adOpenStatic
        mrsDatabase.LockType = adLockBatchOptimistic
    End If
    
    ' Read any records for the selected model.  This should return an empty recordset
    ' since they have been deleted in the previous step.
    ' Then add selected mnb_model_part records back.
    
    With mrsDatabase
        .Source = "select * from v_mnb_model_part " & _
            "where model_number = '" & cboModel.Text & "' and " & _
            "line_id = '" & cboToLine.Text & "'"
        .Open
        If rsModel.RecordCount > 0 Then
            rsModel.MoveFirst
            Do While Not rsModel.EOF
               .AddNew
               !model_number = rsModel!model_number
               !line_id = cboToLine.Text
               !part_id = rsModel!part_id
               !part_sequence_number = rsModel!part_sequence_number
               !original_sequence_number = rsModel!original_sequence_number
               !quantity = rsModel!quantity
               !start_ecn_number = rsModel!start_ecn_number
               !start_ecn_date = rsModel!start_ecn_date
               !start_ecn_flag = rsModel!start_ecn_flag
               !quit_ecn_number = rsModel!quit_ecn_number
               !quit_ecn_date = rsModel!quit_ecn_date
               !quit_ecn_flag = rsModel!quit_ecn_flag
               !level_number = rsModel!level_number
               !Comments = rsModel!Comments
               !part_create_date = rsModel!part_create_date
               !part_reviewed_flag = rsModel!part_reviewed_flag
               !parent_part_number = rsModel!parent_part_number
               !inactive_part_flag = rsModel!inactive_part_flag
               rsModel.MoveNext
            Loop
        End If
        ' Save the new records and check for errors.
        .UpdateBatch
        If gconDatabase.Errors.Count > 0 Then
            Err.Raise gconDatabase.Errors(0).NativeError, "UseDefaults", _
                gconDatabase.Errors(0).Description
        End If
        
        ' Close the recordset.
        .Close
        rsModel.Close
    End With
    
   
PROC_EXIT:

    Set rsModel = Nothing
    
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "CopySplits", Err.Number, Err.Description)
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
            "v_mnb_model_part.line_id = '" & cboFromLine.Text & "'"
        .Open
        xarrSourceModelInfo.LoadRows .GetRows, True
        .Close
        
        ' Retrieve data for destination model
        .Source = "select * from v_mnb_model_part " & _
            "where model_number = '" & cboModel.Text & "' and " & _
            "line_id = '" & cboToLine.Text & "' " & _
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
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
        End With
    
    ' If the recordset is open, close it.
    Else
        If mrsDatabase.State = adStateOpen Then
            mrsDatabase.Close
        End If
        mrsDatabase.CursorType = adOpenStatic
        mrsDatabase.LockType = adLockBatchOptimistic
    End If
    
    ' Read any records for the selected model.  This should return an empty recordset
    ' since only models which have not yet been set up will be displayed on the list.
    With mrsDatabase
        .Source = "select * from v_mnb_model_part_stocking_location " & _
            "where model_number = '" & cboModel.Text & "' and " & _
            "line_id = '" & cboToLine.Text & "'"
        .Open
        
        Dim blnProcessed() As Boolean
        Dim lngPartSequence As Long
        ReDim blnProcessed(xarrDestModelInfo.UpperBound(1))
        lngRecords = 0
        For lngIndex = 0 To xarrSourceModelInfo.UpperBound(1)
            lngRow = xarrDestModelInfo.Find(0, 2, xarrSourceModelInfo(lngIndex, 2))
            If lngRow > -1 Then
                Do While lngRow <= xarrDestModelInfo.UpperBound(1) And blnProcessed(lngRow) And _
                        xarrDestModelInfo(lngRow, 2) = xarrSourceModelInfo(lngIndex, 2)
                    lngRow = lngRow + 1
                Loop
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
    MsgBox Trim(CStr(lngRecords)) & " parts were copied for model " & cboModel.Text & " from line " & _
        cboFromLine.Text & " to line " & cboToLine.Text
        
    cboModel.Text = ""
    cboFromLine.Text = ""
    cboToLine.Text = ""
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "CopyModel", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not mrsDatabase Is Nothing Then
        If mrsDatabase.State = adStateOpen Then
            mrsDatabase.Close
        End If
        Set mrsDatabase = Nothing
    End If
    
    Unload frmProcessing
End Sub
