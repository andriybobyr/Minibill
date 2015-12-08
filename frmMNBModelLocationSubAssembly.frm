VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMNBModelLocationSubAssembly 
   Caption         =   "MiniBill - Model Sub-Assembly By Location Maintenance"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   StartUpPosition =   1  'CenterOwner
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   4425
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7710
      _cx             =   13600
      _cy             =   7805
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      ForeColorDisabled=   -2147483630
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   6
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   2500
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
      Begin VB.CommandButton cmdDelete 
         Height          =   315
         Left            =   465
         Picture         =   "frmMNBModelLocationSubAssembly.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Delete Current Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.TextBox txtModel 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3120
         MaxLength       =   20
         TabIndex        =   14
         Tag             =   "Model:"
         Top             =   1125
         Width           =   2400
      End
      Begin VB.ComboBox cboSubAssy 
         Height          =   360
         Left            =   3135
         TabIndex        =   17
         Tag             =   "Sub-Assembly:"
         Top             =   3165
         Width           =   2340
      End
      Begin VB.ComboBox cboLocation 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3135
         TabIndex        =   16
         Tag             =   "Location:"
         Top             =   2490
         Width           =   2940
      End
      Begin VB.ComboBox cboLine 
         Enabled         =   0   'False
         Height          =   360
         Left            =   3135
         TabIndex        =   15
         Tag             =   "Line:"
         Top             =   1800
         Width           =   930
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   120
         Picture         =   "frmMNBModelLocationSubAssembly.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Add New Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdSave 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   825
         Picture         =   "frmMNBModelLocationSubAssembly.frx":0724
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Save Changes"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdFind 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   1545
         Picture         =   "frmMNBModelLocationSubAssembly.frx":0CCE
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Find An Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdFirst 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   1905
         Picture         =   "frmMNBModelLocationSubAssembly.frx":1264
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Display First Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdPrevious 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   2265
         Picture         =   "frmMNBModelLocationSubAssembly.frx":183E
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Display Previous Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdNext 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   2625
         Picture         =   "frmMNBModelLocationSubAssembly.frx":1E18
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Display Next Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdLast 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   2985
         Picture         =   "frmMNBModelLocationSubAssembly.frx":23F2
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Display Last Entry"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdList 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   3420
         Picture         =   "frmMNBModelLocationSubAssembly.frx":29CC
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Selection List"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdHelp 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   3765
         Picture         =   "frmMNBModelLocationSubAssembly.frx":2B56
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Help"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdClose 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   4125
         Picture         =   "frmMNBModelLocationSubAssembly.frx":2C58
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Close This Form"
         Top             =   120
         Width           =   300
      End
      Begin VB.CommandButton cmdCancel 
         CausesValidation=   0   'False
         Height          =   330
         Left            =   1185
         Picture         =   "frmMNBModelLocationSubAssembly.frx":33BA
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Cancel Changes"
         Top             =   120
         Width           =   300
      End
      Begin MSComctlLib.StatusBar staDBMaint 
         Height          =   390
         Left            =   15
         TabIndex        =   1
         Top             =   4020
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   3
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   6
               Alignment       =   1
               TextSave        =   "10/5/2007"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   5
               Alignment       =   1
               TextSave        =   "2:33 PM"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Bevel           =   2
               Object.Width           =   8414
               Key             =   "Msg"
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   0
         X1              =   120
         X2              =   7440
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   1
         X1              =   120
         X2              =   7440
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
         Height          =   345
         Left            =   5100
         TabIndex        =   2
         Top             =   120
         Width           =   2475
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileCancel 
         Caption         =   "C&ancel Current Update/Add"
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
      Begin VB.Menu mnuViewFind 
         Caption         =   "F&ind"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
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
         Caption         =   "&Last"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "&Selection List..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
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
Attribute VB_Name = "frmMNBModelLocationSubAssembly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents mrsDatabase As ADODB.Recordset
Attribute mrsDatabase.VB_VarHelpID = -1

Private mblnRecChanged As Boolean
Private marrstrLocation() As String

Private Sub cboLine_Change()
    cboFindFirst cboLine
    mblnRecChanged = True
End Sub

Private Sub cboLine_Click()
    Call RetrieveLocationData(txtModel.Text, cboLine.Text)
End Sub

Private Sub cboLine_KeyPress(KeyAscii As Integer)
    cboKeyPress cboLine, KeyAscii
    mblnRecChanged = True
End Sub

Private Sub cboLine_Validate(Cancel As Boolean)
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    ' If there is no data in the field, display a message,
    ' set cancel to true and exit.
    If Len(Trim(cboLine.Text)) = 0 Then
        Cancel = True
        MsgBox "Line Is Required!", _
            vbExclamation + vbOKOnly, _
            "Line Validation"
        GoTo PROC_EXIT
     End If
        
    ' If the type id has changed, look up the new
    ' type id in the listbox.
    If mrsDatabase!line_id <> cboLine.Text Then
        cboFindFirst cboLine
            
        ' If the type id was not found in the list,
        ' display a message, set cancel to true and
        ' exit.
        If cboLine.ListIndex = -1 Then
            MsgBox "Line " & cboLine.Text & " is not valid"
            Cancel = True
            GoTo PROC_EXIT
        End If
            
        mrsDatabase!line_id = cboLine.Text
        Cancel = False
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboLine_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cboLocation_Change()
    cboFindFirst cboLocation
    mblnRecChanged = True
End Sub

Private Sub cboLocation_Click()
    Call RetrievePartData(txtModel.Text, cboLine.Text, _
        marrstrLocation(cboLocation.ItemData(cboLocation.ListIndex)))
End Sub

Private Sub cboLocation_KeyPress(KeyAscii As Integer)
    cboKeyPress cboLocation, KeyAscii
    mblnRecChanged = True
End Sub

Private Sub cboLocation_Validate(Cancel As Boolean)
    
    On Error GoTo PROC_ERR
    
    Dim strLocationId As String
    Dim rsDuplicate As ADODB.Recordset
    
    ' If there is no data in the field, display a message,
    ' set cancel to true and exit.
    If Len(Trim(cboLocation.Text)) = 0 Then
        Cancel = True
        MsgBox "Location Is Required!", _
            vbExclamation + vbOKOnly, _
            "Location Validation"
        GoTo PROC_EXIT
     End If
        
    cboFindFirst cboLocation
    strLocationId = marrstrLocation(cboLocation.ItemData(cboLocation.ListIndex))

    ' If the type id has changed, look up the new
    ' type id in the listbox.
    If Trim(mrsDatabase!stocking_location_id) <> strLocationId Then
            
        Cancel = False
        ' If the type id was not found in the list,
        ' display a message, set cancel to true and
        ' exit.
        If cboLocation.ListIndex = -1 Then
            MsgBox "Location " & cboLocation.Text & " is not valid"
            Cancel = True
            GoTo PROC_EXIT
        End If
            
        Set rsDuplicate = mrsDatabase.Clone
    
        With rsDuplicate
            .MoveFirst
            .Find ("model_number = '" & txtModel.Text & "'")
    
            Do While Not .EOF
                If !model_number <> mrsDatabase!model_number Then
                    Exit Do
                End If
            
                If !line_id = mrsDatabase!line_id And _
                   Trim(!stocking_location_id) = strLocationId Then
                    Cancel = True
                    MsgBox "Model/Line/Location already exists "
                    Exit Do
                End If
                .Find "model_number = '" & txtModel.Text & "'", 1, adSearchForward, .Bookmark
            Loop
            .Close
        End With
    
        Set rsDuplicate = Nothing
        
        If Cancel Then
            GoTo PROC_EXIT
        End If
            
        strLocationId = marrstrLocation(cboLocation.ItemData(cboLocation.ListIndex))
        mrsDatabase!stocking_location_id = strLocationId
    End If
    
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboLocation_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cboSubAssy_Change()
    cboFindFirst cboSubAssy
    mblnRecChanged = True
End Sub

Private Sub cboSubAssy_KeyPress(KeyAscii As Integer)
    cboKeyPress cboSubAssy, KeyAscii
    mblnRecChanged = True
End Sub

Private Sub cboSubAssy_Validate(Cancel As Boolean)
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    ' If there is no data in the field, display a message,
    ' set cancel to true and exit.
    If Len(Trim(cboSubAssy.Text)) = 0 Then
        Cancel = True
        MsgBox "Sub Assembly Is Required!", _
            vbExclamation + vbOKOnly, _
            "Sub Assembly Validation"
        GoTo PROC_EXIT
     End If
        
    ' If the type id has changed, look up the new
    ' type id in the listbox.
    If mrsDatabase!sub_assembly_id <> cboSubAssy.Text Then
        cboFindFirst cboSubAssy
            
        ' If the type id was not found in the list,
        ' display a message, set cancel to true and
        ' exit.
        If cboSubAssy.ListIndex = -1 Then
            MsgBox "Sub Assembly " & cboSubAssy.Text & " is not valid"
            Cancel = True
            GoTo PROC_EXIT
        End If
            
        mrsDatabase!sub_assembly_id = cboSubAssy.Text
        Cancel = False
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboSubAssy_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdAdd_Click()
    ' Purpose:  Tie the click of this button to the FileNew
    '           menu item.
    
    Call mnuFileNew_Click
End Sub

Private Sub cmdClose_Click()
    ' Purpose:  Close the form by the user's request.
    
    Call mnuFileClose_Click
End Sub

Private Sub cmdDelete_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewFind menu item.
    
    Call mnuFileDelete_Click
End Sub

Private Sub cmdFind_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the ViewFind menu item.
    
    Call mnuViewFind_Click
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

Private Sub cmdSave_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the FileSave menu item.
    
    Call mnuFileSave_Click
End Sub

Private Sub Form_Load()
    ' Purpose:  Show the form and login to the server
    
    ' Set up error handling
    On Error GoTo PROC_ERR
       
    ' Connect to the database
    If gconDatabase Is Nothing Then
        Set gconDatabase = gclsSQLServer.Connect( _
            gclsMESApplication.ApplicationRole, _
            gclsMESApplication.ApplicationPassword)
    
        If gconDatabase.State <> adStateOpen Then
            Err.Raise gconDatabase.Errors(0).NativeError, _
                "Form_Load", gconDatabase.Errors(0).Description
        End If
    End If
    
    ' Hide buttons and menu choices for non-update
    If Not gblnUpdate Then
        cmdAdd.Enabled = False
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        
        mnuFileNew.Enabled = False
        mnuFileSave.Enabled = False
        mnuFileCancel.Enabled = False
        
        txtModel.Enabled = False
        cboLocation.Enabled = False
        cboLine.Enabled = False
        cboSubAssy.Enabled = False
    End If
    
    ' Retrieve the data
    Call RetrieveData
    
    If mrsDatabase.RecordCount > 1 Then
        frmMNBModelLocationSubAssemblyDisplay.Show vbModal
        mblnRecChanged = False
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
        ' Validate controls
            If Not ValidEntries Then
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
    ' Purpose:  Close and dereference the connection and
    '           recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Check to see that the connection has been instantiated.
    If Not gconDatabase Is Nothing Then
    
        ' If the connection is open, close it.
        If gconDatabase.State = adStateOpen Then
        
            ' Check to see if the recordset is instantiated.
            If Not mrsDatabase Is Nothing Then
            
                ' If the recordset is open, close it.
                If mrsDatabase.State = adStateOpen Then
                    mrsDatabase.Close
                End If
                Set mrsDatabase = Nothing
            End If
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Unload", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuFileCancel_Click()
    ' Purpose:  This procedure will delete a newly added
    '           record or return an updated record to its
    '           original state.
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ' Set up a field object
    Dim fld As ADODB.Field
    
    ' If the record is newly added, delete it.
    With mrsDatabase
        If .EditMode = adEditAdd Then
            .Delete
            .MoveFirst
            GoTo PROC_EXIT
        End If
        
        ' Loop through the existing record and reset
        ' each field to it original value.
        For Each fld In .Fields
            fld.Value = fld.OriginalValue
        Next fld
        Set fld = Nothing
        mrsDatabase_MoveComplete adRsnMove, Nothing, adStatusOK, _
            Nothing
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuFileCancel_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdCancel_Click()
    mnuFileCancel_Click
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
    With mrsDatabase
        .Delete
        If .RecordCount = 0 Then
            mrsDatabase.AddNew
        Else
            .MoveNext
            If .EOF Then
                .MoveFirst
            End If
        End If
    End With
End Sub

Private Sub mnuFileExit_Click()
    ' Purpose:  Close the application
    
    Unload frmMain
End Sub

Private Sub mnuFileNew_Click()
    ' Purpose:  Create a new record and set fields on the
    '           form to their inital values.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate the controls
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Call a Recordset Add
    mrsDatabase.AddNew
    
    txtModel.Enabled = True
    cboLine.Enabled = True
    cboLocation.Enabled = True
    txtModel.SetFocus
    
    ' Check for an error being returned from the operation.
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "cmdAdd_Click", _
            gconDatabase.Errors(0).Description
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdAdd_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuFileSave_Click()
    ' Purpose:  Save the current changes to the database
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate the controls
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Attempt to update the data
    mrsDatabase.UpdateBatch
    
    ' Check for errors
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuSave_Click", _
            gconDatabase.Errors(0).Description
    End If
    
    ' Reset the Record Changed flag to false.
    mblnRecChanged = False
    
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

Private Sub mnuViewList_Click()
    ' Purpose:  Find a specific record in the database and move
    '            to that record.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
        
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If

    ' Call the procedure to load the form to be used to find
    ' a specific Model/Line.
    frmMNBModelLocationSubAssemblyDisplay.Show vbModal
    

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewFind_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewFind_Click()
    ' Purpose:  Find a specific record in the database and move
    '            to that record.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
        
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If

    ' Call the procedure to load the form to be used to find
    ' a specific Model/Line.
    frmMNBModelLocationSubAssemblyFind.Show vbModal
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewFind_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewFirst_Click()
    ' Purpose:  Move to the first record in the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Move to the first record in the recordset
    mrsDatabase.MoveFirst
    
    ' Check return code from Recordset operation.  Raise an
    ' error if the operation failed.  If the errors collection
    ' of the connection object contains errors, this means the
    ' operation failed.
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuViewFirst_Click", _
            gconDatabase.Errors(0).Description
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewFirst_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub mnuViewPrevious_Click()
    ' Purpose:  Move to the previous record in the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Move to the previous record in the recordset
    With mrsDatabase
        .MovePrevious
        
        ' If the cursor is at the beginning of file, move to
        ' the last record.
        If .BOF Then
            .MoveLast
        End If
    End With
    
    ' Check return code from Recordset operation.  Raise an
    ' error if the operation failed.  If the errors collection
    ' of the connection object contains errors, this means the
    ' operation failed.
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuViewPrevious_Click", _
            gconDatabase.Errors(0).Description
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewPrevious_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewNext_Click()
    ' Purpose:  Move to the next record in the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Move to the next record in the recordset
    With mrsDatabase
        .MoveNext
        
        ' If the cursor is at the end of file, move to
        ' the first record.
        If .EOF Then
            .MoveFirst
        End If
    End With
    
    ' Check return code from Recordset operation.  Raise an
    ' error if the operation failed.  If the errors collection
    ' of the connection object contains errors, this means the
    ' operation failed.
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuViewNext_Click", _
            gconDatabase.Errors(0).Description
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewNext_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuViewLast_Click()
    ' Purpose:  Move to the last record in the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Validate controls before moving
    If Not ValidEntries Then
        GoTo PROC_EXIT
    End If
    
    ' Move to the last record in the recordset
    mrsDatabase.MoveLast
    
    ' Check return code from Recordset operation.  Raise an
    ' error if the operation failed.  If the errors collection
    ' of the connection object contains errors, this means the
    ' operation failed.
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuViewLast_Click", _
            gconDatabase.Errors(0).Description
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewLast_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mrsDatabase_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    ' Purpose:  Fill the fields after a move in the
    '           recordset is complete.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim intI As Integer
    ' If the recordset is at a valid record, fill the
    ' controls.
    With mrsDatabase
        If Not .EOF And Not .BOF Then
            ' Add code here to fill the controls on the form
            ' with the data from the Recordset.
            txtModel.Text = !model_number
            cboLine.Text = !line_id
            cboLocation.Text = !stocking_location_id
            
            If !model_number <> "" And !line_id <> "" Then
                Call RetrieveLocationData(!model_number, !line_id)
            
                For intI = 0 To UBound(marrstrLocation)
                    If marrstrLocation(intI) = Trim(!stocking_location_id) Then
                        cboLocation.ListIndex = intI
                        Exit For
                    End If
                Next
            End If
            
            cboSubAssy.Text = !sub_assembly_id
        End If
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mrsDatabase_MoveComplete", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
    
End Sub

Sub RetrieveData()
    ' Purpose:  Instantiate and open the recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Instantiate the recordset
    Set mrsDatabase = New ADODB.Recordset
    
    ' Set values of fields
    With mrsDatabase
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        ' Change the literal below to the name of your view
        .Source = "select * from v_mnb_model_location_sub_assembly " & _
            "order by model_number, line_id, stocking_location_id"
        .LockType = adLockBatchOptimistic
        .Open
    End With
    
    ' Check for errors returned from the recordset
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise vbObjectError + 1000, "RetrieveData", _
        gconDatabase.Errors(0).Description
    End If
    
    ' if no records were retrieved, add a new record to the
    ' recordset and reset fields to their original value.
    If mrsDatabase.EOF Then
        mrsDatabase.AddNew
        txtModel.Text = vbNullString
        cboLine.Text = vbNullString
        cboLocation.Text = vbNullString
        cboSubAssy.Text = vbNullString
        txtModel.Enabled = True
        cboLine.Enabled = True
        cboLocation.Enabled = True
    Else
        mblnRecChanged = False
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Function ValidEntries() As Boolean
    ' Purpose:  Validate fields before changing reocrds or
    '           updating.
    ' Returns:  Boolean determining whether the updates were
    '           successful.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsDuplicate As ADODB.Recordset
    
    ' Set up variable to hold cancel parameter
    Dim blnCancel As Boolean
    
    ' Validate the Model Number field.
    Call txtModel_Validate(blnCancel)
    
    ' If the Model Number was not valid or was missing, set focus
    ' back to the Model Number and exit the function.
    If blnCancel Then
        txtModel.SetFocus
        GoTo PROC_EXIT
    End If
       
    ' Validate the line id field.
    Call cboLine_Validate(blnCancel)
    
    ' If the line id was not valid or was missing, set focus
    ' back to the line id and exit the function.
    If blnCancel Then
        cboLine.SetFocus
        GoTo PROC_EXIT
    End If
       
    ' Validate the location field.
    Call cboLocation_Validate(blnCancel)
    
    ' If the location was not valid or was missing, set focus
    ' back to the location and exit the function.
    If blnCancel Then
        cboLocation.SetFocus
        GoTo PROC_EXIT
    End If
    
'    If mrsDatabase.EditMode = adEditAdd Then
'        Set rsDuplicate = mrsDatabase.Clone
'
'        With rsDuplicate
'            .Find ("model_number = '" & txtModel.Text & "'")
'
'            Do While Not .EOF
'                If !model_number <> mrsDatabase!model_number Then
'                    Exit Do
'                End If
'
'                If !line_id = mrsDatabase!line_id And _
'                   !stocking_location_id = mrsDatabase!stocking_location_id Then
'                    blnCancel = True
'                    MsgBox "Model/Line/Location already exists "
'                    Exit Do
'                End If
'                .MoveNext
'            Loop
'            .Close
'        End With
'
'        Set rsDuplicate = Nothing
'    End If
'
'    If blnCancel Then
'        GoTo PROC_EXIT
'    End If
    
    ' Validate the sub assembly field.
    Call cboSubAssy_Validate(blnCancel)
    
    ' If the sub assembly was not valid or was missing, set focus
    ' back to the sub assembly and exit the function.
    If blnCancel Then
        cboSubAssy.SetFocus
        GoTo PROC_EXIT
    End If
        
PROC_EXIT:

    ' Set the return value of the function to the opposite of
    ' the cancel boolean.  This is done because the validate
    ' procedures called set the cancel boolean to true if the
    ' validation failed and false if it succeeded.
    ValidEntries = Not blnCancel
    Exit Function
    
PROC_ERR:
    Call ShowError(Me.Name, "ValidEntries", Err.Number, _
        Err.Description)
    blnCancel = True
    GoTo PROC_EXIT
End Function

'Sub RetrieveModelData()
'    ' Purpose:  Fill the model combo box
'
'    ' Set up error handling
'    On Error GoTo PROC_ERR
'
'    Dim rsList As ADODB.Recordset
'    Set rsList = New ADODB.Recordset
'
'    cboModel.Clear
'
'    With rsList
'        Set .ActiveConnection = gconDatabase
'        .CursorLocation = adUseClient
'        .CursorType = adOpenForwardOnly
'        .LockType = adLockReadOnly
'        .Source = "select distinct model_number from v_mnb_model_part " & _
'            "order by model_number asc"
'        .Open
'
'        Do While Not .EOF
'            cboModel.AddItem Trim(!model_number)
'            .MoveNext
'        Loop
'        .Close
'    End With
'
'    Set rsList = Nothing
'
'PROC_EXIT:
'    Exit Sub
'
'PROC_ERR:
'    Call ShowError(Me.Name, "RetrieveModelData", _
'        Err.Number, Err.Description)
'End Sub

Sub RetrieveLineData(strModelNumber As String)
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    
    cboLine.Clear
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select distinct line_id from v_mnb_model_part_stocking_location " & _
            "where model_number = '" & strModelNumber & "' " & _
            "order by line_id asc"
        .Open
        
        Do While Not .EOF
            cboLine.AddItem Trim(!line_id)
            .MoveNext
        Loop
        .Close
    End With
    
    Set rsList = Nothing
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveLineData", _
        Err.Number, Err.Description)
End Sub

Private Sub RetrieveLocationData(strModelNumber As String, strLineId As String)
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    Set rsList = New ADODB.Recordset
    
    cboLocation.Clear
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select distinct a.stocking_location_id, b.stocking_location_description " & _
            "from v_mnb_model_part_stocking_location a " & _
            "join v_prod_stocking_location b on " & _
            "a.stocking_location_id = b.stocking_location_id " & _
            "where model_number = '" & strModelNumber & "' " & _
            "and line_id = '" & strLineId & "' " & _
            "order by b.stocking_location_description asc"
        .Open
        
        If .RecordCount > 0 Then
            ReDim marrstrLocation(.RecordCount - 1)
        End If

        intIndex = 0
        Do While Not .EOF
            cboLocation.AddItem Trim(!stocking_location_description)
            marrstrLocation(intIndex) = Trim(!stocking_location_id)
            cboLocation.ItemData(cboLocation.NewIndex) = intIndex
            intIndex = intIndex + 1
            .MoveNext
        Loop
        .Close
    End With
    
    Set rsList = Nothing
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveLocationData", _
        Err.Number, Err.Description)
End Sub

Sub RetrievePartData(strModelNumber As String, strLineId As String, _
                        strLocationId As String)
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
       
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select distinct part_id from v_mnb_model_part_stocking_location " & _
            "where model_number = '" & strModelNumber & "' " & _
            "and line_id = '" & strLineId & "' " & _
            "and stocking_location_id <> '" & strLocationId & "' " & _
            "order by part_id asc"
        .Open
        
        cboSubAssy.Clear
        
        Do While Not .EOF
            cboSubAssy.AddItem Trim(!part_id)
            .MoveNext
        Loop
        .Close
    End With
    
    Set rsList = Nothing
        
PROC_EXIT:
    
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrievePartData", _
        Err.Number, Err.Description)
End Sub

Private Sub txtModel_Change()
    mblnRecChanged = True
End Sub

Private Sub txtModel_GotFocus()
    ' Purpose:  Select the field for easy update
    
    If Len(txtModel.Text) > 0 Then
        txtModel.SelStart = 0
        txtModel.SelLength = Len(txtModel.Text)
    End If
End Sub

Private Sub txtModel_KeyPress(KeyAscii As Integer)
    ' Purpose:  Change any character to uppercase
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    mblnRecChanged = True
End Sub

Private Sub txtModel_LostFocus()
    RetrieveLineData (txtModel.Text)
End Sub

Private Sub txtModel_Validate(Cancel As Boolean)
    ' Purpose:  Validate that the Model is valid
    '           in length.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
                
    ' Set a variable to hold a copy of the recordset to be used
    ' to check for a duplicate.
    Dim rsModel As ADODB.Recordset
    
    ' If the length of the field is zero, give an error
    If Len(txtModel.Text) = 0 Then
        Cancel = True
        MsgBox "Model is required", _
            vbExclamation + vbOKOnly, _
            "Model Validation"
        GoTo PROC_EXIT
    End If
        
    ' Check to see if the Model has changed.  If it has,
    ' see if the Model entered already exists.  If it
    ' exists, give an error.  If not, change the model number
    ' in the recordset.
    If Trim(txtModel.Text) <> Trim(mrsDatabase!model_number) Then
        Set rsModel = New ADODB.Recordset

        ' Validate Model Number against the database
        With rsModel
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .Source = "select line_id " & _
                "from v_mnb_model_part_stocking_location " & _
                "where model_number = '" & txtModel.Text & "' "
            .LockType = adLockReadOnly
            .Open

            If .RecordCount = 0 Then
                Cancel = True
                MsgBox "Invalid Model Entered", _
                    vbExclamation + vbOKOnly, _
                    "Model Validation"
            Else
                mrsDatabase!model_number = txtModel.Text
            End If
            .Close
        End With

        Set rsModel = Nothing
    End If
    
PROC_EXIT:

    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "txtModel_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

