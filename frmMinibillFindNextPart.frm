VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLineLocation 
   Caption         =   "MiniBill - Line / Stocking Location Association"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8775
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
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   StartUpPosition =   1  'CenterOwner
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   6945
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   8775
      _cx             =   15478
      _cy             =   12250
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
      Begin VB.ComboBox cboLine 
         Height          =   360
         Left            =   1080
         TabIndex        =   0
         Top             =   900
         Width           =   855
      End
      Begin VB.CommandButton cmdFirst 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   60
         Picture         =   "frmMinibillFindNextPart.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Display First Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdPrevious 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   420
         Picture         =   "frmMinibillFindNextPart.frx":05DA
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Display Previous Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdNext 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   780
         Picture         =   "frmMinibillFindNextPart.frx":0BB4
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Display Next Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdLast 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   1140
         Picture         =   "frmMinibillFindNextPart.frx":118E
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Display Last Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdClose 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   2280
         Picture         =   "frmMinibillFindNextPart.frx":1768
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Close This Form"
         Top             =   180
         Width           =   300
      End
      Begin VB.ListBox lstAvailable 
         Height          =   4140
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1980
         Width           =   3000
      End
      Begin VB.ListBox lstSelected 
         Height          =   4140
         Left            =   5460
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1980
         Width           =   3000
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add >"
         Height          =   360
         Left            =   3690
         TabIndex        =   8
         Top             =   2250
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "< &Remove"
         Height          =   390
         Left            =   3690
         TabIndex        =   7
         Top             =   2850
         Width           =   1455
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "<< &Clear All"
         Height          =   360
         Left            =   3720
         TabIndex        =   6
         Top             =   3525
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   1560
         Picture         =   "frmMinibillFindNextPart.frx":1ECA
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Save Changes"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdHelp 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   1920
         Picture         =   "frmMinibillFindNextPart.frx":2474
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Help"
         Top             =   180
         Width           =   300
      End
      Begin MSComctlLib.StatusBar staDBMaint 
         Height          =   390
         Left            =   120
         TabIndex        =   14
         Top             =   6600
         Width           =   8775
         _ExtentX        =   15478
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
               TextSave        =   "4:32 PM"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   10292
            EndProperty
         EndProperty
         Enabled         =   0   'False
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
         Left            =   420
         TabIndex        =   18
         Top             =   923
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Available Locations"
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
         Left            =   300
         TabIndex        =   17
         Top             =   1620
         Width           =   3195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Locations"
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
         Left            =   5460
         TabIndex        =   16
         Top             =   1620
         Width           =   2595
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   1
         X1              =   120
         X2              =   8700
         Y1              =   585
         Y2              =   585
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   0
         X1              =   120
         X2              =   8700
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
         TabIndex        =   15
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
Attribute VB_Name = "frmLineLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents mrsDatabase As ADODB.Recordset
Attribute mrsDatabase.VB_VarHelpID = -1

Private mblnRecChanged As Boolean
Private marrstrLocation() As String
Private marrstrLocationDescription() As String
Private strLocationId As String
Private strLocationDescription As String




Private Sub cboLine_Change()
    cboFindFirst cboLine
End Sub

Private Sub cboLine_Click()
    ' Purpose:  Change the Location settings based on a
    '           change to the Line.
    
    If cboLine.ListIndex > -1 Then
        BuildChoices
    Else
        ClearChoices
    End If
End Sub

Private Sub BuildChoices()
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ' Declare variable
    Dim intIndex As Integer
    Dim intArraySize As Integer
    
    ' Set variable with the number of Location ID's in the
    ' array.
    intArraySize = UBound(marrstrLocation)
    
    ' If the available list does not contain all of the
    ' available products, rebuild the listbox from the
    ' array.
    If lstAvailable.ListCount <> _
            intArraySize + 1 Then
        
        ' Clear the listbox
        lstAvailable.Clear
        
        ' Loop through the array and add entries to the
        ' listbox.
        For intIndex = 0 To intArraySize
            lstAvailable.AddItem _
                marrstrLocation(intIndex) & " " & _
                marrstrLocationDescription(intIndex)
        Next intIndex
    End If
    
    ' If any records have changed, save them.
    If mblnRecChanged Then
        Call SaveChanges
    End If
    
   
    cmdAdd.Enabled = True
    cmdClearAll.Enabled = True
    cmdRemove.Enabled = True
    
    ' Using the recordset
    With mrsDatabase
        
        ' If no records were retrieved, exit the sub
        If .RecordCount = 0 Then
            GoTo PROC_EXIT
        End If
        
        ' Move to the first record
        .MoveFirst
               
        ' Clear the selected list box
        lstSelected.Clear
        
        ' If the Line has changed, find it in the recordset
        If !line_id <> cboLine.Text Then
            .Find "Line_Id = '" & cboLine.Text & "'"
        End If
        
        
        ' Loop through the recordset while the  line is
        ' equal to the line provided.
        Do While Not .EOF
            
            ' If the Line_Id is equal to the one selected..
            If mrsDatabase!line_id <> cboLine.Text Then
                Exit Do
            End If
            
            ' Set the variable holding the Location ID
            strLocationId = !stocking_location_id
            
            ' Find the description for the Location ID
            ' code.
            Call FindDescription
            
            ' Call the procedure to remove the Location Id
            ' from the available list.
            Call RemoveFromAvailable(strLocationId)
            
            ' Add to the selected listbox.
            lstSelected.AddItem strLocationId & _
                " " & strLocationDescription
            
            ' Read the next record.
            .MoveNext
        Loop
    End With
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "CboLine_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub ClearChoices()
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ' Declare variable
    Dim intIndex As Integer
    Dim intArraySize As Integer
    
    ' Set variable with the number of Location ID's in the
    ' array.
    intArraySize = UBound(marrstrLocation)
    
    ' If the available list does not contain all of the
    ' available products, rebuild the listbox from the
    ' array.
    If lstAvailable.ListCount <> _
            intArraySize + 1 Then
        
        ' Clear the listbox
        lstAvailable.Clear
        
        ' Loop through the array and add entries to the
        ' listbox.
        For intIndex = 0 To intArraySize
            lstAvailable.AddItem _
                marrstrLocation(intIndex) & " " & _
                marrstrLocationDescription(intIndex)
        Next intIndex
    End If
    
    ' If any records have changed, save them.
    If mblnRecChanged Then
        Call SaveChanges
    End If
    
               
    ' Clear the selected list box
    lstSelected.Clear
        
    cmdAdd.Enabled = False
    cmdClearAll.Enabled = False
    cmdRemove.Enabled = False
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "CboLine_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub



Private Sub cboLine_GotFocus()
    cboLine.SelStart = 0
    cboLine.SelLength = Len(cboLine.Text)
End Sub

Private Sub cboLine_KeyPress(KeyAscii As Integer)
    cboKeyPress cboLine, KeyAscii
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

Private Sub cmdAdd_Click()
    ' Purpose:  Add the selected record in the available
    '           column to the selected list and to the
    '           recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set up a string to hold the product
    Dim strLocation As String
    Dim intLoopIndex As Integer
    
    ' If no product was selected send an error and exit
    ' the sub.
    If lstAvailable.ListIndex = -1 Then
        MsgBox "Please select an entry to be selected."
        GoTo PROC_EXIT
    End If
    
    ' Add the item to the selected listbox
    lstSelected.AddItem lstAvailable.Text
    
    ' Set the Location code
    strLocationId = Mid(lstAvailable.Text, 1, 5)
    
    ' Add the record to the recordset and load the
    ' fields.
    mrsDatabase.AddNew
    mrsDatabase!line_id = cboLine.Text
    mrsDatabase!stocking_location_id = strLocationId
'    mrsDatabase!Line_LocationLast_updated = Now
'    mrsDatabase!Line_Locationupdated_by = _
'        gclsSQLServer.UserID
    mblnRecChanged = True
    
    ' Remove from the available listbox
    lstAvailable.RemoveItem _
        lstAvailable.ListIndex
    
    ' Set the listindex of the selected listbox to the
    ' index just added.
    lstSelected.ListIndex = lstSelected.NewIndex
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, _
        Err.Description)
    Unload Me

End Sub

Private Sub cmdRemove_Click()
    ' Purpose:  Remove a Location ID from the selected list
    '           and delete the record from the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' If no item was selected in the Selected listbox,
    ' display an error and leave the sub.
    If lstSelected.ListIndex = -1 Then
        MsgBox "There is no entry selected"
        GoTo PROC_EXIT
    End If
    
    ' Find the item in the recordset and delete it.
    With mrsDatabase
        .MoveFirst
        ' Find the Line_Id
        .Find "Line_Id = '" & cboLine.Text & "'"
        
        ' Find the Location code
        strLocationId = Mid(lstSelected.Text, 1, 5)
        If strLocationId <> !stocking_location_id Then
            .Find "stocking_Location_id = '" & strLocationId & "'"
        End If
        
        ' Delete the record.
        .Delete
        mblnRecChanged = True
    End With
        
    ' Add the item to the available listbox
    lstAvailable.AddItem lstSelected.Text
    
    ' Remove the item from the selected listbox
    lstSelected.RemoveItem lstSelected.ListIndex
    
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, _
        Err.Description)
    Unload Me

End Sub

Private Sub cmdClearAll_Click()
    ' Purpose:  Remove all products from the selected
    '           list and from the recordset.
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ' Set the variable for the looping index
    Dim intLoopIndex As Integer
    Dim intReturnValue As Integer
    
    intReturnValue = MsgBox("Are you sure????", vbYesNo, "Delete All Locations in Category")
    If intReturnValue = vbYes Then
    
    ' Loop through the selected listbox.
    For intLoopIndex = 0 To lstSelected.ListCount - 1
    
        ' Find the record in the recordset for the
        ' selected product.
        With mrsDatabase
            .MoveFirst
            .Find "Line_Id = '" & cboLine.Text & "'"
            .Find "stocking_Location_id = '" & _
                Mid(lstSelected.List(intLoopIndex), 1, 5) & "'"
            
            ' Delete the record
            .Delete
        End With
        
        ' Add the item to the available listbox
        lstAvailable.AddItem _
            lstSelected.List(intLoopIndex)
        mblnRecChanged = True
    Next intLoopIndex
    
    ' Clear the selected box.
    lstSelected.Clear
    
    'Related to msgbox question - Cancelling the Clear All command
    End If
    
    
PROC_EXIT:
    
    Exit Sub
    
PROC_ERR:
    
    Call ShowError(Me.Name, "Command4", Err.Number, _
        Err.Description)
    Unload Me
    
    
End Sub

Private Sub Form_Load()
    ' Purpose:  Show the form and login to the server
    
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
    
    ' Disable fields if update is not allowed.
    If Not gblnUpdate Then
        cmdAdd.Enabled = False
        cmdRemove.Enabled = False
        cmdClearAll.Enabled = False
    End If
    
    
    ' Retrieve the data
    Call RetrieveLocationData
    Call RetrieveLineData
    Call RetrieveLineLocationData
    
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
    End If
    
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
    ' Purpose:  When an item in the slected listbox was
    '           clicked, fill the data for that record
    '           and make the fields visible.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    If lstSelected.ListIndex = -1 Then
        Exit Sub
    End If
    
    ' Set a field for the product code.
    Dim strLocationId As String
    
    
    ' Place the slected product code in the field
    strLocationId = Mid(lstSelected.Text, 1, 5)
   
    ' Find the Line code in the recordset and
    ' fill the fields with data.
    With mrsDatabase
        .MoveFirst
        If !line_id <> cboLine.Text Then
            .Find "Line_Id = '" & cboLine.Text & "'"
        End If
        If !stocking_location_id <> strLocationId Then
            .Find "stocking_Location_id = '" & strLocationId & "'"
        End If
    End With

   
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "lstSelected_Click", _
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



Private Sub mnuFileExit_Click()
    ' Purpose:  Exit the application
    Unload frmMain
End Sub

Private Sub mnuFileSave_Click()
    ' Purpose:  Save the current changes to the database
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Attempt to update the data
    mrsDatabase.UpdateBatch
    
    ' Check for errors
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "mnuSave_Click", gconDatabase.Errors(0).Description
    End If
    
    ' Requery the recordset
    mrsDatabase.Requery
    
    ' Reset record changed flag
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


Private Sub mnuViewFirst_Click()
    ' Purpose:  Move to the first record in the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' If any changes have been made, save them before
    ' changing record position.
    If mblnRecChanged Then
        Call SaveChanges
    End If
    
    ' Set the line combo box to the first line.
    cboLine.ListIndex = 0

    
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
    
    ' if not at the first record, move to the previous
    ' item.  If at the first item, move to the last item.
    If cboLine.ListIndex > 0 Then
        cboLine.ListIndex = cboLine.ListIndex - 1
    Else
        cboLine.ListIndex = cboLine.ListCount - 1
    End If
    
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
    
    ' If any changes have been made, save them before
    ' changing record position.
    If mblnRecChanged Then
        Call SaveChanges
    End If
    
    ' If not at the last item, move to the next one,
    ' otherwise move to the first item.
    If cboLine.ListIndex < cboLine.ListCount - 1 Then
        cboLine.ListIndex = cboLine.ListIndex + 1
    Else
        cboLine.ListIndex = 0
    End If
    
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
    
    ' Move to the last item.
    cboLine.ListIndex = cboLine.ListCount - 1
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewLast_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub



Private Sub RetrieveLineLocationData()
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
        .Source = "select * from v_prod_line_stocking_Location " & _
            "order by Line_id asc, stocking_Location_id asc"
        .LockType = adLockBatchOptimistic
        .Open
    End With
    
    ' Check for errors returned from the recordset
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "RetrieveLineLocationData", _
            gconDatabase.Errors(0).Description
    End If
    
    cboLine.ListIndex = 0
            
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveLineLocationData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub RemoveFromAvailable(strProdCode As String)
    ' Purpose:  Removes a Location Code from the available lis
    '           which appears on the Selected list.
    ' Input:    strLocation - The Location Code to be
    '           removed.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare a variable for the index
    Dim intIndex As Integer
    
    ' Loop through the listbox when a match is found,
    ' remove the item and leave the sub.
    For intIndex = 0 To lstAvailable.ListCount - 1
        If Mid(lstAvailable.List(intIndex), 1, Len(strLocationId)) = _
                strLocationId Then
            lstAvailable.RemoveItem (intIndex)
            GoTo PROC_EXIT
        End If
    Next intIndex
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RemoveFromAvailable", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
    
End Sub

Private Sub RetrieveLineData()
    ' Purpose:  Instantiate and open the recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare Recordset to hold line
    Dim RsLine As ADODB.Recordset
    Dim intIndex As Integer
    
    ' Instantiate the recordset
    Set RsLine = New ADODB.Recordset

    ' Set values of fields
    With RsLine
        'tells the recordset where to get its data from
        'i.e. gcondatabase="Provider=SQLOLEDB.1;PASSWORD=090400;USERID=SCHIMC
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        ' Change the literal below to the name of your view
        .Source = "select * from v_prod_Line order by line_id asc"
        '.LockType = adLockBatchOptimistic
        .LockType = adLockReadOnly
        .Open
    End With

    ' Check for errors returned from the recordset
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "RetrieveLineData", _
            gconDatabase.Errors(0).Description
    End If

    ' if no records were retrieved, add a new record to the
    ' recordset and reset fields to their original value.
    If RsLine.EOF Then
        MsgBox ("No records were retrieved from Line table")
        GoTo PROC_EXIT
    End If

    ' Go to the first record in the recordset and set the
    ' line ID
    With RsLine
        ' Loop through the file
        Do While Not .EOF
            cboLine.AddItem RsLine!line_id
            .MoveNext
            intIndex = intIndex + 1
        Loop
        .Close
    End With
    
    Set RsLine = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveLineData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub RetrieveLocationData()
    ' Purpose:  Retrieve the Location ID's from the table and
    '           build the array of Locations and the Location Description
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare the array size
    Dim intArrayCount As Integer
    
    ' Declare the recordset variable
    Dim rsLocation As ADODB.Recordset
    
    ' Instantiate the recordset
    Set rsLocation = New ADODB.Recordset

    ' Set values of fields
    With rsLocation
        'tells the recordset where to get its data from
        'i.e. gcondatabase="Provider=SQLOLEDB.1;PASSWORD=090400;USERID=SCHIMC
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        ' Change the literal below to the name of your view
        .Source = "v_prod_stocking_Location"
        
        '.LockType = adLockBatchOptimistic
        .LockType = adLockReadOnly
        .Open
    End With

    ' Check for errors returned from the recordset
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "RetrieveLocationData", _
            gconDatabase.Errors(0).Description
    End If

    ' Clear the available products.
    lstAvailable.Clear
    
    ' Using the recordset, load the array and the
    ' available listbox.
    With rsLocation
        ' Move to the first record.
        .MoveFirst
        
        ' Loop through the recordset.
        Do While Not .EOF
            
            ' Reset the size of the arrays.
            ReDim Preserve marrstrLocation(intArrayCount)
            ReDim Preserve marrstrLocationDescription(intArrayCount)
            
            ' Add the product information to the arrays.
            marrstrLocation(intArrayCount) = rsLocation _
                !stocking_location_id
            marrstrLocationDescription(intArrayCount) = rsLocation _
                !stocking_location_description
            
            ' Move to the next record.
            .MoveNext
            
            ' Increment the array counter
            intArrayCount = intArrayCount + 1
        Loop
        .Close
    End With
    Set rsLocation = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveLocationData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub






Private Sub SaveChanges()

    On Error GoTo PROC_ERR
    
    ' Declare Return Code variable
    Dim intRetCode As Integer
    
    ' Ask the user if he would like to save the changes.
    intRetCode = MsgBox("Save Changes?", _
        vbQuestion + vbYesNoCancel)
    If intRetCode = vbYes Then
        Call mnuFileSave_Click
    Else
        mrsDatabase.CancelBatch
        mblnRecChanged = False
    End If
        
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "SaveChanges", _
        Err.Number, Err.Description)
End Sub


Private Sub FindDescription()
    ' Purpose:  Find the description for the Location code
    '           code.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set up variables.
    Dim intLoopIndex As Integer

    ' Loop through the array and find the matching
    ' product code.
    For intLoopIndex = 0 To UBound(marrstrLocation)
        If strLocationId = marrstrLocation(intLoopIndex) Then
            strLocationDescription = marrstrLocationDescription(intLoopIndex)
            Exit For
        End If
            
    Next intLoopIndex

   
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FindDescription", _
        Err.Number, Err.Description)
    
End Sub


