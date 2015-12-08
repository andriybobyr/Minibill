VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategoryPart 
   Caption         =   "MiniBill - Category by Part"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   12600
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
   ScaleHeight     =   538
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   840
   StartUpPosition =   1  'CenterOwner
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   8070
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   12600
      _cx             =   22225
      _cy             =   14235
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
      Begin VB.CheckBox chkRacman 
         Caption         =   "Display RacMan Categories"
         Height          =   315
         Left            =   4920
         TabIndex        =   26
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CheckBox chkUnassigned 
         Caption         =   "Display Only Parts With No Category Assignment"
         Height          =   435
         Left            =   2100
         TabIndex        =   25
         Top             =   2460
         Width           =   4815
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   435
         Left            =   7920
         TabIndex        =   24
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtPart 
         Height          =   360
         Left            =   300
         MaxLength       =   15
         TabIndex        =   4
         Top             =   2940
         Width           =   2850
      End
      Begin VB.Frame fraSelection 
         Caption         =   "Part List Specification"
         Height          =   735
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   9075
         Begin VB.ComboBox cboModel 
            Height          =   360
            Left            =   4200
            TabIndex        =   3
            Top             =   240
            Width           =   2715
         End
         Begin VB.OptionButton optByModel 
            Caption         =   "Display By Model"
            Height          =   375
            Left            =   2280
            TabIndex        =   2
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optAll 
            Caption         =   "Display All Parts"
            Height          =   255
            Left            =   300
            TabIndex        =   1
            Top             =   300
            Width           =   1995
         End
      End
      Begin VB.CommandButton cmdFirst 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   60
         Picture         =   "frmCategoryPart.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Display First Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdPrevious 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   420
         Picture         =   "frmCategoryPart.frx":05DA
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Display Previous Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdNext 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   780
         Picture         =   "frmCategoryPart.frx":0BB4
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Display Next Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdLast 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   1140
         Picture         =   "frmCategoryPart.frx":118E
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Display Last Entry"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdClose 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   2280
         Picture         =   "frmCategoryPart.frx":1768
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Close This Form"
         Top             =   180
         Width           =   300
      End
      Begin VB.ComboBox cboCategory 
         Height          =   360
         ItemData        =   "frmCategoryPart.frx":1ECA
         Left            =   4860
         List            =   "frmCategoryPart.frx":1ECC
         TabIndex        =   0
         Text            =   "cboCategory"
         Top             =   900
         Width           =   3855
      End
      Begin VB.ListBox lstAvailable 
         Height          =   4140
         Left            =   300
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   3375
         Width           =   5100
      End
      Begin VB.ListBox lstSelected 
         Height          =   4140
         Left            =   7260
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   3375
         Width           =   4860
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add >"
         Height          =   360
         Left            =   5610
         TabIndex        =   12
         Top             =   3510
         Width           =   1455
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "< &Remove"
         Height          =   390
         Left            =   5610
         TabIndex        =   11
         Top             =   4110
         Width           =   1455
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "<< &Clear All"
         Height          =   360
         Left            =   5640
         TabIndex        =   10
         Top             =   4785
         Width           =   1455
      End
      Begin VB.CommandButton cmdSave 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   1560
         Picture         =   "frmCategoryPart.frx":1ECE
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Save Changes"
         Top             =   180
         Width           =   300
      End
      Begin VB.CommandButton cmdHelp 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   1920
         Picture         =   "frmCategoryPart.frx":2478
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Help"
         Top             =   180
         Width           =   300
      End
      Begin MSComctlLib.StatusBar staDBMaint 
         Height          =   390
         Left            =   0
         TabIndex        =   18
         Top             =   7680
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
               TextSave        =   "4:31 PM"
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   10292
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.Frame Frame1 
         Caption         =   "Category Selections"
         Height          =   975
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   9075
         Begin VB.OptionButton optCategoryIDSequence 
            Caption         =   "Category ID Sequence"
            Height          =   240
            Left            =   720
            TabIndex        =   29
            Top             =   600
            Width           =   3135
         End
         Begin VB.OptionButton optCategoryDescSequence 
            Caption         =   "Category Description Seq."
            Height          =   240
            Left            =   720
            TabIndex        =   28
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   0
         Left            =   3480
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Available Parts"
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
         TabIndex        =   21
         Top             =   2580
         Width           =   3195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Parts"
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
         Left            =   7455
         TabIndex        =   20
         Top             =   2715
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
         TabIndex        =   19
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
Attribute VB_Name = "frmCategoryPart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents mrsDatabase As ADODB.Recordset
Attribute mrsDatabase.VB_VarHelpID = -1

Private mblnRecChanged As Boolean
Private marrstrpart() As String
Private marrstrAllParts() As String
Private marrstrAllPartDesc() As String
Private mblnLoadedAllParts As Boolean
Private marrstrpartDescription() As String
Private arrstrCategory() As String
Private strPartId As String
Private strPartDescription As String
Private mstrCategory As String
Private intLowerBoundry As Integer
Private intUpperBoundry As Integer
Private intSubGroups As Integer

Private Sub cboCategory_Change()
    cboFindFirst cboCategory
End Sub

Private Sub cboCategory_Click()
    ' Purpose:  Change the part settings based on a
    '           change to the Category.
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ' Declare variable
    Dim intIndex As Integer
    Dim intArraySize As Integer
    Dim strCategory As String
    
    ' If any records have changed, save them.
    If mblnRecChanged Then
        Call SaveChanges
    End If
    If cboCategory.ListIndex = -1 Then
        cboFindFirst cboCategory
    End If
    strCategory = arrstrCategory(cboCategory.ListIndex)
    If strCategory = mstrCategory Then
        Exit Sub
    End If
    
    
    ' Set variable with the number of part ID's in the
    ' array.
    intArraySize = UBound(marrstrpart)
    
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
                marrstrpart(intIndex) & " " & _
                marrstrpartDescription(intIndex)
        Next intIndex
    End If
    
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
        
        ' If the Category has changed, find it in the recordset
        If !Category_id <> strCategory Then
            .Find "Category_Id = '" & strCategory & "'"
        End If
        
        ' Loop through the recordset while the  line is
        ' equal to the line provided.
        Do While Not .EOF
            
            ' If the Category_Id is equal to the one selected..
            If mrsDatabase!Category_id = strCategory Then
                ' Set the variable holding the part ID
                strPartId = !part_id
                
                ' Find the description for the part ID
                ' code.
                Call FindDescription
                
                ' Call the procedure to remove the part Id
                ' from the available list.
                Call RemoveFromAvailable(strPartId)
                
                ' Add to the selected listbox.
                lstSelected.AddItem strPartId & _
                    " " & strPartDescription
            
            ' If the Category does not match, leave the loop.
            Else
                Exit Do
            End If
            
            ' Read the next record.
            .MoveNext
        Loop
    End With
    
    lstSelected.Refresh
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "CboCategory_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cboCategory_GotFocus()
    cboCategory.SelStart = 0
    cboCategory.SelLength = Len(cboCategory.Text)
End Sub

Private Sub cboCategory_KeyPress(KeyAscii As Integer)
    cboKeyPress cboCategory, KeyAscii
End Sub

Private Sub cboCategory_Validate(Cancel As Boolean)
    If Len(cboCategory.Text) = 0 Then
        MsgBox "Category is Required."
        Cancel = True
        Exit Sub
    End If
    
    cboFindFirst cboCategory
    If cboCategory.ListIndex = -1 Then
        MsgBox "Invalid Category"
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

Private Sub chkRacman_Click()
    If mblnRecChanged Then
        Call SaveChanges
    End If
    
    RetrieveCategoryData
    If cboCategory.ListCount > 0 Then
        cboCategory.ListIndex = 0
    End If
End Sub

Private Sub chkUnassigned_Click()
    Call RetrievePartData
    cboCategory_Click
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

Private Sub cmdRefresh_Click()
    ' Retrieve the data
    
    If optAll.Value Then
        Exit Sub
    End If
    
    If Len(cboModel.Text) = 0 Then
        MsgBox "Please select model"
        Exit Sub
    End If
    
    cboFindFirst cboModel
    If cboModel.ListIndex = -1 Then
        MsgBox "Please select a valid model"
        Exit Sub
    End If
    
    Call RetrievePartData
    Call cboCategory_Click
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
    Dim intLoopIndex As Integer
    
    ' If no product was selected send an error and exit
    ' the sub.
    If lstAvailable.ListIndex = -1 Then
        MsgBox "Please select an entry to be selected."
        GoTo PROC_EXIT
    End If
    
    ' Add the item to the selected listbox
    lstSelected.AddItem lstAvailable.Text
    
    ' Set the part code
    strPartId = Mid(lstAvailable.Text, 1, 15)
    
    ' Add the record to the recordset and load the
    ' fields.
    mrsDatabase.AddNew
    mrsDatabase!Category_id = arrstrCategory(cboCategory.ListIndex)
    mrsDatabase!part_id = strPartId

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
    ' Purpose:  Remove a part ID from the selected list
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
        ' Find the Category_Id
        .Find "Category_Id = '" & arrstrCategory(cboCategory.ListIndex) & "'"
        
        ' Find the part code
        strPartId = Mid(lstSelected.Text, 1, 15)
        If strPartId <> !part_id Then
            .Find "part_id = '" & strPartId & "'"
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
    
    intReturnValue = MsgBox("Are you sure????", vbYesNo, "Delete All Parts in Category")
    If intReturnValue = vbYes Then
    
    ' Loop through the selected listbox.
    For intLoopIndex = 0 To lstSelected.ListCount - 1
    
        ' Find the record in the recordset for the
        ' selected product.
        With mrsDatabase
            .MoveFirst
            .Find "Category_Id = '" & arrstrCategory(cboCategory.ListIndex) & "'"
            .Find "part_id = '" & _
                Mid(lstSelected.List(intLoopIndex), 1, 15) & "'"
            
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
    
    'End if to bypass clear all to change mind
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
    
    ' Disable fields if update is not allowed.
    If Not gblnUpdate Then
        cmdAdd.Enabled = False
        cmdRemove.Enabled = False
        cmdClearAll.Enabled = False
    End If
    
    gconDatabase.CommandTimeout = 350
    
    ' Retrieve the data
    optAll.Value = True
    Call RetrieveModelData
    mblnLoadedAllParts = False
    Call RetrievePartData
    mblnLoadedAllParts = True
    optCategoryDescSequence.Value = True
'    Call RetrieveCategoryData
    Call RetrieveCategoryPartData
    
    cboModel.Enabled = False
    
    frmProcessing.Hide
    
    
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
    ' Purpose:  When an item in the selected listbox was
    '           clicked, fill the data for that record
    '           and make the fields visible.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set a field for the product code.
    Dim strPartId As String
    
    
    ' Place the slected product code in the field
    strPartId = Mid(lstSelected.Text, 1, 15)
   
    ' Find the Category code in the recordset and
    ' fill the fields with data.
    With mrsDatabase
        .MoveFirst
        If !Category_id <> arrstrCategory(cboCategory.ListIndex) Then
            .Find "Category_Id = '" & arrstrCategory(cboCategory.ListIndex) & "'"
        End If
        If !part_id <> strPartId Then
            .Find "part_id = '" & strPartId & "'"
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
    
    MsgBox "Changes successfully completed!"
    
    txtPart.Text = ""
    
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
    cboCategory.ListIndex = 0

    
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
    If cboCategory.ListIndex > 0 Then
        cboCategory.ListIndex = cboCategory.ListIndex - 1
    Else
        cboCategory.ListIndex = cboCategory.ListCount - 1
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
    If cboCategory.ListIndex < cboCategory.ListCount - 1 Then
        cboCategory.ListIndex = cboCategory.ListIndex + 1
    Else
        cboCategory.ListIndex = 0
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
    cboCategory.ListIndex = cboCategory.ListCount - 1
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuViewLast_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub RetrieveCategoryPartData()
    ' Purpose:  Instantiate and open the recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
     frmProcessing.Label2 = "** Retrieve Category Part List ** " & Time
    frmProcessing.Refresh
    
    ' Instantiate the recordset
    Set mrsDatabase = New ADODB.Recordset
    
    ' Set values of fields
    With mrsDatabase
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        ' Change the literal below to the name of your view
        .Source = "select * from v_mnb_Category_part " & _
            "order by Category_id asc, part_id asc"
        .LockType = adLockBatchOptimistic
        .Open
    End With
    
    ' Check for errors returned from the recordset
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "RetrieveCategoryPartData", _
            gconDatabase.Errors(0).Description
    End If
    
    If cboCategory.ListCount > 0 Then
        cboCategory.ListIndex = 0
    End If
            
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveCategoryPartData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub RemoveFromAvailable(strProdCode As String)
    ' Purpose:  Removes a part Code from the available list
    '           which appears on the Selected list.
    ' Input:    strPart - The part Code to be
    '           removed.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare a variable for the index
    Dim intIndex As Integer
    
    'The v_prod_part tables are very large. This Do Loop and Upper and Lower Boundry
    'is being reset so the part being searched for only has to loop no more than 100 times.
    'This logic change causes the program to loop only 15-150 times vs an
    'average of 5000 - 10,000 times.
    'The tables contain 10,000 to 20,000 parts for each division.
    
    'Setting upper and lower boundries
    intUpperBoundry = lstAvailable.ListCount - 1
    intLowerBoundry = 0
    
    'Grouping for every 100 recs
    intSubGroups = intUpperBoundry / 100
    
    'Determine the group of 100 the Lower and Upper Boundry should be set.
    If intSubGroups > 0 Then
    Do
        If intSubGroups > lstAvailable.ListCount Then
            intUpperBoundry = lstAvailable.ListCount - 1
            Exit Do
        End If
        If strPartId > Mid(lstAvailable.List(intSubGroups), 1, Len(strPartId)) Then
            intLowerBoundry = intSubGroups
        Else
            intUpperBoundry = intSubGroups
            Exit Do
        End If
        intSubGroups = intSubGroups + 100
    Loop
    End If
    
    ' Loop through the listbox when a match is found,
    ' remove the item and leave the sub.
    For intIndex = intLowerBoundry To intUpperBoundry
        If Mid(lstAvailable.List(intIndex), 1, Len(strPartId)) = _
                strPartId Then
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

Private Sub RetrieveCategoryData()
    ' Purpose:  Instantiate and open the recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare Recordset to hold line
    Dim rsCategory As ADODB.Recordset
    Dim intIndex As Integer
    Dim strMinibillFlag As String
    
    frmProcessing.Label2 = "** Retrieve Category List ** " & Time
    frmProcessing.Refresh
    
    If chkRacman.Value = 0 Then
        strMinibillFlag = "1"
    Else
        strMinibillFlag = "0"
    End If
    
    ' Instantiate the recordset
    Set rsCategory = New ADODB.Recordset


    If optCategoryDescSequence.Value = True Then
    ' Set values of fields
        With rsCategory
            'tells the recordset where to get its data from
            'i.e. gcondatabase="Provider=SQLOLEDB.1;PASSWORD=090400;USERID=SCHIMC
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            ' Change the literal below to the name of your view
            .Source = "select category_id, category_description from v_mnb_Category " & _
                "where minibill_only_flag = " & strMinibillFlag & " order by Category_description asc"
            '.LockType = adLockBatchOptimistic
            .LockType = adLockReadOnly
            .Open
        End With
    
    End If

    If optCategoryIDSequence.Value = True Then
    ' Set values of fields
        With rsCategory
            'tells the recordset where to get its data from
            'i.e. gcondatabase="Provider=SQLOLEDB.1;PASSWORD=090400;USERID=SCHIMC
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            ' Change the literal below to the name of your view
            .Source = "select category_id, category_description from v_mnb_Category " & _
                "where minibill_only_flag = " & strMinibillFlag & " order by Category_id asc"
            '.LockType = adLockBatchOptimistic
            .LockType = adLockReadOnly
            .Open
        End With
    
    End If


    ' Check for errors returned from the recordset
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "RetrieveCategoryData", _
            gconDatabase.Errors(0).Description
    End If

    ' if no records were retrieved, add a new record to the
    ' recordset and reset fields to their original value.
    cboCategory.Clear
    If rsCategory.EOF Then
        If chkRacman.Value = 0 Then
            MsgBox ("No Categories have been set up for Minibill")
        Else
            MsgBox "No Categories have been set up for Racman"
        End If
        lstAvailable.Clear
        lstSelected.Clear
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    cmdClearAll.Enabled = False
        GoTo PROC_EXIT
    End If

    cmdAdd.Enabled = True
    cmdRemove.Enabled = True
    cmdClearAll.Enabled = True
    
    ' Go to the first record in the recordset and set the line ID
    With rsCategory
        ' Loop through the file
        Do While Not .EOF
            cboCategory.AddItem (Left(!Category_id, 5) & "      " & Trim(!Category_description)), intIndex
            ReDim Preserve arrstrCategory(intIndex)
            arrstrCategory(intIndex) = !Category_id
            .MoveNext
            intIndex = intIndex + 1
        Loop
        .Close
    End With
    
    Set rsCategory = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveCategoryData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub RetrievePartData()
    ' Purpose:  Retrieve the part ID's from the table and
    '           build the array of parts and the part Description
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare the array size
    Dim intArrayCount As Integer

    
    ' Declare the recordset variable
    Dim rsPart As ADODB.Recordset
    
    ' Instantiate the recordset
    Set rsPart = New ADODB.Recordset

    frmProcessing.Label2 = "** Retrieve Parts List ** " & Time
    frmProcessing.Refresh

    Dim strDisplay As String
    

    ' Set values of fields
    With rsPart
        'tells the recordset where to get its data from
        'i.e. gcondatabase="Provider=SQLOLEDB.1;PASSWORD=090400;USERID=SCHIMC
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        ' Change the literal below to the name of your view
        If Me.chkUnassigned.Value = 0 Then
            If optByModel Then
                .Source = "select distinct v_prod_part.part_id, part_description " & _
                    "from v_mnb_model_part join v_prod_part on " & _
                    "v_mnb_model_part.part_id = v_prod_part.part_id " & _
                    "where model_number = '" & cboModel.Text & "' " & _
                    "order by v_prod_part.part_id asc"
            Else
  '              .Source = "select * from v_prod_part " & _
  '                 "order by v_prod_part.part_id asc"
                 .Source = "select * From v_prod_part " & _
                    "Where part_last_updated > DateAdd(yyyy, -2, getdate()) " & _
                        "order by v_prod_part.part_id asc"
            End If
        Else
            If optByModel Then
                .Source = "select distinct v_prod_part.part_id, part_description " & _
                    "from v_mnb_model_part join v_prod_part on " & _
                    "v_mnb_model_part.part_id = v_prod_part.part_id " & _
                    "left outer join v_mnb_category_part on v_prod_part.part_id = " & _
                    "v_mnb_category_part.part_id " & _
                    "where model_number = '" & cboModel.Text & "' and " & _
                    "category_id is null " & _
                    "order by v_prod_part.part_id asc"
            Else
                .Source = "select v_prod_part.part_id, part_description from v_prod_part " & _
                    "left outer join v_mnb_category_part on v_prod_part.part_id = " & _
                    "v_mnb_category_part.part_id " & _
                    "where category_id is null and part_last_updated > DateAdd(yyyy, -2, getdate()) " & _
                    "order by v_prod_part.part_id asc"


            End If
        End If
        '.LockType = adLockBatchOptimistic
        .LockType = adLockReadOnly
        .Open
        strDisplay = .RecordCount
    End With

    ' Check for errors returned from the recordset
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, _
            "RetrievepartData", _
            gconDatabase.Errors(0).Description
    End If

    ' Clear the available products.
    lstAvailable.Clear
    
    frmProcessing.Label2 = "** Load Parts Array ** " & Time
    frmProcessing.Refresh
    
    ' Using the recordset, load the array and the
    ' available listbox.
    With rsPart
        ' Move to the first record.
        .MoveFirst
        
        ' Reset the size of the arrays.
        ReDim marrstrpart(.RecordCount - 1)
        ReDim marrstrpartDescription(.RecordCount - 1)
        If Not mblnLoadedAllParts Then
            ReDim marrstrAllParts(.RecordCount - 1)
            ReDim marrstrAllPartDesc(.RecordCount - 1)
        End If
        
        ' Loop through the recordset.
        Do While Not .EOF
            
            ' Add the product information to the arrays.
            marrstrpart(intArrayCount) = rsPart _
                !part_id
            marrstrpartDescription(intArrayCount) = rsPart _
                !part_description
            If Not mblnLoadedAllParts Then
                marrstrAllParts(intArrayCount) = marrstrpart(intArrayCount)
                marrstrAllPartDesc(intArrayCount) = marrstrpartDescription(intArrayCount)
            End If
            
            ' Move to the next record.
            .MoveNext
            
            ' Increment the array counter
            intArrayCount = intArrayCount + 1
            
            'Penny counts array before failure '
'            Dim intFactor As Integer
'            Dim intCompareArray As Integer
'
'            If intArrayCount > intCompareArray Then
'                intFactor = intFactor + 1
'                intCompareArray = intFactor * 100
'                MsgBox ("Array is at " & intArrayCount)
'            End If
        Loop
        .Close
    End With
    mblnLoadedAllParts = True
    Set rsPart = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrievepartData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub


Private Function ValidEntries() As Boolean
    ' Purpose:  Validate any fields which have validation
    '           routines.  Send the results back to the caller.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set up a boolean to be used to hold the cancel
    ' variable.
    Dim blnCancel As Boolean
    
    ' Initialize the variable to hold cancel.
    blnCancel = False
    
    ' If the current record has been deleted, do not perform edits.
    If mrsDatabase.EditMode = adEditDelete Then
        blnCancel = False
        GoTo PROC_EXIT
    End If
    
   
PROC_EXIT:
    ValidEntries = Not blnCancel
    Exit Function
    
PROC_ERR:
    Call ShowError(Me.Name, "ValidateEntries", Err.Number, _
        Err.Description)
    blnCancel = True
    GoTo PROC_EXIT

End Function

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
    ' Purpose:  Find the description for the part code.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set up variables.
    Dim intLoopIndex As Integer
    
    'Set upper and lower boundries
    intUpperBoundry = UBound(marrstrAllParts)
    intLowerBoundry = 0
    
    'Break parts into groups of 100
    intSubGroups = intUpperBoundry / 100
    
    'Process the array to determine the range of 100 parts that will be
    'searched.
    If intSubGroups > 0 Then
    Do
        If intSubGroups > UBound(marrstrAllParts) Then
            intUpperBoundry = UBound(marrstrAllParts)
            Exit Do
        End If
        If strPartId > marrstrAllParts(intSubGroups) Then
            intLowerBoundry = intSubGroups
        Else
            intUpperBoundry = intSubGroups
            Exit Do
        End If
        intSubGroups = intSubGroups + 100
    Loop
    End If
    
    ' Loop through the array and find the matching
    ' part number to fill the description.
    For intLoopIndex = intLowerBoundry To intUpperBoundry
        If strPartId = marrstrAllParts(intLoopIndex) Then
            strPartDescription = marrstrAllPartDesc(intLoopIndex)
            Exit For
        End If
            
    Next intLoopIndex
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FindDescription", _
        Err.Number, Err.Description)
    
End Sub

Private Sub optAll_Click()
    If Me.Visible Then
        ' Retrieve the data
        cboModel.Enabled = False
    
        Call RetrievePartData
        cboCategory_Click
    End If
End Sub

Private Sub optByModel_Click()
    cboModel.Enabled = True
End Sub


Sub RetrieveModelData()
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    
    frmProcessing.Label2 = "** Retrieve Model List ** " & Time
    frmProcessing.Refresh
    
    cboModel.Clear
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select distinct model_number from v_mnb_model_line order by model_number asc"
        .Open
        
        Do While Not .EOF
            cboModel.AddItem Trim(!model_number)
            .MoveNext
        Loop
        .Close
    End With
    Set rsList = Nothing
    
PROC_EXIT:

    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveModelData", _
        Err.Number, Err.Description)
    
End Sub

Private Sub optCategoryDescSequence_Click()
                
    Call RetrieveCategoryData
       
End Sub

Private Sub optCategoryIDSequence_Click()

    Call RetrieveCategoryData

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
