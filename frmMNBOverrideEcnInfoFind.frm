VERSION 5.00
Begin VB.Form frmMNBOverrideEcnInfoFind 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MiniBill - Find Override ECN Info..."
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtModelNumber 
      Height          =   375
      Left            =   1845
      MaxLength       =   20
      TabIndex        =   2
      Top             =   720
      Width           =   2475
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   555
      Left            =   3060
      TabIndex        =   1
      Top             =   1800
      Width           =   1275
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   555
      Left            =   780
      TabIndex        =   0
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model Number:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   780
      Width           =   1590
   End
End
Attribute VB_Name = "frmMNBOverrideEcnInfoFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    gblnFindEcnInfoCancel = True
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim varBookmark As Variant
    
    If Len(Trim(txtModelNumber.Text)) = 0 Then
        MsgBox "Please enter the Model Number to find or click Cancel"
        txtModelNumber.SetFocus
        Exit Sub
    End If
    
    If frmMNBOverrideECNInfo.mrsDatabase Is Nothing Then
    
        ' Instantiate the recordset
        
        Set frmMNBOverrideECNInfo.mrsDatabase = New ADODB.Recordset
    
        ' Set values of fields
        
        With frmMNBOverrideECNInfo.mrsDatabase
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            ' Change the literal below to the name of your view
            .Source = "select * from v_mnb_model_Part where model_number = '" & _
                txtModelNumber.Text & "'"
            .LockType = adLockBatchOptimistic
            .Open
        End With
        
        If frmMNBOverrideECNInfo.mrsDatabase.RecordCount = 0 Then
            MsgBox "Model Number not found"
            txtModelNumber.SetFocus
            Set frmMNBOverrideECNInfo.mrsDatabase = Nothing
            Exit Sub
        End If
   
        ' Check for errors returned from the recordset
        If gconDatabase.Errors.Count > 0 Then
            Err.Raise vbObjectError + 1000, "RetrieveData", _
            gconDatabase.Errors(0).Description
        End If
        
    End If
      
' end of addition

    With frmMNBOverrideECNInfo.mrsDatabase
        varBookmark = .Bookmark
        .MoveFirst
        .Find "Model_Number = '" & _
            txtModelNumber.Text & "'"
        If .EOF Then
            MsgBox "Model was not found"
            txtModelNumber.SetFocus
            .Bookmark = varBookmark
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub txtModelNumber_GotFocus()
    txtModelNumber.SelStart = 0
    txtModelNumber.SelLength = Len(txtModelNumber.Text)
End Sub

Private Sub txtModelNumber_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub




