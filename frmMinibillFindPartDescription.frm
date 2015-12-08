VERSION 5.00
Begin VB.Form frmMinibillFindPartDescription 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MiniBill - Find Part Description..."
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
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
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find &Next"
      Default         =   -1  'True
      Height          =   600
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   555
      Left            =   3900
      TabIndex        =   2
      Top             =   1320
      Width           =   1275
   End
   Begin VB.TextBox txtPart 
      Height          =   375
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   3795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Part Description:"
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
      Left            =   165
      TabIndex        =   0
      Top             =   720
      Width           =   1860
   End
End
Attribute VB_Name = "frmMinibillFindPartDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngSaveIndex As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim lngIndex As Long
    Dim strDescription As String * 50
        
'Verify a part number was entered.
    If Len(Trim(txtPart.Text)) = 0 Then
        MsgBox "Please enter the Part Description to find or click Cancel"
        txtPart.SetFocus
        Exit Sub
    End If
    
    
'Verify the part is found in this model
    With frmMiniBillMaintenance
'        If Len(txtPart.Text) < 50 Then
'            txtPart.Text = txtPart.Text & Space(50 - Len(txtPart.Text))
'        End If
        strDescription = txtPart.Text
        lngIndex = .mxarrMiniBill.Find(1, 4, strDescription, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
        'This test was added due to the find not always finding the first part in the list
        If lngIndex < 0 Then
            lngIndex = .mxarrMiniBill.Find(1, 4, txtPart.Text, XORDER_DESCEND, XCOMP_EQ, XTYPE_STRING)
        End If
        If lngIndex < 0 Then
            MsgBox "Part Description is not found"
            txtPart.SetFocus
            Exit Sub
        End If
        
        lngSaveIndex = lngIndex
        
'Once the part is found, compare where new part is compared to the first row
'  on the grid
        .TDBGMiniBill.Scroll 0, lngIndex - .TDBGMiniBill.FirstRow
 
'Clear the grid to refresh the data
        .TDBGMiniBill.Row = 0
    End With
    
'    Unload Me
End Sub
Private Sub cmdFindNext_Click()
    Dim lngIndex As Long
    Dim strDescription As String * 50
    'Verify a part description was entered.
    If Len(Trim(txtPart.Text)) = 0 Then
        MsgBox "Please enter the Part Description to find or click Cancel"
        txtPart.SetFocus
        Exit Sub
    End If
    
'Verify the part is found in this model
    With frmMiniBillMaintenance
'        If Len(txtPart.Text) < 50 Then
'            txtPart.Text = txtPart.Text & Space(50 - Len(txtPart.Text))
'        End If
        strDescription = txtPart.Text
        lngIndex = .mxarrMiniBill.Find(lngSaveIndex + 1, 4, strDescription, XORDER_ASCEND, XCOMP_EQ, XTYPE_STRING)
        lngSaveIndex = lngIndex
        'This test was added due to the find not always finding the first part in the list
        If lngIndex < 0 Then
            lngIndex = .mxarrMiniBill.Find(1, 4, strDescription, XORDER_DESCEND, XCOMP_EQ, XTYPE_STRING)
        End If
        If lngIndex < 0 Then
            MsgBox "Part Description is not found"
            txtPart.SetFocus
            Exit Sub
        End If
      
'Once the part is found, compare where new part is compared to the first row
'  on the grid
        .TDBGMiniBill.Scroll 0, lngIndex - .TDBGMiniBill.FirstRow
 
'Clear the grid to refresh the data
        .TDBGMiniBill.Row = 0
    End With
    
'    Unload Me
End Sub

Private Sub txtPart_GotFocus()
    txtPart.SelStart = 0
    txtPart.SelLength = Len(txtPart.Text)
End Sub

Private Sub txtPart_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


