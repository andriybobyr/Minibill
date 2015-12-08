VERSION 5.00
Begin VB.Form frmMNBModelLineStockLocNotesFind 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MiniBill - Find Model"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
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
   ScaleWidth      =   364
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   555
      Left            =   2820
      TabIndex        =   3
      Top             =   1800
      Width           =   1275
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   495
      Left            =   660
      TabIndex        =   2
      Top             =   1860
      Width           =   1395
   End
   Begin VB.TextBox txtModelNumber 
      Height          =   375
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   1
      Top             =   780
      Width           =   2475
   End
   Begin VB.Label Label1 
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
      Left            =   195
      TabIndex        =   0
      Top             =   900
      Width           =   1590
   End
End
Attribute VB_Name = "frmMNBModelLineStockLocNotesFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim varBookmark As Variant
    
    If Len(Trim(txtModelNumber.Text)) = 0 Then
        MsgBox "Please enter the Model Number to find or click Cancel"
        txtModelNumber.SetFocus
        Exit Sub
    End If
    
    With frmMNBModelLineStockLocNotes.mrsDatabase
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


