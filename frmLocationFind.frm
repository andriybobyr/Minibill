VERSION 5.00
Begin VB.Form frmLocationFind 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MiniBill - Find Stocking Location..."
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   ScaleWidth      =   312
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
   Begin VB.TextBox txtLocationID 
      Height          =   375
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   1
      Top             =   780
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location ID:"
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
      Left            =   540
      TabIndex        =   0
      Top             =   900
      Width           =   1245
   End
End
Attribute VB_Name = "frmLocationFind"
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
    
    If Len(Trim(txtLocationID.Text)) = 0 Then
        MsgBox "Please enter the Location ID to find or click Cancel"
        txtLocationID.SetFocus
        Exit Sub
    End If
    
    With frmLocation.mrsDatabase
        varBookmark = .Bookmark
        .MoveFirst
        .Find "stocking_Location_ID = '" & _
            txtLocationID.Text & "'"
        If .EOF Then
            MsgBox "Location ID was not found"
            txtLocationID.SetFocus
            .Bookmark = varBookmark
            Exit Sub
        End If
    End With
    
    Unload Me
End Sub

Private Sub txtLocationID_GotFocus()
    txtLocationID.SelStart = 0
    txtLocationID.SelLength = Len(txtLocationID.Text)
End Sub

Private Sub txtLocationID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtColorCode_Change()

End Sub

