VERSION 5.00
Begin VB.Form frmLine 
   Caption         =   "Find Parts from Line..."
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   ScaleHeight     =   2490
   ScaleWidth      =   4395
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboLine 
      Height          =   315
      Left            =   1980
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblHeading 
      Caption         =   "Pick Line to update parts:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   3255
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
      Left            =   1320
      TabIndex        =   1
      Top             =   1230
      Width           =   615
   End
End
Attribute VB_Name = "frmLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
