VERSION 5.00
Begin VB.Form frmCopies 
   Caption         =   "Print Copies"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   5340
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   2820
      TabIndex        =   3
      Top             =   2160
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   465
      Left            =   1275
      TabIndex        =   2
      Top             =   2160
      Width           =   1275
   End
   Begin VB.TextBox txtCopies 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4140
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "1"
      Top             =   1035
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter Number of Print Copies:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      TabIndex        =   1
      Top             =   1065
      Width           =   3405
   End
End
Attribute VB_Name = "frmCopies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
'Add any report form here to reflect the number of copies.  If not, number of copies will
'  always be zero.
    
 
    If frmDailyScheduleSheet.mintNumberOfCopies = CInt(txtCopies.Text) Then
        frmDailyScheduleSheet.mblnCancelPrint = False
    End If
    
    If frmDailyScheduleSheetAllModels.mintNumberOfCopies = CInt(txtCopies.Text) Then
        frmDailyScheduleSheetAllModels.mblnCancelPrint = False
    End If
    
    If frmTempPartsList.mintNumberOfCopies = CInt(txtCopies.Text) Then
        frmTempPartsList.mblnCancelPrint = False
    End If
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    frmDailyScheduleSheet.mblnCancelPrint = True
    frmDailyScheduleSheetAllModels.mblnCancelPrint = True
    frmTempPartsList.mblnCancelPrint = True
    
    Unload Me
End Sub
Private Sub Form_Load()

    If CStr(frmDailyScheduleSheet.mintNumberOfCopies) = ActiveForm Then
        txtCopies.Text = CStr(frmDailyScheduleSheet.mintNumberOfCopies)
    End If

    If CStr(frmDailyScheduleSheetAllModels.mintNumberOfCopies) = ActiveForm Then
        txtCopies.Text = CStr(frmDailyScheduleSheetAllModels.mintNumberOfCopies)
    End If
  
        If CStr(frmTempPartsList.mintNumberOfCopies) = ActiveForm Then
            txtCopies.Text = CStr(frmTempPartsList.mintNumberOfCopies)
        End If
   
    txtCopies.SelStart = 0
    txtCopies.SelLength = Len(txtCopies.Text)

End Sub

Private Sub txtCopies_Validate(Cancel As Boolean)

    ' If the length of the field is zero, give an error
    If Len(txtCopies.Text) = 0 Then
        Cancel = True
        MsgBox "Number of Print Copies is required", _
            vbExclamation + vbOKOnly, _
            "Print Copies Validation"
        GoTo PROC_EXIT
    Else
        If Not IsNumeric(txtCopies.Text) Then
            Cancel = True
            MsgBox "Number of Print Copies must be numeric", _
                vbExclamation + vbOKOnly, _
                "Print Copies Validation"
            GoTo PROC_EXIT
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "txtCopies_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub
