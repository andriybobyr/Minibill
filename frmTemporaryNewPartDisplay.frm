VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTemporaryNewPartDisplay 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MiniBill - Temporary Parts Change Selection..."
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
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
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   536
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid dgrdTempNewParts 
      Height          =   4455
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7858
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "part_id"
         Caption         =   "Part Number"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "part_description"
         Caption         =   "Part Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "temp_part_activity"
         Caption         =   "Activity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Quit_ecn_date"
         Caption         =   "Quit Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "M/d/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   139.994
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   60.019
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
End
Attribute VB_Name = "frmTemporaryNewPartDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
    ' Purpose:  If the user has selected an existing Part, close the
    '           form after setting the Part from the
    '           calling form.  If not, send a message that
    '           the field has not yet been selected.
    
    ' Set up error handling
    On Error GoTo PROC_EXIT
    
    Unload Me
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdOK_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
    
End Sub


Private Sub Form_Load()
    ' Purpose:  Build the list of Unit Area ID's from the recordset
    '           received from the calling form.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Fill the data grid
    With dgrdTempNewParts
        Set .DataSource = frmTemporaryNewPart.mrsTempModelPartsMaster
        gconDatabase.Errors.Clear
        .Columns(0).DataField = "Part_Id"
        .Columns(1).DataField = "Part_Description"
        .Columns(2).DataField = "Temp_Part_Activity_code"
        .Columns(3).DataField = "Quit_ECN_Date"
        .ReBind
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "TemporaryNewPartDisplay", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
    
    
End Sub

Private Sub dgrdTempNewParts_dblClick()
    ' Purpose:  Call the cmdOK_Click sub
    
    Call cmdOK_Click
End Sub
