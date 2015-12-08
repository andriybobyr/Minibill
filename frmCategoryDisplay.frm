VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCategoryDisplay 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MiniBill - Category Selection..."
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
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
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid dgrdCategory 
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         DataField       =   ""
         Caption         =   "Category"
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
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   "Description"
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
         DataField       =   ""
         Caption         =   "Minibill Category?"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   4
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   8
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   ""
         Caption         =   "Obsolete Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
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
         MarqueeStyle    =   3
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   71.017
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   165.997
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   128.013
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   111.005
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
End
Attribute VB_Name = "frmCategoryDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
    ' Purpose:  If the user has selected a Location, close the
    '           form after setting the mstrlocationID from the
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
    ' Purpose:  Build the list of Unit Location ID's from the recordset
    '           received from the calling form.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Fill the data grid
    With dgrdCategory
        Set .DataSource = frmCategory.mrsDatabase
        gconDatabase.Errors.Clear
        .Columns(0).DataField = "Category_ID"
        .Columns(1).DataField = "Category_description"
        .Columns(2).DataField = "Minibill_only_flag"
        .Columns(3).DataField = "Category_obsolete_date"
        .ReBind
    End With
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FindLocation", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
    
    
End Sub

Private Sub dgrdCategory_DblClick()
    ' Purpose:  Call the cmdOK_Click sub
    
    Call cmdOK_Click
End Sub
