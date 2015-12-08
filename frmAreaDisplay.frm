VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAreaDisplay 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MiniBill - Area Selection..."
   ClientHeight    =   5580
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
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid dgrdArea 
      Height          =   4275
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7541
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "area_id"
         Caption         =   "Area"
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
         DataField       =   "area_description"
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
         DataField       =   "area_obsolete_date"
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
            ColumnWidth     =   98.986
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   165.997
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   111.005
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   4500
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
End
Attribute VB_Name = "frmAreaDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
    ' Purpose:  If the user has selected a Area, close the
    '           form after setting the mstrAreaID from the
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
    With dgrdArea
        Set .DataSource = frmArea.mrsDatabase
        gconDatabase.Errors.Clear
        .Columns(0).DataField = "stocking_area_id"
        .Columns(1).DataField = "stocking_area_description"
        .Columns(2).DataField = "stocking_area_obsolete_date"
        .ReBind
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FindArea", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
    
    
End Sub

Private Sub dgrdArea_dblClick()
    ' Purpose:  Call the cmdOK_Click sub
    
    Call cmdOK_Click
End Sub
