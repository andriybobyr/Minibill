VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form dlgCalendar 
   BackColor       =   &H00808000&
   Caption         =   "Date Selection..."
   ClientHeight    =   4020
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5370
   StartUpPosition =   1  'CenterOwner
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   4020
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   5370
      _cx             =   9472
      _cy             =   7091
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      BackColor       =   8421376
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
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
      ResizeFonts     =   -1  'True
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   540
         TabIndex        =   1
         Top             =   3540
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   3540
         Width           =   1215
      End
      Begin MSACAL.Calendar ctlCalendar 
         Height          =   3075
         Left            =   300
         TabIndex        =   0
         Top             =   300
         Width           =   4755
         _Version        =   524288
         _ExtentX        =   8387
         _ExtentY        =   5424
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2000
         Month           =   10
         Day             =   21
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "dlgCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This variable contains the date field being passed bsck
' and forth from the calling program.
Public mdteSelectedDate As Date

' This vaiable conains the original date with the form
' was called.  It is used to return the same date to the
' program with the cancel button is clicked.
Private mdteOriginalDate As Date

Private Sub cmdCancel_Click()
    ' Purpose:  When the cancel button is clicked,
    ' reset the passed parameter to its original value
    ' and hide the form.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set the passed date back to the original value
    mdteSelectedDate = mdteOriginalDate
    
    ' Hide the form.  This will return control to the
    ' calling program.
    Me.Hide
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdCancel_click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdOK_Click()
    ' Purpose:  Save the date selected on the calendar to
    ' the date being passed back to the calling form and
    ' hide the form.  This will return control back to the
    ' calling program.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Save the selected date to the public variable used
    ' by the calling program
    mdteSelectedDate = ctlCalendar.Value
    
    ' Hide the form.
    Me.Hide
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "OKButton_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub Form_Activate()
    ' Purpose:  When the form is activated, set the
    ' selected date on the calendar to the date being
    ' passed from the calling program.  Save the original
    ' value.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ctlCalendar.FirstDay = 1
    ctlCalendar.BackColor = RGB(208, 208, 208)
    
    ' if the date being passed is null, set the selected
    ' date to today.
    If Not IsNull(mdteSelectedDate) Then
        ctlCalendar.Value = mdteSelectedDate
        
    ' Otherwise set the selected date to the date that
    ' was passed.
    Else
        ctlCalendar.Value = Date
    End If
    
    ' Save the current date value to the original date
    ' field.
    mdteOriginalDate = ctlCalendar.Value
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Activate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub
