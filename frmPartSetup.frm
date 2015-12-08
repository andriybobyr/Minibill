VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmPartSetup 
   Caption         =   "MiniBill - Part / Stocking Location Specification"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11235
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
   ScaleHeight     =   544
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   749
   StartUpPosition =   1  'CenterOwner
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   8160
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   11235
      _cx             =   19817
      _cy             =   14393
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      BackColor       =   12640511
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   0
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
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      Begin VB.CheckBox chkUnassigned 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Display Only Parts With No Location Assignment"
         Height          =   435
         Left            =   3360
         TabIndex        =   23
         Top             =   2220
         Width           =   4815
      End
      Begin VB.CommandButton cmdSplitPart 
         Caption         =   "&Split Part..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6420
         TabIndex        =   21
         Top             =   4680
         Width           =   1035
      End
      Begin VB.ComboBox cboLine 
         Height          =   360
         Left            =   2460
         TabIndex        =   0
         Top             =   660
         Width           =   855
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   435
         Left            =   7920
         TabIndex        =   19
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox txtPart 
         Height          =   360
         Left            =   420
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2220
         Width           =   2850
      End
      Begin VB.Frame fraSelection 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Assigning defaults for:"
         Height          =   855
         Left            =   600
         TabIndex        =   18
         Top             =   1020
         Width           =   8775
         Begin VB.ComboBox cboModel 
            Height          =   360
            Left            =   4200
            TabIndex        =   4
            Top             =   240
            Width           =   2715
         End
         Begin VB.OptionButton optByModel 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Only Model Listed"
            Height          =   375
            Left            =   2280
            TabIndex        =   3
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optAll 
            BackColor       =   &H00C0E0FF&
            Caption         =   "All Models"
            Height          =   255
            Left            =   300
            TabIndex        =   2
            Top             =   300
            Width           =   1995
         End
      End
      Begin VB.CommandButton cmdClose 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   780
         Picture         =   "frmPartSetup.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Close This Form"
         Top             =   60
         Width           =   300
      End
      Begin VB.ComboBox cboLocation 
         Height          =   360
         Left            =   5640
         TabIndex        =   1
         Top             =   660
         Width           =   3015
      End
      Begin VB.ListBox lstSelected 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   7680
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2760
         Width           =   3360
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add >"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6420
         TabIndex        =   12
         Top             =   2730
         Width           =   1035
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "< &Remove"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6420
         TabIndex        =   11
         Top             =   3330
         Width           =   1035
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "<< &Clear All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6420
         TabIndex        =   10
         Top             =   4005
         Width           =   1035
      End
      Begin VB.CommandButton cmdSave 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   60
         Picture         =   "frmPartSetup.frx":0762
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Save Changes"
         Top             =   60
         Width           =   300
      End
      Begin VB.CommandButton cmdHelp 
         CausesValidation=   0   'False
         Height          =   300
         Left            =   420
         Picture         =   "frmPartSetup.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Help"
         Top             =   60
         Width           =   300
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGAvailable 
         Height          =   5055
         Left            =   420
         TabIndex        =   22
         Top             =   2760
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   8916
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Part"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Description"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Location"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Lvl"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Qty"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2037"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1958"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3598"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=3519"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=2514"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2434"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=741"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=661"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=741"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=661"
         Splits(0)._ColumnProps(27)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(28)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(29)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         TabAction       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=3,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
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
         Index           =   3
         Left            =   720
         TabIndex        =   20
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Location:"
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
         Left            =   3900
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Available Parts"
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
         Index           =   1
         Left            =   420
         TabIndex        =   16
         Top             =   1860
         Width           =   3195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Parts"
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
         Index           =   2
         Left            =   8280
         TabIndex        =   15
         Top             =   2340
         Width           =   2595
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   1
         X1              =   60
         X2              =   11040
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808000&
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         Index           =   0
         X1              =   60
         X2              =   11040
         Y1              =   495
         Y2              =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Whirlpool MES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   8580
         TabIndex        =   14
         Top             =   120
         Width           =   2475
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "&Index..."
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmPartSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsModelPartLocation As ADODB.Recordset
Private mrsPartLocation As ADODB.Recordset

Private mblnRecChanged As Boolean
Private strPartId As String
Private strPartDescription As String
Private mxarrPartInfo As XArrayDB
Private marrstrAllParts() As String
Private marrstrAllPartDesc() As String
Private mblnLoadedAllParts As Boolean
Private mxarrAvailable As XArrayDB

Private mxarrLineLocationInfo As XArrayDB

Private mblnLoadedParts As Boolean
Private mblnSaved As Boolean

Private mstrLocationSave As String
Private mstrLineSave As String
Private mstrModelSave As String
Private mblnAllSave As Boolean
Private mblnUnassignedSave As Boolean



Private Sub cboLine_Change()
    If Len(cboLine.Text) = 2 Then
        cboFindFirst cboLine
    End If
End Sub

Private Sub cboLine_Click()
    ' Purpose:  Fill the location listbox based on the line selected.
    
    If mstrLineSave = cboLine.Text Then
        Exit Sub
    End If
    
    If mblnRecChanged Then
        SaveData
        If Not mblnSaved Then
            Exit Sub
        End If
    End If
    
    mstrLineSave = cboLine.Text
    mstrLocationSave = vbNullString
    Dim lngRow As Long
    cboLocation.Clear
    lngRow = mxarrLineLocationInfo.Find(0, 2, cboLine.Text)
    Do While lngRow <= mxarrLineLocationInfo.UpperBound(1)
        If mxarrLineLocationInfo(lngRow, 2) = cboLine.Text Then
            cboLocation.AddItem mxarrLineLocationInfo(lngRow, 1)
            cboLocation.ItemData(cboLocation.NewIndex) = lngRow
            lngRow = lngRow + 1
        Else
            lngRow = mxarrLineLocationInfo.UpperBound(1) + 1
        End If
    Loop
    
    RetrieveModelData
    If cboModel.ListCount = 0 And Screen.ActiveForm Is Me Then
        MsgBox "No models found for this line."
        Exit Sub
    End If
    
    If cboModel.ListCount = 0 Then
        optAll.Value = 1
        optByModel.Enabled = False
    Else
        optByModel.Enabled = True
    End If
    
    If Not optAll.Value Then
        optAll.Value = True
    Else
        RetrievePartData
    End If
    cboLocation.ListIndex = 0
    cboLocation_Click
End Sub

Private Sub cboLine_Validate(Cancel As Boolean)
    ' Purpose:  Validate the line entered
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' If there is no data in the field, display a message,
    ' set cancel to true and exit.
    If Len(Trim(cboLine.Text)) = 0 Then
        MsgBox "Line is required."
        Cancel = True
        GoTo PROC_EXIT
    
    Else
        ' If the length of the field is one, add a
        ' leading zero.
        If Len(cboLine.Text) = 1 Then
            cboLine.Text = "0" & cboLine.Text
        End If
        
        ' If the line id has changed, look up the new
        ' line in the listbox.
        cboFindFirst cboLine
            
        ' If the line was not found in the list,
        ' display a message, set cancel to true and
        ' exit.
        If cboLine.ListIndex = -1 Then
            MsgBox "Line " & cboLine.Text & " is not valid"
            Cancel = True
            GoTo PROC_EXIT
        End If
        
        Cancel = False
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboLineID_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cboLocation_Change()
    cboFindFirst cboLocation
End Sub

Private Sub cboLocation_Click()
    ' Put code in here to rebuild selected based on the Locaiton selected.
    
    Dim lngRow As Long
    Dim lngNewRow As Long
    Dim intCol As Integer
    
    If mstrLocationSave = cboLocation.Text Then
        Exit Sub
    End If
    
    If mblnRecChanged Then
        SaveData
        If Not mblnSaved Then
            Exit Sub
        End If
    End If
    
    mstrLocationSave = cboLocation.Text
    
    Set mxarrAvailable = New XArrayDB
    lstSelected.Clear
    
    mxarrAvailable.ReDim 0, mxarrPartInfo.UpperBound(1), 0, mxarrPartInfo.UpperBound(2) + 1
    lngNewRow = -1
    For lngRow = 0 To mxarrPartInfo.UpperBound(1)
        If Trim(mxarrPartInfo(lngRow, 2)) = Trim(cboLocation.Text) Then
            lstSelected.AddItem mxarrPartInfo(lngRow, 0) & " " & mxarrPartInfo(lngRow, 1)
            lstSelected.ItemData(lstSelected.NewIndex) = lngRow
        ElseIf chkUnassigned.Value = 0 Or mxarrPartInfo(lngRow, 2) = vbNullString Or IsNull(mxarrPartInfo(lngRow, 2)) Then
            lngNewRow = lngNewRow + 1
            If lngNewRow > 0 Then
                mxarrAvailable.AppendRows
            End If
            For intCol = 0 To mxarrPartInfo.UpperBound(2)
               mxarrAvailable.Set lngNewRow, intCol, mxarrPartInfo(lngRow, intCol)
            Next intCol
            mxarrAvailable.Set lngNewRow, intCol, lngRow
        End If
    Next lngRow
    
    TDBGAvailable.Array = mxarrAvailable
    TDBGAvailable.ReBind
    If optAll.Value Then
        TDBGAvailable.Columns(3).Visible = False
        TDBGAvailable.Columns(4).Visible = False
    Else
        TDBGAvailable.Columns(3).Visible = True
        TDBGAvailable.Columns(4).Visible = True
    End If
    mblnRecChanged = False
    txtPart.Text = vbNullString
End Sub

Private Sub cboLocation_GotFocus()
    cboLocation.SelStart = 0
    cboLocation.SelLength = Len(cboLocation.Text)
End Sub

Private Sub cboLocation_KeyPress(KeyAscii As Integer)
    cboKeyPress cboLocation, KeyAscii
End Sub

Private Sub cboLocation_Validate(Cancel As Boolean)
    ' Purpose:  Check validity of locaiton field.
    
    If Len(Trim(cboLocation.Text)) = 0 Then
        MsgBox "Location is required."
        Cancel = True
    Else
        cboFindFirst cboLocation
        If cboLocation.ListIndex = -1 Then
            MsgBox "Location is invalid."
            Cancel = True
        Else
            Cancel = False
        End If
    End If
End Sub


Private Sub cboModel_Change()
    cboFindFirst cboModel
End Sub

Private Sub cboModel_GotFocus()
    cboModel.SelStart = 0
    cboModel.SelLength = Len(cboModel.Text)
End Sub

Private Sub cboModel_KeyPress(KeyAscii As Integer)
    cboKeyPress cboModel, KeyAscii
End Sub

Private Sub cboModel_Validate(Cancel As Boolean)
    ' Purpose:  Validate model
    
    If Len(Trim(cboModel.Text)) = 0 Then
        MsgBox "Model is required."
        Cancel = True
    Else
        cboFindFirst cboModel
        If cboModel.ListIndex = -1 Then
            MsgBox "Invalid model selected."
            Cancel = True
        Else
            Cancel = False
        End If
    End If
End Sub

Private Sub chkUnassigned_Click()
    Dim lngRow As Long
    Dim lngNewRow As Long
    Dim intCol As Integer
    
    lngNewRow = -1
    mxarrAvailable.Clear
    mxarrAvailable.ReDim 0, mxarrPartInfo.UpperBound(1), 0, mxarrPartInfo.UpperBound(2) + 1
    lngNewRow = -1
    For lngRow = 0 To mxarrPartInfo.UpperBound(1)
        If chkUnassigned.Value = 0 Or IsNull(mxarrPartInfo(lngRow, 2)) Then
            lngNewRow = lngNewRow + 1
            For intCol = 0 To mxarrPartInfo.UpperBound(2)
               mxarrAvailable.Set lngNewRow, intCol, mxarrPartInfo(lngRow, intCol)
            Next intCol
            mxarrAvailable.Set lngNewRow, intCol, lngRow
        End If
    Next lngRow
    
    TDBGAvailable.Array = mxarrAvailable
    TDBGAvailable.ReBind
    txtPart.Text = vbNullString
End Sub

Private Sub cmdClose_Click()
    ' Purpose:  Close the form by the user's request.
    
    Call mnuFileClose_Click
End Sub



Private Sub cmdHelp_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the HelpContents menu item.
    
    Call mnuHelpContents_Click
End Sub


Private Sub cmdRefresh_Click()
    ' Retrieve the data
    
    If mblnRecChanged Then
        SaveData
        If Not mblnSaved Then
            Exit Sub
        End If
    End If
    
    mstrModelSave = cboModel.Text
        
    If optAll.Value Then
        mblnAllSave = True
        Call RetrievePartData
        mstrLocationSave = vbNullString
        Call cboLocation_Click
        cmdSplitPart.Enabled = False
        Exit Sub
    End If
    
    If Len(cboModel.Text) = 0 Then
        MsgBox "Please select model"
        Exit Sub
    End If
    
    cboFindFirst cboModel
    If cboModel.ListIndex = -1 Then
        MsgBox "Please select a valid model"
        Exit Sub
    End If
    
    Call RetrievePartData
    mstrLocationSave = vbNullString
    Call cboLocation_Click
    mblnAllSave = False
    cmdSplitPart.Enabled = True
End Sub

Private Sub cmdSave_Click()
    ' Purpose:  Tie the click of this button to the selection
    '           of the FileSave menu item.
    
    Call mnuFileSave_Click
End Sub

Private Sub cmdAdd_Click()
    ' Purpose:  Add the selected record in the available
    '           column to the selected list and to the
    '           recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim lngRow As Long
    Dim blnMatch As Boolean
    
    lngRow = mxarrAvailable(TDBGAvailable.Bookmark, mxarrAvailable.UpperBound(2))
    If mblnAllSave Then
        With mrsPartLocation
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "part_id = '" & TDBGAvailable.Columns(0).Value & "'"
                Do While Not .EOF
                    If !line_id <> cboLine.Text Then
                        .MoveNext
                        .Find "part_id = '" & TDBGAvailable.Columns(0).Value & "'"
                    Else
                        Exit Do
                    End If
                Loop
            End If
            If .EOF Then
                .AddNew
                !line_id = cboLine.Text
                !part_id = TDBGAvailable.Columns(0).Value
            End If
            !stocking_location_id = mxarrLineLocationInfo(cboLocation.ItemData(cboLocation.ListIndex), 0)
        End With
    Else
        With mrsModelPartLocation
            blnMatch = False
            If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF And Not blnMatch
                .Find "part_id = '" & TDBGAvailable.Columns(0).Value & "'"
                If .EOF Then
                    Exit Do
                End If
                If !part_sequence_number = mxarrPartInfo(lngRow, 5) Then
                    blnMatch = True
                Else
                    .MoveNext
                End If
            Loop
            End If
            If .RecordCount = 0 Or .EOF Then
                .AddNew
                !line_id = mstrLineSave
                !model_number = mstrModelSave
                !part_id = TDBGAvailable.Columns(0).Value
                !part_sequence_number = mxarrPartInfo(lngRow, 5)
            End If
            !stocking_location_id = mxarrLineLocationInfo(cboLocation.ItemData(cboLocation.ListIndex), 0)
        End With
    End If
    
    lstSelected.AddItem TDBGAvailable.Columns(0).Value & " " & TDBGAvailable.Columns(1).Value
    lstSelected.ItemData(lstSelected.NewIndex) = lngRow
    mxarrPartInfo(lngRow, 2) = cboLocation.Text
    TDBGAvailable.Delete
    mblnRecChanged = True

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdAdd_Click", Err.Number, _
        Err.Description)
    Unload Me

End Sub

Private Sub cmdRemove_Click()
    ' Purpose:  Remove a part ID from the selected list
    '           and delete the record from the recordset
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim lngRow As Long
    Dim lngNewRow As Long
    Dim intCol As Integer
    Dim blnMatch As Boolean
    
    If lstSelected.ListIndex = -1 Then
        MsgBox "Please select a part to be removed before clicking the remove button."
        GoTo PROC_EXIT
    End If
    
    lngRow = lstSelected.ItemData(lstSelected.ListIndex)
    If mblnAllSave Then
        With mrsPartLocation
            .MoveFirst
            .Find "part_id = '" & mxarrPartInfo(lstSelected.ItemData(lstSelected.ListIndex), 0) & "'"
            .Delete
        End With
    Else
        With mrsModelPartLocation
            blnMatch = False
            .MoveFirst
            Do While Not .EOF And Not blnMatch
                .Find "part_id = '" & mxarrPartInfo(lstSelected.ItemData(lstSelected.ListIndex), 0) & "'"
                If .EOF Then
                    Exit Do
                End If
                If !part_sequence_number = mxarrPartInfo(lngRow, 5) Then
                    blnMatch = True
                Else
                    .MoveNext
                End If
            Loop
            .Delete
        End With
    End If
    
    lstSelected.RemoveItem lstSelected.ListIndex
    lngNewRow = mxarrAvailable.AppendRows
    For intCol = 0 To mxarrPartInfo.UpperBound(2)
        mxarrAvailable(lngNewRow, intCol) = mxarrPartInfo(lngRow, intCol)
    Next intCol
    mxarrPartInfo(lngRow, 2) = Null
    TDBGAvailable.Refresh
    
    mblnRecChanged = True
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdRemove_Click", Err.Number, _
        Err.Description)
    Unload Me

End Sub

Private Sub cmdClearAll_Click()
    ' Purpose:  Remove all products from the selected
    '           list and from the recordset.
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    Dim intIndex As Integer
    
    For intIndex = lstSelected.ListCount - 1 To 0 Step -1
        lstSelected.ListIndex = intIndex
        cmdRemove_Click
    Next intIndex
    
PROC_EXIT:
    
    Exit Sub
    
PROC_ERR:
    
    Call ShowError(Me.Name, "cmdClearAll_Click", Err.Number, _
        Err.Description)
    Unload Me
    
    
End Sub

Private Sub Form_Load()
    ' Purpose:  Show the form and login to the server
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    frmProcessing.Show
    DoEvents
    
    mstrLocationSave = vbNullString
    mblnLoadedParts = False
    
    If gconDatabase Is Nothing Then
        Set gconDatabase = gclsSQLServer.Connect( _
            gclsMESApplication.ApplicationRole, _
            gclsMESApplication.ApplicationPassword)
    
        If gconDatabase.State <> adStateOpen Then
            Err.Raise gconDatabase.Errors(0).NativeError, _
                "Form_Load", gconDatabase.Errors(0).Description
        End If
    End If
    
    ' Disable fields if update is not allowed.
    If Not gblnUpdate Then
        cmdAdd.Enabled = False
        cmdRemove.Enabled = False
        cmdClearAll.Enabled = False
        cmdSplitPart.Enabled = False
    End If
    
    
    ' Retrieve the data
    optAll.Value = True
    mblnAllSave = True
    cmdSplitPart.Enabled = False
    
    
    Call RetrieveLineAndLocationData
    Call RetrievePartData
    
    mblnLoadedParts = True
    
    cboLocation.ListIndex = 0
    
    cboModel.Enabled = False
    frmProcessing.Hide
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, _
        Err.Description)
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Purpose:  Ask the user if he would like to save changes
    '           before closing the form.  If "Yes", call the
    '           mnuFileSave_Click procedure.  If "Cancel",
    '           set the Cancel flag to true and exit.  If no
    '           exit without setting the Cancel flag to
    '           true.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare Return Code variable
    Dim intRetCode As Integer
    
    ' Check to see if any changes have been made to the
    ' recordset.
    If mblnRecChanged Then
        SaveData
        If mblnSaved Then
            Cancel = False
        Else
            Cancel = True
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_QueryUnload", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Purpose:  Close and de-reference objects used by this
    '           form
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    
    ' Check to see if the recordset exists.  If so, check to see
    ' if it is open.  If it is open, close it.  De-reference it
    ' by setting it to nothing.
    If gconDatabase.State = adStateOpen Then
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Resume Next

End Sub

Private Sub lstSelected_Click()
    ' Purpose:  When an item in the slected listbox was
    '           clicked, fill the data for that record
    '           and make the fields visible.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Set a field for the product code.
    Dim strPartId As String
    
    
    ' Place the slected product code in the field
    strPartId = Mid(lstSelected.Text, 1, 15)
   
    ' Find the Category code in the recordset and
    ' fill the fields with data.
   
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "lstSelected_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub lstSelected_DblClick()
    If lstSelected.ListIndex > -1 Then
        cmdRemove_Click
    End If
End Sub

Private Sub mnuFileClose_Click()
    '  Purpose: Close this form
    
    
    ' Set up error handling
    On Error GoTo PROC_ERR
   
    Unload Me
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuClose_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub



Private Sub mnuFileExit_Click()
    ' Purpose:  Exit the application
    Unload frmMain
End Sub

Private Sub mnuFileSave_Click()
    ' Purpose:  Save the current changes to the database
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    mblnSaved = True
    If mblnAllSave Then
        mrsPartLocation.UpdateBatch
    Else
        mrsModelPartLocation.UpdateBatch
    End If
    If gconDatabase.Errors.Count > 0 Then
        Err.Raise gconDatabase.Errors(0).NativeError, "SaveData", _
            gconDatabase.Errors(0).Description
    End If
    mblnSaved = True
    mblnRecChanged = False
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuFileSave_Click", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuHelpAbout_Click()
    ' Purpose:  Display the About form for the application
    
    ' Declare and instantiate Splash screen object.
    Dim clsSplash As MES.Splash
    Set clsSplash = New MES.Splash
    
    clsSplash.Show
    
End Sub

Private Sub mnuHelpContents_Click()
    ' Purpose:  Display Help Contents
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ShowHelpContents
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuHelpContents_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub mnuHelpIndex_Click()
    ' Purpose:  Display Help Index
    
    ' Set error handling
    On Error GoTo PROC_ERR
    
    ShowHelpIndex
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "mnuHelpContents_Click", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub RetrieveLineAndLocationData()
    ' Purpose:  Instantiate and open the recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare Recordset to hold line
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    
    cboLine.Clear
    cboLocation.Clear
    mstrLineSave = vbNullString
    
    ' Instantiate the recordset
    Set rsList = New ADODB.Recordset

    ' Set values of fields
    With rsList
        'tells the recordset where to get its data from
        'i.e. gcondatabase="Provider=SQLOLEDB.1;PASSWORD=090400;USERID=SCHIMC
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        ' Change the literal below to the name of your view
        .Source = "select line_id from v_prod_line_stocking_location " & _
            "group by line_id " & _
            "order by line_id"
        '.LockType = adLockBatchOptimistic
        .LockType = adLockReadOnly
        .Open
    
        ' Check for errors returned from the recordset
        If gconDatabase.Errors.Count > 0 Then
            Err.Raise gconDatabase.Errors(0).NativeError, _
                "RetrieveLineAndLocationData", _
                gconDatabase.Errors(0).Description
        End If
    
        ' if no records were retrieved, add a new record to the
        ' recordset and reset fields to their original value.
        If .EOF Then
            MsgBox ("No records were retrieved from Line table")
            GoTo PROC_EXIT
        End If
    
        ' Go to the first record in the recordset and set the
        ' line ID
        ' Loop through the file
        Do While Not .EOF
            cboLine.AddItem !line_id
            .MoveNext
            intIndex = intIndex + 1
        Loop
        .Close
        
        .Source = "select v_prod_line_stocking_location.stocking_location_id, " & _
            "rtrim(stocking_location_description) as location_description, line_id " & _
            "from v_prod_line_stocking_location " & _
            "join v_prod_stocking_location on " & _
            "v_prod_line_stocking_location.stocking_location_id = " & _
            "v_prod_stocking_location.stocking_location_id " & _
            "order by line_id, location_description"
        .Open
        Set mxarrLineLocationInfo = New XArrayDB
        mxarrLineLocationInfo.LoadRows .GetRows(), True
        
        .Close
    End With
    
    Set rsList = Nothing
    If cboLine.ListIndex = 0 Then
        cboLine_Click
    Else
        cboLine.ListIndex = 0
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveLineAndLocationData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub RetrievePartData()
    ' Purpose:  Retrieve the part ID's from the table and
    '           build the array of parts and the part Description
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare the array size
    Dim intArrayCount As Integer
    
    ' Declare the recordset variable
    Dim rsPart As ADODB.Recordset
    
    ' Instantiate the recordset
    Set rsPart = New ADODB.Recordset

    Set mxarrPartInfo = New XArrayDB
    
        ' Set values of fields
        With rsPart
            'tells the recordset where to get its data from
            'i.e. gcondatabase="Provider=SQLOLEDB.1;PASSWORD=090400;USERID=SCHIMC
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            ' Change the literal below to the name of your view
            If optByModel Then
                .Source = "select v_prod_part.part_id, part_description, " & _
                    "stocking_location_description, level_number, quantity, v_mnb_model_part.part_sequence_number, " & _
                    "original_sequence_number " & _
                    "from v_mnb_model_part join v_prod_part on " & _
                    "v_mnb_model_part.part_id = v_prod_part.part_id " & _
                    "left outer join v_mnb_model_part_stocking_location on " & _
                    "v_mnb_model_part.model_number = v_mnb_model_part_stocking_location.model_number and " & _
                    "v_mnb_model_part.line_id = v_mnb_model_part_stocking_location.line_id and " & _
                    "v_mnb_model_part.part_id = v_mnb_model_part_stocking_location.part_id and " & _
                    "v_mnb_model_part.part_sequence_number = v_mnb_model_part_stocking_location.part_sequence_number " & _
                    "left outer join v_prod_stocking_location on " & _
                    "v_mnb_model_part_stocking_location.stocking_location_id = " & _
                    "v_prod_stocking_location.stocking_location_id " & _
                    "where v_mnb_model_part.model_number = '" & cboModel.Text & "' " & _
                    "and v_mnb_model_part.line_id = '" & cboLine.Text & "' " & _
                    "order by v_mnb_model_part.original_sequence_number, v_mnb_model_part.part_sequence_number"
            Else
                .Source = "select v_prod_part.part_id, part_description, " & _
                    "stocking_location_description, ' ' as level_number, 0 as quantity " & _
                    "from v_prod_part " & _
                    "left outer join v_mnb_part_line_stocking_location on " & _
                    "v_prod_part.part_id = v_mnb_part_line_stocking_location.part_id and " & _
                    "line_id = '" & cboLine.Text & "' " & _
                    "left outer join v_prod_stocking_location on " & _
                    "v_mnb_part_line_stocking_location.stocking_location_id = " & _
                    "v_prod_stocking_location.stocking_location_id " & _
                    "order by v_prod_part.part_id"
            End If
            
            '.LockType = adLockBatchOptimistic
            .LockType = adLockReadOnly
            .Open
            If .RecordCount = 0 Then
                MsgBox "No Parts Found."
                .Close
                GoTo PROC_EXIT
            End If
            
            mxarrPartInfo.LoadRows .GetRows(), True
            .Close
        End With
        Set rsPart = Nothing
        
        ' Check for errors returned from the recordset
        If gconDatabase.Errors.Count > 0 Then
            Err.Raise gconDatabase.Errors(0).NativeError, _
                "RetrievepartData", _
                gconDatabase.Errors(0).Description
        End If
    
    If optAll.Value Then
        If mrsPartLocation Is Nothing Then
            Set mrsPartLocation = New ADODB.Recordset
            With mrsPartLocation
                Set .ActiveConnection = gconDatabase
                .LockType = adLockBatchOptimistic
                .CursorLocation = adUseClient
                .CursorType = adOpenStatic
            End With
        ElseIf mrsPartLocation.State = adStateOpen Then
            mrsPartLocation.Close
        End If
        With mrsPartLocation
        'Temporarily change to display on parts for line selected
'            .Source = "select * from v_mnb_part_line_stocking_location " & _
'                "order by part_id"
'            .Open
            .Source = "select * from v_mnb_part_line_stocking_location " & _
                "where line_id = '" & cboLine.Text & "' " & _
                "order by part_id"
            .Open
        End With
    Else
        If mrsModelPartLocation Is Nothing Then
            Set mrsModelPartLocation = New ADODB.Recordset
            With mrsModelPartLocation
                Set .ActiveConnection = gconDatabase
                .LockType = adLockBatchOptimistic
                .CursorLocation = adUseClient
                .CursorType = adOpenStatic
            End With
        ElseIf mrsModelPartLocation.State = adStateOpen Then
            mrsModelPartLocation.Close
        End If
        With mrsModelPartLocation
            .Source = "select * from v_mnb_model_part_stocking_location " & _
                "where model_number = '" & cboModel.Text & "' and " & _
                "line_id = '" & cboLine.Text & "' " & _
                "order by part_id, part_sequence_number"
            .Open
        End With
    End If

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrievepartData", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub



Private Sub optAll_Click()
    ' Retrieve the data
    
    cboModel.Enabled = False
End Sub

Private Sub optByModel_Click()
    cboModel.Enabled = True
End Sub


Sub RetrieveModelData()
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    
    cboModel.Clear
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select distinct model_number from v_mnb_model_part where line_id = '" & _
            cboLine.Text & "' order by model_number asc"
        .Open
        
        Do While Not .EOF
            cboModel.AddItem Trim(!model_number)
            .MoveNext
        Loop
        .Close
    End With
    Set rsList = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FindDescription", _
        Err.Number, Err.Description)
    
End Sub

Private Sub TDBGAvailable_DblClick()
    cmdAdd_Click
End Sub

Private Sub txtPart_Change()
    Dim lngIndex As Long
    
    If Len(txtPart.Text) > 0 Then
        lngIndex = mxarrAvailable.Find(1, 0, txtPart.Text, XORDER_ASCEND, XCOMP_GE)
        TDBGAvailable.Scroll 0, lngIndex - TDBGAvailable.Bookmark
    
        TDBGAvailable.Row = 0
    End If
End Sub

Private Sub txtPart_GotFocus()
    txtPart.SelStart = 0
    txtPart.SelLength = Len(txtPart.Text)
End Sub

Private Sub txtPart_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub SaveData()
    ' Purpose:  Check to see if user wants to save changes. If so, save them.
    '           If not, return to the original state.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    ' Declare variables.
    Dim intRetCode As Integer
    
    mblnSaved = False
    
    intRetCode = MsgBox("Save Changes?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Minibill Save")
    If intRetCode = vbNo Then
        If mblnAllSave Then
            mrsPartLocation.CancelBatch
        Else
            mrsModelPartLocation.CancelBatch
        End If
        mblnSaved = True
        mblnRecChanged = False
        GoTo PROC_EXIT
    ElseIf intRetCode = vbCancel Then
        cboLine.Text = mstrLineSave
        cboLocation.Text = mstrLocationSave
        cboModel.Text = mstrModelSave
        optAll.Value = mblnAllSave
        GoTo PROC_EXIT
    Else
        If mblnAllSave Then
            mrsPartLocation.UpdateBatch
            mrsPartLocation.Requery
        Else
            mrsModelPartLocation.UpdateBatch
            mrsModelPartLocation.Requery
        End If
        If gconDatabase.Errors.Count > 0 Then
            Err.Raise gconDatabase.Errors(0).NativeError, "SaveData", _
                gconDatabase.Errors(0).Description
        End If
        mblnSaved = True
        mblnRecChanged = False
    End If

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "SaveData", Err.Number, Err.Description)
    GoTo PROC_EXIT
    
End Sub



Private Sub cmdSplitPart_Click()
    ' Purpose:  To display a form allowing the user to split a part among multiple locations.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim intRetCode As Integer
    Dim rsList As ADODB.Recordset
    Dim lngTotalQuantity As Long
    Dim lngRow As Long
    Dim xarrLocations As XArrayDB
    
    Set xarrLocations = New XArrayDB
    xarrLocations.ReDim 1, cboLocation.ListCount, 0, 1
    For lngRow = 0 To cboLocation.ListCount - 1
        xarrLocations(lngRow + 1, 0) = mxarrLineLocationInfo(cboLocation.ItemData(lngRow), 0)
        xarrLocations(lngRow + 1, 1) = cboLocation.List(lngRow)
    Next lngRow
    
    If mblnRecChanged Then
        SaveData
        If Not mblnSaved Then
            Exit Sub
        End If
    End If
    
    If TDBGAvailable.Row < 0 Then
        MsgBox "Select a part to be split before clicking this button."
        Exit Sub
    End If
    
    Set rsList = New ADODB.Recordset
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select sum(quantity) as Total_Quantity " & _
            "from v_mnb_model_part where model_number = '" & mstrModelSave & _
            "' and line_id = '" & mstrLineSave & "' " & _
            " and part_id = '" & TDBGAvailable.Columns(0).Value & "' and original_sequence_number = " & _
            mxarrAvailable(TDBGAvailable.RowBookmark(TDBGAvailable.Row), 5)
        .Open
        lngTotalQuantity = !total_quantity
        .Close
    End With
    Set rsList = Nothing
    
    If lngTotalQuantity < 2 Then
        MsgBox "Select a part with quantity greater than 1 before clicking this button."
        Exit Sub
    End If
    
    With frmSplitPart
        .txtLine.Text = mstrLineSave
        .txtModel.Text = mstrModelSave
        .txtPart.Text = TDBGAvailable.Columns(0).Value
        .txtPartDescription.Text = TDBGAvailable.Columns(1).Value
        .txtQuantity.Text = lngTotalQuantity
        .mlngPartSequence = mxarrAvailable(TDBGAvailable.RowBookmark(TDBGAvailable.Row), 5)
        Set .mxarrLocations = xarrLocations
        Set .mrsModelPartLocation = mrsModelPartLocation
        .ProcessPart
    End With
    cmdRefresh_Click
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdSplitPart_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub



