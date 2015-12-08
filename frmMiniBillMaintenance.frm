VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmMiniBillMaintenance 
   Caption         =   "MiniBill - Model Maintenance"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
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
   ScaleHeight     =   558
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   795
   StartUpPosition =   1  'CenterOwner
   Begin SizerOneLibCtl.ElasticOne ElasticOne1 
      Height          =   8370
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   11925
      _cx             =   21034
      _cy             =   14764
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
      Begin VB.CommandButton cmdFindPartDescription 
         Caption         =   "Find &Description"
         Height          =   435
         Left            =   2880
         TabIndex        =   19
         Top             =   7800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtModelDescription 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Text            =   " "
         Top             =   960
         Width           =   3135
      End
      Begin VB.CheckBox chkSortByPart 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sort by Part"
         Height          =   315
         Left            =   5100
         TabIndex        =   4
         Top             =   510
         Width           =   3840
      End
      Begin VB.CommandButton cmdFindPart 
         Caption         =   "F&ind Part"
         Height          =   435
         Left            =   840
         TabIndex        =   9
         Top             =   7800
         Width           =   1815
      End
      Begin VB.CheckBox chkECNParts 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Show Only Parts With ECNs"
         Height          =   315
         Left            =   5100
         TabIndex        =   6
         Top             =   1080
         Width           =   3840
      End
      Begin VB.CheckBox chkShowAssigned 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Show Only Assigned Parts"
         Height          =   315
         Left            =   5100
         TabIndex        =   5
         Top             =   795
         Width           =   3840
      End
      Begin VB.CheckBox chkSortByLocation 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sort By Location"
         Height          =   315
         Left            =   5100
         TabIndex        =   3
         Top             =   225
         Width           =   3840
      End
      Begin VB.CommandButton cmdSplitPart 
         Caption         =   "Split &Part"
         Height          =   435
         Left            =   4995
         TabIndex        =   10
         Top             =   7800
         Width           =   1800
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   435
         Left            =   7230
         TabIndex        =   11
         Top             =   7800
         Width           =   1800
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "&Refresh Display"
         Height          =   435
         Left            =   9510
         TabIndex        =   7
         Top             =   540
         Width           =   2130
      End
      Begin VB.CheckBox chkUseDefaults 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Use Default Locations By Part"
         Height          =   315
         Left            =   5100
         TabIndex        =   2
         Top             =   -60
         Width           =   3840
      End
      Begin VB.ComboBox cboLine 
         Height          =   360
         Left            =   975
         TabIndex        =   0
         Top             =   300
         Width           =   690
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         CausesValidation=   0   'False
         Height          =   435
         Left            =   9360
         TabIndex        =   12
         Top             =   7800
         Width           =   1650
      End
      Begin VB.ComboBox cboModel 
         Height          =   360
         Left            =   2640
         TabIndex        =   1
         Top             =   300
         Width           =   1875
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGMiniBill 
         Height          =   6255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   11730
         _ExtentX        =   20690
         _ExtentY        =   11033
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Line"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Model"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Part"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Part Sequence"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Description"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Lvl"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Qty"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Start ECN #"
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Start Date"
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Quit ECN #"
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Quit Date"
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   4
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Default?"
         Columns(11).DataField=   ""
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   1
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Location"
         Columns(12).DataField=   ""
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Step"
         Columns(13).DataField=   ""
         Columns(13).ExternalEditor=   "TDBNumber1"
         Columns(13).ExternalEditor.vt=   8
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   1
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Category"
         Columns(14).DataField=   ""
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   15
         Splits(0)._UserFlags=   0
         Splits(0).AnchorRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   2
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   -1  'True
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=15"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=1217"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=1138"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=3228"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=3149"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=556"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=476"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=609"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=529"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(43)=   "Column(7).Width=1826"
         Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=1746"
         Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(49)=   "Column(8).Width=1588"
         Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=1508"
         Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(53)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(55)=   "Column(9).Width=1720"
         Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=1640"
         Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(59)=   "Column(9).AllowFocus=0"
         Splits(0)._ColumnProps(60)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(61)=   "Column(10).Width=1429"
         Splits(0)._ColumnProps(62)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(10)._WidthInPix=1349"
         Splits(0)._ColumnProps(64)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(65)=   "Column(10).AllowFocus=0"
         Splits(0)._ColumnProps(66)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(67)=   "Column(11).Width=1296"
         Splits(0)._ColumnProps(68)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(11)._WidthInPix=1217"
         Splits(0)._ColumnProps(70)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(71)=   "Column(11)._ColStyle=1"
         Splits(0)._ColumnProps(72)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(73)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(74)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(76)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(77)=   "Column(12).Button=1"
         Splits(0)._ColumnProps(78)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(79)=   "Column(12).DropDownList=1"
         Splits(0)._ColumnProps(80)=   "Column(13).Width=794"
         Splits(0)._ColumnProps(81)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(82)=   "Column(13)._WidthInPix=714"
         Splits(0)._ColumnProps(83)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(84)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(85)=   "Column(14).Width=2487"
         Splits(0)._ColumnProps(86)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(14)._WidthInPix=2408"
         Splits(0)._ColumnProps(88)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(89)=   "Column(14).Button=1"
         Splits(0)._ColumnProps(90)=   "Column(14).AllowFocus=0"
         Splits(0)._ColumnProps(91)=   "Column(14).Order=15"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=975,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=975,.italic=0"
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
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1,.bold=0,.fontsize=750,.italic=0,.underline=0"
         _StyleDefs(25)  =   ":id=13,.strikethrough=0,.charset=0"
         _StyleDefs(26)  =   ":id=13,.fontname=Arial"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=675,.italic=0"
         _StyleDefs(29)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(30)  =   ":id=14,.fontname=Small Fonts"
         _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(34)  =   "Splits(0).EditorStyle:id=17,.parent=7,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(35)  =   ":id=17,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(36)  =   ":id=17,.fontname=MS Sans Serif"
         _StyleDefs(37)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(38)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(39)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(40)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(41)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(42)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(3).Style:id=82,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(3).HeadingStyle:id=79,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(3).FooterStyle:id=80,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(3).EditorStyle:id=81,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(5).Style:id=94,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(5).HeadingStyle:id=91,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(5).FooterStyle:id=92,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(5).EditorStyle:id=93,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(8).Style:id=50,.parent=13"
         _StyleDefs(75)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(9).Style:id=54,.parent=13"
         _StyleDefs(79)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
         _StyleDefs(80)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(10).Style:id=58,.parent=13"
         _StyleDefs(83)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=17"
         _StyleDefs(86)  =   "Splits(0).Columns(11).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(87)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=14"
         _StyleDefs(88)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=15"
         _StyleDefs(89)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=17"
         _StyleDefs(90)  =   "Splits(0).Columns(12).Style:id=70,.parent=13,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(91)  =   ":id=70,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(92)  =   ":id=70,.fontname=MS Sans Serif"
         _StyleDefs(93)  =   "Splits(0).Columns(12).HeadingStyle:id=67,.parent=14"
         _StyleDefs(94)  =   "Splits(0).Columns(12).FooterStyle:id=68,.parent=15"
         _StyleDefs(95)  =   "Splits(0).Columns(12).EditorStyle:id=69,.parent=17,.bold=-1,.fontsize=825"
         _StyleDefs(96)  =   ":id=69,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(97)  =   ":id=69,.fontname=MS Sans Serif"
         _StyleDefs(98)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
         _StyleDefs(99)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
         _StyleDefs(100) =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
         _StyleDefs(101) =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
         _StyleDefs(102) =   "Splits(0).Columns(14).Style:id=74,.parent=13"
         _StyleDefs(103) =   "Splits(0).Columns(14).HeadingStyle:id=71,.parent=14"
         _StyleDefs(104) =   "Splits(0).Columns(14).FooterStyle:id=72,.parent=15"
         _StyleDefs(105) =   "Splits(0).Columns(14).EditorStyle:id=73,.parent=17"
         _StyleDefs(106) =   "Named:id=33:Normal"
         _StyleDefs(107) =   ":id=33,.parent=0,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
         _StyleDefs(108) =   ":id=33,.charset=0"
         _StyleDefs(109) =   ":id=33,.fontname=Arial"
         _StyleDefs(110) =   "Named:id=34:Heading"
         _StyleDefs(111) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(112) =   ":id=34,.wraptext=-1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(113) =   ":id=34,.strikethrough=0,.charset=0"
         _StyleDefs(114) =   ":id=34,.fontname=Arial"
         _StyleDefs(115) =   "Named:id=35:Footing"
         _StyleDefs(116) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(117) =   "Named:id=36:Selected"
         _StyleDefs(118) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(119) =   "Named:id=37:Caption"
         _StyleDefs(120) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(121) =   "Named:id=38:HighlightRow"
         _StyleDefs(122) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(123) =   "Named:id=39:EvenRow"
         _StyleDefs(124) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(125) =   "Named:id=40:OddRow"
         _StyleDefs(126) =   ":id=40,.parent=33"
         _StyleDefs(127) =   "Named:id=41:RecordSelector"
         _StyleDefs(128) =   ":id=41,.parent=34"
         _StyleDefs(129) =   "Named:id=42:FilterBar"
         _StyleDefs(130) =   ":id=42,.parent=33"
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
         Height          =   255
         Left            =   255
         TabIndex        =   16
         Top             =   7680
         Visible         =   0   'False
         Width           =   510
         _Version        =   65536
         _ExtentX        =   900
         _ExtentY        =   450
         Calculator      =   "frmMiniBillMaintenance.frx":0000
         Caption         =   "frmMiniBillMaintenance.frx":0020
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmMiniBillMaintenance.frx":008C
         Keys            =   "frmMiniBillMaintenance.frx":00AA
         Spin            =   "frmMiniBillMaintenance.frx":00F4
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   1
         DecimalPoint    =   "."
         DisplayFormat   =   "########;-########;Null;Zero"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   1
         ForeColor       =   -2147483640
         Format          =   "000"
         HighlightText   =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   999
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   ","
         ShowContextMenu =   0
         ValueVT         =   1
         Value           =   0
         MaxValueVT      =   1464729605
         MinValueVT      =   1312882693
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
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
         Index           =   1
         Left            =   375
         TabIndex        =   18
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   240
         Index           =   6
         Left            =   375
         TabIndex        =   15
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
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
         Left            =   1710
         TabIndex        =   14
         Top             =   360
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmMiniBillMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsModelPartLocation As ADODB.Recordset
Public mrsPartLocation As ADODB.Recordset
Private mrsPartCategory As ADODB.Recordset

Private mblnRecChanged As Boolean

Private mxarrLocations As XArrayDB
Private mxarrLocationCategories As XArrayDB
Private mxarrCategories As XArrayDB
Public mxarrMiniBill As XArrayDB

Private mblnBuildingGrid As Boolean
Private mblnSaved As Boolean
Private col As Integer

Private Sub cboLine_Change()
    If Len(cboLine.Text) = 2 Then
        cboFindFirst cboLine
    End If
    
End Sub
Private Sub cboLine_Click()
    RetrieveModelData
    txtModelDescription.Text = " "
    LoadLocations
    If cboModel.ListCount > 0 And mxarrLocations.UpperBound(1) > 0 And Screen.ActiveForm Is Me Then
        cboModel.SetFocus
    End If
End Sub

Private Sub cboLine_GotFocus()
    cboLine.SelStart = 0
    cboLine.SelLength = Len(cboLine.Text)
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
        
        If mxarrLocations.UpperBound(1) = 0 Then
            Cancel = True
            cboLine_GotFocus
        Else
            Cancel = False
        End If
    End If
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    Call ShowError(Me.Name, "cboLineID_Validate", _
        Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cboModel_Change()
    cboFindFirst cboModel
End Sub


Private Sub cboModel_Click()
    Call RetrieveModelDescription
End Sub

Private Sub cboModel_GotFocus()
    cboModel.SelStart = 0
    cboModel.SelLength = Len(cboModel.Text)
End Sub
Private Sub cboModel_KeyPress(KeyAscii As Integer)
    cboKeyPress cboModel, KeyAscii
End Sub

Private Sub cboModel_Validate(Cancel As Boolean)
    ' Purpose:  Check to see that a model was selected.
    
    If Len(cboModel.Text) = 0 Then
        MsgBox "Model is required."
        Cancel = True
        Exit Sub
    End If
    
    cboFindFirst cboModel
    If cboModel.ListIndex = -1 Then
        MsgBox "Invalid Model selected."
        Cancel = True
        Exit Sub
    End If
    
    Call RetrieveModelDescription
    
    
End Sub

Private Sub chkUseDefaults_Click()
    If Len(cboModel.Text) > 0 Then
        mrsPartLocation.CancelBatch
        mrsPartLocation.Requery
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdDisplay_Click()
    ' Purpose:  Build the model/part display
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    Dim intRetCode As Integer
    Dim strOrderBy As String
    Dim strWhere As String
    Dim blnFound As Boolean
    
    mblnBuildingGrid = True
    
    If mblnRecChanged Then
        intRetCode = MsgBox("Save Changes?", vbYesNo + vbQuestion + vbDefaultButton1, "Update")
        If intRetCode = vbYes Then
            cmdSave_Click
            If Not mblnSaved Then
                GoTo PROC_EXIT
            End If
        Else
            mrsModelPartLocation.CancelBatch
            mrsPartLocation.CancelBatch
            mrsPartCategory.CancelBatch
        End If
    End If
    
    If mxarrMiniBill Is Nothing Then
        Set mxarrMiniBill = New XArrayDB
    ElseIf mxarrMiniBill.UpperBound(1) > 0 Then
        mxarrMiniBill.Clear
    End If
    
    If Me.chkSortByLocation.Value = 1 And Me.chkSortByPart.Value = 1 Then
        MsgBox "Invalid selection, Both Sort By Location and Sort By Part are checked"
        GoTo PROC_EXIT
    End If
    
    If Me.chkSortByLocation.Value = 1 Then
        strOrderBy = "order by stocking_location_id, step_number, original_sequence_number, v_mnb_model_part.part_sequence_number"
    ElseIf Me.chkSortByPart.Value = 1 Then
        strOrderBy = "order by v_mnb_model_part.part_id, original_sequence_number, v_mnb_model_part.part_sequence_number"
    Else
        strOrderBy = "order by original_sequence_number, v_mnb_model_part.part_sequence_number"
    End If
    
    strWhere = ""
    
    If chkECNParts.Value Then
        strWhere = " and (start_ecn_number <> ' ' or quit_ecn_number <> ' ')"
    End If
    
    If Me.chkShowAssigned.Value Then
        strWhere = strWhere & " and v_mnb_model_part_stocking_location.stocking_location_id is not null "
    End If
    
    Set rsList = New ADODB.Recordset
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        If chkUseDefaults.Value Then
            .Source = "select min(v_mnb_model_part.line_id) as line_id, min(v_mnb_model_part.model_number) as model_number, " & _
                "v_mnb_model_part.part_id, v_mnb_model_part.part_sequence_number, " & _
                "min(part_description) as part_description, min(level_number) as level_number, min(quantity) as quantity, min(start_ecn_number) as start_ecn_number, " & _
                "min(case when start_ecn_date = ' ' then ' ' else substring(start_ecn_date,5,2) + '/' + substring(start_ecn_date,7,2) + '/' + substring(start_ecn_date,1,4) end) as start_date, " & _
                "min(quit_ecn_number) as quit_ecn_number, min(case when quit_ecn_date = ' ' then ' ' else substring(quit_ecn_date,5,2) + '/' + substring(quit_ecn_date,7,2) + '/' + substring(quit_ecn_date,1,4) end) as quit_date, " & _
                "min(case when v_mnb_model_part_stocking_location.stocking_location_id is not null " & _
                " and (v_mnb_part_line_stocking_location.stocking_location_id is null or v_mnb_part_line_stocking_location.stocking_location_id <> v_mnb_model_part_stocking_location.stocking_location_id) then 0 " & _
                "else case when v_mnb_part_line_stocking_location.stocking_location_id is null then 0 else -1 end end) as default_setup, " & _
                "min(case when v_mnb_model_part_stocking_location.stocking_location_id is null then v_mnb_part_line_stocking_location.stocking_location_id else " & _
                "v_mnb_model_part_stocking_location.stocking_location_id end) as stocking_location_id, " & _
                "min(step_number) as step_number, min(case when minibill_only_flag = 1 then v_mnb_category_part.category_id else null end) as category_id, min(v_mnb_model_part.original_sequence_number) as original_sequence_number " & _
                "from v_mnb_model_part left outer join v_mnb_model_part_stocking_location on " & _
                "v_mnb_model_part.line_id = v_mnb_model_part_stocking_location.line_id and " & _
                "v_mnb_model_part.model_number = v_mnb_model_part_stocking_location.model_number  and " & _
                "v_mnb_model_part.part_id = v_mnb_model_part_stocking_location.part_id and " & _
                "v_mnb_model_part.part_sequence_number = v_mnb_model_part_stocking_location.part_sequence_number " & _
                "left outer join v_mnb_part_line_stocking_location on " & _
                "v_mnb_model_part.line_id = v_mnb_part_line_stocking_location.line_id and " & _
                "v_mnb_model_part.part_id = v_mnb_part_line_stocking_location.part_id " & _
                "join v_prod_part on v_mnb_model_part.part_id = v_prod_part.part_id " & _
                "left outer join v_mnb_category_part on v_prod_part.part_id = v_mnb_category_part.part_id " & _
                "left outer join v_mnb_category on v_mnb_category_part.category_id = v_mnb_category.category_id " & _
                "where v_mnb_model_part.model_number = '" & cboModel.Text & "' " & _
                " and v_mnb_model_part.line_id = '" & cboLine.Text & "' " & strWhere & " group by v_mnb_model_part.part_id, v_mnb_model_part.part_sequence_number " & _
                strOrderBy
            TDBGMiniBill.Columns(11).Visible = True
        Else
            .Source = "select min(v_mnb_model_part.line_id) as line_id, min(v_mnb_model_part.model_number) as model_number, " & _
                "v_mnb_model_part.part_id, v_mnb_model_part.part_sequence_number, " & _
                "min(part_description) as part_description, min(level_number) as level_number, min(quantity) as quantity, min(start_ecn_number) as start_ecn_number, " & _
                "min(case when start_ecn_date = ' ' then ' ' else substring(start_ecn_date,5,2) + '/' + substring(start_ecn_date,7,2) + '/' + substring(start_ecn_date,1,4) end) as start_date, " & _
                "min(quit_ecn_number) as quit_ecn_number, min(case when quit_ecn_date = ' ' then ' ' else substring(quit_ecn_date,5,2) + '/' + substring(quit_ecn_date,7,2) + '/' + substring(quit_ecn_date,1,4) end) as quit_date, " & _
                "0 as default_setup, " & _
                "min(v_mnb_model_part_stocking_location.stocking_location_id) as stocking_location_id, min(step_number) as step_number, " & _
                "min(case when minibill_only_flag = 1 then v_mnb_category_part.category_id else null end) as category_id, min(v_mnb_model_part.original_sequence_number) as original_sequence_number " & _
                "from v_mnb_model_part left outer join v_mnb_model_part_stocking_location on " & _
                "v_mnb_model_part.line_id = v_mnb_model_part_stocking_location.line_id and " & _
                "v_mnb_model_part.model_number = v_mnb_model_part_stocking_location.model_number  and " & _
                "v_mnb_model_part.part_id = v_mnb_model_part_stocking_location.part_id and " & _
                "v_mnb_model_part.part_sequence_number = v_mnb_model_part_stocking_location.part_sequence_number " & _
                "join v_prod_part on v_mnb_model_part.part_id = v_prod_part.part_id " & _
                "left outer join v_mnb_category_part on v_prod_part.part_id = v_mnb_category_part.part_id " & _
                "left outer join v_mnb_category on v_mnb_category_part.category_id = v_mnb_category.category_id and minibill_only_flag = 1 " & _
                "where v_mnb_model_part.model_number = '" & cboModel.Text & "' " & _
                " and v_mnb_model_part.line_id = '" & cboLine.Text & "' " & strWhere & "  group by v_mnb_model_part.part_id, v_mnb_model_part.part_sequence_number " & _
                strOrderBy
            If TDBGMiniBill.Columns(11).Visible Then
                TDBGMiniBill.Columns(11).Visible = False
            End If
        End If
        .Open
        If .RecordCount > 0 Then
            mxarrMiniBill.LoadRows rsList.GetRows, True
        Else
            MsgBox "No information for the selections made."
            .Close
            GoTo PROC_EXIT
        End If
        .Close
    End With
    Set rsList = Nothing
    LoadDataGrid
    
    With mrsModelPartLocation
        If .State = adStateOpen Then
            .Close
        End If
        .Source = "select * from v_mnb_model_part_stocking_location " & _
            "where model_number = '" & cboModel.Text & "' and line_id = '" & cboLine.Text & "'" & _
            "order by part_id"
        .Open
    End With
    mblnRecChanged = False
    TDBGMiniBill.Row = 0
    
    If Me.chkUseDefaults.Value <> 0 Then
        For intIndex = 0 To mxarrMiniBill.UpperBound(1)
            With mrsPartLocation
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Find "part_id = '" & mxarrMiniBill(intIndex, 2) & "'"
                End If
                Do While Not .EOF
                    If !line_id <> cboLine.Text Then
                        .MoveNext
                        If Not .EOF Then
                            .Find "part_id = '" & mxarrMiniBill(intIndex, 2) & "'"
                        End If
                    Else
                        Exit Do
                    End If
                Loop
                If .EOF Then
                    If mxarrMiniBill(intIndex, 11) > 0 Then
                        .AddNew
                        !line_id = mxarrMiniBill(intIndex, 0)
                        !part_id = mxarrMiniBill(intIndex, 2)
                        If Not IsNull(mxarrMiniBill(intIndex, 12)) Then
                            Debug.Assert False
                        End If
                        !stocking_location_id = mxarrMiniBill(intIndex, 12)
                        mblnRecChanged = True
                    End If
                Else
                    blnFound = False
                    If mrsModelPartLocation.RecordCount > 0 Then
                        mrsModelPartLocation.MoveFirst
                        mrsModelPartLocation.Find "part_id = '" & !part_id & "'"
                        Do While Not mrsModelPartLocation.EOF And Not blnFound
                            If mrsModelPartLocation!part_sequence_number = mxarrMiniBill(intIndex, 3) Then
                                blnFound = True
                                If IsNull(mrsModelPartLocation!stocking_location_id) Then
                                    mrsModelPartLocation!stocking_location_id = !stocking_location_id
                                    mblnRecChanged = True
                                End If
                            Else
                                mrsModelPartLocation.MoveNext
                                If Not mrsModelPartLocation.EOF Then
                                    mrsModelPartLocation.Find "part_id = '" & !part_id & "'"
                                End If
                            End If
                        Loop
                    End If
                    
                    If Not blnFound Then
                        mrsModelPartLocation.AddNew
                        mrsModelPartLocation!model_number = cboModel.Text
                        mrsModelPartLocation!line_id = cboLine.Text
                        mrsModelPartLocation!part_id = !part_id
                        mrsModelPartLocation!stocking_location_id = !stocking_location_id
                        mrsModelPartLocation!part_sequence_number = mxarrMiniBill(intIndex, 3)
                        mblnRecChanged = True
                    End If
                End If
            End With
        Next intIndex
    End If
    
    cmdSplitPart.Enabled = True
    cmdFindPart.Enabled = True
    cmdFindPartDescription.Enabled = True
    cmdSave.Enabled = True
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdDisplay_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdFindPart_Click()
    frmMinibillPartFind.Show vbModal
End Sub
Private Sub cmdFindPartDescription_Click()
    frmMinibillFindPartDescription.Show vbModal
End Sub
Private Sub cmdSplitPart_Click()
    ' Purpose:  To display a form allowing the user to split a part among multiple locations.
    '
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim intRetCode As Integer
    Dim rsList As ADODB.Recordset
    Dim lngTotalQuantity As Long
    Dim lngRow As Long
    
    
    If TDBGMiniBill.Row < 0 Then
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
            "from v_mnb_model_part where model_number = '" & cboModel.Text & _
            "' and line_id = '" & cboLine.Text & "' " & _
            " and part_id = '" & TDBGMiniBill.Columns(2).Value & "' and original_sequence_number = " & _
            mxarrMiniBill(TDBGMiniBill.RowBookmark(TDBGMiniBill.Row), 15)
        .Open
        lngTotalQuantity = !total_quantity
        .Close
    End With
    Set rsList = Nothing
    
    If lngTotalQuantity < 2 Then
        MsgBox "Select a part with quantity greater than 1 before clicking this button."
        Exit Sub
    End If
    
    If mblnRecChanged Then
        intRetCode = MsgBox("Save Changes?", vbYesNo + vbQuestion + vbDefaultButton1, "Update")
        If intRetCode = vbYes Then
            cmdSave_Click
            If Not mblnSaved Then
                GoTo PROC_EXIT
            End If
        Else
            mrsModelPartLocation.CancelBatch
            mrsPartLocation.CancelBatch
            mrsPartCategory.CancelBatch
            mblnRecChanged = False
        End If
    End If
    
    With frmSplitPart
        .txtLine.Text = cboLine.Text
        .txtModel.Text = cboModel.Text
        .txtPart.Text = TDBGMiniBill.Columns(2).Value
        .txtPartDescription.Text = TDBGMiniBill.Columns(4).Value
        .txtQuantity.Text = lngTotalQuantity
        .mlngPartSequence = mxarrMiniBill(TDBGMiniBill.RowBookmark(TDBGMiniBill.Row), 15)
        Set .mxarrLocations = mxarrLocations
        Set .mrsModelPartLocation = mrsModelPartLocation
        .ProcessPart
    End With
    cmdDisplay_Click
        
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdSplitPart_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub cmdSave_Click()
    ' Purpose:  Save current changes.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    mblnSaved = False
    
    With gconDatabase
        .Errors.Clear
        If .Errors.Count > 0 Then
            Err.Raise .Errors(0).NativeError, "cmdSave_click", .Errors(0).Description
        End If
        
        mrsModelPartLocation.UpdateBatch
        If .Errors.Count > 0 Then
            Err.Raise .Errors(0).NativeError, "cmdSave_click", .Errors(0).Description
        End If
        
        mrsPartLocation.UpdateBatch
        If .Errors.Count > 0 Then
            Err.Raise .Errors(0).NativeError, "cmdSave_click", .Errors(0).Description
        End If
        
        mrsPartCategory.UpdateBatch
        If .Errors.Count > 0 Then
            Err.Raise .Errors(0).NativeError, "cmdSave_click", .Errors(0).Description
        End If
        
    End With
    
    gblnMaintPassedUpdates = True
    MsgBox "Records successfully saved."
    mblnRecChanged = False
    mblnSaved = True

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdSave_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    ' Purpose:  Load form data
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    If gconDatabase Is Nothing Then
        Set gconDatabase = gclsSQLServer.Connect( _
            gclsMESApplication.ApplicationRole, _
            gclsMESApplication.ApplicationPassword)
        
        If gconDatabase.State <> adStateOpen Then
            Err.Raise gconDatabase.Errors(0).NativeError, _
                "Form_Load", gconDatabase.Errors(0).Description
        End If
    End If
    
    LoadLines
    LoadCategories
    cboModel.Enabled = False
    cmdSplitPart.Enabled = False
    cmdFindPart.Enabled = False
    cmdFindPartDescription.Enabled = False
    cmdSave.Enabled = False
    mblnRecChanged = False
    
    Set mrsPartLocation = New ADODB.Recordset
    Set mrsPartCategory = New ADODB.Recordset
    Set mrsModelPartLocation = New ADODB.Recordset
    
    With mrsPartLocation
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
    End With
    
    With mrsPartCategory
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Source = "select * from v_mnb_category_part order by part_id"
        .Open
    End With
    
    With mrsModelPartLocation
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub LoadDataGrid()
    ' Purpose:  Load the grid with data from the recordset.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim intIndex As Integer
    Dim Item As TrueOleDBGrid70.ValueItem
    Set Item = New TrueOleDBGrid70.ValueItem
    
    With Me.TDBGMiniBill
        .Array = mxarrMiniBill
        .ReBind
        
        With .Columns(12)
            .AutoCompletion = True
            .ButtonAlways = True
            With .ValueItems
                .Clear
                Item.Value = ""
                Item.DisplayValue = " "
                .Add Item
                For intIndex = 1 To mxarrLocations.UpperBound(1)
                    Item.Value = mxarrLocations.Value(intIndex, 0)
                    Item.DisplayValue = mxarrLocations.Value(intIndex, 1)
                    .Add Item
                Next intIndex
                .Translate = True
            End With
        End With
        
        With .Columns(14)
            .AllowFocus = False
            If .ValueItems.Count = 0 Then
                .AutoCompletion = True
                With .ValueItems
                    For intIndex = 1 To mxarrCategories.UpperBound(1)
                        Item.Value = mxarrCategories.Value(intIndex, 0)
                        Item.DisplayValue = mxarrCategories.Value(intIndex, 1)
                        .Add Item
                    Next intIndex
                    .Translate = True
                End With
            End If
        End With
        .Splits(0).SpringMode = False
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "LoadDataGrid", Err.Number, _
        Err.Description)
    GoTo PROC_EXIT
End Sub

Sub RetrieveModelData()
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    
    cboModel.Clear
    
    ' Changed by CAS to exclude inactive models
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select distinct model_number from v_mnb_model_line " & _
            "where line_id = '" & cboLine.Text & "' " & _
            "order by model_number asc"
        .Open
        
        Do While Not .EOF
            cboModel.AddItem Trim(!model_number)
            .MoveNext
        Loop
        .Close
    End With
    Set rsList = Nothing
    If cboModel.ListCount = 0 Then
        MsgBox "No models were found for the line selected."
        cboModel.Enabled = False
        cboLine.SetFocus
    Else
        cboModel.Enabled = True
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FindDescription", _
        Err.Number, Err.Description)
End Sub
Sub RetrieveModelDescription()
    ' Purpose:  Fill Model Label Field with Model Selected
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsModelDescription As ADODB.Recordset
    Set rsModelDescription = New ADODB.Recordset
    
'    cboModel.Clear
    
    With rsModelDescription
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select distinct Sales_Model_Description from v_prod_model_line " & _
            "join v_prod_model on v_prod_model_line.sales_model_number = v_prod_model.sales_model_number " & _
            "where engineering_model_number = '" & cboModel.Text & "' "
        .Open
        
        If .EOF Then
            txtModelDescription.Text = "Not Found"
            .Close
            GoTo PROC_EXIT
        End If
        
        txtModelDescription.Text = Trim(!Sales_Model_Description)
        
        .Close
    End With
    Set rsModelDescription = Nothing
    
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "RetrieveModelDescription", _
        Err.Number, Err.Description)
End Sub
Private Sub LoadLocations()
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    
    Set rsList = New ADODB.Recordset
    
    Set mxarrLocations = New XArrayDB
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select v_prod_stocking_location.stocking_location_id, " & _
            "min(stocking_location_description) as stocking_location_description, " & _
            "count(category_id) as Categories " & _
            "from v_prod_stocking_location " & _
            "join v_prod_line_stocking_location on " & _
            "v_prod_stocking_location.stocking_location_id = " & _
            "v_prod_line_stocking_location.stocking_location_id " & _
            "left outer join v_mnb_stocking_location_category on " & _
            "v_prod_stocking_location.stocking_location_id = " & _
            "v_mnb_stocking_location_category.stocking_location_id " & _
            "where line_id = '" & cboLine.Text & "' " & _
            "group by v_prod_stocking_location.stocking_location_id " & _
            "order by stocking_location_description"
        .Open
        
        mxarrLocations.ReDim 1, .RecordCount, 0, 1
        
        intIndex = 1
        Do While Not .EOF
            mxarrLocations(intIndex, 0) = .Fields("stocking_location_id").Value
            mxarrLocations(intIndex, 1) = Trim(!stocking_location_description)
            intIndex = intIndex + 1
            .MoveNext
        Loop
        .Close
        If mxarrLocations.UpperBound(1) = 0 Then
            MsgBox "No Locations found for the line selected."
            Set rsList = Nothing
            If Screen.ActiveForm Is Me Then
                cboLine.SetFocus
                cboLine_GotFocus
            End If
            GoTo PROC_EXIT
        End If
    End With
    Set rsList = Nothing
    
    With mrsPartLocation
        If .State = adStateOpen Then
            .Close
        End If
        .Source = "select * from v_mnb_part_line_stocking_location " & _
            "where line_id = '" & cboLine.Text & "' " & _
            "order by part_id"
        .Open
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FindDescription", _
        Err.Number, Err.Description)
End Sub

Private Sub LoadCategories()
    ' Purpose:  Fill the model combo box
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Dim intIndex As Integer
    
    Set rsList = New ADODB.Recordset
    
    Set mxarrCategories = New XArrayDB
    
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select * from v_mnb_category " & _
            "where minibill_only_flag = 1 " & _
            "order by category_description"
        .Open
        
        mxarrCategories.ReDim 1, .RecordCount, 0, 1
        
        intIndex = 1
        Do While Not .EOF
            mxarrCategories(intIndex, 0) = .Fields("category_id").Value
            mxarrCategories(intIndex, 1) = Trim(!Category_description)
            intIndex = intIndex + 1
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Purpose:   Complete processing before closing
    
    Dim intRetCode As Integer
    
    If mblnRecChanged Then
        intRetCode = MsgBox("Save Changes?", vbYesNo + vbQuestion + vbDefaultButton1, "Update")
        If intRetCode = vbYes Then
            cmdSave_Click
            If Not mblnSaved Then
                Cancel = True
                Exit Sub
            End If
        Else
            mrsModelPartLocation.CancelBatch
            mrsPartLocation.CancelBatch
            mrsPartCategory.CancelBatch
        End If
    End If
    
    Set mxarrMiniBill = Nothing
End Sub

Private Sub Form_Resize()
    Me.ElasticOne1.Refresh
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gconDatabase.State = adStateOpen Then
        If Not mrsPartCategory Is Nothing Then
            If mrsPartCategory.State = adStateOpen Then
                mrsPartCategory.Close
            End If
            Set mrsPartCategory = Nothing
        End If
        If Not mrsPartLocation Is Nothing Then
            If mrsPartLocation.State = adStateOpen Then
                mrsPartLocation.Close
            End If
            Set mrsPartLocation = Nothing
        End If
        If Not mrsModelPartLocation Is Nothing Then
            If mrsModelPartLocation.State = adStateOpen Then
                mrsModelPartLocation.Close
            End If
            Set mrsModelPartLocation = Nothing
        End If
    End If
End Sub

Private Sub TDBGMiniBill_AfterColUpdate(ByVal ColIndex As Integer)
    ' Purpose:  Update the recordsets that correspond with the data being set.
    
    mblnRecChanged = True
    Select Case ColIndex
        Case 12
            With mrsModelPartLocation
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Find "part_id = '" & TDBGMiniBill.Columns(2).Value & "'"
                    Do While Not .EOF
                        If !part_sequence_number = Val(TDBGMiniBill.Columns(3).Value) Then
                            Exit Do
                        End If
                        .MoveNext
                        If Not .EOF Then
                            .Find "part_id = '" & TDBGMiniBill.Columns(2).Value & "'"
                        End If
                    Loop
                End If
                If Len(Trim(TDBGMiniBill.Columns(12).Value)) = 0 Then
                    If Not .EOF Then
                        .Delete
                    End If
                Else
                    If .EOF Then
                        .AddNew
                        !line_id = TDBGMiniBill.Columns(0).Value
                        !model_number = TDBGMiniBill.Columns(1).Value
                        !part_id = TDBGMiniBill.Columns(2).Value
                        !part_sequence_number = TDBGMiniBill.Columns(3).Value
                        !step_number = TDBGMiniBill.Columns(13).Value
                    End If
                    !stocking_location_id = TDBGMiniBill.Columns(12).Value
                End If
            End With
            
            With mrsPartLocation
                If Len(Trim(TDBGMiniBill.Columns(12).Value)) > 0 Then
                    If .RecordCount > 0 Then
                        .MoveFirst
                        .Find "part_id = '" & TDBGMiniBill.Columns(2).Value & "'"
                    End If
                    Do While Not .EOF
                        If !line_id <> cboLine.Text Then
                            .MoveNext
                            If Not .EOF Then
                                .Find "part_id = '" & TDBGMiniBill.Columns(2).Value & "'"
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    If .EOF And TDBGMiniBill.Columns(11).Value Then
                        .AddNew
                        !line_id = TDBGMiniBill.Columns(0).Value
                        !part_id = TDBGMiniBill.Columns(2).Value
                    End If
                    If TDBGMiniBill.Columns(11).Value Then
                        !stocking_location_id = TDBGMiniBill.Columns(12).Value
                    ElseIf Not .EOF Then
                        If .EditMode = adEditAdd Then
                            .Delete
                        ElseIf .Fields("stocking_location_id").OriginalValue <> !stocking_location_id Then
                            !stocking_location_id = .Fields("stocking_location_id").OriginalValue
                        End If
                    End If
                    mblnRecChanged = True
                End If
            End With
        Case 13
            With mrsModelPartLocation
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Find "part_id = '" & TDBGMiniBill.Columns(2).Value & "'"
                    Do While Not .EOF
                        If !part_sequence_number = Val(TDBGMiniBill.Columns(3).Value) Then
                            Exit Do
                        End If
                        .MoveNext
                        If Not .EOF Then
                            .Find "part_id = '" & TDBGMiniBill.Columns(2).Value & "'"
                        End If
                    Loop
                    If .EOF Then
                        MsgBox "Step may not be set unless location has been set."
                        TDBGMiniBill.Columns(13) = Null
                        TDBGMiniBill.col = 12
                        Exit Sub
                    End If
                    !step_number = TDBGMiniBill.Columns(13).Value
                    mblnRecChanged = True
                Else
                    MsgBox "Step may not be set unless location has been set."
                    TDBGMiniBill.Columns(13) = Null
                    TDBGMiniBill.col = 12
                End If
            End With
        Case 14
            With mrsPartCategory
                If .RecordCount > 0 Then
                    .MoveFirst
                    .Find "part_id = '" & TDBGMiniBill.Columns(2).Value & "'"
                End If
                If Len(Trim(TDBGMiniBill.Columns(14).Value)) > 0 Then
                    If .EOF Then
                        .AddNew
                        !part_id = TDBGMiniBill.Columns(2).Value
                    End If
                    !Category_id = TDBGMiniBill.Columns(14).Value
                Else
                    If Not .EOF Then
                        .Delete
                    End If
                End If
                mblnRecChanged = True
            End With
        Case 11
            If TDBGMiniBill.Columns(11).Value = -1 Then
                With mrsPartLocation
                    If .RecordCount > 0 Then
                        .MoveFirst
                        .Find "part_id = '" & TDBGMiniBill.Columns(2).Value & "'"
                    End If
                    Do While Not .EOF
                        If !line_id <> cboLine.Text Then
                            .MoveNext
                            If Not .EOF Then
                                .Find "part_id = '" & TDBGMiniBill.Columns(2).Value & "'"
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    If .EOF Then
                        .AddNew
                        !part_id = TDBGMiniBill.Columns(2).Value
                        !line_id = TDBGMiniBill.Columns(0).Value
                    End If
                    !stocking_location_id = TDBGMiniBill.Columns(12).Value
                    mblnRecChanged = True
                End With
            Else
                With mrsPartLocation
                    If .RecordCount > 0 Then
                        .MoveFirst
                        .Find "part_id = '" & TDBGMiniBill.Columns(2).Value & "'"
                    End If
                    Do While Not .EOF
                        If !line_id <> cboLine.Text Then
                            .MoveNext
                            If Not .EOF Then
                                .Find "part_id = '" & TDBGMiniBill.Columns(2).Value & "'"
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    If Not .EOF Then
                        If .EditMode = adEditAdd Then
                            .Delete
                        Else
                            !stocking_location_id = .Fields("stocking_location_id").OriginalValue
                        End If
                        mblnRecChanged = True
                    End If
                End With
            End If
    End Select
    
End Sub

Private Sub LoadLines()
    ' Purpose:  Load lines
    
    ' Set up error handling.
    On Error GoTo PROC_ERR
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    
    cboLine.Clear
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select line_id from v_prod_line order by line_id"
        .Open
        
        Do While Not .EOF
            cboLine.AddItem !line_id
            .MoveNext
        Loop
        .Close
    End With
    
    Set rsList = Nothing
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "LoadLines", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub TDBGMiniBill_KeyDown(KeyCode As Integer, Shift As Integer)
    If TDBGMiniBill.col = 13 Then
        If KeyCode = vbKeyDelete Then
            TDBGMiniBill.Columns(13).Value = Null
            col = TDBGMiniBill.col
'            mrsModelPartLocation!step_number = TDBGMiniBill.Columns(13).Value
            mblnRecChanged = True
            Call TDBGMiniBill_AfterColUpdate(col)
        End If
    End If
End Sub
