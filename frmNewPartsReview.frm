VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmNewPartsReview 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Minbill - New Parts Review"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   10095
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRefreshParts 
      Caption         =   "Refresh Parts"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   1695
   End
   Begin VB.ComboBox cboLine 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   2835
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   435
      Left            =   6720
      TabIndex        =   3
      Top             =   5280
      Width           =   1635
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   8580
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdDisplayModels 
      Caption         =   "Display Models..."
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   5280
      Width           =   1635
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGNewPartsInfo 
      Height          =   4335
      Left            =   180
      TabIndex        =   2
      Top             =   720
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   7646
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   16
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Part Number"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Description"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Create Date"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "Short Date"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   1
      Columns(3)._MaxComboItems=   50
      Columns(3).Caption=   "Category"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   4
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Reviewed?"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2170"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2090"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=5583"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5503"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2117"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2037"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=3254"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Button=1"
      Splits(0)._ColumnProps(24)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=2223"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2143"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=1"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
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
End
Attribute VB_Name = "frmNewPartsReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsModelPart As ADODB.Recordset
Private mrsPartCategory As ADODB.Recordset

Private mxarrPartInfo As XArrayDB
Private mxarrCategories As XArrayDB
Private mblnRecChanged As Boolean
Private mblnSaved As Boolean
Private mblnNoCategories As Boolean
Private arrLines() As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDisplayModels_Click()
    Dim intRetCode As Integer
    
    If TDBGNewPartsInfo.Row < 0 Then
        MsgBox "Please select a part before requesting Model Information."
        Exit Sub
    End If
    
    If mblnRecChanged Then
        intRetCode = MsgBox("Save Changes?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Data Changed")
        If intRetCode = vbNo Then
            mrsModelPart.CancelBatch
            If Not mblnNoCategories Then
                mrsPartCategory.CancelBatch
            End If
        ElseIf intRetCode = vbCancel Then
            Exit Sub
        Else
            cmdSave_Click
            If Not mblnSaved Then
                Exit Sub
            End If
        End If
    End If
    
    With frmModelsByPart
        .mstrPart = TDBGNewPartsInfo.Columns(0).Value
        .mstrPartDescription = TDBGNewPartsInfo.Columns(1).Value
        .DisplayModelsForPart
        .Show vbModal
    End With
End Sub

Private Sub cmdRefreshParts_Click()
    'This button click will refresh the list of parts with only those from the selected Line.  If All Lines
    ' are selected, then will show the entire list.  Passes the line number using a global variable.
    ' This subroutine was built from the Form_Load processing with modifications.
    Dim intRetCode As Integer
    Dim rsData As ADODB.Recordset
    Dim intIndex As Integer
    Dim Item As TrueOleDBGrid70.ValueItem
    Set Item = New TrueOleDBGrid70.ValueItem
        
    If mblnRecChanged Then
        intRetCode = MsgBox("Save Changes?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Data Changed")
        If intRetCode = vbNo Then
            mrsModelPart.CancelBatch
            If Not mblnNoCategories Then
                mrsPartCategory.CancelBatch
            End If
        ElseIf intRetCode = vbCancel Then
            Exit Sub
        Else
            cmdSave_Click
            If Not mblnSaved Then
                Exit Sub
            End If
        End If
    End If
    
    gblnLine = arrLines(cboLine.ListIndex)
    
        ' Set up recordset to hold data for grid and retrieve.
    Set rsData = New ADODB.Recordset
    Set mxarrPartInfo = New XArrayDB
    With rsData
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
'        Temporarily changing this query to return all new parts
        .Source = "select v_mnb_model_part.part_id, min(part_description) as part_description, " & _
            "max(part_create_date) as Create_date, min(category_id) as Category, 0 as Reviewed " & _
            "from v_mnb_model_part " & _
            "join v_prod_part on v_mnb_model_part.part_id = v_prod_part.part_id " & _
            "left outer join v_mnb_category_part on v_mnb_model_part.part_id = v_mnb_category_part.part_id " & _
            "where part_reviewed_flag = 0 and part_create_date is not null and v_mnb_model_part.line_id like '" & _
            gblnLine & "' group by v_mnb_model_part.part_id " & _
            "order by v_mnb_model_part.part_id"
        .Open

        If Not .EOF Then
            mxarrPartInfo.LoadRows .GetRows(), True
        Else
            .Close
            MsgBox "No new parts are available to be reviewed."
'            Unload frmProcessing
'            Unload Me
            Exit Sub
        End If
        .Close
    End With
    Set rsData = Nothing
    
'    LoadCategories
    
    With TDBGNewPartsInfo
        .Array = mxarrPartInfo
        With .Columns(3)
            If mblnNoCategories Then
                .AllowFocus = False
            Else
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
            End If
        End With
    .ReBind
    End With
    
    If Not mblnNoCategories Then
        Set mrsPartCategory = New ADODB.Recordset
        
        With mrsPartCategory
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Source = "select * from v_mnb_category_part order by part_id"
            .Open
        End With
    End If
    
    mblnRecChanged = False
    
    Set mrsModelPart = New ADODB.Recordset
    
    With mrsModelPart
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        'Temporarily changing this query to return all new parts
        .Source = "select * from v_mnb_model_part " & _
            "where part_create_date is not null and part_reviewed_flag = 0 " & _
            "order by part_id"
        .Open
    End With
    mblnRecChanged = False
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdRefreshParts_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub Form_Load()
    ' Purpose:  Build the data for the screen
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim rsData As ADODB.Recordset
    Dim lngIndex As Long
    Dim lngMaxCount As Long
    Dim strPrevModel As String
    Dim strPrevLine As String
    Dim strPrevECN As String
    Dim intIndex As Integer
    Dim Item As TrueOleDBGrid70.ValueItem
    Set Item = New TrueOleDBGrid70.ValueItem
    
    'Option to either review new parts by line or review all new parts.
    
    Call BuildLines
    
    
    'End Option
    
    gconDatabase.CommandTimeout = 350
    
    frmProcessing.Label1 = "Gathering Parts... Please Wait"
    frmProcessing.Label2 = "Get Configuration Parts List at " & Time
    frmProcessing.Show
    
    DoEvents
    
    ' Open connection if it is closed.
    If gconDatabase Is Nothing Then
        Set gconDatabase = gclsSQLServer.Connect( _
            gclsMESApplication.ApplicationRole, _
            gclsMESApplication.ApplicationPassword)
        With gconDatabase
            If .State <> adStateOpen Then
                Err.Raise .Errors(0).NativeError, "Form_Load", .Errors(0).Description
            End If
        End With
    End If
    
    ' Set up recordset to hold data for grid and retrieve.
    Set rsData = New ADODB.Recordset
    Set mxarrPartInfo = New XArrayDB
    With rsData
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
'        Temporarily changing this query to return all new parts
        .Source = "select v_mnb_model_part.part_id, min(part_description) as part_description, " & _
            "max(part_create_date) as Create_date, min(category_id) as Category, 0 as Reviewed " & _
            "from v_mnb_model_part " & _
            "join v_prod_part on v_mnb_model_part.part_id = v_prod_part.part_id " & _
            "left outer join v_mnb_category_part on v_mnb_model_part.part_id = v_mnb_category_part.part_id " & _
            "where part_reviewed_flag = 0 and part_create_date is not null and v_mnb_model_part.line_id like '" & _
            arrLines(cboLine.ListIndex) & "' group by v_mnb_model_part.part_id " & _
            "order by v_mnb_model_part.part_id"
        .Open

        If Not .EOF Then
            mxarrPartInfo.LoadRows .GetRows(), True
        Else
            .Close
            MsgBox "No new parts are available to be reviewed."
            Unload frmProcessing
            Unload Me
            Exit Sub
        End If
        .Close
    End With
    Set rsData = Nothing
    
    LoadCategories
    
    With TDBGNewPartsInfo
        .Array = mxarrPartInfo
        With .Columns(3)
            If mblnNoCategories Then
                .AllowFocus = False
            Else
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
            End If
        End With
    .ReBind
    End With
    
    If Not mblnNoCategories Then
        Set mrsPartCategory = New ADODB.Recordset
        
        With mrsPartCategory
            Set .ActiveConnection = gconDatabase
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Source = "select * from v_mnb_category_part order by part_id"
            .Open
        End With
    End If
    
    mblnRecChanged = False
    
    Set mrsModelPart = New ADODB.Recordset
    
    With mrsModelPart
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        'Temporarily changing this query to return all new parts
        .Source = "select * from v_mnb_model_part " & _
            "where part_create_date is not null and part_reviewed_flag = 0 " & _
            "order by part_id"
'        .Source = "select * from v_mnb_model_part " & _
'            "where part_reviewed_flag = 0 " & _
'            "order by part_id"
        .Open
    End With
    mblnRecChanged = False
    
    frmProcessing.Hide
    
    DoEvents
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "Form_Load", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim intRetCode As Integer
    
    If mblnRecChanged Then
        intRetCode = MsgBox("Save Changes?", vbQuestion + vbYesNoCancel + vbDefaultButton1, "Data Changed")
        If intRetCode = vbYes Then
            cmdSave_Click
        ElseIf intRetCode = vbNo Then
            If Not mblnNoCategories Then
                mrsPartCategory.CancelBatch
            End If
            mrsModelPart.CancelBatch
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gconDatabase.State = adStateOpen Then
        If Not mrsPartCategory Is Nothing Then
            If mrsPartCategory.State = adStateOpen Then
                mrsPartCategory.Close
            End If
            Set mrsPartCategory = Nothing
        End If
        If Not mrsModelPart Is Nothing Then
            If mrsModelPart.State = adStateOpen Then
                mrsModelPart.Close
            End If
            Set mrsModelPart = Nothing
        End If
    End If
    Set mxarrPartInfo = Nothing
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
        
        If .RecordCount = 0 Then
            mblnNoCategories = True
        Else
            mblnNoCategories = False
            mxarrCategories.ReDim 1, .RecordCount, 0, 1
            
            intIndex = 1
            Do While Not .EOF
                mxarrCategories(intIndex, 0) = .Fields("category_id").Value
                mxarrCategories(intIndex, 1) = Trim(!Category_description)
                intIndex = intIndex + 1
                .MoveNext
            Loop
        End If
        .Close
    End With
    Set rsList = Nothing
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "FindDescription", _
        Err.Number, Err.Description)
End Sub




Private Sub TDBGNewPartsInfo_AfterColUpdate(ByVal ColIndex As Integer)
    mblnRecChanged = True
    Select Case ColIndex
'        Case 3
'            With mrsPartCategory
'                If .RecordCount > 0 Then
'                    .MoveFirst
'                    .Find "part_id = '" & TDBGNewPartsInfo.Columns(0).Value & "'"
'                End If
'                If Len(Trim(TDBGNewPartsInfo.Columns(3).Value)) > 0 Then
'                    If .EOF Then
'                        .AddNew
'                        !part_id = TDBGNewPartsInfo.Columns(0).Value
'                    End If
'                    !Category_id = TDBGNewPartsInfo.Columns(3).Value
'                Else
'                    If Not .EOF Then
'                        .Delete
'                    End If
'                End If
'                mblnRecChanged = True
'            End With
        Case 4
            With mrsModelPart
                .MoveFirst
                .Find "part_id = '" & TDBGNewPartsInfo.Columns(0).Value & "'"
                Do While Not .EOF
                    If TDBGNewPartsInfo.Columns(4).Value = 0 Then
                        !part_reviewed_flag = 0
                    Else
                        !part_reviewed_flag = 1
                    End If
                    .MoveNext
                    If Not .EOF Then
                        If Trim(!part_id) <> Trim(TDBGNewPartsInfo.Columns(0).Value) Then
                            Exit Do
                        End If
                    End If
                Loop
            End With
        Case Else
    End Select
End Sub

Private Sub cmdSave_Click()
    ' Purpose:  Save current changes.
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    mblnSaved = False
    
    With gconDatabase
        mrsModelPart.UpdateBatch
        If .Errors.Count > 0 Then
            Err.Raise .Errors(0).NativeError, "cmdSave_click", .Errors(0).Description
        End If
        
        If Not mblnNoCategories Then
            mrsPartCategory.UpdateBatch
            If .Errors.Count > 0 Then
                Err.Raise .Errors(0).NativeError, "cmdSave_click", .Errors(0).Description
            End If
        End If
    End With
    
    MsgBox "Records successfully saved."
    mblnRecChanged = False
    mblnSaved = True

PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    Call ShowError(Me.Name, "cmdSave_Click", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub


Private Sub TDBGNewPartsInfo_DblClick()
    cmdDisplayModels_Click
End Sub
Public Sub BuildLines()
    ' Purpose:  Generate list of Assembly Lines used by Minibill
    
    ' Set up error handling
    On Error GoTo PROC_ERR
    
    Dim intIndex As Integer
    Dim intCtr As Integer
    
    
    Dim rsList As ADODB.Recordset
    Set rsList = New ADODB.Recordset
    With rsList
        Set .ActiveConnection = gconDatabase
        .CursorLocation = adUseClient
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Source = "select '%' as line_id, '--All Lines--' as line_description union " & _
                    "SELECT distinct lsl.line_id,line_description FROM V_PROD_Line_Stocking_Location lsl " & _
                    "join V_PROD_line line on lsl.line_id = line.line_id order by lsl.line_id"
        .Open

        intCtr = 0
        ReDim arrLines(.RecordCount - 1)
        cboLine.Clear
        Do While Not .EOF
            cboLine.AddItem !line_description
            arrLines(intCtr) = !line_id
            intCtr = intCtr + 1
            .MoveNext
        Loop
        If cboLine.ListCount > 0 Then
            cboLine.ListIndex = 0
        End If
        .Close
      
    End With

'    mblnRecChanged = False
    
PROC_EXIT:
            
    Exit Sub
    
PROC_ERR:

    Call ShowError(Me.Name, "BuildLines", Err.Number, Err.Description)
    GoTo PROC_EXIT
End Sub

