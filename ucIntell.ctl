VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl ucIntell 
   BackColor       =   &H80000005&
   ClientHeight    =   2010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2160
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   144
   Begin MSComctlLib.ImageList imgLst 
      Left            =   1680
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucIntell.ctx":0000
            Key             =   "prop"
            Object.Tag             =   "prop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucIntell.ctx":015A
            Key             =   "sub"
            Object.Tag             =   "sub"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucIntell.ctx":06F4
            Key             =   "value"
            Object.Tag             =   "value"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvMembers 
      CausesValidation=   0   'False
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imgLst"
      SmallIcons      =   "imgLst"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "name"
         Object.Tag             =   "name"
         Text            =   "Name"
         Object.Width           =   1984
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "obj"
         Object.Tag             =   "obj"
         Text            =   "Obj"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "trigger"
         Object.Tag             =   "trigger"
         Text            =   "trigger"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "ucIntell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Private Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long


Private Declare Function SendMessageLong Lib "user32" Alias _
        "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function DrawText Lib "user32" Alias _
    "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
    ByVal nCount As Long, lpRect As RECT, ByVal wFormat _
    As Long) As Long
    
Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Const LVM_FIRST = &H1000
Private Const DT_CALCRECT = &H400

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'Border Styles for the control
Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_ADJUST = &H2000
Private Const BF_FLAT = &H4000
Private Const BF_TOP = &H2
Private Const BF_LEFT = &H1
Private Const BF_MIDDLE = &H800
Private Const BF_MONO = &H8000
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_SOFT = &H1000
Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL = &H10
Private Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Private Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Private Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Private Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_STATICEDGE = &H20000
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_DLGMODALFRAME = &H1&

Private Const HWND_TOP = 0


Public Enum BorderStyles
  bsbump
  bsEtched
  bsRaised
  bsSunken = EDGE_SUNKEN
End Enum

'Events

'The click and dblClick event now returns the selected item
Public Event Click(Item As Integer)
Public Event DblClick(Item As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OnError(Description As String, Source As String)

' member variable for ShadowColor property
Private m_ShadowColor As OLE_COLOR
' member variable for ListIndex property
Private m_ListIndex As Integer
' member variable for ListCount property
Private m_ListCount As Integer
Private m_W As Long
Private m_H As Long
' member variable for DataFile property
Private m_DataFile As String
Private m_XML As DOMDocument
Private m_DataLoaded As Boolean
' member variable for ForeColor property
Private m_ForeColor As OLE_COLOR
' member variable for BackColor property
Private m_BackColor As OLE_COLOR


Public Type MethodData
  MD_OBJECT As String
  MD_TRIGGER As String
  MD_NAME As String
End Type
' member variable for LastMethod property
Private m_LastMethod As MethodData
' member variable for FlatScrollBars property
Private m_FlatScrollBars As Boolean
' member variable for GridLines property
Private m_GridLines As Boolean
' member variable for HideSelection property
Private m_HideSelection As Boolean
' member variable for DisplayTotal property
Private m_DisplayTotal As Long
' member variable for HasShadow property
Private m_HasShadow As Boolean
' member variable for HotTracking property
Private m_HotTracking As Boolean
' member variable for HoverSelection property
Private m_HoverSelection As Boolean
' member variable for BorderStyle property
'Private m_BorderStyle As BorderStyles
' member variable for ListIcons property
' member variable for ListIcon property
Private m_ListIcon As ListItem
' member variable for ListIconCount property
Private m_ListIconCount As Long


' Returns The Icon for the Index

Property Get ListIcon(Index As Variant) As ListItem
Attribute ListIcon.VB_Description = "Returns The Icon for the Index"
  Set ListIcon = imgLst.ListImages(Index)
End Property

' Returns the number of Icons.

Property Get ListIconCount() As Long
Attribute ListIconCount.VB_Description = "Returns the number of Icons."
  ListIconCount = imgLst.ListImages.Count
End Property










' Sets / Returns the Lists border style

'Property Get BorderStyle() As BorderStyles
'  BorderStyle = m_BorderStyle
'End Property
'
'Property Let BorderStyle(ByVal newValue As BorderStyles)
'  m_BorderStyle = newValue
'  UserControl_Paint
'  PropertyChanged "BorderStyle"
'End Property




' Sets / Returns if HoverSelection is Enabled for the list.

Property Get HoverSelection() As Boolean
Attribute HoverSelection.VB_Description = "Sets / Returns if HoverSelection is Enabled for the list."
  HoverSelection = m_HoverSelection
End Property

Property Let HoverSelection(ByVal newValue As Boolean)
  m_HoverSelection = newValue
  UserControl_Paint
  PropertyChanged "HoverSelection"
End Property

Public Sub AddIcon(pIcon As IPictureDisp, Optional key As Variant)
  imgLst.ListImages.Add , key, pIcon
End Sub
Public Sub RemoveIcon(Index As Variant)
  imgLst.ListImages.Remove (Index)
End Sub


Public Sub RemoveItem(Index As Variant)
  lvMembers.ListItems.Remove (Index)
End Sub
'Manually add and item
Public Sub AddItem(Item As String, Optional Index As Integer, Optional key As Variant)
  On Error GoTo err_handler
  lvMembers.ListItems.Add Index, key, Item
  Exit Sub
err_handler:
  RaiseEvent OnError(Err.Description, UserControl.Name & "::AddSunkenBorder()")
  Resume Next
End Sub

Sub AddSunkenBorder(objX As Object)
  On Error GoTo err_handler
  Dim lstStyle As Long, m_BorderStyle As BorderStyles
  ' Get objects extended window style
  lstStyle = GetWindowLong(objX.hWnd, GWL_EXSTYLE)
  ' Append the border to the current extended window style
  lstStyle = lstStyle Or EDGE_BUMP
  ' Apply the change to the control
  Call SetWindowLong(objX.hWnd, GWL_EXSTYLE, lstStyle)
  Exit Sub
err_handler:
  RaiseEvent OnError(Err.Description, UserControl.Name & "::AddSunkenBorder()")
  Resume Next
End Sub

Private Sub AutoSizeColumn(LV As ListView, Optional Column _
 As ColumnHeader = Nothing)
 On Error GoTo err_handler
 Dim C As ColumnHeader
 If Column Is Nothing Then
  For Each C In LV.ColumnHeaders
   SendMessage LV.hWnd, LVM_FIRST + 30, C.Index - 1, -1
  Next
 Else
  SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
 End If
 LV.Refresh
 UserControl.Width = PixelsToTwipsX(LV.ColumnHeaders(1).Width + 26)
 Exit Sub
err_handler:
 RaiseEvent OnError(Err.Description, UserControl.Name & "::AutoSizeColumn()")
 Resume Next
End Sub

' Sets / Returns background color

Property Get BackColor() As OLE_COLOR
  BackColor = m_BackColor
End Property

Property Let BackColor(ByVal newValue As OLE_COLOR)
  m_BackColor = newValue
  lvMembers.BackColor = m_BackColor
  PropertyChanged "BackColor"
End Property

Public Sub Clear()
  lvMembers.ListItems.Clear
End Sub

' The DataFile property

Property Get DataFile() As String
  DataFile = m_DataFile
End Property

Property Let DataFile(ByVal newValue As String)
  m_DataFile = newValue
  PropertyChanged "FileName"
End Property

' The DataLoaded property

Property Get DataLoaded() As Boolean
  DataLoaded = m_DataLoaded
End Property


'**********Can't seem to get this to work*********
'***It will display the correct font in design mode
'***But once it's ran it goes back to the default font
'***If anyone can help with this please email me
'***witenite87@excite.com***

' member variable for Font property
'Private m_Font As Font

' Sets / Returns the Font of the List

'Property Get Font() As Font
'  Set Font = m_Font
'End Property
'
'Property Set Font(ByVal newValue As Font)
'  Set m_Font = newValue
'  Set lvMembers.Font = m_Font
'  Set UserControl.Font = m_Font
'  PropertyChanged "Font"
'End Property



' Sets / Returns the Max number of Items to display in the list.
Property Get DisplayTotal() As Long
Attribute DisplayTotal.VB_Description = "Sets / Returns the Max number of Items to display in the list."
  DisplayTotal = m_DisplayTotal
End Property

Property Let DisplayTotal(ByVal newValue As Long)
  m_DisplayTotal = newValue
  PropertyChanged "DisplayTotal"
End Property

' Returns / Sets if the list should have Flatscrollbars

Property Get FlatScrollBars() As Boolean
Attribute FlatScrollBars.VB_Description = "Returns / Sets if the list should have Flatscrollbars"
  FlatScrollBars = m_FlatScrollBars
End Property

Property Let FlatScrollBars(ByVal newValue As Boolean)
  m_FlatScrollBars = newValue
  UserControl_Paint
  PropertyChanged "FlatScrollBars"
End Property

' sets / returns the foreground color

Property Get ForeColor() As OLE_COLOR
  ForeColor = m_ForeColor
End Property

Property Let ForeColor(ByVal newValue As OLE_COLOR)
  m_ForeColor = newValue
  lvMembers.ForeColor = m_ForeColor
  PropertyChanged "ForeColor"
End Property

' Returns / Sets if the list has grid lines

Property Get GridLines() As Boolean
Attribute GridLines.VB_Description = "Returns / Sets if the list has grid lines"
  GridLines = m_GridLines
End Property

Property Let GridLines(ByVal newValue As Boolean)
  m_GridLines = newValue
  UserControl_Paint
  PropertyChanged "GridLines"
End Property

' Sets / Returns if the Controls has a shadow.

Property Get HasShadow() As Boolean
Attribute HasShadow.VB_Description = "Sets / Returns if the Controls has a shadow."
  HasShadow = m_HasShadow
End Property

Property Let HasShadow(ByVal newValue As Boolean)
  m_HasShadow = newValue
  If m_HasShadow Then
    UserControl.BackStyle = 1
  Else
    UserControl.BackStyle = 0 'Transparent
  End If
  PropertyChanged "HasShadow"
End Property

'Hides the popup
Public Function HidePopup() As Boolean
  On Error GoTo oops
  Dim R As RECT, hDWP As Long
  R.Left = 0
  R.Top = 0
  R.Bottom = 0
  R.Right = 0
  'Just makes the width and height 0
  
  hDWP = BeginDeferWindowPos(1)
  DeferWindowPos hDWP, hWnd, HWND_TOP, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, &H40
  EndDeferWindowPos hDWP
  'disables the list so it can't get focus
  lvMembers.Enabled = False
  HidePopup = True
  
  Exit Function
oops:
  RaiseEvent OnError(Err.Description, UserControl.Name & "::HidePopup()")
  HidePopup = False
End Function

' Returns /Sets if the list hides the selected item when focus is lost

Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Returns /Sets if the list hides the selected item when focus is lost"
  HideSelection = m_HideSelection
End Property

Property Let HideSelection(ByVal newValue As Boolean)
  m_HideSelection = newValue
  lvMembers.HideSelection = m_HideSelection
  PropertyChanged "HideSelection"
End Property

' Sets / Returns if list allows Hot Tracking.

Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Sets / Returns if list allows Hot Tracking."
  HotTracking = m_HotTracking
End Property

Property Let HotTracking(ByVal newValue As Boolean)
  m_HotTracking = newValue
  UserControl_Paint
  PropertyChanged "HotTracking"
End Property

' Returns information about the last Select Item

Property Get LastMethod() As MethodData
Attribute LastMethod.VB_Description = "Returns information about the last Select Item"
  LastMethod = m_LastMethod
End Property

Public Function List(Index As Integer) As String
  'return the selected list item
  List = lvMembers.ListItems(Index).Text
End Function

' The ListCount property

Property Get ListCount() As Integer
  m_ListCount = lvMembers.ListItems.Count
  ListCount = m_ListCount
End Property

' The ListIndex property

Property Get ListIndex() As Integer
  m_ListIndex = lvMembers.SelectedItem.Index
  ListIndex = m_ListIndex
End Property


Private Sub lvMembers_Click()
  On Error GoTo err_handler
  Dim Item As Integer
  Item = lvMembers.SelectedItem.Index
  'Store the Member
  m_LastMethod.MD_NAME = lvMembers.ListItems(Item).Text
  m_LastMethod.MD_OBJECT = lvMembers.ListItems(Item).SubItems(1)
  m_LastMethod.MD_TRIGGER = lvMembers.ListItems(Item).SubItems(2)
  'Debug.Print "OBJ = ", m_LastMethod.MD_OBJECT
  RaiseEvent Click(Item)
  Exit Sub
err_handler:
  RaiseEvent OnError(Err.Description, UserControl.Name & "::Click(" & Item & ")")
  Resume Next
End Sub

Private Sub lvMembers_DblClick()
  On Error GoTo err_handler
  Dim Item As Integer
  Item = lvMembers.SelectedItem.Index
  'Store the Member
  m_LastMethod.MD_NAME = lvMembers.ListItems(Item).Text
  m_LastMethod.MD_OBJECT = lvMembers.ListItems(Item).SubItems(1)
  m_LastMethod.MD_TRIGGER = lvMembers.ListItems(Item).SubItems(2)
  'Debug.Print "OBJ = ", m_LastMethod.MD_OBJECT
  RaiseEvent DblClick(Item)
  Exit Sub
err_handler:
  RaiseEvent OnError(Err.Description, UserControl.Name & "::DblClick(" & Item & ")")
  Resume Next
End Sub

Private Sub lvMembers_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub lvMembers_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lvMembers_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub lvMembers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub lvMembers_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub lvMembers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'Pixel in Twips X
Function PixelsToTwipsX(Pixels As Long) As Long
   On Error GoTo err_handler
   PixelsToTwipsX = Pixels * Screen.TwipsPerPixelX
   Exit Function
err_handler:
   RaiseEvent OnError(Err.Description, UserControl.Name & "::PixelsToTwipsX()")
   Resume Next
End Function

'Pixel in Twips Y
Function PixelsToTwipsY(Pixels As Long) As Long
   On Error GoTo err_handler
   PixelsToTwipsY = Pixels * Screen.TwipsPerPixelY
   Exit Function
err_handler:
   RaiseEvent OnError(Err.Description, UserControl.Name & "::PixelsToTwipsY()")
   Resume Next
End Function


'This is not used But I didn't want to remove it :)
'Private Function SelectNode(strName As String) As IXMLDOMNodeList
'  'On Error Resume Next
'  'this selects the Nodes
'  m_XML.setProperty "SelectionLanguage", "XPath"
'  Set SelectNode = m_XML.documentElement.selectSingleNode("/child::item[@name='" & strName & "']")
'End Function


Private Function SelectNodes(strName As String, strTrigger As String) As IXMLDOMNodeList
  On Error GoTo err_handler
  'this selects the Nodes
  m_XML.setProperty "SelectionLanguage", "XPath"
  
  Set SelectNodes = m_XML.documentElement.SelectNodes("child::obj[attribute::name='" & strName & "'][attribute::trigger='" & strTrigger & "']").Item(0).childNodes
  Exit Function
err_handler:
   RaiseEvent OnError(Err.Description, UserControl.Name & "::SelectNodes()")
   Resume Next
End Function

' sets / returns the shadow color
Property Get ShadowColor() As OLE_COLOR
Attribute ShadowColor.VB_Description = "sets / returns the shadow color"
  ShadowColor = m_ShadowColor
End Property

Property Let ShadowColor(ByVal newValue As OLE_COLOR)
  m_ShadowColor = newValue
  PropertyChanged ("ShadowColor")
  UserControl_Paint
End Property

Public Function ShowPopup(x As Long, y As Long, Optional ObjectRef As String = "", Optional strTrigger As String = ".") As Boolean
  On Error GoTo oops
  Dim imgIndex As Integer
  Dim R As RECT, hDWP As Long, retVal As Boolean
  Dim NodeList As IXMLDOMNodeList
  Dim Node As IXMLDOMNode
  Dim fso As New FileSystemObject
  Set fso = New FileSystemObject
  Dim LI As ListItem
  Dim xx As Long

  'check that we need to show it
  If Height = 0 And Width = 0 Then
    'enable the list
    lvMembers.Enabled = True
    'look for the data file
    If fso.FileExists(m_DataFile) Then
      'load the data if we need to
      If Not m_DataLoaded Then
        m_DataLoaded = m_XML.Load(m_DataFile)
        If m_DataLoaded = False Then MsgBox (m_XML.parseError.reason)
      End If
    Else
      m_DataLoaded = False
    End If


    'what are we loading?
    If ObjectRef <> "" And m_DataLoaded Then
      'get the nodes
      Set NodeList = SelectNodes(ObjectRef, strTrigger)
      lvMembers.ListItems.Clear
      For xx = 0 To NodeList.length - 1
        'populate the list
        On Error Resume Next
        imgIndex = imgLst.ListImages(NodeList(xx).Attributes(3).Text).Index
        If imgIndex <= 0 Then imgIndex = 1
        On Error GoTo oops
        Set LI = lvMembers.ListItems.Add(, , NodeList(xx).Text, imgIndex, imgIndex)
        LI.ListSubItems.Add , , NodeList(xx).Attributes(1).Text
        LI.ListSubItems.Add , , NodeList(xx).Attributes(2).Text
        DoEvents
      Next xx

      'there were no items
      If lvMembers.ListItems.Count = 0 Then HidePopup: ShowPopup = False: Exit Function
      'lstMembers.ListIndex = 0
      lvMembers.SetFocus

      
      'Set the position
      R.Left = x
      R.Top = y
      R.Bottom = TwipsToPixelsY(m_H)
      R.Right = TwipsToPixelsX(m_W)
      'Position the control at the X,Y position
      hDWP = BeginDeferWindowPos(1)
      DeferWindowPos hDWP, hWnd, HWND_TOP, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, &H40
      EndDeferWindowPos hDWP
      DoEvents
      
      'Set it to the correct height
      If lvMembers.ListItems.Count < m_DisplayTotal Then
        'The list has less then 10 items so lets set the size to fit the items
        UserControl.Height = PixelsToTwipsY(22 * lvMembers.ListItems.Count)
      Else
        'This will show m_DisplayTotal items
        UserControl.Height = PixelsToTwipsY(19 * m_DisplayTotal)
      End If

      'UserControl_Resize
    End If

    'This will resize the control and list to be wide enough for the longest item
    AutoSizeColumn lvMembers, lvMembers.ColumnHeaders("name")
    ShowPopup = True
  Else
oops:
    HidePopup  'incase it loaded for some reason
    ShowPopup = False
    RaiseEvent OnError(Err.Description, UserControl.Name & "::ShowPopup()")
  End If
End Function

'Twips in Pixel X
Function TwipsToPixelsX(Twips As Long) As Long
   On Error GoTo err_handler
   TwipsToPixelsX = Twips / Screen.TwipsPerPixelX
   Exit Function
err_handler:
   RaiseEvent OnError(Err.Description, UserControl.Name & "::TwipsToPixelsX()")
   Resume Next
End Function

'Twips in Pixel Y
Function TwipsToPixelsY(Twips As Long) As Long
   On Error GoTo err_handler
   TwipsToPixelsY = Twips / Screen.TwipsPerPixelY
   Exit Function
err_handler:
   RaiseEvent OnError(Err.Description, UserControl.Name & "::TwipsToPixelsY()")
   Resume Next
End Function

Private Sub UserControl_Initialize()
  On Error GoTo err_handler
  AddSunkenBorder lvMembers
  Dim fso As New FileSystemObject
  Set fso = New FileSystemObject
  Set m_XML = New DOMDocument
  m_DataFile = App.Path & "\defs.dat"
  If fso.FileExists(m_DataFile) Then
    m_DataLoaded = m_XML.Load(m_DataFile)
    If m_DataLoaded = False Then MsgBox (m_XML.parseError.reason)
  Else
    m_DataLoaded = False
  End If
  
  
  Set m_ListIcon = Nothing
  lvMembers.ListItems.Clear
  lvMembers.ListItems.Add , , "Intellisense", 2, 2
  'Set m_Font = UserControl.Font
  'Set lvMembers.Font = m_Font
  m_H = Height
  m_W = Width
  Exit Sub
err_handler:
  RaiseEvent OnError(Err.Description, UserControl.Name & "::Initialize()")
  Resume Next
End Sub

Private Sub UserControl_InitProperties()
  On Error GoTo err_handler
  m_ShadowColor = vbButtonShadow
  m_ListIndex = -1
  
  m_DataFile = App.Path & "\defs.dat"

  m_ForeColor = vbWindowText
  m_BackColor = vbWindowBackground
  m_FlatScrollBars = False
  m_GridLines = False
  m_HideSelection = False
  m_DisplayTotal = 7
  'Set m_Font = UserControl.Font
  'Set lvMembers.Font = m_Font
  m_HasShadow = True
  m_HotTracking = False
  m_HoverSelection = False
  Exit Sub
err_handler:
  RaiseEvent OnError(Err.Description, UserControl.Name & "::InitProperties()")
  Resume Next
  m_ListIconCount = 0
End Sub

Private Sub UserControl_Paint()
  On Error GoTo err_handler
  Dim R As RECT
  
  lvMembers.ForeColor = m_ForeColor
  lvMembers.BackColor = m_BackColor
  lvMembers.FlatScrollBar = m_FlatScrollBars
  lvMembers.GridLines = m_GridLines
  lvMembers.HotTracking = m_HotTracking
  lvMembers.HoverSelection = m_HoverSelection
  
  If m_HasShadow Then
    UserControl.BackStyle = 1
    Cls
    Line (ScaleWidth - 3, 2)-(ScaleWidth - 3, ScaleHeight - 2), m_ShadowColor
    Line (ScaleWidth - 2, 2)-(ScaleWidth - 2, ScaleHeight - 2), m_ShadowColor
    Line (ScaleWidth - 1, 2)-(ScaleWidth - 1, ScaleHeight - 2), m_ShadowColor
    Line (ScaleWidth, 2)-(ScaleWidth, ScaleHeight - 2), m_ShadowColor
    
    Line (2, ScaleHeight - 3)-(ScaleWidth, ScaleHeight - 3), m_ShadowColor
    Line (2, ScaleHeight - 2)-(ScaleWidth, ScaleHeight - 2), m_ShadowColor
    Line (2, ScaleHeight - 1)-(ScaleWidth, ScaleHeight - 1), m_ShadowColor
    Line (2, ScaleHeight)-(ScaleWidth, ScaleHeight), m_ShadowColor
  Else
    UserControl.BackStyle = 0 'Transparent
  End If
  Exit Sub
err_handler:
  RaiseEvent OnError(Err.Description, UserControl.Name & "::Paint()")
  Resume Next
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  On Error GoTo err_handler
  With PropBag
    m_ShadowColor = .ReadProperty("ShadowColor", vbButtonShadow)
    m_DataFile = .ReadProperty("FileName", App.Path & "\defs.dat")
    m_FlatScrollBars = .ReadProperty("FlatScrollBars", False)
    m_ForeColor = .ReadProperty("ForeColor", vbWindowText)
    m_BackColor = .ReadProperty("BackColor", vbWindowBackground)
    m_GridLines = .ReadProperty("GridLines", False)
    m_HideSelection = .ReadProperty("HideSelection", False)
    m_DisplayTotal = .ReadProperty("DisplayTotal", 7)
    m_HasShadow = .ReadProperty("HasShadow", True)
    m_HotTracking = .ReadProperty("HotTracking", False)
    m_HoverSelection = .ReadProperty("HoverSelection", False)
    'm_BorderStyle = .ReadProperty("BorderStyle", bsbump)
    'Set m_Font = .ReadProperty("Font", UserControl.Font)
  End With
  Exit Sub
err_handler:
  RaiseEvent OnError(Err.Description, UserControl.Name & "::ReadProperties()")
  Resume Next
End Sub

Private Sub UserControl_Resize()
  On Error GoTo err_handler
  lvMembers.Move 0, 0, ScaleWidth - 3, ScaleHeight - 3
  Exit Sub
err_handler:
  RaiseEvent OnError(Err.Description, UserControl.Name & "::Resize()")
  Resume Next
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  On Error GoTo err_handler
  With PropBag
    .WriteProperty "ShadowColor", m_ShadowColor, vbButtonShadow
    .WriteProperty "FileName", m_DataFile, App.Path & "\defs.dat"
    .WriteProperty "FlatScrollBars", m_FlatScrollBars, False
    .WriteProperty "ForeColor", m_ForeColor, vbWindowText
    .WriteProperty "BackColor", m_BackColor, vbWindowBackground
    .WriteProperty "GridLines", m_GridLines, False
    .WriteProperty "HideSelection", m_HideSelection, False
    .WriteProperty "DisplayTotal", m_DisplayTotal, 7
    .WriteProperty "HasShadow", m_HasShadow, True
    .WriteProperty "HotTracking", m_HotTracking, False
    .WriteProperty "HoverSelection", m_HoverSelection, False
    '.WriteProperty "BorderStyle", m_BorderStyle, bsbump
    '.WriteProperty "Font", m_Font, UserControl.Font
  End With
  Exit Sub
err_handler:
  RaiseEvent OnError(Err.Description, UserControl.Name & "::WriteProperties()")
  Resume Next
End Sub
