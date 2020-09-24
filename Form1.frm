VERSION 5.00
Object = "*\AprjIntell.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "chip"
            Object.Tag             =   "chip"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":27B2
            Key             =   "win"
            Object.Tag             =   "win"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3384
            Key             =   "unknown"
            Object.Tag             =   "unknown"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":34DE
            Key             =   "pad"
            Object.Tag             =   "pad"
         EndProperty
      EndProperty
   End
   Begin prjIntell.ucIntell ucIntell1 
      Height          =   1935
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3413
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetCaretPos Lib "user32" _
    (lpPoint As POINTAPI) As Long
Private Type POINTAPI
  X As Long
  Y As Long
End Type
Dim mdData As MethodData


Private Sub Form_Load()
  Dim X As Long
  For X = 1 To ImageList1.ListImages.Count
    'Add icons from our image list to the Intellisense Control
    ucIntell1.AddIcon ImageList1.ListImages(X).Picture, ImageList1.ListImages(X).Key
    DoEvents
  Next X
  
  Caption = "Image List has " & ucIntell1.ListIconCount & " Images"
  ucIntell1.HidePopup
End Sub


Private Sub Form_Resize()
  On Error Resume Next
  Text1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  Dim pXY As POINTAPI
  Dim strT As String
  
  Select Case KeyAscii
  Case 46
    GetCaretPos pXY
    If Len(Text1.Text) > 1 Then
      
      
      If InStrRev(Text1, vbCrLf) Then
        'We got a crlf
        strT = Mid$(Text1.Text, InStrRev(Text1, vbCrLf) + 2, Text1.SelStart)
      ElseIf InStrRev(Text1, " ") Then
        'we got a space to go by
        strT = Mid$(Text1.Text, InStrRev(Text1, " ") + 1, Text1.SelStart)
      ElseIf InStrRev(Text1, ".") Then
        'We got a "."
        strT = Mid$(Text1.Text, InStrRev(Text1, ".") + 1, Text1.SelStart)
      Else
        'we should be on the first line
        strT = Mid$(Text1.Text, 1, Text1.SelStart)
      End If
      ucIntell1.ShowPopup pXY.X + 10, pXY.Y + 10, strT, "."
    End If

  Case Asc("=")
    GetCaretPos pXY
    If Len(Text1.Text) > 1 Then
    
    ''''Test Multiple Obj
    If mdData.MD_OBJECT <> "" And mdData.MD_TRIGGER = "=" Then
      strT = mdData.MD_OBJECT
      ucIntell1.ShowPopup pXY.X + 10, pXY.Y + 10, Trim$(strT), "="
      Exit Sub
    End If
    
      If InStrRev(Text1, " ") Then
        'we got a space to go by
        strT = Mid$(Text1.Text, InStrRev(Text1, " ", Text1.SelStart - 2) + 1, Text1.SelStart)
      ElseIf InStrRev(Text1, vbCrLf) Then
        'We got a crlf
        strT = Mid$(Text1.Text, InStrRev(Text1, vbCrLf) + 2, Text1.SelStart)
      Else
        'we should be on the first line
        strT = Mid$(Text1.Text, 1, Text1.SelStart)
      End If
      ucIntell1.ShowPopup pXY.X + 10, pXY.Y + 10, Trim$(strT), "="
    End If
  Case 8
    ucIntell1.HidePopup
  End Select


End Sub

Private Sub ucIntell1_DblClick(Item As Integer)
  If Mid$(Text1, Text1.SelStart, 1) = "=" Then
    Text1.SelText = " " & ucIntell1.List(Item)
  Else
    Text1.SelText = ucIntell1.List(Item)
  End If
  Text1.SetFocus
  mdData = ucIntell1.LastMethod
  ucIntell1.HidePopup
End Sub

Private Sub ucIntell1_KeyPress(KeyAscii As Integer)
Dim intTemp As Integer
  Select Case KeyAscii
    Case Asc(".")
      ucIntell1_DblClick ucIntell1.ListIndex
      Text1.SelText = "."
      Text1_KeyPress (Asc("."))
      Text1.SetFocus
    Case Asc("=")
      ucIntell1_DblClick ucIntell1.ListIndex
      Text1.SelText = " ="
      Text1_KeyPress (Asc("="))
      Text1.SetFocus
    Case 13
      ucIntell1_DblClick ucIntell1.ListIndex
      Text1.SelText = vbCrLf
    Case Asc(" ")
      ucIntell1_DblClick ucIntell1.ListIndex
      Text1.SelText = " "
      Text1.SetFocus
    Case 8 'Backspace
      ucIntell1.HidePopup
      intTemp = Text1.SelStart - 1
      Debug.Print Text1.SelStart
      Text1 = Mid$(Text1, 1, intTemp)
      Text1.SelStart = intTemp
      Debug.Print Text1.SelStart
      Text1.SetFocus
  End Select
  
End Sub
