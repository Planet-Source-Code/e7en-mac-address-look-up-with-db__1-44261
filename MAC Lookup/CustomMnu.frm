VERSION 5.00
Begin VB.Form CustomMnu 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   1080
   End
   Begin VB.Image Normal 
      Height          =   270
      Left            =   240
      Picture         =   "CustomMnu.frx":0000
      Top             =   720
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Hover 
      Height          =   270
      Left            =   240
      Picture         =   "CustomMnu.frx":155A
      Top             =   480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image BarPic 
      Height          =   60
      Left            =   240
      Picture         =   "CustomMnu.frx":2AB4
      Top             =   960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image TopPic 
      Height          =   60
      Left            =   240
      Picture         =   "CustomMnu.frx":2FA6
      Top             =   360
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Img 
      Height          =   255
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Cap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   480
   End
   Begin VB.Image BottomBar 
      Height          =   30
      Left            =   240
      Picture         =   "CustomMnu.frx":3498
      Top             =   1680
      Width           =   1500
   End
End
Attribute VB_Name = "CustomMnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================================'
'| Created By: §e7eN                                                  |'
'| Description: This will allow you to make custom menus              |'
'|              useing Images.                                        |'
'|                                                                    |'
'|                                                                    |'
'| Contact: hate_114@hotmail.com                                      |'
'|                                                                    |'
'| *If you wish to use this in one of your Programs please E-mail me* |'
'======================================================================


'Returns Mouse Position
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Returns windows Dimentions
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Returns hWnd of selected window
Private Declare Function GetActiveWindow Lib "user32" () As Long
'Changes Window Location and Dimentions
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'Used to Return/Set the Dimentions of a window
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'Used to Return/Set the Co-ordinates of the mouse
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Dim ImgCount As Integer


'------------------
'Createing Pictures
'------------------
'
Private Sub Form_Load()
'Set our Bar Image up top
With Img(0)
    .Top = 0
    .Left = 0
    .Picture = TopPic.Picture
    .Visible = True
    .Tag = "-"
End With
End Sub


Public Sub CreateMenuButton(Text As String)
'Count up Next Image
ImgCount = ImgCount + 1

If Text = "-" Then 'If Image is a bar then

Load Img(ImgCount) 'Load New Image Control

'Position Image and set image
With Img(ImgCount)
    .Picture = BarPic.Picture
    .Top = Img(ImgCount - 1).Top + Img(ImgCount - 1).Height
    .Left = 0
    .Visible = True
    .Tag = "-"
End With

Else

'Load New Image Control and Caption
Load Img(ImgCount)
Load Cap(ImgCount)

'Position Image and set image
With Img(ImgCount)
    .Picture = Normal.Picture
    .Top = Img(ImgCount - 1).Top + Img(ImgCount - 1).Height
    .Left = 0
    .Visible = True
    .Tag = ""
End With

'Position and set the caption
With Cap(ImgCount)
    .Caption = Text
    .Top = Img(ImgCount).Top + (Img(ImgCount).Height / 2) - 100
    .Left = 100
    .ZOrder
    .Visible = True
End With

End If

With BottomBar
    .Top = Img(ImgCount).Top + Img(ImgCount).Height
    .Left = 0
    .Tag = "-"
End With

'Resize the form to adjust for the new menu added
Me.Height = Img(ImgCount).Top + Img(ImgCount).Height + BottomBar.Height
Me.Width = Img(ImgCount).Width
End Sub

'------------------
'Change Pictures on Mouse Events
'------------------
'
Private Sub Cap_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If mouse is over a button then change the picture
    Rollover Index
End Sub

Private Sub Img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Rollover Index
End Sub


Private Sub Img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If button clicked change picture to normal
    If Img(Index).Tag = "-" Then Else Img(Index).Picture = Normal.Picture
End Sub
'------------------
'When Buttons Clicked
'------------------
'
Private Sub Img_Click(Index As Integer)
    'Return the Button click
    ButtonClick Index
End Sub

Private Sub Cap_Click(Index As Integer)
    ButtonClick Index
End Sub

Sub ButtonClick(Index As Integer)
Dim sMAC As String

Me.Visible = False

Select Case LCase(Cap(Index).Caption)
    Case "lookup"
        If Main.txtIP.Text = "" Then Exit Sub
        
        sMAC = GetMAC(Main.txtIP.Text)
        Main.txtMAC.Text = sMAC
        Main.txtVendor.Text = MAClookup(sMAC)
    Case "about"
        MsgBox "This program was brought to you today by the number 5 and the letter Z." & vbCrLf & "Author: §e7eN (Jake Paternoster)" & vbCrLf & "Contact: hate_114@hotmail.com", vbApplicationModal + vbInformation + vbOKOnly, Main.Caption
    Case "exit"
        End
    Case Else
        MsgBox Cap(Index).Caption & " was clicked."
End Select
End Sub

'------------------
'Other stuff
'------------------
'
Private Sub Timer1_Timer()
Dim MouseCurPos As POINTAPI
Dim WinRECT As RECT

'Get Mouse Co-ordinates
Call GetCursorPos(MouseCurPos)
'Get Window Dimentions
Call GetWindowRect(Me.hwnd, WinRECT)

'If the menu is NOT active (Selected) window then Hide
If GetActiveWindow() <> Me.hwnd Then Me.Visible = False

'If the mouse is not over the menu then restore the button pictures to normal
If MouseCurPos.X > WinRECT.Left And MouseCurPos.X < WinRECT.Right _
And MouseCurPos.Y > WinRECT.Top And MouseCurPos.Y < WinRECT.Bottom Then

Else

For X = 1 To Img.Count - 1
    If Img(X).Tag = "-" Or Img(X).Picture = Normal.Picture Then Else Img(X).Picture = Normal.Picture
Next

End If
End Sub

Public Sub ShowMenu()
Dim MouseCurPos As POINTAPI
Dim WinRECT As RECT

'Get mouse Co-ordinates
Call GetCursorPos(MouseCurPos)
'Get window Dimentions
Call GetWindowRect(Me.hwnd, WinRECT)
'Move the menu to the mouses location
Call MoveWindow(Me.hwnd, MouseCurPos.X, MouseCurPos.Y, WinRECT.Right - WinRECT.Left, WinRECT.Bottom - WinRECT.Top, True)

'Show the menu
Me.Visible = True
End Sub

Sub Rollover(Index As Integer)
'Set all menu button pictures to normal state
For X = 0 To ImgCount
    If Img(X).Tag = "-" Then Else Img(X).Picture = Normal.Picture
Next

'Change only the picture that is being Hovered
    If Img(Index).Tag = "-" Then Else Img(Index).Picture = Hover.Picture
    
End Sub


