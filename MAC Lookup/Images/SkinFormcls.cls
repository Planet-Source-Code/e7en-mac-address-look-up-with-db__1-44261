VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SkinFormcls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dim Time As Timer

'----------------------
'/Define Title Bar Controls
'----------------------
Dim TitleBarLeft As Image
Dim TitleBarMiddle As Image
Dim TitleBarRight As Image

'----------------------
'/Define Title Bar Focused Pictures
'----------------------
Dim TitleBarLeftFocusY As IPictureDisp
Dim TitleBarMiddleFocusY As IPictureDisp
Dim TitleBarRightFocusY  As IPictureDisp

'----------------------
'/Define Title Bar Not Focused Pictures
'----------------------
Dim TitleBarLeftFocusN As IPictureDisp
Dim TitleBarMiddleFocusN As IPictureDisp
Dim TitleBarRightFocusN  As IPictureDisp

'----------------------
'/Define Title Bar Title Values
'----------------------
Dim TitleOffsetX As Integer
Dim TitleOffsetY As Integer
Dim TitleFont As IFontDisp
Public Focused As Boolean

'----------------------
'/Define Other
'----------------------
Public SkinedForm As Form

Sub CreateImageControls()
'----------------------
'/Set Title Bar Focused Pictures
'----------------------
Set TitleBarLeft = SkinedForm.Controls.Add("VB.Image", "TitleBarLeft", SkinedForm)
TitleBarLeft.Visible = True

Set TitleBarMiddle = SkinedForm.Controls.Add("VB.Image", "TitleBarMiddle", SkinedForm)
TitleBarMiddle.Visible = True

Set TitleBarRight = SkinedForm.Controls.Add("VB.Image", "TitleBarRight", SkinedForm)
TitleBarRight.Visible = True

Call SetWindowLong(SkinedForm.hwnd, GWL_STYLE, GetWindowLong(SkinedForm.hwnd, GWL_STYLE) + WS_DLGFRAME)
'Set Timer = SkinedForm.Controls.Add("VB.Timer", "Timer", SkinedForm)
End Sub

Sub SetForm(F As Form)
Set SkinedForm = F
End Sub


Sub ResizeImages()
'----------------------
'/Resize Title Bar
'----------------------
With TitleBarLeft
    .Left = 0
    .Top = 0
End With

With TitleBarRight
    .Left = SkinedForm.Width - TitleBarRight.Width - 100
    .Top = 0
End With

With TitleBarMiddle
    .Left = TitleBarLeft.Width
    .Top = 0
    .Width = SkinedForm.Width - (TitleBarLeft.Width + TitleBarRight.Width + 100)
End With

With SkinedForm.Drag
    .Left = 0
    .Top = 0
    .Height = TitleBarLeft.Height
    .Width = SkinedForm.Width
End With

End Sub

Sub FormFocused()
'----------------------
'/Set Title Bar Focused Pictures
'----------------------
Focused = True
ImageStrech False
Set TitleBarLeft.Picture = TitleBarLeftFocusY
Set TitleBarMiddle.Picture = TitleBarMiddleFocusY
Set TitleBarRight.Picture = TitleBarRightFocusY
ImageStrech True
ResizeImages
End Sub

Sub FormNotFocused()
'----------------------
'/Set Title Bar Not Focused Pictures
'----------------------
Focused = False
ImageStrech False
Set TitleBarLeft.Picture = TitleBarLeftFocusN
Set TitleBarMiddle.Picture = TitleBarMiddleFocusN
Set TitleBarRight.Picture = TitleBarRightFocusN
ImageStrech True
ResizeImages
End Sub

Sub LoadPictures()
Dim Path As String

Path = App.Path & "\Images"

Set TitleBarLeftFocusY = LoadPicture(Path & "\TBFRY.bmp")
Set TitleBarMiddleFocusY = LoadPicture(Path & "\TBFMY.bmp")
Set TitleBarRightFocusY = LoadPicture(Path & "\TBFLY.bmp")

Set TitleBarLeftFocusN = LoadPicture(Path & "\TBFLN.bmp")
Set TitleBarMiddleFocusN = LoadPicture(Path & "\TBFMN.bmp")
Set TitleBarRightFocusN = LoadPicture(Path & "\TBFRN.bmp")
End Sub

Sub FocusCheck()
End Sub

Sub ImageStrech(State As Boolean)
TitleBarLeft.Stretch = State
TitleBarMiddle.Stretch = State
TitleBarRight.Stretch = State
End Sub

