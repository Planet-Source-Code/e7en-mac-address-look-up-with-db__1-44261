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

'----------------------
'/Define Other
'----------------------
Dim SkinedForm As Form

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

End Sub

Sub ResizeImages()
'----------------------
'/Resize Title Bar
'----------------------
With TitleBarLeft
    .Left = 0
    .Top = 0
End With

With TitleBarLeft
    .Left = SkinedForm.Width - TitleBarLeft.Width
    .Top = 0
End With

With TitleBarMiddle
    .Left = TitleBarLeft.Width
    .Top = 0
    .Width = SkinedForm.Width - (TitleBarLeft.Width + TitleBarRight.Width)
End With

End Sub

Sub FormFocused()
'----------------------
'/Set Title Bar Focused Pictures
'----------------------
Set TitleBarLeft.Picture = TitleBarLeftFocusY
Set TitleBarMiddle.Picture = TitleBarMiddleFocusY
Set TitleBarRight.Picture = TitleBarRightFocusY

End Sub

Sub FormNotFocused()
'----------------------
'/Set Title Bar Not Focused Pictures
'----------------------
Set TitleBarLeft.Picture = TitleBarLeftFocusN
Set TitleBarMiddle.Picture = TitleBarMiddleFocusN
Set TitleBarRight.Picture = TitleBarRightFocusN

End Sub

Sub LoadPictures()
Dim Path As String

Path = App.Path & "\Images"

Set TitleBarLeftFocusY = LoadPicture(Path & "\TBTLFY.png")
Set TitleBarMiddleFocusY = LoadPicture(Path & "\TBTLMY.png")
Set TitleBarRightFocusY = LoadPicture(Path & "\TBTLRY.png")

Set TitleBarLeftFocusN = LoadPicture(Path & "\TBTLFN.png")
Set TitleBarMiddleFocusN = LoadPicture(Path & "\TBTLMN.png")
Set TitleBarRightFocusN = LoadPicture(Path & "\TBTLRN.png")
End Sub

Private Sub Class_Initialize()
CreateImageControls
End Sub
