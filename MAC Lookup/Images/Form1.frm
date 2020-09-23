VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "MAC Address Look-up"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   3810
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtIP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "192.168.0.1"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtVendor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton cmdRetreive 
      Caption         =   "Retreive"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtMAC 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6600
      Top             =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IP Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MAC Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1005
   End
   Begin VB.Image Drag 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FF As SkinFormcls

Private Sub cmdRetreive_Click()
Dim sMAC As String

If Main.txtIP.Text = "" Then Exit Sub

sMAC = GetMAC(txtIP.Text)
txtMAC.Text = sMAC
txtVendor.Text = MAClookup(sMAC)
End Sub

Private Sub Drag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Move Form
ReleaseCapture
SendMessage Me.hwnd, &HA1, 2, 0&
End Sub



Private Sub Form_Load()

'/-----Skin the form------/
Set FF = New SkinFormcls

With FF
    .SetForm Me
    .CreateImageControls
    .LoadPictures
    .ResizeImages
End With

Me.Show
'/-----End------/

'/-----Custom Menu------/

Load CustomMnu

With CustomMnu
    .CreateMenuButton "Lookup"
    .CreateMenuButton "-"
    .CreateMenuButton "About"
    .CreateMenuButton "-"
    .CreateMenuButton "Exit"
End With
'/-----End------/

LoadMACdb 'Load the Database
End Sub

Private Sub Form_LostFocus()
'When Not selected, change picture
FF.FormNotFocused
End Sub

Private Sub Form_GotFocus()
'When selected, change picture
FF.FormFocused
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    CustomMnu.ShowMenu
End If
End Sub

Private Sub Form_Resize()
'on form Resize, resize images
FF.ResizeImages
End Sub

Private Sub Timer1_Timer()
If GetActiveWindow() = Me.hwnd And FF.Focused = False Then FF.FormFocused

If GetActiveWindow() <> Me.hwnd And FF.Focused = True Then FF.FormNotFocused
End Sub
