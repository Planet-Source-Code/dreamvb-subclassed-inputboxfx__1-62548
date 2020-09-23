VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Subclassed InputBox - No More Custom forms"
   ClientHeight    =   3555
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBackPatten 
      Caption         =   "Allow Back Patten"
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   1935
      Width           =   2445
   End
   Begin VB.CheckBox chkFlatEdit 
      Caption         =   "Flatten EditBox"
      Height          =   315
      Left            =   135
      TabIndex        =   11
      Top             =   1545
      Width           =   1500
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   420
      Left            =   2295
      TabIndex        =   10
      Top             =   3000
      Width           =   1050
   End
   Begin VB.PictureBox PicFc 
      BackColor       =   &H0000FFFF&
      Height          =   360
      Left            =   4140
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   9
      Top             =   780
      Width           =   690
   End
   Begin VB.PictureBox PicBk 
      BackColor       =   &H000080FF&
      Height          =   360
      Left            =   4140
      ScaleHeight     =   300
      ScaleWidth      =   630
      TabIndex        =   7
      Top             =   255
      Width           =   690
   End
   Begin VB.TextBox txtAlpha 
      Height          =   285
      Left            =   1950
      TabIndex        =   5
      Text            =   "160"
      Top             =   1095
      Width           =   615
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   285
      Left            =   1950
      TabIndex        =   3
      Text            =   "50"
      Top             =   690
      Width           =   630
   End
   Begin VB.CheckBox chkTimeOut 
      Caption         =   "EnableTimeOut"
      Height          =   315
      Left            =   210
      TabIndex        =   2
      Top             =   675
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show InputBoxEX"
      Height          =   420
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   2025
   End
   Begin VB.CheckBox chkDisR 
      Caption         =   "Enable Right Click on EditBox"
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   2625
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "You can even fnd out what button was pressed"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2430
      Width           =   3360
   End
   Begin VB.Label lblButton 
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   2685
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   4635
      Picture         =   "demo.frx":0000
      Top             =   2580
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Forecolor"
      Height          =   195
      Left            =   3255
      TabIndex        =   8
      Top             =   855
      Width           =   660
   End
   Begin VB.Label lblbk 
      AutoSize        =   -1  'True
      Caption         =   "Backcolor"
      Height          =   195
      Left            =   3225
      TabIndex        =   6
      Top             =   330
      Width           =   720
   End
   Begin VB.Label lblalpha 
      AutoSize        =   -1  'True
      Caption         =   "AlphaBlend Level"
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   1125
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
    MsgBox "InputBoxEx " & vbCrLf & vbTab & "By Ben Jones" & vbCrLf & "Please Vote", vbInformation: End
    
End Sub

Private Sub Command1_Click()
Dim x As Long, y As Long
Dim lResult As Variant
Dim sPromt As String

    x = (Screen.Width \ Screen.TwipsPerPixelX) \ 2
    y = (Screen.Height \ Screen.TwipsPerPixelY) \ 2
    
    
    InputFx.FlatEditBox = chkFlatEdit
    InputFx.fFontSize = 18 'Fontsize
    InputFx.fFontName = "Comic Sans MS" ' Font for the inputbox
    InputFx.AllowRClick = chkDisR ' Disable or enable right clicking on edit box
    InputFx.DisableX = True ' Disbale X Close button
    InputFx.EnableTimeOut = chkTimeOut 'Enable self close
    InputFx.TimeOut = Val(txtTimeOut.Text) 'Set for 50 seconds
    InputFx.AllowBackImage = chkBackPatten 'Turn on texture support
    InputFx.Image = Image1 'Allow the inputbox to have a tiled bacground
    
    sPromt = "Please enter your name:" & vbCrLf & "X Button is Disbaled" _
    & vbCrLf & "Right Click is Disbaled"
    
    
    lResult = InputFx.InputBoxFx(sPromt, "This box will close in 50 seconeds", , x, y, , _
    , PicBk.BackColor, PicFc.BackColor, Val(txtAlpha.Text), vbBlue, vbWhite)
    
    Debug.Print "Result : " & lResult
    lblButton.Caption = "Button Pressed : " & InputFx.ButtonPressed
    
End Sub



