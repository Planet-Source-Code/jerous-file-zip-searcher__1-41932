VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   4245
   ScaleWidth      =   7620
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   4425
      Top             =   1785
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   2760
      ScaleHeight     =   2415
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   2985
      Visible         =   0   'False
      Width           =   3915
   End
   Begin MSComctlLib.ImageList imgLst 
      Left            =   4575
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   195
      ImageHeight     =   113
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   2955
      Top             =   2040
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim W%, H%, Y%, XPos%, YPos%, XFactor%, YFactor%
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private OnJer As Boolean
Private OnRich As Boolean
Private Sub Form_Deactivate()
Form_MouseDown 0, 0, 0, 0
End Sub

Private Sub Form_Load()
Caption = AppName
W = imgLst.ImageWidth
H = imgLst.ImageHeight
imgLst.ListImages(1).Draw pic.hdc, 0, 0
SaveSetting AppName, "Options", "AboutBoxShowed", True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FontSize = 7
If Y <= Height - TextHeight("W") * 4 Or Y > Height - TextHeight("W") * 3 Then
  Unload Me
Else
  If OnJer Then
    Call ShellExecute(0&, vbNullString, "mailto:jeremyvandeneynde@hotmail.com?subject=GR Productions File Searcher", vbNullString, vbNullString, vbNormalFocus)
  End If
  If OnRich Then
    Call ShellExecute(0&, vbNullString, "mailto:richard@richsoftcomputing.co.uk?subject=Richsoft VBZip", vbNullString, vbNullString, vbNormalFocus)
  End If
  OnJer = False
  OnRich = False
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
FontSize = 7
OnJer = False
OnRich = False
If Y <= Height - TextHeight("W") * 4 Or Y > Height - TextHeight("W") * 3 Then
  Me.MousePointer = 0
Else
  Me.MousePointer = 99
  Me.MouseIcon = LoadPicture(Form1.RightPath(App.Path) + "hand.ico")
  If X <= Width / 2 Then OnJer = True Else OnRich = True
End If
End Sub

Private Sub Timer1_Timer()
Dim X%, Pos As POINTAPI
Const Stap = 3
GetCursorPos Pos
If XPos > Width - W * Screen.TwipsPerPixelX Then XPos = Width - W * Screen.TwipsPerPixelX: XFactor = XFactor * -1
If YPos > Height - H * Screen.TwipsPerPixelY - 300 Then YPos = Height - H * Screen.TwipsPerPixelY - 300: YFactor = YFactor * -1

If XPos < 0 Then XPos = 0: XFactor = XFactor * -1
If YPos < 0 Then YPos = 0: YFactor = YFactor * -1

Pos.X = (Pos.X - W / 2) * Screen.TwipsPerPixelX - Left
Pos.Y = (Pos.Y - H / 2) * Screen.TwipsPerPixelY - Top
If Pos.X > XPos Then XFactor = XFactor + 5
If Pos.X < XPos Then XFactor = XFactor - 5
If Pos.Y > YPos Then YFactor = YFactor + 5
If Pos.Y < YPos Then YFactor = YFactor - 5
'move picture
XPos = XPos + XFactor
YPos = YPos + YFactor
'slow down movements
XFactor = XFactor * 0.98
YFactor = YFactor * 0.98

Cls
'draw logo in 'flag' form
For X = 0 To W Step Stap
  BitBlt Me.hdc, XPos / Screen.TwipsPerPixelX + X, YPos / Screen.TwipsPerPixelY + 3 * Sin(Y + X / Stap / 3), 10, H, pic.hdc, X, 0, vbSrcCopy
Next

'print info
CurrentX = 2000
CurrentY = 1500
FontSize = 20
Print AppName
Print ""
FontSize = 7
CurrentY = Height - TextHeight(String(5, vbCrLf))
Line (0, CurrentY - 50)-(Width, CurrentY - 50)

CurrentX = 0
Print "For info or suggestions: mail to"
Print vbTab + "GR Productions"
FontUnderline = True
ForeColor = vbBlack
If OnJer Then ForeColor = vbBlue
Print vbTab + "jeremyvandeneynde@hotmail.com"
ForeColor = vbBlack
FontUnderline = False

CurrentX = Width / 2
CurrentY = Height - TextHeight(String(5, vbCrLf))
Print "Much thanks to (for Zip functions):"
CurrentX = Width / 2
Print vbTab + "Richsoft Computing 2001 "
FontUnderline = True
CurrentX = Width / 2
ForeColor = vbBlack
If OnRich Then ForeColor = vbBlue
Print vbTab + "richard@richsoftcomputing.co.uk "
ForeColor = vbBlack
FontUnderline = False
End Sub

Private Sub Timer2_Timer()
Y = Y + 3.14 / 4
End Sub
