VERSION 5.00
Begin VB.UserControl XPWIN 
   BackColor       =   &H00F7DED6&
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2850
   ControlContainer=   -1  'True
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   2415
   ScaleWidth      =   2850
   ToolboxBitmap   =   "XPWIN2.ctx":0000
   Begin VB.PictureBox picmain 
      BackColor       =   &H00F7DED6&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1095
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image imgstate 
      Height          =   285
      Index           =   3
      Left            =   1320
      Picture         =   "XPWIN2.ctx":0312
      ToolTipText     =   "Up"
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgstate 
      Height          =   285
      Index           =   2
      Left            =   960
      Picture         =   "XPWIN2.ctx":07C8
      ToolTipText     =   "Down"
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   1
      Left            =   240
      Picture         =   "XPWIN2.ctx":0C7E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "XPWIN2.ctx":0CBF
      Top             =   120
      Width           =   45
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   3
      Left            =   2640
      Picture         =   "XPWIN2.ctx":0D04
      Top             =   120
      Width           =   30
   End
   Begin VB.Image imgstate 
      Height          =   285
      Index           =   0
      Left            =   240
      Picture         =   "XPWIN2.ctx":0D49
      ToolTipText     =   "Down"
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgstate 
      Height          =   285
      Index           =   1
      Left            =   600
      Picture         =   "XPWIN2.ctx":11CD
      ToolTipText     =   "Up"
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgbutton 
      Height          =   285
      Left            =   2160
      Picture         =   "XPWIN2.ctx":164C
      Tag             =   "0"
      Top             =   180
      Width           =   285
   End
   Begin VB.Image imgborder 
      Height          =   1575
      Index           =   0
      Left            =   120
      Picture         =   "XPWIN2.ctx":1AD0
      Stretch         =   -1  'True
      Top             =   600
      Width           =   30
   End
   Begin VB.Image imgborder 
      Height          =   1575
      Index           =   1
      Left            =   2640
      Picture         =   "XPWIN2.ctx":1B07
      Stretch         =   -1  'True
      Top             =   600
      Width           =   30
   End
   Begin VB.Image imgborder 
      Height          =   30
      Index           =   2
      Left            =   120
      Picture         =   "XPWIN2.ctx":1B3E
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2565
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   2
      Left            =   1080
      Picture         =   "XPWIN2.ctx":1B75
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "XPWIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim xpwin_state As Boolean
Dim xp_speed As Long
Dim xpwin_height As Single
Dim xpwin_icon  As Boolean
Dim mousex As Single, mousey As Single
Public Event statechange(state As Boolean)
Public Event resize()
Public Event Click(x As Single, Y As Single)
Public Event MouseDown(x As Single, Y As Single, Button As Integer)
Public Event MouseUp(x As Single, Y As Single, Button As Integer)
Dim m_smooth As Boolean
Public Property Let canchangestate(state As Boolean)
    imgbutton.enabled = state
End Property
Public Property Let CanResize(resizeable As Boolean)
    imgbutton.enabled = resizeable
End Property
Public Property Get smooth() As Boolean
    smooth = m_smooth
End Property
Public Property Let smooth(issmooth As Boolean)
    m_smooth = issmooth
End Property

Public Property Get speed() As Long
    speed = xp_speed
End Property
Public Property Let speed(scrollspeed As Long)
    xp_speed = scrollspeed
End Property

Public Property Get CanResize() As Boolean
    CanResize = imgbutton.enabled
End Property
Public Property Let state(state As Boolean)
    If xpwin_state <> state Then imgbutton_Click
End Property
Public Property Get state() As Boolean
    state = xpwin_state
End Property

Public Sub imgbutton_Click()
xpwin_state = Not xpwin_state
imgbutton.enabled = False
Select Case xpwin_state
     Case True 'was down, move to up
          'UserControl.Height = xpwin_height
          expand
          imgbutton.Picture = imgstate(1).Picture
     Case False
          xpwin_height = UserControl.height
          'UserControl.Height = imghead(0).Height
          contract
          imgbutton.Picture = imgstate(0).Picture
End Select
imgbutton.enabled = True
RaiseEvent statechange(xpwin_state)
End Sub
Public Sub expand()
If m_smooth Then
Do Until UserControl.height >= xpwin_height
    UserControl.height = UserControl.height + xp_speed
    DoEvents
Loop
End If
UserControl.height = xpwin_height
End Sub
Public Sub contract()
If m_smooth Then
Do Until UserControl.height <= imghead(0).height
    UserControl.height = UserControl.height - xp_speed
    DoEvents
Loop
End If
UserControl.height = imghead(0).height
End Sub
Private Sub imghead_Click(Index As Integer)
    RaiseEvent Click(mousex, mousey)
End Sub

Private Sub imghead_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseDown(mousex, mousey, Button)
End Sub

Private Sub imghead_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
mousex = x + imghead(Index).Left
mousey = Y + imghead(Index).Top
End Sub

Private Sub imghead_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseUp(mousex, mousey, Button)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click(mousex, mousey)
End Sub

Private Sub UserControl_Initialize()
     xpwin_state = True
     xp_speed = 90
     xpwin_icon = False
     imgbutton.Picture = imgstate(1).Picture
     smooth = True
End Sub
Private Sub imgbutton_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If state = True Then
    imgbutton.Picture = imgstate(3).Picture
Else
    imgbutton.Picture = imgstate(2).Picture
End If
End Sub

Private Sub imgbutton_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If state = True Then
    imgbutton.Picture = imgstate(1).Picture
Else
    imgbutton.Picture = imgstate(0).Picture
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseDown(mousex, mousey, Button)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
mousey = Y
mousex = x
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseUp(mousex, mousey, Button)
End Sub

Private Sub UserControl_Resize()
imghead(0).Move 0, 0
imghead(1).Move imghead(0).Left, 0
imghead(3).Move UserControl.width - imghead(3).width, 0
imghead(2).Move imghead(1).Left + imghead(1).width, 0, imghead(3).Left - imghead(1).Left - imghead(1).width
imgbutton.Move UserControl.width - imgbutton.width * 1.3, imghead(0).height / 2 - imgbutton.height / 2
If UserControl.height > imghead(0).height Then
    imgborder(0).Move 0, imghead(0).height, imgborder(0).width, UserControl.height - imghead(0).height
    imgborder(1).Move UserControl.width - imgborder(1).width, imghead(0).height, imgborder(1).width, UserControl.height - imghead(0).height
    imgborder(2).Move 0, UserControl.height - imgborder(2).height, UserControl.width
    picmain.Move imgborder(0).width, imghead(0).height, UserControl.width - imgborder(0).width * 2, UserControl.height - imghead(0).height - imgborder(2).height
End If
RaiseEvent resize
End Sub
