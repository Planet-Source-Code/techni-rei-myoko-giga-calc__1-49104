VERSION 5.00
Begin VB.UserControl XPCMD2 
   AutoRedraw      =   -1  'True
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1560
   ControlContainer=   -1  'True
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   104
   ToolboxBitmap   =   "XPCMD2.ctx":0000
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   5
      Left            =   1320
      Picture         =   "XPCMD2.ctx":0312
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   5
      ToolTipText     =   "disabled"
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   4
      Left            =   1080
      Picture         =   "XPCMD2.ctx":0504
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   4
      ToolTipText     =   "mouse out, in focus"
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   3120
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   3
      Left            =   840
      Picture         =   "XPCMD2.ctx":06F6
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   3
      ToolTipText     =   "mouse out, in focus"
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   2
      Left            =   600
      Picture         =   "XPCMD2.ctx":08E8
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   2
      ToolTipText     =   "Mouse Up, over, in or out of focus"
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   360
      Picture         =   "XPCMD2.ctx":0ADA
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   1
      ToolTipText     =   "Mouse Up, Out, not in focus"
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   0
      Left            =   120
      Picture         =   "XPCMD2.ctx":0CCC
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   0
      ToolTipText     =   "Mouse Down"
      Top             =   120
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "XPCMD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long 'use vbSrcCopy for dwRop
Private Declare Function GetCursorPos Lib "USER32" (lpPoint As Point) As Long
Private Declare Function WindowFromPoint Lib "USER32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Dim mousestate As Boolean ' true is up, false is down
Dim MouseOver  As Boolean ' true is over, false is not
Dim hasfocus   As Boolean ' true is yes, false is not

Dim temp As Point
Dim loc As Point
Dim mousebutton As Integer
Private Type Point
    x As Long
    Y As Long
End Type

Public Event Click(x As Long, Y As Long, Button As Integer)

Public Event MouseMove(x As Long, Y As Long)
Public Event MouseOver()
Public Event MouseLeave()
Public Event MouseDown(Button As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, x As Single, Y As Single)

Public Property Let enabled(state As Boolean)
    Timer.enabled = state
    statechanged
End Property
Public Property Get enabled() As Boolean
    enabled = Timer.enabled
End Property
Public Property Let state(UpIsTrueDownIsFalse As Boolean)
    If mousestate = Not UpIsTrueDownIsFalse Then
        mousestate = UpIsTrueDownIsFalse
        statechanged
    End If
End Property
Public Sub statechanged()
Dim draw As Long
If mousestate = False Then 'down
    draw = 0
Else
    If MouseOver = True Then
        If hasfocus = False Then
            draw = 2
        Else
            draw = 4
        End If
    Else
        If hasfocus = True Then
            draw = 3
        Else
            draw = 1
        End If
    End If
End If
If Timer.enabled = False And mousestate Then draw = 5

UserControl.Picture = LoadPicture("")
UserControl.BackColor = GetPixel(picmain(draw).hdc, 5, 5)

BitBlt UserControl.hdc, 0, 0, 4, 4, picmain(draw).hdc, 0, 0, vbSrcCopy
BitBlt UserControl.hdc, UserControl.width / 15 - 4, 0, 4, 4, picmain(draw).hdc, 8, 0, vbSrcCopy
BitBlt UserControl.hdc, 0, UserControl.height / 15 - 4, 4, 4, picmain(draw).hdc, 0, 8, vbSrcCopy
BitBlt UserControl.hdc, UserControl.width / 15 - 4, UserControl.height / 15 - 4, 4, 4, picmain(draw).hdc, 8, 8, vbSrcCopy

tilevertical UserControl.hdc, picmain(draw).hdc, 0, 4, 4, 4, 0, 4, UserControl.height / 15 - 8
tilevertical UserControl.hdc, picmain(draw).hdc, 8, 4, 4, 4, UserControl.width / 15 - 4, 4, UserControl.height / 15 - 8

tilehorizontal UserControl.hdc, picmain(draw).hdc, 4, 0, 4, 4, 4, 0, UserControl.width / 15 - 8
tilehorizontal UserControl.hdc, picmain(draw).hdc, 4, 8, 4, 4, 4, UserControl.height / 15 - 4, UserControl.width / 15 - 8

UserControl.Refresh
End Sub

Private Sub tilehorizontal(destHDC As Long, srchdc As Long, Left As Single, Top As Single, width As Single, height As Single, x As Single, Y As Single, newwidth As Single)
On Error Resume Next
    Dim count As Single
    For count = x To x + newwidth Step width
        If count + width > x + newwidth Then GoTo last:
        BitBlt destHDC, count, Y, width, height, srchdc, Left, Top, vbSrcCopy
    Next
last:
If width <= newwidth Then
    BitBlt destHDC, x + newwidth - width, Y, width, height, srchdc, Left, Top, vbSrcCopy
Else
    BitBlt destHDC, x, Y, newwidth, height, srchdc, Left, Top, vbSrcCopy
End If
End Sub
Private Sub tilevertical(destHDC As Long, srchdc As Long, Left As Single, Top As Single, width As Single, height As Single, x As Single, Y As Single, newheight As Single)
On Error Resume Next
    Dim count As Single
    For count = Y To Y + newheight Step width
        If count + height > Y + newheight Then GoTo last:
        BitBlt destHDC, x, count, width, height, srchdc, Left, Top, vbSrcCopy
    Next
last:
If height <= newheight Then
    BitBlt destHDC, x, Y + newheight - height, width, height, srchdc, Left, Top, vbSrcCopy
Else
    BitBlt destHDC, x, Y, width, newheight, srchdc, Left, Top, vbSrcCopy
End If
End Sub

Private Sub Timer_Timer()
GetCursorPos temp
If WindowFromPoint(temp.x, temp.Y) = UserControl.hWnd Then
    If MouseOver = False Then
        RaiseEvent MouseOver
        MouseOver = True
        statechanged
    End If
Else
    If MouseOver = True Then
        RaiseEvent MouseLeave
        MouseOver = False
        statechanged
    End If
End If
End Sub

Private Sub UserControl_Click()
RaiseEvent Click(loc.x, loc.Y, mousebutton)
End Sub

Private Sub UserControl_GotFocus()
If hasfocus = False Then
    hasfocus = True
    statechanged
End If
End Sub

Private Sub UserControl_Initialize()
mousestate = True
End Sub

Private Sub UserControl_LostFocus()
If hasfocus = True Then
    hasfocus = False
    statechanged
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
mousebutton = Button
Timer.enabled = False
If mousestate = True Then
    mousestate = False
    statechanged
End If
RaiseEvent MouseDown(Button, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
RaiseEvent MouseMove(x * 1, Y * 1)
loc.x = x
loc.Y = Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
mousebutton = Button
Timer.enabled = True
If mousestate = False Then
    mousestate = True
    statechanged
End If
RaiseEvent MouseUp(Button, x, Y)
End Sub

Private Sub UserControl_Resize()
statechanged
End Sub
