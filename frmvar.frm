VERSION 5.00
Begin VB.Form frmvar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Variable Edit/Create"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4545
   Icon            =   "frmvar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkarray 
      Caption         =   "Is this even an array?"
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtdirections 
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Text            =   "frmvar.frx":000C
      Top             =   3480
      Width           =   4335
   End
   Begin VB.TextBox VarDelimeters 
      Height          =   285
      Left            =   1440
      TabIndex        =   18
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox VarDimensions 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "0"
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox VarName 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   1455
   End
   Begin VB.ListBox lstmain 
      Height          =   1425
      ItemData        =   "frmvar.frx":055D
      Left            =   120
      List            =   "frmvar.frx":055F
      TabIndex        =   9
      Top             =   960
      Width           =   4335
   End
   Begin VB.TextBox VarValue 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin GigaCalc.XPCMD2 XPCMD 
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Begin VB.Label lblcmd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   60
         Width           =   735
      End
   End
   Begin GigaCalc.XPCMD2 XPCMD 
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   5280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Begin VB.Label lblcmd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   60
         Width           =   735
      End
   End
   Begin GigaCalc.XPCMD2 XPCMD 
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   13
      Top             =   2640
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Begin VB.Label lblcmd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   14
         Top             =   0
         Width           =   135
      End
   End
   Begin GigaCalc.XPCMD2 XPCMD 
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   15
      Top             =   2880
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Begin VB.Label lblcmd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   16
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Label lblvar 
      Caption         =   "Help:"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblvar 
      Caption         =   "Delimeters:"
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   17
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblvar 
      Caption         =   "Dimensions:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblvar 
      Caption         =   "Type:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblvar 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lblvar 
      Caption         =   "Value:"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblvar 
      Caption         =   "Name:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmvar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const maxdim As Long = 10

Public Function varedit(Optional ByRef name As String, Optional ByRef value As String) As String
    Dim temp As Long, tempstr() As String
    Me.VarName.text = Empty
    Me.VarName.enabled = True
    Me.VarValue = Empty
    Me.VarDelimeters = Empty
    Me.VarDimensions = "0"
    chkarray.value = vbUnchecked
    lstmain.enabled = True
    
    If name <> Empty Then
        Me.VarName.enabled = False
        Me.VarName.text = name
        lstmain.enabled = False
    End If
    If value <> Empty Then Me.VarValue = value
       
    lstmain.Clear
    tempstr = Split(enumTypes(publicvars, varcount), "|")
    For temp = 0 To UBound(tempstr)
        If tempstr(temp) <> "array" Then lstmain.additem tempstr(temp)
    Next
    lstmain.ListIndex = 1
    
    Me.Show vbModal, frmmain
    
    name = Me.VarName
    value = Me.VarValue
    
    varedit = Me.VarValue
End Function

Private Sub lblcmd_Click(Index As Integer)
    XPCMD_Click Index, 0, 0, 0
End Sub

Private Sub lblcmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    XPCMD(Index).state = False
End Sub

Private Sub lblcmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    XPCMD(Index).state = True
End Sub
Private Sub Form_Load()
Dim temp As Long
For temp = 0 To 3
    XPCMD(temp).enabled = True
Next
End Sub

Private Sub VarDimensions_KeyPress(KeyAscii As Integer)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then KeyAscii = 0 ':Debug.Print KeyAscii
End Sub

Private Sub VarName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then VarValue.SetFocus
End Sub

Private Sub VarValue_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then XPCMD_Click 0, 0, 0, 0
End Sub

Public Sub XPCMD_Click(Index As Integer, x As Long, Y As Long, Button As Integer)
    If Index < 2 Then
        If Index = 0 Then
            If Len(VarDelimeters) <> Val(VarDimensions) And getvartype = "array" Then
                MsgBox "You must have the same amount of delimeters as dimensions", vbCritical, "Array Delimeter/Dimension Error"
                Exit Sub
            End If
        End If
        If Index = 1 Then Me.VarValue = Empty
        Me.Visible = False
    End If
    'If getvartype = "array" Then
    If Index = 2 Then 'up
        If Val(VarDimensions) < maxdim Then
            VarDimensions = Val(VarDimensions) + 1
        Else
            VarDimensions = 0
        End If
    End If
    If Index = 3 Then 'down
        If Val(VarDimensions) > 0 Then
            VarDimensions = Val(VarDimensions) - 1
        Else
            VarDimensions = maxdim
        End If
    End If
    'End If
End Sub
Public Function getvartype() As String
    On Error Resume Next
    getvartype = LCase(lstmain.List(lstmain.ListIndex))
End Function
