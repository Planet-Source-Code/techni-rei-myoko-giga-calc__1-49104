VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Giga-Calulator"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9015
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin GigaCalc.XPWIN XPWIN 
      Height          =   5655
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   9975
      Begin GigaCalc.XPCMD2 XPCMD 
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   13
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Begin VB.Label lblcmd 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Delete"
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
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   60
            Width           =   735
         End
      End
      Begin GigaCalc.XPCMD2 XPCMD 
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Begin VB.Label lblcmd 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Edit"
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
            TabIndex        =   5
            Top             =   60
            Width           =   735
         End
      End
      Begin GigaCalc.XPCMD2 XPCMD 
         Height          =   375
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Begin VB.Label lblcmd 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Add"
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
            TabIndex        =   3
            Top             =   60
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView lstmain 
         Height          =   4920
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   735
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   8678
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Value"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbltitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Variables"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   2055
      End
   End
   Begin GigaCalc.XPWIN XPWIN 
      Height          =   5655
      Index           =   1
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9975
      Begin VB.ListBox Listmain 
         Height          =   4545
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtmain 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   5220
         Width           =   4335
      End
      Begin GigaCalc.XPCMD2 XPCMD 
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   8
         Top             =   5160
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Begin VB.Label lblcmd 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Evaluate"
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
            Index           =   3
            Left            =   120
            TabIndex        =   9
            Top             =   60
            Width           =   735
         End
      End
      Begin VB.Label lbltitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Equation Evaluation"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   60
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const pi As String = "3.141592653589793238462"
Const e As String = "2.718281828459045235360"

Private Sub Form_Load()
    Dim temp As Long
    For temp = 0 To 3
        XPCMD(temp).enabled = True
    Next
    
    declare_var publicvars, varcount, Array("set", "pi", "to", pi)
    declare_var publicvars, varcount, Array("set", "e", "to", e)
    declare_var publicvars, varcount, Array("set", "answer", "to", "0")
    
    If command <> Empty Then
        loadvars publicvars, varcount, getfromquotes(command)
    End If
    
    refreshvars
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If command <> Empty Then
        savevars publicvars, varcount, getfromquotes(command)
    End If
End Sub

Private Sub lblcmd_Click(Index As Integer)
    XPCMD_Click Index, 0, 0, 0
End Sub

Private Sub lblcmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    XPCMD(Index).state = False
End Sub

Private Sub lblcmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    XPCMD(Index).state = True
End Sub

Private Sub Listmain_Click()
On Error Resume Next
    If Listmain.ListIndex Mod 2 = 0 Then
        txtmain = Listmain.List(Listmain.ListIndex) & Replace(Listmain.List(Listmain.ListIndex + 1), vbTab, Empty)
    Else
        txtmain = Listmain.List(Listmain.ListIndex - 1) & Replace(Listmain.List(Listmain.ListIndex), vbTab, Empty)
    End If
End Sub

Private Sub txtmain_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then XPCMD_Click 3, 0, 0, 0
End Sub

Public Sub XPCMD_Click(Index As Integer, x As Long, Y As Long, Button As Integer)
    On Error Resume Next
    Dim tempname As String, tempvalue As String
    
    If Index = 1 Or Index = 0 Then
        If Index = 1 Then
            tempname = lstmain(0).selecteditem.text
            tempvalue = lstmain(0).selecteditem.SubItems(2)
        End If
        frmvar.varedit tempname, tempvalue
        If tempname <> Empty Then
        If functionexists(tempname) Then
            MsgBox tempname & " is a reserved name. Please choose another" & vbNewLine & vbNewLine & "You can not use:" & vbNewLine & Replace(registeredfunctions, "|", ", "), vbCritical, "This name is reserved"
        Else
            If Not varexists(publicvars, varcount, tempname) Then
                If frmvar.chkarray = vbUnchecked Then
                    declare_var publicvars, varcount, Array("declare", tempname, "as", frmvar.getvartype, "with", tempvalue)
                Else
                    declare_var publicvars, varcount, Array("declare", tempname, "as", "array", "with", frmvar.VarDelimeters & " " & frmvar.getvartype)
                    declare_var publicvars, varcount, Array("set", tempname, "to", tempvalue)
                End If
            Else
                declare_var publicvars, varcount, Array("set", tempname, "to", tempvalue)
            End If
            refreshvars
        End If
        End If
    End If
    
    If Index = 2 Then
        tempname = lstmain(0).selecteditem.text
        deletevar publicvars, varcount, tempname
        refreshvars
    End If
    
    If Index = 3 Then
        If LCase(txtmain.text) = "end" Then End
        tempvalue = Eval(txtmain.text)
        Listmain.additem txtmain.text
        Listmain.additem vbTab & "= " & tempvalue
        txtmain.text = Empty
        Listmain.ListIndex = Listmain.ListCount
        declare_var publicvars, varcount, Array("set", "Answer", "to", tempvalue)
        refreshvars
    End If
End Sub
Public Sub refreshvars()
    lstmain(0).ListItems.Clear
    Dim temp As Long
    For temp = 1 To varcount
        additem lstmain(0), temp = varcount, publicvars(temp).var_name, publicvars(temp).var_type, publicvars(temp).var_valu
    Next
End Sub
