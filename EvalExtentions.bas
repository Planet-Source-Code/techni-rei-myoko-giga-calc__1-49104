Attribute VB_Name = "EvalExtentions"
Option Explicit
Public publicvars() As variable, varcount As Long
Public Const registeredfunctions As String = "hex|replace|sin|cos|tan|rnd|len|left|right|mid|not|sqrt|abs|exp|log|time|ucase|lcase|round|asc|chr|trim|instr|instrrev|datediff|strreverse|savevars|loadvars"
Public Sub processvars(strarray)
    Dim temp As Long, vari As String
    For temp = LBound(strarray) To UBound(strarray)
        If chartype(strarray(temp) & "") = ch_routine And Left(strarray(temp) & "", 1) <> """" Then
            If varexists(publicvars, varcount, strarray(temp) & Empty) Or functionexists(strarray(temp) & Empty) Then
                vari = getvariable(strarray(temp) & Empty)
                strarray(temp) = Eval(vari)
            End If
        End If
    Next
End Sub
Public Sub setvariable(name As String, value As String)
    declare_var publicvars, varcount, Array("set", name, "to", value)
End Sub
Public Function getvariable(ByVal VarName As String) As String
    Dim temp As String, isneg As Boolean
    If Left(VarName, 1) = "-" Then
        isneg = True
        VarName = Right(VarName, Len(VarName) - 1)
    End If
    If InStr(VarName, "(") > 0 Or InStr(VarName, "[") > 0 Then VarName = evalvarname(VarName)
    If functionexists(VarName) Then
            
            temp = getfunct(VarName)
        Else
            temp = declare_var(publicvars, varcount, Array("get", VarName))
    End If
    temp = Right(temp, Len(temp) - InStrRev(temp, "="))
    If isneg And isanumber(temp) Then
        getvariable = Val(temp) * 1
    Else
        getvariable = temp
    End If
End Function
Public Sub leftv(Optional ByRef temp As String, Optional ByRef VarName As String, Optional ByRef temp2 As String, Optional ByRef lbrack As String, Optional ByRef rbrack As String, Optional offset As Long = 0)
    If InStr(VarName, lbrack) > 0 Then
        temp = Left(VarName, InStr(VarName, lbrack) - offset)
        VarName = Right(VarName, Len(VarName) - InStr(VarName, lbrack))
        temp2 = Right(VarName, Len(VarName) - InStrRev(VarName, rbrack) + 1)
        VarName = Left(VarName, InStr(VarName, rbrack) - 1)
    End If
End Sub
Public Function evalvarname(ByVal VarName As String, Optional temp As String, Optional ByRef tempstr, Optional temp2 As String) As String
    'example 'as[x,3534,fgfddf].value'
    Dim count As Long
    
    'Cut of the 'as[' and '].value'
    leftv temp, VarName, temp2, "[", "]"
    leftv temp, VarName, temp2, "(", ")"
    
    If InStr(VarName, ",") > 0 Then
        tempstr = Split(VarName, ",")
        For count = 0 To UBound(tempstr)
            tempstr(count) = Eval(tempstr(count))
        Next
        evalvarname = temp & Join(tempstr, ",") & temp2
    Else
        evalvarname = temp & Eval(VarName) & temp2
    End If
End Function
Public Function functionexists(ByVal functname As String) As Boolean
    On Error Resume Next
    Dim test As String, temp As Long, flist() As String
    leftv test, functname, , "[", "]", 1
    leftv test, functname, , "(", ")", 1
    If test = Empty Then test = functname
    flist = filter(Split(registeredfunctions, "|"), test, , vbTextCompare)
    functionexists = flist(0) = test
End Function

Public Function getfunct(ByVal VarName As String) As String
'On Error Resume Next
'On Error GoTo err:
 Dim leftside As String, middle As String, rightside As String, parameters() As String, count As Long
 middle = VarName
 leftv leftside, middle, rightside, "[", "]", 1
 leftv leftside, middle, rightside, "(", ")", 1
 splitbychartype middle, parameters
 leftside = LCase(leftside)
 If leftside = Empty Then leftside = VarName
 Select Case leftside
    Case "abs": If verifyparameters(parameters, True, 1, Empty, "num") Then getfunct = Abs(getcell(parameters, 1))
    Case "asc": If verifyparameters(parameters, True, 1, Empty, "any") Then getfunct = Asc(Left(getfromquotes(getcell(parameters, 1)), 1))
    Case "chr": If verifyparameters(parameters, True, 1, Empty, "num") Then getfunct = Cos(getcell(parameters, 1))
    Case "cos": If verifyparameters(parameters, True, 1, Empty, "num") Then getfunct = Chr(getcell(parameters, 1))
    Case "datediff": If verifyparameters(parameters, True, 3, Empty, "any", "any", "any") Then getfunct = DateDiff(getfromquotes(getcell(parameters, 1)), getfromquotes(getcell(parameters, 2)), getfromquotes(getcell(parameters, 3)))
    Case "exp": If verifyparameters(parameters, True, 1, Empty, "num") Then getfunct = Exp(getcell(parameters, 1))
    Case "hex": If verifyparameters(parameters, True, 1, Empty, "num") Then getfunct = Hex(getcell(parameters, 1))
    Case "instr": If verifyparameters(parameters, True, 2, Empty, "any", "any") Then getfunct = InStr(getfromquotes(getcell(parameters, 1)), getfromquotes(getcell(parameters, 2)))
    Case "instrrev": If verifyparameters(parameters, True, 2, Empty, "any", "any") Then getfunct = InStrRev(getfromquotes(getcell(parameters, 1)), getfromquotes(getcell(parameters, 2)))
    Case "lcase": If verifyparameters(parameters, True, 1, Empty, "any") Then getfunct = LCase(getcell(parameters, 1))
    Case "left": If verifyparameters(parameters, True, 2, Empty, "any", "num") Then getfunct = """" & Left(getfromquotes(getcell(parameters, 1)), getcell(parameters, 2)) & """"
    Case "len": If verifyparameters(parameters, True, 1, Empty, "any") Then getfunct = Len(getcell(parameters, 1))
    Case "loadvars": If verifyparameters(parameters, True, 1, Empty, "txt") Then getfunct = loadvars(publicvars, varcount, getfromquotes(getcell(parameters, 1))): frmmain.refreshvars
    Case "log": If verifyparameters(parameters, True, 1, Empty, "num") Then getfunct = Log(getcell(parameters, 1))
    Case "mid": If verifyparameters(parameters, True, 3, Empty, "any", "num", "num") Then getfunct = """" & Mid(getfromquotes(getcell(parameters, 1)), getcell(parameters, 2), getcell(parameters, 3)) & """"
    Case "not": If verifyparameters(parameters, True, 0, Empty, "bool") Then getfunct = IIf(LCase(getcell(parameters, 1)) = "true", "false", "true")
    Case "replace": If verifyparameters(parameters, True, 3, Empty, "any", "any", "any") Then getfunct = Replace(getfromquotes(getcell(parameters, 1)), getfromquotes(getcell(parameters, 2)), getfromquotes(getcell(parameters, 3)))
    Case "right": If verifyparameters(parameters, True, 2, Empty, "any", "num") Then getfunct = """" & Right(getfromquotes(getcell(parameters, 1)), getcell(parameters, 2)) & """"
    Case "rnd": If verifyparameters(parameters, True, 0, Empty, "num") Then Randomize getcell(parameters, 1, "0"): getfunct = Rnd
    Case "round": If verifyparameters(parameters, True, 2, Empty, "num", "num") Then getfunct = Round(getcell(parameters, 1), getcell(parameters, 2))
    Case "savevars": If verifyparameters(parameters, True, 1, Empty, "txt") Then getfunct = savevars(publicvars, varcount, getfromquotes(getcell(parameters, 1))): frmmain.refreshvars
    Case "sin": If verifyparameters(parameters, True, 1, Empty, "num") Then getfunct = Sin(getcell(parameters, 1))
    Case "sqrt": If verifyparameters(parameters, True, 1, Empty, "num") Then getfunct = Sqr(Abs(getcell(parameters, 1)))
    Case "strreverse": If verifyparameters(parameters, True, 1, Empty, "any") Then getfunct = StrReverse(getcell(parameters, 1))
    Case "tan": If verifyparameters(parameters, True, 1, Empty, "num") Then getfunct = Tan(getcell(parameters, 1))
    Case "time": getfunct = Time & Empty
    Case "trim": If verifyparameters(parameters, True, 1, Empty, "any") Then getfunct = Trim(getcell(parameters, 1))
    Case "ucase": If verifyparameters(parameters, True, 1, Empty, "any") Then getfunct = UCase(getcell(parameters, 1))
 End Select
'err:
 '   If err.Number > 0 Then
 '       MsgBox "Unable to check " & VarName & vbNewLine & err.Number & ") " & err.Description, vbCritical, "Equation Evaluation: Parameter Verification"
 '       err.Clear
 '   End If
End Function

Public Sub fail(haserrors As Boolean, errors As String, reason As String)
    haserrors = True
    errors = errors & IIf(errors = Empty, Empty, vbNewLine) & reason
End Sub

Public Function verifyparameters(parameters, msg As Boolean, minparams As Long, errors As String, ParamArray properties() As Variant)
    Dim count As Long, hasfailed As Boolean
    If UBound(parameters) < minparams Then fail hasfailed, errors, "This function requires " & minparams & " parameters. You specified " & UBound(parameters)
    For count = 0 To UBound(properties)
        Select Case LCase(properties(count))
            Case "bool", "boolean"
                If LCase(getcell(parameters, count + 1)) <> "true" And LCase(getcell(parameters, count + 1)) <> "false" Then
                    fail hasfailed, errors, "Parameter " & count + 1 & " must be either true or false"
                End If
            Case "text", "string", "quotes", "txt"
                If getfromquotes(getcell(parameters, count + 1)) = getcell(parameters, count + 1) Then
                    fail hasfailed, errors, "Parameter " & count + 1 & " must be text enclosed in quotations"
                End If
            Case "number", "numerical", "num"
                If Not isanumber(getcell(parameters, count + 1)) And getcell(parameters, count + 1) <> Empty Then
                    fail hasfailed, errors, "Parameter " & count + 1 & " must be numerical"
                End If
        End Select
    Next
    verifyparameters = Not hasfailed
    If hasfailed And msg Then MsgBox errors, vbCritical, "Equation Evaluation: Parameter Verification"
End Function
