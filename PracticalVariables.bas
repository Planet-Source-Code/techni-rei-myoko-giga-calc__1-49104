Attribute VB_Name = "PracticalVariables"
Option Explicit

Public Type variable
    var_name As String
    var_type As String
    var_valu As String
    var_dime As String
End Type

Public Const novar As Long = -1 'the variable does not exist
Public Const delimeter As String = "|" 'standard seprator
Public optionexplicit As Boolean

Public Function findentry(strlist, strname As String) As Long
    Dim count As Long
    findentry = novar
    strname = LCase(strname)
    For count = 0 To UBound(strlist)
        If strname = LCase(strlist(count)) Then
            findentry = count
            Exit For
        End If
    Next
End Function

Public Function get_var_index(ByRef varlist() As variable, ByRef varcount As Long, vname As String) As Long
    Dim count As Long
    vname = LCase(vname)
    get_var_index = novar
    For count = 1 To varcount
        If vname = varlist(count).var_name Then
            If varlist(count).var_type = "reference" Then
                get_var_index = get_var_index(varlist, varcount, varlist(count).var_valu)
                Exit For
            Else
                get_var_index = count
                Exit For
            End If
        End If
    Next
End Function

Public Function istype(ByRef varlist() As variable, ByRef varcount As Long, vtype As String, Optional parameter As String) As String
Dim count As Long
    Select Case vtype
        Case "number", "text", "type", "array", "reference"
        Case Else
            count = get_var_index(varlist, varcount, vtype)
            If count = novar Then
                istype = vtype & " has not been declared"
            Else
                If varlist(count).var_type <> "type" Then
                    istype = vtype & " is not a type of variable"
                End If
                If parameter <> Empty Then
                    count = findentry(Split(varlist(count).var_valu, delimeter), LCase(parameter))
                    If count = -1 Then
                        istype = vtype & " does not have " & parameter & " as a parameter. It does however have " & Replace(varlist(count).var_valu, delimeter, ", ")
                    End If
                End If
            End If
    End Select
End Function

Public Function add_var(ByRef varlist() As variable, ByRef varcount As Long, vname As String, vtype As String, Optional vvalu As String) As String
    Dim count As Long, temp As String, arraydime As String
    vtype = Trim(LCase(vtype))
    temp = istype(varlist, varcount, vtype)
    If temp <> Empty Then
        add_var = temp
        Exit Function
    End If
    count = get_var_index(varlist, varcount, vname)
    If count = novar Then
        varcount = varcount + 1
        ReDim Preserve varlist(1 To varcount)
        count = varcount
    End If
    arraydime = LCase(vname)
    varlist(count).var_name = arraydime
    If vtype = "array" Then
        arraydime = vvalu
        vvalu = Empty
    End If
    varlist(count).var_type = vtype
    If vtype = "type" Then vvalu = LCase(Replace(vvalu, " ", Empty))
    varlist(count).var_valu = vvalu
    varlist(count).var_dime = arraydime
    add_var = vname & "=" & varlist(count).var_valu
End Function

Public Function extractbrackets(text As String) As String
    extractbrackets = Mid(text, InStr(text, "(") + 1, Len(text) - InStr(text, "(") - 1)
End Function

Public Function get_var(ByRef varlist() As variable, ByRef varcount As Long, vname As String, Optional default As String) As String
    Dim temp As Long
    temp = get_var_index(varlist, varcount, vname)
    If temp > novar Then
        get_var = vname & "=" & varlist(temp).var_valu
    Else
        get_var = vname & "=" & default
    End If
End Function

Public Function forcearray(text As String, delimeter As String, max As Long) As String
    Dim count As Long
    count = countchars(text, delimeter)  'make sure its the right dimensions
    forcearray = text
    If count < max Then 'make the variable the proper number of parameters
        forcearray = text & String(max - count, delimeter)
    End If
End Function

Public Function get_property_index(ByRef varlist() As variable, ByRef varcount As Long, vtype As String, parameter As String) As Long
    Dim count As Long
    get_property_index = novar
    If LCase(parameter) <> "value" Then
    count = get_var_index(varlist, varcount, vtype)
    If count > novar Then
        If varlist(count).var_type = "type" Then
            get_property_index = findentry(Split(varlist(count).var_valu, delimeter), LCase(Trim(parameter)))
        End If
    End If
    End If
End Function

Public Function countchars(text As String, char As String) As Long
Dim temp As Long, temp2 As Long
temp2 = InStr(1, text, char, vbTextCompare)
Do Until temp2 = 0
    temp = temp + 1
    temp2 = InStr(temp2 + 1, text, char, vbTextCompare)
Loop
countchars = temp
End Function
Public Function extractproperty(text As String) As String
    If InStr(text, "=") > 0 Then
        extractproperty = Right(text, Len(text) - InStr(text, "="))
    End If
End Function

Public Function getparameter(command, parameter As Long, Optional default As String, Optional allelse As Boolean = False) As String
    On Error Resume Next
    getparameter = default
    If parameter >= LBound(command) And parameter <= UBound(command) Then
        getparameter = command(parameter)
        If allelse Then
            Dim temp As Long, tempstr As String
            For temp = parameter To UBound(command)
                If tempstr = Empty Then
                    tempstr = command(temp)
                Else
                    tempstr = tempstr & " " & command(temp)
                End If
            Next
        End If
    End If
End Function
Public Function get_array_cell(text As String, delimeter As String, Index As Long) As String
    get_array_cell = getparameter(Split(text, delimeter), Index)
End Function

Public Function set_property(ByRef varlist() As variable, ByRef varcount As Long, vname As String, Optional vproperty As String = "value", Optional vvalue As String = Empty) As String
    'leave vvalue empty to return the property without changing it
    Dim tempstr() As String, tempstr2() As String, count As Long, count2 As Long, count3 As Long, count4 As Long
    count = get_var_index(varlist, varcount, vname)
    If count > novar Then 'if variable vname exists
        count2 = get_var_index(varlist, varcount, varlist(count).var_type)
        If vproperty <> "value" Then
            If count2 > novar Then 'if variable type of vname exists
                If varlist(count2).var_type = "type" Then
                    tempstr2 = Split(varlist(count2).var_valu, delimeter) 'contents of type
                    count3 = findentry(tempstr2, vproperty)
                    If count3 > novar Then 'if property of variable type of vname exists
                        varlist(count).var_valu = forcearray(varlist(count).var_valu, delimeter, UBound(tempstr2))
                        tempstr = Split(varlist(count).var_valu, delimeter) 'contents of var
                        If vvalue <> Empty Then
                            tempstr(count3) = vvalue
                            varlist(count).var_valu = Join(tempstr, delimeter)
                        End If
                        set_property = vname & "." & vproperty & "=" & tempstr(count3)
                    Else
                        set_property = varlist(count).var_type & " does not have a property by the name of " & vproperty & ". It does however have " & extractproperty(Replace(varlist(count2).var_valu, delimeter, ", "))
                    End If
                Else
                    set_property = varlist(count2).var_type & " was declared as a " & varlist(count2).var_type & ". It needs to be a 'type'"
                End If
            Else
                set_property = "The type this variable is defined as has not been declared. " & vname & " was declared as " & varlist(count).var_type
            End If
        Else
            If vvalue <> Empty Then
                varlist(count).var_valu = vvalue
            End If
            set_property = vname & "." & vproperty & "=" & varlist(count).var_valu
        End If
    Else
        If vproperty = "value" And optionexplicit = False Then
            set_property = add_var(varlist, varcount, vname, "text", vvalue)
        Else
            set_property = vname & " has not been declared"
        End If
    End If
End Function

Public Function declare_var(ByRef varlist() As variable, ByRef varcount As Long, command) As String
'On Error Resume Next
declare_var = "Not enough parameters. Example: Declare <variable name> as <variable type> [with <parameters>]"
Dim vname As String, vtype As String, vvalu As String, vparam As String, temp As Long, tempstring As String, currparam As Long, isanarray As Boolean, arrayindex As String

    vname = LCase(getparameter(command, 1))  'name of variable
    vname = Replace(vname, "[", "(")
    vname = Replace(vname, "]", ")")
    vtype = "text"
    vparam = "value"
    
    If InStr(vname, ".") > 0 Then 'make sure its not a parameter of a var you're declaring/setting
        vparam = Trim(Right(vname, Len(vname) - InStr(vname, ".")))
        vname = Trim(Left(vname, InStr(vname, ".") - 1))
    End If
    
    currparam = 2
    If LCase(getparameter(command, 2)) = "as" Then
        vtype = LCase(Trim(getparameter(command, 3, "text"))) 'type of variable
        tempstring = istype(varlist, varcount, vtype)
        If tempstring <> Empty Then
            declare_var = tempstring
            Exit Function
        End If
        currparam = 4
    End If
    
    If LCase(getparameter(command, currparam)) = "with" Or LCase(getparameter(command, currparam)) = "to" Then
        vvalu = getparameter(command, currparam + 1, Empty, True)
    End If
    
    If vtype = "array" Or countchars(vname, "(") > 0 Then isanarray = True
    
    If isanarray = False Then
        Select Case LCase(getparameter(command, 0))
            Case "declare": declare_var = add_var(varlist, varcount, vname, vtype, vvalu)
            Case "set":     declare_var = set_property(varlist, varcount, vname, vparam, vvalu)
            Case "get":     declare_var = set_property(varlist, varcount, vname, vparam)
            Case Else:      declare_var = "This function was not set up to run the command " & getparameter(command, 0)
        End Select
    Else
        Select Case LCase(getparameter(command, 0))
            Case "declare": declare_var = add_var(varlist, varcount, vname, vtype, vvalu)
            Case "set":     declare_var = get_or_set_array_property(varlist, varcount, False, vname, vparam, vvalu)
            Case "get":     declare_var = get_or_set_array_property(varlist, varcount, True, vname, vparam, vvalu)
            Case Else:      declare_var = "This function was not set up to run the command " & getparameter(command, 0)
        End Select
    End If
End Function

Public Function set_array_index(text As String, ByVal indexes As String, Optional value As String) As String
    Dim tempstr() As String, tempstring As String, count As Long
    If InStr(indexes, ",") = 0 Then
        'has received single dimensional array, split it up and process
        tempstring = get_array_cell(indexes, " ", 1) ' <index desired> <delimeter>
        count = Val(get_array_cell(indexes, " ", 0))
        text = forcearray(text, tempstring, count)
        If text = Empty Then
            set_array_index = String(count, tempstring) & value
        Else
            text = forcearray(text, tempstring, count)
            tempstr = Split(text, tempstring)  'split by delimeter
            tempstr(count) = value
            set_array_index = Join(tempstr, tempstring)
        End If
    Else
        'has received multidimensional array, split it up, pass the cell matching the first index along removing the first index
        tempstring = Left(indexes, InStr(indexes, ",") - 1) '(first index)
        indexes = Right(indexes, Len(indexes) - InStr(indexes, ",")) 'remove first index
        count = Val(get_array_cell(tempstring, " ", 0))
        If text = Empty Then
            set_array_index = String(count, get_array_cell(tempstring, " ", 1)) & value
        Else
            text = forcearray(text, get_array_cell(tempstring, " ", 1), count)
            tempstr = Split(text, get_array_cell(tempstring, " ", 1)) 'split text by delimeter from first index
            tempstr(count) = set_array_index(tempstr(count), indexes, value)
            set_array_index = Join(tempstr, get_array_cell(tempstring, " ", 1))
        End If
    End If
End Function

Public Function get_array_index(ByVal text As String, ByVal indexes As String, Optional default As String) As String
    Dim tempstr() As String, tempstring As String, count As Long, delim As String
    Do Until InStr(indexes, ",") = 0
        tempstring = Left(indexes, InStr(indexes, ",") - 1) '(first index)
        indexes = Right(indexes, Len(indexes) - InStr(indexes, ",")) 'remove first index
        count = Val(get_array_cell(tempstring, " ", 0)) 'get index desired
        delim = get_array_cell(tempstring, " ", 1) 'get delimeter
        text = forcearray(text, delim, count) ' make sure its the correct dimension
        tempstr = Split(text, delim) 'split text by delimeter from first index
        text = getparameter(tempstr, count, default) 'get the parameter desired
    Loop
    get_array_index = get_array_cell(text, get_array_cell(indexes, " ", 1), Val(get_array_cell(indexes, " ", 0)))
End Function

Public Function get_or_set_array_property(ByRef varlist() As variable, ByRef varcount As Long, GETorSET As Boolean, ByVal vname As String, Optional property As String = "value", Optional value As String) As String
    'oh goodie, the hard part. extractbrackets, set_array_index, get_array_index
    Dim varindex As Long, varray As String, vproperty As Long, vdelimeters As String, vtype As String, count As Long, tempstr() As String, temp As Long, temparray As String
    varray = extractbrackets(vname)
    temparray = varray
    vname = Left(vname, InStr(vname, "(") - 1)
    count = get_var_index(varlist, varcount, vname)
    If count > novar Then
        varindex = count
        vtype = get_array_cell(varlist(count).var_dime, " ", 1)
        vdelimeters = get_array_cell(varlist(count).var_dime, " ", 0)
        vproperty = get_property_index(varlist, varcount, vtype, property)
        
        varray = forcearray(varray, ",", 1)
        tempstr = Split(varray, ",")
        For count = 0 To UBound(tempstr)
            If tempstr(count) <> Empty Then
                tempstr(count) = tempstr(count) & " " & Mid(vdelimeters, count + 1, 1)
            End If
        Next
        varray = Join(tempstr, ",")
        If vproperty > novar Then varray = varray & IIf(Right(varray, 1) = ",", Empty, ",") & vproperty & " " & delimeter
        
        get_or_set_array_property = vname & vbTab & varray & vbTab & value
        If GETorSET Then
            get_or_set_array_property = vname & "(" & temparray & ")" & "." & property & "=" & get_array_index(varlist(varindex).var_valu, varray, value)
        Else
            varlist(varindex).var_valu = set_array_index(varlist(varindex).var_valu, varray, value)
            get_or_set_array_property = vname & "(" & temparray & ")" & "." & property & "=" & varlist(varindex).var_valu
        End If
    Else
        get_or_set_array_property = vname & " has not been declared, and arrays can not be auto declared"
    End If
End Function

Public Function varexists(ByRef varlist() As variable, ByRef varcount As Long, ByVal VarName As String, Optional ByRef temp As Long) As Boolean
    If InStr(VarName, "[") > 0 Then VarName = Left(VarName, InStr(VarName, "[") - 1)
    If InStr(VarName, "(") > 0 Then VarName = Left(VarName, InStr(VarName, "(") - 1)
    If InStr(VarName, ".") > 0 Then VarName = Left(VarName, InStr(VarName, ".") - 1)
    temp = get_var_index(varlist, varcount, VarName)
    varexists = temp <> novar
End Function

Public Sub deletevar(ByRef varlist() As variable, ByRef varcount As Long, VarName As String)
    Dim temp As Long, temp2 As Long
    If varexists(varlist, varcount, VarName, temp) Then
        For temp2 = temp + 1 To varcount
            varlist(temp2 - 1) = varlist(temp2)
        Next
        varcount = varcount - 1
        If varcount = 0 Then
            ReDim varlist(0)
        Else
            ReDim Preserve varlist(1 To varcount)
        End If
    End If
End Sub

Public Function enumTypes(ByRef varlist() As variable, ByRef varcount As Long)
    Dim temp As Long, typelist As String
    typelist = "text|number|type|array"
    For temp = 1 To varcount
        If varlist(temp).var_type = "type" Then
            typelist = typelist & "|" & varlist(temp).var_name
        End If
    Next
    enumTypes = typelist
End Function

Public Function savevars(ByRef varlist() As variable, ByRef varcount As Long, filename As String) As Boolean
    On Error Resume Next
    Dim tempfile As Long, temp As Long
    tempfile = FreeFile
    Open filename For Output As tempfile
        For temp = 1 To varcount
            Write #tempfile, varlist(temp).var_dime, varlist(temp).var_name, varlist(temp).var_type, varlist(temp).var_valu
        Next
    Close tempfile
    savevars = True
End Function

Public Function loadvars(ByRef varlist() As variable, ByRef varcount As Long, filename As String) As Boolean
    On Error Resume Next
    varcount = 0
    Dim tempfile As Long, tempvar As variable
    tempfile = FreeFile
    If Dir(filename) <> Empty Then
        Open filename For Input As tempfile
            Do Until EOF(tempfile)
                Input #tempfile, tempvar.var_dime, tempvar.var_name, tempvar.var_type, tempvar.var_valu
                varcount = varcount + 1
                If varcount = 1 Then
                    ReDim varlist(1 To 1)
                Else
                    ReDim Preserve varlist(1 To varcount)
                End If
                varlist(varcount) = tempvar
            Loop
        Close tempfile
        loadvars = True
    End If
End Function
