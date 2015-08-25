VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FoxproMemFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public data As Scripting.Dictionary

Public Sub Class_Initialize()
  Set data = New Scripting.Dictionary()
End Sub

Public Sub Init(ByVal filename As String)
    Dim i As Integer
    Dim s As String
    Dim v As Variant
    
    i = FreeFile()
    Open filename For Binary Access Read As i Len = 255
    
    While Not EOF(i)
        v = ReadInput(s, i)
        If Not IsNull(v) Then
            data.Add s, v
        End If
    Wend
    
    Close #i
End Sub

Private Function ReadInput(ByRef s As String, ByVal i As Integer) As Variant
    Dim vname As String * 11
    Dim vlname As String
    Dim vtype As String * 2
    Dim vlen As Long
    Dim padding As String * 8, zstr As String * 7
    
    Get #i, , vname
    Get #i, , vtype
    Get #i, , vlen
    Get #i, , padding
    Get #i, , zstr
    
    If EOF(i) Then
        ReadInput = Null
        Exit Function
    End If
    
    If StrComp(padding, Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)) <> 0 Then
        Error 1001
    End If
    
    If StrComp(zstr, "z" & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(0)) <> 0 Then
        Error 1002
    End If
    
    ' must change the order?
    vlen = (vlen And 255) * &H1000000 + _
                ((vlen / &H100) And &HFF) * &H10000 + _
                ((vlen / &H10000) And &HFF) * &H100 + _
                ((vlen / &H1000000) And &HFF) * &H1
            

    vlname = Mid(vname, 1, InStr(vname, Chr(0)) - 1)
    'MsgBox StrComp(vname, "XZSTP")
    'MsgBox vlname
    'MsgBox vtype
    'MsgBox vlen
    
    Dim vByteArray() As Byte
    Dim vNullChar As Byte
    Dim vString As String
    Dim j As Integer
    
    If StrComp("", vlname) = 0 Then
        ' need to read the variable name
        ' coincidentally the same format as VBA's?
        If StrComp(LCase(vtype), vtype) <> 0 Then
            Error 1000
        End If
        
        Dim vnamelen As Integer
        
        Get #i, , vnamelen
        ReDim vByteArray(1 To vnamelen) As Byte
        Get #i, , vByteArray
        vlname = ""
        For j = 1 To vnamelen
            vlname = vlname & Chr(vByteArray(j))
        Next
    End If
    
    Dim vltype As String
    vltype = UCase(Left(vtype, 1))
    If StrComp(vltype, "N") = 0 Then
        Dim vDouble As Double
        If vlen <> 10 Then
        Error 1003
        End If
        Get #i, , vDouble
        ReadInput = vDouble
    ElseIf StrComp(vltype, "C") = 0 Then
        
        ReDim vByteArray(1 To vlen) As Byte
        Get #i, , vByteArray
        For j = 1 To vlen - 1
            vString = vString & Chr(vByteArray(j))
        Next
        
        vNullChar = vByteArray(vlen)
        If vNullChar <> 0 Then
            Error 1005
        End If
        ReadInput = vString
    ElseIf StrComp(vltype, "L") = 0 Then
        Dim vByte As Byte
        Get #i, , vByte
        ReadInput = CBool(vByte)
    ElseIf StrComp(vltype, "D") = 0 Then
        Dim vDate As Date
        Get #i, , vDate
        vDate = DateAdd("D", -2415019, vDate)
        ReadInput = vDate
    Else
        MsgBox "unknown type " & vltype
        Error 1004
    End If
    
    s = vlname
    
'    Close #1
End Function

