Attribute VB_Name = "MIEEE754"
Option Explicit

Public INDef  As Double
Public posINF As Double
Public negINF As Double
Public NaN    As Double

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal bLength As Long)

Public Sub Init()
    GetINDef INDef
    posINF = GetINF
    negINF = GetINF(-1)
    GetNaN NaN
End Sub

' v ############################## v '    Create functions    ' v ############################## v '
'either with error handling:
Public Function GetINFE(Optional ByVal sign As Long = 1) As Double
Try: On Error Resume Next
    GetINFE = Sgn(sign) / 0
Catch: On Error GoTo 0
End Function

' or without error handling:
Public Function GetINF(Optional ByVal sign As Long = 1) As Double
    Dim L(1 To 2) As Long
    If Sgn(sign) > 0 Then
        L(2) = &H7FF00000
    ElseIf Sgn(sign) < 0 Then
        L(2) = &HFFF00000
    End If
    Call RtlMoveMemory(GetINF, L(1), 8)
End Function

Public Sub GetNaN(ByRef Value As Double)
    Dim L(1 To 2) As Long
    L(1) = 1
    L(2) = &H7FF00000
    Call RtlMoveMemory(Value, L(1), 8)
End Sub

Public Sub GetINDef(ByRef Value As Double)
Try: On Error Resume Next
    Value = 0# / 0#
Catch: On Error GoTo 0
End Sub
' ^ ############################## ^ '    Create functions    ' ^ ############################## ^ '

' v ############################## v '     Bool functions     ' v ############################## v '
Public Function IsINDef(ByRef Value As Double) As Boolean
Try: On Error Resume Next
    IsINDef = (CStr(Value) = CStr(INDef))
Catch: On Error GoTo 0
End Function

Public Function IsNaN(ByRef Value As Double) As Boolean
    Dim b(0 To 7) As Byte
    Dim i As Long
    
    RtlMoveMemory b(0), Value, 8
    
    If (b(7) = &H7F) Or (b(7) = &HFF) Then
        If (b(6) >= &HF0) Then
            For i = 0 To 5
                If b(i) <> 0 Then
                    IsNaN = True
                    Exit Function
                End If
            Next
        End If
    End If
End Function

Public Function IsPosINF(ByVal Value As Double) As Boolean
    IsPosINF = (Value = posINF)
End Function

Public Function IsNegINF(ByVal Value As Double) As Boolean
    IsNegINF = (Value = negINF)
End Function
' ^ ############################## ^ '     Bool functions     ' ^ ############################## ^ '

' v ############################## v '    Output functions    ' v ############################## v '
Public Function INDefToString() As String
    On Error Resume Next
    INDefToString = CStr(INDef)
    On Error GoTo 0
End Function

Public Function NaNToString() As String
    On Error Resume Next
    If App.LogMode = 0 Then
        NaNToString = "1.#QNAN"
    Else
        NaNToString = CStr(NaN)
    End If
    On Error GoTo 0
End Function

Public Function PosINFToString() As String
    PosINFToString = CStr(posINF)
End Function

Public Function NegINFToString() As String
    NegINFToString = CStr(negINF)
End Function
' ^ ############################## ^ '    Output functions    ' ^ ############################## ^ '

' v ############################## v '     Input function     ' v ############################## v '
Public Function Double_TryParse(s As String, Value_out As Double) As Boolean
Try: On Error GoTo Catch
    If Len(s) = 0 Then Exit Function
    s = Replace(s, ",", ".")
    If StrComp(s, "1.#QNAN") = 0 Then
        GetNaN Value_out
    ElseIf StrComp(s, "1.#INF") = 0 Then
        Value_out = GetINF
    ElseIf StrComp(s, "-1.#INF") = 0 Then
        Value_out = GetINF(-1)
    ElseIf StrComp(s, "-1.#IND") = 0 Then
        GetINDef Value_out
    Else
        Value_out = Val(s)
    End If
    Double_TryParse = True
Catch:
End Function
' ^ ############################## ^ '     Input function     ' ^ ############################## ^ '

