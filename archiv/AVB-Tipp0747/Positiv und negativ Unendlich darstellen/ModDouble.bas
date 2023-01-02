Attribute VB_Name = "ModDouble"
Option Explicit

Public posINF As Double
Public negINF As Double
Public NaN    As Double

Public Declare Sub RtlMoveMemory Lib "kernel32" ( _
    ByRef pDst As Any, ByRef pSrc As Any, ByVal bLength As Long)

Public Sub Init()
    posINF = GetINF
    negINF = GetINF(-1)
    Call GetNaN(NaN)
End Sub

'entweder mit Fehlerbehandlung:
Public Function GetINFE(Optional ByVal sign As Long = 1) As Double
    On Error Resume Next
    GetINFE = Sgn(sign) / 0
    On Error GoTo 0
End Function

' oder ohne Fehlerbehandlung:
Public Function GetINF(Optional ByVal sign As Long = 1) As Double
    Dim L(1 To 2) As Long
    If Sgn(sign) > 0 Then
        L(2) = &H7FF00000
    ElseIf Sgn(sign) < 0 Then
        L(2) = &HFFF00000
    End If
    Call RtlMoveMemory(GetINF, L(1), 8)
End Function

Public Sub GetNaN(ByRef DblVal As Double)
    Dim L(1 To 2) As Long
    L(1) = 1
    L(2) = &H7FF00000
    Call RtlMoveMemory(DblVal, L(1), 8)
End Sub

Public Function IsNaN(ByRef DblVal As Double) As Boolean
    Dim b(0 To 7) As Byte
    Dim i As Long
    
    Call RtlMoveMemory(b(0), DblVal, 8)
    
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

Public Function IsPosINF(ByVal DblVal As Double) As Boolean
    IsPosINF = (DblVal = posINF)
End Function

Public Function IsNegINF(ByVal DblVal As Double) As Boolean
    IsNegINF = (DblVal = negINF)
End Function

Public Function NaNToString() As String
    On Error Resume Next
    NaNToString = CStr(NaN)
    On Error GoTo 0
End Function

Public Function PosINFToString() As String
    PosINFToString = CStr(posINF)
End Function

Public Function NegINFToString() As String
    NegINFToString = CStr(negINF)
End Function

Public Sub DoubleParse(d As Double, StrVal As String)
    If Len(StrVal) > 0 Then
        StrVal = Replace(StrVal, ",", ".")
        If StrComp(StrVal, "1.#QNAN") = 0 Then
            Call GetNaN(d)
        ElseIf StrComp(StrVal, "1.#INF") = 0 Then
            d = GetINF
        ElseIf StrComp(StrVal, "-1.#INF") = 0 Then
            d = GetINF(-1)
        Else
            d = CDbl(StrVal)
        End If
    End If
End Sub
