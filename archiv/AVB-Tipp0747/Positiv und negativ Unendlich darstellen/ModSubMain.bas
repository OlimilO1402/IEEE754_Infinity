Attribute VB_Name = "ModSubMain"
Option Explicit

Sub Main()
    Call ModDouble.Init
    
    MsgBox "Not a Number: " & NaNToString
    
    MsgBox "Positive Infinity: " & PosINFToString
    
    MsgBox "Negative Infinity: " & NegINFToString
    
    MsgBox "? IsNaN(NaN): " & CStr(IsNaN(NaN))
    
    MsgBox "? IsPosINF(posINF): " & CStr(IsPosINF(posINF))
    
    MsgBox "? IsNegINF(negINF): " & CStr(IsNegINF(negINF))
    
    MsgBox "? IsNaN(1#): " & CStr(IsNaN(1#))
    
    MsgBox "? IsPosINF(1#): " & CStr(IsPosINF(1#))
    
    MsgBox "? IsNegINF(1#): " & CStr(IsNegINF(1#))
    
    Dim d As Double
    Dim s As String
    s = NaNToString
    Call DoubleParse(d, s)
    MsgBox CStr(d)
    
    s = PosINFToString
    Call DoubleParse(d, s)
    MsgBox CStr(d)
    
    s = NegINFToString
    Call DoubleParse(d, s)
    MsgBox CStr(d)
End Sub
