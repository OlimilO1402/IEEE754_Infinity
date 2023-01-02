VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "IEEE-754"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.CommandButton BtnTestIEEE754 
      Caption         =   "Test IEEE754"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnTestIEEE754_Click()
    
    'MIEEE754.Init <- call before executing any function from Module MIEEE754, will be done in "Sub Main"
    
    Debug_Print "Indefinite:        " & INDefToString
    
    Debug_Print "Not a Number:       " & NaNToString
    
    Debug_Print "Positive Infinity:  " & PosINFToString
    
    Debug_Print "Negative Infinity: " & NegINFToString
    
    Debug_Print "? IsINDef(INDef):   " & CStr(IsINDef(INDef))
    
    Debug_Print "? IsNaN(NaN):       " & CStr(IsNaN(NaN))
    
    Debug_Print "? IsPosINF(posINF): " & CStr(IsPosINF(posINF))
    
    Debug_Print "? IsNegINF(negINF): " & CStr(IsNegINF(negINF))
    
    Debug_Print "? IsINDef(1#):      " & CStr(IsINDef(1#))
    
    Debug_Print "? IsNaN(1#):        " & CStr(IsNaN(1#))
    
    Debug_Print "? IsPosINF(1#):     " & CStr(IsPosINF(1#))
    
    Debug_Print "? IsNegINF(1#):     " & CStr(IsNegINF(1#))
    
    Dim d As Double
    Dim s As String
    
    s = INDefToString
    If Double_TryParse(s, d) Then
        Debug_Print "Double_TryParse:   " & CStr(d)
    End If
    d = 0
    
    s = PosINFToString
    If Double_TryParse(s, d) Then
        Debug_Print "Double_TryParse:    " & CStr(d)
    End If
    d = 0
    
    s = NegINFToString
    If Double_TryParse(s, d) Then
        Debug_Print "Double_TryParse:   " & CStr(d)
    End If
    d = 0
    
    s = NaNToString
    If Double_TryParse(s, d) Then
        On Error Resume Next
        Debug_Print "Double_TryParse:    " & CStr(d) 'works only compiled to pe-exe!
    End If
    d = 0
    
End Sub

Private Sub Debug_Print(ByVal s As String)
    Text1.Text = Text1.Text & s & vbCrLf
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - Text1.Top
    If W > 0 And H > 0 Then Text1.Move 0, Text1.Top, W, H
End Sub
