Attribute VB_Name = "mColor"
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
'//////////////////////////////////// Special for Vietnamese /////////////////////////////////////
Option Explicit

Public Sub GradientColor(Color1 As Long, Color2 As Long, Depth As Integer, Result() As Long)
Dim VR, VG, VB As Single
Dim r, G, B, R2, G2, B2 As Integer
Dim t As Long
    t = (Color1 And 255)
    r = t And 255
    t = Int(Color1 / 256)
    G = t And 255
    t = Int(Color1 / 65536)
    B = t And 255
    t = (Color2 And 255)
    R2 = t And 255
    t = Int(Color2 / 256)
    G2 = t And 255
    t = Int(Color2 / 65536)
    B2 = t And 255
    VR = Abs(r - R2) / (Depth)
    VG = Abs(G - G2) / (Depth)
    VB = Abs(B - B2) / (Depth)
    If R2 < r Then VR = -VR
    If G2 < G Then VG = -VG
    If B2 < B Then VB = -VB
    ReDim Result(Depth)
    For t = 0 To Depth
        R2 = r + VR * t
        G2 = G + VG * t
        B2 = B + VB * t
        Result(t) = RGB(R2, G2, B2)
    Next t
End Sub

Sub SplitGradientColor(CS() As Long, Length As Long, Result() As Long)
Dim i As Long
Dim j As Long
Dim tn As Integer
Dim td As Long
Dim c1() As Long
Dim t As Long
    ReDim Result(0)
    tn = Length \ UBound(CS)
    td = Length Mod UBound(CS)
    For i = 0 To UBound(CS) - 1
        If td > 1 And i < td - 1 Then
            GradientColor2 CS(i), CS(i + 1), tn + 2, c1
        ElseIf td = 0 And i = UBound(CS) - 1 Then
            GradientColor2 CS(i), CS(i + 1), tn, c1
        Else
            GradientColor2 CS(i), CS(i + 1), tn + 1, c1
        End If
        t = UBound(Result)
        ReDim Preserve Result(t + UBound(c1))
        For j = 0 To UBound(c1)
            Result(j + t) = c1(j)
        Next j
    Next i
End Sub

Public Sub GradientColor2(Color1 As Long, Color2 As Long, Depth As Integer, Result() As Long)
Dim VR, VG, VB As Single
Dim r, G, B, R2, G2, B2 As Integer
Dim t As Long
    If Depth < 1 Then Exit Sub
    If Depth = 1 Then
        ReDim Result(0)
        Result(0) = Color1
        Exit Sub
    End If
    t = (Color1 And 255)
    r = t And 255
    t = Int(Color1 / 256)
    G = t And 255
    t = Int(Color1 / 65536)
    B = t And 255
    t = (Color2 And 255)
    R2 = t And 255
    t = Int(Color2 / 256)
    G2 = t And 255
    t = Int(Color2 / 65536)
    B2 = t And 255
    VR = Abs(r - R2) / (Depth - 1)
    VG = Abs(G - G2) / (Depth - 1)
    VB = Abs(B - B2) / (Depth - 1)
    If R2 < r Then VR = -VR
    If G2 < G Then VG = -VG
    If B2 < B Then VB = -VB
    ReDim Result(Depth - 1)
    For t = 0 To Depth - 1
        R2 = r + VR * t
        G2 = G + VG * t
        B2 = B + VB * t
        Result(t) = RGB(R2, G2, B2)
    Next t
End Sub

