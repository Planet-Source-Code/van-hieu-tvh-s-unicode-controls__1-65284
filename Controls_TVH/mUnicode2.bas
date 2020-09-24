Attribute VB_Name = "mUnicode"
'///////////////////////////////////////// Truong Van Hieu ////////////////////////////////////////
'////////////////////////////////// tvhhh2003@yahoo.com /////////////////////////////////////
'//////////////////////////////////// Special for Vietnamese /////////////////////////////////////

Option Explicit

Private VNI_A() As Byte
Private VNI_E() As Byte
Private VNI_I() As Byte
Private VNI_O() As Byte
Private VNI_U() As Byte
Private VNI_Y() As Byte
Private VNI_D() As Byte

Private Const Max_A = 17 * 2
Private Const Max_E = 11 * 2
Private Const Max_I = 5 * 2
Private Const Max_O = 17 * 2
Private Const Max_U = 11 * 2
Private Const Max_Y = 5 * 2
Private Const Max_D = 2

Private t_Dau(1 To 9) As String
Private t_TV(1 To 9) As String

Private Type Size
    cx As Long
    cy As Long
End Type

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hdc As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As Size) As Long

Private Sub setU(T() As Byte, id As Byte, k1 As Byte, k2 As Byte)
    T(id, 0) = k1
    T(id, 1) = k2
End Sub

Public Sub KhoiTao()
    ReDim VNI_A(Max_A, 2)
    ReDim VNI_E(Max_E, 2)
    ReDim VNI_I(Max_I, 2)
    ReDim VNI_O(Max_O, 2)
    ReDim VNI_U(Max_U, 2)
    ReDim VNI_Y(Max_Y, 2)
    ReDim VNI_D(Max_D, 2)
'Set A
    setU VNI_A, 0, 193, 0       ' A'
    setU VNI_A, 1, 192, 0       ' A`
    setU VNI_A, 2, 162, 30      ' A?
    setU VNI_A, 3, 195, 0       ' A~
    setU VNI_A, 4, 160, 30     ' A.
    setU VNI_A, 5, 194, 0       ' A^
    setU VNI_A, 6, 164, 30      ' A^'
    setU VNI_A, 7, 166, 30      ' A^`
    setU VNI_A, 8, 168, 30      ' A^?
    setU VNI_A, 9, 170, 30      ' A^~
    setU VNI_A, 10, 172, 30     ' A^.
    setU VNI_A, 11, 2, 1          '  A(
    setU VNI_A, 12, 174, 30     ' A('
    setU VNI_A, 13, 176, 30     ' A(`
    setU VNI_A, 14, 178, 30     ' A(?
    setU VNI_A, 15, 180, 30     ' A(~
    setU VNI_A, 16, 182, 30     ' A(.
    setU VNI_A, 17, 225, 0      ' a'
    setU VNI_A, 18, 224, 0
    setU VNI_A, 19, 163, 30
    setU VNI_A, 20, 227, 0
    setU VNI_A, 21, 161, 30
    setU VNI_A, 22, 226, 0      'a^
    setU VNI_A, 23, 165, 30
    setU VNI_A, 24, 167, 30
    setU VNI_A, 25, 169, 30
    setU VNI_A, 26, 171, 30
    setU VNI_A, 27, 173, 30
    setU VNI_A, 28, 3, 1            'a(
    setU VNI_A, 29, 175, 30
    setU VNI_A, 30, 177, 30
    setU VNI_A, 31, 179, 30
    setU VNI_A, 32, 181, 30
    setU VNI_A, 33, 183, 30
'Set E
    setU VNI_E, 0, 201, 0       ' E'
    setU VNI_E, 1, 200, 0
    setU VNI_E, 2, 186, 30
    setU VNI_E, 3, 188, 30
    setU VNI_E, 4, 184, 30
    setU VNI_E, 5, 202, 0       ' E^
    setU VNI_E, 6, 190, 30
    setU VNI_E, 7, 192, 30
    setU VNI_E, 8, 194, 30
    setU VNI_E, 9, 196, 30
    setU VNI_E, 10, 198, 30
    setU VNI_E, 11, 233, 0      ' e'
    setU VNI_E, 12, 232, 0
    setU VNI_E, 13, 187, 30
    setU VNI_E, 14, 189, 30
    setU VNI_E, 15, 185, 30
    setU VNI_E, 16, 234, 0      ' e^
    setU VNI_E, 17, 191, 30
    setU VNI_E, 18, 193, 30
    setU VNI_E, 19, 195, 30
    setU VNI_E, 20, 197, 30
    setU VNI_E, 21, 199, 30
'Set I
    setU VNI_I, 0, 205, 0
    setU VNI_I, 1, 204, 0
    setU VNI_I, 2, 200, 30
    setU VNI_I, 3, 40, 1
    setU VNI_I, 4, 202, 30
    setU VNI_I, 5, 237, 0
    setU VNI_I, 6, 236, 0
    setU VNI_I, 7, 201, 30
    setU VNI_I, 8, 41, 1
    setU VNI_I, 9, 203, 30
'Set O
    setU VNI_O, 0, 211, 0       ' O'
    setU VNI_O, 1, 210, 0
    setU VNI_O, 2, 206, 30
    setU VNI_O, 3, 213, 0
    setU VNI_O, 4, 204, 30
    setU VNI_O, 5, 212, 0       ' O^
    setU VNI_O, 6, 208, 30
    setU VNI_O, 7, 210, 30
    setU VNI_O, 8, 212, 30
    setU VNI_O, 9, 214, 30
    setU VNI_O, 10, 216, 30
    setU VNI_O, 11, 160, 1      ' O*
    setU VNI_O, 12, 218, 30
    setU VNI_O, 13, 220, 30
    setU VNI_O, 14, 222, 30
    setU VNI_O, 15, 224, 30
    setU VNI_O, 16, 226, 30
    setU VNI_O, 17, 243, 0      ' o'
    setU VNI_O, 18, 242, 0
    setU VNI_O, 19, 207, 30
    setU VNI_O, 20, 245, 0
    setU VNI_O, 21, 205, 30
    setU VNI_O, 22, 244, 0      ' o^
    setU VNI_O, 23, 209, 30
    setU VNI_O, 24, 211, 30
    setU VNI_O, 25, 213, 30
    setU VNI_O, 26, 215, 30
    setU VNI_O, 27, 217, 30
    setU VNI_O, 28, 161, 1      ' o*
    setU VNI_O, 29, 219, 30
    setU VNI_O, 30, 221, 30
    setU VNI_O, 31, 223, 30
    setU VNI_O, 32, 225, 30
    setU VNI_O, 33, 227, 30
'Set U
    setU VNI_U, 0, 218, 0       ' U'
    setU VNI_U, 1, 217, 0
    setU VNI_U, 2, 230, 30
    setU VNI_U, 3, 104, 1
    setU VNI_U, 4, 228, 30
    setU VNI_U, 5, 175, 1       ' U*
    setU VNI_U, 6, 232, 30
    setU VNI_U, 7, 234, 30
    setU VNI_U, 8, 236, 30
    setU VNI_U, 9, 238, 30
    setU VNI_U, 10, 240, 30
    setU VNI_U, 11, 250, 0      ' u'
    setU VNI_U, 12, 249, 0
    setU VNI_U, 13, 231, 30
    setU VNI_U, 14, 105, 1
    setU VNI_U, 15, 229, 30
    setU VNI_U, 16, 176, 1      ' u*
    setU VNI_U, 17, 233, 30
    setU VNI_U, 18, 235, 30
    setU VNI_U, 19, 237, 30
    setU VNI_U, 20, 239, 30
    setU VNI_U, 21, 241, 30
'Set Y
    setU VNI_Y, 0, 221, 0
    setU VNI_Y, 1, 242, 30
    setU VNI_Y, 2, 246, 30
    setU VNI_Y, 3, 248, 30
    setU VNI_Y, 4, 244, 30
    setU VNI_Y, 5, 253, 0
    setU VNI_Y, 6, 243, 30
    setU VNI_Y, 7, 247, 30
    setU VNI_Y, 8, 249, 30
    setU VNI_Y, 9, 245, 30
'Set D
    setU VNI_D, 0, 16, 1
    setU VNI_D, 1, 17, 1
    
'Set t_Dau
    Dim ts(0 To 4)
    ts(0) = ""
    ts(0) = ts(0) & "A" & GetKt(VNI_A, 5) & GetKt(VNI_A, 11) & "a" & GetKt(VNI_A, 22) & GetKt(VNI_A, 28) 'a
    ts(0) = ts(0) & "E" & GetKt(VNI_E, 5) & "e" & GetKt(VNI_E, 16) 'e
    ts(0) = ts(0) & "I" & "i" 'i
    ts(0) = ts(0) & "O" & GetKt(VNI_O, 5) & GetKt(VNI_O, 11) & "o" & GetKt(VNI_O, 22) & GetKt(VNI_O, 28) 'o
    ts(0) = ts(0) & "U" & GetKt(VNI_U, 5) & "u" & GetKt(VNI_U, 16) 'u
    ts(0) = ts(0) & "Y" & "y" 'y
    Dim i As Byte
    For i = 0 To 4
        t_Dau(i + 1) = ts(0)
        ts(1) = ts(1) & GetKt(VNI_A, 0 + i) & GetKt(VNI_A, 17 + i) _
                        & GetKt(VNI_E, 0 + i) & GetKt(VNI_E, 11 + i) _
                        & GetKt(VNI_O, 0 + i) & GetKt(VNI_O, 17 + i)
        ts(2) = ts(2) & GetKt(VNI_O, 0 + i) & GetKt(VNI_O, 17 + i) _
                        & GetKt(VNI_U, 0 + i) & GetKt(VNI_U, 11 + i)
        ts(3) = ts(3) & GetKt(VNI_A, 0 + i) & GetKt(VNI_A, 17 + i)
    Next i
    t_Dau(6) = "A" & "a" & "E" & "e" & "O" & "o" & ts(1)
    t_Dau(7) = "O" & "o" & "U" & "u" & ts(2)
    t_Dau(8) = "A" & "a" & ts(3)
    t_Dau(9) = "D" & "d"
'Set t_TV
    ts(1) = "": ts(2) = "": ts(3) = ""
    For i = 0 To 4
        ts(0) = ""
        ts(0) = ts(0) & GetKt(VNI_A, 0 + i) & GetKt(VNI_A, 6 + i) & GetKt(VNI_A, 12 + i) & GetKt(VNI_A, 17 + i) & GetKt(VNI_A, 23 + i) & GetKt(VNI_A, 29 + i) 'a
        ts(0) = ts(0) & GetKt(VNI_E, 0 + i) & GetKt(VNI_E, 6 + i) & GetKt(VNI_E, 11 + i) & GetKt(VNI_E, 17 + i) 'e
        ts(0) = ts(0) & GetKt(VNI_I, 0 + i) & GetKt(VNI_I, 5 + i) 'i
        ts(0) = ts(0) & GetKt(VNI_O, 0 + i) & GetKt(VNI_O, 6 + i) & GetKt(VNI_O, 12 + i) & GetKt(VNI_O, 17 + i) & GetKt(VNI_O, 23 + i) & GetKt(VNI_O, 29 + i) 'o
        ts(0) = ts(0) & GetKt(VNI_U, 0 + i) & GetKt(VNI_U, 6 + i) & GetKt(VNI_U, 11 + i) & GetKt(VNI_U, 17 + i) 'u
        ts(0) = ts(0) & GetKt(VNI_Y, 0 + i) & GetKt(VNI_Y, 5 + i) 'y
        t_TV(i + 1) = ts(0)
        ts(1) = ts(1) & GetKt(VNI_A, 6 + i) & GetKt(VNI_A, 23 + i) _
                & GetKt(VNI_E, 6 + i) & GetKt(VNI_E, 17 + i) _
                & GetKt(VNI_O, 6 + i) & GetKt(VNI_O, 23 + i)
        ts(2) = ts(2) & GetKt(VNI_O, 12 + i) & GetKt(VNI_O, 29 + i) _
                & GetKt(VNI_U, 6 + i) & GetKt(VNI_U, 17 + i)
        ts(3) = ts(3) & GetKt(VNI_A, 12 + i) & GetKt(VNI_A, 29 + i)
    Next i
    
    t_TV(6) = GetKt(VNI_A, 5) & GetKt(VNI_A, 22) _
                & GetKt(VNI_E, 5) & GetKt(VNI_E, 16) _
                & GetKt(VNI_O, 5) & GetKt(VNI_O, 22) & ts(1)
    t_TV(7) = GetKt(VNI_O, 11) & GetKt(VNI_O, 28) _
                & GetKt(VNI_U, 5) & GetKt(VNI_U, 16) & ts(2)
    t_TV(8) = GetKt(VNI_A, 11) & GetKt(VNI_A, 28) & ts(3)
    t_TV(9) = GetKt(VNI_D, 0) & GetKt(VNI_D, 1)
End Sub

'Kiem tra phai danh dau hay khong
Private Function isNum(s As String) As Byte
    If Asc(s) >= 49 And Asc(s) <= 57 Then
        isNum = s
    Else
        isNum = 0
    End If
End Function

Private Function GetKt(VNI_T() As Byte, id As Byte) As String
Dim T(0 To 1) As Byte
    T(0) = VNI_T(id, 0)
    T(1) = VNI_T(id, 1)
    GetKt = T
End Function

Private Function DauExist(dau As Byte, s As String) As Byte
    Dim i As Byte
    DauExist = 0
    If dau < 1 Then Exit Function
    For i = 1 To Len(t_Dau(dau))
        If Mid(t_Dau(dau), i, 1) = s Then
            DauExist = i
            Exit Function
        End If
    Next i
End Function

Public Function Text_To_Unicode(s As String) As String
Dim i As Long, j As Long
Dim ts As String
Dim tk As String
Dim dau As Byte
    If Trim(s) = "" Then Text_To_Unicode = "": Exit Function
    ts = Left(s, 1)
    j = 2
    For i = 2 To Len(s)
        tk = Mid(s, i, 1)
        dau = isNum(tk)
        If dau <> 0 And DauExist(dau, Mid(ts, j - 1, 1)) Then
            ts = Mid(ts, 1, Len(ts) - 1) & Mid(t_TV(dau), DauExist(dau, Mid(ts, j - 1, 1)), 1)
            j = j - 1
        Else
            ts = ts & tk
        End If
        j = j + 1
    Next i
    Text_To_Unicode = ts
End Function

Private Function FindDau(VNI_T() As Byte, Max As Byte, kt As String, Replace As String) As String
    Dim i As Byte
    For i = 0 To Max - 1
        If GetKt(VNI_T, i) = kt Then
            If i < Max / 2 Then
                FindDau = Left(Replace, 1)
            Else
                FindDau = Right(Replace, 1)
            End If
            Exit Function
        End If
    Next i
    FindDau = kt
End Function

'Ha`m xo'a da^'u cho Tieng Viet Unicode
Public Function XoaDau(s As String) As String
Dim ts As String
Dim i As Long, j As Long
Dim kt As String
    ts = ""
    For i = 1 To Len(s)
        kt = Mid(s, i, 1)
        j = Asc(kt)
        If j < 64 Or j > 126 Then
            Dim old As String
            old = kt
            kt = FindDau(VNI_A, Max_A, kt, "Aa")
            If old = kt Then kt = FindDau(VNI_E, Max_E, kt, "Ee")
            If old = kt Then kt = FindDau(VNI_I, Max_I, kt, "Ii")
            If old = kt Then kt = FindDau(VNI_O, Max_O, kt, "Oo")
            If old = kt Then kt = FindDau(VNI_U, Max_U, kt, "Uu")
            If old = kt Then kt = FindDau(VNI_Y, Max_Y, kt, "Yy")
            If old = kt Then kt = FindDau(VNI_D, Max_D, kt, "Dd")
        End If
        ts = ts & kt
    Next i
    XoaDau = ts
End Function

Public Function UniTextWidth(hdc As Long, s As String) As Long
Dim sz As Size
    GetTextExtentPoint32 hdc, StrPtr(s), Len(s), sz
    UniTextWidth = sz.cx
End Function

