Attribute VB_Name = "mUnicode"
'///////////////////////////////////////// Code by: Truong Van Hieu ///////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////// tvhhh2003@yahoo.com ////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////// Special for Vietnamese ////////////////////////////////////////////////////////////////////////

Option Explicit

Private Type Size
    cx As Long
    cy As Long
End Type

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hdc As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As Size) As Long

Function UniTextWidth(hdc As Long, s As String) As Long
Dim sz As Size
    GetTextExtentPoint32 hdc, StrPtr(s), Len(s), sz
    UniTextWidth = sz.cx
End Function

'cSkip la` ky' tu ko duoc hien thi, dung de ngan cach khong bo dau. Mac dinh: "¦" (Alt+179)
'Vi du:
'    +s=" A1 " --> VNI_Unicode=" A' "
'    +s=" A¦1 " --> VNI_Unicode=" A1 "
Function VNI_Unicode(s As String, Optional cSkip = "¦") As String
Dim i&, j&, kq$, C$
    If Len(s) = 0 Then VNI_Unicode = "": Exit Function
    kq = Left(s, 1)
    For i = 2 To Len(s)
        j = Len(kq) - 1
        Select Case Mid(s, i, 1)
            Case 1:
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(225)
                    Case "A": C = ChrW(193)
                    Case ChrW(259): C = ChrW(7855) 'a(
                    Case ChrW(258): C = ChrW(7854) 'A(
                    Case ChrW(226): C = ChrW(7845) 'a^
                    Case ChrW(194): C = ChrW(7844) 'A^
                    
                    Case "e": C = ChrW(233)
                    Case "E": C = ChrW(201)
                    Case ChrW(234): C = ChrW(7871) 'e^
                    Case ChrW(202): C = ChrW(7870) 'E^
                
                    Case "i": C = ChrW(237)
                    Case "I": C = ChrW(205)
                    
                    Case "o": C = ChrW(243)
                    Case "O": C = ChrW(211)
                    Case ChrW(244): C = ChrW(7889) 'o^
                    Case ChrW(212): C = ChrW(7888) 'O^
                    Case ChrW(417): C = ChrW(7899) 'o*
                    Case ChrW(416): C = ChrW(7898) 'O*
                    
                    Case "u": C = ChrW(250)
                    Case "U": C = ChrW(218)
                    Case ChrW(432): C = ChrW(7913) 'u*
                    Case ChrW(431): C = ChrW(7912) 'U*
                    
                    Case "y": C = ChrW(253)
                    Case "Y": C = ChrW(221)
                    
                    Case Else
                        j = j + 1
                        C = "1"
                End Select
            Case 2:
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(224)
                    Case "A": C = ChrW(192)
                    Case ChrW(259): C = ChrW(7857) 'a(
                    Case ChrW(258): C = ChrW(7856) 'A(
                    Case ChrW(226): C = ChrW(7847) 'a^
                    Case ChrW(194): C = ChrW(7846) 'A^
                    
                    Case "e": C = ChrW(232)
                    Case "E": C = ChrW(200)
                    Case ChrW(234): C = ChrW(7873) 'e^
                    Case ChrW(202): C = ChrW(7872) 'E^
                
                    Case "i": C = ChrW(236)
                    Case "I": C = ChrW(204)
                    
                    Case "o": C = ChrW(242)
                    Case "O": C = ChrW(210)
                    Case ChrW(244): C = ChrW(7891) 'o^
                    Case ChrW(212): C = ChrW(7890) 'O^
                    Case ChrW(417): C = ChrW(7901) 'o*
                    Case ChrW(416): C = ChrW(7900) 'O*
                    
                    Case "u": C = ChrW(249)
                    Case "U": C = ChrW(217)
                    Case ChrW(432): C = ChrW(7915) 'u*
                    Case ChrW(431): C = ChrW(7914) 'U*
                    
                    Case "y": C = ChrW(7923)
                    Case "Y": C = ChrW(7922)
                    
                    Case Else
                        j = j + 1
                        C = "2"
                End Select
            Case 3:
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(7843)
                    Case "A": C = ChrW(7842)
                    Case ChrW(259): C = ChrW(7859) 'a(
                    Case ChrW(258): C = ChrW(7858) 'A(
                    Case ChrW(226): C = ChrW(7849) 'a^
                    Case ChrW(194): C = ChrW(7848) 'A^
                    
                    Case "e": C = ChrW(7867)
                    Case "E": C = ChrW(7866)
                    Case ChrW(234): C = ChrW(7875) 'e^
                    Case ChrW(202): C = ChrW(7874) 'E^
                
                    Case "i": C = ChrW(7881)
                    Case "I": C = ChrW(7880)
                    
                    Case "o": C = ChrW(7887)
                    Case "O": C = ChrW(7886)
                    Case ChrW(244): C = ChrW(7893) 'o^
                    Case ChrW(212): C = ChrW(7892) 'O^
                    Case ChrW(417): C = ChrW(7903) 'o*
                    Case ChrW(416): C = ChrW(7902) 'O*
                    
                    Case "u": C = ChrW(7911)
                    Case "U": C = ChrW(7910)
                    Case ChrW(432): C = ChrW(7917) 'u*
                    Case ChrW(431): C = ChrW(7916) 'U*
                    
                    Case "y": C = ChrW(7927)
                    Case "Y": C = ChrW(7926)
                    
                    Case Else
                        j = j + 1
                        C = "3"
                End Select
            Case 4:
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(227)
                    Case "A": C = ChrW(195)
                    Case ChrW(259): C = ChrW(7861) 'a(
                    Case ChrW(258): C = ChrW(7860) 'A(
                    Case ChrW(226): C = ChrW(7851) 'a^
                    Case ChrW(194): C = ChrW(7850) 'A^
                    
                    Case "e": C = ChrW(7869)
                    Case "E": C = ChrW(7868)
                    Case ChrW(234): C = ChrW(7877) 'e^
                    Case ChrW(202): C = ChrW(7876) 'E^
                
                    Case "i": C = ChrW(297)
                    Case "I": C = ChrW(296)
                    
                    Case "o": C = ChrW(245)
                    Case "O": C = ChrW(213)
                    Case ChrW(244): C = ChrW(7895) 'o^
                    Case ChrW(212): C = ChrW(7894) 'O^
                    Case ChrW(417): C = ChrW(7905) 'o*
                    Case ChrW(416): C = ChrW(7904) 'O*
                    
                    Case "u": C = ChrW(361)
                    Case "U": C = ChrW(360)
                    Case ChrW(432): C = ChrW(7919) 'u*
                    Case ChrW(431): C = ChrW(7918) 'U*
                    
                    Case "y": C = ChrW(7929)
                    Case "Y": C = ChrW(7928)
                    
                    Case Else
                        j = j + 1
                        C = "4"
                End Select
            Case 5:
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(7841)
                    Case "A": C = ChrW(7840)
                    Case ChrW(259): C = ChrW(7863) 'a(
                    Case ChrW(258): C = ChrW(7862) 'A(
                    Case ChrW(226): C = ChrW(7853) 'a^
                    Case ChrW(194): C = ChrW(7852) 'A^
                    
                    Case "e": C = ChrW(7865)
                    Case "E": C = ChrW(7864)
                    Case ChrW(234): C = ChrW(7879) 'e^
                    Case ChrW(202): C = ChrW(7878) 'E^
                
                    Case "i": C = ChrW(7883)
                    Case "I": C = ChrW(7882)
                    
                    Case "o": C = ChrW(7885)
                    Case "O": C = ChrW(7884)
                    Case ChrW(244): C = ChrW(7897) 'o^
                    Case ChrW(212): C = ChrW(7896) 'O^
                    Case ChrW(417): C = ChrW(7907) 'o*
                    Case ChrW(416): C = ChrW(7906) 'O*
                    
                    Case "u": C = ChrW(7909)
                    Case "U": C = ChrW(7908)
                    Case ChrW(432): C = ChrW(7921) 'u*
                    Case ChrW(431): C = ChrW(7920) 'U*
                    
                    Case "y": C = ChrW(7925)
                    Case "Y": C = ChrW(7924)
                    
                    Case Else
                        j = j + 1
                        C = "5"
                End Select
            Case 6:
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(226)
                    Case "A": C = ChrW(194)
                    Case ChrW(225): C = ChrW(7845) 'a'
                    Case ChrW(193): C = ChrW(7844) 'A'
                    Case ChrW(224): C = ChrW(7847) 'a`
                    Case ChrW(192): C = ChrW(7846) 'A`
                    Case ChrW(7843): C = ChrW(7849) 'a?
                    Case ChrW(7842): C = ChrW(7848) 'A?
                    Case ChrW(227): C = ChrW(7851) 'a~
                    Case ChrW(195): C = ChrW(7850) 'A~
                    Case ChrW(7841): C = ChrW(7853) 'a.
                    Case ChrW(7840): C = ChrW(7852) 'A.
                    
                    Case "e": C = ChrW(234)
                    Case "E": C = ChrW(202)
                    Case ChrW(233): C = ChrW(7871) 'e'
                    Case ChrW(201): C = ChrW(7870) 'E'
                    Case ChrW(232): C = ChrW(7873) 'e`
                    Case ChrW(200): C = ChrW(7872) 'E`
                    Case ChrW(7867): C = ChrW(7875) 'e?
                    Case ChrW(7866): C = ChrW(7874) 'E?
                    Case ChrW(7869): C = ChrW(7877) 'e~
                    Case ChrW(7868): C = ChrW(7876) 'E~
                    Case ChrW(7865): C = ChrW(7879) 'e.
                    Case ChrW(7864): C = ChrW(7878) 'E.
                    
                    Case "o": C = ChrW(244)
                    Case "O": C = ChrW(212)
                    Case ChrW(243): C = ChrW(7889) 'o'
                    Case ChrW(211): C = ChrW(7888) 'O'
                    Case ChrW(242): C = ChrW(7891) 'o`
                    Case ChrW(210): C = ChrW(7890) 'O`
                    Case ChrW(7887): C = ChrW(7893) 'o?
                    Case ChrW(7886): C = ChrW(7892) 'O?
                    Case ChrW(245): C = ChrW(7895) 'o~
                    Case ChrW(213): C = ChrW(7894) 'O~
                    Case ChrW(7885): C = ChrW(7897) 'o.
                    Case ChrW(7884): C = ChrW(7896) 'O.
                    
                    Case Else
                        j = j + 1
                        C = "6"
                End Select
            Case 7:
                Select Case Mid(kq, j + 1, 1)
                    Case "o": C = ChrW(417)
                    Case "O": C = ChrW(416)
                    Case ChrW(243): C = ChrW(7899) 'o'
                    Case ChrW(211): C = ChrW(7898) 'O'
                    Case ChrW(242): C = ChrW(7901) 'o`
                    Case ChrW(210): C = ChrW(7900) 'O`
                    Case ChrW(7887): C = ChrW(7903) 'o?
                    Case ChrW(7886): C = ChrW(7902) 'O?
                    Case ChrW(245): C = ChrW(7905) 'o~
                    Case ChrW(213): C = ChrW(7904) 'O~
                    Case ChrW(7885): C = ChrW(7907) 'o.
                    Case ChrW(7884): C = ChrW(7906) 'O.
                    
                    Case "u": C = ChrW(432)
                    Case "U": C = ChrW(431)
                    Case ChrW(250): C = ChrW(7913) 'u'
                    Case ChrW(218): C = ChrW(7912) 'U'
                    Case ChrW(249): C = ChrW(7915) 'u`
                    Case ChrW(217): C = ChrW(7914) 'U`
                    Case ChrW(7911): C = ChrW(7917) 'u?
                    Case ChrW(7910): C = ChrW(7916) 'U?
                    Case ChrW(361): C = ChrW(7919) 'u~
                    Case ChrW(360): C = ChrW(7918) 'U~
                    Case ChrW(7909): C = ChrW(7921) 'u.
                    Case ChrW(7908): C = ChrW(7920) 'U.
                    
                    Case Else
                        j = j + 1
                        C = "7"
                End Select
            Case 8:
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(259)
                    Case "A": C = ChrW(258)
                    Case ChrW(225): C = ChrW(7855) 'a'
                    Case ChrW(193): C = ChrW(7854) 'A'
                    Case ChrW(224): C = ChrW(7857) 'a`
                    Case ChrW(192): C = ChrW(7856) 'A`
                    Case ChrW(7843): C = ChrW(7859) 'a?
                    Case ChrW(7842): C = ChrW(7858) 'A?
                    Case ChrW(227): C = ChrW(7861) 'a~
                    Case ChrW(195): C = ChrW(7860) 'A~
                    Case ChrW(7841): C = ChrW(7863) 'a.
                    Case ChrW(7840): C = ChrW(7862) 'A.
                    
                    Case Else
                        j = j + 1
                        C = "8"
                End Select
            Case 9:
                Select Case Mid(kq, j + 1, 1)
                    Case "d": C = ChrW(273)
                    Case "D": C = ChrW(272)
                    Case Else
                        j = j + 1
                        C = "9"
                End Select
            Case Else
                j = j + 1
                C = Mid(s, i, 1)
        End Select
        kq = Mid(kq, 1, j) & C
    Next i
    VNI_Unicode = Replace(kq, cSkip, "")
End Function

'cSkip la` ky' tu ko duoc hien thi, dung de ngan cach khong bo dau. Mac dinh: "¦" (Alt+179)
'Vi du:
'    +s=" AS " --> Telex_Unicode=" A' "
'    +s=" A¦S " --> Telex_Unicode=" AS "
Function Telex_Unicode(s As String, Optional cSkip = "¦") As String
Dim i&, j&, kq$, C$
    If Len(s) = 0 Then Telex_Unicode = "": Exit Function
    kq = Left(s, 1)
    For i = 2 To Len(s)
        j = Len(kq) - 1
        Select Case Mid(s, i, 1)
            Case "s", "S":
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(225)
                    Case "A": C = ChrW(193)
                    Case ChrW(259): C = ChrW(7855) 'a(
                    Case ChrW(258): C = ChrW(7854) 'A(
                    Case ChrW(226): C = ChrW(7845) 'a^
                    Case ChrW(194): C = ChrW(7844) 'A^
                    
                    Case "e": C = ChrW(233)
                    Case "E": C = ChrW(201)
                    Case ChrW(234): C = ChrW(7871) 'e^
                    Case ChrW(202): C = ChrW(7870) 'E^
                
                    Case "i": C = ChrW(237)
                    Case "I": C = ChrW(205)
                    
                    Case "o": C = ChrW(243)
                    Case "O": C = ChrW(211)
                    Case ChrW(244): C = ChrW(7889) 'o^
                    Case ChrW(212): C = ChrW(7888) 'O^
                    Case ChrW(417): C = ChrW(7899) 'o*
                    Case ChrW(416): C = ChrW(7898) 'O*
                    
                    Case "u": C = ChrW(250)
                    Case "U": C = ChrW(218)
                    Case ChrW(432): C = ChrW(7913) 'u*
                    Case ChrW(431): C = ChrW(7912) 'U*
                    
                    Case "y": C = ChrW(253)
                    Case "Y": C = ChrW(221)
                    
                    Case Else
                        j = j + 1
                        C = Mid(s, i, 1)
                End Select
            Case "f", "F":
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(224)
                    Case "A": C = ChrW(192)
                    Case ChrW(259): C = ChrW(7857) 'a(
                    Case ChrW(258): C = ChrW(7856) 'A(
                    Case ChrW(226): C = ChrW(7847) 'a^
                    Case ChrW(194): C = ChrW(7846) 'A^
                    
                    Case "e": C = ChrW(232)
                    Case "E": C = ChrW(200)
                    Case ChrW(234): C = ChrW(7873) 'e^
                    Case ChrW(202): C = ChrW(7872) 'E^
                
                    Case "i": C = ChrW(236)
                    Case "I": C = ChrW(204)
                    
                    Case "o": C = ChrW(242)
                    Case "O": C = ChrW(210)
                    Case ChrW(244): C = ChrW(7891) 'o^
                    Case ChrW(212): C = ChrW(7890) 'O^
                    Case ChrW(417): C = ChrW(7901) 'o*
                    Case ChrW(416): C = ChrW(7900) 'O*
                    
                    Case "u": C = ChrW(249)
                    Case "U": C = ChrW(217)
                    Case ChrW(432): C = ChrW(7915) 'u*
                    Case ChrW(431): C = ChrW(7914) 'U*
                    
                    Case "y": C = ChrW(7923)
                    Case "Y": C = ChrW(7922)
                    
                    Case Else
                        j = j + 1
                        C = Mid(s, i, 1)
                End Select
            Case "r", "R":
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(7843)
                    Case "A": C = ChrW(7842)
                    Case ChrW(259): C = ChrW(7859) 'a(
                    Case ChrW(258): C = ChrW(7858) 'A(
                    Case ChrW(226): C = ChrW(7849) 'a^
                    Case ChrW(194): C = ChrW(7848) 'A^
                    
                    Case "e": C = ChrW(7867)
                    Case "E": C = ChrW(7866)
                    Case ChrW(234): C = ChrW(7875) 'e^
                    Case ChrW(202): C = ChrW(7874) 'E^
                
                    Case "i": C = ChrW(7881)
                    Case "I": C = ChrW(7880)
                    
                    Case "o": C = ChrW(7887)
                    Case "O": C = ChrW(7886)
                    Case ChrW(244): C = ChrW(7893) 'o^
                    Case ChrW(212): C = ChrW(7892) 'O^
                    Case ChrW(417): C = ChrW(7903) 'o*
                    Case ChrW(416): C = ChrW(7902) 'O*
                    
                    Case "u": C = ChrW(7911)
                    Case "U": C = ChrW(7910)
                    Case ChrW(432): C = ChrW(7917) 'u*
                    Case ChrW(431): C = ChrW(7916) 'U*
                    
                    Case "y": C = ChrW(7927)
                    Case "Y": C = ChrW(7926)
                    
                    Case Else
                        j = j + 1
                        C = Mid(s, i, 1)
                End Select
            Case "x", "X":
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(227)
                    Case "A": C = ChrW(195)
                    Case ChrW(259): C = ChrW(7861) 'a(
                    Case ChrW(258): C = ChrW(7860) 'A(
                    Case ChrW(226): C = ChrW(7851) 'a^
                    Case ChrW(194): C = ChrW(7850) 'A^
                    
                    Case "e": C = ChrW(7869)
                    Case "E": C = ChrW(7868)
                    Case ChrW(234): C = ChrW(7877) 'e^
                    Case ChrW(202): C = ChrW(7876) 'E^
                
                    Case "i": C = ChrW(297)
                    Case "I": C = ChrW(296)
                    
                    Case "o": C = ChrW(245)
                    Case "O": C = ChrW(213)
                    Case ChrW(244): C = ChrW(7895) 'o^
                    Case ChrW(212): C = ChrW(7894) 'O^
                    Case ChrW(417): C = ChrW(7905) 'o*
                    Case ChrW(416): C = ChrW(7904) 'O*
                    
                    Case "u": C = ChrW(361)
                    Case "U": C = ChrW(360)
                    Case ChrW(432): C = ChrW(7919) 'u*
                    Case ChrW(431): C = ChrW(7918) 'U*
                    
                    Case "y": C = ChrW(7929)
                    Case "Y": C = ChrW(7928)
                    
                    Case Else
                        j = j + 1
                        C = Mid(s, i, 1)
                End Select
            Case "j", "J":
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(7841)
                    Case "A": C = ChrW(7840)
                    Case ChrW(259): C = ChrW(7863) 'a(
                    Case ChrW(258): C = ChrW(7862) 'A(
                    Case ChrW(226): C = ChrW(7853) 'a^
                    Case ChrW(194): C = ChrW(7852) 'A^
                    
                    Case "e": C = ChrW(7865)
                    Case "E": C = ChrW(7864)
                    Case ChrW(234): C = ChrW(7879) 'e^
                    Case ChrW(202): C = ChrW(7878) 'E^
                
                    Case "i": C = ChrW(7883)
                    Case "I": C = ChrW(7882)
                    
                    Case "o": C = ChrW(7885)
                    Case "O": C = ChrW(7884)
                    Case ChrW(244): C = ChrW(7897) 'o^
                    Case ChrW(212): C = ChrW(7896) 'O^
                    Case ChrW(417): C = ChrW(7907) 'o*
                    Case ChrW(416): C = ChrW(7906) 'O*
                    
                    Case "u": C = ChrW(7909)
                    Case "U": C = ChrW(7908)
                    Case ChrW(432): C = ChrW(7921) 'u*
                    Case ChrW(431): C = ChrW(7920) 'U*
                    
                    Case "y": C = ChrW(7925)
                    Case "Y": C = ChrW(7924)
                    
                    Case Else
                        j = j + 1
                        C = Mid(s, i, 1)
                End Select
            Case "a", "A"
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(226)
                    Case "A": C = ChrW(194)
                    Case Else
                        j = j + 1
                        C = Mid(s, i, 1)
                End Select
            Case "e", "E"
                Select Case Mid(kq, j + 1, 1)
                    Case "e": C = ChrW(234)
                    Case "E": C = ChrW(202)
                    Case Else
                        j = j + 1
                        C = Mid(s, i, 1)
                End Select
            Case "o", "O"
                Select Case Mid(kq, j + 1, 1)
                    Case "o": C = ChrW(244)
                    Case "O": C = ChrW(212)
                    Case Else
                        j = j + 1
                        C = Mid(s, i, 1)
                End Select
            Case "w", "W"
                Select Case Mid(kq, j + 1, 1)
                    Case "a": C = ChrW(259)
                    Case "A": C = ChrW(258)
                    Case "o": C = ChrW(417)
                    Case "O": C = ChrW(416)
                    Case "u": C = ChrW(432)
                    Case "U": C = ChrW(431)
                    Case Else
                        j = j + 1
                        C = Mid(s, i, 1)
                End Select
            Case "d", "D":
                Select Case Mid(kq, j + 1, 1)
                    Case "d": C = ChrW(273)
                    Case "D": C = ChrW(272)
                    Case Else
                        j = j + 1
                        C = Mid(s, i, 1)
                End Select
            Case Else
                j = j + 1
                C = Mid(s, i, 1)
        End Select
        kq = Mid(kq, 1, j) & C
    Next i
    Telex_Unicode = Replace(kq, cSkip, "")
End Function


