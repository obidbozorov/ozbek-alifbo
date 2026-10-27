Attribute VB_Name = "OzbekAlifbo"
'==========================================================================
'  YANGI O'ZBEK ALIFBOSI - PowerPoint uchun makros (avtomatik yig'ilgan: tools/build_apps.py)
'  O' -> O-umlaut, G' -> G-breve, Sh -> S-sedil, Ch -> C-sedil
'  (Ng o'zgarmaydi: yangi alifboda n+g deb yoziladi)
'
'  Makroslar:
'    YangiAlifbo_Hujjat      - butun hujjatni o'giradi (kolontitul, izohlar bilan)
'    YangiAlifbo_Tanlangan   - faqat belgilangan matnni o'giradi
'    YangiAlifbo_Kirill      - kirill matnini yangi lotinga o'giradi (butun hujjat)
'    YangiAlifbo_KirillTanlangan - kirill -> yangi lotin, faqat belgilangan matn
'    YangiAlifbo_Yozishda    - yozish paytida avtomatik o'girishni yoqadi/o'chiradi
'    YangiAlifbo_Istisnolar  - istisnolarni boshqarish: ko'rish, qo'shish,
'                              o'chirish, standartlarni yoqish/o'chirish, tekshirish
'    YangiAlifbo_IstisnoQoshish - kursor turgan so'zni istisnoga qo'shish
'    YangiAlifbo_KontekstMenyu  - o'ng tugma menyusiga "Istisnolarga qo'shish" qo'shish
'    YangiAlifbo_LotinKirill / _LotinKirillTanlangan - lotin (yangi yoki amaldagi) -> kirill
'    YangiAlifbo_EskiLotin / _EskiLotinTanlangan - yangi lotin -> eski (amaldagi) lotin
'    YangiAlifbo_Tugmalar    - tezkor tugmalarni (Alt + A/E/K/C/Z/I/O) o'rnatish/olib tashlash
'    YangiAlifbo_KirillIstisnolar     - kirilldan o'girish istisnolari (Maykrosoft=Microsoft ...)
'    YangiAlifbo_KirillIstisnoQoshish - belgilangan kirill so'zini istisnoga qo'shish
'
'  Eslatma: bu fayl ataylab faqat ASCII belgilarda yozilgan
'  (VBA muharriri UTF-8 ni to'g'ri o'qimaydi). Maxsus harflar ChrW() orqali.
'==========================================================================
Option Explicit
Option Compare Binary

' ------------------------- SOZLAMALAR -------------------------------------
' ma'no -> ma<tutuq>no (U+02BC) qilib tutuq belgisini ham bir xillashtirish
Private Const TUTUQNI_TOGRILASH As Boolean = False
' Inglizcha abzaslarni (annotatsiya, adabiyotlar) o'tkazib yuborish
Private Const INGLIZCHA_ABZASNI_OTKAZISH As Boolean = True
' Kirilldan o'girishda ruscha abzaslarni o'tkazib yuborish
Private Const RUSCHA_ABZASNI_OTKAZISH As Boolean = True
' Standart istisnolar. Katta harf bo'lsa - registrga sezgir.
' "*" = shu bilan boshlanadigan so'zlar, "#" = keyin raqam / - / kelishi mumkin
Private Const STANDART_ISTISNOLAR As String = _
    "ishoq*,mushaf*,ashob*,ishol*,SHA#,SSH#,SSHv2,OpenSSH,SHAKE#,SHACAL#,ChaCha#," & _
    "CHACHA#,XChaCha#,hash,hashes,hashing,hashed,hashcash,hashdump,HashiCorp," & _
    "shellcode,Shellshock,phishing,smishing,vishing,spearphishing,Schnorr,Shamir," & _
    "Shannon,Schneier,Chaum,Shoup,Shacham,Bleichenbacher,Threefish,Shodan,Shamoon," & _
    "SharpHound,Dashlane,Chainlink,Chainalysis,bash,zsh,shell,powershell,push,flash," & _
    "fetch,match,search,research,branch,approach,chain,change,channel,machine," & _
    "dashboard,shift,mesh,refresh,publish,flush,splash,hashtag,threshold,Scheme," & _
    "Apache,PyTorch,Torch,chrome,chromium,chromebook,chatgpt,sharepoint,photoshop," & _
    "shopify,Shotcut,Chevrolet,Chery,Changan,Bosch,Porsche,english,Shakespeare," & _
    "Michigan,Massachusetts,Munich,Zurich,Utrecht,O'Brien,O'Connor,O'Connell,O'Neill," & _
    "O'Neil,O'Reilly,O'Hara,O'Donnell,O'Sullivan,O'Keefe,O'Leary,O'Malley,O'Toole," & _
    "O'Higgins,o'clock,g'day"
Private Const MONO_SHRIFTLAR As String = _
    "|courier new|courier|consolas|lucida console|cascadia code|cascadia mono|" & _
    "source code pro|menlo|monaco|jetbrains mono|fira code|"
Private Const EN_SOZLAR As String = _
    "|the|and|of|are|for|with|this|that|from|which|was|were|have|has|been|these|their|into|its|"
' --------------------------------------------------------------------------

#If VBA7 Then
Private Declare PtrSafe Function GetKeyboardLayout Lib "user32" (ByVal idThread As Long) As LongPtr
Private Declare PtrSafe Function ToUnicodeEx Lib "user32" (ByVal wVirtKey As Long, ByVal wScanCode As Long, _
    ByRef lpKeyState As Byte, ByVal pwszBuff As LongPtr, ByVal cchBuff As Long, ByVal wFlags As Long, _
    ByVal dwhkl As LongPtr) As Long
Private Declare PtrSafe Function MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExW" (ByVal uCode As Long, _
    ByVal uMapType As Long, ByVal dwhkl As LongPtr) As Long
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal idThread As Long) As Long
Private Declare Function ToUnicodeEx Lib "user32" (ByVal wVirtKey As Long, ByVal wScanCode As Long, _
    ByRef lpKeyState As Byte, ByVal pwszBuff As Long, ByVal cchBuff As Long, ByVal wFlags As Long, _
    ByVal dwhkl As Long) As Long
Private Declare Function MapVirtualKeyEx Lib "user32" Alias "MapVirtualKeyExW" (ByVal uCode As Long, _
    ByVal uMapType As Long, ByVal dwhkl As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

Private Const VK_H As Long = 72
Private Const VK_APOS As Long = 222     ' wdKeySingleQuote
Private Const VK_BACKTICK As Long = 192 ' wdKeyBackSingleQuote

Private mApos As String
Private mTutuq As String
Private mExc() As String
Private mExcCount As Long
Private mInit As Boolean

Private mPos() As Long
Private mLen() As Long
Private mFrom() As String
Private mTo() As String
Private mCount As Long
Private mSkipped As Long
Private mMode As Long        ' 0 - lotin, 1 - kirill -> lotin, 2 - lotin -> kirill, 3 - yangi -> eski lotin
' kirill istisnolari
Private mCyrDef() As String
Private mCyrDefCount As Long
Private mCyrKey() As String
Private mCyrPrefix() As Boolean
Private mCyrLatin() As String
Private mCyrHasLatin() As Boolean
Private mCyrCount As Long
Private mCyrIndex As Collection
' yozish paytida o'girish holati
Private mLive As Boolean
Private mRibbon As Object      ' lenta (nomlarni yangilash uchun)
Private mLiveLoaded As Boolean
Private mTickOn As Boolean
' lotin -> kirill lug'ati
Private mLcIndex As Collection
Private mLcCyr() As String
Private mLcPrefix() As Boolean
Private mLcCount As Long

'==========================================================================
'  ASOSIY MAKROSLAR
'==========================================================================
Private Sub InitAll()
    If mInit Then Exit Sub
    mApos = "'" & ChrW(&H2018) & ChrW(&H2019) & ChrW(&H2BB) & ChrW(&H2BC) & "`" & ChrW(&HB4) & ChrW(&H2032) & _
        ChrW(&H2B9) & ChrW(&H2BD) & ChrW(&H2BE) & ChrW(&H2BF) & ChrW(&H2C8) & ChrW(&H2CA) & ChrW(&H559) & _
        ChrW(&H55A) & ChrW(&H5F3) & ChrW(&H1FFD) & ChrW(&H1FFE) & ChrW(&H201B) & ChrW(&HA78C)
    mTutuq = ChrW(&H2BC)
    LoadExceptions
    LoadCyrExceptions
    LoadLatCyr
    mInit = True
End Sub

' Matndagi o'zgarishlarni hisoblaydi; natija mPos/mLen/mFrom/mTo (1..mCount)
Private Function ComputeEdits(ByVal t As String) As Long
    Dim n As Long, i As Long, st As Long, quoteOpen As Boolean
    If mMode = 1 Then ComputeEdits = ComputeEditsCyr(t): Exit Function
    If mMode = 2 Then ComputeEdits = ComputeEditsLatCyr(t): Exit Function
    If mMode = 3 Then ComputeEdits = ComputeEditsNewToOld(t): Exit Function
    mCount = 0
    ComputeEdits = 0
    If INGLIZCHA_ABZASNI_OTKAZISH Then
        If IsEnglishText(t) Then Exit Function
    End If
    n = Len(t)
    i = 1
    Do While i <= n
        If IsSpaceCh(Mid$(t, i, 1)) Then
            i = i + 1
        Else
            st = i
            Do While i <= n
                If IsSpaceCh(Mid$(t, i, 1)) Then Exit Do
                i = i + 1
            Loop
            quoteOpen = ProcessToken(Mid$(t, st, i - st), st - 1, quoteOpen)
        End If
    Loop
    ComputeEdits = mCount
End Function

Private Function ProcessToken(ByVal orig As String, ByVal basePos As Long, ByVal quoteOpen As Boolean) As Boolean
    Dim ln As Long, lead As Long, lim As Long, t As Long, aposRun As Long, i As Long
    Dim c As String, nx As String, nx2 As String, rep As String, core As String
    Dim tok As String, fixed As Boolean, c0 As Long, h As Long, q As Long, covered As Boolean
    tok = FixHomoglyphs(orig)
    fixed = (tok <> orig)
    c0 = mCount

    ln = Len(tok)
    ' boshidagi tinish belgilari; apostrof - qo'shtirnoq ochilishi
    lead = 1
    Do While lead <= ln
        c = Mid$(tok, lead, 1)
        If IsLetterCh(c) Or IsDigitCh(c) Then Exit Do
        If IsApos(c) Then quoteOpen = True
        lead = lead + 1
    Loop
    If lead > ln Then ProcessToken = quoteOpen: Exit Function

    ' oxiridagi tinish belgilari (apostrofdan tashqari)
    lim = ln
    t = ln
    Do While t > lead
        c = Mid$(tok, t, 1)
        If IsLetterCh(c) Or IsApos(c) Or IsDigitCh(c) Then Exit Do
        t = t - 1
    Loop
    If IsApos(Mid$(tok, t, 1)) Then
        aposRun = 0
        Do While t - aposRun > lead
            If Not IsApos(Mid$(tok, t - aposRun, 1)) Then Exit Do
            aposRun = aposRun + 1
        Loop
        If aposRun >= 2 And Not quoteOpen Then
            lim = t - aposRun - 1          ' boo''  -> ataylab yozilgan
        ElseIf quoteOpen Then
            lim = t - 1                    ' yopuvchi qo'shtirnoq
            quoteOpen = False
        End If
    End If

    core = CoreOf(tok)
    i = lead
    If LooksLikeLinkOrCode(tok) Then
        i = lim + 1
    ElseIf IsExceptionWord(core) Then
        i = lim + 1
    ElseIf LooksForeign(core) Then
        i = lim + 1
    End If
    Do While i <= lim
        c = Mid$(tok, i, 1)
        nx = CharAt(tok, i + 1, lim)
        If (c = "o" Or c = "O" Or c = "g" Or c = "G") And IsApos(nx) Then
            nx2 = CharAt(tok, i + 2, lim)
            If IsApos(nx2) Then
                i = i + 3
            Else
                Select Case c
                    Case "o": rep = ChrW(&HF6)
                    Case "O": rep = ChrW(&HD6)
                    Case "g": rep = ChrW(&H11F)
                    Case Else: rep = ChrW(&H11E)
                End Select
                AddEdit basePos + i - 1, 2, Mid$(orig, i, 2), rep
                i = i + 2
            End If
        ElseIf (c = "s" Or c = "S" Or c = "c" Or c = "C") And (nx = "h" Or nx = "H") Then
            Select Case c
                Case "s": rep = ChrW(&H15F)
                Case "S": rep = ChrW(&H15E)
                Case "c": rep = ChrW(&HE7)
                Case Else: rep = ChrW(&HC7)
            End Select
            AddEdit basePos + i - 1, 2, Mid$(orig, i, 2), rep
            i = i + 2
        ElseIf TutuqCase(tok, i, lead, c, nx) Then
            AddEdit basePos + i - 1, 1, c, mTutuq
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
    If fixed Then
        ' o'girilmagan joylardagi o'xshash kirill harflari - bittadan almashtiriladi
        For h = 1 To ln
            If Mid$(orig, h, 1) <> Mid$(tok, h, 1) Then
                covered = False
                For q = c0 + 1 To mCount
                    If basePos + h - 1 >= mPos(q) And basePos + h - 1 < mPos(q) + mLen(q) Then covered = True: Exit For
                Next q
                If Not covered Then AddEdit basePos + h - 1, 1, Mid$(orig, h, 1), Mid$(tok, h, 1)
            End If
        Next h
        SortEdits c0 + 1, mCount
    End If
    ProcessToken = quoteOpen
End Function

' Lotin so'z ichiga adashib kirgan kirill harflari (ko'rinishi bir xil): Toshkent (kirill o) -> Toshkent.
' So'zda lotin harfi bo'lsa va boshqa kirill harfi bo'lmasa - tuzatilgan shakl, aks holda o'zi.
Private Function FixHomoglyphs(ByVal tok As String) As String
    Dim i As Long, c As String, hasLat As Boolean, hasHomo As Boolean, out As String, m As String, code As Long
    FixHomoglyphs = tok
    For i = 1 To Len(tok)
        c = Mid$(tok, i, 1)
        code = AscW(c)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Then
            hasLat = True
        ElseIf Len(HomoOf(c)) > 0 Then
            hasHomo = True
        ElseIf code >= &H400 And code <= &H4FF Then
            Exit Function
        End If
    Next i
    If Not (hasLat And hasHomo) Then Exit Function
    For i = 1 To Len(tok)
        c = Mid$(tok, i, 1)
        m = HomoOf(c)
        If Len(m) > 0 Then out = out & m Else out = out & c
    Next i
    FixHomoglyphs = out
End Function

Private Function HomoOf(ByVal c As String) As String
    Select Case AscW(c)
        Case &H430: HomoOf = "a"
        Case &H435: HomoOf = "e"
        Case &H43E: HomoOf = "o"
        Case &H440: HomoOf = "p"
        Case &H441: HomoOf = "c"
        Case &H445: HomoOf = "x"
        Case &H443: HomoOf = "y"
        Case &H43A: HomoOf = "k"
        Case &H410: HomoOf = "A"
        Case &H412: HomoOf = "B"
        Case &H415: HomoOf = "E"
        Case &H41A: HomoOf = "K"
        Case &H41C: HomoOf = "M"
        Case &H41D: HomoOf = "H"
        Case &H41E: HomoOf = "O"
        Case &H420: HomoOf = "P"
        Case &H421: HomoOf = "C"
        Case &H422: HomoOf = "T"
        Case &H425: HomoOf = "X"
        Case Else: HomoOf = ""
    End Select
End Function

' mPos bo'yicha o'sish tartibida saralash (a..b oralig'ida; kichik ro'yxatlar uchun)
Private Sub SortEdits(ByVal a As Long, ByVal b As Long)
    Dim i As Long, j As Long, p As Long, l As Long, f As String, t As String
    For i = a + 1 To b
        p = mPos(i): l = mLen(i): f = mFrom(i): t = mTo(i)
        j = i - 1
        Do While j >= a
            If mPos(j) <= p Then Exit Do
            mPos(j + 1) = mPos(j): mLen(j + 1) = mLen(j): mFrom(j + 1) = mFrom(j): mTo(j + 1) = mTo(j)
            j = j - 1
        Loop
        mPos(j + 1) = p: mLen(j + 1) = l: mFrom(j + 1) = f: mTo(j + 1) = t
    Next i
End Sub

Private Function TutuqCase(ByVal tok As String, ByVal i As Long, ByVal lead As Long, _
                           ByVal c As String, ByVal nx As String) As Boolean
    TutuqCase = False
    If Not TUTUQNI_TOGRILASH Then Exit Function
    If Not IsApos(c) Then Exit Function
    If c = mTutuq Then Exit Function
    If i <= lead Then Exit Function
    If Not IsLetterCh(Mid$(tok, i - 1, 1)) Then Exit Function
    If Not IsLetterCh(nx) Then Exit Function
    TutuqCase = True
End Function

Private Sub AddEdit(ByVal pos As Long, ByVal ln As Long, ByVal fromS As String, ByVal toS As String)
    mCount = mCount + 1
    If mCount = 1 Then
        ReDim mPos(1 To 64): ReDim mLen(1 To 64): ReDim mFrom(1 To 64): ReDim mTo(1 To 64)
    ElseIf mCount > UBound(mPos) Then
        ReDim Preserve mPos(1 To mCount * 2): ReDim Preserve mLen(1 To mCount * 2)
        ReDim Preserve mFrom(1 To mCount * 2): ReDim Preserve mTo(1 To mCount * 2)
    End If
    mPos(mCount) = pos: mLen(mCount) = ln: mFrom(mCount) = fromS: mTo(mCount) = toS
End Sub

Private Function CharAt(ByVal s As String, ByVal idx As Long, ByVal lim As Long) As String
    If idx < 1 Or idx > lim Or idx > Len(s) Then CharAt = "" Else CharAt = Mid$(s, idx, 1)
End Function

Private Function IsApos(ByVal c As String) As Boolean
    If Len(c) <> 1 Then IsApos = False Else IsApos = (InStr(1, mApos, c, vbBinaryCompare) > 0)
End Function

Private Function IsLetterCh(ByVal c As String) As Boolean
    If Len(c) = 0 Then IsLetterCh = False Else IsLetterCh = (LCase$(c) <> UCase$(c))
End Function

Private Function IsDigitCh(ByVal c As String) As Boolean
    If Len(c) = 0 Then IsDigitCh = False Else IsDigitCh = (c >= "0" And c <= "9")
End Function

Private Function IsSpaceCh(ByVal c As String) As Boolean
    Dim a As Long
    a = AscW(c) And &HFFFF&
    IsSpaceCh = (a <= 32 Or a = 160 Or a = &H2007 Or a = &H202F)
End Function

Private Function CoreOf(ByVal tok As String) As String
    Dim s As Long, e As Long
    s = 1: e = Len(tok)
    Do While s <= e
        If IsLetterCh(Mid$(tok, s, 1)) Or IsDigitCh(Mid$(tok, s, 1)) Then Exit Do
        s = s + 1
    Loop
    Do While e >= s
        If IsLetterCh(Mid$(tok, e, 1)) Or IsDigitCh(Mid$(tok, e, 1)) Then Exit Do
        e = e - 1
    Loop
    If e >= s Then CoreOf = Mid$(tok, s, e - s + 1) Else CoreOf = ""
End Function

Private Function LooksForeign(ByVal core As String) As Boolean
    Dim i As Long, c As String, nx As String
    For i = 1 To Len(core)
        c = Mid$(core, i, 1)
        If c = "w" Or c = "W" Then LooksForeign = True: Exit Function
        If c = "c" Or c = "C" Then
            nx = Mid$(core, i + 1, 1)
            If nx <> "h" And nx <> "H" Then LooksForeign = True: Exit Function
        End If
    Next i
    LooksForeign = False
End Function

Private Function LooksLikeLinkOrCode(ByVal tok As String) As Boolean
    Dim lt As String, d As Long, ext As String, i As Long, c As String
    lt = LCase$(tok)
    LooksLikeLinkOrCode = True
    If InStr(lt, "://") > 0 Then Exit Function
    If Left$(lt, 4) = "www." Then Exit Function
    If InStr(lt, "@") > 0 Then Exit Function
    If Left$(lt, 1) = "#" Then Exit Function
    If InStr(lt, "\") > 0 Then Exit Function
    d = InStrRev(lt, ".")
    If d > 1 And d < Len(lt) Then
        ext = Mid$(lt, d + 1)
        If InStr("|com|uz|org|net|ru|io|edu|gov|doc|docx|pdf|xls|xlsx|ppt|pptx|exe|py|js|htm|html|txt|zip|", "|" & ext & "|") > 0 Then
            For i = 1 To Len(lt)
                c = Mid$(lt, i, 1)
                If Not ((c >= "a" And c <= "z") Or IsDigitCh(c) Or c = "." Or c = "-" Or c = "_") Then Exit For
            Next i
            If i > Len(lt) Then Exit Function
        End If
    End If
    LooksLikeLinkOrCode = False
End Function

Private Function IsEnglishText(ByVal t As String) As Boolean
    Dim i As Long, c As String, w As String, words As Long, hits As Long
    t = LCase$(t) & " "
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If IsLetterCh(c) Or c = "'" Then
            w = w & c
        ElseIf Len(w) > 0 Then
            words = words + 1
            If InStr(EN_SOZLAR, "|" & w & "|") > 0 Then hits = hits + 1
            w = ""
        End If
    Next i
    If words < 3 Then IsEnglishText = False: Exit Function
    IsEnglishText = (hits >= 2 And hits / words >= 0.08)
End Function

Private Function IsMonoFont(ByVal nm As String) As Boolean
    If Len(nm) = 0 Then IsMonoFont = False Else IsMonoFont = (InStr(MONO_SHRIFTLAR, "|" & LCase$(nm) & "|") > 0)
End Function

'---------------------------- Istisnolar ----------------------------------
Private Function IsExceptionWord(ByVal core As String) As Boolean
    IsExceptionWord = (Len(FindException(core)) > 0)
End Function

' So'zni himoya qilayotgan istisno yozuvi ("" - yo'q)
Private Function FindException(ByVal core As String) As String
    Dim parts() As String, i As Long, j As Long
    FindException = ""
    If Len(core) = 0 Then Exit Function
    parts = Split(Replace(core, "/", "-"), "-")
    For j = 0 To mExcCount - 1
        If MatchException(core, mExc(j)) Then FindException = mExc(j): Exit Function
        For i = 0 To UBound(parts)
            If MatchException(parts(i), mExc(j)) Then FindException = mExc(j): Exit Function
        Next i
    Next j
End Function

Private Function MatchException(ByVal w As String, ByVal entry As String) As Boolean
    Dim cs As Boolean, a As String, e As String, lastC As String, p As String, rest As String, i As Long
    MatchException = False
    If Len(entry) = 0 Or Len(w) = 0 Then Exit Function
    cs = (LCase$(entry) <> entry)
    If cs Then a = w: e = entry Else a = LCase$(w): e = LCase$(entry)
    lastC = Right$(e, 1)
    If lastC = "*" Then
        p = Left$(e, Len(e) - 1)
        MatchException = (Left$(a, Len(p)) = p)
    ElseIf lastC = "#" Then
        p = Left$(e, Len(e) - 1)
        If Left$(a, Len(p)) <> p Then Exit Function
        rest = Mid$(a, Len(p) + 1)
        For i = 1 To Len(rest)
            If InStr("0123456789-/", Mid$(rest, i, 1)) = 0 Then Exit Function
        Next i
        MatchException = True
    Else
        MatchException = (a = e)
    End If
End Function

' Standart ro'yxat ("!so'z" bilan o'chirilganlardan tashqari) + foydalanuvchi so'zlari
Private Sub LoadExceptions()
    Dim arr() As String, i As Long, users As Collection, offs As Collection, v As Variant
    arr = Split(STANDART_ISTISNOLAR, ",")
    ReadState users, offs
    mExcCount = 0
    ReDim mExc(0 To UBound(arr) + users.Count + 1)
    For i = 0 To UBound(arr)
        If Not InColl(offs, Trim$(arr(i)), True) Then
            mExc(mExcCount) = Trim$(arr(i))
            mExcCount = mExcCount + 1
        End If
    Next i
    For Each v In users
        mExc(mExcCount) = CStr(v)
        mExcCount = mExcCount + 1
    Next v
End Sub

' istisnolar.txt ni o'qiydi: oddiy qator - so'z, "!so'z" - o'chirilgan standart, ";" - izoh
Private Sub ReadState(users As Collection, offs As Collection, Optional ByVal path As String = "")
    Dim txt As String, lines() As String, i As Long, ln As String
    Set users = New Collection
    Set offs = New Collection
    If Len(path) = 0 Then path = ExceptionsPath()
    txt = ReadUtf8File(path)
    If Len(txt) = 0 Then Exit Sub
    If AscW(Left$(txt, 1)) = &HFEFF Then txt = Mid$(txt, 2)
    lines = Split(Replace(txt, vbCr, ""), vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) > 0 Then
            If Left$(ln, 1) = "!" Then
                ln = Trim$(Mid$(ln, 2))
                If Len(ln) > 0 Then
                    If Not InColl(offs, ln, True) Then offs.Add ln
                End If
            ElseIf Left$(ln, 1) <> ";" Then
                If InStr(ln, "=") > 0 Then
                    ' "o'zak=Lotin": bir o'zak uchun bitta yozuv (oxirgisi qoladi)
                    RemoveKey users, ExcKey(ln)
                    users.Add ln
                ElseIf Not InColl(users, ln, False) Then
                    users.Add ln
                End If
            End If
        End If
    Next i
End Sub

Private Sub SaveState(ByVal users As Collection, ByVal offs As Collection, Optional ByVal path As String = "")
    Dim t As String, v As Variant
    If Len(path) = 0 Then path = ExceptionsPath()
    If path = CyrExceptionsPath() Then
        t = "; Yangi alifbo - kirilldan o'girish istisnolari. Har qatorga bitta." & vbCrLf & _
            "; so'z = o'girilmaydi; so'z* = shu bilan boshlanganlar; o'zak=Lotin = maxsus shakl." & vbCrLf & _
            "; !o'zak = standart yozuvni o'chirish." & vbCrLf
    Else
        t = "; Yangi alifbo - istisno so'zlar. Har qatorga bitta." & vbCrLf & _
            "; Katta harf bor bo'lsa - registrga sezgir. * = shu bilan boshlanadi, # = keyin raqam." & vbCrLf & _
            "; !so'z = standart istisnoni o'chirish." & vbCrLf
    End If
    For Each v In offs
        t = t & "!" & v & vbCrLf
    Next v
    For Each v In users
        t = t & v & vbCrLf
    Next v
    WriteUtf8File path, t
    mInit = False   ' keyingi ishda qayta o'qiladi
End Sub

Private Function InColl(ByVal c As Collection, ByVal s As String, ByVal ignoreCase As Boolean) As Boolean
    Dim v As Variant
    InColl = False
    For Each v In c
        If ignoreCase Then
            If LCase$(CStr(v)) = LCase$(s) Then InColl = True: Exit Function
        Else
            If CStr(v) = s Then InColl = True: Exit Function
        End If
    Next v
End Function

Private Sub RemoveFromColl(ByVal c As Collection, ByVal s As String, ByVal ignoreCase As Boolean)
    Dim i As Long
    For i = c.Count To 1 Step -1
        If ignoreCase Then
            If LCase$(CStr(c(i))) = LCase$(s) Then c.Remove i
        Else
            If CStr(c(i)) = s Then c.Remove i
        End If
    Next i
End Sub

Private Function ExceptionsPath() As String
    ExceptionsPath = Environ$("APPDATA") & "\OzbekAlifbo\istisnolar.txt"
End Function

Private Function ReadUtf8File(ByVal path As String) As String
    Dim st As Object
    ReadUtf8File = ""
    If Len(Dir$(path)) = 0 Then Exit Function
    On Error GoTo Oxir
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    ReadUtf8File = st.ReadText(-1)
    st.Close
Oxir:
End Function

Private Sub WriteUtf8File(ByVal path As String, ByVal text As String)
    Dim st As Object, folder As String
    folder = Left$(path, InStrRev(path, "\") - 1)
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText text
    st.SaveToFile path, 2
    st.Close
End Sub

' Unicode harflarni ham ko'rsata oladigan xabar oynasi (MsgBox faqat ANSI)
Private Sub UMsg(ByVal text As String, Optional ByVal icon As Long = 64)
    On Error GoTo Oddiy
    CreateObject("WScript.Shell").Popup text, 0, AppNomi() & " - istisno so'zlar", icon
    Exit Sub
Oddiy:
    MsgBox text, vbInformation, AppNomi() & " - istisno so'zlar"
End Sub

' O'girilgan so'zni eski yozuvga qaytaradi: S-sedil+amir -> Shamir
Private Function RevertWord(ByVal w As String) As String
    Dim i As Long, c As String, out As String, letters As String, allCaps As Boolean
    For i = 1 To Len(w)
        c = Mid$(w, i, 1)
        If IsLetterCh(c) Then letters = letters & c
    Next i
    allCaps = (Len(letters) > 1 And letters = UCase$(letters))
    For i = 1 To Len(w)
        c = Mid$(w, i, 1)
        Select Case AscW(c)
            Case &H15F: out = out & "sh"
            Case &HE7: out = out & "ch"
            Case &HF6: out = out & "o" & ChrW(&H2BB)
            Case &H11F: out = out & "g" & ChrW(&H2BB)
            Case &H15E: out = out & IIf(allCaps, "SH", "Sh")
            Case &HC7: out = out & IIf(allCaps, "CH", "Ch")
            Case &HD6: out = out & "O" & ChrW(&H2BB)
            Case &H11E: out = out & "G" & ChrW(&H2BB)
            Case Else: out = out & c
        End Select
    Next i
    RevertWord = out
End Function

' So'z qo'shadi; natija haqida xabar matnini qaytaradi
Private Function AddUserWord(ByVal w As String) As String
    Dim users As Collection, offs As Collection, arr() As String, i As Long
    w = Trim$(w)
    If Len(w) = 0 Then AddUserWord = "So'z kiritilmadi.": Exit Function
    If InStr(w, " ") > 0 Then AddUserWord = "Faqat bitta so'z kiriting (bo'shliqsiz).": Exit Function
    ReadState users, offs
    arr = Split(STANDART_ISTISNOLAR, ",")
    For i = 0 To UBound(arr)
        If LCase$(Trim$(arr(i))) = LCase$(w) Then
            If InColl(offs, w, True) Then
                RemoveFromColl offs, w, True
                SaveState users, offs
                AddUserWord = "'" & Trim$(arr(i)) & "' standart ro'yxatda bor - qayta yoqildi."
            Else
                AddUserWord = "'" & Trim$(arr(i)) & "' standart ro'yxatda allaqachon bor."
            End If
            Exit Function
        End If
    Next i
    If InColl(users, w, False) Then AddUserWord = "'" & w & "' ro'yxatda allaqachon bor.": Exit Function
    users.Add w
    SaveState users, offs
    AddUserWord = "'" & w & "' istisnolarga qo'shildi."
End Function

' Bitta so'z nima bo'lishini tushuntiradi
Private Function ExplainWord(ByVal w As String) As String
    Dim core As String, ex As String, n As Long, k As Long, out As String
    InitAll
    w = Trim$(w)
    core = CoreOf(w)
    If LooksLikeLinkOrCode(w) Then ExplainWord = w & vbCrLf & "O'girilmaydi - havola yoki fayl nomi.": Exit Function
    ex = FindException(core)
    If Len(ex) > 0 Then ExplainWord = w & vbCrLf & "O'girilmaydi - istisno: " & ex: Exit Function
    If LooksForeign(core) Then ExplainWord = w & vbCrLf & "O'girilmaydi - chet so'z (w yoki yolg'iz c harfi bor).": Exit Function
    n = ComputeEdits(w)
    If n = 0 Then ExplainWord = w & vbCrLf & "O'zgarmaydi (o'giriladigan harf yo'q).": Exit Function
    out = w
    For k = n To 1 Step -1
        out = Left$(out, mPos(k)) & mTo(k) & Mid$(out, mPos(k) + mLen(k) + 1)
    Next k
    ExplainWord = w & "  ->  " & out
End Function

'==========================================================================
'  ISTISNOLARNI BOSHQARISH
'==========================================================================
Public Sub YangiAlifbo_Istisnolar()
    Dim choice As String, users As Collection, offs As Collection, nDef As Long
    Do
        ReadState users, offs
        nDef = UBound(Split(STANDART_ISTISNOLAR, ",")) + 1
        choice = InputBox( _
            "Mening so'zlarim: " & users.Count & " ta" & vbCrLf & _
            "Standart: " & (nDef - offs.Count) & " / " & nDef & " yoqilgan" & vbCrLf & vbCrLf & _
            "1 - Ro'yxatni ko'rish" & vbCrLf & _
            "2 - So'z qo'shish" & vbCrLf & _
            "3 - So'zni o'chirish" & vbCrLf & _
            "4 - Belgilangan so'zni qo'shish" & vbCrLf & _
            "5 - Standart istisnoni yoqish / o'chirish" & vbCrLf & _
            "6 - So'zni tekshirish" & vbCrLf & _
            "7 - Faylni Bloknotda tahrirlash" & vbCrLf & _
            "8 - Kirill istisnolari (kirilldan o'girish uchun)" & vbCrLf & vbCrLf & _
            "Raqamni kiriting:", AppNomi() & " - istisno so'zlar", "1")
        Select Case Trim$(choice)
            Case "": Exit Do
            Case "1": ShowExceptionList: Exit Do
            Case "2": PromptAdd
            Case "3": PromptRemove
            Case "4": YangiAlifbo_IstisnoQoshish
            Case "5": PromptToggleDefault
            Case "6": PromptTest
            Case "7": OpenInNotepad: Exit Do
            Case "8": YangiAlifbo_KirillIstisnolar: Exit Do
            Case Else: UMsg "1 dan 8 gacha raqam kiriting.", 48
        End Select
    Loop
End Sub

' Kursor turgan yoki belgilangan so'zni istisnoga qo'shadi.
' So'z allaqachon o'girilgan bo'lsa (S-sedil+amir), hujjatda Shamir ga qaytaradi.
Private Sub ShowExceptionList()
    Dim users As Collection, offs As Collection, arr() As String, i As Long, t As String, v As Variant
    ReadState users, offs
    arr = Split(STANDART_ISTISNOLAR, ",")
    t = AppNomi() & " - istisno so'zlar" & vbCr & vbCr
    t = t & "MENING SO'ZLARIM (" & users.Count & ")" & vbCr
    If users.Count = 0 Then t = t & "    (hali yo'q)" & vbCr
    For Each v In users
        t = t & "    " & v & vbCr
    Next v
    t = t & vbCr & "STANDART ISTISNO SO'ZLAR (" & (UBound(arr) + 1 - offs.Count) & " / " & (UBound(arr) + 1) & " yoqilgan)" & vbCr
    For i = 0 To UBound(arr)
        If InColl(offs, Trim$(arr(i)), True) Then
            t = t & "    " & Trim$(arr(i)) & "    [o'chirilgan]" & vbCr
        Else
            t = t & "    " & Trim$(arr(i)) & vbCr
        End If
    Next i
    t = t & vbCr & "Shablonlar: phish* - shu bilan boshlanadi; SHA# - keyin raqam yoki '-'." & vbCr & _
        "Katta harf bor bo'lsa - registrga sezgir." & vbCr & _
        "Fayl: " & ExceptionsPath() & vbCr & _
        "O'zgartirish: YangiAlifbo_Istisnolar makrosi yoki faylni tahrirlash."
    ShowText t
End Sub

Private Sub PromptAdd()
    Dim w As String
    w = InputBox("Yangi istisno so'z:" & vbCrLf & vbCrLf & _
                 "Shablonlar:" & vbCrLf & _
                 "  phish*  - shu bilan boshlanadigan so'zlar" & vbCrLf & _
                 "  SHA#    - keyin raqam yoki '-' (SHA-256)" & vbCrLf & _
                 "Katta harf bor bo'lsa - registrga sezgir.", AppNomi() & " - so'z qo'shish")
    If Len(Trim$(w)) > 0 Then UMsg AddUserWord(w)
End Sub

Private Sub PromptRemove()
    Dim users As Collection, offs As Collection, t As String, i As Long, ans As String, w As String
    ReadState users, offs
    If users.Count = 0 Then UMsg "Siz qo'shgan so'zlar yo'q.": Exit Sub
    For i = 1 To users.Count
        If Len(t) < 800 Then t = t & i & ") " & users(i) & vbCrLf
    Next i
    If Len(t) >= 800 Then t = t & "..." & vbCrLf
    ans = Trim$(InputBox(t & vbCrLf & "O'chiriladigan so'z raqami yoki o'zini kiriting:", AppNomi() & " - so'zni o'chirish"))
    If Len(ans) = 0 Then Exit Sub
    If IsNumeric(ans) Then
        If CLng(ans) >= 1 And CLng(ans) <= users.Count Then w = users(CLng(ans))
    ElseIf InColl(users, ans, True) Then
        For i = 1 To users.Count
            If LCase$(users(i)) = LCase$(ans) Then w = users(i): Exit For
        Next i
    End If
    If Len(w) = 0 Then UMsg "Bunday so'z topilmadi: " & ans, 48: Exit Sub
    RemoveFromColl users, w, False
    SaveState users, offs
    UMsg "'" & w & "' o'chirildi."
End Sub

Private Sub PromptToggleDefault()
    Dim users As Collection, offs As Collection, arr() As String, i As Long, t As String, ans As String, w As String
    ReadState users, offs
    arr = Split(STANDART_ISTISNOLAR, ",")
    For i = 0 To UBound(arr)
        t = t & (i + 1) & ")" & IIf(InColl(offs, Trim$(arr(i)), True), "[-]", "") & Trim$(arr(i)) & "  "
        If (i + 1) Mod 4 = 0 Then t = t & vbCrLf
    Next i
    ans = Trim$(InputBox(t & vbCrLf & vbCrLf & "[-] = o'chirilgan." & vbCrLf & _
                         "Yoqish/o'chirish uchun raqamni kiriting:", AppNomi() & " - standart istisno so'zlar"))
    If Len(ans) = 0 Then Exit Sub
    If Not IsNumeric(ans) Then UMsg "Raqam kiriting.", 48: Exit Sub
    i = CLng(ans) - 1
    If i < 0 Or i > UBound(arr) Then UMsg "1 dan " & (UBound(arr) + 1) & " gacha raqam kiriting.", 48: Exit Sub
    w = Trim$(arr(i))
    If InColl(offs, w, True) Then
        RemoveFromColl offs, w, True
        SaveState users, offs
        UMsg "'" & w & "' qayta yoqildi - endi o'girilmaydi."
    Else
        offs.Add w
        SaveState users, offs
        UMsg "'" & w & "' o'chirildi - endi bu so'z ham o'giriladi."
    End If
End Sub

Private Sub PromptTest()
    Dim w As String
    w = InputBox("Tekshiriladigan so'z:", AppNomi() & " - so'zni tekshirish")
    If Len(Trim$(w)) > 0 Then UMsg ExplainWord(w)
End Sub

Private Sub OpenInNotepad()
    Dim users As Collection, offs As Collection
    If Len(Dir$(ExceptionsPath())) = 0 Then
        ReadState users, offs
        SaveState users, offs
    End If
    Shell "notepad.exe """ & ExceptionsPath() & """", vbNormalFocus
    mInit = False
End Sub

'---------------------------- Kirill istisnolari --------------------------
' "so'z" - o'girilmaydi; "so'z*" - shu bilan boshlanganlar; "o'zak=Lotin" - maxsus shakl
' (qo'shimchalari odatdagidek o'giriladi: Maykrosoftning -> Microsoftning)

Private Function CyrExceptionsPath() As String
    CyrExceptionsPath = Environ$("APPDATA") & "\OzbekAlifbo\kirill-istisnolar.txt"
End Function

Private Sub AddCyrDefault(ByVal e As String)
    mCyrDefCount = mCyrDefCount + 1
    ReDim Preserve mCyrDef(1 To mCyrDefCount)
    mCyrDef(mCyrDefCount) = e
End Sub

' Standart kirill yozuvlari (ASCII fayl bo'lgani uchun ChrW orqali)
Private Sub InitCyrDefaults()
    mCyrDefCount = 0
    AddCyrDefault ChrW(&H41C) & ChrW(&H430) & ChrW(&H439) & ChrW(&H43A) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H43E) & ChrW(&H444) & ChrW(&H442) & "=Microsoft"
    AddCyrDefault ChrW(&H413) & ChrW(&H443) & ChrW(&H433) & ChrW(&H43B) & "=Google"
    AddCyrDefault ChrW(&H412) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H43E) & ChrW(&H443) & ChrW(&H441) & "=Windows"
    AddCyrDefault ChrW(&H424) & ChrW(&H435) & ChrW(&H439) & ChrW(&H441) & ChrW(&H431) & ChrW(&H443) & ChrW(&H43A) & "=Facebook"
    AddCyrDefault ChrW(&H42E) & ChrW(&H442) & ChrW(&H443) & ChrW(&H431) & "=YouTube"
    AddCyrDefault ChrW(&H42D) & ChrW(&H43F) & ChrW(&H43F) & ChrW(&H43B) & "=Apple"
    AddCyrDefault ChrW(&H425) & ChrW(&H443) & ChrW(&H430) & ChrW(&H432) & ChrW(&H435) & ChrW(&H439) & "=Huawei"
    AddCyrDefault ChrW(&H41B) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H443) & ChrW(&H43A) & ChrW(&H441) & "=Linux"
    AddCyrDefault ChrW(&H410) & ChrW(&H439) & ChrW(&H444) & ChrW(&H43E) & ChrW(&H43D) & "=iPhone"
    AddCyrDefault ChrW(&H41F) & ChrW(&H430) & ChrW(&H443) & ChrW(&H44D) & ChrW(&H440) & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H439) & ChrW(&H43D) & ChrW(&H442) & "=PowerPoint"
    AddCyrDefault ChrW(&H42D) & ChrW(&H43A) & ChrW(&H441) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & "=Excel"
    AddCyrDefault ChrW(&H412) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H434) & "=Word"
    AddCyrDefault ChrW(&H412) & ChrW(&H430) & ChrW(&H442) & ChrW(&H441) & ChrW(&H430) & ChrW(&H43F) & "=WhatsApp"
    AddCyrDefault ChrW(&H422) & ChrW(&H432) & ChrW(&H438) & ChrW(&H442) & ChrW(&H442) & ChrW(&H435) & ChrW(&H440) & "=Twitter"
    AddCyrDefault ChrW(&H421) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H439) & ChrW(&H43F) & "=Skype"
    AddCyrDefault ChrW(&H430) & ChrW(&H431) & ChrW(&H437) & ChrW(&H430) & ChrW(&H446) & "=abzas"
    AddCyrDefault ChrW(&H430) & ChrW(&H432) & ChrW(&H438) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H435) & ChrW(&H446) & "=avianoses"
    AddCyrDefault ChrW(&H431) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H435) & ChrW(&H446) & "=bronenoses"
    AddCyrDefault ChrW(&H43C) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H435) & ChrW(&H446) & "=minonoses"
    AddCyrDefault ChrW(&H43C) & ChrW(&H430) & ChrW(&H440) & ChrW(&H433) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H446) & "=marganes"
    AddCyrDefault ChrW(&H440) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H446) & "=ranes"
    AddCyrDefault ChrW(&H43A) & ChrW(&H43E) & ChrW(&H44D) & ChrW(&H444) & ChrW(&H444) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442) & "=koeffitsient"
    AddCyrDefault ChrW(&H446) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H438) & ChrW(&H434) & "=selluloid"
    AddCyrDefault ChrW(&H446) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H437) & ChrW(&H430) & "=selluloza"
    AddCyrDefault ChrW(&H43E) & ChrW(&H43A) & ChrW(&H442) & ChrW(&H44F) & ChrW(&H431) & ChrW(&H440) & "=oktabr"
    AddCyrDefault ChrW(&H441) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H44F) & ChrW(&H431) & ChrW(&H440) & "=sentabr"
    AddCyrDefault ChrW(&H43F) & ChrW(&H430) & ChrW(&H440) & ChrW(&H442) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440) & "=parter"
    AddCyrDefault ChrW(&H442) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H43D) & "=tulen"
End Sub

Private Function CyrLowerStr(ByVal s As String) As String
    Dim i As Long, out As String
    For i = 1 To Len(s)
        out = out & LCase$(CyrLower(Mid$(s, i, 1)))
    Next i
    CyrLowerStr = out
End Function

' Yozuv kaliti: "o'zak=Lotin" dan o'zak qismi, kichik harfda
Private Function ExcKey(ByVal e As String) As String
    Dim p As Long
    p = InStr(e, "=")
    If p > 0 Then e = Left$(e, p - 1)
    ExcKey = CyrLowerStr(Trim$(e))
End Function

Private Sub RemoveKey(ByVal c As Collection, ByVal key As String)
    Dim i As Long
    For i = c.Count To 1 Step -1
        If ExcKey(CStr(c(i))) = key Then c.Remove i
    Next i
End Sub

Private Function KeyInColl(ByVal c As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    KeyInColl = False
    For Each v In c
        If ExcKey(CStr(v)) = key Then KeyInColl = True: Exit Function
    Next v
End Function

' Standart + foydalanuvchi yozuvlaridan qoidalar jadvali (kalit bo'yicha qidiriladi)
Private Sub LoadCyrExceptions()
    Dim users As Collection, offs As Collection, v As Variant, i As Long
    InitCyrDefaults
    ReadState users, offs, CyrExceptionsPath()
    Set mCyrIndex = New Collection
    mCyrCount = 0
    ReDim mCyrKey(1 To mCyrDefCount + users.Count + 1)
    ReDim mCyrPrefix(1 To mCyrDefCount + users.Count + 1)
    ReDim mCyrLatin(1 To mCyrDefCount + users.Count + 1)
    ReDim mCyrHasLatin(1 To mCyrDefCount + users.Count + 1)
    For i = 1 To mCyrDefCount
        If Not KeyInColl(offs, ExcKey(mCyrDef(i))) Then AddCyrRule mCyrDef(i)
    Next i
    For Each v In users
        AddCyrRule CStr(v)          ' foydalanuvchi yozuvi standartni almashtiradi
    Next v
End Sub

Private Sub AddCyrRule(ByVal e As String)
    Dim p As Long, k As String
    e = Trim$(e)
    p = InStr(e, "=")
    mCyrCount = mCyrCount + 1
    mCyrHasLatin(mCyrCount) = False: mCyrPrefix(mCyrCount) = False: mCyrLatin(mCyrCount) = ""
    If p > 0 Then
        k = CyrLowerStr(Trim$(Left$(e, p - 1)))
        mCyrLatin(mCyrCount) = Trim$(Mid$(e, p + 1))
        mCyrHasLatin(mCyrCount) = True
        mCyrPrefix(mCyrCount) = True
    ElseIf Right$(e, 1) = "*" Then
        k = CyrLowerStr(Left$(e, Len(e) - 1))
        mCyrPrefix(mCyrCount) = True
    Else
        k = CyrLowerStr(e)
    End If
    If Len(k) = 0 Then mCyrCount = mCyrCount - 1: Exit Sub
    mCyrKey(mCyrCount) = k
    On Error Resume Next
    mCyrIndex.Remove k                ' shu kalit oldin bo'lsa - almashtiriladi
    On Error GoTo 0
    mCyrIndex.Add mCyrCount, k
End Sub

Private Function CyrIndexOf(ByVal k As String) As Long
    On Error GoTo Yoq
    CyrIndexOf = mCyrIndex.Item(k)
    Exit Function
Yoq:
    CyrIndexOf = 0
End Function

' Eng uzun mos kalit raqami (0 - yo'q)
Private Function FindCyrRule(ByVal run As String) As Long
    Dim rl As String, ln As Long, idx As Long
    FindCyrRule = 0
    If mCyrIndex Is Nothing Then Exit Function
    rl = CyrLowerStr(run)
    For ln = Len(rl) To 1 Step -1
        idx = CyrIndexOf(Left$(rl, ln))
        If idx > 0 Then
            If mCyrPrefix(idx) Or ln = Len(rl) Then FindCyrRule = idx: Exit Function
        End If
    Next ln
End Function

'==========================================================================
'  KIRILL -> YANGI LOTIN  (converter.js dagi cyrToLatin bilan bir xil)
'==========================================================================
Private Function CyrLower(ByVal c As String) As String
    Dim a As Long
    a = AscW(c) And &HFFFF&
    If a >= &H410 And a <= &H42F Then
        CyrLower = ChrW(a + &H20)
    ElseIf a >= &H400 And a <= &H40F Then
        CyrLower = ChrW(a + &H50)
    ElseIf a >= &H460 And a <= &H4FF And (a Mod 2 = 0) Then
        CyrLower = ChrW(a + 1)
    Else
        CyrLower = c
    End If
End Function

Private Function IsCyrCh(ByVal c As String) As Boolean
    Dim a As Long
    If Len(c) = 0 Then IsCyrCh = False: Exit Function
    a = AscW(c) And &HFFFF&
    IsCyrCh = (a >= &H400 And a <= &H4FF)
End Function

Private Function IsCyrVowel(ByVal lc As String) As Boolean
    If Len(lc) = 0 Then IsCyrVowel = False: Exit Function
    Select Case AscW(lc) And &HFFFF&
        Case &H430, &H435, &H451, &H438, &H43E, &H443, &H44D, &H44E, &H44F, &H45E, &H44B
            IsCyrVowel = True
        Case Else
            IsCyrVowel = False
    End Select
End Function

' Lotin natijani katta harfga: s-sedil, c-sedil, o-umlaut, g-breve ham
Private Function UpperLat(ByVal s As String) As String
    Dim i As Long, c As String, out As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        Select Case AscW(c) And &HFFFF&
            Case &H15F: out = out & ChrW(&H15E)
            Case &HE7: out = out & ChrW(&HC7)
            Case &HF6: out = out & ChrW(&HD6)
            Case &H11F: out = out & ChrW(&H11E)
            Case Else: out = out & UCase$(c)
        End Select
    Next i
    UpperLat = out
End Function

Private Function IsRussianText(ByVal t As String) As Boolean
    Dim i As Long, c As String, w As String, words As Long, hits As Long, ru As String
    IsRussianText = False
    ' o'zbekcha harflar bo'lsa - o'zbek matni
    For i = 1 To Len(t)
        Select Case AscW(Mid$(t, i, 1)) And &HFFFF&
            Case &H45E, &H49B, &H493, &H4B3, &H40E, &H49A, &H492, &H4B2
                Exit Function
        End Select
    Next i
    ru = "|" & ChrW(&H432) & "|" & ChrW(&H43D) & ChrW(&H430) & "|" & ChrW(&H43D) & ChrW(&H435) & "|" & ChrW(&H447) & ChrW(&H442) & ChrW(&H43E) & "|" & ChrW(&H434) & ChrW(&H43B) & ChrW(&H44F) & "|" & ChrW(&H44D)
    ru = ru & ChrW(&H442) & ChrW(&H43E) & "|" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43A) & "|" & ChrW(&H43F) & ChrW(&H440) & ChrW(&H438) & "|" & ChrW(&H438) & ChrW(&H43B) & ChrW(&H438) & "|" & ChrW(&H43F) & ChrW(&H43E)
    ru = ru & "|" & ChrW(&H436) & ChrW(&H435) & "|" & ChrW(&H431) & ChrW(&H44B) & ChrW(&H43B) & "|" & ChrW(&H431) & ChrW(&H44B) & ChrW(&H43B) & ChrW(&H430) & "|" & ChrW(&H431) & ChrW(&H44B) & ChrW(&H43B) & ChrW(&H438)
    ru = ru & "|" & ChrW(&H442) & ChrW(&H430) & ChrW(&H43A) & ChrW(&H436) & ChrW(&H435) & "|" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H44B) & ChrW(&H439) & "|" & ChrW(&H43A) & ChrW(&H43E)
    ru = ru & ChrW(&H442) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H44B) & ChrW(&H435) & "|" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H430) & ChrW(&H44F) & "|" & ChrW(&H435) & ChrW(&H433)
    ru = ru & ChrW(&H43E) & "|" & ChrW(&H441) & "|" & ChrW(&H438) & "|" & ChrW(&H43A) & "|"
    t = t & " "
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If IsCyrCh(c) Then
            w = w & CyrLower(c)
        ElseIf Len(w) > 0 Then
            words = words + 1
            If InStr(ru, "|" & w & "|") > 0 Then hits = hits + 1
            w = ""
        End If
    Next i
    If words < 3 Then Exit Function
    IsRussianText = (hits >= 2 And hits / words >= 0.08)
End Function

Private Function CyrChar(ByVal t As String, ByVal i As Long, ByVal runStart As Long) As String
    Dim lc As String, prevC As String, nextC As String
    lc = CyrLower(Mid$(t, i, 1))
    If i > runStart Then prevC = CyrLower(Mid$(t, i - 1, 1)) Else prevC = ""
    If i < Len(t) Then nextC = CyrLower(Mid$(t, i + 1, 1)) Else nextC = ""
    Select Case AscW(lc) And &HFFFF&
        Case &H435  ' ye / e
            If i = runStart Or IsCyrVowel(prevC) Or prevC = ChrW(&H44A) Or prevC = ChrW(&H44C) Then CyrChar = "ye" Else CyrChar = "e"
        Case &H446  ' ts / s
            If i = runStart Then
                CyrChar = "s"
            ElseIf IsCyrVowel(prevC) Then
                CyrChar = "ts"
            Else
                CyrChar = "s"
            End If
        Case &H44A  ' qattiq belgi
            If nextC = ChrW(&H435) Or nextC = ChrW(&H451) Or nextC = ChrW(&H44E) Or nextC = ChrW(&H44F) Then CyrChar = "" Else CyrChar = ChrW(&H2BC)
        Case &H44C  ' yumshoq belgi
            If nextC = ChrW(&H43E) Then CyrChar = "y" Else CyrChar = ""
        Case &H441  ' s (+ h -> s<tutuq>h)
            If nextC = ChrW(&H4B3) Then CyrChar = "s" & ChrW(&H2BC) Else CyrChar = "s"
        Case &H430: CyrChar = "a"
        Case &H431: CyrChar = "b"
        Case &H432: CyrChar = "v"
        Case &H433: CyrChar = "g"
        Case &H434: CyrChar = "d"
        Case &H436: CyrChar = "j"
        Case &H437: CyrChar = "z"
        Case &H438: CyrChar = "i"
        Case &H439: CyrChar = "y"
        Case &H43A: CyrChar = "k"
        Case &H43B: CyrChar = "l"
        Case &H43C: CyrChar = "m"
        Case &H43D: CyrChar = "n"
        Case &H43E: CyrChar = "o"
        Case &H43F: CyrChar = "p"
        Case &H440: CyrChar = "r"
        Case &H442: CyrChar = "t"
        Case &H443: CyrChar = "u"
        Case &H444: CyrChar = "f"
        Case &H445: CyrChar = "x"
        Case &H447: CyrChar = ChrW(&HE7)
        Case &H448, &H449: CyrChar = ChrW(&H15F)
        Case &H44B: CyrChar = "i"
        Case &H44D: CyrChar = "e"
        Case &H451: CyrChar = "yo"
        Case &H44E: CyrChar = "yu"
        Case &H44F: CyrChar = "ya"
        Case &H45E: CyrChar = ChrW(&HF6)
        Case &H49B: CyrChar = "q"
        Case &H493: CyrChar = ChrW(&H11F)
        Case &H4B3: CyrChar = "h"
        Case Else: CyrChar = Mid$(t, i, 1)
    End Select
End Function

' Har bir kirill so'zi - bitta o'zgarish (mPos/mLen/mFrom/mTo)
Private Function ComputeEditsCyr(ByVal t As String) As Long
    Dim n As Long, i As Long, st As Long, k As Long, run As String, outS As String, piece As String, c As String
    Dim allCaps As Boolean, hasUpper As Boolean, hasLower As Boolean, rule As Long, fromK As Long
    mCount = 0
    ComputeEditsCyr = 0
    If RUSCHA_ABZASNI_OTKAZISH Then
        If IsRussianText(t) Then Exit Function
    End If
    n = Len(t)
    i = 1
    Do While i <= n
        If Not IsCyrCh(Mid$(t, i, 1)) Then
            i = i + 1
        Else
            st = i
            Do While i <= n
                If Not IsCyrCh(Mid$(t, i, 1)) Then Exit Do
                i = i + 1
            Loop
            run = Mid$(t, st, i - st)
            hasUpper = False: hasLower = False
            For k = 1 To Len(run)
                c = Mid$(run, k, 1)
                If CyrLower(c) <> c Then hasUpper = True Else hasLower = True
            Next k
            allCaps = (Len(run) > 1 And hasUpper And Not hasLower)
            outS = ""
            fromK = st
            rule = FindCyrRule(run)
            If rule > 0 Then
                If mCyrHasLatin(rule) Then
                    outS = mCyrLatin(rule)
                    If allCaps Then
                        outS = UpperLat(outS)
                    ElseIf CyrLower(Left$(run, 1)) <> Left$(run, 1) And LCase$(outS) = outS And Len(outS) > 0 Then
                        outS = UpperLat(Left$(outS, 1)) & Mid$(outS, 2)     ' Abzas
                    End If
                    fromK = st + Len(mCyrKey(rule))
                Else
                    fromK = i       ' istisno: o'girilmaydi
                    outS = run
                End If
            End If
            For k = fromK To i - 1
                c = Mid$(t, k, 1)
                piece = CyrChar(t, k, st)
                If Len(piece) > 0 And CyrLower(c) <> c Then
                    If allCaps Then
                        piece = UpperLat(piece)
                    Else
                        piece = UpperLat(Left$(piece, 1)) & Mid$(piece, 2)
                    End If
                End If
                outS = outS & piece
            Next k
            If outS <> run Then AddEdit st - 1, Len(run), run, outS
        End If
    Loop
    ComputeEditsCyr = mCount
End Function

'==========================================================================
'  LOTIN -> KIRILL  (converter.js dagi latToCyr bilan bir xil)
'  Yangi (o-umlaut, s-sedil ...) va amaldagi (o', sh, ch) lotin qabul qilinadi.
'==========================================================================
Private Function LcChar(ByVal c As String) As String
    Select Case AscW(c) And &HFFFF&
        Case &H61: LcChar = ChrW(&H430)
        Case &H62: LcChar = ChrW(&H431)
        Case &H64: LcChar = ChrW(&H434)
        Case &H65: LcChar = ChrW(&H435)
        Case &H66: LcChar = ChrW(&H444)
        Case &H67: LcChar = ChrW(&H433)
        Case &H68: LcChar = ChrW(&H4B3)
        Case &H69: LcChar = ChrW(&H438)
        Case &H6A: LcChar = ChrW(&H436)
        Case &H6B: LcChar = ChrW(&H43A)
        Case &H6C: LcChar = ChrW(&H43B)
        Case &H6D: LcChar = ChrW(&H43C)
        Case &H6E: LcChar = ChrW(&H43D)
        Case &H6F: LcChar = ChrW(&H43E)
        Case &H70: LcChar = ChrW(&H43F)
        Case &H71: LcChar = ChrW(&H49B)
        Case &H72: LcChar = ChrW(&H440)
        Case &H73: LcChar = ChrW(&H441)
        Case &H74: LcChar = ChrW(&H442)
        Case &H75: LcChar = ChrW(&H443)
        Case &H76: LcChar = ChrW(&H432)
        Case &H78: LcChar = ChrW(&H445)
        Case &H79: LcChar = ChrW(&H439)
        Case &H7A: LcChar = ChrW(&H437)
        Case &HF6: LcChar = ChrW(&H45E)
        Case &H11F: LcChar = ChrW(&H493)
        Case &H15F: LcChar = ChrW(&H448)
        Case &HE7: LcChar = ChrW(&H447)
        Case Else: LcChar = c
    End Select
End Function

Private Function LowerLat(ByVal c As String) As String
    Select Case AscW(c) And &HFFFF&
        Case &HD6: LowerLat = ChrW(&HF6)
        Case &H11E: LowerLat = ChrW(&H11F)
        Case &H15E: LowerLat = ChrW(&H15F)
        Case &HC7: LowerLat = ChrW(&HE7)
        Case Else: LowerLat = LCase$(c)
    End Select
End Function

Private Function LowerLatStr(ByVal s As String) As String
    Dim i As Long, o As String
    For i = 1 To Len(s)
        o = o & LowerLat(Mid$(s, i, 1))
    Next i
    LowerLatStr = o
End Function

Private Function IsNewLetter(ByVal c As String) As Boolean
    If Len(c) = 0 Then IsNewLetter = False: Exit Function
    Select Case AscW(c) And &HFFFF&
        Case &HF6, &H11F, &H15F, &HE7: IsNewLetter = True
        Case Else: IsNewLetter = False
    End Select
End Function

' o'zbek lotin harfi (c va w yo'q), kichik harf kutiladi
Private Function IsUzLat(ByVal c As String) As Boolean
    If Len(c) = 0 Then IsUzLat = False: Exit Function
    If IsNewLetter(c) Then IsUzLat = True: Exit Function
    IsUzLat = (InStr(1, "abdefghijklmnopqrstuvxyz", c, vbBinaryCompare) > 0)
End Function

Private Function CyrUpper(ByVal c As String) As String
    Dim a As Long
    a = AscW(c) And &HFFFF&
    If a >= &H430 And a <= &H44F Then
        CyrUpper = ChrW(a - &H20)
    ElseIf a >= &H450 And a <= &H45F Then
        CyrUpper = ChrW(a - &H50)
    ElseIf a >= &H461 And a <= &H4FF And (a Mod 2 = 1) Then
        CyrUpper = ChrW(a - 1)
    Else
        CyrUpper = UCase$(c)
    End If
End Function

Private Function CyrUpperStr(ByVal s As String) As String
    Dim i As Long, o As String
    For i = 1 To Len(s)
        o = o & CyrUpper(Mid$(s, i, 1))
    Next i
    CyrUpperStr = o
End Function

Private Function IsCyrVowelOut(ByVal c As String) As Boolean
    If Len(c) = 0 Then IsCyrVowelOut = False: Exit Function
    Select Case AscW(c) And &HFFFF&
        Case &H430, &H435, &H451, &H438, &H43E, &H443, &H44D, &H44E, &H44F, &H45E: IsCyrVowelOut = True
        Case Else: IsCyrVowelOut = False
    End Select
End Function

' eski lotin kalit shakli: kichik harf, yangi harflar -> sh/ch/o'/g', apostroflar -> '
Private Function LatKey(ByVal w As String) As String
    Dim i As Long, c As String, o As String
    For i = 1 To Len(w)
        c = LowerLat(Mid$(w, i, 1))
        Select Case AscW(c) And &HFFFF&
            Case &H15F: o = o & "sh"
            Case &HE7: o = o & "ch"
            Case &HF6: o = o & "o'"
            Case &H11F: o = o & "g'"
            Case Else
                If IsApos(c) Then o = o & "'" Else o = o & c
        End Select
    Next i
    LatKey = o
End Function

Private Sub AddLatCyr(ByVal e As String)
    Dim p As Long, k As String, v As String, star As Boolean
    e = Trim$(e)
    p = InStr(e, "=")
    If p <= 1 Then Exit Sub
    k = Trim$(Left$(e, p - 1)): v = CyrLowerStr(Trim$(Mid$(e, p + 1)))
    If Right$(k, 1) = "*" Then star = True: k = Left$(k, Len(k) - 1)
    k = LatKey(k)
    If Len(k) = 0 Or Len(v) = 0 Then Exit Sub
    mLcCount = mLcCount + 1
    If mLcCount > UBound(mLcCyr) Then
        ReDim Preserve mLcCyr(1 To mLcCount + 100): ReDim Preserve mLcPrefix(1 To mLcCount + 100)
    End If
    mLcCyr(mLcCount) = v
    mLcPrefix(mLcCount) = (star Or Len(k) >= 5)
    On Error Resume Next
    mLcIndex.Remove k
    On Error GoTo 0
    mLcIndex.Add mLcCount, k
End Sub

' Lug'at: kirill istisnolarining teskarisi (oddiy so'zlar), keyin standart lug'at (ustun)
Private Sub LoadLatCyr()
    Dim i As Long, e As String, p As Long, lat As String, k As Long, ok As Boolean, c As String
    Dim users As Collection, offs As Collection, v As Variant
    Set mLcIndex = New Collection
    mLcCount = 0
    ReDim mLcCyr(1 To 400): ReDim mLcPrefix(1 To 400)
    InitCyrDefaults
    ReadState users, offs, CyrExceptionsPath()
    For i = 1 To mCyrDefCount + users.Count
        If i <= mCyrDefCount Then e = mCyrDef(i) Else e = CStr(users(i - mCyrDefCount))
        p = InStr(e, "=")
        If p > 1 Then
            lat = Trim$(Mid$(e, p + 1))
            ok = (Len(lat) > 0 And LowerLatStr(lat) = lat)
            For k = 1 To Len(lat)
                c = Mid$(lat, k, 1)
                If Not (IsUzLat(c) Or c = "c" Or IsApos(c) Or c = "-") Then ok = False
            Next k
            If ok Then AddLatCyr lat & "=" & Trim$(Left$(e, p - 1))
        End If
    Next i
    InitLatCyr1
    InitLatCyr2
    InitLatCyr3
End Sub

Private Function LcIndexOf(ByVal k As String) As Long
    On Error GoTo Yoq
    LcIndexOf = mLcIndex.Item(k)
    Exit Function
Yoq:
    LcIndexOf = 0
End Function

' Eng uzun mos kalit: natija - lug'at raqami, keyLen - kalit uzunligi
Private Function FindLatRule(ByVal key As String, ByRef keyLen As Long) As Long
    Dim ln As Long, idx As Long
    FindLatRule = 0
    If mLcIndex Is Nothing Then Exit Function
    For ln = Len(key) To 1 Step -1
        idx = LcIndexOf(Left$(key, ln))
        If idx > 0 Then
            If ln = Len(key) Or mLcPrefix(idx) Then FindLatRule = idx: keyLen = ln: Exit Function
        End If
    Next ln
End Function

Private Function ChAt(ByVal s As String, ByVal i As Long) As String
    If i < 1 Or i > Len(s) Then ChAt = "" Else ChAt = Mid$(s, i, 1)
End Function

' Bitta so'z (harflar + ichki apostroflar) -> kirill; chet so'z bo'lsa "" va isNull=True
Private Function LatWordToCyr(ByVal run As String, ByVal oldMode As Boolean, ByRef isNull As Boolean) As String
    Dim lw As String, o As String, i As Long, n As Long, c As String, nx As String, nx2 As String
    Dim a3 As String, a4 As String, letters As String, k As Long, allCaps As Boolean, capFirst As Boolean
    Dim idx As Long, keyLen As Long, consumed As Long, kk As Long, cyr As String, atStart As Boolean, pv As Boolean
    isNull = False
    lw = LowerLatStr(run)
    n = Len(lw)
    For k = 1 To n
        c = Mid$(run, k, 1)
        If IsLetterCh(c) Then letters = letters & c
    Next k
    allCaps = (Len(letters) > 1 And HasNoLower(letters))
    capFirst = (Left$(run, 1) <> Left$(lw, 1))
    If Not allCaps Then
        For k = 2 To Len(run)
            If (Mid$(run, k - 1, 1) Like "[a-z]") And (Mid$(run, k, 1) Like "[A-Z]") Then isNull = True: Exit Function
        Next k
    End If

    i = 1
    idx = FindLatRule(LatKey(run), keyLen)
    If idx > 0 Then
        consumed = 0: kk = 0
        Do While kk < keyLen And consumed < n
            consumed = consumed + 1
            If IsNewLetter(Mid$(lw, consumed, 1)) Then kk = kk + 2 Else kk = kk + 1
        Loop
        cyr = mLcCyr(idx)
        If consumed < n And Right$(cyr, 1) = ChrW(&H44C) Then cyr = Left$(cyr, Len(cyr) - 1)
        o = cyr
        i = consumed + 1
    End If

    Do While i <= n
        c = Mid$(lw, i, 1): nx = ChAt(lw, i + 1): nx2 = ChAt(lw, i + 2)
        atStart = (Len(o) = 0)
        pv = IsCyrVowelOut(Right$(o, 1))
        If (c = "o" Or c = "g") And IsApos(nx) Then
            If c = "o" Then o = o & ChrW(&H45E) Else o = o & ChrW(&H493)
            i = i + 2
        ElseIf oldMode And c = "s" And nx = "h" Then
            o = o & ChrW(&H448): i = i + 2
        ElseIf oldMode And c = "c" And nx = "h" Then
            o = o & ChrW(&H447): i = i + 2
        ElseIf IsApos(c) Then
            If ChAt(lw, i - 1) = "s" And nx = "h" Then
                i = i + 1                                   ' s'h -> s h
            ElseIf i > 1 And IsUzLat(nx) Then
                o = o & ChrW(&H44A): i = i + 1              ' tutuq -> qattiq belgi
            Else
                o = o & Mid$(run, i, 1): i = i + 1
            End If
        ElseIf (c = "t" Or c = "k") And nx = "s" And nx2 = "i" And (ChAt(lw, i + 3) = "y" Or ChAt(lw, i + 3) = "o") Then
            If c = "t" Then o = o & ChrW(&H446) Else o = o & ChrW(&H43A) & ChrW(&H446)
            i = i + 2                                       ' -tsiya, -ksiya
        ElseIf c = "t" And nx = "s" And i > 1 And TsIsC(lw, i) Then
            o = o & ChrW(&H446): i = i + 2                  ' atseton, defitsit
        ElseIf c = "e" And Mid$(lw, i + 1, 3) = "nsi" And (ChAt(lw, i + 4) = "y" Or Mid$(lw, i + 4, 2) = "al") _
               And ChAt(lw, i - 1) <> "p" And Len(o) > 0 Then
            If pv Then o = o & ChrW(&H44D) Else o = o & ChrW(&H435)
            o = o & ChrW(&H43D) & ChrW(&H446): i = i + 3    ' konferensiya, konfidensial
        ElseIf Mid$(lw, i, 5) = "porsi" Then
            o = o & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H446): i = i + 4
        ElseIf Mid$(lw, i, 8) = "stansiya" Then
            o = o & ChrW(&H441) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H446): i = i + 5
        ElseIf c = "y" Then
            If nx = "o" And Not IsApos(nx2) Then
                o = o & ChrW(&H451): i = i + 2
            ElseIf nx = "a" Then
                o = o & ChrW(&H44F): i = i + 2
            ElseIf nx = "u" Then
                o = o & ChrW(&H44E): i = i + 2
            ElseIf nx = "e" Then
                If atStart Or pv Then o = o & ChrW(&H435) Else o = o & ChrW(&H44A) & ChrW(&H435)
                i = i + 2
            Else
                o = o & ChrW(&H439): i = i + 1
            End If
        ElseIf c = "e" Then
            If atStart Or pv Then o = o & ChrW(&H44D) Else o = o & ChrW(&H435)
            i = i + 1
        Else
            If IsUzLat(c) Then o = o & LcChar(c) Else o = o & Mid$(run, i, 1)
            i = i + 1
        End If
    Loop
    If allCaps Then
        LatWordToCyr = CyrUpperStr(o)
    ElseIf capFirst And Len(o) > 0 Then
        LatWordToCyr = CyrUpper(Left$(o, 1)) & Mid$(o, 2)
    Else
        LatWordToCyr = o
    End If
End Function

Private Function IsOG(ByVal c As String) As Boolean
    IsOG = (c = "o" Or c = "g" Or c = "O" Or c = "G")
End Function

Private Function HasNoLower(ByVal s As String) As Boolean
    Dim i As Long, c As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If LowerLat(c) = c And UCase$(c) <> c Then HasNoLower = False: Exit Function
    Next i
    HasNoLower = True
End Function

' ts + e  yoki  ts + i + undosh (z, n dan tashqari; nt mumkin)  yoki  ts + i + a/e  -> ts
Private Function TsIsC(ByVal lw As String, ByVal i As Long) As Boolean
    Dim nx2 As String, a3 As String, a4 As String
    nx2 = ChAt(lw, i + 2): a3 = ChAt(lw, i + 3): a4 = ChAt(lw, i + 4)
    TsIsC = False
    If nx2 = "e" Then TsIsC = True: Exit Function
    If nx2 <> "i" Or Len(a3) = 0 Then Exit Function
    If a3 = "z" Then Exit Function
    If a3 = "n" And a4 <> "t" Then Exit Function
    If Not IsUzLat(a3) Then Exit Function
    If InStr("aeiou", a3) = 0 Or a3 = "a" Or a3 = "e" Then TsIsC = True
End Function

' Lotin -> kirill o'zgarishlari (mPos/mLen/mFrom/mTo)
Private Function ComputeEditsLatCyr(ByVal t As String) As Long
    Dim n As Long, i As Long, st As Long, tok As String, lead As Long, lim As Long, tt As Long, core As String
    Dim j As Long, ws As Long, run As String, rl As String, q As Long, ch As String, foreign As Boolean
    Dim hasNew As Boolean, cyr As String, isNull As Boolean, quoteOpen As Boolean, keyLen As Long
    mCount = 0
    ComputeEditsLatCyr = 0
    If INGLIZCHA_ABZASNI_OTKAZISH Then
        If IsEnglishText(t) Then Exit Function
    End If
    n = Len(t)
    i = 1
    Do While i <= n
        If IsSpaceCh(Mid$(t, i, 1)) Then
            i = i + 1
        Else
            st = i
            Do While i <= n
                If IsSpaceCh(Mid$(t, i, 1)) Then Exit Do
                i = i + 1
            Loop
            tok = Mid$(t, st, i - st)
            ' qo'shtirnoqlar
            lead = 1
            Do While lead <= Len(tok)
                If IsLetterCh(Mid$(tok, lead, 1)) Or IsDigitCh(Mid$(tok, lead, 1)) Then Exit Do
                If IsApos(Mid$(tok, lead, 1)) Then quoteOpen = True
                lead = lead + 1
            Loop
            lim = Len(tok): tt = Len(tok)
            Do While tt > lead
                If IsLetterCh(Mid$(tok, tt, 1)) Or IsApos(Mid$(tok, tt, 1)) Then Exit Do
                tt = tt - 1
            Loop
            If tt > lead Then
                If IsApos(Mid$(tok, tt, 1)) And quoteOpen Then lim = tt - 1: quoteOpen = False
            End If
            core = CoreOf(tok)
            If Len(core) > 0 And Not LooksLikeLinkOrCode(tok) And Not HasCyr(core) Then
                If Len(FindException(core)) = 0 Or FindLatRule(LatKey(core), keyLen) > 0 Then
                    j = lead
                    Do While j <= lim
                        If Not IsLetterCh(Mid$(tok, j, 1)) Then
                            j = j + 1
                        Else
                            ws = j
                            Do While j <= lim
                                ch = Mid$(tok, j, 1)
                                If IsLetterCh(ch) Then
                                    j = j + 1
                                ElseIf IsApos(ch) And (IsLetterCh(ChAt(tok, j + 1)) Or IsOG(ChAt(tok, j - 1))) Then
                                    j = j + 1
                                Else
                                    Exit Do
                                End If
                            Loop
                            run = Mid$(tok, ws, j - ws)
                            rl = LowerLatStr(run)
                            hasNew = False: foreign = False
                            For q = 1 To Len(rl)
                                If IsNewLetter(Mid$(rl, q, 1)) Then hasNew = True
                            Next q
                            For q = 1 To Len(rl)
                                ch = Mid$(rl, q, 1)
                                If Not IsApos(ch) Then
                                    If ch = "c" Then
                                        If Not (ChAt(rl, q + 1) = "h" And Not hasNew) Then foreign = True: Exit For
                                    ElseIf Not IsUzLat(ch) Then
                                        foreign = True: Exit For
                                    End If
                                End If
                            Next q
                            If Not foreign Then
                                cyr = LatWordToCyr(run, Not hasNew, isNull)
                                If Not isNull And cyr <> run Then AddEdit st - 1 + ws - 1, Len(run), run, cyr
                            End If
                        End If
                    Loop
                End If
            End If
        End If
    Loop
    ComputeEditsLatCyr = mCount
End Function

'==========================================================================
'  YANGI LOTIN -> ESKI (AMALDAGI) LOTIN:  o-umlaut -> o', s-sedil -> sh ...
'==========================================================================
Private Function ComputeEditsNewToOld(ByVal t As String) As Long
    Dim n As Long, i As Long, st As Long, tok As String, o As String, k As Long, has As Boolean
    mCount = 0
    n = Len(t)
    i = 1
    Do While i <= n
        If IsSpaceCh(Mid$(t, i, 1)) Then
            i = i + 1
        Else
            st = i
            Do While i <= n
                If IsSpaceCh(Mid$(t, i, 1)) Then Exit Do
                i = i + 1
            Loop
            tok = Mid$(t, st, i - st)
            has = False
            For k = 1 To Len(tok)
                If IsNewLetter(LowerLat(Mid$(tok, k, 1))) Then has = True: Exit For
            Next k
            If has Then
                o = RevertWord(tok)
                If o <> tok Then AddEdit st - 1, Len(tok), tok, o
            End If
        End If
    Loop
    ComputeEditsNewToOld = mCount
End Function

'==========================================================================
'  YOZISH PAYTIDA O'GIRISH
'  h, Shift+h, ' va ` tugmalari makrosga bog'lanadi. Makros klaviaturaning
'  joriy tilini aniqlab, tugma beradigan harfni yozadi (kirillcha ham ishlaydi),
'  so'ng oxirgi harflarni tekshiradi:  sh -> s-sedil,  o' -> o-umlaut ...
'  Bekor qilish: o'sha tugmani yana bir marta bosing (s-sedil + h -> sh)
'  yoki Ctrl+Z.
'==========================================================================
Public Function AppNomi() As String
    AppNomi = ChrW(&HD6) & "zbek" & ChrW(&HE7) & "a"
End Function

Public Sub YangiAlifbo_KirillIstisnolar()
    Dim choice As String, users As Collection, offs As Collection
    Do
        InitCyrDefaults
        ReadState users, offs, CyrExceptionsPath()
        choice = InputBox( _
            "Kirilldan o'girish istisnolari" & vbCrLf & _
            "Mening yozuvlarim: " & users.Count & " ta,  standart: " & (mCyrDefCount - offs.Count) & " / " & mCyrDefCount & vbCrLf & vbCrLf & _
            "1 - Ro'yxatni ko'rish (raqamlar bilan)" & vbCrLf & _
            "2 - Belgilangan kirill so'zini qo'shish" & vbCrLf & _
            "3 - Yozuvni o'chirish (raqami bo'yicha)" & vbCrLf & _
            "4 - Standart yozuvni yoqish / o'chirish" & vbCrLf & _
            "5 - Faylni Bloknotda tahrirlash" & vbCrLf & vbCrLf & _
            "Raqamni kiriting:", AppNomi() & " - kirill istisnolari", "1")
        Select Case Trim$(choice)
            Case "": Exit Do
            Case "1": ShowCyrList: Exit Do
            Case "2": YangiAlifbo_KirillIstisnoQoshish
            Case "3": PromptCyrRemove
            Case "4": PromptCyrToggle
            Case "5": OpenCyrInNotepad: Exit Do
            Case Else: UMsg "1 dan 5 gacha raqam kiriting.", 48
        End Select
    Loop
End Sub

' Belgilangan (yoki kursor turgan) kirill so'zini istisnoga qo'shadi
Private Sub ShowCyrList()
    Dim users As Collection, offs As Collection, i As Long, t As String
    InitCyrDefaults
    ReadState users, offs, CyrExceptionsPath()
    t = AppNomi() & " - kirill istisnolari" & vbCr & vbCr
    t = t & "MENING YOZUVLARIM (" & users.Count & ")" & vbCr
    If users.Count = 0 Then t = t & "    (hali yo'q)" & vbCr
    For i = 1 To users.Count
        t = t & "  " & i & ")  " & CyrEntryText(CStr(users(i))) & vbCr
    Next i
    t = t & vbCr & "STANDART YOZUVLAR" & vbCr
    For i = 1 To mCyrDefCount
        t = t & "  " & i & ")  " & CyrEntryText(mCyrDef(i))
        If KeyInColl(offs, ExcKey(mCyrDef(i))) Then t = t & "    [o'chirilgan]"
        t = t & vbCr
    Next i
    t = t & vbCr & "so'z = o'girilmaydi;  so'z* = shu bilan boshlanganlar;  o'zak=Lotin = maxsus shakl." & vbCr & _
        "Fayl: " & CyrExceptionsPath()
    ShowText t
End Sub

Private Function CyrEntryText(ByVal e As String) As String
    Dim p As Long
    p = InStr(e, "=")
    If p > 0 Then CyrEntryText = Left$(e, p - 1) & "  ->  " & Mid$(e, p + 1) Else CyrEntryText = e & "  (o'girilmaydi)"
End Function

Private Sub PromptCyrRemove()
    Dim users As Collection, offs As Collection, ans As String, n As Long, e As String
    ReadState users, offs, CyrExceptionsPath()
    If users.Count = 0 Then UMsg "Sizning yozuvlaringiz yo'q.": Exit Sub
    ans = Trim$(InputBox("O'chiriladigan yozuv raqamini kiriting (1 - " & users.Count & ")." & vbCrLf & _
                         "Raqamlarni ko'rish: menyudagi 1-band.", AppNomi() & " - kirill istisnosini o'chirish"))
    If Len(ans) = 0 Then Exit Sub
    If Not IsNumeric(ans) Then UMsg "Raqam kiriting.", 48: Exit Sub
    n = CLng(ans)
    If n < 1 Or n > users.Count Then UMsg "1 dan " & users.Count & " gacha raqam kiriting.", 48: Exit Sub
    e = CStr(users(n))
    users.Remove n
    SaveState users, offs, CyrExceptionsPath()
    UMsg "O'chirildi: " & CyrEntryText(e)
End Sub

Private Sub PromptCyrToggle()
    Dim users As Collection, offs As Collection, ans As String, n As Long, t As String, i As Long, k As String, p As Long
    InitCyrDefaults
    ReadState users, offs, CyrExceptionsPath()
    For i = 1 To mCyrDefCount
        p = InStr(mCyrDef(i), "=")
        t = t & i & ")" & IIf(KeyInColl(offs, ExcKey(mCyrDef(i))), "[-]", "") & Mid$(mCyrDef(i), p + 1) & "  "
        If i Mod 4 = 0 Then t = t & vbCrLf
    Next i
    ans = Trim$(InputBox(t & vbCrLf & vbCrLf & "[-] = o'chirilgan." & vbCrLf & _
                         "Yoqish/o'chirish uchun raqamni kiriting:", AppNomi() & " - standart kirill yozuvlari"))
    If Len(ans) = 0 Then Exit Sub
    If Not IsNumeric(ans) Then UMsg "Raqam kiriting.", 48: Exit Sub
    n = CLng(ans)
    If n < 1 Or n > mCyrDefCount Then UMsg "1 dan " & mCyrDefCount & " gacha raqam kiriting.", 48: Exit Sub
    k = ExcKey(mCyrDef(n))
    If KeyInColl(offs, k) Then
        RemoveKey offs, k
        SaveState users, offs, CyrExceptionsPath()
        UMsg CyrEntryText(mCyrDef(n)) & vbCrLf & "qayta yoqildi."
    Else
        p = InStr(mCyrDef(n), "=")
        offs.Add Left$(mCyrDef(n), p - 1)
        SaveState users, offs, CyrExceptionsPath()
        UMsg CyrEntryText(mCyrDef(n)) & vbCrLf & "o'chirildi - endi odatdagidek o'giriladi."
    End If
End Sub

Private Sub OpenCyrInNotepad()
    Dim users As Collection, offs As Collection
    If Len(Dir$(CyrExceptionsPath())) = 0 Then
        ReadState users, offs, CyrExceptionsPath()
        SaveState users, offs, CyrExceptionsPath()
    End If
    Shell "notepad.exe """ & CyrExceptionsPath() & """", vbNormalFocus
    mInit = False
End Sub

Private Function HasCyr(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If IsCyrCh(Mid$(s, i, 1)) Then HasCyr = True: Exit Function
    Next i
    HasCyr = False
End Function

Private Sub RefreshRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

Private Sub InitLatCyr1()
    AddLatCyr "abzas=" & ChrW(&H430) & ChrW(&H431) & ChrW(&H437) & ChrW(&H430) & ChrW(&H446)
    AddLatCyr "aksent=" & ChrW(&H430) & ChrW(&H43A) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442)
    AddLatCyr "avianoses=" & ChrW(&H430) & ChrW(&H432) & ChrW(&H438) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H435) & ChrW(&H446)
    AddLatCyr "batsilla=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H430)
    AddLatCyr "biomitsin=" & ChrW(&H431) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43D)
    AddLatCyr "botsman=" & ChrW(&H431) & ChrW(&H43E) & ChrW(&H446) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H43D)
    AddLatCyr "bronenoses=" & ChrW(&H431) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H435) & ChrW(&H446)
    AddLatCyr "brutselloz=" & ChrW(&H431) & ChrW(&H440) & ChrW(&H443) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H437)
    AddLatCyr "cherepitsa=" & ChrW(&H447) & ChrW(&H435) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43F) & ChrW(&H438) & ChrW(&H446) & ChrW(&H430)
    AddLatCyr "dotsent=" & ChrW(&H434) & ChrW(&H43E) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442)
    AddLatCyr "fransuz=" & ChrW(&H444) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H443) & ChrW(&H437)
    AddLatCyr "gaubitsa=" & ChrW(&H433) & ChrW(&H430) & ChrW(&H443) & ChrW(&H431) & ChrW(&H438) & ChrW(&H446) & ChrW(&H430)
    AddLatCyr "gers=" & ChrW(&H433) & ChrW(&H435) & ChrW(&H440) & ChrW(&H446)
    AddLatCyr "glitserin=" & ChrW(&H433) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H446) & ChrW(&H435) & ChrW(&H440) & ChrW(&H438) & ChrW(&H43D)
    AddLatCyr "gorchitsa=" & ChrW(&H433) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H447) & ChrW(&H438) & ChrW(&H446) & ChrW(&H430)
    AddLatCyr "gusenitsa=" & ChrW(&H433) & ChrW(&H443) & ChrW(&H441) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H446) & ChrW(&H430)
    AddLatCyr "inersiya=" & ChrW(&H438) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H440) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "inssenirovka=" & ChrW(&H438) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H432) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "kalsiy=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H446) & ChrW(&H438) & ChrW(&H439)
    AddLatCyr "kansler=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "koeffitsient=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H44D) & ChrW(&H444) & ChrW(&H444) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442)
    AddLatCyr "konferens-zal=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H444) & ChrW(&H435) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H446) & "-" & ChrW(&H437) & ChrW(&H430) & ChrW(&H43B)
    AddLatCyr "konsepsiya=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43F) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "konsern=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H435) & ChrW(&H440) & ChrW(&H43D)
    AddLatCyr "konsert=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H435) & ChrW(&H440) & ChrW(&H442)
    AddLatCyr "konslager=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H433) & ChrW(&H435) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "kultivatsiya=" & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438) & ChrW(&H432) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "kvars=" & ChrW(&H43A) & ChrW(&H432) & ChrW(&H430) & ChrW(&H440) & ChrW(&H446)
    AddLatCyr "litsey=" & ChrW(&H43B) & ChrW(&H438) & ChrW(&H446) & ChrW(&H435) & ChrW(&H439)
    AddLatCyr "lotsman=" & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H446) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H43D)
    AddLatCyr "marganes=" & ChrW(&H43C) & ChrW(&H430) & ChrW(&H440) & ChrW(&H433) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H446)
    AddLatCyr "minonoses=" & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H435) & ChrW(&H446)
    AddLatCyr "munitsipalitet=" & ChrW(&H43C) & ChrW(&H443) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43F) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H442) & ChrW(&H435) & ChrW(&H442)
    AddLatCyr "ofitsiant=" & ChrW(&H43E) & ChrW(&H444) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H442)
    AddLatCyr "penitsillin=" & ChrW(&H43F) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43D)
    AddLatCyr "platsdarm=" & ChrW(&H43F) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H446) & ChrW(&H434) & ChrW(&H430) & ChrW(&H440) & ChrW(&H43C)
    AddLatCyr "platskarta=" & ChrW(&H43F) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H446) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H440) & ChrW(&H442) & ChrW(&H430)
    AddLatCyr "politsmeyster=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H446) & ChrW(&H43C) & ChrW(&H435) & ChrW(&H439) & ChrW(&H441) & ChrW(&H442) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "ranes=" & ChrW(&H440) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H446)
    AddLatCyr "ritsar=" & ChrW(&H440) & ChrW(&H438) & ChrW(&H446) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "sellofan=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H444) & ChrW(&H430) & ChrW(&H43D)
    AddLatCyr "selluloid=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H438) & ChrW(&H434)
    AddLatCyr "selluloza=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H437) & ChrW(&H430)
    AddLatCyr "selsiy=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H441) & ChrW(&H438) & ChrW(&H439)
    AddLatCyr "sement=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43C) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442)
    AddLatCyr "sentner=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "sentrifuga=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H440) & ChrW(&H438) & ChrW(&H444) & ChrW(&H443) & ChrW(&H433) & ChrW(&H430)
    AddLatCyr "senz=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H437)
    AddLatCyr "senzor=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H437) & ChrW(&H43E) & ChrW(&H440)
    AddLatCyr "senzura=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H437) & ChrW(&H443) & ChrW(&H440) & ChrW(&H430)
    AddLatCyr "sex=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H445)
    AddLatCyr "shnitsel=" & ChrW(&H448) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "shveysar=" & ChrW(&H448) & ChrW(&H432) & ChrW(&H435) & ChrW(&H439) & ChrW(&H446) & ChrW(&H430) & ChrW(&H440)
    AddLatCyr "siferblat=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H444) & ChrW(&H435) & ChrW(&H440) & ChrW(&H431) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H442)
    AddLatCyr "silindr=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H440)
    AddLatCyr "silindrik=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H440) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "sink=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H43A)
    AddLatCyr "sirk=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H440) & ChrW(&H43A)
    AddLatCyr "sirkulyar=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H440) & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44F) & ChrW(&H440)
    AddLatCyr "sisterna=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H441) & ChrW(&H442) & ChrW(&H435) & ChrW(&H440) & ChrW(&H43D) & ChrW(&H430)
    AddLatCyr "sitrus=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H442) & ChrW(&H440) & ChrW(&H443) & ChrW(&H441)
    AddLatCyr "sotsiologik=" & ChrW(&H441) & ChrW(&H43E) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "sotsiologiya=" & ChrW(&H441) & ChrW(&H43E) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "ssenariy=" & ChrW(&H441) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H440) & ChrW(&H438) & ChrW(&H439)
    AddLatCyr "kultivator=" & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438) & ChrW(&H432) & ChrW(&H430) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H440)
    AddLatCyr "kuryer=" & ChrW(&H43A) & ChrW(&H443) & ChrW(&H440) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "lager=" & ChrW(&H43B) & ChrW(&H430) & ChrW(&H433) & ChrW(&H435) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "losos=" & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H44C)
    AddLatCyr "mebel=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H431) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "medal=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H434) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "medalyon=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H434) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "menshevik=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H448) & ChrW(&H435) & ChrW(&H432) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "menshevizm=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H448) & ChrW(&H435) & ChrW(&H432) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "migren=" & ChrW(&H43C) & ChrW(&H438) & ChrW(&H433) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "mikrofilm=" & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H444) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C)
    AddLatCyr "mil=" & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "model=" & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H434) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "neft=" & ChrW(&H43D) & ChrW(&H435) & ChrW(&H444) & ChrW(&H442) & ChrW(&H44C)
    AddLatCyr "nikel=" & ChrW(&H43D) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "nippel=" & ChrW(&H43D) & ChrW(&H438) & ChrW(&H43F) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "nol=" & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "noyabr=" & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H44F) & ChrW(&H431) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "oktabr=" & ChrW(&H43E) & ChrW(&H43A) & ChrW(&H442) & ChrW(&H44F) & ChrW(&H431) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "palto=" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H43E)
    AddLatCyr "panel=" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "parallel=" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "parol=" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "parter=" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H440) & ChrW(&H442) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "patrul=" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H442) & ChrW(&H440) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "pavilyon=" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H432) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "pedal=" & ChrW(&H43F) & ChrW(&H435) & ChrW(&H434) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "plastir=" & ChrW(&H43F) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "pochtalyon=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H447) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "porshen=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H448) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "portfel=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H442) & ChrW(&H444) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "povest=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H432) & ChrW(&H435) & ChrW(&H441) & ChrW(&H442) & ChrW(&H44C)
    AddLatCyr "predoxranitel=" & ChrW(&H43F) & ChrW(&H440) & ChrW(&H435) & ChrW(&H434) & ChrW(&H43E) & ChrW(&H445) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "premyera=" & ChrW(&H43F) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43C) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440) & ChrW(&H430)
    AddLatCyr "pristan=" & ChrW(&H43F) & ChrW(&H440) & ChrW(&H438) & ChrW(&H441) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "puls=" & ChrW(&H43F) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H441)
    AddLatCyr "pyesa=" & ChrW(&H43F) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H441) & ChrW(&H430)
    AddLatCyr "rels=" & ChrW(&H440) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H441)
    AddLatCyr "relyef=" & ChrW(&H440) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H444)
    AddLatCyr "rentabel=" & ChrW(&H440) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H430) & ChrW(&H431) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "rezba=" & ChrW(&H440) & ChrW(&H435) & ChrW(&H437) & ChrW(&H44C) & ChrW(&H431) & ChrW(&H430)
    AddLatCyr "rol=" & ChrW(&H440) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "royal=" & ChrW(&H440) & ChrW(&H43E) & ChrW(&H44F) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "rul=" & ChrW(&H440) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "seld=" & ChrW(&H441) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H434)
    AddLatCyr "sentabr=" & ChrW(&H441) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H44F) & ChrW(&H431) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "shinel=" & ChrW(&H448) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "shpatel=" & ChrW(&H448) & ChrW(&H43F) & ChrW(&H430) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "shpilka=" & ChrW(&H448) & ChrW(&H43F) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "shpindel=" & ChrW(&H448) & ChrW(&H43F) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "shtapel=" & ChrW(&H448) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "shtempel=" & ChrW(&H448) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43C) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "shtepsel=" & ChrW(&H448) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43F) & ChrW(&H441) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "spektakl=" & ChrW(&H441) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "spiral=" & ChrW(&H441) & ChrW(&H43F) & ChrW(&H438) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "sterjen=" & ChrW(&H441) & ChrW(&H442) & ChrW(&H435) & ChrW(&H440) & ChrW(&H436) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "sudya=" & ChrW(&H441) & ChrW(&H443) & ChrW(&H434) & ChrW(&H44C) & ChrW(&H44F)
    AddLatCyr "sulfat=" & ChrW(&H441) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H444) & ChrW(&H430) & ChrW(&H442)
    AddLatCyr "tabel=" & ChrW(&H442) & ChrW(&H430) & ChrW(&H431) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "tekstil=" & ChrW(&H442) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "tokar=" & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "tulen=" & ChrW(&H442) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "tunnel=" & ChrW(&H442) & ChrW(&H443) & ChrW(&H43D) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "umivalnik=" & ChrW(&H443) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H432) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "valeryanka=" & ChrW(&H432) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H440) & ChrW(&H44C) & ChrW(&H44F) & ChrW(&H43D) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "vals=" & ChrW(&H432) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H441)
    AddLatCyr "veksel=" & ChrW(&H432) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H441) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "velvet=" & ChrW(&H432) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H432) & ChrW(&H435) & ChrW(&H442)
    AddLatCyr "ventil=" & ChrW(&H432) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "vermishel=" & ChrW(&H432) & ChrW(&H435) & ChrW(&H440) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H448) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "vimpel=" & ChrW(&H432) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "violonchel=" & ChrW(&H432) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H447) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "volfram=" & ChrW(&H432) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H444) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43C)
    AddLatCyr "volt=" & ChrW(&H432) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442)
    AddLatCyr "volta=" & ChrW(&H432) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H430)
    AddLatCyr "voltmetr=" & ChrW(&H432) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H43C) & ChrW(&H435) & ChrW(&H442) & ChrW(&H440)
    AddLatCyr "vulgar=" & ChrW(&H432) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H433) & ChrW(&H430) & ChrW(&H440)
    AddLatCyr "yakor=" & ChrW(&H44F) & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "yanvar=" & ChrW(&H44F) & ChrW(&H43D) & ChrW(&H432) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "yuriskonsult=" & ChrW(&H44E) & ChrW(&H440) & ChrW(&H438) & ChrW(&H441) & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442)
    AddLatCyr "ansambl=" & ChrW(&H430) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H431) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "artel=" & ChrW(&H430) & ChrW(&H440) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "artikl=" & ChrW(&H430) & ChrW(&H440) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "aryergard=" & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440) & ChrW(&H433) & ChrW(&H430) & ChrW(&H440) & ChrW(&H434)
    AddLatCyr "asfalt=" & ChrW(&H430) & ChrW(&H441) & ChrW(&H444) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442)
    AddLatCyr "atelye=" & ChrW(&H430) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H435)
    AddLatCyr "avtomobil=" & ChrW(&H430) & ChrW(&H432) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H431) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "balzam=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H437) & ChrW(&H430) & ChrW(&H43C)
    AddLatCyr "banderol=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H435) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "batalyon=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "bilyard=" & ChrW(&H431) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H44F) & ChrW(&H440) & ChrW(&H434)
    AddLatCyr "binokl=" & ChrW(&H431) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "bolshevik=" & ChrW(&H431) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H448) & ChrW(&H435) & ChrW(&H432) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "budilnik=" & ChrW(&H431) & ChrW(&H443) & ChrW(&H434) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "bulvar=" & ChrW(&H431) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H432) & ChrW(&H430) & ChrW(&H440)
    AddLatCyr "dalton=" & ChrW(&H434) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "dekabr=" & ChrW(&H434) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H431) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "delfin=" & ChrW(&H434) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H444) & ChrW(&H438) & ChrW(&H43D)
    AddLatCyr "devalvatsiya=" & ChrW(&H434) & ChrW(&H435) & ChrW(&H432) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H432) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "dirijabl=" & ChrW(&H434) & ChrW(&H438) & ChrW(&H440) & ChrW(&H438) & ChrW(&H436) & ChrW(&H430) & ChrW(&H431) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "dizel=" & ChrW(&H434) & ChrW(&H438) & ChrW(&H437) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "dizel-motor=" & ChrW(&H434) & ChrW(&H438) & ChrW(&H437) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & "-" & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H440)
    AddLatCyr "duel=" & ChrW(&H434) & ChrW(&H443) & ChrW(&H44D) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "dvigatel=" & ChrW(&H434) & ChrW(&H432) & ChrW(&H438) & ChrW(&H433) & ChrW(&H430) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "emulsiya=" & ChrW(&H44D) & ChrW(&H43C) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H441) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "eskadrilya=" & ChrW(&H44D) & ChrW(&H441) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H434) & ChrW(&H440) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H44F)
    AddLatCyr "fakultativ=" & ChrW(&H444) & ChrW(&H430) & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H430) & ChrW(&H442) & ChrW(&H438) & ChrW(&H432)
    AddLatCyr "fakultet=" & ChrW(&H444) & ChrW(&H430) & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H435) & ChrW(&H442)
    AddLatCyr "falsifikator=" & ChrW(&H444) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H441) & ChrW(&H438) & ChrW(&H444) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H440)
    AddLatCyr "falsifikatsiya=" & ChrW(&H444) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H441) & ChrW(&H438) & ChrW(&H444) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "feldmarshal=" & ChrW(&H444) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H434) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H440) & ChrW(&H448) & ChrW(&H430) & ChrW(&H43B)
    AddLatCyr "feldsher=" & ChrW(&H444) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H434) & ChrW(&H448) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "festival=" & ChrW(&H444) & ChrW(&H435) & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H432) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "fevral=" & ChrW(&H444) & ChrW(&H435) & ChrW(&H432) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "filtr=" & ChrW(&H444) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H440)
    AddLatCyr "folklor=" & ChrW(&H444) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H440)
    AddLatCyr "fotoalbom=" & ChrW(&H444) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H431) & ChrW(&H43E) & ChrW(&H43C)
    AddLatCyr "fotoatelye=" & ChrW(&H444) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H430) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H435)
    AddLatCyr "gantel=" & ChrW(&H433) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "gastrol=" & ChrW(&H433) & ChrW(&H430) & ChrW(&H441) & ChrW(&H442) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "gilza=" & ChrW(&H433) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H437) & ChrW(&H430)
    AddLatCyr "gospital=" & ChrW(&H433) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H43F) & ChrW(&H438) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "grifel=" & ChrW(&H433) & ChrW(&H440) & ChrW(&H438) & ChrW(&H444) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "impuls=" & ChrW(&H438) & ChrW(&H43C) & ChrW(&H43F) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H441)
    AddLatCyr "insult=" & ChrW(&H438) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442)
End Sub

Private Sub InitLatCyr2()
    AddLatCyr "intervyu=" & ChrW(&H438) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H435) & ChrW(&H440) & ChrW(&H432) & ChrW(&H44C) & ChrW(&H44E)
    AddLatCyr "inyeksiya=" & ChrW(&H438) & ChrW(&H43D) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "italyan=" & ChrW(&H438) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H44F) & ChrW(&H43D)
    AddLatCyr "iyul=" & ChrW(&H438) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "iyun=" & ChrW(&H438) & ChrW(&H44E) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "kabel=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H431) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "kalendar=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "kanifol=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H444) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "karamel=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "kartel=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H440) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "karusel=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H440) & ChrW(&H443) & ChrW(&H441) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "katapulta=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43F) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H430)
    AddLatCyr "kinofestival=" & ChrW(&H43A) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H444) & ChrW(&H435) & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H432) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "kinofilm=" & ChrW(&H43A) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H444) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C)
    AddLatCyr "kisel=" & ChrW(&H43A) & ChrW(&H438) & ChrW(&H441) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "kitel=" & ChrW(&H43A) & ChrW(&H438) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "kobalt=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H431) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442)
    AddLatCyr "kompanyon=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43C) & ChrW(&H43F) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "konferansye=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H444) & ChrW(&H435) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H44C) & ChrW(&H435)
    AddLatCyr "syomka=" & ChrW(&H441) & ChrW(&H44A) & ChrW(&H451) & ChrW(&H43C) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "korrupsiya=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H440) & ChrW(&H443) & ChrW(&H43F) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "konyunktura=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H44A) & ChrW(&H44E) & ChrW(&H43D) & ChrW(&H43A) & ChrW(&H442) & ChrW(&H443) & ChrW(&H440) & ChrW(&H430)
    AddLatCyr "adyunktura=" & ChrW(&H430) & ChrW(&H434) & ChrW(&H44A) & ChrW(&H44E) & ChrW(&H43D) & ChrW(&H43A) & ChrW(&H442) & ChrW(&H443) & ChrW(&H440) & ChrW(&H430)
    AddLatCyr "ishoq=" & ChrW(&H438) & ChrW(&H441) & ChrW(&H4B3) & ChrW(&H43E) & ChrW(&H49B)
    AddLatCyr "mushaf=" & ChrW(&H43C) & ChrW(&H443) & ChrW(&H441) & ChrW(&H4B3) & ChrW(&H430) & ChrW(&H444)
    AddLatCyr "ashob=" & ChrW(&H430) & ChrW(&H441) & ChrW(&H4B3) & ChrW(&H43E) & ChrW(&H431)
    AddLatCyr "ishol=" & ChrW(&H438) & ChrW(&H441) & ChrW(&H4B3) & ChrW(&H43E) & ChrW(&H43B)
    AddLatCyr "mayor=" & ChrW(&H43C) & ChrW(&H430) & ChrW(&H439) & ChrW(&H43E) & ChrW(&H440)
    AddLatCyr "rayon=" & ChrW(&H440) & ChrW(&H430) & ChrW(&H439) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "mikrorayon=" & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H430) & ChrW(&H439) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "mayonez=" & ChrW(&H43C) & ChrW(&H430) & ChrW(&H439) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H437)
    AddLatCyr "yogurt=" & ChrW(&H439) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H443) & ChrW(&H440) & ChrW(&H442)
    AddLatCyr "yoga*=" & ChrW(&H439) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H430)
    AddLatCyr "poyabzal=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H439) & ChrW(&H430) & ChrW(&H431) & ChrW(&H437) & ChrW(&H430) & ChrW(&H43B)
    AddLatCyr "poyandoz=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H439) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H43E) & ChrW(&H437)
    AddLatCyr "poyafzal=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H439) & ChrW(&H430) & ChrW(&H444) & ChrW(&H437) & ChrW(&H430) & ChrW(&H43B)
    AddLatCyr "poyustun=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H439) & ChrW(&H443) & ChrW(&H441) & ChrW(&H442) & ChrW(&H443) & ChrW(&H43D)
    AddLatCyr "moyupa=" & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H439) & ChrW(&H443) & ChrW(&H43F) & ChrW(&H430)
    AddLatCyr "konveyer=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H432) & ChrW(&H435) & ChrW(&H439) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "pleyer=" & ChrW(&H43F) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H439) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "stayer=" & ChrW(&H441) & ChrW(&H442) & ChrW(&H430) & ChrW(&H439) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "foye*=" & ChrW(&H444) & ChrW(&H43E) & ChrW(&H439) & ChrW(&H435)
    AddLatCyr "kompyuter=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43C) & ChrW(&H43F) & ChrW(&H44C) & ChrW(&H44E) & ChrW(&H442) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "bulyon=" & ChrW(&H431) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "aprel=" & ChrW(&H430) & ChrW(&H43F) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "film*=" & ChrW(&H444) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C)
    AddLatCyr "rubl*=" & ChrW(&H440) & ChrW(&H443) & ChrW(&H431) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "prinsip=" & ChrW(&H43F) & ChrW(&H440) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43F)
    AddLatCyr "sitata=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H442) & ChrW(&H430) & ChrW(&H442) & ChrW(&H430)
    AddLatCyr "sivilizatsiya=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H432) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H437) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "sikl*=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H43B)
    AddLatCyr "siklon=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "sentr*=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H440)
    AddLatCyr "konsentr=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H440)
    AddLatCyr "spetsif=" & ChrW(&H441) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H446) & ChrW(&H438) & ChrW(&H444)
    AddLatCyr "spetsial=" & ChrW(&H441) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H446) & ChrW(&H438) & ChrW(&H430) & ChrW(&H43B)
    AddLatCyr "patsient=" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442)
    AddLatCyr "ofitsial=" & ChrW(&H43E) & ChrW(&H444) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H430) & ChrW(&H43B)
    AddLatCyr "stansiya=" & ChrW(&H441) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "instansiya=" & ChrW(&H438) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "distansiya=" & ChrW(&H434) & ChrW(&H438) & ChrW(&H441) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "kvitansiya=" & ChrW(&H43A) & ChrW(&H432) & ChrW(&H438) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "tendensiya=" & ChrW(&H442) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "potensial=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H430) & ChrW(&H43B)
    AddLatCyr "litsenziya=" & ChrW(&H43B) & ChrW(&H438) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H437) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "absorbsiya=" & ChrW(&H430) & ChrW(&H431) & ChrW(&H441) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H431) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "abssess=" & ChrW(&H430) & ChrW(&H431) & ChrW(&H441) & ChrW(&H446) & ChrW(&H435) & ChrW(&H441) & ChrW(&H441)
    AddLatCyr "adsorbsiya=" & ChrW(&H430) & ChrW(&H434) & ChrW(&H441) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H431) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "adyunkt=" & ChrW(&H430) & ChrW(&H434) & ChrW(&H44A) & ChrW(&H44E) & ChrW(&H43D) & ChrW(&H43A) & ChrW(&H442)
    AddLatCyr "aerofotosyomka=" & ChrW(&H430) & ChrW(&H44D) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H444) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H44A) & ChrW(&H451) & ChrW(&H43C) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "aerosyomka=" & ChrW(&H430) & ChrW(&H44D) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H44A) & ChrW(&H451) & ChrW(&H43C) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "agnostitsizm=" & ChrW(&H430) & ChrW(&H433) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "aksioma=" & ChrW(&H430) & ChrW(&H43A) & ChrW(&H441) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43C) & ChrW(&H430)
    AddLatCyr "aksiz=" & ChrW(&H430) & ChrW(&H43A) & ChrW(&H446) & ChrW(&H438) & ChrW(&H437)
    AddLatCyr "akvarel=" & ChrW(&H430) & ChrW(&H43A) & ChrW(&H432) & ChrW(&H430) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "albatros=" & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H431) & ChrW(&H430) & ChrW(&H442) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H441)
    AddLatCyr "albom=" & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H431) & ChrW(&H43E) & ChrW(&H43C)
    AddLatCyr "alfa-zarralar=" & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H444) & ChrW(&H430) & "-" & ChrW(&H437) & ChrW(&H430) & ChrW(&H440) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H440)
    AddLatCyr "alfa=" & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H444) & ChrW(&H430)
    AddLatCyr "alkogol=" & ChrW(&H430) & ChrW(&H43B) & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "alma-terapiya=" & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C) & ChrW(&H430) & "-" & ChrW(&H442) & ChrW(&H435) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43F) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "altchi=" & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H447) & ChrW(&H438)
    AddLatCyr "altimetr=" & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H435) & ChrW(&H442) & ChrW(&H440)
    AddLatCyr "alyans=" & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H44F) & ChrW(&H43D) & ChrW(&H441)
    AddLatCyr "amalgama=" & ChrW(&H430) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H433) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H430)
    AddLatCyr "anneksiya=" & ChrW(&H430) & ChrW(&H43D) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H441) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "apelsin=" & ChrW(&H430) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H441) & ChrW(&H438) & ChrW(&H43D)
    AddLatCyr "apopleksiya=" & ChrW(&H430) & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H43F) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H441) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "asfiksiya=" & ChrW(&H430) & ChrW(&H441) & ChrW(&H444) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H441) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "aviamodel=" & ChrW(&H430) & ChrW(&H432) & ChrW(&H438) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H434) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "avtomagistral=" & ChrW(&H430) & ChrW(&H432) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H433) & ChrW(&H438) & ChrW(&H441) & ChrW(&H442) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "balneolog=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H433)
    AddLatCyr "balneologik=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "balneologiya=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "balneoterapiya=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H435) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43F) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "balneotexnika=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H435) & ChrW(&H445) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "barelyef=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H444)
    AddLatCyr "barrel=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H440) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "baryer=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "bazalt=" & ChrW(&H431) & ChrW(&H430) & ChrW(&H437) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442)
    AddLatCyr "beletaj=" & ChrW(&H431) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H44D) & ChrW(&H442) & ChrW(&H430) & ChrW(&H436)
    AddLatCyr "belgiyalik=" & ChrW(&H431) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H433) & ChrW(&H438) & ChrW(&H44F) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "belting=" & ChrW(&H431) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H433)
    AddLatCyr "belveder=" & ChrW(&H431) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H432) & ChrW(&H435) & ChrW(&H434) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "bemeyor=" & ChrW(&H431) & ChrW(&H435) & ChrW(&H43C) & ChrW(&H435) & ChrW(&H44A) & ChrW(&H451) & ChrW(&H440)
    AddLatCyr "biofiltr=" & ChrW(&H431) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H444) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H440)
    AddLatCyr "biolyuminessensiya=" & ChrW(&H431) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44E) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H441) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "bolonya=" & ChrW(&H431) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H44F)
    AddLatCyr "bolshevizm=" & ChrW(&H431) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H448) & ChrW(&H435) & ChrW(&H432) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "brakonyerlik=" & ChrW(&H431) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "broneavtomobil=" & ChrW(&H431) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H430) & ChrW(&H432) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H431) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "buldenej=" & ChrW(&H431) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H434) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H436)
    AddLatCyr "buldog=" & ChrW(&H431) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H434) & ChrW(&H43E) & ChrW(&H433)
    AddLatCyr "buldozer=" & ChrW(&H431) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H434) & ChrW(&H43E) & ChrW(&H437) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "buldozerchi=" & ChrW(&H431) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H434) & ChrW(&H43E) & ChrW(&H437) & ChrW(&H435) & ChrW(&H440) & ChrW(&H447) & ChrW(&H438)
    AddLatCyr "burgutsifat=" & ChrW(&H431) & ChrW(&H443) & ChrW(&H440) & ChrW(&H433) & ChrW(&H443) & ChrW(&H442) & ChrW(&H441) & ChrW(&H438) & ChrW(&H444) & ChrW(&H430) & ChrW(&H442)
    AddLatCyr "burgutsimon=" & ChrW(&H431) & ChrW(&H443) & ChrW(&H440) & ChrW(&H433) & ChrW(&H443) & ChrW(&H442) & ChrW(&H441) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "byulleten=" & ChrW(&H431) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "chizel=" & ChrW(&H447) & ChrW(&H438) & ChrW(&H437) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "daraxtsimon=" & ChrW(&H434) & ChrW(&H430) & ChrW(&H440) & ChrW(&H430) & ChrW(&H445) & ChrW(&H442) & ChrW(&H441) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "delta=" & ChrW(&H434) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H430)
    AddLatCyr "detal=" & ChrW(&H434) & ChrW(&H435) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "diafilm=" & ChrW(&H434) & ChrW(&H438) & ChrW(&H430) & ChrW(&H444) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C)
    AddLatCyr "differensiatsiya=" & ChrW(&H434) & ChrW(&H438) & ChrW(&H444) & ChrW(&H444) & ChrW(&H435) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "distansion=" & ChrW(&H434) & ChrW(&H438) & ChrW(&H441) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "drel=" & ChrW(&H434) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "emal=" & ChrW(&H44D) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "empiriokrititsizm=" & ChrW(&H44D) & ChrW(&H43C) & ChrW(&H43F) & ChrW(&H438) & ChrW(&H440) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43A) & ChrW(&H440) & ChrW(&H438) & ChrW(&H442) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "endshpil=" & ChrW(&H44D) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H448) & ChrW(&H43F) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "ensefalit=" & ChrW(&H44D) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H435) & ChrW(&H444) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H442)
    AddLatCyr "ensefalogramma=" & ChrW(&H44D) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H435) & ChrW(&H444) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43C) & ChrW(&H430)
    AddLatCyr "ensiklopedik=" & ChrW(&H44D) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H434) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "ensiklopedist=" & ChrW(&H44D) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H434) & ChrW(&H438) & ChrW(&H441) & ChrW(&H442)
    AddLatCyr "ensiklopediya=" & ChrW(&H44D) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H434) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "ensiklopediyachi=" & ChrW(&H44D) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H434) & ChrW(&H438) & ChrW(&H44F) & ChrW(&H447) & ChrW(&H438)
    AddLatCyr "esminets=" & ChrW(&H44D) & ChrW(&H441) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H446)
    AddLatCyr "feldyeger=" & ChrW(&H444) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H434) & ChrW(&H44A) & ChrW(&H435) & ChrW(&H433) & ChrW(&H435) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "felyeton=" & ChrW(&H444) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "felyetonchi=" & ChrW(&H444) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H447) & ChrW(&H438)
    AddLatCyr "filmoy=" & ChrW(&H444) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H439)
    AddLatCyr "fitonsid=" & ChrW(&H444) & ChrW(&H438) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H434)
    AddLatCyr "fizkultura=" & ChrW(&H444) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H443) & ChrW(&H440) & ChrW(&H430)
    AddLatCyr "fizkulturachi=" & ChrW(&H444) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H443) & ChrW(&H440) & ChrW(&H430) & ChrW(&H447) & ChrW(&H438)
    AddLatCyr "fleksiya=" & ChrW(&H444) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H441) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "folga=" & ChrW(&H444) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H433) & ChrW(&H430)
    AddLatCyr "fonar=" & ChrW(&H444) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "fortepyano=" & ChrW(&H444) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43F) & ChrW(&H44C) & ChrW(&H44F) & ChrW(&H43D) & ChrW(&H43E)
    AddLatCyr "fotosyomka=" & ChrW(&H444) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H44A) & ChrW(&H451) & ChrW(&H43C) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "fransiya=" & ChrW(&H444) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "frits=" & ChrW(&H444) & ChrW(&H440) & ChrW(&H438) & ChrW(&H446)
    AddLatCyr "gallitsizm=" & ChrW(&H433) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "gallyutsinatsiya=" & ChrW(&H433) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H44E) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "galvanik=" & ChrW(&H433) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H432) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "galvanizatsiya=" & ChrW(&H433) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H432) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H437) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "galvanometr=" & ChrW(&H433) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H432) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H43C) & ChrW(&H435) & ChrW(&H442) & ChrW(&H440)
    AddLatCyr "gavan=" & ChrW(&H433) & ChrW(&H430) & ChrW(&H432) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "gelmint=" & ChrW(&H433) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H442)
    AddLatCyr "gelmintologiya=" & ChrW(&H433) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "gelmintoz=" & ChrW(&H433) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H437)
    AddLatCyr "geraldika=" & ChrW(&H433) & ChrW(&H435) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H434) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "gersog=" & ChrW(&H433) & ChrW(&H435) & ChrW(&H440) & ChrW(&H446) & ChrW(&H43E) & ChrW(&H433)
    AddLatCyr "gidropodyomnik=" & ChrW(&H433) & ChrW(&H438) & ChrW(&H434) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H434) & ChrW(&H44A) & ChrW(&H451) & ChrW(&H43C) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "giposulfit=" & ChrW(&H433) & ChrW(&H438) & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H444) & ChrW(&H438) & ChrW(&H442)
    AddLatCyr "golf=" & ChrW(&H433) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H444)
    AddLatCyr "gorelyef=" & ChrW(&H433) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H444)
    AddLatCyr "guash=" & ChrW(&H433) & ChrW(&H443) & ChrW(&H430) & ChrW(&H448) & ChrW(&H44C)
    AddLatCyr "instarsiya=" & ChrW(&H438) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H442) & ChrW(&H430) & ChrW(&H440) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "interyer=" & ChrW(&H438) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H435) & ChrW(&H440) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "inventar=" & ChrW(&H438) & ChrW(&H43D) & ChrW(&H432) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "itsigek=" & ChrW(&H438) & ChrW(&H442) & ChrW(&H441) & ChrW(&H438) & ChrW(&H433) & ChrW(&H435) & ChrW(&H43A)
    AddLatCyr "jannatsimon=" & ChrW(&H436) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H442) & ChrW(&H441) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "jenshen=" & ChrW(&H436) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H448) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "kafel=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H444) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "kalka=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "kalkalamoq=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H49B)
    AddLatCyr "kalkulyatsiya=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44F) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "kanserogen=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H435) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H435) & ChrW(&H43D)
    AddLatCyr "kapelmeyster=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C) & ChrW(&H435) & ChrW(&H439) & ChrW(&H441) & ChrW(&H442) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "kapsyul=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43F) & ChrW(&H441) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "kartech=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H440) & ChrW(&H442) & ChrW(&H435) & ChrW(&H447) & ChrW(&H44C)
    AddLatCyr "karyer=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "kastryul=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H441) & ChrW(&H442) & ChrW(&H440) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "kastryulka=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H441) & ChrW(&H442) & ChrW(&H440) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "kaxeksiya=" & ChrW(&H43A) & ChrW(&H430) & ChrW(&H445) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H441) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "kinokonsert=" & ChrW(&H43A) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H435) & ChrW(&H440) & ChrW(&H442)
    AddLatCyr "kinossenariy=" & ChrW(&H43A) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H440) & ChrW(&H438) & ChrW(&H439)
    AddLatCyr "klassitsizm=" & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H441) & ChrW(&H441) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "knyaz=" & ChrW(&H43A) & ChrW(&H43D) & ChrW(&H44F) & ChrW(&H437) & ChrW(&H44C)
    AddLatCyr "kokteyl=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43A) & ChrW(&H442) & ChrW(&H435) & ChrW(&H439) & ChrW(&H43B) & ChrW(&H44C)
End Sub

Private Sub InitLatCyr3()
    AddLatCyr "kolonsifra=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H444) & ChrW(&H440) & ChrW(&H430)
    AddLatCyr "kommersiya=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43C) & ChrW(&H43C) & ChrW(&H435) & ChrW(&H440) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "konki=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H43A) & ChrW(&H438)
    AddLatCyr "konkichi=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H43A) & ChrW(&H438) & ChrW(&H447) & ChrW(&H438)
    AddLatCyr "konsessiya=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H435) & ChrW(&H441) & ChrW(&H441) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "konsorsium=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H446) & ChrW(&H438) & ChrW(&H443) & ChrW(&H43C)
    AddLatCyr "konsultant=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H442)
    AddLatCyr "konsultativ=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H430) & ChrW(&H442) & ChrW(&H438) & ChrW(&H432)
    AddLatCyr "konsultatsiya=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "kontrol=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "konyak=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H44F) & ChrW(&H43A)
    AddLatCyr "korol=" & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "krovat=" & ChrW(&H43A) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H432) & ChrW(&H430) & ChrW(&H442) & ChrW(&H44C)
    AddLatCyr "kulminatsion=" & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "kulminatsiya=" & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "kulturizm=" & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H443) & ChrW(&H440) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "kyat=" & ChrW(&H43A) & ChrW(&H44C) & ChrW(&H44F) & ChrW(&H442)
    AddLatCyr "lanset=" & ChrW(&H43B) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H435) & ChrW(&H442)
    AddLatCyr "latun=" & ChrW(&H43B) & ChrW(&H430) & ChrW(&H442) & ChrW(&H443) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "losyon=" & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H44C) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "lourensiy=" & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H443) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H438) & ChrW(&H439)
    AddLatCyr "marselyeza=" & ChrW(&H43C) & ChrW(&H430) & ChrW(&H440) & ChrW(&H441) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H437) & ChrW(&H430)
    AddLatCyr "matritsa=" & ChrW(&H43C) & ChrW(&H430) & ChrW(&H442) & ChrW(&H440) & ChrW(&H438) & ChrW(&H446) & ChrW(&H430)
    AddLatCyr "mayub=" & ChrW(&H43C) & ChrW(&H430) & ChrW(&H44A) & ChrW(&H44E) & ChrW(&H431)
    AddLatCyr "mayus=" & ChrW(&H43C) & ChrW(&H430) & ChrW(&H44A) & ChrW(&H44E) & ChrW(&H441)
    AddLatCyr "mayuslanmoq=" & ChrW(&H43C) & ChrW(&H430) & ChrW(&H44A) & ChrW(&H44E) & ChrW(&H441) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H49B)
    AddLatCyr "mayuslik=" & ChrW(&H43C) & ChrW(&H430) & ChrW(&H44A) & ChrW(&H44E) & ChrW(&H441) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "mayusona=" & ChrW(&H43C) & ChrW(&H430) & ChrW(&H44A) & ChrW(&H44E) & ChrW(&H441) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H430)
    AddLatCyr "meditsina=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H434) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H430)
    AddLatCyr "mehnatsevar=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H4B3) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H442) & ChrW(&H441) & ChrW(&H435) & ChrW(&H432) & ChrW(&H430) & ChrW(&H440)
    AddLatCyr "mehnatsevarlik=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H4B3) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H442) & ChrW(&H441) & ChrW(&H435) & ChrW(&H432) & ChrW(&H430) & ChrW(&H440) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "melxior=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H445) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H440)
    AddLatCyr "mexanitsizm=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H445) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "meyor=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H44A) & ChrW(&H451) & ChrW(&H440)
    AddLatCyr "meyoriy=" & ChrW(&H43C) & ChrW(&H435) & ChrW(&H44A) & ChrW(&H451) & ChrW(&H440) & ChrW(&H438) & ChrW(&H439)
    AddLatCyr "mikroinsult=" & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442)
    AddLatCyr "mistitsizm=" & ChrW(&H43C) & ChrW(&H438) & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "mizanssena=" & ChrW(&H43C) & ChrW(&H438) & ChrW(&H437) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H430)
    AddLatCyr "modelyer=" & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H434) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "molbert=" & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H431) & ChrW(&H435) & ChrW(&H440) & ChrW(&H442)
    AddLatCyr "monastir=" & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "motel=" & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "multimediya=" & ChrW(&H43C) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H435) & ChrW(&H434) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "multimillioner=" & ChrW(&H43C) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "multiplikator=" & ChrW(&H43C) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43F) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H440)
    AddLatCyr "multiplikatsion=" & ChrW(&H43C) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43F) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "multiplikatsiya=" & ChrW(&H43C) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43F) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "nenets=" & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H446)
    AddLatCyr "nenetslar=" & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H446) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H440)
    AddLatCyr "nimpalto=" & ChrW(&H43D) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H43F) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H43E)
    AddLatCyr "oblast=" & ChrW(&H43E) & ChrW(&H431) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H441) & ChrW(&H442) & ChrW(&H44C)
    AddLatCyr "ochered=" & ChrW(&H43E) & ChrW(&H447) & ChrW(&H435) & ChrW(&H440) & ChrW(&H435) & ChrW(&H434) & ChrW(&H44C)
    AddLatCyr "oftalmologiya=" & ChrW(&H43E) & ChrW(&H444) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "okkultizm=" & ChrW(&H43E) & ChrW(&H43A) & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "opsion=" & ChrW(&H43E) & ChrW(&H43F) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "ordinarets=" & ChrW(&H43E) & ChrW(&H440) & ChrW(&H434) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H440) & ChrW(&H435) & ChrW(&H446)
    AddLatCyr "otel=" & ChrW(&H43E) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "palma=" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C) & ChrW(&H430)
    AddLatCyr "palmazor=" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H437) & ChrW(&H43E) & ChrW(&H440)
    AddLatCyr "palpatsiya=" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43F) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "pechat=" & ChrW(&H43F) & ChrW(&H435) & ChrW(&H447) & ChrW(&H430) & ChrW(&H442) & ChrW(&H44C)
    AddLatCyr "pechenye=" & ChrW(&H43F) & ChrW(&H435) & ChrW(&H447) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H435)
    AddLatCyr "penalti=" & ChrW(&H43F) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438)
    AddLatCyr "petlitsa=" & ChrW(&H43F) & ChrW(&H435) & ChrW(&H442) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H446) & ChrW(&H430)
    AddLatCyr "pinset=" & ChrW(&H43F) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H435) & ChrW(&H442)
    AddLatCyr "pitssa=" & ChrW(&H43F) & ChrW(&H438) & ChrW(&H446) & ChrW(&H446) & ChrW(&H430)
    AddLatCyr "platskart=" & ChrW(&H43F) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H446) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H440) & ChrW(&H442)
    AddLatCyr "plebissit=" & ChrW(&H43F) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H431) & ChrW(&H438) & ChrW(&H441) & ChrW(&H446) & ChrW(&H438) & ChrW(&H442)
    AddLatCyr "porsillamoq=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H441) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H49B)
    AddLatCyr "porsilloq=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H441) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H49B)
    AddLatCyr "porsin=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H441) & ChrW(&H438) & ChrW(&H43D)
    AddLatCyr "portsigar=" & ChrW(&H43F) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H442) & ChrW(&H441) & ChrW(&H438) & ChrW(&H433) & ChrW(&H430) & ChrW(&H440)
    AddLatCyr "premyer-ministr=" & ChrW(&H43F) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43C) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440) & "-" & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H441) & ChrW(&H442) & ChrW(&H440)
    AddLatCyr "press-papye=" & ChrW(&H43F) & ChrW(&H440) & ChrW(&H435) & ChrW(&H441) & ChrW(&H441) & "-" & ChrW(&H43F) & ChrW(&H430) & ChrW(&H43F) & ChrW(&H44C) & ChrW(&H435)
    AddLatCyr "press-sekretar=" & ChrW(&H43F) & ChrW(&H440) & ChrW(&H435) & ChrW(&H441) & ChrW(&H441) & "-" & ChrW(&H441) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H440) & ChrW(&H435) & ChrW(&H442) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "profil=" & ChrW(&H43F) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H444) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "provinsializm=" & ChrW(&H43F) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H432) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "provinsiya=" & ChrW(&H43F) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H432) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "pulmonologiya=" & ChrW(&H43F) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "pulpit=" & ChrW(&H43F) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43F) & ChrW(&H438) & ChrW(&H442)
    AddLatCyr "pult=" & ChrW(&H43F) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442)
    AddLatCyr "pulverizator=" & ChrW(&H43F) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H432) & ChrW(&H435) & ChrW(&H440) & ChrW(&H438) & ChrW(&H437) & ChrW(&H430) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H440)
    AddLatCyr "qaynegachi=" & ChrW(&H49B) & ChrW(&H430) & ChrW(&H439) & ChrW(&H43D) & ChrW(&H44D) & ChrW(&H433) & ChrW(&H430) & ChrW(&H447) & ChrW(&H438)
    AddLatCyr "qimmatsiramoq=" & ChrW(&H49B) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H442) & ChrW(&H441) & ChrW(&H438) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H49B)
    AddLatCyr "radiospektakl=" & ChrW(&H440) & ChrW(&H430) & ChrW(&H434) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "rantye=" & ChrW(&H440) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H44C) & ChrW(&H435)
    AddLatCyr "raya-andisha=" & ChrW(&H440) & ChrW(&H430) & ChrW(&H44A) & ChrW(&H44F) & "-" & ChrW(&H430) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H438) & ChrW(&H448) & ChrW(&H430)
    AddLatCyr "raya=" & ChrW(&H440) & ChrW(&H430) & ChrW(&H44A) & ChrW(&H44F)
    AddLatCyr "reket=" & ChrW(&H440) & ChrW(&H44D) & ChrW(&H43A) & ChrW(&H435) & ChrW(&H442)
    AddLatCyr "rekognossirovka=" & ChrW(&H440) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H446) & ChrW(&H438) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H432) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "relsli=" & ChrW(&H440) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H441) & ChrW(&H43B) & ChrW(&H438)
    AddLatCyr "relssiz=" & ChrW(&H440) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H441) & ChrW(&H441) & ChrW(&H438) & ChrW(&H437)
    AddLatCyr "retush=" & ChrW(&H440) & ChrW(&H435) & ChrW(&H442) & ChrW(&H443) & ChrW(&H448) & ChrW(&H44C)
    AddLatCyr "revalvatsiya=" & ChrW(&H440) & ChrW(&H435) & ChrW(&H432) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H432) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "revolver=" & ChrW(&H440) & ChrW(&H435) & ChrW(&H432) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H432) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "riyel=" & ChrW(&H440) & ChrW(&H438) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "rubilnik=" & ChrW(&H440) & ChrW(&H443) & ChrW(&H431) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "salto=" & ChrW(&H441) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H43E)
    AddLatCyr "sanga=" & ChrW(&H446) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H433) & ChrW(&H430)
    AddLatCyr "sapfa=" & ChrW(&H446) & ChrW(&H430) & ChrW(&H43F) & ChrW(&H444) & ChrW(&H430)
    AddLatCyr "sedra=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H434) & ChrW(&H440) & ChrW(&H430)
    AddLatCyr "sekretar=" & ChrW(&H441) & ChrW(&H435) & ChrW(&H43A) & ChrW(&H440) & ChrW(&H435) & ChrW(&H442) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "selderey=" & ChrW(&H441) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H434) & ChrW(&H435) & ChrW(&H440) & ChrW(&H435) & ChrW(&H439)
    AddLatCyr "sent=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H442)
    AddLatCyr "senyor=" & ChrW(&H441) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H43E) & ChrW(&H440)
    AddLatCyr "senyora=" & ChrW(&H441) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H430)
    AddLatCyr "seytnot=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H439) & ChrW(&H442) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H442)
    AddLatCyr "seziy=" & ChrW(&H446) & ChrW(&H435) & ChrW(&H437) & ChrW(&H438) & ChrW(&H439)
    AddLatCyr "shagren=" & ChrW(&H448) & ChrW(&H430) & ChrW(&H433) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "shampun=" & ChrW(&H448) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43F) & ChrW(&H443) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "sherst=" & ChrW(&H448) & ChrW(&H435) & ChrW(&H440) & ChrW(&H441) & ChrW(&H442) & ChrW(&H44C)
    AddLatCyr "shifonyer=" & ChrW(&H448) & ChrW(&H438) & ChrW(&H444) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "shmutstitul=" & ChrW(&H448) & ChrW(&H43C) & ChrW(&H443) & ChrW(&H446) & ChrW(&H442) & ChrW(&H438) & ChrW(&H442) & ChrW(&H443) & ChrW(&H43B)
    AddLatCyr "shprits=" & ChrW(&H448) & ChrW(&H43F) & ChrW(&H440) & ChrW(&H438) & ChrW(&H446)
    AddLatCyr "shtangensirkul=" & ChrW(&H448) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H433) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H440) & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "sian=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H430) & ChrW(&H43D)
    AddLatCyr "sianli=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H43B) & ChrW(&H438)
    AddLatCyr "sinkalamoq=" & ChrW(&H441) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H49B)
    AddLatCyr "sinkograf=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H440) & ChrW(&H430) & ChrW(&H444)
    AddLatCyr "sinkografiya=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H440) & ChrW(&H430) & ChrW(&H444) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "siren=" & ChrW(&H441) & ChrW(&H438) & ChrW(&H440) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H44C)
    AddLatCyr "sirkchi=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H440) & ChrW(&H43A) & ChrW(&H447) & ChrW(&H438)
    AddLatCyr "sirkoniy=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H440) & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H439)
    AddLatCyr "sirkul=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H440) & ChrW(&H43A) & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "sirroz=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H440) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H437)
    AddLatCyr "sistit=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H442)
    AddLatCyr "sitodiagnostika=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H434) & ChrW(&H438) & ChrW(&H430) & ChrW(&H433) & ChrW(&H43D) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "sitokimyo=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43A) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H451)
    AddLatCyr "sitoliz=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H437)
    AddLatCyr "sitologiya=" & ChrW(&H446) & ChrW(&H438) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H433) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "skalpel=" & ChrW(&H441) & ChrW(&H43A) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43F) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "skeptitsizm=" & ChrW(&H441) & ChrW(&H43A) & ChrW(&H435) & ChrW(&H43F) & ChrW(&H442) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "slanets=" & ChrW(&H441) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H446)
    AddLatCyr "slesar=" & ChrW(&H441) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H441) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C)
    AddLatCyr "sobol=" & ChrW(&H441) & ChrW(&H43E) & ChrW(&H431) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "sokol=" & ChrW(&H446) & ChrW(&H43E) & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "ssenarist=" & ChrW(&H441) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H440) & ChrW(&H438) & ChrW(&H441) & ChrW(&H442)
    AddLatCyr "statya=" & ChrW(&H441) & ChrW(&H442) & ChrW(&H430) & ChrW(&H442) & ChrW(&H44C) & ChrW(&H44F)
    AddLatCyr "stelka=" & ChrW(&H441) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "stil=" & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "stoitsizm=" & ChrW(&H441) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H438) & ChrW(&H446) & ChrW(&H438) & ChrW(&H437) & ChrW(&H43C)
    AddLatCyr "stronsiy=" & ChrW(&H441) & ChrW(&H442) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H439)
    AddLatCyr "sunami=" & ChrW(&H446) & ChrW(&H443) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H438)
    AddLatCyr "sutemizuvchilar=" & ChrW(&H441) & ChrW(&H443) & ChrW(&H442) & ChrW(&H44D) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H437) & ChrW(&H443) & ChrW(&H432) & ChrW(&H447) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H440)
    AddLatCyr "tablitsa=" & ChrW(&H442) & ChrW(&H430) & ChrW(&H431) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H446) & ChrW(&H430)
    AddLatCyr "talk=" & ChrW(&H442) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43A)
    AddLatCyr "telefilm=" & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H444) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C)
    AddLatCyr "teleinssenirovka=" & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H435) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H446) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H432) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "tendensioz=" & ChrW(&H442) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H437)
    AddLatCyr "tendensiozlik=" & ChrW(&H442) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H434) & ChrW(&H435) & ChrW(&H43D) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H437) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "teplitsa=" & ChrW(&H442) & ChrW(&H435) & ChrW(&H43F) & ChrW(&H43B) & ChrW(&H438) & ChrW(&H446) & ChrW(&H430)
    AddLatCyr "terset=" & ChrW(&H442) & ChrW(&H435) & ChrW(&H440) & ChrW(&H446) & ChrW(&H435) & ChrW(&H442)
    AddLatCyr "tersiya=" & ChrW(&H442) & ChrW(&H435) & ChrW(&H440) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "tigel=" & ChrW(&H442) & ChrW(&H438) & ChrW(&H433) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "tonnel=" & ChrW(&H442) & ChrW(&H43E) & ChrW(&H43D) & ChrW(&H43D) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "transkripsion=" & ChrW(&H442) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H43A) & ChrW(&H440) & ChrW(&H438) & ChrW(&H43F) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "transkripsiya=" & ChrW(&H442) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H43A) & ChrW(&H440) & ChrW(&H438) & ChrW(&H43F) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "transkripsiyalamoq=" & ChrW(&H442) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H441) & ChrW(&H43A) & ChrW(&H440) & ChrW(&H438) & ChrW(&H43F) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F) & ChrW(&H43B) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H49B)
    AddLatCyr "tyul=" & ChrW(&H442) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "ultimatum=" & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H430) & ChrW(&H442) & ChrW(&H443) & ChrW(&H43C)
    AddLatCyr "ultrabinafsha=" & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H440) & ChrW(&H430) & ChrW(&H431) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H430) & ChrW(&H444) & ChrW(&H448) & ChrW(&H430)
    AddLatCyr "ultramikroskop=" & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H43A) & ChrW(&H43E) & ChrW(&H43F)
    AddLatCyr "ultraqisqa=" & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H440) & ChrW(&H430) & ChrW(&H49B) & ChrW(&H438) & ChrW(&H441) & ChrW(&H49B) & ChrW(&H430)
    AddLatCyr "ultratovush=" & ChrW(&H443) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H440) & ChrW(&H430) & ChrW(&H442) & ChrW(&H43E) & ChrW(&H432) & ChrW(&H443) & ChrW(&H448)
    AddLatCyr "util=" & ChrW(&H443) & ChrW(&H442) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "vaksina=" & ChrW(&H432) & ChrW(&H430) & ChrW(&H43A) & ChrW(&H446) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H430)
    AddLatCyr "valvatsiya=" & ChrW(&H432) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H432) & ChrW(&H430) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "vanil=" & ChrW(&H432) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "varyete=" & ChrW(&H432) & ChrW(&H430) & ChrW(&H440) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H442) & ChrW(&H435)
    AddLatCyr "vedomost=" & ChrW(&H432) & ChrW(&H435) & ChrW(&H434) & ChrW(&H43E) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H442) & ChrW(&H44C)
    AddLatCyr "verf=" & ChrW(&H432) & ChrW(&H435) & ChrW(&H440) & ChrW(&H444) & ChrW(&H44C)
    AddLatCyr "vernyer=" & ChrW(&H432) & ChrW(&H435) & ChrW(&H440) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H440)
    AddLatCyr "vestibyul=" & ChrW(&H432) & ChrW(&H435) & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H431) & ChrW(&H44E) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "videofilm=" & ChrW(&H432) & ChrW(&H438) & ChrW(&H434) & ChrW(&H435) & ChrW(&H43E) & ChrW(&H444) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43C)
    AddLatCyr "viklyuchatel=" & ChrW(&H432) & ChrW(&H438) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H44E) & ChrW(&H447) & ChrW(&H430) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "vintsimon=" & ChrW(&H432) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H442) & ChrW(&H441) & ChrW(&H438) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H43D)
    AddLatCyr "vinyetka=" & ChrW(&H432) & ChrW(&H438) & ChrW(&H43D) & ChrW(&H44C) & ChrW(&H435) & ChrW(&H442) & ChrW(&H43A) & ChrW(&H430)
    AddLatCyr "vklyuchatel=" & ChrW(&H432) & ChrW(&H43A) & ChrW(&H43B) & ChrW(&H44E) & ChrW(&H447) & ChrW(&H430) & ChrW(&H442) & ChrW(&H435) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "vodevil=" & ChrW(&H432) & ChrW(&H43E) & ChrW(&H434) & ChrW(&H435) & ChrW(&H432) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "volost=" & ChrW(&H432) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H442) & ChrW(&H44C)
    AddLatCyr "voltli=" & ChrW(&H432) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H442) & ChrW(&H43B) & ChrW(&H438)
    AddLatCyr "xemosorbsiya=" & ChrW(&H445) & ChrW(&H435) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H441) & ChrW(&H43E) & ChrW(&H440) & ChrW(&H431) & ChrW(&H446) & ChrW(&H438) & ChrW(&H44F)
    AddLatCyr "xolodilnik=" & ChrW(&H445) & ChrW(&H43E) & ChrW(&H43B) & ChrW(&H43E) & ChrW(&H434) & ChrW(&H438) & ChrW(&H43B) & ChrW(&H44C) & ChrW(&H43D) & ChrW(&H438) & ChrW(&H43A)
    AddLatCyr "xrustal=" & ChrW(&H445) & ChrW(&H440) & ChrW(&H443) & ChrW(&H441) & ChrW(&H442) & ChrW(&H430) & ChrW(&H43B) & ChrW(&H44C)
    AddLatCyr "yodli=" & ChrW(&H439) & ChrW(&H43E) & ChrW(&H434) & ChrW(&H43B) & ChrW(&H438)
    AddLatCyr "yotsiramoq=" & ChrW(&H451) & ChrW(&H442) & ChrW(&H441) & ChrW(&H438) & ChrW(&H440) & ChrW(&H430) & ChrW(&H43C) & ChrW(&H43E) & ChrW(&H49B)
    AddLatCyr "yuan=" & ChrW(&H44E) & ChrW(&H430) & ChrW(&H43D) & ChrW(&H44C)
End Sub

'==========================================================================
'  EXCEL VA POWERPOINT UCHUN UMUMIY QISM
'  (yuqoridagi konvertor Word versiyasi bilan aynan bir xil;
'   bu yerda faqat ilovaga bog'liq bo'lmagan makroslar va lenta tugmalari)
'==========================================================================

' Matnga o'zgarishlarni qo'llaydi (mMode avvaldan o'rnatiladi); n - o'zgarishlar soni
Private Function ConvertString(ByVal t As String, ByRef n As Long) As String
    Dim n1 As Long, n2 As Long, o As String
    If mMode = 4 Or mMode = 5 Then          ' 4: kirill + eski lotin -> yangi; 5: kirill + yangi -> eski lotin
        Dim m As Long
        m = mMode
        mMode = 1: o = ConvertString1(t, n1)
        mMode = IIf(m = 4, 0, 3): o = ConvertString1(o, n2)
        mMode = m
        n = n1 + n2
        ConvertString = o
    Else
        ConvertString = ConvertString1(t, n)
    End If
End Function

Private Function ConvertString1(ByVal t As String, ByRef n As Long) As String
    Dim k As Long, o As String
    n = ComputeEdits(t)
    o = t
    For k = n To 1 Step -1
        o = Left$(o, mPos(k)) & mTo(k) & Mid$(o, mPos(k) + mLen(k) + 1)
    Next k
    ConvertString1 = o
End Function

' TextRange2 (shakl, jadval katakchasi, slayd matni) ichidagi matnni formatini saqlagan holda o'giradi
Private Function ConvertTR(ByVal tr As Object, Optional ByVal whole As Boolean = False) As Long
    Dim m As Long
    If mMode = 4 Or mMode = 5 Then
        m = mMode
        mMode = 1: ConvertTR = ConvertTR1(tr, whole)
        mMode = IIf(m = 4, 0, 3): ConvertTR = ConvertTR + ConvertTR1(tr, whole)
        mMode = m
    Else
        ConvertTR = ConvertTR1(tr, whole)
    End If
End Function

Private Function ConvertTR1(ByVal tr As Object, ByVal whole As Boolean) As Long
    Dim i As Long, k As Long, p As Object, t As String, n As Long, cnt As Long, ch As Object, np As Long
    On Error Resume Next
    If whole Then
        np = 1
    Else
        np = tr.Paragraphs.Count
    End If
    For i = np To 1 Step -1
        If whole Then Set p = tr Else Set p = tr.Paragraphs(i)
        t = p.Text
        If Not IsMonoFont(CStr(p.Font.Name)) Then
            n = ComputeEdits(t)
            For k = n To 1 Step -1
                Set ch = p.Characters(mPos(k) + 1, mLen(k))
                If ch.Text = mFrom(k) Then
                    ch.Text = mTo(k)
                    cnt = cnt + 1
                Else
                    mSkipped = mSkipped + 1
                End If
            Next k
        End If
    Next i
    ConvertTR1 = cnt
End Function

' Shakl: guruh, jadval yoki matnli shakl
Private Function ConvertShape(ByVal shp As Object) As Long
    Dim s As Object, r As Long, c As Long, cnt As Long, ht As Boolean, tp As Long, hasT As Boolean
    On Error Resume Next
    tp = -1
    tp = shp.Type
    If tp = 6 Then                                    ' msoGroup
        For Each s In shp.GroupItems
            cnt = cnt + ConvertShape(s)
        Next s
        ConvertShape = cnt
        Exit Function
    End If
    ht = False
    ht = shp.HasTable
    If ht Then
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                cnt = cnt + ConvertTR(shp.Table.Cell(r, c).Shape.TextFrame2.TextRange)
            Next c
        Next r
        ConvertShape = cnt
        Exit Function
    End If
    hasT = False
    hasT = (shp.TextFrame2.HasText <> 0)
    If hasT Then cnt = ConvertTR(shp.TextFrame2.TextRange)
    ConvertShape = cnt
End Function

Private Sub ResultMsg(ByVal total As Long, ByVal mode As Long, ByVal extra As String)
    MsgBox total & IIf(mode = 1, " ta so'z kirilldan o'girildi.", IIf(mode = 2, " ta so'z kirillga o'girildi.", _
        IIf(mode = 3, " ta so'z eski lotinga o'girildi.", IIf(mode = 4, " ta joy yangi alifboga o'girildi.", IIf(mode = 5, " ta joy eski lotinga o'girildi.", " ta harf o'girildi."))))) & _
        IIf(mSkipped > 0, vbCrLf & mSkipped & " ta joy o'tkazib yuborildi.", "") & _
        IIf(Len(extra) > 0, vbCrLf & extra, ""), vbInformation, AppNomi()
End Sub

' Matnni vaqtinchalik faylga yozib, Bloknotda ochadi
Private Sub ShowText(ByVal t As String)
    Dim p As String
    p = Environ$("TEMP") & "\Ozbekca-royxat.txt"
    WriteUtf8File p, Replace(t, vbCr, vbCrLf)
    Shell "notepad.exe """ & p & """", vbNormalFocus
End Sub

' Matndan birinchi so'zni (chetdagi tinish belgilarisiz) ajratadi
Private Function FirstWord(ByVal s As String) As String
    Dim i As Long, st As Long
    s = Replace(Replace(s, vbCr, " "), vbLf, " ")
    s = Trim$(s)
    i = InStr(s, " ")
    If i > 0 Then s = Left$(s, i - 1)
    Do While Len(s) > 0
        If IsLetterCh(Left$(s, 1)) Or IsDigitCh(Left$(s, 1)) Then Exit Do
        s = Mid$(s, 2)
    Loop
    Do While Len(s) > 0
        If IsLetterCh(Right$(s, 1)) Or IsDigitCh(Right$(s, 1)) Then Exit Do
        s = Left$(s, Len(s) - 1)
    Loop
    FirstWord = s
End Function

'------------------------------ makroslar ---------------------------------
Public Sub YangiAlifbo_Hujjat()
    RunConvert False
End Sub

Public Sub YangiAlifbo_Tanlangan()
    RunConvert True
End Sub

' Yangi alifboga: kirill ham, eski lotin ham
Public Sub YangiAlifbo_Lotin()
    RunConvert HostHasSelection(), 4
End Sub

Public Sub YangiAlifbo_Kirill()
    RunConvert False, 1
End Sub

Public Sub YangiAlifbo_KirillTanlangan()
    RunConvert True, 1
End Sub

Public Sub YangiAlifbo_KirillAvto()
    RunConvert HostHasSelection(), 4
End Sub

Public Sub YangiAlifbo_EskiLotin()
    RunConvert False, 3
End Sub

Public Sub YangiAlifbo_EskiLotinTanlangan()
    RunConvert True, 3
End Sub

Public Sub YangiAlifbo_EskiLotinAvto()
    RunConvert HostHasSelection(), 5
End Sub

Public Sub YangiAlifbo_LotinKirill()
    RunConvert False, 2
End Sub

Public Sub YangiAlifbo_LotinKirillTanlangan()
    RunConvert True, 2
End Sub

Public Sub YangiAlifbo_LotinKirillAvto()
    RunConvert HostHasSelection(), 2
End Sub

' Belgilangan so'zni (yoki qo'lda yozilganini) istisno so'zlarga qo'shadi
Public Sub YangiAlifbo_IstisnoQoshish()
    Dim w As String
    InitAll
    w = FirstWord(HostSelectedText())
    If IsNewLetterWord(w) Then w = RevertWord(w)
    w = InputBox("Istisno so'zlarga qo'shiladigan so'z:" & vbCrLf & vbCrLf & _
                 "Shablonlar:" & vbCrLf & _
                 "  phish*  - shu bilan boshlanadigan so'zlar" & vbCrLf & _
                 "  SHA#    - keyin raqam yoki '-' (SHA-256)", AppNomi() & " - so'z qo'shish", w)
    If Len(Trim$(w)) > 0 Then UMsg AddUserWord(w)
    mInit = False
End Sub

Private Function IsNewLetterWord(ByVal w As String) As Boolean
    Dim i As Long
    For i = 1 To Len(w)
        If IsNewLetter(Mid$(w, i, 1)) Then IsNewLetterWord = True: Exit Function
    Next i
End Function

' Kirill so'zni kirill istisnolariga qo'shadi
Public Sub YangiAlifbo_KirillIstisnoQoshish()
    Dim w As String, lat As String, users As Collection, offs As Collection, entry As String, i As Long
    InitAll
    w = FirstWord(HostSelectedText())
    If Not HasCyr(w) Then w = ""
    w = Trim$(InputBox("O'girilmaydigan (yoki maxsus shakldagi) kirill so'z:", AppNomi() & " - kirill istisnosi", w))
    If Len(w) = 0 Then Exit Sub
    For i = 1 To Len(w)
        If Not IsCyrCh(Mid$(w, i, 1)) Then UMsg "Faqat bitta kirill so'z kiriting.", 48: Exit Sub
    Next i
    lat = InputBox("Lotin shaklini yozing (masalan: Microsoft)." & vbCrLf & _
                   "Bo'sh qoldirsangiz - bu so'z o'girilmaydi." & vbCrLf & _
                   "Qo'shimchalari (-ning, -dan ...) odatdagidek o'giriladi.", AppNomi() & " - kirill istisnosi")
    If StrPtr(lat) = 0 Then Exit Sub
    lat = Trim$(lat)
    If Len(lat) > 0 Then entry = w & "=" & lat Else entry = w
    ReadState users, offs, CyrExceptionsPath()
    RemoveKey users, ExcKey(entry)
    users.Add entry
    SaveState users, offs, CyrExceptionsPath()
    mInit = False
    If Len(lat) > 0 Then
        UMsg "Qo'shildi: " & w & " -> " & lat
    Else
        UMsg "Qo'shildi: " & w & " (o'girilmaydi)"
    End If
End Sub

Public Sub YangiAlifbo_IstisnoAvto()
    If HasCyr(HostSelectedText()) Then
        YangiAlifbo_KirillIstisnoQoshish
    Else
        YangiAlifbo_IstisnoQoshish
    End If
End Sub

Private Function KeyText(ByVal id As String) As String
    Dim l As String
    l = HostKeyLetter(id)
    If Len(l) > 0 Then KeyText = "Alt+" & l
End Function

Private Function HelpKey(ByVal id As String, ByVal what As String) As String
    If Len(KeyText(id)) > 0 Then HelpKey = "   " & KeyText(id) & " - " & what & vbCrLf
End Function

Public Sub YangiAlifbo_Yordam()
    Dim t As String
    t = UCase$(AppNomi()) & " 1.2 - " & HostName() & " uchun" & vbCrLf & vbCrLf & _
        "Yangi alifboga - kirill ham, eski lotin ham yangi alifboga o'tadi." & vbCrLf & _
        "Eskichaga - yangi alifbodagi va kirill matnni amaldagi lotinga o'giradi." & vbCrLf & _
        "Kirillga - har qanday lotin yozuvini kirillga o'giradi." & vbCrLf & _
        "   Belgilangan " & HostSelWhat() & " bo'lsa - faqat ular, aks holda hammasi." & vbCrLf & _
        "Istisno so'zlar - o'girilmaydigan so'zlar (SHA-256, hash ...);" & vbCrLf & _
        "   ro'yxat Word bilan umumiy." & vbCrLf & vbCrLf
    If Len(KeyText("Lotin")) > 0 Then
        t = t & "TEZKOR TUGMALAR:" & vbCrLf & _
            HelpKey("Lotin", "yangi alifboga") & HelpKey("Eski", "eskichaga") & _
            HelpKey("LotinKirill", "kirillga") & _
            HelpKey("Istisno", "so'zni istisno so'zlarga qo'shish") & HelpKey("Istisnolar", "istisno so'zlarni boshqarish") & _
            "   (bir nechta katakcha belgilangan bo'lsa - faqat ular, aks holda butun kitob)" & vbCrLf & vbCrLf
    End If
    t = t & HostUndoHint() & vbCrLf & vbCrLf & _
        "Istisno so'zlar: " & ExceptionsPath() & vbCrLf & _
        "Kirill istisno so'zlari: " & CyrExceptionsPath() & vbCrLf & vbCrLf & _
        AppNomi() & " 1.2  -  muallif: Bozorov Obidjon Norqobilovich" & vbCrLf & _
        "Axborot xavfsizligi kafedrasi, O'zMU  -  o.bozorov@nuu.uz"
    MsgBox t, vbInformation, AppNomi() & " - yordam"
End Sub

'------------------------------ lenta -------------------------------------
Public Sub RB_OnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
    HostSetupKeys
End Sub

' tag="nom|amal" -> "nom (Alt+A)" (tezkor tugma bo'lmasa - faqat nom)
Public Sub RB_Label(control As IRibbonControl, ByRef returnedVal)
    Dim t As String, p As Long, k As String
    On Error Resume Next
    t = control.Tag
    p = InStr(t, "|")
    If p = 0 Then returnedVal = t: Exit Sub
    k = KeyText(Mid$(t, p + 1))
    If Len(k) > 0 Then
        returnedVal = Replace(Left$(t, p - 1), " ", ChrW(160)) & " (" & k & ")"
    Else
        returnedVal = Left$(t, p - 1)
    End If
End Sub

Public Sub RB_Hujjat(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_Hujjat
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_Tanlangan(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_Tanlangan
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_Kirill(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_Kirill
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_KirillTanlangan(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_KirillTanlangan
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_LotinKirill(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_LotinKirill
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_LotinKirillTanlangan(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_LotinKirillTanlangan
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_EskiLotin(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_EskiLotin
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_EskiLotinTanlangan(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_EskiLotinTanlangan
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_Istisnolar(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_Istisnolar
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_IstisnoAvto(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_IstisnoAvto
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_IstisnoRoyxat(control As IRibbonControl)
    On Error GoTo Xato
    ShowExceptionList
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_KirillIstisnolar(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_KirillIstisnolar
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_KirillIstisnoQoshish(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_KirillIstisnoQoshish
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_Yordam(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_Yordam
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_YangiAlifboga(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_Lotin
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_Eskichaga(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_EskiLotinAvto
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

Public Sub RB_Kirillga(control As IRibbonControl)
    On Error GoTo Xato
    YangiAlifbo_LotinKirillAvto
    Exit Sub
Xato:
    MsgBox "Xatolik: " & Err.Description, vbExclamation, AppNomi()
End Sub

'==========================================================================
'  POWERPOINT QISMI
'  Slaydlardagi barcha matn: sarlavha, matn bloklari, shakllar, guruhlar,
'  jadvallar va ma'ruzachi izohlari. Formatlash saqlanadi.
'==========================================================================
Private Function HostName() As String
    HostName = "PowerPoint"
End Function

Private Function HostSelWhat() As String
    HostSelWhat = "matn, shakllar yoki slaydlar"
End Function

Private Function HostUndoHint() As String
    HostUndoHint = "Bekor qilish: Ctrl+Z."
End Function

' PowerPoint tashqi makroslarga tezkor tugma bog'lashga imkon bermaydi
Private Function HostKeyLetter(ByVal id As String) As String
    HostKeyLetter = ""
End Function

Private Sub HostSetupKeys()
End Sub

Private Function SelType() As Long
    On Error Resume Next
    SelType = 0
    SelType = ActiveWindow.Selection.Type          ' 0 yo'q, 1 slaydlar, 2 shakllar, 3 matn
End Function

Private Function HostHasSelection() As Boolean
    Dim st As Long
    On Error Resume Next
    st = SelType()
    If st = 3 Then
        HostHasSelection = (ActiveWindow.Selection.TextRange2.Length > 0)
    Else
        HostHasSelection = (st = 1 Or st = 2)
    End If
End Function

Private Function HostSelectedText() As String
    On Error Resume Next
    Select Case SelType()
        Case 3: HostSelectedText = ActiveWindow.Selection.TextRange2.Text
        Case 2: HostSelectedText = ActiveWindow.Selection.ShapeRange(1).TextFrame2.TextRange.Text
    End Select
End Function

Private Sub RunConvert(ByVal onlySel As Boolean, Optional ByVal mode As Long = 0)
    Dim total As Long, sld As Object, shp As Object, st As Long, hadErr As Boolean
    If Presentations.Count = 0 Then Exit Sub
    If onlySel And Not HostHasSelection() Then
        MsgBox "Avval o'giriladigan matnni, shakllarni yoki slaydlarni belgilang.", vbInformation, AppNomi()
        Exit Sub
    End If
    InitAll
    mMode = mode
    mSkipped = 0
    On Error Resume Next
    Application.StartNewUndoEntry                   ' hammasi bitta Ctrl+Z bilan qaytsin
    On Error GoTo Xato
    If onlySel Then
        st = SelType()
        If st = 3 Then
            total = ConvertTR(ActiveWindow.Selection.TextRange2, True)
        ElseIf st = 2 Then
            For Each shp In ActiveWindow.Selection.ShapeRange
                total = total + ConvertShape(shp)
            Next shp
        Else
            For Each sld In ActiveWindow.Selection.SlideRange
                total = total + ConvertSlide(sld)
            Next sld
        End If
    Else
        For Each sld In ActivePresentation.Slides
            total = total + ConvertSlide(sld)
        Next sld
    End If
Tugash:
    mMode = 0
    If Not hadErr Then ResultMsg total, mode, HostUndoHint()
    Exit Sub
Xato:
    hadErr = True
    MsgBox "Xatolik: " & Err.Description, vbCritical, AppNomi()
    Resume Tugash
End Sub

' Slayd shakllari va ma'ruzachi izohi
Private Function ConvertSlide(ByVal sld As Object) As Long
    Dim shp As Object, cnt As Long, pt As Long
    For Each shp In sld.Shapes
        cnt = cnt + ConvertShape(shp)
    Next shp
    On Error Resume Next
    For Each shp In sld.NotesPage.Shapes
        pt = 0
        pt = shp.PlaceholderFormat.Type
        If pt = 2 Then cnt = cnt + ConvertShape(shp)      ' ppPlaceholderBody - izoh matni
    Next shp
    ConvertSlide = cnt
End Function
