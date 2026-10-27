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
