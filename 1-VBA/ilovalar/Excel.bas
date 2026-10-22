'==========================================================================
'  EXCEL QISMI
'  Katakchalardagi matn (formulalar emas) va shakllardagi matn o'giriladi.
'  Bekor qilish: Ctrl+Z (katakchalar uchun - Application.OnUndo orqali).
'==========================================================================
Private mUndoObj As Collection
Private mUndoVal As Collection

Private Function HostName() As String
    HostName = "Excel"
End Function

Private Function HostSelWhat() As String
    HostSelWhat = "katakchalar yoki shakllar"
End Function

Private Function HostUndoHint() As String
    HostUndoHint = "Bekor qilish: Ctrl+Z (katakchalar uchun)."
End Function

Private Function HostKeyLetter(ByVal id As String) As String
    Select Case id
        Case "Lotin": HostKeyLetter = "A"
        Case "Eski": HostKeyLetter = "E"
        Case "LotinKirill": HostKeyLetter = "C"
        Case "Istisno": HostKeyLetter = "I"
        Case "Istisnolar": HostKeyLetter = "O"
        Case Else: HostKeyLetter = ""
    End Select
End Function

Private Function MacroRef(ByVal m As String) As String
    MacroRef = "'" & ThisWorkbook.Name & "'!" & m
End Function

' Alt + harf (Excel OnKey: % = Alt)
Private Sub HostSetupKeys()
    On Error Resume Next
    Application.OnKey "%a", MacroRef("YangiAlifbo_Lotin")
    Application.OnKey "%e", MacroRef("YangiAlifbo_EskiLotinAvto")
    Application.OnKey "%c", MacroRef("YangiAlifbo_LotinKirillAvto")
    Application.OnKey "%i", MacroRef("YangiAlifbo_IstisnoAvto")
    Application.OnKey "%o", MacroRef("YangiAlifbo_Istisnolar")
End Sub

' Bir nechta katakcha yoki shakl belgilangan bo'lsa - faqat ular; bitta katakcha - butun kitob
Private Function HostHasSelection() As Boolean
    On Error Resume Next
    HostHasSelection = False
    If TypeName(Selection) = "Range" Then
        HostHasSelection = (Selection.Cells.CountLarge > 1)
    Else
        HostHasSelection = True
    End If
End Function

Private Function HostSelectedText() As String
    On Error Resume Next
    If TypeName(Selection) = "Range" Then
        HostSelectedText = CStr(ActiveCell.Value)
    Else
        HostSelectedText = Selection.ShapeRange(1).TextFrame2.TextRange.Text
    End If
End Function

Private Sub RunConvert(ByVal onlySel As Boolean, Optional ByVal mode As Long = 0)
    Dim total As Long, ws As Object, rng As Object, shp As Object, prot As Long, hadErr As Boolean
    If ActiveWorkbook Is Nothing Then Exit Sub
    InitAll
    mMode = mode
    mSkipped = 0
    Set mUndoObj = New Collection
    Set mUndoVal = New Collection
    On Error GoTo Xato
    Application.ScreenUpdating = False
    If onlySel Then
        If TypeName(Selection) = "Range" Then
            If Selection.Cells.CountLarge = 1 Then
                total = ConvertCells(Selection)
            Else
                Set rng = Application.Intersect(Selection, Selection.Worksheet.UsedRange)
                If Not rng Is Nothing Then total = ConvertCells(rng)
            End If
        Else
            For Each shp In Selection.ShapeRange
                total = total + ConvertShape(shp)
            Next shp
        End If
    Else
        For Each ws In ActiveWorkbook.Worksheets
            If ws.ProtectContents Then
                prot = prot + 1
            Else
                total = total + ConvertCells(ws.UsedRange)
                For Each shp In ws.Shapes
                    total = total + ConvertShape(shp)
                Next shp
            End If
        Next ws
    End If
Tugash:
    mMode = 0
    Application.ScreenUpdating = True
    If mUndoObj.Count > 0 Then
        On Error Resume Next
        Application.OnUndo AppNomi() & ": o'girishni bekor qilish", MacroRef("OzbekAlifbo_Qaytarish")
        On Error GoTo 0
    End If
    If Not hadErr Then
        ResultMsg total, mode, IIf(prot > 0, prot & " ta himoyalangan varaq o'tkazib yuborildi." & vbCrLf, "") & HostUndoHint()
    End If
    Exit Sub
Xato:
    hadErr = True
    MsgBox "Xatolik: " & Err.Description, vbCritical, AppNomi()
    Resume Tugash
End Sub

' Matnli (formula bo'lmagan) katakchalar
Private Function ConvertCells(ByVal rng As Object) As Long
    Dim c As Object, t As String, nt As String, n As Long, cnt As Long, pre As String
    For Each c In rng.Cells
        If Not c.HasFormula Then
            If VarType(c.Value) = vbString Then
                t = c.Value
                nt = ConvertString(t, n)
                If n > 0 Then
                    pre = ""
                    If c.PrefixCharacter = "'" Then pre = "'"
                    mUndoObj.Add c
                    mUndoVal.Add pre & t
                    c.Value = pre & nt
                    cnt = cnt + n
                End If
            End If
        End If
    Next c
    ConvertCells = cnt
End Function

' Ctrl+Z: oxirgi o'girishdan oldingi qiymatlarni qaytaradi
Public Sub OzbekAlifbo_Qaytarish()
    Dim i As Long
    On Error Resume Next
    If mUndoObj Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    For i = mUndoObj.Count To 1 Step -1
        mUndoObj.Item(i).Value = mUndoVal.Item(i)
    Next i
    Application.ScreenUpdating = True
    Set mUndoObj = New Collection
    Set mUndoVal = New Collection
End Sub
