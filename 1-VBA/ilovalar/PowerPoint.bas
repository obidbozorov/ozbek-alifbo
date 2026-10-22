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
