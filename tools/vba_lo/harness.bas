' ---- test harness (LibreOffice) ----
Function ReadAll(path As String) As String
    Dim sfa As Object, ins As Object, tis As Object, s As String
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    ins = sfa.openFileRead(ConvertToURL(path))
    tis = CreateUnoService("com.sun.star.io.TextInputStream")
    tis.setInputStream(ins)
    tis.setEncoding("UTF-8")
    Do While Not tis.isEOF()
        s = s & tis.readLine() & Chr(10)
    Loop
    tis.closeInput()
    ReadAll = s
End Function

Sub WriteAll(path As String, s As String)
    Dim sfa As Object, outs As Object, tos As Object
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If sfa.exists(ConvertToURL(path)) Then sfa.kill(ConvertToURL(path))
    outs = sfa.openFileWrite(ConvertToURL(path))
    tos = CreateUnoService("com.sun.star.io.TextOutputStream")
    tos.setOutputStream(outs)
    tos.setEncoding("UTF-8")
    tos.writeString(s)
    tos.closeOutput()
End Sub

Function RunMode(ByVal t As String, ByVal md As Long) As String
    Dim n As Long, k As Long, o As String
    mMode = md
    n = ComputeEdits(t)
    o = t
    For k = n To 1 Step -1
        o = Left(o, mPos(k)) & mTo(k) & Mid(o, mPos(k) + mLen(k) + 1)
    Next k
    RunMode = o
End Function

Sub Main
    Dim dir As String, files As Variant, modes As Variant, f As Long, txt As String, lines As Variant, i As Long, res As String, p As Variant
    On Error GoTo Xato
    dir = "@DIR@/"
    InitAll
    files = Array("in_lat.tsv", "in_cyr.tsv", "in_l2c.tsv", "in_cyr2.tsv", "in_lat2.tsv", "in_l2cw.tsv", "in_l2ct.tsv", "in_n2o.tsv")
    modes = Array(0, 1, 2, 1, 0, 2, 2, 3)
    For f = 0 To 7
        txt = ReadAll(dir & files(f))
        lines = Split(txt, Chr(10))
        res = ""
        For i = 0 To UBound(lines)
            If Len(lines(i)) > 0 Then
                p = Split(lines(i), Chr(9))
                res = res & RunMode(p(0), modes(f)) & Chr(10)
            End If
        Next i
        WriteAll dir & "out_" & files(f), res
    Next f
    WriteAll dir & "status.txt", "OK"
    Exit Sub
Xato:
    WriteAll dir & "status.txt", "XATO " & Err & ": " & Error$ & " (qator " & Erl & ")"
End Sub
