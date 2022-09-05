Dim msXML As XMLHTTP60
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub TumIPUlkeleriniGetir()

    Dim strIpAdresi, strUlkeAdi As String

    UlkeHucreleriniTemizle
    
    Application.ScreenUpdating = True
    
    Dim c As Range, r As Range
    Dim Sum As Long
    For Each r In Selection.Rows
        For Each c In r.Cells
            strIpAdresi = c.Value
            strUlkeAdi = getCountry(strIpAdresi)
            ActiveSheet.Cells(c.Row, 12).Value = strUlkeAdi
            Debug.Print strIpAdresi, strUlkeAdi
            Application.Wait (Now + TimeValue("00:00:" & Int((3 * Rnd) + 1)))
            ' Sleep 500
        Next c
    Next r
End Sub

Sub UlkeHucreleriniTemizle()
    Range("L2:L" & Rows.Count).Clear
End Sub

Public Function getCountry(ByVal ip As String) As String
    Dim resp, country As String
    Dim Json As Object
    
    JsonConverter.JsonOptions.AllowUnquotedKeys = True
    
    resp = getHTTP("https://ipleak.net/?mode=json&ip=" & ip, "GET")
    
    Set Json = JsonConverter.ParseJson(resp)
    getCountry = Json("country_name")
End Function

Public Function getHTTP(ByVal url As String, ByVal method As String) As String
    If msXML Is Nothing Then Set msXML = New XMLHTTP60
    With msXML
        .Open method, url, False: .Send
        getHTTP = StrConv(.responseBody, vbUnicode)
    End With
End Function
