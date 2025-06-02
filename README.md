Public Function LlamarOllama(p As String) As String
    Dim h As Object
    Dim j As String
    Dim r As String

    ' Escapar caracteres problemáticos
    j = "{""model"":""llama3.2"",""prompt"":""" & _
        Replace(Replace(p, "\", "\\"), """", "\""") & _
        """,""stream"":false}"

    ' Crear petición HTTP
    Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.Open "POST", "http://localhost:11434/api/generate", False
    h.SetRequestHeader "Content-Type", "application/json"
    h.Send j

    r = h.ResponseText
   
    r = Mid(h.ResponseText, InStr(h.ResponseText, """response"":""") + 12)
    r = Left(r, InStr(r, """") - 1)
	  r = Replace(r, "\n", "") ' <- Esto quita los \n literales
    LlamarOllama = Trim(r)
End Function
