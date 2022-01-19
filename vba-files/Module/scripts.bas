Attribute VB_Name = "scripts"
Sub log(usuario, msg)

    Sheets("log").Visible = True
    Sheets("log").Select
    ultimaLinha = Range("A1000000").End(xlUp).Row + 1
    Range("A" & ultimaLinha).Value = Now()
    Range("B" & ultimaLinha).Value = usuario
    Range("C" & ultimaLinha).Value = msg
    Sheets("log").Visible = False

End Sub
