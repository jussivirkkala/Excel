# Excel macros

Do not enable macro content before inspecting the code!

Älä aktivoi makroja ennenkuin olet tarkistanut koodin!

# HS-koronavirus-avoindata

https://github.com/HS-Datadesk/koronavirus-avoindata datan nouto ja visualisointi.

- 2020-03-14 Lisätty graafi. Jossain koneissa ei toimi "Virhe tulkita dataa: ActiveX component can't create object".
- 2020-03-13 Ensimmäinen versio. Automaattinen päivitys ei toimi ensimmäisellä kerralla makron hyväksynnän jälkeen.

VBA koodi alla. On oltava nimettynä solu Tapaukset, jossa kaava =sum(A:A). Data kirjoitetaan rivistä 16 eteenpäin sarakkeille A-F
```
Option Explicit

' Hakee HS datasta Korona tapaukset
' https://github.com/HS-Datadesk/koronavirus-avoindata
' https://w3qa5ydb4l.execute-api.eu-west-1.amazonaws.com/prod/finnishCoronaData
'
' https://github.com/jussivirkkala/excel/
' https://twitter.com/jussivirkkala
' 2020-03-13 First version
' 2020-03-14 More error handling. Modified text.

Sub Workbook_open()
    If MsgBox("Haluatko hakea https://github.com/HS-Datadesk/koronavirus-avoindata myös 15 min välein? Excel on oltava auki. Saat uusista tapauksista ilmoituksen.", _
    vbYesNo, Application.Name) = vbYes Then
        Timer
    Else
        Update
    End If
End Sub

Sub Timer()
    Update
    On Error GoTo err:
    Application.OnTime Now + TimeValue("00:15:00"), "ThisWorkbook.Timer"
    Exit Sub
err:
    MsgBox ("Ajastus ei onnistunut. Avaa tallennettu tiedosto uudestaan")
End Sub

Sub Update()
    Dim DATA As String
    DATA = "https://w3qa5ydb4l.execute-api.eu-west-1.amazonaws.com/prod/finnishCoronaData"
    
    ' GET data
    On Error GoTo err_get
    Dim request
    Set request = New ServerXMLHTTP60
    request.Open "GET", DATA, False
    request.send

    ' Parse JSON
    On Error GoTo err_json
    Dim sc, json
    Set sc = CreateObject("ScriptControl"): sc.Language = "JScript"
    Set json = sc.Eval("(" + request.responseText + ")")
    
    Dim n As Long
    n = ActiveSheet.Range("Tapauksia")
        
    Dim row As Long
    Dim subject As Object
    For Each subject In CallByName(json, "confirmed", VbGet)
        row = subject.id + 15
        ActiveSheet.Cells(row, 1) = subject.id
        Dim d As String
        d = CallByName(subject, "date", VbGet)
        ActiveSheet.Cells(row, 2) = d
        ActiveSheet.Cells(row, 3) = subject.healthCareDistrict
        ActiveSheet.Cells(row, 4) = subject.infectionSourceCountry
        ActiveSheet.Cells(row, 5) = subject.infectionSource
        ActiveSheet.Cells(row, 6) = DateValue(Mid(d, 1, 10)) + TimeValue(Mid(d, 12, 8))
    Next

    On Error GoTo err_other
    Application.CalculateFull
    If ActiveSheet.Range("Tapauksia") <> n Then
        MsgBox "Tapaukset ovat lisäntyneet " + Format(ActiveSheet.Range("Tapauksia").Value - n) + " kappaletta" _
        , , Application.Name
    End If
    Exit Sub
err_get:
    MsgBox "Virhe lukea " + DATA + " :" + err.Description
    Exit Sub
err_json:
    MsgBox "Virhe tulkita dataa:", err.Description
    Exit Sub
err_other:
    MsgBox "Virher Muu virhe: ", err.Description
End Sub

' End
```
