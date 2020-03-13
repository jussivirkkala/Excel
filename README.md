# Excel macros

Do not enable macro content before inspecting the code!

Älä aktivoi makroja ennenkuin olet tarkistanut koodin!

# HS-koronavirus-avoindata

https://github.com/HS-Datadesk/koronavirus-avoindata Excel datan nouto ja visualisointi.

VBA koodi alla. On nimettävä solut Tapaukset =sum(A:A) sekä Paivitetty. Data kirjoitetaan rivistä 11 eteenpäin sarakkeille A-F
```
Option Explicit

' Hakee HS datasta Korona tapaukset
' https://github.com/HS-Datadesk/koronavirus-avoindata
' https://w3qa5ydb4l.execute-api.eu-west-1.amazonaws.com/prod/finnishCoronaData
'
' https://github.com/jussivirkkala/excel/
' https://twitter.com/jussivirkkala
' 2020-03-13 First version

Dim dialog As Boolean


Sub Workbook_open()
    dialog = False
    If MsgBox("Haluatko hakea  https://github.com/HS-Datadesk/koronavirus-avoindata 15 min välein. Excel oltava auki", vbYesNo) = vbYes Then
        Timer
    Else
        dialog = True
        Update
    End If
End Sub


Sub Timer()
    Update
    Application.OnTime Now + TimeValue("00:15:00"), "ThisWorkbook.Timer"
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
        row = subject.id + 10
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
        MsgBox "Tapaukset ovat lisäntyneet " + Format(ActiveSheet.Range("Tapauksia").Value - n) + _
        " kappaletta edellisestä päivityksestä " + ActiveSheet.Range("Paivitetty").Text, , ActiveWorkbook.Name
    End If
    ActiveSheet.Range("Paivitetty") = datetime.Now
    Exit Sub
err_get:
    MsgBox "Virhe lukea " + DATA
    Exit Sub
err_json:
    MsgBox "Virhe tulkita dataa"
err_other:
    MsgBox "Muu virhe"
    Exit Sub

End Sub

' End
```
