# Excel macros

Do not enable macro content before inspecting the code!

Älä aktivoi makroja ennenkuin olet tarkistanut koodin!

#  [all-exposure-checks.xlsx](https://github.com/jussivirkkala/excel/blob/master/all-exposure-checks.xlsx)

Parsing COVID-19 exposure notification file all-exposure-checks.json into Excel graph. No Excel macro needed. Having history data of https://www.koronavilkku.fi/.

- 2020-09-27 Changed to line graph. See #koronavilkku information https://thl.fi/fi/web/hyvinvoinnin-ja-terveyden-edistamisen-johtaminen/ajankohtaista/koronan-vaikutukset-yhteiskuntaan-ja-palveluihin#Koronavilkkua
- 2020-09-26 First version. Supporting FIN, UK until end of 2020 (format of times).

# [HS-koronavirus-avoindata.xlsm](https://github.com/jussivirkkala/excel/blob/master/hs-koronavirus-avoindata.xlsm)

https://github.com/HS-Datadesk/koronavirus-avoindata datan nouto ja visualisointi. 

- 2020-09-29 Päivitetty datalähde https://w3qa5ydb4l.execute-api.eu-west-1.amazonaws.com/prod/finnishCoronaData/v2. Lähteessä ei enää maata eikä lähdettä. Poistettu automaattinen päivitys sekä ajastus.
- 2020-04-01 Datalähdettä ei päivitetä. THL uuden datan sijainti https://thl.fi/fi/tilastot-ja-data/aineistot-ja-palvelut/avoin-data/varmistetut-koronatapaukset-suomessa-covid-19-
- 2020-03-31 Korjattu kuvaajan x-akselin vaihtuminen esim. tallennuksen yhteydessä.
- 2020-03-16 Lisätty Päivitä painike. Lisätty sairaahoitopiirit Tilastot välilehdelle. Joissain koneissa "Virhe lukea ... Toiminnon aikakatkaisu" vaikka sivulle https://w3qa5ydb4l.execute-api.eu-west-1.amazonaws.com/prod/finnishCoronaData pääsee.
- 2020-03-15 Lisätty välilehdet. Puuttuvat id numerot eivät enää aiheuta tyhjiä riviä Data välilehdellä. Tämä helpottaa lajittelua ja auto filter käyttöä. Ei toimi myöskään Mac koneissa: Can't find project or library ServerXMLHTTP60
- 2020-03-14 Lisätty graafi. Joissain koneissa ei toimi: "Virhe tulkita dataa: ActiveX component can't create object".
- 2020-03-13 Ensimmäinen versio. Automaattinen ajastus ei toimi ensimmäisellä kerralla makron hyväksynnän jälkeen.

VBA koodi alla. On oltava nimettynä solu Data väliehdellä Tapaukset, jossa kaava =sum(A:A). Data kirjoitetaan rivistä 2 eteenpäin sarakkeille A-F
```
Option Explicit

' Hakee HS datasta Korona tapaukset
' https://github.com/HS-Datadesk/koronavirus-avoindata
' https://w3qa5ydb4l.execute-api.eu-west-1.amazonaws.com/prod/finnishCoronaData
'
' https://github.com/jussivirkkala/excel/
' https://twitter.com/jussivirkkala
'
' 2020-09-29 Disabled automatic update and timer
' 2020-06-13 Added button
' 2020-03-15 Dialog for no cases. Time of last get.
' 2020-03-14 More error handling. Modified text.
' 2020-03-13 First version

Dim DIALOG As Boolean


Sub Workbook_open_off()
    DIALOG = True
    If MsgBox("Haluatko hakea https://github.com/HS-Datadesk/koronavirus-avoindata myös 15 min välein? Excel on oltava auki. Saat uusista tapauksista ilmoituksen.", _
    vbYesNo, Application.Name) = vbYes Then
        Timer
        DIALOG = False
    Else
        UpdateDialog
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

Sub UpdateDialog()
    DIALOG = True
    Update
    DIALOG = False
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
    n = Sheets("Data").Range("Tapauksia")
    
    Dim row As Long
    row = 1
    Dim subject As Object
    Application.Calculation = xlCalculationManual
    For Each subject In CallByName(json, "confirmed", VbGet)
        row = row + 1
        Sheets("Data").Cells(row, 1) = subject.id
        Dim d As String
        d = CallByName(subject, "date", VbGet)
        Sheets("Data").Cells(row, 2) = d
        Sheets("Data").Cells(row, 3) = subject.healthCareDistrict
        Sheets("Data").Cells(row, 4) = subject.infectionSourceCountry
        Sheets("Data").Cells(row, 5) = subject.infectionSource
        Sheets("Data").Cells(row, 6) = DateValue(Mid(d, 1, 10)) + TimeValue(Mid(d, 12, 8))
    Next

    On Error GoTo err_other
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull
    If Sheets("Data").Range("Tapauksia") <> n Then
        MsgBox "Tapaukset ovat lisäntyneet " + Format(Sheets("Data").Range("Tapauksia").Value - n) + " kappaletta " _
        + Sheets("Data").Range("Paivitetty").Text + " jälkeen.", , Application.Name
        Sheets("Data").Range("Paivitetty") = datetime.Now()
    Else
        If DIALOG Then MsgBox "Ei uusia tilastoituja tapauksia " + Sheets("Data").Range("Paivitetty").Text + " jälkeen.", , Application.Name
    End If
    
    Exit Sub
err_get:
    MsgBox "Virhe lukea " + DATA + ": " + err.Description
    Exit Sub
err_json:
    MsgBox "Virhe tulkita dataa: " + err.Description
    Exit Sub
err_other:
    MsgBox "Muu virhe: " + err.Description
End Sub

' End
```
