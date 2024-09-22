# kod.vba    Private Sub Worksheet_Change(ByVal Target As Range)
    Call SyncColors
End Sub



Dim nextRun As Date

Sub StartTimer()
    nextRun = Now + TimeValue("00:30:00") ' Kör varje 30 minut
    Application.OnTime nextRun, "RunSyncColors"
End Sub

Sub RunSyncColors()
    Call SyncColors
    StartTimer ' Starta om timern för nästa körning
End Sub

Sub StopTimer()
    On Error Resume Next
    Application.OnTime nextRun, "RunSyncColors", , False
End Sub

Sub SyncColors()
    Dim sourceSheet As Worksheet
    Dim mondaySheet As Worksheet
    Dim oneSheet As Worksheet
    Dim tuesdaySheet As Worksheet
    Dim additionalSheet As Worksheet
    Dim thursdaySheet As Worksheet
    Dim fridaySheet As Worksheet
    Dim saturdaySheet As Worksheet
    Dim sundaySheet As Worksheet
    Dim week8Sheet As Worksheet
    Dim week9Sheet As Worksheet
    Dim week10Sheet As Worksheet
    Dim week11Sheet As Worksheet
    Dim week12Sheet As Worksheet
    Dim week13Sheet As Worksheet
    Dim week14Sheet As Worksheet
    Dim week15Sheet As Worksheet
    Dim week16Sheet As Worksheet
    Dim week17Sheet As Worksheet
    Dim week18Sheet As Worksheet
    Dim week19Sheet As Worksheet
    Dim week20Sheet As Worksheet
    Dim week21Sheet As Worksheet
    Dim password As String

    password = "1981"

    ' Ställ in arken
    On Error Resume Next
    Set sourceSheet = ThisWorkbook.Sheets("Månadsschema")
    Set mondaySheet = ThisWorkbook.Sheets("Måndag-")
    Set oneSheet = ThisWorkbook.Sheets("-Måndag")
    Set tuesdaySheet = ThisWorkbook.Sheets("Tisdag-")
    Set additionalSheet = ThisWorkbook.Sheets("-Tisdag")
    Set thursdaySheet = ThisWorkbook.Sheets("Torsdag-")
    Set fridaySheet = ThisWorkbook.Sheets("Fredag-")
    Set saturdaySheet = ThisWorkbook.Sheets("Lördag-")
    Set sundaySheet = ThisWorkbook.Sheets("söndag-")
    Set week8Sheet = ThisWorkbook.Sheets("--Måndag")
    Set week9Sheet = ThisWorkbook.Sheets("--Tisdag")
    Set week10Sheet = ThisWorkbook.Sheets("--Onsdag")
    Set week11Sheet = ThisWorkbook.Sheets("--Torsdag")
    Set week12Sheet = ThisWorkbook.Sheets("--Fredag")
    Set week13Sheet = ThisWorkbook.Sheets("--Lördag")
    Set week14Sheet = ThisWorkbook.Sheets("--Söndag")
    Set week15Sheet = ThisWorkbook.Sheets("---Måndag")
    Set week16Sheet = ThisWorkbook.Sheets("---Tisdag")
    Set week17Sheet = ThisWorkbook.Sheets("---Onsdag")
    Set week18Sheet = ThisWorkbook.Sheets("---Torsdag")
    Set week19Sheet = ThisWorkbook.Sheets("---Fredag")
    Set week20Sheet = ThisWorkbook.Sheets("---Lördag")
    Set week21Sheet = ThisWorkbook.Sheets("---Söndag")
    On Error GoTo 0
    
    ' Kontrollera att arken finns
    If sourceSheet Is Nothing Or mondaySheet Is Nothing Or oneSheet Is Nothing Or tuesdaySheet Is Nothing Or additionalSheet Is Nothing Or thursdaySheet Is Nothing Or fridaySheet Is Nothing Or saturdaySheet Is Nothing Or sundaySheet Is Nothing Or week8Sheet Is Nothing Or week9Sheet Is Nothing Or week10Sheet Is Nothing Or week11Sheet Is Nothing Or week12Sheet Is Nothing Or week13Sheet Is Nothing Or week14Sheet Is Nothing Or week15Sheet Is Nothing Or week16Sheet Is Nothing Or week17Sheet Is Nothing Or week18Sheet Is Nothing Or week19Sheet Is Nothing Or week20Sheet Is Nothing Or week21Sheet Is Nothing Then
        MsgBox "Ett eller flera av de angivna bladen kunde inte hittas. Kontrollera att bladen finns och att namnen är korrekta.", vbCritical
        Exit Sub
    End If

    ' Ta bort skyddet från skyddade blad
    On Error Resume Next
    mondaySheet.Unprotect password
    oneSheet.Unprotect password
    tuesdaySheet.Unprotect password
    additionalSheet.Unprotect password
    thursdaySheet.Unprotect password
    fridaySheet.Unprotect password
    saturdaySheet.Unprotect password
    sundaySheet.Unprotect password
    week8Sheet.Unprotect password
    week9Sheet.Unprotect password
    week10Sheet.Unprotect password
    week11Sheet.Unprotect password
    week12Sheet.Unprotect password
    week13Sheet.Unprotect password
    week14Sheet.Unprotect password
    week15Sheet.Unprotect password
    week16Sheet.Unprotect password
    week17Sheet.Unprotect password
    week18Sheet.Unprotect password
    week19Sheet.Unprotect password
    week20Sheet.Unprotect password
    week21Sheet.Unprotect password
    On Error GoTo 0

    ' Synka färger från källområde till målområde
    SyncColorRange sourceSheet, "F5:F49", "-Måndag!D4:D48", "Måndag-!D4:D48"
    SyncColorRange sourceSheet, "G5:G49", "-Måndag!E4:E48", "Måndag-!E4:E48"
    SyncColorRange sourceSheet, "I5:I49", "-Måndag!X4:X48", "Måndag-!X4:X48"
    SyncColorRange sourceSheet, "J5:J49", "-Måndag!Y4:Y48", "Måndag-!Y4:Y48"

    ' Lägg till nya områden och blad här
    SyncColorRange sourceSheet, "R5:R49", "-Onsdag!D4:D48", "Onsdag-!D4:D48"
    SyncColorRange sourceSheet, "S5:S49", "-Onsdag!E4:E48", "Onsdag-!E4:E48"
    SyncColorRange sourceSheet, "U5:U49", "-Onsdag!X4:X48", "Onsdag-!X4:X48"
    SyncColorRange sourceSheet, "V5:V49", "-Onsdag!Y4:Y48", "Onsdag-!Y4:Y48"
    SyncColorRange sourceSheet, "L5:L49", "-Tisdag!D4:D48", "Tisdag-!D4:D48"
    SyncColorRange sourceSheet, "M5:M49", "-Tisdag!E4:E48", "Tisdag-!E4:E48"
    SyncColorRange sourceSheet, "O5:O49", "-Tisdag!X4:X48", "Tisdag-!X4:X48"
    SyncColorRange sourceSheet, "P5:P49", "-Tisdag!Y4:Y48", "Tisdag-!Y4:Y48"
    
    ' Nya områden för torsdag
    SyncColorRange sourceSheet, "X5:X49", "-Torsdag!D4:D48", "Torsdag-!D4:D48"
    SyncColorRange sourceSheet, "Y5:Y49", "-Torsdag!E4:E48", "Torsdag-!E4:E48"
    SyncColorRange sourceSheet, "AA5:AA49", "-Torsdag!X4:X48", "Torsdag-!X4:X48"
    SyncColorRange sourceSheet, "AB5:AB49", "-Torsdag!Y4:Y48", "Torsdag-!Y4:Y48"

    ' Nya områden för fredag
    SyncColorRange sourceSheet, "AD5:AD49", "-Fredag!D4:D48", "Fredag-!D4:D48"
    SyncColorRange sourceSheet, "AE5:AE49", "-Fredag!E4:E48", "Fredag-!E4:E48"
    SyncColorRange sourceSheet, "AG5:AG49", "-Fredag!X4:X48", "Fredag-!X4:X48"
    SyncColorRange sourceSheet, "AH5:AH49", "-Fredag!Y4:Y48", "Fredag-!Y4:Y48"
    
    ' Nya områden för lördag
    SyncColorRange sourceSheet, "AJ5:AJ49", "-Lördag!D4:D48", "Lördag-!D4:D48"
    SyncColorRange sourceSheet, "AK5:AK49", "-Lördag!E4:E48", "Lördag-!E4:E48"
    SyncColorRange sourceSheet, "AM5:AM49", "-Lördag!X4:X48", "Lördag-!X4:X48"
    SyncColorRange sourceSheet, "AN5:AN49", "-Lördag!Y4:Y48", "Lördag-!Y4:Y48"
    
    ' Nya områden för söndag
    SyncColorRange sourceSheet, "AP5:AP49", "-Söndag!D4:D48", "söndag-!D4:D48"
    SyncColorRange sourceSheet, "AQ5:AQ49", "-Söndag!E4:E48", "söndag-!E4:E48"
    SyncColorRange sourceSheet, "AS5:AS49", "-Söndag!X4:X48", "söndag-!X4:X48"
    SyncColorRange sourceSheet, "AT5:AT49", "-Söndag!Y4:Y48", "söndag-!Y4:Y48"

    ' Nya områden för vecka 8
    SyncColorRange sourceSheet, "BD5:BD49", "--Måndag!D4:D48"
    SyncColorRange sourceSheet, "BE5:BE49", "--Måndag!E4:E48"
    SyncColorRange sourceSheet, "BG5:BG49", "--Måndag!X4:X48"
    SyncColorRange sourceSheet, "BH5:BH49", "--Måndag!Y4:Y48"

    ' Nya områden för vecka 9
    SyncColorRange sourceSheet, "BJ5:BJ49", "--Tisdag!D4:D48"
    SyncColorRange sourceSheet, "BK5:BK49", "--Tisdag!E4:E48"
    SyncColorRange sourceSheet, "BM5:BM49", "--Tisdag!X4:X48"
    SyncColorRange sourceSheet, "BN5:BN49", "--Tisdag!Y4:Y48"

    ' Nya områden för vecka 10
    SyncColorRange sourceSheet, "BP5:BP49", "--Onsdag!D4:D48"
    SyncColorRange sourceSheet, "BQ5:BQ49", "--Onsdag!E4:E48"
    SyncColorRange sourceSheet, "BS5:BS49", "--Onsdag!X4:X48"
    SyncColorRange sourceSheet, "BT5:BT49", "--Onsdag!Y4:Y48"

    ' Nya områden för vecka 11
    SyncColorRange sourceSheet, "BV5:BV49", "--Torsdag!D4:D48"
    SyncColorRange sourceSheet, "BW5:BW49", "--Torsdag!E4:E48"
    SyncColorRange sourceSheet, "BY5:BY49", "--Torsdag!X4:X48"
    SyncColorRange sourceSheet, "BZ5:BZ49", "--Torsdag!Y4:Y48"

    ' Nya områden för vecka 12
    SyncColorRange sourceSheet, "CB5:CB49", "--Fredag!D4:D48"
    SyncColorRange sourceSheet, "CC5:CC49", "--Fredag!E4:E48"
    SyncColorRange sourceSheet, "CE5:CE49", "--Fredag!X4:X48"
    SyncColorRange sourceSheet, "CF5:CF49", "--Fredag!Y4:Y48"

    ' Nya områden för vecka 13
    SyncColorRange sourceSheet, "CH5:CH49", "--Lördag!D4:D48"
    SyncColorRange sourceSheet, "CI5:CI49", "--Lördag!E4:E48"
    SyncColorRange sourceSheet, "CK5:CK49", "--Lördag!X4:X48"
    SyncColorRange sourceSheet, "CL5:CL49", "--Lördag!Y4:Y48"

    ' Nya områden för vecka 14
    SyncColorRange sourceSheet, "CN5:CN49", "--Söndag!D4:D48"
    SyncColorRange sourceSheet, "CO5:CO49", "--Söndag!E4:E48"
    SyncColorRange sourceSheet, "CQ5:CQ49", "--Söndag!X4:X48"
    SyncColorRange sourceSheet, "CR5:CR49", "--Söndag!Y4:Y48"

    ' Nya områden för vecka 15
    SyncColorRange sourceSheet, "CZ5:CZ49", "---Måndag!D4:D48"
    SyncColorRange sourceSheet, "DA5:DA49", "---Måndag!E4:E48"
    SyncColorRange sourceSheet, "DC5:DC49", "---Måndag!X4:X48"
    SyncColorRange sourceSheet, "DD5:DD49", "---Måndag!Y4:Y48"

    ' Nya områden för vecka 16
    SyncColorRange sourceSheet, "DF5:DF49", "---Tisdag!D4:D48"
    SyncColorRange sourceSheet, "DG5:DG49", "---Tisdag!E4:E48"
    SyncColorRange sourceSheet, "DI5:DI49", "---Tisdag!X4:X48"
    SyncColorRange sourceSheet, "DJ5:DJ49", "---Tisdag!Y4:Y48"

    ' Nya områden för vecka 17
    SyncColorRange sourceSheet, "DL5:DL49", "---Onsdag!D4:D48"
    SyncColorRange sourceSheet, "DM5:DM49", "---Onsdag!E4:E48"
    SyncColorRange sourceSheet, "DO5:DO49", "---Onsdag!X4:X48"
    SyncColorRange sourceSheet, "DP5:DP49", "---Onsdag!Y4:Y48"

    ' Nya områden för vecka 18
    SyncColorRange sourceSheet, "DR5:DR49", "---Torsdag!D4:D48"
    SyncColorRange sourceSheet, "DS5:DS49", "---Torsdag!E4:E48"
    SyncColorRange sourceSheet, "DU5:DU49", "---Torsdag!X4:X48"
    SyncColorRange sourceSheet, "DV5:DV49", "---Torsdag!Y4:Y48"

    ' Nya områden för vecka 19
    SyncColorRange sourceSheet, "DX5:DX49", "---Fredag!D4:D48"
    SyncColorRange sourceSheet, "DY5:DY49", "---Fredag!E4:E48"
    SyncColorRange sourceSheet, "EA5:EA49", "---Fredag!X4:X48"
    SyncColorRange sourceSheet, "EB5:EB49", "---Fredag!Y4:Y48"

    ' Nya områden för vecka 20
    SyncColorRange sourceSheet, "ED5:ED49", "---Lördag!D4:D48"
    SyncColorRange sourceSheet, "EE5:EE49", "---Lördag!E4:E48"
    SyncColorRange sourceSheet, "EG5:EG49", "---Lördag!X4:X48"
    SyncColorRange sourceSheet, "EH5:EH49", "---Lördag!Y4:Y48"

    ' Nya områden för vecka 21
    SyncColorRange sourceSheet, "EJ5:EJ49", "---Söndag!D4:D48"
    SyncColorRange sourceSheet, "EK5:EK49", "---Söndag!E4:E48"
    SyncColorRange sourceSheet, "EM5:EM49", "---Söndag!X4:X48"
    SyncColorRange sourceSheet, "EN5:EN49", "---Söndag!Y4:Y48"

    ' Nya områden för månad 22
    SyncColorRange sourceSheet, "EV5:EV49", "----Måndag!D4:D48"
    SyncColorRange sourceSheet, "EW5:EW49", "----Måndag!E4:E48"
    SyncColorRange sourceSheet, "EY5:EY49", "----Måndag!X4:X48"
    SyncColorRange sourceSheet, "EZ5:EZ49", "----Måndag!Y4:Y48"

    ' Nya områden för månad 23
    SyncColorRange sourceSheet, "FB5:FB49", "----Tisdag!D4:D48"
    SyncColorRange sourceSheet, "FC5:FC49", "----Tisdag!E4:E48"
    SyncColorRange sourceSheet, "FE5:FE49", "----Tisdag!X4:X48"
    SyncColorRange sourceSheet, "FF5:FF49", "----Tisdag!Y4:Y48"

    ' Nya områden för månad 24
    SyncColorRange sourceSheet, "FH5:FH49", "----Onsdag!D4:D48"
    SyncColorRange sourceSheet, "FI5:FI49", "----Onsdag!E4:E48"
    SyncColorRange sourceSheet, "FK5:FK49", "----Onsdag!X4:X48"
    SyncColorRange sourceSheet, "FL5:FL49", "----Onsdag!Y4:Y48"

    ' Nya områden för månad 25
    SyncColorRange sourceSheet, "FN5:FN49", "----Torsdag!D4:D48"
    SyncColorRange sourceSheet, "FO5:FO49", "----Torsdag!E4:E48"
    SyncColorRange sourceSheet, "FQ5:FQ49", "----Torsdag!X4:X48"
    SyncColorRange sourceSheet, "FR5:FR49", "----Torsdag!Y4:Y48"

    ' Nya områden för månad 26
    SyncColorRange sourceSheet, "FT5:FT49", "----Fredag!D4:D48"
    SyncColorRange sourceSheet, "FU5:FU49", "----Fredag!E4:E48"
    SyncColorRange sourceSheet, "FW5:FW49", "----Fredag!X4:X48"
    SyncColorRange sourceSheet, "FX5:FX49", "----Fredag!Y4:Y48"

    ' Nya områden för månad 27
    SyncColorRange sourceSheet, "FZ5:FZ49", "----Lördag!D4:D48"
    SyncColorRange sourceSheet, "GA5:GA49", "----Lördag!E4:E48"
    SyncColorRange sourceSheet, "GC5:GC49", "----Lördag!X4:X48"
    SyncColorRange sourceSheet, "GD5:GD49", "----Lördag!Y4:Y48"

    ' Nya områden för månad 28
    SyncColorRange sourceSheet, "GF5:GF49", "----Söndag!D4:D48"
    SyncColorRange sourceSheet, "GG5:GG49", "----Söndag!E4:E48"
    SyncColorRange sourceSheet, "GI5:GI49", "----Söndag!X4:X48"
    SyncColorRange sourceSheet, "GJ5:GJ49", "----Söndag!Y4:Y48"

    ' Nya områden för månad 29
    SyncColorRange sourceSheet, "GR5:GR49", "------Måndag!D4:D48"
    SyncColorRange sourceSheet, "GS5:GS49", "------Måndag!E4:E48"
    SyncColorRange sourceSheet, "GU5:GU49", "------Måndag!X4:X48"
    SyncColorRange sourceSheet, "GV5:GV49", "------Måndag!Y4:Y48"

    ' Nya områden för månad 30
    SyncColorRange sourceSheet, "GX5:GX49", "------Tisdag!D4:D48"
    SyncColorRange sourceSheet, "GY5:GY49", "------Tisdag!E4:E48"
    SyncColorRange sourceSheet, "HA5:HA49", "------Tisdag!X4:X48"
    SyncColorRange sourceSheet, "HB5:HB49", "------Tisdag!Y4:Y48"

    ' Nya områden för månad 31
    SyncColorRange sourceSheet, "HD5:HD49", "------Onsdag!D4:D48"
    SyncColorRange sourceSheet, "HE5:HE49", "------Onsdag!E4:E48"
    SyncColorRange sourceSheet, "HG5:HG49", "------Onsdag!X4:X48"
    SyncColorRange sourceSheet, "HH5:HH49", "------Onsdag!Y4:Y48"

    ' Nya områden för månad 32
    SyncColorRange sourceSheet, "HJ5:HJ49", "Torsdag!D4:D48"
    SyncColorRange sourceSheet, "HK5:HK49", "Torsdag!E4:E48"
    SyncColorRange sourceSheet, "HM5:HM49", "Torsdag!X4:X48"
    SyncColorRange sourceSheet, "HN5:HN49", "Torsdag!Y4:Y48"

    ' Nya områden för månad 33
    SyncColorRange sourceSheet, "HP5:HP49", "Fredag!D4:D48"
    SyncColorRange sourceSheet, "HQ5:HQ49", "Fredag!E4:E48"
    SyncColorRange sourceSheet, "HS5:HS49", "Fredag!X4:X48"
    SyncColorRange sourceSheet, "HT5:HT49", "Fredag!Y4:Y48"

    ' Nya områden för månad 34
    SyncColorRange sourceSheet, "HV5:HV49", "Lördag!D4:D48"
    SyncColorRange sourceSheet, "HW5:HW49", "Lördag!E4:E48"
    SyncColorRange sourceSheet, "HY5:HY49", "Lördag!X4:X48"
    SyncColorRange sourceSheet, "HZ5:HZ49", "Lördag!Y4:Y48"

    ' Nya områden för månad 35
    SyncColorRange sourceSheet, "IB5:IB49", "Söndag!D4:D48"
    SyncColorRange sourceSheet, "IC5:IC49", "Söndag!E4:E48"
    SyncColorRange sourceSheet, "IE5:IE49", "Söndag!X4:X48"
    SyncColorRange sourceSheet, "IF5:IF49", "Söndag!Y4:Y48"

    ' Nya områden för månad 36
    SyncColorRange sourceSheet, "IN5:IN49", "Måndag!D4:D48"
    SyncColorRange sourceSheet, "IO5:IO49", "Måndag!E4:E48"
    SyncColorRange sourceSheet, "IQ5:IQ49", "Måndag!X4:X48"
    SyncColorRange sourceSheet, "IR5:IR49", "Måndag!Y4:Y48"

    ' Nya områden för månad 37
    SyncColorRange sourceSheet, "IT5:IT49", "Tisdag !D4:D48"
    SyncColorRange sourceSheet, "IU5:IU49", "Tisdag !E4:E48"
    SyncColorRange sourceSheet, "IW5:IW49", "Tisdag !X4:X48"
    SyncColorRange sourceSheet, "IX5:IX49", "Tisdag !Y4:Y48"
    
    ' Återställ skyddet
    mondaySheet.Protect password
    oneSheet.Protect password
    tuesdaySheet.Protect password
    additionalSheet.Protect password
    thursdaySheet.Protect password
    fridaySheet.Protect password
    saturdaySheet.Protect password
    sundaySheet.Protect password
    week8Sheet.Protect password
    week9Sheet.Protect password
    week10Sheet.Protect password
    week11Sheet.Protect password
    week12Sheet.Protect password
    week13Sheet.Protect password
    week14Sheet.Protect password
    week15Sheet.Protect password
    week16Sheet.Protect password
    week17Sheet.Protect password
    week18Sheet.Protect password
    week19Sheet.Protect password
    week20Sheet.Protect password
    week21Sheet.Protect password

End Sub

Sub SyncColorRange(sourceSheet As Worksheet, sourceRangeAddress As String, ParamArray targetRangeAddresses() As Variant)
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim address As Variant
    Dim cell As Range

    ' Definiera källområdet
    Set sourceRange = sourceSheet.Range(sourceRangeAddress)

    ' Loopa igenom varje målområde
    For Each address In targetRangeAddresses
        Set targetRange = ThisWorkbook.Sheets(Split(address, "!")(0)).Range(Split(address, "!")(1))

        ' Kopiera färger från källområdet till målområdet
        For Each cell In sourceRange
            targetRange.Cells(cell.row - sourceRange.row + 1, cell.Column - sourceRange.Column + 1).Interior.Color = cell.Interior.Color
        Next cell
    Next address
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Kör synkronisering av färger innan du sparar arbetsboken
    Call SyncColors
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Kör synkronisering av färger innan du sparar arbetsboken
    Call SyncColors
End Sub

Private Sub Workbook_Open()
    Call VisaBladVeckaFörVecka
    Call SkickaEpost
    Call SkapaKopia
End Sub

Sub VisaBladVeckaFörVecka()
    Dim ws As Worksheet
    Dim dagensDatum As Date
    Dim bladDatum As Date
    Dim veckaSlutDatum As Date
    Dim aktuellTid As Double

    ' Hämta dagens datum och aktuell tid
    dagensDatum = Date
    aktuellTid = Time

    ' Bestäm veckans slutdatum (7 dagar framåt)
    veckaSlutDatum = dagensDatum + 6

    ' Loopa genom alla blad
    For Each ws In ThisWorkbook.Sheets
        ' Håll specifika blad synliga
        If ws.Name = "Månadsschema" Or _
           ws.Name = "Dagliguppföljning" Or _
           ws.Name = "Veckor & Månaduppföljning" Or _
           ws.Name = "Sjukskrivningar" Or _
           ws.Name = "Utskrift" Or _
           ws.Name = "Månadsuppföljning" Then
            ws.Visible = xlSheetVisible
        Else
            ' Kontrollera datum i cell C2
            On Error Resume Next
            bladDatum = CDate(ws.Range("C2").Value)
            On Error GoTo 0

            ' Dölja eller visa baserat på datum om tiden är efter kl 02:00
            If aktuellTid >= TimeValue("02:00:00") Then
                If bladDatum >= dagensDatum And bladDatum <= veckaSlutDatum Then
                    ws.Visible = xlSheetVisible
                Else
                    ws.Visible = xlSheetHidden
                End If
            End If
        End If
    Next ws
End Sub

Sub SkickaEpost()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim igår As Date
    Dim cell As Range
    Dim row As Range
    Dim rowContent As String
    Dim hasContent As Boolean
    igår = Date - 1 ' Dagen som har passerat

    Set OutlookApp = CreateObject("Outlook.Application")

    For Each ws In ThisWorkbook.Worksheets
        If IsDate(ws.Range("C2").Value) Then
            ' Kontrollera om datumet i cell C2 är igår
            If ws.Range("C2").Value = igår Then
                Set OutlookMail = OutlookApp.CreateItem(0)
                With OutlookMail
                    .To = "Ilker_sali@hotmail.com" ' Din e-postadress
                    .Subject = "Automatisk e-post för " & ws.Range("C2").Value
                    .Body = "Detta är ett meddelande för " & ws.Range("C2").Value & vbCrLf & vbCrLf

                    ' Loopa genom raderna i L1:T36
                    For Each row In ws.Range("L1:T36").Rows
                        rowContent = ""
                        hasContent = False ' Återställ flaggan för varje rad

                        ' Loopa genom cellerna i raden
                        For Each cell In row.Cells
                            If cell.Value <> 0 Then ' Kontrollera om cellen inte är noll
                                rowContent = rowContent & cell.Value & vbTab
                                hasContent = True ' Om vi har innehåll
                            End If
                        Next cell

                        ' Lägg till raden i meddelandet om den har innehåll
                        If hasContent Then
                            .Body = .Body & rowContent & vbCrLf
                        End If
                    Next row

                    .Send
                End With
                Set OutlookMail = Nothing
            End If
        End If
    Next ws

    Set OutlookApp = Nothing
End Sub

Sub SkapaKopia()
    ' Variabler
    Dim sheetToCopy As Worksheet
    Dim newSheet As Worksheet
    Dim lösenordBlad As String
    Dim newSheetName As String
    Dim skyddatOmråde1 As Range
    Dim skyddatOmråde2 As Range
    Dim skyddatOmråde3 As Range
    Dim skyddatOmråde4 As Range
    Dim todayDate As Date

    ' Lösenord och datum
    lösenordBlad = "1981"
    newSheetName = "Schema-Lagring"
    todayDate = Date

    ' Kontrollera om idag är den 1:a i månaden
    If Day(todayDate) = 1 Then
        ' Kontrollera om bladet "Månadsschema" redan finns
        On Error Resume Next
        Set sheetToCopy = ThisWorkbook.Sheets("Månadsschema")
        On Error GoTo 0

        If Not sheetToCopy Is Nothing Then
            ' Skydda originalbladet med lösenord 1981
            sheetToCopy.Protect password:=lösenordBlad

            ' Skapa kopia av bladet "Månadsschema"
            sheetToCopy.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            Set newSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

            ' Namnge det nya bladet
            On Error Resume Next
            newSheet.Name = newSheetName
            If Err.Number <> 0 Then
                MsgBox "Kunde inte namnge det nya bladet. Fel: " & Err.Description
                Exit Sub
            End If
            On Error GoTo 0

            ' Skydda hela kopian med lösenord 1981
            newSheet.Protect password:=lösenordBlad

            ' Definiera de skyddade områdena
            Set skyddatOmråde1 = newSheet.Range("D5:D50")   ' Område D
            Set skyddatOmråde2 = newSheet.Range("KB5:KB50")  ' Område KB
            Set skyddatOmråde3 = newSheet.Range("E1:E2")      ' Område E1 till E2
            Set skyddatOmråde4 = newSheet.Range("AV1:AV2")    ' Område AV1 till AV2

            ' Ta bort skyddet för att skapa redigerbara områden
            newSheet.Unprotect password:=lösenordBlad

            ' Skapa redigerbara områden
            newSheet.Protection.AllowEditRanges.Add Title:="Skyddat område D", _
                Range:=skyddatOmråde1, password:=""
            newSheet.Protection.AllowEditRanges.Add Title:="Skyddat område KB", _
                Range:=skyddatOmråde2, password:=""
            newSheet.Protection.AllowEditRanges.Add Title:="Skyddat område E", _
                Range:=skyddatOmråde3, password:=""
            newSheet.Protection.AllowEditRanges.Add Title:="Skyddat område AV", _
                Range:=skyddatOmråde4, password:=""

            ' Skydda hela kopian igen
            newSheet.Protect password:=lösenordBlad, UserInterfaceOnly:=True

            ' Meddela användaren om att kopian har skapats
            MsgBox "En kopia av schemat har sparats som " & newSheetName

        Else
            MsgBox "Bladet 'Månadsschema' hittades inte. Inget sparades."
        End If
    Else
        MsgBox "Idag är inte den 1:a i månaden. Ingen kopia skapades."
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim ws As Worksheet
    Dim dagensMål As Double
    Dim kvällensMål As Double
    Dim dagensResultat As Variant
    Dim kvällensResultat As Variant
    Dim lösenord As String
    lösenord = "1981"

    For Each ws In ThisWorkbook.Worksheets
        If ws.Range("C2").value = Date Then
            ws.Unprotect password:=lösenord
            
            If Not IsEmpty(ws.Range("P6")) And Not IsEmpty(ws.Range("P19")) And _
               Not IsEmpty(ws.Range("O33")) And Not IsEmpty(ws.Range("O34")) Then

                dagensMål = ws.Range("P6").value
                kvällensMål = ws.Range("P19").value
                dagensResultat = ws.Range("O33").value
                kvällensResultat = ws.Range("O34").value

                ws.Range("O33:O34").Interior.ColorIndex = xlNone

                ' Färga och meddelande för dagens resultat
                Dim meddelande As String
                meddelande = ""

                If IsNumeric(dagensResultat) And dagensResultat > 0 Then
                    If dagensResultat >= dagensMål Then
                        ws.Range("O33").Interior.Color = RGB(255, 0, 0) ' Röd
                        meddelande = Choose(Application.WorksheetFunction.RandBetween(1, 3), _
                            "Bra försök! Tänk på vad du kan göra annorlunda nästa gång!", _
                            "Varje steg räknas! Fortsätt kämpa!", _
                            "Ingen är perfekt! Använd detta som en lärdom för framtiden!")
                    Else
                        ws.Range("O33").Interior.Color = RGB(0, 255, 0) ' Grön
                        meddelande = Choose(Application.WorksheetFunction.RandBetween(1, 3), _
                            "Fantastiskt arbete idag! Du har nått ditt mål!", _
                            "Bra jobbat! Du är på väg att nå dina drömmar!", _
                            "Utmärkt prestation! Ditt hårda arbete lönar sig!")
                    End If
                End If

                ' Färga och meddelande för kvällens resultat
                If IsNumeric(kvällensResultat) And kvällensResultat > 0 Then
                    If kvällensResultat >= kvällensMål Then
                        ws.Range("O34").Interior.Color = RGB(255, 0, 0) ' Röd
                        meddelande = meddelande & vbCrLf & Choose(Application.WorksheetFunction.RandBetween(1, 3), _
                            "Fortsätt kämpa! Du har vad som krävs!", _
                            "Det är alltid möjligt att förbättra sig. Du kan göra det!", _
                            "Kom ihåg, varje försök räknas. Tveka inte att ge allt nästa gång!")
                    Else
                        ws.Range("O34").Interior.Color = RGB(0, 255, 0) ' Grön
                        meddelande = meddelande & vbCrLf & Choose(Application.WorksheetFunction.RandBetween(1, 3), _
                            "Strålande insats ikväll! Fortsätt att bygga på detta!", _
                            "Du har verkligen gjort framsteg! Håll i det!", _
                            "Fantastiskt! Din insats ikväll är imponerande!")
                    End If
                End If

                If meddelande <> "" Then
                    MsgBox meddelande, vbInformation, "Motiverande Meddelande"
                End If
            End If
            
            ws.Protect password:=lösenord
        End If
    Next ws
End Sub



