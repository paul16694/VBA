Attribute VB_Name = "Modul16"
Sub SendEmail()
    ' Variablen für Outlook und Email
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    
    ' Variablen für die Excel-Zellenwerte
    Dim stueck As String
    Dim material As String
    Dim Materialbeistellung As String
    Dim empfanger As String
    Dim betreff As String
    Dim empfangsadresse As String
    
       Dim zeile As Integer
    
    ' Benutzer nach der Zeilennummer fragen
    zeile = InputBox("Bitte geben Sie die Zeilennummer ein:", "Zeilennummer eingeben")
    
    ' Überprüfen, ob eine gültige Zeilennummer eingegeben wurde
    If zeile <= 0 Or zeile > Rows.Count Then
        MsgBox "Ungültige Zeilennummer. Das Makro wird beendet.", vbExclamation
        Exit Sub
    End If
    
    ' Outlook starten
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0) ' 0 steht für olMailItem
    
    ' Werte aus den Excel-Zellen holen
    stueck = Sheets("Bestellung").Range("H3").Value
    material = Sheets("Bestellung").Range("G3").Value
    Materialbeistellung = Sheets("Bestellung").Range("I3").Value
    empfanger = Sheets("Bestellung").Range("J3").Value
    betreff = "Preisanfrage " & Sheets("Bestellung").Range("A3").Value
    bezeichnung = Sheets("Bestellung").Range("F3").Value
    
    ' Empfängeradresse basierend auf dem Wert in Bestellung!J3 bestimmen
    If empfanger = "EWE" Then
        empfangsadresse = "paul.schulze@online.de"
    ElseIf empfanger = "SSL" Then
        empfangsadresse = "@ddfdfdf"
    Else
        empfangsadresse = "" ' Leere Adresse, falls nichts zutrifft
    End If
    
    ' Email erstellen
    With OutlookMail
        .To = empfangsadresse
        .Subject = betreff
        .Body = "Guten Morgen," & vbNewLine & vbNewLine & _
                "wir haben eine Anfrage über " & stueck & " Stück " & bezeichnung & " im Hause." & vbNewLine & _
                "Material ist " & material & ". Materialbeistellung " & Materialbeistellung & "." & vbNewLine & vbNewLine & _
                "Können Sie uns hierzu ein Angebot anbieten?" & vbNewLine & vbNewLine & _
                "Vielen Dank vorab!" & vbNewLine & _
                "Mit freundlichen Grüßen" & vbNewLine & vbNewLine & _
                "Paul Schulze"
        .Display ' E-Mail wird geöffnet, aber nicht gesendet
    End With
    
    ' Aufräumen
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub


