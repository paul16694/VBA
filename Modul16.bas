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
    
    zeile = InputBox("Bitte geben Sie die Zeilennummer ein:", "Zeilennummer eingeben")
    

    If zeile <= 0 Or zeile > Rows.Count Then
        MsgBox "Ungültige Zeilennummer. Das Makro wird beendet.", vbExclamation
        Exit Sub
    End If
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0) 

    stueck = Sheets("Bestellung").Range("H3").Value
    material = Sheets("Bestellung").Range("G3").Value
    Materialbeistellung = Sheets("Bestellung").Range("I3").Value
    empfanger = Sheets("Bestellung").Range("J3").Value
    betreff = "Preisanfrage " & Sheets("Bestellung").Range("A3").Value
    bezeichnung = Sheets("Bestellung").Range("F3").Value
    

    If empfanger = "Lieferant1" Then
        empfangsadresse = "paul.schulze@online.de"
    ElseIf empfanger = "Lieferant2" Then
        empfangsadresse = "@ddfdfdf"
    Else
        empfangsadresse = "" 
    End If
    

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
        .Display 
    End With
    

    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub


