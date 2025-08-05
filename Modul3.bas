Attribute VB_Name = "Modul3"
Option Explicit

Sub KUNDENABerstellenPDF()

    Dim Quellentabelle As Worksheet
    Dim Zieltabelle As Worksheet
    Dim Zeilennummer As Long
    Dim Bestellnummer As String
    Dim speicherPfad As String
    Dim dateiname As String
    Dim Liefertermin As Variant
    Dim antwort As VbMsgBoxResult

    Set Quellentabelle = ThisWorkbook.Worksheets("Bestellungen")
    Set Zieltabelle = ThisWorkbook.Worksheets("ABKunde")

      Zeilennummer = CLng(InputBox("Bitte geben Sie die Zeilennummer in Spalte A ein:", "Zeilennummer"))

    Zieltabelle.Range("A3").Value = Quellentabelle.Range("D" & Zeilennummer).Value
    Zieltabelle.Range("C10").Value = Quellentabelle.Range("C" & Zeilennummer).Value
    Zieltabelle.Range("C11").Value = Quellentabelle.Range("A" & Zeilennummer).Value
    Zieltabelle.Range("C13").Value = Quellentabelle.Range("E" & Zeilennummer).Value
    Zieltabelle.Range("C14").Value = Quellentabelle.Range("F" & Zeilennummer).Value
    Zieltabelle.Range("A17").Value = Quellentabelle.Range("B" & Zeilennummer).Value
    
    Liefertermin = Quellentabelle.Range("AV" & Zeilennummer).Value
    antwort = MsgBox("Kann der Liefertermin " & Liefertermin & " eingehalten werden? Klicken Sie 'Ja', um den Termin zu übernehmen, oder 'Nein', um ein neues Datum einzugeben.", vbYesNo + vbQuestion, "Liefertermin Abfrage")

    If antwort = vbYes Then

        Zieltabelle.Range("C25").Value = Liefertermin
    Else

        Dim neuerLiefertermin As Variant
        neuerLiefertermin = InputBox("Bitte geben Sie den neuen Liefertermin ein (TT.MM.JJJJ):", "Neuer Liefertermin", Liefertermin)
        
        If IsDate(neuerLiefertermin) Then
            Zieltabelle.Range("C25").Value = CDate(neuerLiefertermin)
        Else
            MsgBox "Ungültiges Datum eingegeben. Der Liefertermin wird nicht übernommen.", vbExclamation, "Fehler"
        End If
    End If
   
    Zieltabelle.Range("B17").Value = Quellentabelle.Range("H" & Zeilennummer).Value
    Zieltabelle.Range("C17").Value = Quellentabelle.Range("I" & Zeilennummer).Value
    Zieltabelle.Range("E17").Value = Quellentabelle.Range("J" & Zeilennummer).Value
    Zieltabelle.Range("F17").Value = Quellentabelle.Range("K" & Zeilennummer).Value
                         
    Bestellnummer = Quellentabelle.Range("A" & Zeilennummer).Value

    speicherPfad = "D:\LWM\14 Zertifizierung\0 Vorlagen\Auftragsbestätigung "
    dateiname = Zieltabelle.Range("C13").Value

    ' PDF erstellen und speichern
    Zieltabelle.ExportAsFixedFormat Type:=xlTypePDF, Filename:=speicherPfad & dateiname & ".pdf"

    MsgBox "Die Bestellanforderung wurde als PDF gespeichert.", vbInformation

End Sub

