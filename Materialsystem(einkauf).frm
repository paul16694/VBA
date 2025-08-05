VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Materialsystem"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7965
   OleObjectBlob   =   "Materialsystem(einkauf).frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOpenCalculator_Click()
     Me.Hide  ' Verberge die aktuelle UserForm
    Gewichtrechner1.Show
    
End Sub

Private Sub CommandButton1_Click()
    ' Verberge die aktuelle UserForm (z.B. UserForm4)
    Me.Hide
    
    ' Überprüfe, ob UserForm4 bereits sichtbar ist
    If UserForm4.Visible = False Then
        UserForm4.Show
    End If
End Sub


Private Sub cmdSendEmail_Click()
    ' Variablen für Outlook und Email
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim empfanger As String
    Dim betreff As String
    Dim EmailBody As String
    Dim lieferant As String
    Dim bestelldetails As String
    
    ' Überprüfen, ob ein Lieferant (Option Button) ausgewählt wurde
    If Not (optLotter.Value Or optIBL.Value Or optstappert.Value Or optBP.Value Or optHandelshof.Value) Then
        MsgBox "Bitte wählen Sie einen Lieferanten aus.", vbExclamation
        Exit Sub
    End If
    
    ' Outlook starten
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0) ' 0 steht für olMailItem
    
    ' Lieferanten basierend auf der Auswahl festlegen
    If optLotter.Value Then
        empfanger = "i@example.de"
        lieferant = "Frau Schulze"
    ElseIf optIBL.Value Then
        empfanger = "paul.schulze@online.de"
        lieferant = "Herr Schulze"
    ElseIf optstappert.Value Then
        empfanger = "merzsch@info.de"
        lieferant = "Herr Merz"
    ElseIf optBP.Value Then
        empfanger = "Lieferant1@info.de"
        lieferant = "Blechprofil-Team"
    ElseIf optHandelshof.Value Then
        empfanger = "Lieferant3@info.de"
        lieferant = "Handelshof-Team"
    End If
    
    ' Betreff festlegen
    betreff = "Preisanfrage"
    
    ' Bestelldetails aus dem Textfeld übernehmen
    bestelldetails = Trim(txtRows.Value)
    
    ' Überprüfen, ob das Textfeld leer ist
    If bestelldetails = "" Then
        MsgBox "Bitte geben Sie die Bestelldetails in das Textfeld ein.", vbExclamation
        Exit Sub
    End If
    
    ' Email-Body erstellen
    EmailBody = "Hallo " & lieferant & "," & vbNewLine & vbNewLine & _
                "wir bitten um Preis und Lieferzeit für folgende Position:" & vbNewLine & vbNewLine & _
                bestelldetails & vbNewLine & vbNewLine & _
                "Vielen Dank für Ihre Bemühungen." & vbNewLine
    
    ' Email erstellen
    With OutlookMail
        .To = empfanger
        .Subject = betreff
        .Body = EmailBody
        .Display ' E-Mail wird geöffnet, aber nicht gesendet
    End With
    
    ' Aufräumen
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub



Private Sub UserForm_Click()

End Sub
