Attribute VB_Name = "Saisie_macros"
Option Explicit

'==============================================================================
' SAISIE RAPIDE - Journée agent
' Ajoute une ligne dans l'onglet "Justificatifs" à partir de l'onglet "Saisie".
' Les boutons et le raccourci Ctrl+Maj+A sont créés automatiquement à l'ouverture.
'==============================================================================

Private Const FEUILLE_SAISIE As String = "Saisie"
Private Const FEUILLE_JUSTIF As String = "Justificatifs"
Private Const LIGNE_DEBUT As Long = 5
Private Const LIGNE_FIN As Long = 5004

'------------------------------------------------------------------------------
' Ajoute la ligne saisie dans l'onglet Justificatifs
'------------------------------------------------------------------------------
Public Sub AjouterLigne()
    Dim sh As Worksheet, jf As Worksheet
    Dim ctrl As String, msg As String, r As Long
    Dim ancienCalcul As Long

    ancienCalcul = Application.Calculation
    On Error GoTo Erreur

    Set sh = ThisWorkbook.Worksheets(FEUILLE_SAISIE)
    Set jf = ThisWorkbook.Worksheets(FEUILLE_JUSTIF)

    If ActiveSheet.Name <> FEUILLE_SAISIE Then
        sh.Activate
        MsgBox "Remplissez d'abord la feuille « Saisie », puis relancez.", vbInformation, "Ajouter une ligne"
        Exit Sub
    End If

    sh.Calculate
    ctrl = CStr(sh.Range("F8").Value)
    If ctrl <> "OK" Then
        If Left$(ctrl, 13) = "Prix manquant" Then
            If MsgBox(ctrl & vbCrLf & vbCrLf & "Ajouter quand même la ligne (montant 0 " & ChrW(8364) & ") ?", _
                      vbYesNo + vbQuestion, "Ajouter une ligne") <> vbYes Then Exit Sub
        Else
            MsgBox "La ligne n'a pas été ajoutée :" & vbCrLf & ctrl, vbExclamation, "Ajouter une ligne"
            Exit Sub
        End If
    End If

    r = ProchaineLigneLibre(jf)
    If r > LIGNE_FIN Then
        MsgBox "L'onglet Justificatifs est plein.", vbExclamation, "Ajouter une ligne"
        Exit Sub
    End If

    ' message de confirmation (construit avant d'effacer les champs)
    msg = "Ligne ajoutée le " & Format$(Int(sh.Range("C5").Value2), "dd/mm/yyyy") & " : " & _
          sh.Range("C6").Value & " · " & sh.Range("C7").Value
    If Len(CStr(sh.Range("C9").Value)) > 0 Then msg = msg & " · quantité " & sh.Range("C9").Value
    If IsNumeric(sh.Range("F7").Value) Then
        msg = msg & " · " & Format$(sh.Range("F7").Value, "#,##0.00") & " " & ChrW(8364)
    End If

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    jf.Cells(r, 1).Value2 = Int(sh.Range("C5").Value2)
    jf.Cells(r, 2).Value = sh.Range("C6").Value
    jf.Cells(r, 3).Value = sh.Range("C7").Value
    If Val(CStr(sh.Range("I5").Value)) > 1 Then
        jf.Cells(r, 4).Value = sh.Range("C8").Value
    Else
        jf.Cells(r, 4).ClearContents
    End If
    If Len(CStr(sh.Range("C9").Value)) > 0 Then
        jf.Cells(r, 5).Value2 = sh.Range("C9").Value2
    Else
        jf.Cells(r, 5).ClearContents
    End If
    jf.Cells(r, 11).Value = sh.Range("C10").Value
    jf.Cells(r, 12).Value = sh.Range("C11").Value

    ' mémorise la ligne (pour pouvoir l'annuler) et prépare la saisie suivante
    sh.Range("I1").Value = r
    sh.Range("C7").ClearContents
    sh.Range("C8").ClearContents
    sh.Range("C9").Value = 1
    sh.Range("C10").ClearContents
    sh.Range("C11").ClearContents
    sh.Range("B16").Value = msg & "   (ligne " & r & " de Justificatifs)"

    Application.Calculation = ancienCalcul
    Application.ScreenUpdating = True
    sh.Range("C7").Select
    Exit Sub

Erreur:
    Application.ScreenUpdating = True
    Application.Calculation = ancienCalcul
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, "Ajouter une ligne"
End Sub

'------------------------------------------------------------------------------
' Supprime la dernière ligne ajoutée avec AjouterLigne
'------------------------------------------------------------------------------
Public Sub AnnulerDerniereLigne()
    Dim sh As Worksheet, jf As Worksheet
    Dim r As Long, txt As String, v As Variant

    Set sh = ThisWorkbook.Worksheets(FEUILLE_SAISIE)
    Set jf = ThisWorkbook.Worksheets(FEUILLE_JUSTIF)

    v = sh.Range("I1").Value
    If Len(CStr(v)) = 0 Then
        MsgBox "Aucune ligne à annuler (rien n'a été ajouté avec le bouton).", vbInformation, "Annuler la dernière ligne"
        Exit Sub
    End If
    r = CLng(v)
    If r < LIGNE_DEBUT Or r > LIGNE_FIN Then Exit Sub

    If Len(CStr(jf.Cells(r, 1).Value)) = 0 Then
        MsgBox "La dernière ligne ajoutée a déjà été effacée.", vbInformation, "Annuler la dernière ligne"
        sh.Range("I1").ClearContents
        Exit Sub
    End If

    txt = Format$(jf.Cells(r, 1).Value, "dd/mm/yyyy") & " · " & jf.Cells(r, 2).Value & " · " & jf.Cells(r, 3).Value
    If Len(CStr(jf.Cells(r, 4).Value)) > 0 Then txt = txt & " · " & jf.Cells(r, 4).Value
    If MsgBox("Supprimer la dernière ligne ajoutée ?" & vbCrLf & vbCrLf & txt, _
              vbYesNo + vbQuestion, "Annuler la dernière ligne") <> vbYes Then Exit Sub

    jf.Range(jf.Cells(r, 1), jf.Cells(r, 5)).ClearContents
    jf.Cells(r, 11).ClearContents
    jf.Cells(r, 12).ClearContents
    sh.Range("I1").ClearContents
    sh.Range("B16").Value = "Ligne annulée : " & txt
End Sub

'------------------------------------------------------------------------------
' Première ligne libre de l'onglet Justificatifs (colonnes Date / Technicien / Projet)
'------------------------------------------------------------------------------
Private Function ProchaineLigneLibre(jf As Worksheet) As Long
    Dim c As Long, r As Long, m As Long
    m = LIGNE_DEBUT - 1
    For c = 1 To 3
        r = jf.Cells(jf.Rows.Count, c).End(xlUp).Row
        If r > m Then m = r
    Next c
    ProchaineLigneLibre = m + 1
End Function

'==============================================================================
' À l'ouverture : raccourci Ctrl+Maj+A, date du jour, boutons sur la feuille Saisie
'==============================================================================
Public Sub Auto_Open()
    On Error Resume Next
    Application.OnKey "^+a", "AjouterLigne"
    Application.OnKey "^+s", "SynchroniserWeb"
    ThisWorkbook.Worksheets(FEUILLE_SAISIE).Range("C5").Formula = "=TODAY()"
    CreerBoutons
    ThisWorkbook.Worksheets(FEUILLE_SAISIE).Activate
    SynchroAuto
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Application.OnKey "^+a"
    Application.OnKey "^+s"
    ArreterSynchro
End Sub

' Crée les deux boutons sur la feuille Saisie (sans doublon). Peut aussi être lancée à la main.
Public Sub CreerBoutons()
    Dim sh As Worksheet, b As Button, c As Range
    Dim aAjouter As Boolean, aAnnuler As Boolean, aSynchro As Boolean

    Set sh = ThisWorkbook.Worksheets(FEUILLE_SAISIE)
    For Each b In sh.Buttons
        If b.Name = "btnAjouter" Then aAjouter = True
        If b.Name = "btnAnnuler" Then aAnnuler = True
        If b.Name = "btnSynchro" Then aSynchro = True
    Next b

    If Not aAjouter Then
        Set c = sh.Range("B13:C14")
        Set b = sh.Buttons.Add(c.Left + 2, c.Top + 2, c.Width - 4, c.Height - 4)
        b.Name = "btnAjouter"
        b.Caption = "AJOUTER LA LIGNE   (Ctrl+Maj+A)"
        b.OnAction = "AjouterLigne"
        b.Font.Bold = True
        b.Font.Size = 12
    End If

    If Not aAnnuler Then
        Set c = sh.Range("D13:F14")
        Set b = sh.Buttons.Add(c.Left + 6, c.Top + 2, 230, c.Height - 4)
        b.Name = "btnAnnuler"
        b.Caption = "Annuler la dernière ligne"
        b.OnAction = "AnnulerDerniereLigne"
        b.Font.Size = 10
    End If

    If Not aSynchro Then
        Set c = sh.Range("G13:H14")
        Set b = sh.Buttons.Add(c.Left + 6, c.Top + 2, 230, c.Height - 4)
        b.Name = "btnSynchro"
        b.Caption = "Synchroniser le web   (Ctrl+Maj+S)"
        b.OnAction = "SynchroniserWeb"
        b.Font.Size = 10
    End If
End Sub

'==============================================================================
' SYNCHRO WEB - récupère les saisies faites par les techniciens sur la page web
' (Google Sheet « Journée agent ») et les ajoute dans l'onglet Justificatifs.
' - Seules les colonnes Date, Technicien, Projet, Prestation, Quantité,
'   N° ticket et Commentaire sont écrites : les colonnes automatiques gardent
'   leurs formules. Prix sur devis : PU et montant écrits en orange.
' - Chaque ligne venue du web porte son identifiant en colonne U : elle n'est
'   jamais ajoutée deux fois, et elle est retirée si elle a été supprimée sur le web.
' - Les lignes saisies directement dans Excel (colonne U vide) ne sont jamais touchées.
'==============================================================================

Private Const COL_ID As Long = 21                 ' colonne U de Justificatifs
Private Const NOM_LIEN As String = "LienSaisieWeb"
Private Const MINUTES_SYNCHRO As Long = 5
Private prochaineSynchro As Date

' Bouton « Synchroniser » et raccourci Ctrl+Maj+S
Public Sub SynchroniserWeb()
    Synchroniser False
End Sub

' Synchro silencieuse : à l'ouverture puis toutes les 5 minutes
Public Sub SynchroAuto()
    If Len(LienWeb()) > 0 Then Synchroniser True
    ProgrammerSynchro
End Sub

Private Sub ProgrammerSynchro()
    On Error Resume Next
    prochaineSynchro = Now + TimeSerial(0, MINUTES_SYNCHRO, 0)
    Application.OnTime prochaineSynchro, "SynchroAuto"
End Sub

Private Sub ArreterSynchro()
    On Error Resume Next
    If prochaineSynchro > 0 Then Application.OnTime prochaineSynchro, "SynchroAuto", , False
End Sub

' Enregistre (ou remplace) le lien de la page web
Public Sub ChangerLienWeb()
    Dim url As String
    url = Trim$(InputBox("Collez le lien de la page de saisie (il finit par /exec) :", "Lien de la page web", LienWeb()))
    If Len(url) = 0 Then Exit Sub
    If InStr(1, url, "script.google.com", vbTextCompare) = 0 Then
        MsgBox "Ce lien ne ressemble pas à celui de la page de saisie (script.google.com/.../exec).", vbExclamation, "Lien de la page web"
        Exit Sub
    End If
    On Error Resume Next
    ThisWorkbook.Names(NOM_LIEN).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=NOM_LIEN, RefersTo:="=""" & Replace(url, """", "") & """", Visible:=False
End Sub

Private Function LienWeb() As String
    On Error Resume Next
    LienWeb = CStr(Evaluate(ThisWorkbook.Names(NOM_LIEN).RefersTo))
    If Err.Number <> 0 Then LienWeb = ""
End Function

Private Sub Synchroniser(silencieux As Boolean)
    Dim jf As Worksheet, url As String, texte As String
    Dim lignes() As String, f() As String, i As Long, r As Long
    Dim distant As Object, local_ As Object, k As Variant
    Dim nAjout As Long, nRetrait As Long, ancienCalcul As Long

    If Len(LienWeb()) = 0 Then
        If silencieux Then Exit Sub
        ChangerLienWeb
        If Len(LienWeb()) = 0 Then Exit Sub
    End If

    ancienCalcul = Application.Calculation
    On Error GoTo Erreur
    Set jf = ThisWorkbook.Worksheets(FEUILLE_JUSTIF)
    If Len(CStr(jf.Cells(4, COL_ID).Value)) = 0 Then jf.Cells(4, COL_ID).Value = "ID web (ne pas modifier)"

    url = LienWeb()
    url = url & IIf(InStr(url, "?") > 0, "&", "?") & "export=tsv"
    texte = Telecharger(url)
    If Left$(texte, 3) <> "ID" & vbTab Then
        If Not silencieux Then MsgBox "La page web n'a pas renvoyé les saisies." & vbCrLf & _
            "Vérifiez le lien (macro ChangerLienWeb) et la connexion Internet.", vbExclamation, "Synchroniser"
        Exit Sub
    End If

    ' saisies présentes sur le web : identifiant -> champs
    Set distant = CreateObject("Scripting.Dictionary")
    lignes = Split(Replace(texte, vbCr, ""), vbLf)
    For i = 1 To UBound(lignes)
        If Len(lignes(i)) > 0 Then
            f = Split(lignes(i), vbTab)
            If UBound(f) >= 9 Then distant(f(0)) = f
        End If
    Next i

    ' lignes déjà venues du web dans Justificatifs : identifiant -> n° de ligne
    Set local_ = CreateObject("Scripting.Dictionary")
    For r = LIGNE_DEBUT To LIGNE_FIN
        k = CStr(jf.Cells(r, COL_ID).Value)
        If Len(k) > 0 Then local_(k) = r
    Next r

    ' si le web ne renvoie plus rien alors qu'Excel a des lignes venues du web, on demande avant de retirer
    If distant.Count = 0 And local_.Count > 0 Then
        If silencieux Then Exit Sub
        If MsgBox("La page web ne contient plus aucune saisie." & vbCrLf & _
                  "Retirer les " & local_.Count & " lignes venues du web de l'onglet Justificatifs ?", _
                  vbYesNo + vbQuestion, "Synchroniser") <> vbYes Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' 1. retire les lignes supprimées sur le web
    For Each k In local_.Keys
        If Not distant.Exists(k) Then
            ViderLigne jf, CLng(local_(k))
            nRetrait = nRetrait + 1
        End If
    Next k

    ' 2. ajoute les nouvelles saisies
    r = LIGNE_DEBUT
    For Each k In distant.Keys
        If Not local_.Exists(k) Then
            Do While r <= LIGNE_FIN
                If Len(CStr(jf.Cells(r, 1).Value)) = 0 And Len(CStr(jf.Cells(r, 2).Value)) = 0 _
                   And Len(CStr(jf.Cells(r, 3).Value)) = 0 And Len(CStr(jf.Cells(r, COL_ID).Value)) = 0 Then Exit Do
                r = r + 1
            Loop
            If r > LIGNE_FIN Then
                Application.Calculation = ancienCalcul
                Application.ScreenUpdating = True
                MsgBox "L'onglet Justificatifs est plein : " & nAjout & " lignes ajoutées seulement.", vbExclamation, "Synchroniser"
                Exit Sub
            End If
            EcrireLigne jf, r, distant(k)
            nAjout = nAjout + 1
            r = r + 1
        End If
    Next k

    Application.Calculation = ancienCalcul
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(FEUILLE_SAISIE).Range("B16").Value = "Synchro web du " & Format$(Now, "dd/mm/yyyy à hh:nn") & _
        " : " & nAjout & " ligne(s) ajoutée(s), " & nRetrait & " retirée(s)."
    If Not silencieux Then MsgBox nAjout & " ligne(s) ajoutée(s), " & nRetrait & " retirée(s).", vbInformation, "Synchroniser"
    Exit Sub

Erreur:
    Application.Calculation = ancienCalcul
    Application.ScreenUpdating = True
    If Not silencieux Then MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, "Synchroniser"
End Sub

' f : ID, Date (aaaa-mm-jj), Technicien, Projet, Prestation, Quantité, Ticket, Commentaire, PU, Montant
Private Sub EcrireLigne(jf As Worksheet, r As Long, f As Variant)
    Dim q As Double, pu As Double, montant As Double
    jf.Cells(r, 1).Value2 = CDbl(DateSerial(CInt(Left$(f(1), 4)), CInt(Mid$(f(1), 6, 2)), CInt(Mid$(f(1), 9, 2))))
    jf.Cells(r, 2).Value = Texte(f(2))
    jf.Cells(r, 3).Value = Texte(f(3))
    If Len(f(4)) > 0 Then jf.Cells(r, 4).Value = Texte(f(4)) Else jf.Cells(r, 4).ClearContents
    q = Val(f(5))
    If q = 1 Or Len(f(5)) = 0 Then jf.Cells(r, 5).ClearContents Else jf.Cells(r, 5).Value2 = q
    If Len(f(6)) > 0 Then jf.Cells(r, 11).Value = Texte(f(6)) Else jf.Cells(r, 11).ClearContents
    If Len(f(7)) > 0 Then jf.Cells(r, 12).Value = Texte(f(7)) Else jf.Cells(r, 12).ClearContents
    jf.Cells(r, COL_ID).Value = Texte(f(0))

    ' prix sur devis (absent du BPU Excel) : on reprend le prix tapé sur le web
    pu = Val(f(8)): montant = Val(f(9))
    jf.Range(jf.Cells(r, 1), jf.Cells(r, 20)).Calculate
    If pu > 0 And Val(CStr(jf.Cells(r, 9).Value)) = 0 Then
        jf.Cells(r, 9).Value2 = pu
        jf.Cells(r, 10).Value2 = montant
        jf.Range(jf.Cells(r, 9), jf.Cells(r, 10)).Interior.Color = RGB(252, 213, 180)
    End If
End Sub

' Efface une ligne venue du web et remet les formules PU / Montant si elles avaient été remplacées
Private Sub ViderLigne(jf As Worksheet, r As Long)
    jf.Range(jf.Cells(r, 1), jf.Cells(r, 5)).ClearContents
    jf.Cells(r, 11).ClearContents
    jf.Cells(r, 12).ClearContents
    jf.Cells(r, COL_ID).ClearContents
    If Not jf.Cells(r, 9).HasFormula Then
        jf.Cells(r, 9).Formula = "=IF($T" & r & "="""","""",IFERROR(N(INDEX(BPU!$F$6:$F$400,$T" & r & ")),0))"
        jf.Cells(r, 10).Formula = "=IF($T" & r & "="""","""",IF($E" & r & "="""",1,N($E" & r & "))*$I" & r & ")"
        jf.Range(jf.Cells(r, 9), jf.Cells(r, 10)).Interior.Pattern = xlNone
    End If
End Sub

' Un texte qui commence par = + - @ ne doit pas devenir une formule
Private Function Texte(v As Variant) As String
    Texte = CStr(v)
    If Len(Texte) > 0 Then
        If InStr("=+-@", Left$(Texte, 1)) > 0 Then Texte = "'" & Texte
    End If
End Function

' Télécharge le texte (UTF-8) d'une adresse web
Private Function Telecharger(url As String) As String
    Dim http As Object, flux As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Option(6) = True                          ' suit les redirections de Google
    http.SetTimeouts 10000, 10000, 30000, 60000
    http.Open "GET", url, False
    http.Send
    If http.Status <> 200 Then Exit Function
    Set flux = CreateObject("ADODB.Stream")
    flux.Type = 1: flux.Open
    flux.Write http.ResponseBody
    flux.Position = 0: flux.Type = 2: flux.Charset = "utf-8"
    Telecharger = flux.ReadText
    flux.Close
End Function
