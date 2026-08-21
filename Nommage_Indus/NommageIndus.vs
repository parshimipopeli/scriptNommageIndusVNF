Option Explicit

Sub NommageIndus()

    Dim wsSource As Worksheet
    Dim wsNommage As Worksheet
    Dim wsDest As Worksheet

    Dim derniereLigneSource As Long
    Dim derniereLigneNommage As Long

    Dim i As Long
    Dim j As Long
    Dim ligneDest As Long
    Dim id As Long

    Dim nomOuvrage As String
    Dim adresseIP As String

    Dim trouveNom As Range

    Dim codeNomOuvrage As String
    Dim codeF As String
    Dim trigramme As String

    Dim ligneNommage As Long

    Dim dictHotes As Object

    Const POLLER As String = "Poller-PCC_COMPIEGNE"
    Const COMMUNAUTE As String = "PccCom"

    Set wsSource = Sheets("Source_GB_OISE_CANALISEE")
    Set wsNommage = Sheets("NommageBis")
    Set wsDest = Sheets("Destination")

    Set dictHotes = CreateObject("Scripting.Dictionary")

    derniereLigneSource = wsSource.Cells(wsSource.Rows.Count, 2).End(xlUp).Row
    derniereLigneNommage = wsNommage.Cells(wsNommage.Rows.Count, 11).End(xlUp).Row

    ligneDest = 2
    id = 1

    Application.ScreenUpdating = False

    For i = 2 To derniereLigneSource

        nomOuvrage = Trim(wsSource.Cells(i, 2).Value)

        If nomOuvrage <> "" Then

            Set trouveNom = ChercherOuvrage(wsNommage, nomOuvrage)

            If trouveNom Is Nothing Then

                EcrireErreur wsDest, ligneDest, _
                    nomOuvrage & " - Ouvrage non trouvé dans la base"

                ligneDest = ligneDest + 1

            Else

                codeNomOuvrage = wsNommage.Cells(trouveNom.Row, 7).Value
                codeF = wsNommage.Cells(trouveNom.Row, 6).Value
                nomOuvrage = wsNommage.Cells(trouveNom.Row, 8).Value

                trigramme = ObtenirTrigramme(nomOuvrage)

                For j = 6 To 200

                    adresseIP = Trim(wsSource.Cells(i, j).Value)

                    If adresseIP <> "" Then

                        adresseIP = NettoyerIP(adresseIP)

                        ligneNommage = TrouverCorrespondanceIP( _
                                            wsNommage, _
                                            ConstruireCleRecherche(adresseIP), _
                                            derniereLigneNommage)

                        If ligneNommage > 0 Then

                            AjouterEquipement _
                                wsNommage, _
                                wsDest, _
                                ligneNommage, _
                                ligneDest, _
                                id, _
                                nomOuvrage, _
                                adresseIP, _
                                trigramme, _
                                codeF, _
                                codeNomOuvrage, _
                                dictHotes, _
                                POLLER, _
                                COMMUNAUTE

                            ligneDest = ligneDest + 1
                            id = id + 1

                        Else

                            EcrireErreur wsDest, ligneDest, _
                                nomOuvrage & " " & adresseIP & _
                                " - Aucune correspondance IP"

                            ligneDest = ligneDest + 1

                        End If

                    End If

                Next j

            End If

        End If

    Next i

    Application.ScreenUpdating = True

    MsgBox "Traitement terminé. " & _
           ligneDest - 2 & _
           " ligne(s) écrite(s)."

End Sub

Private Function ChercherOuvrage( _
                    ws As Worksheet, _
                    nomOuvrage As String) As Range

    Set ChercherOuvrage = ws.Columns(8).Find( _
                                What:=nomOuvrage, _
                                LookAt:=xlPart, _
                                MatchCase:=False)

End Function

    
Private Function NettoyerIP( _
                    adresseIP As String) As String

    NettoyerIP = Split(adresseIP, "/")(0)

End Function

Private Function ConstruireCleRecherche( _
                    adresseIP As String) As String

    Dim ip() As String

    ip = Split(adresseIP, ".")

    Select Case ip(2)

        Case "10", "11", "12", "13", _
             "14", "17", "18", "19", _
             "24", "25", "30", "255"

            ConstruireCleRecherche = _
                ip(2) & "." & ip(3)

        Case Else

            ConstruireCleRecherche = ip(3)

    End Select

End Function

Private Function TrouverCorrespondanceIP( _
                    wsNommage As Worksheet, _
                    valeurRecherche As String, _
                    derniereLigne As Long) As Long

    Dim k As Long
    Dim v As Variant
    Dim valeur As String

    For k = 3 To derniereLigne

        valeur = Trim(wsNommage.Cells(k, 12).Value)

        For Each v In Split(valeur, "||")

            If StrComp(Trim(v), _
                       valeurRecherche, _
                       vbTextCompare) = 0 Then

                TrouverCorrespondanceIP = k
                Exit Function

            End If

        Next v

    Next k

End Function

Private Function GenererNomHote( _
                    dictHotes As Object, _
                    codeF As String, _
                    codeNomOuvrage As String, _
                    valeurJ As String) As String

    Dim baseNom As String

    baseNom = Format(codeF, "00") & _
              codeNomOuvrage & _
              valeurJ

    If dictHotes.Exists(baseNom) Then

        dictHotes(baseNom) = _
            dictHotes(baseNom) + 1

    Else

        dictHotes.Add baseNom, 1

    End If

    GenererNomHote = _
        baseNom & dictHotes(baseNom)

End Function

Private Function ObtenirCriticite( _
                    dernierOctet As Long) As String

    Select Case dernierOctet

        Case 98
            ObtenirCriticite = "C2"

        Case 101 To 106
            ObtenirCriticite = "C2"

        Case 109 To 117
            ObtenirCriticite = "C2"

        Case 119 To 121
            ObtenirCriticite = "C2"

        Case Else
            ObtenirCriticite = "C1"

    End Select

End Function

Private Sub AjouterEquipement( _
                    wsNommage As Worksheet, _
                    wsDest As Worksheet, _
                    ligneNom As Long, _
                    ligneDest As Long, _
                    id As Long, _
                    nomOuvrage As String, _
                    adresseIP As String, _
                    trigramme As String, _
                    codeF As String, _
                    codeNomOuvrage As String, _
                    dictHotes As Object, _
                    poller As String, _
                    communaute As String)

    Dim nomHote As String
    Dim valeurJ As String
    Dim domaine As String
    Dim systeme As String
    Dim dernierOctet As Long

    valeurJ = wsNommage.Cells(ligneNom, 10).Value
    domaine = wsNommage.Cells(ligneNom, 14).Value
    systeme = wsNommage.Cells(ligneNom, 13).Value

    nomHote = GenererNomHote( _
                    dictHotes, _
                    codeF, _
                    codeNomOuvrage, _
                    valeurJ)

    dernierOctet = CLng(Split(adresseIP, ".")(3))

    With wsDest

        .Cells(ligneDest, 1).Value = id
        .Cells(ligneDest, 2).Value = Date

        .Cells(ligneDest, 5).Value = UCase(nomOuvrage)
        .Cells(ligneDest, 6).Value = UCase(domaine)
        .Cells(ligneDest, 7).Value = UCase(systeme)

        .Cells(ligneDest, 8).Value = UCase(nomHote)
        .Cells(ligneDest, 9).Value = trigramme

        .Cells(ligneDest, 10).Value = _
            UCase(nomHote & "." & trigramme & ".V2I")

        .Cells(ligneDest, 11).Value = adresseIP
        .Cells(ligneDest, 12).Value = ObtenirCriticite(dernierOctet)

        .Cells(ligneDest, 14).Value = poller
        .Cells(ligneDest, 17).Value = communaute

    End With

    With wsDest.Range("A" & ligneDest & ":R" & ligneDest)

        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)

    End With

End Sub

Private Sub EcrireErreur( _
                    wsDest As Worksheet, _
                    ligneDest As Long, _
                    messageErreur As String)

    wsDest.Cells(ligneDest, 8).Value = messageErreur

    With wsDest.Range("A" & ligneDest & ":R" & ligneDest)

        .Font.Bold = True
        .Font.Color = RGB(192, 0, 0)

    End With

End Sub


Function GenererTrigramme(nomOuvrage As String) As String

    Dim mots() As String
    Dim element As Variant
    Dim trigramme As String

    mots = Split(nomOuvrage, " ")

    For Each element In mots

    Select Case LCase(Trim(element))

        Case "ecluse", "écluse", _
             "barrage", "ouvrage", _
             "de", "du", "des", _
             "la", "le", "les", _
             "pcc"

            ' Ignoré

        Case Else

            If LCase(Left(Trim(element), 2)) = "l'" Then

                trigramme = trigramme & UCase(Mid(Trim(element), 3, 3))

            Else

                trigramme = trigramme & UCase(Left(Trim(element), 3))

            End If

    End Select

Next element

    GenererTrigramme = Left(trigramme, 3)
   ' Debug.Print "Résultat final : " & GenererTrigramme

End Function