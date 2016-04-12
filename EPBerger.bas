Attribute VB_Name = "Berger"
Option Explicit

Const opSaveHTML As Boolean = True      'EXPORT DES FICHIERS EN HTLL
Const opDelName As Boolean = True       'EFFACE LE NOM DU JOUEUR DE LA FEUILLE
Const opOrderByResult As Boolean = True 'TRIE PAR RESULTAT
Const opAddLink As Boolean = True       'AJOUTE LES LIENS VERS LES AUTRES FEUILLES
Const opAddLichessLink As Boolean = True 'AJOUTE LE LIEN VERS LE PROFIL LICHESS



Sub ExtraireParcoursBerger()

Dim nbJoueur As Integer
Dim i As Integer
Dim l As Integer
Dim NumJoueur As Integer
Dim ligneDebut As Integer
Dim ligneFin As Integer
Dim NomJoueur As String
Dim nbRondeGagnee As Integer
Dim nbRondeNulle As Integer
Dim nbRondePerdue As Integer
Dim nbRondeRestante As Integer
Dim Pourcentage As Integer
Dim Score As Single
Dim StrResult As String

    If MsgBox("Voulez vous exécuter la macro ?" & vbCrLf & _
             vbCrLf & _
            "Export en HTML : " & IIf(opSaveHTML, "OUI vers " & ActiveWorkbook.Path, "NON") & vbCrLf & vbCrLf & _
            "Liens Fiches : " & IIf(opAddLink, "OUI", "NON") & vbCrLf & _
            "Liens Lichess : " & IIf(opAddLichessLink, "OUI", "NON") & vbCrLf & _
            "" & vbCrLf _
            , vbExclamation + vbOKCancel, "Génération des parcours") = vbCancel Then
            Exit Sub
    End If
    
        
            


    'EFFACE LES DEUX PREMIERES LIGNES
    Rows("1:2").Select
    Range("A2").Activate
    Selection.Delete Shift:=xlUp

    
    'CALCUL LE NOMBRE DE JOUEUR
    i = 2
    While Range("A" & i).Text <> "RONDE 2"
        i = i + 1
    Wend
    nbJoueur = (i - 2) * 2
    
    'EFFACE LES LIGNES CONTENANT "RONDE X"
    For i = 0 To nbJoueur - 1
        Rows(i * nbJoueur / 2 + 1).Select
        Selection.Delete Shift:=xlUp
    Next i
    
    'INSERE UNE COLONNE
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    'ECRIT LE NUMERO DE LA RONDE DEVANT CHAQUE PARTIE
    For i = 0 To nbJoueur - 2
        For l = 1 To nbJoueur / 2
            Range("A" & i * nbJoueur / 2 + l).Select
            If i + 1 < 10 Then
            ActiveCell.FormulaR1C1 = "RONDE 0" & i + 1
            Else
                ActiveCell.FormulaR1C1 = "RONDE " & i + 1
            End If
        Next l
    Next i
    
    'SELECTIONNE TOUTES LES LIGNES
    Range("A1:E" & ((nbJoueur - 1) * nbJoueur / 2)).Select
    
    'TRIE PAR JOUEURS BLANCS
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:=Range( _
        "B1:B" & ((nbJoueur - 1) * nbJoueur / 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:=Range( _
        "A1:A" & ((nbJoueur - 1) * nbJoueur / 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets(1).Sort
        .SetRange Range("A1:E" & ((nbJoueur - 1) * nbJoueur / 2))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'PLACE LE SCORE POUR LE JOUEUR BLANC DANS LA COLONNE E
    For i = 1 To nbJoueur / 2 * (nbJoueur - 1)
        StrResult = Range("C" & i).Text
        Range("E" & i) = -1
        If InStr(1, StrResult, "0 - 1") > 0 Then
            Range("E" & i) = 0
        End If
        If InStr(1, StrResult, "1 - 0") > 0 Then
            Range("E" & i) = 1
        End If
        If InStr(1, StrResult, "X - X") > 0 Then
            Range("E" & i) = 0.5
        End If
    Next i
    
    
    'DUPLIQUE L'ENSEMBLE DES LIGNES
    Selection.Copy
    Range("A" & ((nbJoueur - 1) * nbJoueur / 2) + 1).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
    'TRIE PAR JOUEURS NOIRS
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:=Range( _
        "D" & ((nbJoueur - 1) * nbJoueur / 2 + 1) & ":D" & ((nbJoueur - 1) * nbJoueur)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets(1).Sort.SortFields.Add Key:=Range( _
        "A" & ((nbJoueur - 1) * nbJoueur / 2 + 1) & ":A" & ((nbJoueur - 1) * nbJoueur)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets(1).Sort
        .SetRange Range("A" & ((nbJoueur - 1) * nbJoueur / 2 + 1) & " :E" & ((nbJoueur - 1) * nbJoueur))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'PLACE LE SCORE POUR LE JOUEUR NOIR DANS LA COLONNE E
    For i = nbJoueur / 2 * (nbJoueur - 1) + 1 To 2 * (nbJoueur / 2 * (nbJoueur - 1))
        StrResult = Range("C" & i).Text
        Range("E" & i) = -1
        If InStr(1, StrResult, "0 - 1") > 0 Then
            Range("E" & i) = 1
        End If
        If InStr(1, StrResult, "1 - 0") > 0 Then
            Range("E" & i) = 0
        End If
        If InStr(1, StrResult, "X - X") > 0 Then
            Range("E" & i) = 0.5
        End If
    Next i
    
    'COPIE LES RONDES PAR JOUEUR BLANC DANS DES NOUVELLES FEUILLES
    NumJoueur = 1
    ligneDebut = 1
    ligneFin = 1
    
    While NumJoueur <= CInt(nbJoueur)
        NomJoueur = Range("B" & ligneDebut).Text
        While Range("B" & ligneFin).Text = NomJoueur
            ligneFin = ligneFin + 1
        Wend
        ligneFin = ligneFin - 1
        Range("A" & ligneDebut & ":E" & ligneFin).Select
        Selection.Copy
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(NumJoueur + 1).Select
        Sheets(NumJoueur + 1).Name = NomJoueur
        Range("A1").Select
        ActiveSheet.Paste
        Sheets(1).Select
        ligneDebut = ligneFin + 1
        ligneFin = ligneFin + 1
        NumJoueur = NumJoueur + 1
    Wend
    
    
    
    
    'COPIE LES RONDES PAR JOUEUR NOIR DANS DES NOUVELLES FEUILLES
    NumJoueur = 1
    While NumJoueur <= CInt(nbJoueur)
        NomJoueur = Range("D" & ligneDebut).Text
        While Range("D" & ligneFin).Text = NomJoueur
            ligneFin = ligneFin + 1
        Wend
        ligneFin = ligneFin - 1
        Range("A" & ligneDebut & ":E" & ligneFin).Select
        Selection.Copy
        Sheets(NumJoueur + 1).Select
        Range("A" & nbJoueur).Select
        ActiveSheet.Paste
        
        
        If opDelName Then
            'EFFACE LE NOM DU JOUEUR BLANC
            Range("B1:B" & nbJoueur - 1).Clear
            
            'EFFACE LE NOM DU JOUEUR NOIR
            Range("D" & nbJoueur & ":D" & nbJoueur * 2).Clear
        End If
        
        'TRIE PAR RONDE
        Range("A1:E" & nbJoueur * 2).Select
        ActiveWorkbook.Worksheets(NumJoueur + 1).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(NumJoueur + 1).Sort.SortFields.Add Key:=Range( _
            "A1:A" & nbJoueur * 2), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With ActiveWorkbook.Worksheets(NumJoueur + 1).Sort
            .SetRange Range("A1:E" & nbJoueur * 2)
            .Header = xlGuess
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        If opOrderByResult Then
        
             'TRIE PAR RESULTAT
             ActiveWorkbook.Worksheets(NumJoueur + 1).Sort.SortFields.Clear
             ActiveWorkbook.Worksheets(NumJoueur + 1).Sort.SortFields.Add Key:=Range( _
                 "E1:E" & nbJoueur - 1), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
                 xlSortNormal
             With ActiveWorkbook.Worksheets(NumJoueur + 1).Sort
                 .SetRange Range("A1:E" & nbJoueur - 1)
                 .Header = xlGuess
                 .MatchCase = False
                 .Orientation = xlTopToBottom
                 .SortMethod = xlPinYin
                 .Apply
             End With
             
             Range("E" & nbJoueur) = -1
             
             'COMPTE LE NOMBRE DE RONDE GAGNEE
             nbRondeGagnee = 0
             ligneDebut = 1
             If Range("E1").Text = 1 Then
                 While Range("E" & ligneDebut) = 1
                     ligneDebut = ligneDebut + 1
                 Wend
                 nbRondeGagnee = ligneDebut - 1
             End If
             
             'COMPTE LE NOMBRE DE RONDE NULLE
             nbRondeNulle = 0
             If Range("E" & ligneDebut).Text = 0.5 Then
                 While Range("E" & ligneDebut) = 0.5
                     ligneDebut = ligneDebut + 1
                 Wend
                 nbRondeNulle = ligneDebut - nbRondeGagnee - 1
             End If
             
             'COMPTE LE NOMBRE DE RONDE PERDUE
             nbRondePerdue = 0
             If Range("E" & ligneDebut).Text = 0 Then
                 While Range("E" & ligneDebut) = 0
                     ligneDebut = ligneDebut + 1
                 Wend
                 nbRondePerdue = ligneDebut - nbRondeGagnee - nbRondeNulle - 1
             End If
             
             'COMPTE LE NOMBRE DE RONDE RESTANTE
             nbRondeRestante = nbJoueur - 1 - nbRondeGagnee - nbRondeNulle - nbRondePerdue
             
             'EFFACE LA COLONNE E
             Columns("E:E").Select
             Selection.Delete Shift:=xlToLeft
             
             
             'AJOUTE LA LIGNE "RONDES RESTANTES"
             If nbRondeRestante > 0 Then
             ligneDebut = (nbJoueur - 1) - nbRondeRestante + 1
                 Rows(ligneDebut & ":" & ligneDebut).Select
                 Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                 Range("A" & ligneDebut & ":D" & ligneDebut).Select
                 With Selection
                     .HorizontalAlignment = xlCenter
                     .VerticalAlignment = xlCenter
                     .WrapText = False
                     .Orientation = 0
                     .AddIndent = False
                     .IndentLevel = 0
                     .ShrinkToFit = False
                     .ReadingOrder = xlContext
                     .MergeCells = False
                 End With
                 Selection.Merge
                 Range("A" & ligneDebut & ":D" & ligneDebut).Select
                 Selection.Style = "Accent1"
                 Selection.Font.Bold = True
                 If nbRondeRestante = 1 Then
                    ActiveCell.FormulaR1C1 = nbRondeRestante & " RONDE RESTANTE"
                 Else
                    ActiveCell.FormulaR1C1 = nbRondeRestante & " RONDES RESTANTES"
                 End If

            End If
             
             'AJOUTE LA LIGNE "RONDES PERDUES"
             If nbRondePerdue > 0 Then
                 ligneDebut = (nbJoueur - 1) - nbRondeRestante - nbRondePerdue + 1
                 Rows(ligneDebut & ":" & ligneDebut).Select
                 Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                 Range("A" & ligneDebut & ":D" & ligneDebut).Select
                 With Selection
                     .HorizontalAlignment = xlCenter
                     .VerticalAlignment = xlCenter
                     .WrapText = False
                     .Orientation = 0
                     .AddIndent = False
                     .IndentLevel = 0
                     .ShrinkToFit = False
                     .ReadingOrder = xlContext
                     .MergeCells = False
                 End With
                 Selection.Merge
                 Range("A" & ligneDebut & ":D" & ligneDebut).Select
                 Selection.Style = "Insatisfaisant"
                 Selection.Font.Bold = True
                 
                 If nbRondePerdue = 1 Then
                    ActiveCell.FormulaR1C1 = nbRondePerdue & " RONDE PERDUE"
                 Else
                    ActiveCell.FormulaR1C1 = nbRondePerdue & " RONDES PERDUES"
                 End If
                 
             End If
             
             'AJOUTE LA LIGNE "RONDES NULLES"
             If nbRondeNulle > 0 Then
                 ligneDebut = (nbJoueur - 1) - nbRondeRestante - nbRondePerdue - nbRondeNulle + 1
                 Rows(ligneDebut & ":" & ligneDebut).Select
                 Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                 Range("A" & ligneDebut & ":D" & ligneDebut).Select
                 With Selection
                     .HorizontalAlignment = xlCenter
                     .VerticalAlignment = xlCenter
                     .WrapText = False
                     .Orientation = 0
                     .AddIndent = False
                     .IndentLevel = 0
                     .ShrinkToFit = False
                     .ReadingOrder = xlContext
                     .MergeCells = False
                 End With
                 Selection.Merge
                 Range("A" & ligneDebut & ":D" & ligneDebut).Select
                 Selection.Style = "Neutre"
                 Selection.Font.Bold = True
                 
                 If nbRondeNulle = 1 Then
                    ActiveCell.FormulaR1C1 = nbRondeNulle & " RONDE NULLE"
                 Else
                    ActiveCell.FormulaR1C1 = nbRondeNulle & " RONDES NULLES"
                 End If
             End If
             
             'AJOUTE LA LIGNE "RONDES GAGNEES"
             If nbRondeGagnee > 0 Then
                 ligneDebut = 1
                 Rows(ligneDebut & ":" & ligneDebut).Select
                 Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                 Range("A" & ligneDebut & ":D" & ligneDebut).Select
                 With Selection
                     .HorizontalAlignment = xlCenter
                     .VerticalAlignment = xlCenter
                     .WrapText = False
                     .Orientation = 0
                     .AddIndent = False
                     .IndentLevel = 0
                     .ShrinkToFit = False
                     .ReadingOrder = xlContext
                     .MergeCells = False
                 End With
                 Selection.Merge
                 Range("A" & ligneDebut & ":D" & ligneDebut).Select
                 Selection.Style = "Satisfaisant"
                 Selection.Font.Bold = True
                 
                 If nbRondeGagnee = 1 Then
                    ActiveCell.FormulaR1C1 = nbRondeGagnee & " RONDE GAGNEE"
                 Else
                    ActiveCell.FormulaR1C1 = nbRondeGagnee & " RONDES GAGNEES"
                 End If
             End If
       
        
            'AJOUTE LA LIGNE TITRE
        
            ligneDebut = 1
            Rows(ligneDebut & ":" & ligneDebut).Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("A" & ligneDebut & ":D" & ligneDebut).Select
            With Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            Selection.Merge
            Range("A" & ligneDebut & ":D" & ligneDebut).Select
            
            'AJOUTE LE LIEN VERS LE PROFIL LICHESS
            If opAddLichessLink Then
                ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="http://fr.lichess.org/@/" & Sheets(NumJoueur + 1).Name
            End If
            
            Selection.Style = "Commentaire"
            With Selection.Font
                .Bold = True
                .Size = 18
                .ColorIndex = xlAutomatic
                .Underline = xlUnderlineStyleNone
            End With
            Score = nbRondeGagnee + nbRondeNulle / 2
            If (nbJoueur - 1 - nbRondeRestante) > 0 Then
                Pourcentage = CInt(Score / (nbJoueur - 1 - nbRondeRestante) * 100)
            Else
                Pourcentage = 0
            End If
            
            ActiveCell.FormulaR1C1 = Sheets(NumJoueur + 1).Name & _
                              IIf(nbRondeGagnee > 0, " +" & nbRondeGagnee, "") & _
                              IIf(nbRondeNulle > 0, " =" & nbRondeNulle, "") & _
                            IIf(nbRondePerdue > 0, " -" & nbRondePerdue, "") & _
                            IIf(nbRondeRestante > 0, " #" & nbRondeRestante, "") & _
                            IIf(Score > 0, " " & "." & " " & Score & "/" & nbJoueur - 1 - nbRondeRestante & _
                            " " & "." & " " & Pourcentage & "%", "")
         End If
        
        'ADAPTE LA LARGEUR
        Columns("A:A").ColumnWidth = 10
        Columns("B:B").ColumnWidth = 25
        Columns("C:C").ColumnWidth = 6
        Columns("D:D").ColumnWidth = 25
        
        'ADAPTE LA HAUTEUR
        For i = 1 To nbJoueur * 2
            Rows(i & ":" & i).RowHeight = 20
        Next i
        
        'AJOUTE DES LIENS SUR LES NOMS
        If opAddLink Then
            For i = 3 To nbJoueur + 5
                If Range("B" & i).Text <> "" And Range("D" & i).Text = "" Then
                    Range("B" & i).Select
                    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=Range("B" & i).Text & ".html", TextToDisplay:=Range("B" & i).Text
                End If
                If Range("B" & i).Text = "" And Range("D" & i).Text <> "" Then
                    Range("D" & i).Select
                    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=Range("D" & i).Text & ".html", TextToDisplay:=Range("D" & i).Text
                End If
                With Selection.Font
                    .ColorIndex = xlAutomatic
                    .TintAndShade = 0
                End With
                Selection.Font.Underline = xlUnderlineStyleNone
            Next i
        End If
        
        If opSaveHTML Then
            'SAUVEGARDE EN HTML
            With ActiveWorkbook.PublishObjects.Add(xlSourceSheet, _
                 ActiveWorkbook.Path & "\" & NomJoueur & ".html", _
                NomJoueur, "", xlHtmlStatic, NomJoueur, "")
                .Publish (True) 'ATTENTION ERREUR SI LES DROIT D ECRITURE WINDOWS NE SONT PAS CORRECT
                .AutoRepublish = False
            End With
        End If
        
        'ON PASSE AU JOUEUR SUIVANT
        Sheets(1).Select
        ligneDebut = ligneFin + 1
        ligneFin = ligneFin + 1
        NumJoueur = NumJoueur + 1
        
    Wend
    
    

End Sub





