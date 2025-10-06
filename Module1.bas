Attribute VB_Name = "Module1"
Option Explicit

Public Sub UpdateDailyData()
    '======================================================================
    '  UpdateDailyData – version 26-06-2025  (clé Ordre + Opération)
    '======================================================================

    '---- Variables générales -------------------------------------------
    Dim wbSource As Workbook, wbDest As Workbook
    Dim ws As Worksheet, destWs As Worksheet
    Dim srcPath As String, srcFullPath As String
    Dim expectedFile As String, fileName As String, latestFile As String
    Dim latestDate As Date, fileDate As Date
    Dim srcFileName As String, sheetName As String
    Dim lastRow As Long, destLastRow As Long
    Dim i As Long, newRow As Long
    Dim prevValue As Variant, currentColor As Long
    Dim dictCouleurs As Object, numOrdre As String, couleurCell As Long
    Dim prepDict As Object

    Const PREP_FILE_PATH As String = "T:\DT atelier cartes\Demande de transfert Atelier cartes.xlsx"
    Const PREP_KEY_COLUMN As Long = 3
    Const PREP_VALUE_COLUMN As Long = 8



    '---- Optimisation Excel --------------------------------------------
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With
    
    
    On Error GoTo Cleanup


    '---- Recherche du fichier source -----------------------------------
    srcPath = "W:\CHARGE_SAP\"
    'srcPath = "D:\CHARGE_SAP\"
    'srcPath = "G:\Mon Drive\Traitement\Source\"
    'srcPath = "D:\_ Suivi fab carte journalier\"

    expectedFile = srcPath & "Plan_jal_aval_mef_" & Format(Date, "d_m_yyyy") & ".xlsx"

    If Dir(expectedFile) <> "" Then
        srcFullPath = expectedFile
    Else
        fileName = Dir(srcPath & "Plan_jal_aval_mef_*.xlsx")
        If fileName = "" Then Err.Raise vbObjectError + 513, , "Aucun fichier source trouvé."
        Do While fileName <> ""
            fileDate = FileDateTime(srcPath & fileName)
            If fileDate > latestDate Then latestDate = fileDate: latestFile = fileName
            fileName = Dir
        Loop
        srcFullPath = srcPath & latestFile
    End If
    
   
    
    ' Désactiver le partage si activé
    If ThisWorkbook.MultiUserEditing Then
        ThisWorkbook.ExclusiveAccess
    End If
    
    'Réinitialisation des filtres et des colonnes masquées
    Call DesactiverFiltresEtAfficherColonnes
    RemovePreparationColumn ThisWorkbook.Worksheets(2), "Pr�paration"
    RemovePreparationColumn ThisWorkbook.Worksheets(3), "Pr�paration"

    
    '====================================================================
    '  SAUVEGARDE DES COULEURS COLONNE Q (basée sur numéro d'ordre col B)
    '====================================================================
    Set destWs = ThisWorkbook.Worksheets(2) ' Définir destWs pour la sauvegarde
    Set dictCouleurs = CreateObject("Scripting.Dictionary")
    
    ' Sauvegarder les couleurs existantes de la colonne Q basées sur le numéro d'ordre (colonne B)
    destLastRow = destWs.Cells(destWs.Rows.Count, "B").End(xlUp).Row
    If destLastRow >= 2 Then
        For i = 2 To destLastRow
            numOrdre = CStr(destWs.Cells(i, "B").Value)
            If Len(numOrdre) > 0 Then
                couleurCell = destWs.Cells(i, "Q").Interior.Color
                ' Stocker la couleur si elle n'est pas la couleur par défaut
                If couleurCell <> RGB(255, 255, 255) And couleurCell <> xlColorIndexNone Then
                    dictCouleurs(numOrdre) = couleurCell
                End If
            End If
        Next i
    End If
    '====================================================================
  
    
    'Supprimer les mise en forme conditionnelles et les couleurs de la feuille 2 (CMS)
    Call SupprimerMFC_Et_CouleursDeFondSansPremiereLigne
    

    '---- Ouverture & préparation Feuil1 --------------------------------
    Set wbDest = ThisWorkbook
    Set ws = wbDest.Worksheets(1)
    Set destWs = wbDest.Worksheets(2)

    Set wbSource = Workbooks.Open(srcFullPath, ReadOnly:=True)
    srcFileName = wbSource.Name
    sheetName = Left(srcFileName, Len(srcFileName) - 5)

    ws.Cells.Clear
    On Error Resume Next
    ws.Name = sheetName
    If Err.Number <> 0 Then Err.Clear: ws.Name = "Données"
    On Error GoTo Cleanup


    wbSource.Worksheets(1).UsedRange.Copy ws.Range("A1")
    ws.Rows(1).Delete xlUp                      'en-tête en double
    
    ' Fermer le source sans sauvegarder
    wbSource.Close SaveChanges:=False

    

    '---- Nettoyage & filtres -------------------------------------------
    ws.Range("C:C,H:H,I:I,J:J,L:L,P:P,Q:Q,W:W,X:X,AA:AA,AB:AB,AC:AC," & _
             "AD:AD,AE:AE,AF:AF,AH:AH,AI:AI,AJ:AJ,AK:AK,AL:AL,AM:AM,AN:AN").Delete xlToLeft

    ws.Columns("Q").Cut
    ws.Columns("K").Insert xlToRight
    Application.CutCopyMode = False

    'A – ne garder que « OF ordo »
    With ws
        Dim rngA As Range: Set rngA = .Range("A1", .Cells(.Rows.Count, "A").End(xlUp))
        rngA.AutoFilter 1, "<>OF ordo"
        On Error Resume Next: rngA.Offset(1).SpecialCells(xlCellTypeVisible).EntireRow.Delete xlUp
        On Error GoTo Cleanup: .AutoFilterMode = False
    End With

    'F – **garder uniquement** les lignes commençant par "x" ou "X"
    With ws
        Dim rngF As Range: Set rngF = .Range("F1", .Cells(.Rows.Count, "F").End(xlUp))
        rngF.AutoFilter 1, "<>x*"         'supprime tout sauf x*
        On Error Resume Next: rngF.Offset(1).SpecialCells(xlCellTypeVisible).EntireRow.Delete xlUp
        On Error GoTo Cleanup: .AutoFilterMode = False
    End With

    'G – supprimer « OUV* »
    With ws
        Dim rngG As Range: Set rngG = .Range("G1", .Cells(.Rows.Count, "G").End(xlUp))
        rngG.AutoFilter 1, "OUV*"
        On Error Resume Next: rngG.Offset(1).SpecialCells(xlCellTypeVisible).EntireRow.Delete xlUp
        On Error GoTo Cleanup: .AutoFilterMode = False
    End With

    'Formule Reste à produire
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    ws.Range("L2:L" & lastRow).Formula = "=J2-K2"

    'R – supprimer charge restante = 0
    With ws
        Dim rngR As Range: Set rngR = .Range("R1", .Cells(.Rows.Count, "R").End(xlUp))
        rngR.AutoFilter 1, "0"
        On Error Resume Next: rngR.Offset(1).SpecialCells(xlCellTypeVisible).EntireRow.Delete xlUp
        On Error GoTo Cleanup: .AutoFilterMode = False
    End With

    
    'Supprimer Code gestionnaire, Statut Système, Date début opération, Date fin opération
    ws.Range("F:I").Delete xlToLeft
    
    ' Changer le titre par "Opérateur" en A1 et nettoyer le reste des valeurs de la colonne.
    With ws
        .Range("A1").Value = "Opérateur"
        .Range("A2:A" & .Rows.Count).ClearContents
        .Range("O1").Value = "Semaine"
        .Columns("N").Copy
        .Columns("O").PasteSpecial Paste:=xlPasteFormats
    End With
    
    Application.CutCopyMode = False
    
    
    '------------CopierValeursDepuisZPVB-------------------------------------
    Dim wbZPVB  As Workbook
    Dim wsZPVB  As Worksheet
    Dim chemin    As String
    Dim cle       As Variant
    Dim resultat  As Variant
    Dim lastData As Long
    Dim key As Variant, pos As Variant
    Dim colSemaine As Long
    
    'Construire le chemin vers ZPVB.XLSX (même dossier que ThisWorkbook)
    chemin = ThisWorkbook.Path & "\ZPVB.XLSX"
    
    
    'Ouvrir le classeur source en lecture seule
    On Error Resume Next
    Set wbZPVB = Workbooks.Open(fileName:=chemin, ReadOnly:=True)
    If wbZPVB Is Nothing Then
        MsgBox "Impossible d'ouvrir '" & chemin & "'. Vérifiez qu'il existe.", _
               vbExclamation, "Fichier introuvable"
        Exit Sub
    End If
    On Error GoTo 0
    
    'Pointer sur la feuille Sheet1 du classeur source
    Set wsZPVB = wbZPVB.Worksheets("Sheet1")
    
   ' Rechercher la colonne "Semaine" dans la ligne 1
    colSemaine = 0
    For i = 1 To wsZPVB.Cells(1, wsZPVB.Columns.Count).End(xlToLeft).Column
        If Trim(UCase(wsZPVB.Cells(1, i).Value)) = "SEMAINE" Then
            colSemaine = i
            Exit For
        End If
    Next i

    ' Vérification que la colonne "Semaine" a bien été trouvée
    If colSemaine = 0 Then
        MsgBox "La colonne 'Semaine' est introuvable dans Sheet1 de ZPVB.", vbCritical
        wbZPVB.Close SaveChanges:=False
        Exit Sub
    End If
    
    
    ' Déterminer la dernière ligne utilisée dans la source
    lastData = wsZPVB.Cells(wsZPVB.Rows.Count, "B").End(xlUp).Row
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    
    
    ' Boucle principale
    For i = 2 To lastRow
        key = ToText(CStr(ws.Cells(i, "B").Value))
        If IsNumeric(key) Then key = CDbl(key)

        On Error Resume Next
        pos = Application.WorksheetFunction.Match(key, _
              wsZPVB.Range("B2:B" & lastData), 0)
        If Err.Number <> 0 Or IsError(pos) Then
            ws.Cells(i, "O").Value = ""
            Err.Clear
        Else
            ws.Cells(i, "O").Value = wsZPVB.Cells(pos + 1, colSemaine).Value
        End If
        On Error GoTo 0
    Next i

    ' Fermer le classeur source sans enregistrer
    wbZPVB.Close SaveChanges:=False
    
    'L – garder « cablage cms* »
'    With ws
'        Dim rngO As Range: Set rngO = .Range("K1", .Cells(.Rows.Count, "K").End(xlUp))
'        rngO.AutoFilter 1, "=cablage cms*"
'    End With


    'Poste de travail – garder : "CMS-POSE" et "CMS-L1"
    
    With ws
        Dim rngO As Range
        Set rngO = .Range("J1", .Cells(.Rows.Count, "J").End(xlUp))
        
        rngO.AutoFilter Field:=1, Criteria1:=Array("CMS-POSE", "CMS-L1"), Operator:=xlFilterValues
    End With

    
    ' Remise en place onglet "CMS" à l'origine avant mise à jour
    With destWs
        .Columns(10).Cut: .Columns(6).Insert Shift:=xlToRight
        .Columns(11).Cut: .Columns(7).Insert Shift:=xlToRight
        .Columns(12).Cut: .Columns(8).Insert Shift:=xlToRight
        .Columns(15).Cut: .Columns(10).Insert Shift:=xlToRight
        .Columns(13).Cut: .Columns(11).Insert Shift:=xlToRight
    End With



    '====================================================================
    '  MÀJ INCRÉMENTALE   (clé = colonne B + colonne K)
    '====================================================================
    Const DATA_COLS As Long = 15   'A-O copiées
    Const OP_COL   As Long = 10   'colonne K absolue

    Dim dictSrc  As Object: Set dictSrc = CreateObject("Scripting.Dictionary")
    Dim dictDest As Object: Set dictDest = CreateObject("Scripting.Dictionary")

    Dim rngVis As Range, cel As Range
    Dim ordre As String, oper As String
    Dim srcRows As Collection, destRows As Collection
    Dim srcCnt As Long, destCnt As Long
    
    Dim rowsToDelete As Object: Set rowsToDelete = CreateObject("System.Collections.ArrayList")
    

    '1) Source ? dictSrc
    On Error Resume Next
    Set rngVis = ws.Range("B2", ws.Cells(ws.Rows.Count, "B").End(xlUp)) _
                    .SpecialCells(xlCellTypeVisible)
    On Error GoTo Cleanup


    If Not rngVis Is Nothing Then
        For Each cel In rngVis
            ordre = CStr(cel.Value)
            oper = LCase$(CStr(cel.Cells(1, OP_COL).Value))
            key = ordre & "|" & oper
            If dictSrc.Exists(key) Then
                dictSrc(key).Add cel.EntireRow
            Else
                Set srcRows = New Collection: srcRows.Add cel.EntireRow
                dictSrc.Add key, srcRows
            End If
        Next cel
    End If

    '2) Destination ? dictDest
    destLastRow = destWs.Cells(destWs.Rows.Count, "B").End(xlUp).Row
    For i = 2 To destLastRow
        ordre = CStr(destWs.Cells(i, 2).Value)
        oper = LCase$(CStr(destWs.Cells(i, OP_COL + 1).Value))
        key = ordre & "|" & oper
        If dictDest.Exists(key) Then
            dictDest(key).Add i
        Else
            Set destRows = New Collection: destRows.Add i
            dictDest.Add key, destRows
        End If
    Next i

    '3) Synchroniser clé par clé
    For Each key In dictSrc.Keys
        Set srcRows = dictSrc(key)
        If dictDest.Exists(key) Then
            Set destRows = dictDest(key)
        Else
            Set destRows = New Collection
        End If

        srcCnt = srcRows.Count: destCnt = destRows.Count

        '3-a MAJ lignes communes (A-O)
        For i = 1 To Application.Min(srcCnt, destCnt)
            destWs.Cells(destRows(i), 1).Resize(1, DATA_COLS).Value = _
                srcRows(i).Resize(1, DATA_COLS).Value
        Next i

        '3-b  (remplacer la suppression directe)
        For i = destCnt To srcCnt + 1 Step -1
            rowsToDelete.Add destRows(i)   'on stocke le numéro
        Next i
        
        '3-c Ajouter manquants
        For i = destCnt + 1 To srcCnt
            newRow = destWs.Cells(destWs.Rows.Count, "B").End(xlUp).Row + 1
            destWs.Cells(newRow, 1).Resize(1, DATA_COLS).Value = _
                srcRows(i).Resize(1, DATA_COLS).Value
        Next i

        If dictDest.Exists(key) Then dictDest.Remove key
    Next key

    '4) Clés restantes
    For Each key In dictDest.Keys
        For i = 1 To dictDest(key).Count
            rowsToDelete.Add dictDest(key)(i)
        Next i
    Next key

    '-----------------------------------------------
    '  Suppression réelle – à exécuter une seule fois
    If rowsToDelete.Count > 0 Then
        rowsToDelete.Sort
        For i = rowsToDelete.Count - 1 To 0 Step -1
            destWs.Rows(rowsToDelete(i)).Delete
        Next i
    End If
            
    
    
    
    '-----------------------------------------------------------------------


    ' Supprimer le filtrage de la feuille 1
    With ws
        If .FilterMode Then .ShowAllData
    End With

    ' Coloration alternée en Feuille1 selon changements de valeur en colonne A
    With ws
        destLastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        Dim paleBlue As Long, paleYellow As Long
        paleBlue = RGB(204, 229, 255)    ' Bleu pâle
        paleYellow = RGB(255, 255, 204)  ' Jaune pâle
        If destLastRow >= 2 Then
            prevValue = .Cells(2, "B").Value
            currentColor = paleBlue
            For i = 2 To destLastRow
                If .Cells(i, "B").Value <> prevValue Then
                    If currentColor = paleBlue Then
                        currentColor = paleYellow
                    Else
                        currentColor = paleBlue
                    End If
                    prevValue = .Cells(i, "B").Value
                End If
                .Rows(i).Interior.Color = currentColor
            Next i
        End If
    End With
    
    
    '--- Tri Feuil2 : 1) Semaine (colonne O), 2) N° d'ordre (numérique, colonne B)
    
    Tri destWs, "O", "AA"
    
    
    
    ' Coloration alternée en Feuille2 selon changements de valeur en colonne A
'    With destWs
'        destLastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
'        paleBlue = RGB(204, 229, 255)    ' Bleu pâle
'        paleYellow = RGB(255, 255, 204)  ' Jaune pâle
'        If destLastRow >= 2 Then
'            prevValue = .Cells(2, "B").Value
'            currentColor = paleBlue
'            For i = 2 To destLastRow
'                If .Cells(i, "B").Value <> prevValue Then
'                    If currentColor = paleBlue Then
'                        currentColor = paleYellow
'                    Else
'                        currentColor = paleBlue
'                    End If
'                    prevValue = .Cells(i, "B").Value
'                End If
'                .Rows(i).Interior.Color = currentColor
'            Next i
'        End If
'    End With
    
   
    
    ' Mise en place validation de données
    With destWs.Range("P2:P" & destLastRow).Validation
        ' 2) On supprime d'abord toute validation existante
        .Delete
        
        ' 3) On crée la validation Liste
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="CMS-POSE,CMS-L1,YAMAHA,PROBLEME,PAS DE PROG,EN COURS,FAIT"
    End With

    ' Limiter les nombres décimaux à deux chiffres après la virgule uniquement sur les colonnes M et N des feuilles 1 et 2
    ws.Columns("M:N").NumberFormat = "0.00"
    destWs.Columns("M:N").NumberFormat = "0.00"


    ' Ajuster automatiquement la largeur des colonnes sur les feuilles 1 et 2
'    ws.UsedRange.Columns.AutoFit
'    destWs.UsedRange.Columns.AutoFit

    ' Centrer le contenu sur les feuilles 1 et 2
    With ws.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With destWs.UsedRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    '===  Feuil1 ? Feuil3  =============================================
    Call UpdateSheet3(ws, wbDest.Worksheets(3))   'ws = Feuil1
    
    
    
    Call IgnorerErreurs(ws)
    Call IgnorerErreurs(destWs)

    ' Supprimer colonne Opérateur sur feuille1
    ws.Range("A:A").Delete xlToLeft



    ' Mise en forme finale et positionnement des 3 feuilles
    ws.Activate
    ws.Range("A1").Select

    
    With destWs
        .Columns(9).Cut: .Columns(6).Insert Shift:=xlToRight
        .Columns(12).Cut: .Columns(7).Insert Shift:=xlToRight
        .Columns(13).Cut: .Columns(8).Insert Shift:=xlToRight
        .Columns(13).Cut: .Columns(9).Insert Shift:=xlToRight
        .Columns(14).Cut: .Columns(13).Insert Shift:=xlToRight
        .Columns(15).Cut: .Columns(14).Insert Shift:=xlToRight
        .Activate
        .Range("A1").Select
    End With
    
    
    
    'wbDest.Worksheets(3).Activate
    'wbDest.Worksheets(3).Range("A1").Select

    Call VerifierTousLesIPR_ColonneC
    
    Call ActiverFiltresEtEffacerFormatHorsDonnees
    
    ' Mise en forme de la colonne A de la feuille "Planning"
    
    With wbDest.Worksheets(3)
        .Activate
        .Range("A1").VerticalAlignment = xlTop
        .Columns("A").ColumnWidth = 42
        .Range("A1").Select
        
    End With


    Call AppliquerMiseEnFormeConditionnelleCMS
    
    '====================================================================
    '  RESTAURATION DES COULEURS COLONNE Q (basée sur numéro d'ordre col B)
    '====================================================================
    ' Restaurer les couleurs sauvegardées dans la colonne Q
    destLastRow = destWs.Cells(destWs.Rows.Count, "B").End(xlUp).Row
    If destLastRow >= 2 Then
        For i = 2 To destLastRow
            numOrdre = CStr(destWs.Cells(i, "B").Value)
            If Len(numOrdre) > 0 And dictCouleurs.Exists(numOrdre) Then
                destWs.Cells(i, "Q").Interior.Color = dictCouleurs(numOrdre)
            End If
        Next i
    End If
    
    '====================================================================
        
    
    
    
    Set prepDict = LoadPreparationData(PREP_FILE_PATH, PREP_KEY_COLUMN, PREP_VALUE_COLUMN)
    InsertPreparationColumnWithData destWs, 10, "Pr�paration", prepDict
    InsertPreparationColumnWithData wbDest.Worksheets(3), 5, "Pr�paration", prepDict

    destWs.Range("A:A,E:H").EntireColumn.Hidden = True

    MsgBox "Mise à jour terminée : données importées de " & srcFileName, vbInformation

    ' Réactiver le partage
    On Error Resume Next
    ThisWorkbook.SaveAs fileName:=ThisWorkbook.FullName, AccessMode:=xlShared
    On Error GoTo 0
    


Cleanup:
    With Application
        .DisplayAlerts = True
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    If Err.Number <> 0 Then MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

'----------------------------------------------------------------------
Private Sub RemovePreparationColumn(ws As Worksheet, headerName As String)
    Dim headerCell As Range

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set headerCell = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0

    If Not headerCell Is Nothing Then
        ws.Columns(headerCell.Column).Delete
    End If
End Sub

'----------------------------------------------------------------------
Private Function LoadPreparationData(filePath As String, keyColumn As Long, valueColumn As Long) As Object
    Dim wb As Workbook, wsSrc As Worksheet
    Dim lastRow As Long, r As Long
    Dim dict As Object
    Dim key As String

    If Len(filePath) = 0 Then Exit Function
    If Dir(filePath) = "" Then Exit Function

    On Error GoTo CleanExit

    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    Set wsSrc = wb.Worksheets(1)

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, keyColumn).End(xlUp).Row
    If lastRow < 2 Then GoTo CleanExit

    Set dict = CreateObject("Scripting.Dictionary")

    For r = 2 To lastRow
        key = ToText(wsSrc.Cells(r, keyColumn).Value)
        If Len(key) > 0 Then
            If Not dict.Exists(key) Then
                dict.Add key, wsSrc.Cells(r, valueColumn).Value
            End If
        End If
    Next r

CleanExit:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set LoadPreparationData = dict
End Function

'----------------------------------------------------------------------
Private Sub InsertPreparationColumnWithData(ws As Worksheet, colIndex As Long, headerName As String, prepDict As Object)
    Dim lastRow As Long, r As Long
    Dim key As String

    If ws Is Nothing Then Exit Sub
    If colIndex < 1 Then Exit Sub

    ws.Columns(colIndex).Insert Shift:=xlToRight
    ws.Cells(1, colIndex).Value = headerName

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For r = 2 To lastRow
        key = ToText(ws.Cells(r, "B").Value)
        If Len(key) > 0 Then
            If Not prepDict Is Nothing And prepDict.Exists(key) Then
                ws.Cells(r, colIndex).Value = prepDict(key)
            Else
                ws.Cells(r, colIndex).Value = ""
            End If
        Else
            ws.Cells(r, colIndex).Value = ""
        End If
    Next r
End Sub


'======================================================================
'  UpdateSheet3  –  Feuil1  ?  Feuil3
'  Colonnes gardées : A B C D F G H I K L N O  = 1-2-3-4-6-7-8-9-11-12-14-15
'  Clé = "N° d'ordre" (col. B) + "Opération" (col. K)
'======================================================================
Public Sub UpdateSheet3(ByVal srcWs As Worksheet, ByVal destWs As Worksheet)

    Dim keepCols As Variant: keepCols = Array(1, 2, 3, 4, 6, 7, 8, 9, 11, 12, 14, 15)
    Const ordreCol& = 2          'colonne B
    Const operCol& = 11          'colonne K

    '------------------------------------------------------------------
    ' 1)  Index de Feuil3  (clé ? ligne)
    '------------------------------------------------------------------
    Dim dictDest As Object: Set dictDest = CreateObject("Scripting.Dictionary")
    Dim matched  As Object: Set matched = CreateObject("Scripting.Dictionary")
    Dim lastDest&, r&
    Dim key As Variant
    

    lastDest = destWs.Cells(destWs.Rows.Count, "B").End(xlUp).Row
    For r = 2 To lastDest
        key = CStr(destWs.Cells(r, ordreCol).Value) & "|" & _
              CStr(destWs.Cells(r, operCol - 2).Value)
        If Len(key) > 1 Then dictDest(key) = r
    Next r

    '------------------------------------------------------------------
    ' 2)  Parcours intégral de Feuil1
    '------------------------------------------------------------------
    Dim lastSrc&, newRow&
    lastSrc = srcWs.Cells(srcWs.Rows.Count, ordreCol).End(xlUp).Row
    Dim nextFreeRow As Long
    nextFreeRow = destWs.Cells(destWs.Rows.Count, ordreCol).End(xlUp).Row + 1
             'ordreCol = 2  ? on vise la colonne B

    For r = 2 To lastSrc
        key = CStr(srcWs.Cells(r, ordreCol).Value) & "|" & _
              CStr(srcWs.Cells(r, operCol).Value)
    
        If dictDest.Exists(key) Then
            '--- mise à jour (col. A préservée)
            CopyRowFixed srcWs, r, destWs, dictDest(key), keepCols, False
            matched(key) = True
    
        Else
            '--- AJOUT : on colle ligne après ligne grâce au pointeur
            CopyRowFixed srcWs, r, destWs, nextFreeRow, keepCols, True
            dictDest(key) = nextFreeRow
            matched(key) = True
            nextFreeRow = nextFreeRow + 1            '? avance le curseur
        End If
    Next r
    '------------------------------------------------------------------
    ' 3)  Suppression des lignes obsolètes
    '------------------------------------------------------------------
    Dim toDel() As Variant, i&
    For Each key In dictDest.Keys
        If Not matched.Exists(key) Then
            ReDim Preserve toDel(0 To i)
            toDel(i) = dictDest(key): i = i + 1
        End If
    Next key
    If i > 0 Then
        Dim j&
        For j = UBound(toDel) To 0 Step -1
            destWs.Rows(toDel(j)).Delete
        Next j
    End If

    '------------------------------------------------------------------
    ' 4)  Mise en forme finale
    '------------------------------------------------------------------
    
    Tri destWs, "L", "AA"
    ColorBands destWs, ordreCol
    With destWs.UsedRange
        .Columns.AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

'----------------------------------------------------------------------
'  Copie des colonnes fixées (contiguës A?L)
'----------------------------------------------------------------------
Private Sub CopyRowFixed(srcWs As Worksheet, srcRow&, _
                         destWs As Worksheet, destRow&, _
                         keepCols As Variant, ByVal copyColA As Boolean)

    Dim idx&, srcCol&, dstCol&
    For idx = LBound(keepCols) To UBound(keepCols)
        srcCol = keepCols(idx)         'colonne dans Feuil1
        dstCol = idx + 1               'position contiguë dans Feuil3

        If dstCol = 1 And Not copyColA Then
            'préserver la colonne A
        Else
            destWs.Cells(destRow, dstCol).Value = _
                srcWs.Cells(srcRow, srcCol).Value
        End If
    Next idx
End Sub

'----------------------------------------------------------------------
'  Coloration alternée (ordre)
'----------------------------------------------------------------------
Private Sub ColorBands(ws As Worksheet, keyCol&)
    Const Col1& = &HFFE5CC       'bleu pâle
    Const Col2& = &HCCFFFF       'jaune pâle

    Dim lastRow&, r&, curCol&, prevVal
    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    curCol = Col1
    prevVal = ws.Cells(2, keyCol).Value

    For r = 2 To lastRow
        If ws.Cells(r, keyCol).Value <> prevVal Then
            curCol = IIf(curCol = Col1, Col2, Col1)
            prevVal = ws.Cells(r, keyCol).Value
        End If
        ws.Rows(r).Interior.Color = curCol
    Next r
End Sub


'----------------------------------------------------------------------
'--- Tri Feuil : 1) Semaine , 2) N° d'ordre (numérique, colonne B)
'----------------------------------------------------------------------
    
Private Sub Tri(ws As Worksheet, Col1 As String, ColTemp As String)
  
    With ws
        Dim lastDataRow As Long
        lastDataRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        
        'Colonne temporaire AA : conversion texte en nombre pour N° d'ordre (colonne B)
        .Columns(ColTemp).Insert
        .Range(ColTemp & "1").Value = "_OrdreNum" ' En-tête pour la colonne temporaire
        .Range(ColTemp & "2:" & ColTemp & lastDataRow).FormulaR1C1 = "=--RC2" ' Convertit B en numérique dans AA
        
        'Tri sur O (Semaine) puis AA (N° d'ordre numérique)
        ' Assurez-vous que la plage à trier inclut la colonne O et la colonne temporaire AA.
        ' .UsedRange est généralement suffisant si toutes les données sont contiguës et que AA est la dernière colonne.
        .UsedRange.Sort Key1:=.Range(Col1 & "1"), Order1:=xlAscending, _
                        Key2:=.Range(ColTemp & "1"), Order2:=xlAscending, _
                        Header:=xlYes
        
        .Columns(ColTemp).Delete ' Supprime la colonne temporaire
    End With

 End Sub
    









