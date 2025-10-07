Attribute VB_Name = "Module2"
Sub LargeursColonnesFeuilleActive()
    Dim ws As Worksheet
    Dim nbCols As Integer
    Dim i As Integer
    Dim txt As String
    Dim largeur As String

    Set ws = ActiveSheet

    nbCols = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    txt = ""
    For i = 1 To nbCols
        largeur = CStr(ws.Columns(i).ColumnWidth)
        largeur = Replace(largeur, ",", ".")
        txt = txt & largeur & ", "
    Next i

    If Len(txt) > 2 Then txt = Left(txt, Len(txt) - 2)
    InputBox "Voici la liste à copier/coller dans ta macro :", "Largeurs de colonnes", txt
End Sub


Sub IgnorerErreurs(ws As Worksheet)
    Dim cell As Range
    Dim plageB As Range
    Dim plageI As Range
    Dim plageTotale As Range

    ' Détection automatique des lignes utilisées dans les colonnes B et K de la feuille spécifiée
    On Error Resume Next
    Set plageB = ws.Range("B1:B" & ws.Cells(ws.Rows.Count, "B").End(xlUp).Row)
    Set plageI = ws.Range("I1:I" & ws.Cells(ws.Rows.Count, "I").End(xlUp).Row)
    Set plageTotale = Union(plageB, plageI)
    On Error GoTo 0

    For Each cell In plageTotale
        If cell.Errors(xlNumberAsText).Value Then
            cell.Errors(xlNumberAsText).Ignore = True
        End If
    Next cell

End Sub


Sub AjouterMenuMacro()
    Dim cb As CommandBar, cbc As CommandBarControl

    On Error Resume Next
    ' Supprimer l'ancien menu s'il existe
    Application.CommandBars("Worksheet Menu Bar").Controls("Mes macros").Delete
    On Error GoTo 0

    ' Ajouter un nouveau menu
    Set cb = Application.CommandBars("Worksheet Menu Bar")
    With cb.Controls.Add(Type:=msoControlPopup, Temporary:=True)
        .Caption = "Mes macros"

        ' Ajouter un sous-menu pour lancer la macro
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Importer Données"
            .OnAction = "UpdateDailyData" ' nom exact de ta macro
            .FaceId = 59 ' Icône par défaut (modifiable)
        End With
    End With
End Sub





'—–––––––––––––––––––––––––––––––––––––––––––––––––––––––––––
' Fonction CleanString : équivalent de CLEAN + TRIM
Function CleanString(ByVal s As String) As String
    Dim i As Long, ch As String, tmp As String
    tmp = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        ' Ne conserver que les caractères dont le code ASCII = 32
        If Asc(ch) >= 32 Then tmp = tmp & ch
    Next i
    CleanString = VBA.Trim(tmp)
End Function

' Fonction ToNumber : force un texte numérique en Double
Function ToNumber(ByVal s As String) As Variant
    On Error Resume Next
    ToNumber = CDbl(CleanString(s))
    If Err.Number <> 0 Then
        ToNumber = CVErr(xlErrValue)
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Fonction ToText : force n'importe quoi en String
Function ToText(ByVal v As Variant) As String
    ToText = CStr(CleanString(CStr(v)))
End Function
'—–––––––––––––––––––––––––––––––––––––––––––––––––––––––––––


Sub GénérerOps()
    '--- feuilles ---
    Const SHEET_PLAN As String = "Plan"      'table avec les 2 listes
    Const SHEET_DATA As String = "CMS"   'ta feuille de travail
                                            '? adapte si besoin
    
    Dim wsPlan As Worksheet, wsD As Worksheet
    Set wsPlan = Worksheets(SHEET_PLAN)
    Set wsD = Worksheets(SHEET_DATA)
    
    '--- charge les deux listes dans des arrays ---
    Dim startArr As Variant, targetArr As Variant
    startArr = Application.Transpose(wsPlan.Range("A1", wsPlan.Cells(wsPlan.Rows.Count, "A").End(xlUp)))
    targetArr = Application.Transpose(wsPlan.Range("B1", wsPlan.Cells(wsPlan.Rows.Count, "B").End(xlUp)))
    
    '--- dictionnaire positions courantes ---
    Dim pos As Object: Set pos = CreateObject("scripting.dictionary")
    Dim i As Long
    For i = LBound(startArr) To UBound(startArr)
        pos(startArr(i)) = i                    'index 1-based = n° de colonne
    Next i
    
    '--- nettoie la zone de sortie (C:D) ---
    wsPlan.Range("C:D").ClearContents
    Dim rowOut As Long: rowOut = 1
    
    '--- parcours de gauche à droite (ordre final) ---
    Dim curPos As Long, tgtPos As Long
    Dim k As Variant
    For i = LBound(targetArr) To UBound(targetArr)
        If Len(targetArr(i)) = 0 Then Exit For                      'sécurité
        If Not pos.Exists(targetArr(i)) Then
            wsPlan.Cells(rowOut, "C").Value = "? Introuvable : " & targetArr(i)
            rowOut = rowOut + 1
        Else
            curPos = pos(targetArr(i))          'position actuelle
            tgtPos = i                          'position voulue
            
            If curPos <> tgtPos Then
                '--- ligne descriptive ---
                wsPlan.Cells(rowOut, "C").Value = _
                    "Déplacer '" & targetArr(i) & "' de col " & curPos & " ? col " & tgtPos
                '--- ligne de code prête à coller ---
                wsPlan.Cells(rowOut, "D").Value = _
                    "ws.Columns(" & curPos & ").Cut : ws.Columns(" & tgtPos & ").Insert Shift:=xlToRight"
                rowOut = rowOut + 1
                
                '--- met à jour les positions dans le dico ---
                For Each k In pos.Keys
                    If curPos > tgtPos Then            'on décale vers la gauche
                        If pos(k) >= tgtPos And pos(k) < curPos Then pos(k) = pos(k) + 1
                    Else                                'vers la droite
                        If pos(k) <= tgtPos And pos(k) > curPos Then pos(k) = pos(k) - 1
                    End If
                Next k
                pos(targetArr(i)) = tgtPos
            End If
        End If
    Next i
    
    MsgBox "Liste des opérations générée !", vbInformation
End Sub





'------------------------------------------------------------------------------
'----------- RECHERCHER IPR ---------------------------------------------------
'------------------------------------------------------------------------------
'Private Sub Tri(ws As Worksheet, Col1 As String, ColTemp As String)

Sub RechIPR(codecar As String)


'remplacement / par- dans code article
slash = Len(codecar)
line20:
For remsla = 1 To slash
    If Mid(codecar, remsla, 1) = "/" Then
        Mid(codecar, remsla, 1) = "-"
        If remsla < slash Then
            GoTo line20
        End If
    End If
   
Next

'Sheets(1).Select

fichIPR = Dir("S:\Methodes Production\0- IPR VALIDE" & "\" & codecar & ".xls*")
fichIPRw = Dir("S:\Methodes Production\0- IPR VALIDE" & "\" & codecar & ".doc*")
If fichIPR <> "" Then
    MsgBox ("code trouvé dans IPR VALIDE")
    Workbooks.Open fileName:="S:\Methodes Production\0- IPR VALIDE" & "\" & codecar & ".xls*"
    GoTo line30
End If
If fichIPRw <> "" Then
    MsgBox ("code trouvé dans IPR VALIDE, Fichier Word")
    Set ww = CreateObject("word.application")
    ww.Visible = True
    If UCase(Right(fichIPRw, 1)) = "X" Then
        ww.Documents.Open fileName:="S:\Methodes Production\0- IPR VALIDE" & "\" & codecar & ".docx"
    Else
        ww.Documents.Open fileName:="S:\Methodes Production\0- IPR VALIDE" & "\" & codecar & ".doc"
    End If
    GoTo line30
End If
 
 
fichIPR = Dir("S:\Methodes Production\1- IPR AUTORISEES" & "\" & codecar & ".xls*")
fichIPRw = Dir("S:\Methodes Production\1- IPR AUTORISEES" & "\" & codecar & ".doc*")
If fichIPR <> "" Then
    MsgBox ("code trouvé dans IPR AUTORISES")
    Workbooks.Open fileName:="S:\Methodes Production\1- IPR AUTORISEES" & "\" & codecar & ".xls*"
    GoTo line30
End If
If fichIPRw <> "" Then
    MsgBox ("code trouvé dans IPR AUTORISES, Fichier Word")
    Set ww = CreateObject("word.application")
    ww.Visible = True
    If UCase(Right(fichIPRw, 1)) = "X" Then
        ww.Documents.Open fileName:="S:\Methodes Production\1- IPR AUTORISEES" & "\" & codecar & ".docx"
    Else
        ww.Documents.Open fileName:="S:\Methodes Production\1- IPR AUTORISEES" & "\" & codecar & ".doc"
    End If
    GoTo line30
End If


fichIPR = Dir("S:\Methodes Production\2- IPR en COURS" & "\" & codecar & ".xls*")
fichIPRw = Dir("S:\Methodes Production\2- IPR en COURS" & "\" & codecar & ".doc*")
If fichIPR <> "" Then
    pos = MsgBox("code trouvé dans IPR en cours,                                                              N'UTILISER QUE LES POSTES EN VERT    ", vbOKOnly, "Attention")
    Workbooks.Open fileName:="S:\Methodes Production\2- IPR en COURS" & "\" & codecar & ".xls"
    GoTo line30
End If
If fichIPRw <> "" Then
    MsgBox ("code trouvé dans IPR en COURS,                                                              N'UTILISER QUE LES POSTES EN VERT    , fichier Word")
    Set ww = CreateObject("word.application")
    ww.Visible = True
    If UCase(Right(fichIPRw, 1)) = "X" Then
        ww.Documents.Open fileName:="S:\Methodes Production\2- IPR en COURS" & "\" & codecar & ".docx"
    Else
        ww.Documents.Open fileName:="S:\Methodes Production\2- IPR en COURS" & "\" & codecar & ".doc"
    End If
    GoTo line30
End If


fichIPR = Dir("S:\Methodes Production\3- IPR ARCHIVES" & "\" & codecar & ".xls*")
fichIPRw = Dir("S:\Methodes Production\3- IPR ARCHIVES" & "\" & codecar & ".doc*")
If fichIPR <> "" Then
    pos = MsgBox("code trouvé dans IPR ARCHIVES, ne pas utiliser, consulter les méthodes", vbOKOnly, "Attention")
    'Workbooks.Open Filename:="S:\Methodes Production\0- IPR ARCHIVES" & "\" & codecar & ".xls"
    GoTo line30
End If
If fichIPRw <> "" Then
    MsgBox ("code trouvé dans IPR ARCHIVES, ne pas utiliser, consulter les méthodes, fichier Word ")
    GoTo line30
End If

MsgBox ("pas d'IPR trouvé")
   
line30:

'sauvegarde fichier si lecture ecriture
Application.DisplayAlerts = False
If ActiveWorkbook.ReadOnly = False Then
    ActiveWorkbook.Save
End If
Application.DisplayAlerts = True

End Sub


Sub ActiverFiltresEtQuadrillage()
    Dim ws As Worksheet
    Dim plageDonnees As Range

    On Error Resume Next ' En cas d'erreur (feuille vide, etc.)

    For Each ws In ThisWorkbook.Worksheets
        With ws
            If .AutoFilterMode Then .AutoFilterMode = False
            
            ' Vérifie s'il y a des données sur la feuille
            If Application.WorksheetFunction.CountA(.Cells) > 0 Then
                ' Détecte la plage de données à partir de A1
                Set plageDonnees = .Range("A1").CurrentRegion
                
                ' Active le filtre automatique sur la plage détectée
                plageDonnees.AutoFilter

                ' Réinitialise les bordures de la plage
                plageDonnees.Borders.LineStyle = xlContinuous
                plageDonnees.Borders.Weight = xlThin
                plageDonnees.Borders.ColorIndex = xlAutomatic
            End If
        End With
    Next ws

    On Error GoTo 0
End Sub


Sub ActiverFiltresEtEffacerFormatHorsDonnees()
    Dim ws As Worksheet
    Dim plageDonnees As Range
    Dim derniereLigne As Long
    Dim derniereColonne As Long
    Dim i As Long
    Dim col As Range

    On Error Resume Next

    For Each ws In ThisWorkbook.Worksheets
        With ws
            If .AutoFilterMode Then .AutoFilterMode = False

            If Application.WorksheetFunction.CountA(.Cells) > 0 Then
                ' Détecte la plage des données à partir de A1
                Set plageDonnees = .Range("A1").CurrentRegion

                ' Applique le filtre automatique
                plageDonnees.AutoFilter

                ' Applique les bordures sur les données
                With plageDonnees.Borders
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With

                ' Détection des dernières lignes et colonnes utilisées
                derniereLigne = plageDonnees.Row + plageDonnees.Rows.Count - 1
                derniereColonne = plageDonnees.Column + plageDonnees.Columns.Count - 1

                ' Effacer le formatage des colonnes après les données
                If derniereColonne < ws.Columns.Count Then
                    .Range(.Cells(1, derniereColonne + 1), .Cells(ws.Rows.Count, ws.Columns.Count)).ClearFormats
                End If

                ' Effacer le formatage des lignes après les données
                If derniereLigne < ws.Rows.Count Then
                    .Range(.Cells(derniereLigne + 1, 1), .Cells(ws.Rows.Count, ws.Columns.Count)).ClearFormats
                End If

                ' Effacer le formatage des colonnes avant la zone de données (si besoin)
                If plageDonnees.Column > 1 Then
                    .Range(.Cells(1, 1), .Cells(ws.Rows.Count, plageDonnees.Column - 1)).ClearFormats
                End If

                ' Effacer le formatage des lignes avant la zone de données (si besoin)
                If plageDonnees.Row > 1 Then
                    .Range(.Cells(1, 1), .Cells(plageDonnees.Row - 1, ws.Columns.Count)).ClearFormats
                End If
            End If
            For Each col In .UsedRange.Columns
                If Not col.EntireColumn.Hidden Then
                    col.EntireColumn.AutoFit
                End If
            Next col
        End With
    Next ws

    On Error GoTo 0
End Sub





Sub DesactiverFiltresEtAfficherColonnes()
    Dim ws As Worksheet

    On Error Resume Next ' En cas d'erreur sur une feuille vide, ignorer

    For Each ws In ThisWorkbook.Worksheets
        With ws
            ' Désactiver le filtre s'il est actif
            If .AutoFilterMode Then
                .AutoFilterMode = False
            End If
            
            ' Réafficher toutes les colonnes (colonnes 1 à 16384 = A à XFD)
            .Columns.Hidden = False
        End With
    Next ws

    On Error GoTo 0
End Sub


Sub RepositionnementColFeuilleCMS()

    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(2)


    
    With ws
        .Columns(9).Cut: .Columns(6).Insert Shift:=xlToRight
        .Columns(12).Cut: .Columns(7).Insert Shift:=xlToRight
        .Columns(13).Cut: .Columns(8).Insert Shift:=xlToRight
        .Columns(13).Cut: .Columns(9).Insert Shift:=xlToRight
        .Columns(14).Cut: .Columns(13).Insert Shift:=xlToRight
        .Columns(15).Cut: .Columns(14).Insert Shift:=xlToRight

    End With

End Sub


Sub SupprimerMFC_Et_CouleursDeFondSansPremiereLigne()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim plageCible As Range

    ' Définir la feuille concernée
    Set ws = ThisWorkbook.Worksheets(2)

    ' Supprimer toutes les règles de mise en forme conditionnelle
    ws.Cells.FormatConditions.Delete

    ' Déterminer la dernière ligne et la dernière colonne utilisées
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Définir la plage à partir de la ligne 2 jusqu’à la dernière ligne utilisée
    If lastRow >= 2 And lastCol >= 1 Then
        Set plageCible = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
        plageCible.Interior.ColorIndex = xlColorIndexNone
    End If

End Sub




Sub ValidationDonneesFeuilleCMS()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim destLastRow As Long
    Dim i As Long
    Dim cellValue As String

    Set wb = ThisWorkbook
    Set ws = wb.Worksheets(2)

    destLastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Suppression et ajout de validation de données
    With ws.Range("P2:P" & destLastRow).Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="CMS-POSE,CMS-L1,YAMAHA,PROBLEME,PAS DE PROG,EN COURS,FAIT"
    End With

End Sub

Sub AppliquerMiseEnFormeConditionnelleCMS()

    Dim ws As Worksheet
    Dim plageConditionnelle As Range
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(2) ' À adapter si besoin

    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    Set plageConditionnelle = ws.Range("A2:Y" & lastRow)

    ' Supprimer les anciennes règles
    plageConditionnelle.FormatConditions.Delete

    With plageConditionnelle
        ' 1. YAMAHA --> Bleu
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$P2=""YAMAHA"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 200, 255)

        ' 2. PROBLEME --> Orange
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$P2=""PROBLEME"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 192, 0)

        ' 3. PAS DE PROG --> Rouge
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$P2=""PAS DE PROG"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 0, 0)

        ' 4. EN COURS --> Jaune
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$P2=""EN COURS"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 0)

        ' 5. FAIT --> Vert
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$P2=""FAIT"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(0, 176, 80)

        ' 6. CMS-L1 --> Blanc
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$P2=""CMS-L1"""
        .FormatConditions(.FormatConditions.Count).Interior.ColorIndex = xlNone
        
        ' 7. CMS-POSE --> Blanc
        .FormatConditions.Add Type:=xlExpression, Formula1:="=$P2=""CMS-POSE"""
        .FormatConditions(.FormatConditions.Count).Interior.ColorIndex = xlNone


    End With
    
    ws.Columns("Q").FormatConditions.Delete

End Sub





Sub RechIPR2(codecar As String)
    Dim dossierList As Variant
    Dim dossierName As Variant
    Dim basePath As String
    Dim fichExcel As String, fichWord As String
    Dim ww As Object
    Dim cheminComplet As String
    Dim msgTexte As String
    Dim msgType As VbMsgBoxStyle
    Dim ub As Integer

    ' Remplacement des / par -
    codecar = Replace(codecar, "/", "-")

    ' Liste des sous-dossiers avec messages et types d'alerte
    dossierList = Array( _
        Array("0- IPR VALIDE", "code trouvé dans IPR VALIDE"), _
        Array("1- IPR AUTORISEES", "code trouvé dans IPR AUTORISEES"), _
        Array("2- IPR en COURS", "code trouvé dans IPR EN COURS - N'UTILISER QUE LES POSTES EN VERT", vbOKOnly + vbExclamation), _
        Array("3- IPR ARCHIVES", "code trouvé dans IPR ARCHIVES - ne pas utiliser, consulter les méthodes", vbOKOnly + vbCritical) _
    )

    ' Parcours des dossiers
    For Each dossierName In dossierList
        basePath = "S:\Methodes Production\" & dossierName(0) & "\"

        ' Vérifie combien d'éléments contient le sous-tableau
        ub = UBound(dossierName)

        ' Affecte message et type
        msgTexte = dossierName(1)
        If ub >= 2 Then
            msgType = dossierName(2)
        Else
            msgType = vbInformation
        End If

        ' Recherche Excel
        fichExcel = Dir(basePath & codecar & ".xls*")
        If fichExcel <> "" Then
            MsgBox msgTexte, msgType
            Workbooks.Open fileName:=basePath & fichExcel
            GoTo Fin
        End If

        ' Recherche Word
        fichWord = Dir(basePath & codecar & ".doc*")
        If fichWord <> "" Then
            MsgBox msgTexte & " (fichier Word)", msgType
            Set ww = CreateObject("Word.Application")
            ww.Visible = True
            ww.Documents.Open basePath & fichWord
            GoTo Fin
        End If
    Next dossierName

    ' Aucun fichier trouvé
    MsgBox "Aucun fichier IPR trouvé pour le code : " & codecar, vbExclamation
    Exit Sub

Fin:
    ' Sauvegarde si non en lecture seule
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not ActiveWorkbook.ReadOnly Then ActiveWorkbook.Save
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub


Sub RechIPR_InfosUniquement(codecar As String)
    Dim dossierList As Variant
    Dim dossierName As Variant
    Dim basePath As String
    Dim fichExcel As String, fichWord As String
    Dim msgTexte As String
    Dim msgType As VbMsgBoxStyle
    Dim ub As Integer
    Dim resultatTrouve As Boolean

    ' Remplacement des / par -
    codecar = Replace(codecar, "/", "-")
    resultatTrouve = False

    ' Liste des sous-dossiers avec messages et types d'alerte
    dossierList = Array( _
        Array("0- IPR VALIDE", "code trouvé dans IPR VALIDE"), _
        Array("1- IPR AUTORISEES", "code trouvé dans IPR AUTORISEES"), _
        Array("2- IPR en COURS", "code trouvé dans IPR EN COURS - N'UTILISER QUE LES POSTES EN VERT", vbOKOnly + vbExclamation), _
        Array("3- IPR ARCHIVES", "code trouvé dans IPR ARCHIVES - ne pas utiliser, consulter les méthodes", vbOKOnly + vbCritical) _
    )

    ' Parcours des dossiers
    For Each dossierName In dossierList
        basePath = "S:\Methodes Production\" & dossierName(0) & "\"

        ub = UBound(dossierName)
        msgTexte = dossierName(1)
        If ub >= 2 Then
            msgType = dossierName(2)
        Else
            msgType = vbInformation
        End If

        ' Recherche Excel
        fichExcel = Dir(basePath & codecar & ".xls*")
        If fichExcel <> "" Then
            MsgBox msgTexte & " (fichier Excel détecté)", msgType
            resultatTrouve = True
        End If

        ' Recherche Word
        fichWord = Dir(basePath & codecar & ".doc*")
        If fichWord <> "" Then
            MsgBox msgTexte & " (fichier Word détecté)", msgType
            resultatTrouve = True
        End If
    Next dossierName

    ' Aucun fichier trouvé
    If Not resultatTrouve Then
        MsgBox "Aucun fichier IPR trouvé pour le code : " & codecar, vbExclamation
    End If
End Sub


Sub VerifierTousLesIPR_ColonneC()
    Dim ws As Worksheet
    Dim ligne As Long
    Dim derniereLigne As Long
    Dim codecar As String
    Dim resultat As String
    Dim celluleResultat As Range

    Set ws = Worksheets(2) ' Feuille 2

    ' Détecter la dernière ligne avec une valeur dans la colonne C
    derniereLigne = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    ' Parcourir chaque ligne
    For ligne = 2 To derniereLigne ' en supposant une ligne d'en-tête
        codecar = Trim(ws.Cells(ligne, "C").Value)
        Set celluleResultat = ws.Cells(ligne, "Q")

        If codecar <> "" Then
            resultat = RepertoireIPR(codecar)
            celluleResultat.Value = resultat

            ' Appliquer fond rouge uniquement si "pas_trouvé"
'            If resultat = "pas_trouvé" Then
'                celluleResultat.Interior.Color = vbRed
'            End If
'            ' Sinon : ne rien faire (laisser la couleur en place)
        Else
            celluleResultat.Value = ""
            ' Ne rien changer à la couleur si cellule vide
        End If
    Next ligne

End Sub


Function RepertoireIPR(codecar As String) As String
    Dim dossierList As Variant
    Dim dossierName As Variant
    Dim basePath As String
    Dim fichExcel As String, fichWord As String

    ' Remplacer / par -
    codecar = Replace(codecar, "/", "-")

    ' Liste des sous-dossiers et les mots-clés à retourner
    dossierList = Array( _
        Array("0- IPR VALIDE", "VALIDE"), _
        Array("1- IPR AUTORISEES", "AUTORISEES"), _
        Array("2- IPR en COURS", "EN_COURS"), _
        Array("3- IPR ARCHIVES", "ARCHIVES") _
    )

    ' Parcourir chaque dossier
    For Each dossierName In dossierList
        basePath = "S:\Methodes Production\" & dossierName(0) & "\"

        ' Chercher fichier Excel
        fichExcel = Dir(basePath & codecar & ".xls*")
        If fichExcel <> "" Then
            RepertoireIPR = dossierName(1)
            Exit Function
        End If

        ' Chercher fichier Word
        fichWord = Dir(basePath & codecar & ".doc*")
        If fichWord <> "" Then
            RepertoireIPR = dossierName(1)
            Exit Function
        End If
    Next dossierName

    ' Aucun fichier trouvé
    RepertoireIPR = "pas_trouvé"
End Function


Sub FiltrerOperateur(op As String)
    With ThisWorkbook.Worksheets("Planning")
        ' 1. Se positionner sur la cellule A1 pour ne pas être décalé après sélection
        .Range("A1").Select
        Application.Goto Reference:=Range("A1"), Scroll:=True
        
        ' 2. Trouve la dernière ligne non vide de la colonne B
        Dim lastRow As Long
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        
        ' 3. Définit la plage de données de A1 à B:lastRow
        Dim rngData As Range
        Set rngData = .Range("A1:A" & lastRow)
        
        ' 4. Applique le filtre sur la 1ère colonne de rngData (c'est la colonne A)
        rngData.AutoFilter Field:=1, Criteria1:="=*" & op & "*"
    End With
End Sub


Sub FiltrerParForme()
    Dim ws As Worksheet, op As String
    Set ws = ThisWorkbook.Worksheets("Planning")
    
    ' Application.Caller renvoie le nom de la forme ou du bouton FormControl
    op = ws.Shapes(Application.Caller).TextFrame.Characters.Text
    
    FiltrerOperateur op
End Sub

Sub EffacerFiltreColA()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Planning")
    
    With ws
        ' 1. Si les flèches de filtre sont présentes
        If .AutoFilterMode Then
            ' 2. On récupère la plage AutoFilter et...
            ' 3. ... on efface le critère de la 1ère colonne (A)
            .AutoFilter.Range.AutoFilter Field:=1
        End If
    End With
End Sub

Sub EffacerVidesColA()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Planning")
    
    With ws
        If .AutoFilterMode Then
            ' Réapplique AutoFilter sur la 1ère colonne sans critère = supprime le filtre "vide"
            .AutoFilter.Range.AutoFilter Field:=1, Criteria1:="<>"
        End If
    End With
End Sub









