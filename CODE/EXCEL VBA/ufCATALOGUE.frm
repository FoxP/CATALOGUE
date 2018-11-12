VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufCATALOGUE 
   Caption         =   "Cr�ation d'une fiche"
   ClientHeight    =   9570.001
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11910
   OleObjectBlob   =   "ufCATALOGUE.frx":0000
End
Attribute VB_Name = "ufCATALOGUE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================================================================================
'
'                                                               CATALOGUE
'
'Auteur :
'           Paul RENARD
'
'Fonction :
'           Catalogue Excel / HTML de fonctions, morceaux de codes, astuces, m�thodes, fichiers, ... fr�quemment utilis�s
'           Capitalise le savoir tout en fonctionnant sur des syst�mes d'exploitation avec des politiques de s�curit� strictes
'           N�cessite Microsoft Excel (2010 ou sup�rieur), Visual Basic for Applications, ainsi qu'un navigateur internet r�cent
'
'D�pendances VBA :
'           - Microsoft Scripting Runtime
'           - Visual Basic for Applications
'           - Microsoft Forms XX Object Library
'           - Microsoft Excel XX Object Library
'           - Microsoft Office XX Object Library
'           - Microsoft VBScript Regular Expressions 5.5
'
'===============================================================================================================================================

Option Explicit

'El�ment du UserForm qui a eu le focus en dernier
Public focusedControl As Control

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es au formatage Markdown des champs "Probl�me" et "Solution" : gras, italique, barr�, soulign�, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Renseignement de l'�l�ment du UserForm qui a eu le focus en dernier : Set focusedControl = ...
'-----------------------------------------------------------------------------------------------------------------------------------------------

'Lors le champ "Probl�me" est actif, s'il n'est pas maximis� (panel de pi�ces jointes cach�) :
'On agrandit l�g�rement le champ, et modifie la position / taille des champs "Solution" et "Code" en cons�quence
Private Sub tbProblem_Enter()
    If frameFiles.Visible = True Then
        Me.tbProblem.Height = 195   '100
        Me.tbProblem.Top = 66       '66
        Me.tbSolution.Height = 55   '100
        Me.tbSolution.Top = 265     '170
        Me.tbCode.Height = 55       '105
        Me.tbCode.Top = 325         '275
        
        lbProblem.Top = 66 + 88     '108
        lbSolution.Top = 265 + 18   '210
        lbCode.Top = 325 + 18       '318
    End If
    Set focusedControl = Me.tbProblem
End Sub

'Lors le champ "Solution" est actif, s'il n'est pas maximis� (panel de pi�ces jointes cach�) :
'On agrandit l�g�rement le champ, et modifie la position / taille des champs "Probl�me" et "Code" en cons�quence
Private Sub tbSolution_Enter()
    If frameFiles.Visible = True Then
        Me.tbProblem.Height = 55    '100
        Me.tbProblem.Top = 66       '66
        Me.tbSolution.Height = 194  '100
        Me.tbSolution.Top = 126     '170
        Me.tbCode.Height = 55       '105
        Me.tbCode.Top = 325         '275
        
        lbProblem.Top = 66 + 18     '108
        lbSolution.Top = 126 + 88   '210
        lbCode.Top = 325 + 18       '318
    End If
    Set focusedControl = Me.tbSolution
End Sub

'Lors le champ "Code" est actif, s'il n'est pas maximis� (panel de pi�ces jointes cach�) :
'On agrandit l�g�rement le champ, et modifie la position / taille des champs "Probl�me" et "Solution" en cons�quence
Private Sub tbCode_Enter()
    If frameFiles.Visible = True Then
        Me.tbProblem.Height = 55    '100
        Me.tbProblem.Top = 66       '66
        Me.tbSolution.Height = 55   '100
        Me.tbSolution.Top = 126     '170
        Me.tbCode.Height = 194      '105
        Me.tbCode.Top = 186         '275
        
        lbProblem.Top = 66 + 18     '108
        lbSolution.Top = 126 + 18   '210
        lbCode.Top = 186 + 88       '318
    End If
    Set focusedControl = Me.tbCode
End Sub

Private Sub tbTitle_Enter()
    Set focusedControl = Me.tbTitle
End Sub

Private Sub tbVersion_Enter()
    Set focusedControl = Me.tbVersion
End Sub

Private Sub tbID_Enter()
    Set focusedControl = Me.tbID
End Sub

Private Sub tbKeywords_Enter()
    Set focusedControl = Me.tbKeywords
End Sub

Private Sub cbSoftware_Enter()
    Set focusedControl = Me.cbSoftware
End Sub

Private Sub cbType_Enter()
    Set focusedControl = Me.cbType
End Sub

Private Sub cbStatus_Enter()
    Set focusedControl = Me.cbStatus
End Sub

Private Sub cbLanguage_Enter()
    Set focusedControl = Me.cbLanguage
End Sub

Private Sub fileList_Enter()
    Set focusedControl = Me.fileList
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Formatage Markdown gras pour les champs "Probl�me" et "Solution"
'Ajoute "__" avant / apr�s la portion de texte s�lectionn�e
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbBold_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        Call addTagToSelectedText(focusedControl, "__", "__")
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Formatage HTML citation pour les champs "Probl�me" et "Solution"
'Ajoute "<blockquote>" avant et "</blockquote>" apr�s la portion de texte s�lectionn�e
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbQuote_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        Call addTagToSelectedText(focusedControl, "<blockquote>", "</blockquote>")
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Convertit les caract�res sp�ciaux en entit�s HTML et inversement pour les champs "Probl�me" et "Solution"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbHtmlSpecialChars_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        If Len(focusedControl.SelText) > 0 Then
            Dim lPos As Long
            lPos = focusedControl.SelStart
            Dim lLength As Long
            If InStr(focusedControl.SelText, "<") <> 0 Or InStr(focusedControl.SelText, ">") <> 0 Or InStr(focusedControl.SelText, "'") <> 0 Or InStr(focusedControl.SelText, """") <> 0 Then
                lLength = Len(htmlSpecialChars(focusedControl.SelText))
                focusedControl.SelText = htmlSpecialChars(focusedControl.SelText)
                focusedControl.SelStart = lPos
                focusedControl.SelLength = lLength
            ElseIf InStr(focusedControl.SelText, "&lt;") <> 0 Or InStr(focusedControl.SelText, "&gt;") <> 0 Or InStr(focusedControl.SelText, "&#039;") <> 0 Or InStr(focusedControl.SelText, "&quot;") <> 0 Then
                lLength = Len(specialCharsHtml(focusedControl.SelText))
                focusedControl.SelText = specialCharsHtml(focusedControl.SelText)
                focusedControl.SelStart = lPos
                focusedControl.SelLength = lLength
            End If
        End If
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Statistiques pour les champs "Probl�me", "Solution", "Code" et "Mots-cl�s" : nombre de mots, caract�res et lignes
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbStats_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Or Me.tbCode Is focusedControl Or Me.tbKeywords Is focusedControl Or Me.tbTitle Is focusedControl Then
        Dim regWords As New VBScript_RegExp_55.RegExp
        regWords.Pattern = "\S+"
        regWords.IgnoreCase = True
        regWords.MultiLine = True
        regWords.Global = True
        Dim lWords As Long
        lWords = regWords.Execute(focusedControl.Text).Count()
        Dim lLines As Long
        lLines = UBound(Split(focusedControl.Text, vbCrLf)) + 1
        Dim lChars As Long
        lChars = Len(focusedControl.Text)
        MsgBox "- Mots : " & lWords & vbNewLine & "- Lignes : " & lLines & vbNewLine & "- Caract�res : " & lChars, vbInformation, "Statistiques"
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Maximise la taille du champ "Probl�me", "Solution" ou "Code"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbMaximizeTextBox_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Or Me.tbCode Is focusedControl Then
        'Masque le panel de pi�ces jointes
        frameFiles.Visible = False
        'Maximise le champ actif
        focusedControl.Top = 66
        focusedControl.Height = 373
        focusedControl.Width = 824
        'Si le panel de pi�ces jointes n'�tait pas ouvert, on agrandit / recentre la fen�tre
        If Me.Width < 900 Then
            With Me
                .Width = 900
                .Left = Me.Left - ((900 - 600) / 2)
            End With
        End If
        'Masque les champs / labels inutiles selon le cas
        If Me.tbProblem Is focusedControl Then
            Me.tbSolution.Visible = False
            Me.tbCode.Visible = False
            Me.lbProblem.Top = 220
            Me.lbSolution.Visible = False
            Me.lbCode.Visible = False
            Me.tbProblem.ZOrder (0)
        ElseIf Me.tbSolution Is focusedControl Then
            Me.tbProblem.Visible = False
            Me.tbCode.Visible = False
            Me.lbSolution.Top = 220
            Me.lbProblem.Visible = False
            Me.lbCode.Visible = False
            Me.tbSolution.ZOrder (0)
        Else
            Me.tbProblem.Visible = False
            Me.tbSolution.Visible = False
            Me.lbCode.Top = 220
            Me.lbProblem.Visible = False
            Me.lbSolution.Visible = False
            Me.tbCode.ZOrder (0)
        End If
        lbKeywords.Visible = False
        lbLanguage.Visible = False
        cbLanguage.Visible = False
        cbSoftware.Visible = False
        cbType.Visible = False
        tbID.Enabled = False
        tbVersion.Enabled = False
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Minimise la taille du champ "Probl�me", "Solution" ou "Code"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbMinimizeTextBox_Click()
    'En cas de cr�ation de fiche, le champ "Titre" peut �tre actif : v�rifions
    Dim bIsTitleFocused As Boolean
    If Me.tbTitle Is Me.ActiveControl Then
        bIsTitleFocused = True
    End If
    'R�affiche le panel de pi�ces jointes
    frameFiles.Visible = True
    'R�affiche les champs / labels selon le cas, minimise le champ maximis�
    If Me.tbProblem.Width > 524 Then
        Me.tbProblem.Width = 524
        Me.tbSolution.Visible = True
        Me.tbCode.Visible = True
        Call tbProblem_Enter
        Me.lbSolution.Visible = True
        Me.lbCode.Visible = True
    ElseIf Me.tbSolution.Width > 524 Then
        Me.tbSolution.Width = 524
        Me.tbProblem.Visible = True
        Me.tbCode.Visible = True
        Call tbSolution_Enter
        Me.lbProblem.Visible = True
        Me.lbCode.Visible = True
    ElseIf Me.tbCode.Width > 524 Then
        Me.tbCode.Width = 524
        Me.tbProblem.Visible = True
        Me.tbSolution.Visible = True
        Call tbCode_Enter
        Me.lbProblem.Visible = True
        Me.lbSolution.Visible = True
    End If
    lbKeywords.Visible = True
    lbLanguage.Visible = True
    cbLanguage.Visible = True
    cbSoftware.Visible = True
    cbType.Visible = True
    tbID.Enabled = True
    tbVersion.Enabled = True
    'Si le champ titre �tait actif, on r�applique cet �tat
    If bIsTitleFocused Then
        Set focusedControl = Me.tbTitle
        Me.tbTitle.SetFocus
    Else
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ajout d'un s�parateur horizontal Markdown pour les champs "Probl�me" et "Solution"
'Ajoute "---" apr�s la position actuelle du curseur
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbSeparator_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        'Position du curseur
        Dim lPos As Long
        lPos = focusedControl.SelStart
        'Caract�re apr�s la position du curseur
        focusedControl.SelStart = lPos
        focusedControl.SelLength = 1
        Dim sAfter As String
        sAfter = focusedControl.SelText
        'Caract�re avant la position du curseur
        If lPos <> 0 Then
            focusedControl.SelStart = lPos - 1
            focusedControl.SelLength = 1
        Else
            focusedControl.SelStart = lPos
            focusedControl.SelLength = 0
        End If
        Dim sBefore As String
        sBefore = focusedControl.SelText
        'RAZ de la position du curseur
        focusedControl.SelStart = lPos
        focusedControl.SelLength = 0
        'Selon le cas, ajoute "---" et le bon nombre de sauts de lignes, avec positionnement du curseur
        If sAfter = vbCr And sBefore = vbCr Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & "---" & vbNewLine
            focusedControl.SelStart = focusedControl.SelStart + 1
        ElseIf sAfter = vbCr And sBefore = "" Then
            focusedControl.SelText = focusedControl.SelText & "---" & vbNewLine
            focusedControl.SelStart = focusedControl.SelStart + 1
        ElseIf sAfter = vbCr And sBefore <> "" Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & vbNewLine & "---" & vbNewLine
            focusedControl.SelStart = focusedControl.SelStart + 1
        ElseIf sAfter = "" And sBefore = vbCr Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & "---" & vbNewLine & vbNewLine
        ElseIf sAfter = "" And sBefore <> "" Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & vbNewLine & "---" & vbNewLine & vbNewLine
        ElseIf sAfter = "" And sBefore = "" Then
            focusedControl.SelText = focusedControl.SelText & "---" & vbNewLine & vbNewLine
        ElseIf sAfter <> "" And sBefore = vbCr Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & "---" & vbNewLine & vbNewLine
        ElseIf sAfter <> "" And sBefore = "" Then
            focusedControl.SelText = focusedControl.SelText & "---" & vbNewLine & vbNewLine
        ElseIf sAfter <> "" And sBefore <> "" Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & vbNewLine & "---" & vbNewLine & vbNewLine
        Else
            focusedControl.SelText = focusedControl.SelText & vbNewLine & vbNewLine & "---" & vbNewLine & vbNewLine
        End If
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ajout d'un tableau Markdown pour les champs "Probl�me" et "Solution"
'Ajoute la s�quence suivante apr�s la position actuelle du curseur :
'
'   Colonne 1 | Colonne 2 | Colonne 3
'   --- | --- | ---
'   L.1 Col.1 | L.1 Col.2 | L.1 Col.3
'   L.2 Col.1 | L.2 Col.2 | L.2 Col.3
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbTable_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        'Position du curseur
        Dim lPos As Long
        lPos = focusedControl.SelStart
        'Caract�re apr�s la position du curseur
        focusedControl.SelStart = lPos
        focusedControl.SelLength = 1
        Dim sAfter As String
        sAfter = focusedControl.SelText
        'Caract�re avant la position du curseur
        If lPos <> 0 Then
            focusedControl.SelStart = lPos - 1
            focusedControl.SelLength = 1
        Else
            focusedControl.SelStart = lPos
            focusedControl.SelLength = 0
        End If
        Dim sBefore As String
        sBefore = focusedControl.SelText
        'RAZ de la position du curseur
        focusedControl.SelStart = lPos
        focusedControl.SelLength = 0
        'G�n�ration du tableau Markdown
        Dim sMarkdownTable As String
        Dim lColsNbr As Long
        Dim lRowsNbr As Long
        ufMarkdownTable.Show
        If bCancelTable = True Then
            Exit Sub
        End If
        lColsNbr = ufMarkdownTable.tbColsNbr
        lRowsNbr = ufMarkdownTable.tbRowsNbr
        Dim i As Long, j As Long
        For i = 2 To lColsNbr
            If i = 2 Then
                sMarkdownTable = "Colonne " & i - 1 & " | Colonne " & i
            Else
                sMarkdownTable = sMarkdownTable & " | Colonne " & i
            End If
        Next
        sMarkdownTable = sMarkdownTable & vbNewLine & "--- | --- | ---" & vbNewLine
        For j = 1 To lRowsNbr
            For i = 1 To lColsNbr
                If i = lColsNbr Then
                    If j = lRowsNbr Then
                        sMarkdownTable = sMarkdownTable & "L." & j & " Col." & i
                    Else
                        sMarkdownTable = sMarkdownTable & "L." & j & " Col." & i & vbNewLine
                    End If
                Else
                    sMarkdownTable = sMarkdownTable & "L." & j & " Col." & i & " | "
                End If
            Next
        Next
        'Selon le cas, ajoute le bon nombre de sauts de lignes, avec positionnement du curseur
        If sAfter = vbCr And sBefore = vbCr Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & sMarkdownTable & vbNewLine
            focusedControl.SelStart = focusedControl.SelStart + 1
        ElseIf sAfter = vbCr And sBefore = "" Then
            focusedControl.SelText = focusedControl.SelText & sMarkdownTable & vbNewLine
            focusedControl.SelStart = focusedControl.SelStart + 1
        ElseIf sAfter = vbCr And sBefore <> "" Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & vbNewLine & sMarkdownTable & vbNewLine
            focusedControl.SelStart = focusedControl.SelStart + 1
        ElseIf sAfter = "" And sBefore = vbCr Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & sMarkdownTable & vbNewLine & vbNewLine
        ElseIf sAfter = "" And sBefore <> "" Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & vbNewLine & sMarkdownTable & vbNewLine & vbNewLine
        ElseIf sAfter = "" And sBefore = "" Then
            focusedControl.SelText = focusedControl.SelText & sMarkdownTable & vbNewLine & vbNewLine
        ElseIf sAfter <> "" And sBefore = vbCr Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & sMarkdownTable & vbNewLine & vbNewLine
        ElseIf sAfter <> "" And sBefore = "" Then
            focusedControl.SelText = focusedControl.SelText & sMarkdownTable & vbNewLine & vbNewLine
        ElseIf sAfter <> "" And sBefore <> "" Then
            focusedControl.SelText = focusedControl.SelText & vbNewLine & vbNewLine & sMarkdownTable & vbNewLine & vbNewLine
        Else
            focusedControl.SelText = focusedControl.SelText & vbNewLine & vbNewLine & sMarkdownTable & vbNewLine & vbNewLine
        End If
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ajout d'un lien Markdown pour les champs "Probl�me" et "Solution"
'Format : [Texte du lien](http://adresse-du-lien.com)
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbLink_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        'Position du curseur
        Dim lPos As Long
        lPos = focusedControl.SelStart
        'Configuration du UserForm ufMarkdownLink
        If Len(focusedControl.SelText) > 0 Then
            'Du texte est s�lectionn�, il deviendra le texte du lien
            ufMarkdownLink.tbText.Value = focusedControl.SelText
        Else
            'Pas de texte s�lectionn�, on vide la TextBox li�e
            ufMarkdownLink.tbText.Value = ""
        End If
        ufMarkdownLink.tbLink.Value = ""
        ufMarkdownLink.Show
        If bCancelLink = True Then
            Exit Sub
        End If
        'D�finition de l'expression r�guli�re de d�tection d'URL de fiche
        Dim regSheetURL As New VBScript_RegExp_55.RegExp
        regSheetURL.Pattern = ".*[\/\\]SHEETS*[\/\\]([0-9]{6}_.*_[0-9]+\.html)"
        Dim sURL As String
        'Si c'est une URL de fiche, on la transforme en lien relatif
        If regSheetURL.test(ufMarkdownLink.tbLink.Value) Then
            sURL = regSheetURL.Execute(ufMarkdownLink.tbLink.Value)(0).SubMatches(0)
        Else
            sURL = ufMarkdownLink.tbLink.Value
        End If
        'Si du texte est s�lectionn�, on construit le balisage Markdown autour
        If Len(focusedControl.SelText) > 0 Then
            Dim sSelectedText As String
            sSelectedText = focusedControl.SelText
            'Si l'utilisateur ne souhaite pas modifier le texte s�lectionn�
            If ufMarkdownLink.tbText.Value = sSelectedText Then
                Call addTagToSelectedText(focusedControl, "[", "](" & sURL & ")")
            'Si l'utilisateur souhaite modifier le texte s�lectionn�
            Else
                focusedControl.SelText = "[" & ufMarkdownLink.tbText.Value & "](" & sURL & ")"
                focusedControl.SelStart = lPos + 1
                focusedControl.SelLength = Len(ufMarkdownLink.tbText.Value)
            End If
        'Si pas de texte s�lectionn�, on construit le balisage Markdown � la position actuelle du curseur
        Else
            focusedControl.SelText = focusedControl.SelText & "[" & ufMarkdownLink.tbText.Value & "](" & sURL & ")"
        End If
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'A partir des pi�ces jointes, ajout d'un lien ou d'une image Markdown pour les champs "Probl�me" et "Solution"
'Format : [Nom du document](url-du-document) si document ou ![Nom de l'image](url-de-l'image) si image
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbFile_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        If strFile.Count > 0 Then
            'Pas de copie si titre de la fiche non d�fini
            If delAllSpace(Me.tbTitle) <> "" Then
                'Si le titre de la fiche peut �tre modifi� (cr�ation de fiche), on demande � le verrouiller
                'Si l'utilisateur change le nom de la fiche apr�s l'insertion d'images, les liens seront bris�s
                If Me.tbTitle.Enabled = True Then
                    If MsgBox("Le titre de la fiche va �tre verrouill� et ne pourra plus �tre modifi�, souhaitez vous continuer ?", vbYesNo + vbQuestion, "Continuer ?") = vbNo Then
                        Exit Sub
                    End If
                End If
                'Verrouillage du titre si cr�ation de fiche
                Me.tbTitle.Enabled = False
                'Position du curseur
                Dim lPos As Long
                lPos = focusedControl.SelStart
                'Ajout des noms de pi�ces jointes tri�s par ordre alphab�tique dans la ComboBox
                ufMarkdownFile.cbFile.Clear
                Dim strFileTemp As New Collection
                Set strFileTemp = strFile
                Call sortCollection(strFileTemp)
                Dim sFilePath As Variant
                For Each sFilePath In strFileTemp
                    ufMarkdownFile.cbFile.AddItem cleanNameForMarkdown(getFilenameFromPath(sFilePath))
                Next
                'Pour chaque entr�e de la ListBox de pi�ces jointes
                Dim i As Integer
                For i = Me.fileList.ListCount - 1 To 0 Step -1
                    'Si l'entr�e est s�lectionn�e, on fait de m�me pour la ComboBox
                    If Me.fileList.Selected(i) = True Then
                        ufMarkdownFile.cbFile.Value = Me.fileList.List(i)
                        'Uniquement pour la 1i�re entr�e s�lectionn�e
                        Exit For
                    End If
                Next
                'Si aucune entr�e n'�tait s�lectionn�e dans la ListBox de pi�ces jointes
                If ufMarkdownFile.cbFile.Value = "" Then
                    'S'il y a des pi�ces jointes, on s�lectionne la 1i�re dans la ComboBox
                    If strFileTemp.Count > 0 Then
                        ufMarkdownFile.cbFile.Value = cleanNameForMarkdown(getFilenameFromPath(strFileTemp(1)))
                    End If
                End If
                'Configuration du UserForm ufMarkdownFile
                If Len(focusedControl.SelText) > 0 Then
                    'Du texte est s�lectionn�, il deviendra la description de la pi�ce jointe
                    ufMarkdownFile.tbDescription.Value = focusedControl.SelText
                Else
                    'Pas de texte s�lectionn�, le nom de la pi�ce jointe deviendra sa description
                    ufMarkdownFile.tbDescription.Value = ufMarkdownFile.cbFile.Value
                End If
                ufMarkdownFile.Show
                'Si l'utilisateur a cliqu� sur "Annuler" ou ferm� le UserForm, on quitte
                If bCancelFile = True Then
                    Exit Sub
                End If
                'URL de la pi�ce jointe en relatif
                Dim sFileURL As String
                sFileURL = filesPath & "\" & replaceIllegalChar(tbID & "_" & Me.tbTitle & "_" & Me.tbVersion) & "\" & ufMarkdownFile.cbFile.Value
                'Si du texte est s�lectionn�, on construit le balisage Markdown autour
                If Len(focusedControl.SelText) > 0 Then
                    Dim sSelectedText As String
                    sSelectedText = focusedControl.SelText
                    'Si l'utilisateur ne souhaite pas modifier le texte s�lectionn�
                    If ufMarkdownFile.tbDescription.Value = sSelectedText Then
                        If isInArray(LCase(getExtensionFromPath(ufMarkdownFile.cbFile.Value)), Split(imgArray, ",")) Then
                            Call addTagToSelectedText(focusedControl, "![", "](" & sFileURL & ")")
                        Else
                            Call addTagToSelectedText(focusedControl, "[", "](" & sFileURL & ")")
                        End If
                    'Si l'utilisateur souhaite modifier le texte s�lectionn�
                    Else
                        If isInArray(LCase(getExtensionFromPath(ufMarkdownFile.cbFile.Value)), Split(imgArray, ",")) Then
                            focusedControl.SelText = "![" & ufMarkdownFile.tbDescription.Value & "](" & sFileURL & ")"
                            focusedControl.SelStart = lPos + 2
                        Else
                            focusedControl.SelText = "[" & ufMarkdownFile.tbDescription.Value & "](" & sFileURL & ")"
                            focusedControl.SelStart = lPos + 1
                        End If
                        focusedControl.SelLength = Len(ufMarkdownFile.tbDescription.Value)
                    End If
                'Si pas de texte s�lectionn�, on construit le balisage Markdown � la position actuelle du curseur
                Else
                    If isInArray(LCase(getExtensionFromPath(ufMarkdownFile.cbFile.Value)), Split(imgArray, ",")) Then
                        focusedControl.SelText = focusedControl.SelText & "![" & ufMarkdownFile.tbDescription.Value & "](" & sFileURL & ")"
                    Else
                        focusedControl.SelText = focusedControl.SelText & "[" & ufMarkdownFile.tbDescription.Value & "](" & sFileURL & ")"
                    End If
                End If
                focusedControl.SetFocus
            Else
                MsgBox "Le titre de la fiche doit �tre renseign� avant de pouvoir ins�rer le code Markdown d'une pi�ce jointe.", vbInformation, "Titre invalide"
            End If
        Else
            MsgBox "Des pi�ces jointes doivent �tre rattach�es � la fiche avant de pouvoir ins�rer leur code Markdown.", vbInformation, "Aucune pi�ce jointe"
        End If
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Balise Markdown code pour les champs "Probl�me" et "Solution"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbCode_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        'Si texte s�lectionn� sur plusieurs lignes
        If UBound(Split(focusedControl.SelText, vbCr)) > 0 Then
            Call addTagToSelectedText(focusedControl, "```" & vbNewLine, vbNewLine & "```")
            focusedControl.SetFocus
        'Sinon, texte s�lectionn� sur une seule ligne
        Else
            Call addTagToSelectedText(focusedControl, "`", "`")
            focusedControl.SetFocus
        End If
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Formatage MAJUSCULE pour les champs "Probl�me", "Solution", "Code" et "Mots-cl�s"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbMaj_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Or Me.tbCode Is focusedControl Or Me.tbKeywords Is focusedControl Then
        If Len(focusedControl.SelText) > 0 Then
            Dim lPos As Long
            lPos = focusedControl.SelStart
            Dim lLength As Long
            lLength = focusedControl.SelLength
            focusedControl.SelText = UCase(focusedControl.SelText)
            focusedControl.SelStart = lPos
            focusedControl.SelLength = lLength
            focusedControl.SetFocus
        End If
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Formatage minuscule pour les champs "Probl�me", "Solution", "Code" et "Mots-cl�s"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbMin_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Or Me.tbCode Is focusedControl Or Me.tbKeywords Is focusedControl Then
        If Len(focusedControl.SelText) > 0 Then
            Dim lPos As Long
            lPos = focusedControl.SelStart
            Dim lLength As Long
            lLength = focusedControl.SelLength
            focusedControl.SelText = LCase(focusedControl.SelText)
            focusedControl.SelStart = lPos
            focusedControl.SelLength = lLength
            focusedControl.SetFocus
        End If
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Formatage HTML surlign� pour les champs "Probl�me" et "Solution"
'Ajoute "<mark>" avant et "</mark>" apr�s la portion de texte s�lectionn�e
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbMark_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        Call addTagToSelectedText(focusedControl, "<mark>", "</mark>")
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Formatage HTML couleur pour les champs "Probl�me" et "Solution"
'Ajoute "<span style="color:#HEX_COLOR">" avant et "</span>" apr�s la portion de texte s�lectionn�e
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbColor_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        If Len(focusedControl.SelText) > 0 Then
            Call addTagToSelectedText(focusedControl, "<span style=""" & "color:#" & decimalColor2Hex(pickNewColor) & """>", "</span>")
            focusedControl.SetFocus
        End If
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Formatage Markdown italique pour les champs "Probl�me" et "Solution"
'Ajoute "_" avant / apr�s la portion de texte s�lectionn�e
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbItalic_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        Call addTagToSelectedText(focusedControl, "_", "_")
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Formatage Markdown barr� pour les champs "Probl�me" et "Solution"
'Ajoute "~~" avant / apr�s la portion de texte s�lectionn�e
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbStrike_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        Call addTagToSelectedText(focusedControl, "~~", "~~")
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Formatage Markdown soulign� pour les champs "Probl�me" et "Solution"
'Ajoute "<u>" avant et "</u>" apr�s la portion de texte s�lectionn�e
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbUnderline_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Then
        Call addTagToSelectedText(focusedControl, "<u>", "</u>")
        focusedControl.SetFocus
    End If
End Sub

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es aux champs "Code", "Probl�me" et "Solution" : copie du contenu du presse papier, r�initialisation, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Copie du contenu du presse papier dans le champ "Code", "Probl�me" ou "Solution"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbPasteCode_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Or Me.tbCode Is focusedControl Then
        Dim sTextBoxName As String
        If Me.tbProblem Is focusedControl Then
            sTextBoxName = "Probl�me"
        ElseIf Me.tbSolution Is focusedControl Then
            sTextBoxName = "Solution"
        Else
            sTextBoxName = "Code"
        End If
        If Replace(focusedControl.Value, " ", "") <> "" Then
            If MsgBox("Le champ """ & sTextBoxName & """ n'est pas vide, souhaitez vous le remplacer par le contenu du presse papier ? Les donn�es existantes seront perdues.", vbYesNo + vbQuestion, "Continuer ?") = vbNo Then
                focusedControl.SetFocus
                Exit Sub
            End If
        End If
        focusedControl.Value = getClipboardContent()
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'R�initialisation du contenu du champ "Code", "Probl�me" ou "Solution"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbReset_Click()
    If Me.tbProblem Is focusedControl Or Me.tbSolution Is focusedControl Or Me.tbCode Is focusedControl Then
        Dim sTextBoxName As String
        If Me.tbProblem Is focusedControl Then
            sTextBoxName = "Probl�me"
        ElseIf Me.tbSolution Is focusedControl Then
            sTextBoxName = "Solution"
        Else
            sTextBoxName = "Code"
        End If
        If Replace(focusedControl.Value, " ", "") <> "" Then
            If MsgBox("Le champ """ & sTextBoxName & """ n'est pas vide, souhaitez vous vraiment effacer la totalit� de son contenu ? Les donn�es existantes seront perdues.", vbYesNo + vbQuestion, "Continuer ?") = vbNo Then
                focusedControl.SetFocus
                Exit Sub
            End If
        End If
        focusedControl.Value = ""
        focusedControl.SetFocus
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Affichage de la table des caract�res de Windows
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbCharMap_Click()
    Call showCharMap
    focusedControl.SetFocus
End Sub

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es aux pi�ces jointes : tri, ajout, suppression, ouverture, copie du chemin dans le presse papier, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Tri des pi�ces jointes de la ListView par ordre alphab�tique
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbSortFiles_Click()
    Call sortListBox(Me.fileList)
    Set focusedControl = Me.cbSortFiles
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ouverture du dossier des pi�ces jointes de la fiche
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbOpenFilesFolder_Click()
    'Si le dossier des pi�ces jointes de la fiche existe
    If folderExists(cataloguePath & sheetsPath & filesPath & "\" & replaceIllegalChar(Me.tbID & "_" & Me.tbTitle & "_" & Me.tbVersion)) Then
        Shell Environ("WINDIR") & "\explorer.exe " & cataloguePath & sheetsPath & filesPath & "\" & replaceIllegalChar(Me.tbID & "_" & Me.tbTitle & "_" & Me.tbVersion), vbNormalFocus
    Else
        MsgBox "Le dossier de pi�ces jointes n'existe pas (ou pas encore).", vbInformation, "Aucune pi�ce jointe"
    End If
    Set focusedControl = Me.cbOpenFilesFolder
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur appuie sur "Ajouter" : Ajout d'une ou de plusieurs pi�ces jointes � la fiche
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbAddFile_Click()

    Dim intChoice As Integer
    Dim i As Integer
    
    'S�lection de plusieurs fichiers autoris�e
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
    'Afficher la boite de dialogue de s�lection de fichier
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    'Si des fichiers ont �t� s�lectionn�s
    If intChoice <> 0 Then
        'Pour chaque fichier s�lectionn�
        For i = 1 To Application.FileDialog(msoFileDialogOpen).SelectedItems.Count
            'Un fichier ne peut pas �tre ajout� deux fois
            If Not keyExistsInColl(strFile, cleanNameForMarkdown(getFilenameFromPath(Application.FileDialog(msoFileDialogOpen).SelectedItems(i)))) Then
                'Ajout du nom du fichier dans la ListBox
                Me.fileList.AddItem cleanNameForMarkdown(getFilenameFromPath(Application.FileDialog(msoFileDialogOpen).SelectedItems(i)))
                'Ajout du chemin du fichier dans la collection de pi�ces jointes � copier
                'La cl� du fichier dans la collection est le nom du fichier lui-m�me
                strFile.Add Application.FileDialog(msoFileDialogOpen).SelectedItems(i), cleanNameForMarkdown(getFilenameFromPath(Application.FileDialog(msoFileDialogOpen).SelectedItems(i)))
            End If
        Next i
        lbFilesCpt.Caption = strFile.Count
    End If
    Set focusedControl = Me.cbAddFile

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur appuie sur "Supprimer" : Suppression d'une ou de plusieurs pi�ces jointes de la fiche
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbDelFile_Click()

    Dim i As Integer
    'Pour chaque entr�e de la ListBox
    For i = Me.fileList.ListCount - 1 To 0 Step -1
        'Si l'entr�e est s�lectionn�e
        If Me.fileList.Selected(i) = True Then
            'Suppression du fichier dans la collection de pi�ces jointes � supprimer
            'N�cessaire en cas d'�dition de fiche dans laquelle on supprime une pi�ce jointe existante
            strToDelete.Add strFile.item(Me.fileList.List(i)), cleanNameForMarkdown(getFilenameFromPath(strFile.item(Me.fileList.List(i))))
            'Suppression du fichier dans la collection de pi�ces jointes � copier
            strFile.Remove Me.fileList.List(i)
            'Suppression du fichier dans la ListBox
            Me.fileList.RemoveItem (i)
        End If
    Next
    lbFilesCpt.Caption = strFile.Count
    Set focusedControl = Me.cbDelFile

End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Copie du code Markdown d'insertion d'une image (issue des pi�ces jointes) dans le presse papier lors du double clic sur son nom dans la ListBox
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub fileList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    'Si au moins une pi�ce jointe
    If fileList.ListCount > 0 Then
        'Pas de copie si titre de la fiche non d�fini
        If delAllSpace(Me.tbTitle) <> "" Then
            'Si le titre de la fiche peut �tre modifi� (cr�ation de fiche), on demande � le verrouiller
            'Si l'utilisateur change le nom de la fiche apr�s l'insertion d'images, les liens seront bris�s
            If Me.tbTitle.Enabled = True Then
                If MsgBox("Le titre de la fiche va �tre verrouill� et ne pourra plus �tre modifi�, souhaitez vous continuer ?", vbYesNo + vbQuestion, "Continuer ?") = vbNo Then
                    Exit Sub
                End If
            End If
            'Verrouillage du titre si cr�ation de fiche
            Me.tbTitle.Enabled = False
            If isInArray(LCase(getExtensionFromPath(Me.fileList.List(Me.fileList.ListIndex))), Split(imgArray, ",")) Then
                'Image
                Call CopyToClipboard("![" & getFilenameFromPath(Me.fileList.List(Me.fileList.ListIndex)) & "](" & filesPath & "\" & replaceIllegalChar(tbID & "_" & Me.tbTitle & "_" & Me.tbVersion) & "\" & Me.fileList.List(Me.fileList.ListIndex) & ")")
            Else
                'Document
                Call CopyToClipboard("[" & getFilenameFromPath(Me.fileList.List(Me.fileList.ListIndex)) & "](" & filesPath & "\" & replaceIllegalChar(tbID & "_" & Me.tbTitle & "_" & Me.tbVersion) & "\" & Me.fileList.List(Me.fileList.ListIndex) & ")")
            End If
        Else
            MsgBox "Le titre de la fiche doit �tre renseign� avant de pouvoir copier un lien dans le presse-papier.", vbInformation, "Titre invalide"
        End If
    End If
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ouverture d'une pi�ce jointe (issue de la liste des pi�ces jointes) lors d'un clic sur son nom dans la ListBox avec la molette de la souris
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub fileList_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim i As Integer
    'Pour chaque entr�e de la ListBox
    For i = Me.fileList.ListCount - 1 To 0 Step -1
        'Si l'entr�e est s�lectionn�e
        If Me.fileList.Selected(i) = True Then
            'Si clic de la molette
            If Button = 4 Then
                'Ouverture de la pi�ce jointe
                OpenIt (strFile(Me.fileList.List(i)))
            End If
        End If
    Next
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Agrandit la largeur du UserForm pour afficher le panel de pi�ces jointes
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub hideShowFilesPanel_Click()

    'Le panel de pi�ces jointes est d�j� affich�
    If Me.Width = 900 Then
        'Masque le panel de pi�ces jointes
        With Me
            'Si le champ "Probl�me", "Solution" ou "Code" n'est pas maximis�, on r�duit la taille de la fen�tre
            If .frameFiles.Visible = True Then
                .Width = 600
                .Left = Me.Left + ((900 - 600) / 2)
                '.Left = Application.Left + Application.Width / 2 - .Width / 2
            Else
                'Le champ "Probl�me", "Solution" ou "Code" est maximis�, on le minimise pour afficher le panel de pi�ces jointes
                Call cbMinimizeTextBox_Click
            End If
        End With
    'Le panel de pi�ces jointes n'est pas affich�
    Else
        'Affiche le panel de pi�ces jointes
        With Me
            .Width = 900
            .Left = Me.Left - ((900 - 600) / 2)
            '.Left = Application.Left + Application.Width / 2 - .Width / 2
        End With
    End If
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Agrandit la largeur du UserForm pour afficher le panel de pi�ces jointes
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub lbFilesCpt_Click()

    'Le panel de pi�ces jointes est d�j� affich�
    If Me.Width = 900 Then
        'Masque le panel de pi�ces jointes
        With Me
            'Si le champ "Probl�me", "Solution" ou "Code" n'est pas maximis�, on r�duit la taille de la fen�tre
            If .frameFiles.Visible = True Then
                .Width = 600
                .Left = Me.Left + ((900 - 600) / 2)
                '.Left = Application.Left + Application.Width / 2 - .Width / 2
            Else
                'Le champ "Probl�me", "Solution" ou "Code" est maximis�, on le minimise pour afficher le panel de pi�ces jointes
                Call cbMinimizeTextBox_Click
            End If
        End With
    'Le panel de pi�ces jointes n'est pas affich�
    Else
        'Affiche le panel de pi�ces jointes
        With Me
            .Width = 900
            .Left = Me.Left - ((900 - 600) / 2)
            '.Left = Application.Left + Application.Width / 2 - .Width / 2
        End With
    End If

End Sub

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es � la cr�ation / �dition de fiches : pr�visualisation, v�rification de l'orthographe, aide au langage Markdown, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Aper�u format� en Markdown des champs "Probl�me", "Solution" et "Code"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbPreview_Click()

    'Si au moins un des champs "Probl�me", "Solution" et "Code" n'est pas vide
    If Not (delAllSpace(Me.tbProblem.Value) = "" And delAllSpace(Me.tbSolution.Value) = "" And delAllSpace(Me.tbCode.Value) = "") Then
        Dim iFilesCount As Integer
        Dim strProblemContent As String
        Dim strSolutionContent As String
        Dim bBadImgName As Boolean
        
        'Duplication du template HTML d�di� � la pr�visualisation : htmlPreviewName -> htmlPreviewTmpName
        Call copyFileFromTo(cataloguePath & sheetsPath & htmlPreviewName, cataloguePath & sheetsPath & htmlPreviewTmpName, True)
        Dim strIn 'As String
        'R�cup�ration du contenu du template HTML d�di� � la pr�visualisation
        strIn = getFileContent(cataloguePath & sheetsPath & htmlPreviewName)
        
        'Expression r�guli�re permettant de r�cup�rer l'url des images ins�r�es en Markdown
        'Quand le code Markdown d'une image est ins�r� depuis les pi�ces jointes mais que la fiche n'a pas encore �t� sauvegard�e :
        '   - la pi�ce jointe n'a pas encore �t� d�plac�e dans le dossier de pi�ces jointes de la fiche
        '   - � la pr�visualisation de la fiche, les urls des pi�ces jointes pas encore d�plac�es sont alors mauvaises
        'Solution : d�tecter l'url des images ins�r�es en Markdown, et les remplacer par leur emplacement avant d�placement
        Dim regMarkdownImg As New VBScript_RegExp_55.RegExp
        regMarkdownImg.Pattern = "(!\[.*?\]\()(.+?)(\))"
        'Plusieurs lignes, plusieurs r�sultats possibles
        regMarkdownImg.MultiLine = True
        regMarkdownImg.Global = True
        'Collection de r�sultats de l'expression r�guli�re
        Dim matchesColl As MatchCollection
        
        'Si le champ "Probl�me" n'est pas vide
        If delAllSpace(Me.tbProblem.Value) <> "" Then
            strProblemContent = Me.tbProblem.Value
            'Si des images ont �t� ins�r�es en Markdown
            If regMarkdownImg.test(strProblemContent) Then
                Set matchesColl = regMarkdownImg.Execute(strProblemContent)
                'Pour chaque image ins�r�e en Markdown d�tect�e
                For iFilesCount = 0 To matchesColl.Count - 1
                    'Si l'image a �t� ajout�e dans la collection de pi�ces jointes
                    If keyExistsInColl(strFile, cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strProblemContent)(iFilesCount).SubMatches(1)))) Then
                        'Si une des images dont le code Markdown a �t� ins�r� dans la fiche contient un des caract�res suivants : [ ] ( )
                        If InStr(strFile(cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strProblemContent)(iFilesCount).SubMatches(1)))), "[") <> 0 Or _
                        InStr(strFile(cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strProblemContent)(iFilesCount).SubMatches(1)))), "]") <> 0 Or _
                        InStr(strFile(cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strProblemContent)(iFilesCount).SubMatches(1)))), "(") <> 0 Or _
                        InStr(strFile(cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strProblemContent)(iFilesCount).SubMatches(1)))), ")") <> 0 Then
                            bBadImgName = True
                        End If
                        'Remplacement de son url par l'url issue de la collection de pi�ces jointes
                        strProblemContent = Replace(strProblemContent, regMarkdownImg.Execute(strProblemContent)(iFilesCount).SubMatches(1), strFile(cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strProblemContent)(iFilesCount).SubMatches(1)))))
                    End If
                Next
            End If
            'Remplacement de la balise "{TEMPLATE_PREVIEW_PROBLEM}" par le contenu du champ "Probl�me"
            strIn = Replace(strIn, "{TEMPLATE_PREVIEW_PROBLEM}", strProblemContent)
        Else
            'Suppression de la section
            strIn = Replace(strIn, GetBetween(strIn, "<section id=" & Chr(34) & "problem" & Chr(34) & ">", "</section>"), "")
        End If
        
        'Si le champ "Solution" n'est pas vide
        If delAllSpace(Me.tbSolution.Value) <> "" Then
            strSolutionContent = Me.tbSolution.Value
            'Si des images ont �t� ins�r�es en Markdown
            If regMarkdownImg.test(strSolutionContent) Then
                Set matchesColl = regMarkdownImg.Execute(strSolutionContent)
                'Pour chaque image ins�r�e en Markdown d�tect�e
                For iFilesCount = 0 To matchesColl.Count - 1
                    'Si l'image a �t� ajout�e dans la collection de pi�ces jointes
                    If keyExistsInColl(strFile, cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strSolutionContent)(iFilesCount).SubMatches(1)))) Then
                        'Si une des images dont le code Markdown a �t� ins�r� dans la fiche contient un des caract�res suivants : [ ] ( )
                        If InStr(strFile(cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strSolutionContent)(iFilesCount).SubMatches(1)))), "[") <> 0 Or _
                        InStr(strFile(cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strSolutionContent)(iFilesCount).SubMatches(1)))), "]") <> 0 Or _
                        InStr(strFile(cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strSolutionContent)(iFilesCount).SubMatches(1)))), "(") <> 0 Or _
                        InStr(strFile(cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strSolutionContent)(iFilesCount).SubMatches(1)))), ")") <> 0 Then
                            bBadImgName = True
                        End If
                        'Remplacement de son url par l'url issue de la collection de pi�ces jointes
                        strSolutionContent = Replace(strSolutionContent, regMarkdownImg.Execute(strSolutionContent)(iFilesCount).SubMatches(1), strFile(cleanNameForMarkdown(getFilenameFromPath(regMarkdownImg.Execute(strSolutionContent)(iFilesCount).SubMatches(1)))))
                    End If
                Next
            End If
            'Remplacement de la balise "{TEMPLATE_PREVIEW_SOLUTION}" par le contenu du champ "Solution"
            strIn = Replace(strIn, "{TEMPLATE_PREVIEW_SOLUTION}", strSolutionContent)
        Else
            'Suppression de la section
            strIn = Replace(strIn, GetBetween(strIn, "<section id=" & Chr(34) & "solution" & Chr(34) & ">", "</section>"), "")
        End If
        
        'Si le champ "Code" n'est pas vide
        If delAllSpace(Me.tbCode.Value) <> "" Then
            'Remplacement de la balise "{TEMPLATE_PREVIEW_CODE}" par le contenu du champ "Solution"
            strIn = Replace(strIn, "{TEMPLATE_PREVIEW_CODE}", Me.tbCode.Value)
            'Remplacement de la balise "{TEMPLATE_PREVIEW_IDLANG}" par l'ID du langage correspondant
            strIn = Replace(strIn, "{TEMPLATE_PREVIEW_IDLANG}", getIDLanguage(Me.cbLanguage.Value))
        Else
            'Suppression de la section
            strIn = Replace(strIn, GetBetween(strIn, "<section id=" & Chr(34) & "code" & Chr(34) & ">", "</section>"), "")
        End If
        
        'Ecriture dans le fichier HTML temporaire : htmlPreviewTmpName
        Call writeToFile(cataloguePath & sheetsPath & htmlPreviewTmpName, strIn)
        
        'Message d'erreur si une des images dont le code Markdown a �t� ins�r� dans la fiche contient un des caract�res suivants : [ ] ( )
        If bBadImgName Then
            MsgBox "Une ou plusieurs images des pi�ces jointes dont le code Markdown a �t� ins�r� dans la fiche contiennent un des caract�res suivants :" & vbNewLine & vbNewLine & _
                   "[  ]  (  )" & vbNewLine & vbNewLine & _
                   "La pr�visualisation n'affichera pas correctement ces images (limitation Markdown). Lorsque vous validerez la cr�ation / modification de la fiche, elles seront cependant renomm�es automatiquement.", vbInformation, "Saisie incompl�te"
        End If
        'Ouverture du fichier HTML temporaire dans le navigateur
        Call browseURL(cataloguePath & sheetsPath & htmlPreviewTmpName)
    Else
        MsgBox "Veuillez remplir au moins un des champs suivants :" & vbNewLine & vbNewLine & _
               "- Probl�me" & vbNewLine & _
               "- Solution" & vbNewLine & _
               "- Code" & vbNewLine, _
               vbInformation, "Incompatibilit� Markdown"
    End If
    Set focusedControl = Me.cbPreview
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'V�rifie l'orthographe des TextBox "Titre, "Solution", "Probl�me" et "Tags"
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbCheckSpell_Click()

    'Si cr�ation de fiche, l'utilisateur peut encore modifier le titre
    If Me.tbTitle.Enabled = True Then
        'V�rification du titre de la fiche
        Call checkSpellTextbox(Me.tbTitle, optionsSheet)
    End If
    'V�rification du champ "Probl�me"
    Call checkSpellTextbox(Me.tbProblem, optionsSheet)
    'V�rification du champ "Probl�me"
    Call checkSpellTextbox(Me.tbSolution, optionsSheet)
    'V�rification du champ "Tags"
    Call checkSpellTextbox(Me.tbKeywords, optionsSheet)
    Set focusedControl = Me.cbCheckSpell
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Bouton d'aide au langage Markdown : Ouvre la fiche n�187028
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbHelp_Click()

    Call browseURL(cataloguePath & sheetsPath & "187028_Syntaxe_et_exemples_du_langage_de_balisage_Markdown_1" & ".html")
    Set focusedControl = Me.cbHelp
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Ouvre le catalogue HTML dans le navigateur par d�faut
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbCatalogue_Click()

    Call browseURL(cataloguePath & htmlCataloguePath & htmlCatalogueName)
    Set focusedControl = Me.cbCatalogue
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur appuie sur "Valider" : Cr�ation / �dition de fiche
'R�cup�ration du contenu des champs du UserForm, remplissage des variables publiques
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbOK_Click()

    'Suppression du 1er caract�re du titre si apostrophe car pas affich� par Excel
    'Les espaces en trop sont �galement automatiquement supprim�s par delAllSpace()
    If VBA.Left(delAllSpace(Me.tbTitle.Value), 1) = "'" Then
        strTitle = VBA.Right(delAllSpace(Me.tbTitle.Value), Len(delAllSpace(Me.tbTitle.Value)) - 1)
    Else
        strTitle = delAllSpace(Me.tbTitle.Value)
    End If
    strCode = Me.tbCode.Value                                                                   'Code
    strKeywords = delAllSpace(Replace(Replace(Me.tbKeywords.Value, ";", ""), ",", ""))          'Mots cl�s
    strSoftware = Me.cbSoftware.Value                                                           'Logiciel
    strLanguage = Me.cbLanguage.Value                                                           'Langage
    strType = Me.cbType.Value                                                                   'Type
    strStatus = Me.cbStatus.Value                                                               'Statut
    
    'L'ancienne release passe en superseded, si au moins une release existe d�j�
    If strStatus = "Released" And intVersion > 1 Then
        isSuperseded = True
    Else
        isSuperseded = False
    End If
    
    intVersion = Me.tbVersion.Value                        'Version
    dblId = Me.tbID.Value                                  'ID unique
    strProblem = Me.tbProblem.Value                        'Probl�me
    strSolution = Me.tbSolution.Value                      'Solution
    
    'Si tous les champs du UserForm n'ont pas �t� remplis, affiche une erreur
    Dim sRequiredInputList As String
    sRequiredInputList = ""
    Dim iRequiredCount As Integer
    iRequiredCount = 0
    'Titre
    If Replace(strTitle, " ", "") = "" Then
        sRequiredInputList = sRequiredInputList & "- Titre" & vbNewLine
        iRequiredCount = iRequiredCount + 1
    End If
    'Probl�me
    If Replace(strProblem, " ", "") = "" Then
        sRequiredInputList = sRequiredInputList & "- Probl�me" & vbNewLine
        iRequiredCount = iRequiredCount + 1
    End If
    'Mots cl�s
    If Replace(strKeywords, " ", "") = "" Then
        sRequiredInputList = sRequiredInputList & "- Tags" & vbNewLine
        iRequiredCount = iRequiredCount + 1
    End If
    'Statut
    If Replace(strStatus, " ", "") = "" Then
        sRequiredInputList = sRequiredInputList & "- Statut" & vbNewLine
        iRequiredCount = iRequiredCount + 1
    End If
    'Langage
    If Replace(strLanguage, " ", "") = "" Then
        sRequiredInputList = sRequiredInputList & "- Langage" & vbNewLine
        iRequiredCount = iRequiredCount + 1
    End If
    'Logiciel
    If Replace(strSoftware, " ", "") = "" Then
        sRequiredInputList = sRequiredInputList & "- Logiciel" & vbNewLine
        iRequiredCount = iRequiredCount + 1
    End If
    'Type
    If Replace(strType, " ", "") = "" Then
        sRequiredInputList = sRequiredInputList & "- Type" & vbNewLine
        iRequiredCount = iRequiredCount + 1
    End If
    If iRequiredCount = 1 Then
        MsgBox "Veuillez remplir le champ suivant :" & vbNewLine & vbNewLine & sRequiredInputList, vbInformation, "Saisie incompl�te"
    ElseIf iRequiredCount > 1 Then
        MsgBox "Veuillez remplir les champs suivants :" & vbNewLine & vbNewLine & sRequiredInputList, vbInformation, "Saisie incompl�te"
    Else
        'Si la fiche existe d�j� dans le catalogue, demande confirmation avant �crasement
        If fileExists(cataloguePath & sheetsPath & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html") Then
            If MsgBox("Le fichier suivant existe d�j�, souhaitez vous l'�craser ?" & vbNewLine & vbNewLine & replaceIllegalChar(dblId & "_" & strTitle & "_" & intVersion) & ".html", vbYesNo + vbQuestion, "Confirmation d'�crasement") = vbYes Then
                Me.Hide
            End If
        Else
            'Fermeture de la fen�tre de cr�ation / �dition de fiche
            Me.Hide
        End If
    End If
    Set focusedControl = Me.cbOK
    
End Sub

'===============================================================================================================================================
'===============================================================================================================================================
'
'Fonctions li�es � l'initialisation du UserForm, de ses �l�ments, et des diff�rentes interactions utilisateur : fermeture, ...
'
'===============================================================================================================================================
'===============================================================================================================================================

'-----------------------------------------------------------------------------------------------------------------------------------------------
'Modifie le contenu de la ComboBox de statuts suivant le statut de la fiche
'-----------------------------------------------------------------------------------------------------------------------------------------------
Sub populateStatusComboBox(ByRef cbStatus As ComboBox, ByVal strStatus As String)

    'R�initialisation de la ComboBox
    cbStatus.Clear
    'Bouton "Valider" activ� par d�faut : �dition de la fiche possible
    Me.cbOK.Enabled = True
    
    Select Case strStatus
        Case "Draft"
            cbStatus.AddItem optionsSheet.Cells(2, colOptStatus).Value    'Draft
            cbStatus.AddItem optionsSheet.Cells(5, colOptStatus).Value    'Submitted
            cbStatus.AddItem optionsSheet.Cells(3, colOptStatus).Value    'Released
            cbStatus.AddItem optionsSheet.Cells(6, colOptStatus).Value    'Obsolete
        Case "Submitted"
            cbStatus.AddItem optionsSheet.Cells(2, colOptStatus).Value    'Draft
            cbStatus.AddItem optionsSheet.Cells(5, colOptStatus).Value    'Submitted
            cbStatus.AddItem optionsSheet.Cells(3, colOptStatus).Value    'Released
            cbStatus.AddItem optionsSheet.Cells(6, colOptStatus).Value    'Obsolete
        Case "Released"
            'Si �dition de fiche Released, un nouveau Draft est cr��
            cbStatus.AddItem optionsSheet.Cells(2, colOptStatus).Value    'Draft
            cbStatus.AddItem optionsSheet.Cells(5, colOptStatus).Value    'Submitted
        Case "Superseded"
            cbStatus.AddItem optionsSheet.Cells(4, colOptStatus).Value    'Superseded
            Me.cbOK.Enabled = False                                       'Bouton "Valider" d�sactiv� pour emp�cher l'�dition de la fiche
        Case "Obsolete"
            cbStatus.AddItem optionsSheet.Cells(6, colOptStatus).Value    'Obsolete
            Me.cbOK.Enabled = False                                       'Bouton "Valider" d�sactiv� pour emp�cher l'�dition de la fiche
        Case Else
            cbStatus.AddItem optionsSheet.Cells(2, colOptStatus).Value    'Draft
            cbStatus.AddItem optionsSheet.Cells(5, colOptStatus).Value    'Submitted
    End Select

End Sub

Private Sub UserForm_Initialize()

    'Initialisation des ComboBox logiciels, langages, et types
    Call initializeComboBox
        
    'Call UserForm_activate

End Sub

Private Sub UserForm_activate()

    'Modifie le contenu de la ComboBox de statuts suivant le statut de la fiche
    Call populateStatusComboBox(Me.cbStatus, strStatus)
    
    'R�initialisation de la position et de la taille des champs "Probl�me", "Solution" et "Code"
    Me.tbProblem.Height = 100
    Me.tbProblem.Top = 66
    Me.tbSolution.Height = 100
    Me.tbSolution.Top = 170
    Me.tbCode.Height = 105
    Me.tbCode.Top = 275
    
    lbProblem.Top = 108
    lbSolution.Top = 210
    lbCode.Top = 318
    
    'Si la UserForm a �t� quitt�e alors que le champ "Probl�me", "Solution" ou "Code" �tait maximis� :
    'R�affiche le panel de pi�ces jointes, re-minimise les champs, r�affiche les champs / labels masqu�s
    If frameFiles.Visible = False Then
        frameFiles.Visible = True
        
        Me.tbProblem.Width = 524
        Me.tbSolution.Width = 524
        Me.tbCode.Width = 524
        
        Me.tbProblem.Visible = True
        Me.tbSolution.Visible = True
        Me.tbCode.Visible = True
        
        Me.lbProblem.Visible = True
        Me.lbSolution.Visible = True
        Me.lbCode.Visible = True
        
        lbKeywords.Visible = True
        lbLanguage.Visible = True
        cbLanguage.Visible = True
        cbSoftware.Visible = True
        cbType.Visible = True
        tbID.Enabled = True
        tbVersion.Enabled = True
    End If
    
    'Si cr�ation d'une nouvelle fiche, r�initialisation des champs
    If strLanguage = "" Then
        Me.cbLanguage.Value = ""                                       'Langage
        Me.cbSoftware.Value = ""                                       'Logiciel
        Me.cbType.Value = ""                                           'Type
        Me.cbStatus.Value = Me.cbStatus.List(0)                        'Statut
        
        Me.tbVersion.Value = 1                                         'Version
        
        dblId = 100
        Do While Len(CStr(dblId)) <> 6 Or isUniqueID(dblId) = False
            dblId = GetUniqueID
        Loop
        Me.tbID.Value = dblId                                          'ID unique
        
        Me.tbProblem.Value = ""                                        'Probl�me
        Me.tbSolution.Value = ""                                       'Solution
        Me.tbCode.Value = ""                                           'Code
'        Me.tbCode.Value = getClipboardContent()                        'R�cup�re le contenu du presse papier
        Me.fileList.Clear                                              'Pi�ces jointes
        lbFilesCpt.Caption = strFile.Count
        
        Me.tbTitle.Enabled = True                                      'Cr�ation de fiche : - Le titre de la fiche peut �tre �dit�
        Me.Caption = "Cr�ation d'une fiche"
        
        Me.Width = 600                                                 'Cr�ation de fiche : - Le panel de pi�ces jointes est ferm�
        Me.tbTitle.SetFocus                                            'Cr�ation de fiche : - Focus sur la TextBox de titre
    Else
        'Si �dition de fiche, remplissage des champs � partir des variables publiques
        Me.cbLanguage.Value = strLanguage                              'Langage
        Me.cbSoftware.Value = strSoftware                              'Logiciel
        Me.cbType.Value = strType                                      'Type
        If strStatus = "Released" Then                                 'Si �dition d'une fiche released :
            Me.tbVersion.Value = intVersion + 1                        '   - Version : Incr�mentation de la version
            Me.cbStatus.Value = Me.cbStatus.List(0)                    '   - Statut : Statut draft par d�faut
        Else
            Me.tbVersion.Value = intVersion                            'Version
            Me.cbStatus.Value = strStatus                              'Statut
        End If
        Me.tbID.Value = dblId                                          'ID unique
        Me.tbProblem.Value = strProblem                                'Probl�me
        Me.tbSolution.Value = strSolution                              'Solution
        Me.tbCode.Value = specialCharsHtml(strCode)                    'Code
        Me.fileList.Clear                                              'Pi�ces jointes
        lbFilesCpt.Caption = strFile.Count
        Dim i As Integer
        For i = 1 To strFile.Count
            Me.fileList.AddItem getFilenameFromPath(strFile.item(i))   'Remplit la ListBox � partir de la collection de pi�ces jointes � copier / conserver
        Next
        
        On Error Resume Next                                           'Affiche les donn�es de chaque TextBox � la partir de la 1i�re ligne
        If delAllSpace(Me.tbSolution.Value) <> "" Then
            With Me.tbSolution
                .SetFocus
                .SelStart = 0
            End With
        End If
        If delAllSpace(Me.tbProblem.Value) <> "" Then
            With Me.tbProblem
                .SetFocus
                .SelStart = 0
            End With
        End If
        If delAllSpace(Me.tbCode.Value) <> "" Then
            With Me.tbCode
                .SetFocus
                .SelStart = 0
            End With
        End If
        On Error GoTo 0
        
        If strFile.Count <> 0 Then                                     'Agrandit la largeur du UserForm pour afficher le panel de pi�ces jointes s'il y en a
            Me.Width = 900
        Else
            Me.Width = 600
        End If
        
        Me.tbTitle.Enabled = False                                     'Edition de fiche : Le titre de la fiche ne peut pas �tre �dit�
        Me.Caption = "Edition d'une fiche"
    End If
    
    'Positionnement du UserForm au milieu de la fen�tre Excel
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    End With

    Me.tbTitle.Value = strTitle                                        'Titre
    Me.tbKeywords.Value = strKeywords                                  'Mots cl�s
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur appuie sur "Annuler" : arr�t de la macro
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub cbCancel_Click()

    If MsgBox("Les modifications apport�es � la fiche vont �tre perdues, souhaitez vous continuer ? ""Oui"" pour continuer, ""Non"" pour poursuivre l'�dition.", vbYesNo + vbQuestion, "Continuer ?") = vbYes Then
        Application.Visible = True
        Me.Hide
        
        'Vidage des variables publiques
        Call resetPublicVar
        
        bExit = True
    End If
    Set focusedControl = Me.cbCancel
    
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
'L'utilisateur ferme le UserForm (croix rouge) : arr�t de la macro
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If MsgBox("Les modifications apport�es � la fiche vont �tre perdues, souhaitez vous continuer ? ""Oui"" pour continuer, ""Non"" pour poursuivre l'�dition.", vbYesNo + vbQuestion, "Continuer ?") = vbYes Then
        Application.Visible = True
        Me.Hide
        
        'Vidage des variables publiques
        Call resetPublicVar
        
        bExit = True
    Else
        'Ne pas fermer le UserForm si l'utilisateur a choisi "Non"
        Cancel = (CloseMode = 0)
    End If

End Sub
