Function MoyennePonderee(ByVal pond() As Double, ByVal seuil() As Double, ByVal ParamArray similarite() As Variant) As Double
    Dim i As Integer
    Dim totalPond As Double
    Dim totalSimilarite As Double
    
    ' Vérifier la cohérence des paramètres passés
    If UBound(pond) <> UBound(seuil) Or UBound(pond) <> UBound(similarite) Then
        MoyennePonderee = CVErr(xlErrRef)
        Exit Function
    End If
    
    ' Initialiser les totaux
    totalPond = 0
    totalSimilarite = 0
    
    ' Calculer la somme pondérée des similarités
    For i = LBound(pond) To UBound(pond)
        If similarite(i) >= seuil(i) Then
            totalSimilarite = totalSimilarite + pond(i) * similarite(i)
            totalPond = totalPond + pond(i)
        End If
    Next i
    
    ' Calculer la moyenne pondérée si la somme des pondérations est non nulle
    If totalPond <> 0 Then
        MoyennePonderee = totalSimilarite / totalPond
    Else
        MoyennePonderee = 0 ' Ou une autre valeur par défaut
    End If
End Function
	
	
# bien

Function MoyennePonderee(ByVal pond() As Double, ByVal seuil() As Double, ByVal ParamArray similarite() As Variant) As Double
    Dim i As Integer
    Dim totalPond As Double
    Dim totalSimilarite As Double
    
    ' Vérifier la cohérence des paramètres passés
    If UBound(pond) <> UBound(seuil) Or UBound(pond) <> UBound(similarite) Then
        MoyennePonderee = CVErr(xlErrRef)
        Exit Function
    End If
    
    ' Initialiser les totaux
    totalPond = 0
    totalSimilarite = 0
    
    ' Calculer la somme pondérée des similarités
    For i = LBound(pond) To UBound(pond)
        If similarite(i) >= seuil(i) Then
            ' Appeler la fonction de calcul de similarité appropriée en fonction du type de modèle
            Select Case LCase(similarityType(i))
                Case "levenshtein":
                    totalSimilarite = totalSimilarite + pond(i) * Levenshtein(similarite(i))
                Case "damereau":
                    totalSimilarite = totalSimilarite + pond(i) * Damereau(similarite(i))
                Case "jarowinkler":
                    totalSimilarite = totalSimilarite + pond(i) * JaroWinkler(similarite(i))
                ' Ajouter d'autres cas pour d'autres types de modèles si nécessaire
                Case Else
                    ' Gérer les cas non pris en charge
                    totalSimilarite = totalSimilarite + 0 ' Ou une autre valeur par défaut
            End Select
            totalPond = totalPond + pond(i)
        End If
    Next i
    
    ' Calculer la moyenne pondérée si la somme des pondérations est non nulle
    If totalPond <> 0 Then
        MoyennePonderee = totalSimilarite / totalPond
    Else
        MoyennePonderee = 0 ' Ou une autre valeur par défaut
    End If
End Function

' Fonction pour calculer la similarité Levenshtein
Function Levenshtein(ByVal value As Variant) As Double
    ' Mettre ici la logique de calcul de la distance de Levenshtein
End Function

' Fonction pour calculer la similarité de Damereau
Function Damereau(ByVal value As Variant) As Double
    ' Mettre ici la logique de calcul de la distance de Damereau
End Function

' Fonction pour calculer la similarité de Jaro-Winkler
Function JaroWinkler(ByVal value As Variant) As Double
    ' Mettre ici la logique de calcul de la distance de Jaro-Winkler
End Function

# Bien mieux

Function MoyennePonderee(ByVal pond() As Double, ByVal ParamArray models() As Variant) As Double
    Dim i As Integer
    Dim totalPond As Double
    Dim totalSimilarite As Double
    Dim similarites() As Double
    Dim seuils() As Double
    
    ' Vérifier la cohérence des paramètres passés
    If UBound(pond) <> UBound(models) Then
        MoyennePonderee = CVErr(xlErrRef)
        Exit Function
    End If
    
    ' Initialiser les totaux
    totalPond = 0
    totalSimilarite = 0
    
    ' Redimensionner les tableaux pour stocker les similarités et les seuils
    ReDim similarites(LBound(models) To UBound(models))
    ReDim seuils(LBound(models) To UBound(models))
    
    ' Calculer les similarités et stocker les seuils correspondants
    For i = LBound(models) To UBound(models)
        Select Case LCase(models(i))
            Case "levenshtein":
                similarites(i) = Levenshtein() ' Appeler la fonction de calcul de similarité Levenshtein
            Case "damereau":
                similarites(i) = Damereau() ' Appeler la fonction de calcul de similarité Damereau
            Case "jarowinkler":
                similarites(i) = JaroWinkler() ' Appeler la fonction de calcul de similarité JaroWinkler
            ' Ajouter d'autres cas pour d'autres types de modèles si nécessaire
        End Select
        
        ' Stocker le seuil correspondant pour chaque modèle
        seuils(i) = GetSeuil(models(i)) ' Appeler la fonction pour obtenir le seuil correspondant
        
        ' Vérifier si la similarité est supérieure ou égale au seuil correspondant
        If similarites(i) >= seuils(i) Then
            totalSimilarite = totalSimilarite + pond(i) * similarites(i)
            totalPond = totalPond + pond(i)
        End If
    Next i
    
    ' Calculer la moyenne pondérée si la somme des pondérations est non nulle
    If totalPond <> 0 Then
        MoyennePonderee = totalSimilarite / totalPond
    Else
        MoyennePonderee = 0 ' Ou une autre valeur par défaut
    End If
End Function

***
            ' Vérifier si cette similarité est meilleure que celle stockée pour ce couple
            If Not similarityDict.Exists(cell1.Address & "|" & cell2.Address) Then
                similarityDict(cell1.Address & "|" & cell2.Address) = similarity
            ElseIf similarity > similarityDict(cell1.Address & "|" & cell2.Address) Then
                similarityDict(cell1.Address & "|" & cell2.Address) = similarity
            End If