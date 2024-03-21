Sub TrouverH2O()
    Dim phrase1 As String
    Dim phrase2 As String
    Dim bigrammes1() As String
    Dim bigrammes2() As String
    Dim i As Long
    Dim countH2O1 As Long
    Dim countH2O2 As Long
    Dim pourcentage1 As Double
    Dim pourcentage2 As Double
    
    ' Phrases à rechercher
    phrase1 = "mandate h2o AM"
    phrase2 = "H2o Multi agrégate d’une"
    
    ' Créer des bigrammes pour la première phrase
    bigrammes1 = ExtraireBigrammes(phrase1)
    
    ' Créer des bigrammes pour la deuxième phrase
    bigrammes2 = ExtraireBigrammes(phrase2)
    
    ' Compter le nombre de bigrammes contenant "h2o" dans la première phrase
    For i = LBound(bigrammes1) To UBound(bigrammes1)
        If InStr(1, bigrammes1(i), "h2o", vbTextCompare) > 0 Then
            countH2O1 = countH2O1 + 1
        End If
    Next i
    
    ' Compter le nombre de bigrammes contenant "h2o" dans la deuxième phrase
    For i = LBound(bigrammes2) To UBound(bigrammes2)
        If InStr(1, bigrammes2(i), "h2o", vbTextCompare) > 0 Then
            countH2O2 = countH2O2 + 1
        End If
    Next i
    
    ' Calculer le pourcentage de correspondance pour chaque phrase
    pourcentage1 = (countH2O1 / UBound(bigrammes1)) * 100
    pourcentage2 = (countH2O2 / UBound(bigrammes2)) * 100
    
    ' Afficher les résultats
    MsgBox "Pourcentage de correspondance pour la première phrase : " & pourcentage1 & "%" & vbCrLf & _
           "Pourcentage de correspondance pour la deuxième phrase : " & pourcentage2 & "%"
End Sub

Function ExtraireBigrammes(phrase As String) As String()
    Dim mots() As String
    Dim bigrammes() As String
    Dim i As Long
    
    ' Diviser la phrase en mots
    mots = Split(phrase, " ")
    
    ' Créer des bigrammes à partir des mots
    ReDim bigrammes(1 To UBound(mots))
    For i = LBound(mots) To UBound(mots) - 1
        bigrammes(i) = mots(i) & " " & mots(i + 1)
    Next i
    
    ExtraireBigrammes = bigrammes
End Function



#


Function CreateNGram(strInput, intN)
    Dim arrNGram, intBound, i, j, strGram, didInc, arrTemp
 
    If Len(strInput) = 0 Then Exit Function
 
    ReDim arrNGram(Len(strInput) + 1, 1)
    strInput = Chr(0) & UCase(Trim(strInput)) & Chr(0)
    intBound = -1
 
    For i = 1 To Len(strInput)-intN+1
        strGram = Mid(strInput, i, intN)
        didInc = False
 
        For j = 0 To intBound
            If strGram = arrNGram(j, 0) Then
                arrNGram(j, 1) = arrNGram(j, 1) + 1
                didInc = True
                Exit For
            End If
        Next
 
        If Not didInc Then
            intBound = intBound + 1
            arrNGram(intBound, 0) = strGram
            arrNGram(intBound, 1) = 1
        End If
    Next
 
    ReDim arrTemp(intBound, 1)
    For i = 0 To intBound
        arrTemp(i, 0) = arrNGram(i, 0)
        arrTemp(i, 1) = arrNGram(i, 1)
    Next
 
    CreateNGram = arrTemp
End Function
 
Function CompareNGram(arr1, arr2)
    Dim i, j, intMatches, intCount1, intCount2
 
    intMatches = 0
    intCount1 = 0
 
    For i = 0 To UBound(arr1)
        intCount1 = intCount1 + arr1(i, 1)
        intCount2 = 0
 
        For j = 0 To UBound(arr2)
            intCount2 = intCount2 + arr2(j, 1)
 
            If arr1(i, 0) = arr2(j, 0) Then
                If arr1(i, 1) >= arr2(j, 1) Then
                    intMatches = intMatches + arr2(j, 1)
                Else
                    intMatches = intMatches + arr1(i, 1)
                End If
            End If
        Next
    Next
 
    CompareNGram = 2 * intMatches / (intCount1 + intCount2)
End Function


#

```
Private Function getNgramProfile(s As String, Optional k As Long = 3) As Scripting.Dictionary
    Dim i As Long
    Dim old As Long
    Dim ngram As String
    Dim ngrams As Scripting.Dictionary
    Dim string_no_space As String
    string_no_space = normalize(s, " ,./;'[]\!@#$%^&*()_") 'Replace(s, " ", "")
    Set ngrams = New Scripting.Dictionary
    
    
    For i = 1 To (Len(string_no_space) - k + 1)
        ngram = Mid(string_no_space, i, k)
        If ngrams.Exists(ngram) Then
            old = ngrams.Item(ngram)
            ngrams(ngram) = old + 1
        Else
            ngrams(ngram) = 1
        End If
        
    Next i
    
    Set getNgramProfile = ngrams
    
End Function

Private Function normalize(s As String, special_characters As String) As String
    Dim i As Long
    
    For i = 1 To Len(special_characters)
        s = replace(s, Mid(special_characters, i, 1), "")
    Next i
    
    normalize = s
End Function

```

# Similitude avec Tversky

L'indice de Tversky est une mesure d'algorithme de similarité qui quantifie le degré de chevauchement entre deux ensembles, en tenant compte à la fois des faux positifs et des faux négatifs. Ceci est particulièrement utile lorsqu'il s'agit de données déséquilibrées ou de situations dans lesquelles la présence ou l'absence d'éléments dans les ensembles a une importance variable. Nous avons donc ici le choix de mettre l'accent sur les mots (caractères) communs ou sur ceux non partagés.
```
Option Explicit

Function TverskySimilarity(set1 As Object, set2 As Object, alpha As Double, beta As Double) As Double
    Dim commonCount As Integer
    Dim nonCommonCount1 As Integer
    Dim nonCommonCount2 As Integer
    Dim similarity As Double
    
    ' Comptage des éléments communs
    commonCount = 0
    For Each elem In set1
        If Not IsError(Application.Match(elem, set2, 0)) Then
            commonCount = commonCount + 1
        End If
    Next elem
    
    ' Comptage des éléments non communs dans set1
    nonCommonCount1 = set1.Count - commonCount
    
    ' Comptage des éléments non communs dans set2
    nonCommonCount2 = set2.Count - commonCount
    
    ' Calcul de la similitude de Tversky
    similarity = commonCount / (commonCount + alpha * nonCommonCount1 + beta * nonCommonCount2)
    
    TverskySimilarity = similarity
End Function

```
# Similitude de chevauchement

Le coefficient de chevauchement, similarité , ou coefficient de Szymkiewicz-Simpson, est une simple similarité entre deux ensembles. Il calcule le rapport entre la taille de l’intersection des ensembles et la taille du plus petit ensemble. Le coefficient de chevauchement est particulièrement utile lorsqu'il s'agit de données binaires ou de situations où la présence ou l'absence d'éléments est importante. Sa formule suivante est :

```
Option Explicit

Function OverlapSimilarity(s1 As String, s2 As String) As Double
    Dim overlap As Integer
    Dim i As Integer, j As Integer
    Dim max_overlap As Integer
    
    ' Recherche du chevauchement maximal
    max_overlap = 0
    For i = 1 To Len(s1)
        For j = 1 To Len(s2)
            If Mid$(s1, i) = Left$(s2, Len(s1) - i + 1) Then
                overlap = Len(s1) - i + 1
                If overlap > max_overlap Then
                    max_overlap = overlap
                End If
            End If
        Next j
    Next i
    
    ' Calcul de la similitude de chevauchement
    If Len(s1) + Len(s2) = 0 Then
        OverlapSimilarity = 0
    Else
        OverlapSimilarity = 2 * max_overlap / (Len(s1) + Len(s2))
    End If
End Function

```

# séquence

## Similitude Ratcliff-Obershelp
La similarité Ratcliff-Obershelp ou correspondance de modèles Gestalt est une mesure de similarité de chaînes qui se concentre sur la recherche de la similarité entre deux chaînes en fonction de leurs sous-chaînes communes. Il est particulièrement utile pour comparer des chaînes ayant des structures similaires mais pouvant différer en termes de modifications, suppressions ou insertions mineures. L'algorithme attribue un score de similarité basé sur la longueur de la sous-chaîne commune la plus longue entre les deux chaînes.
```
Function RatcliffObershelp(str1 As String, str2 As String) As Double
    Dim length1 As Integer, length2 As Integer
    Dim common As Integer, maxCommon As Integer
    Dim i As Integer, j As Integer
    Dim maxLen As Integer
    Dim temp As String
    
    length1 = Len(str1)
    length2 = Len(str2)
    
    maxLen = WorksheetFunction.Max(length1, length2)
    
    If maxLen = 0 Then
        RatcliffObershelp = 1
        Exit Function
    End If
    
    maxCommon = 0
    For i = 1 To length1
        For j = 1 To length2
            If Mid(str1, i, 1) = Mid(str2, j, 1) Then
                common = 1
                Do While (i + common <= length1) And (j + common <= length2) And (Mid(str1, i + common, 1) = Mid(str2, j + common, 1))
                    common = common + 1
                Loop
                If common > maxCommon Then
                    maxCommon = common
                    temp = Mid(str1, i, common)
                End If
            End If
        Next j
    Next i
    
    RatcliffObershelp = (2 * maxCommon) / maxLen
End Function

```
+ rapide
```
Option Explicit

Function RatcliffObershelp(s1 As String, s2 As String) As Double
    Dim matrix() As Integer
    Dim len_s1 As Integer, len_s2 As Integer
    Dim i As Integer, j As Integer
    Dim max_match As Integer, match As Integer
    Dim char1 As String, char2 As String

    len_s1 = Len(s1)
    len_s2 = Len(s2)

    ReDim matrix(0 To len_s1)

    max_match = 0

    For i = 1 To len_s1
        For j = len_s2 To 1 Step -1
            char1 = Mid$(s1, i, 1)
            char2 = Mid$(s2, j, 1)
            If char1 = char2 Then
                matrix(i) = matrix(i - 1) + 1
                If matrix(i) > max_match Then
                    max_match = matrix(i)
                End If
            Else
                matrix(i) = 0
            End If
        Next j
    Next i

    If len_s1 = 0 Or len_s2 = 0 Then
        RatcliffObershelp = 0
    Else
        match = 2 * max_match
        RatcliffObershelp = match / (len_s1 + len_s2)
    End If
End Function

```


# Similitude de sous-chaîne/sous-séquence commune la plus longue

```
Option Explicit

Function LongestCommonSubstring(s1 As String, s2 As String) As Double
    Dim len_s1 As Integer, len_s2 As Integer
    Dim i As Integer, j As Integer
    Dim lcs() As Integer
    Dim max_length As Integer, current_length As Integer
    Dim char1 As String, char2 As String
    
    len_s1 = Len(s1)
    len_s2 = Len(s2)
    
    ReDim lcs(1 To len_s2 + 1)
    
    max_length = 0
    
    For i = 1 To len_s1
        For j = len_s2 To 1 Step -1
            char1 = Mid$(s1, i, 1)
            char2 = Mid$(s2, j, 1)
            If char1 = char2 Then
                lcs(j + 1) = lcs(j) + 1
                If lcs(j + 1) > max_length Then
                    max_length = lcs(j + 1)
                End If
            Else
                lcs(j + 1) = 0
            End If
        Next j
    Next i
    
    If len_s1 = 0 Or len_s2 = 0 Then
        LongestCommonSubstring = 0
    Else
        LongestCommonSubstring = max_length / len_s1 ' Normalizing by length of first string
    End If
End Function

```

# test ++ levenshtein


```
Function Levenshtein(ByVal string1 As String, ByVal string2 As String) As Long

Dim i As Long, j As Long
Dim string1_length As Long
Dim string2_length As Long
Dim distance() As Long

string1_length = Len(string1)
string2_length = Len(string2)
ReDim distance(string1_length, string2_length)

For i = 0 To string1_length
    distance(i, 0) = i
Next

For j = 0 To string2_length
    distance(0, j) = j
Next

For i = 1 To string1_length
    For j = 1 To string2_length
        If Asc(Mid$(string1, i, 1)) = Asc(Mid$(string2, j, 1)) Then
            distance(i, j) = distance(i - 1, j - 1)
        Else
            distance(i, j) = Application.WorksheetFunction.Min _
            (distance(i - 1, j) + 1, _
             distance(i, j - 1) + 1, _
             distance(i - 1, j - 1) + 1)
        End If
    Next
Next

Levenshtein = distance(string1_length, string2_length)

End Function

```


```
Option Explicit

  Public Declare Function GetTickCount Lib "kernel32" () As Long

  Sub test()
  Dim s1 As String, s2 As String, lTime As Long, i As Long
  s1 = Space(100)
  s2 = String(100, "a")
  lTime = GetTickCount
  For i = 1 To 100
     LevenshteinStrings s1, s2  ' the original fn from Wikibooks and Stackoverflow
  Next
  Debug.Print GetTickCount - lTime; " ms" '  3900  ms for all diff

  lTime = GetTickCount
  For i = 1 To 100
     Levenshtein s1, s2
  Next
  Debug.Print GetTickCount - lTime; " ms" ' 234  ms

  End Sub

  'Option Base 0 assumed

  'POB: fn with byte array is 17 times faster
  Function Levenshtein(ByVal string1 As String, ByVal string2 As String) As Long

  Dim i As Long, j As Long, bs1() As Byte, bs2() As Byte
  Dim string1_length As Long
  Dim string2_length As Long
  Dim distance() As Long
  Dim min1 As Long, min2 As Long, min3 As Long

  string1_length = Len(string1)
  string2_length = Len(string2)
  ReDim distance(string1_length, string2_length)
  bs1 = string1
  bs2 = string2

  For i = 0 To string1_length
      distance(i, 0) = i
  Next

  For j = 0 To string2_length
      distance(0, j) = j
  Next

  For i = 1 To string1_length
      For j = 1 To string2_length
          'slow way: If Mid$(string1, i, 1) = Mid$(string2, j, 1) Then
          If bs1((i - 1) * 2) = bs2((j - 1) * 2) Then   ' *2 because Unicode every 2nd byte is 0
              distance(i, j) = distance(i - 1, j - 1)
          Else
              'distance(i, j) = Application.WorksheetFunction.Min _
              (distance(i - 1, j) + 1, _
               distance(i, j - 1) + 1, _
               distance(i - 1, j - 1) + 1)
              ' spell it out, 50 times faster than worksheetfunction.min
              min1 = distance(i - 1, j) + 1
              min2 = distance(i, j - 1) + 1
              min3 = distance(i - 1, j - 1) + 1
              If min1 <= min2 And min1 <= min3 Then
                  distance(i, j) = min1
              ElseIf min2 <= min1 And min2 <= min3 Then
                  distance(i, j) = min2
              Else
                  distance(i, j) = min3
              End If

          End If
      Next
  Next

  Levenshtein = distance(string1_length, string2_length)

  End Function

```