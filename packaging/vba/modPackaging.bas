Attribute VB_Name = "modPackaging"
Option Explicit
' =====================================================================
'  Packaging Engine MMS - suggestion de caisses (All Seasons / Summer)
'  ---------------------------------------------------------------
'  Feuilles utilisees :
'    - Produits             : SKU, Produit, Par Saison?, sub_Item_1..3 (+quantite, Longueur, Largeur, Hauteur)
'    - Caisses              : Nom de Caisse, Saison, Longueur, Largeur, Hauteur
'    - Packaging Calculator : Marge vide All Seasons / Marge vide Summer + table Item / Quantity / ...
'    - Mixed Pack Helper    : Order ID / Item / Quantity  ->  Order ID / Suggested ... (par commande)
'  Macros a lancer :
'    - Packaging_Setup              : cree les feuilles / en-tetes / boutons "Remplir" (a lancer 1 fois)
'    - RemplirPackagingCalculator   : bouton de la feuille Packaging Calculator
'    - RemplirMixedPackHelper       : bouton de la feuille Mixed Pack Helper
'  Algorithme :
'    - Chaque produit est decompose en sous-articles (boites) avec dimensions.
'    - Hauteur utilisable d'une caisse = Hauteur x (1 - marge vide)  (la marge est en haut).
'    - Placement 3D "espaces maximaux" (pas de debordement, 6 orientations possibles),
'      plusieurs strategies essayees, la meilleure est retenue.
'    - Priorite au minimum de caisses : 1 seule caisse si possible (la plus petite qui convient),
'      sinon on remplit la caisse qui prend le plus de volume et on recommence avec le reste.
'    - Format du resultat : "1:Caisse Super Small_2:Caisse V2"  (qte:nom, separateur "_").
' =====================================================================

Public Const SHEET_PRODUITS As String = "Produits"
Public Const SHEET_CAISSES As String = "Caisses"
Public Const SHEET_CALC As String = "Packaging Calculator"
Public Const SHEET_MIX As String = "Mixed Pack Helper"

' True  : un article peut etre couche sur n'importe quelle face (6 orientations)
' False : un article reste debout, seule la rotation a plat est permise (2 orientations)
Public Const ALLOW_TIP_OVER As Boolean = True

Private Const EPS As Double = 0.000000001
Private Const BIG_N As Long = 150          ' au-dela, on saute les strategies "best fit" (plus lentes)
Private Const MAX_ITEMS As Long = 5000
Private Const MAX_CAISSES As Long = 200

Private Type TItem
    L As Double
    W As Double
    H As Double
    Vol As Double
    MaxDim As Double
    Ref As String
    Done As Boolean      ' deja place dans une caisse validee
    Placed As Boolean    ' place dans la tentative en cours
End Type

Private Type TSpace
    X As Double
    Y As Double
    Z As Double
    L As Double
    W As Double
    H As Double
End Type

Private Type TCaisse
    Idx As Long
    CaisseName As String
    IsSummer As Boolean
    L As Double
    W As Double
    H As Double
End Type

Private Type TProduct
    Key As String        ' SKU normalise
    NameKey As String    ' nom produit normalise
    ProdName As String
    Seasonal As Boolean  ' Par Saison? = Oui
    NSub As Long
    SubName(1 To 3) As String
    SubQty(1 To 3) As Double
    SubL(1 To 3) As Double
    SubW(1 To 3) As Double
    SubH(1 To 3) As Double
    SubOK(1 To 3) As Boolean
End Type

' ---- donnees chargees ----
Private mProd() As TProduct
Private mNProd As Long
Private mCaisses() As TCaisse
Private mNCaisses As Long

' ---- ordres de tri des articles (calcules une fois par commande) ----
Private mOrdVol() As Long
Private mOrdDim() As Long

' =====================================================================
'                         POINTS D'ENTREE
' =====================================================================

Public Sub RemplirPackagingCalculator()
    Dim ws As Worksheet, r As Long, lastRow As Long, hdr As Long
    Dim mA As Double, mS As Double
    Dim skus() As String, qtys() As Double
    Dim resA As String, resS As String, info As String
    Dim items() As TItem, n As Long, seasonal As Boolean, errMsg As String
    Dim t0 As Double

    On Error GoTo Fail
    t0 = Timer
    Set ws = ThisWorkbook.Worksheets(SHEET_CALC)
    If Not ReadMargins(ws, mA, mS) Then Exit Sub
    hdr = FindRowStartingWith(ws, "Item", 1, 30)
    If hdr = 0 Then
        MsgBox "En-tete 'Item' introuvable en colonne A de la feuille '" & SHEET_CALC & "'. Lancez Packaging_Setup.", vbExclamation
        Exit Sub
    End If
    LoadData
    Application.ScreenUpdating = False
    ReDim skus(1 To 1): ReDim qtys(1 To 1)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = hdr + 1 To lastRow
        If Len(Trim(CStr(ws.Cells(r, 1).Value))) = 0 Then
            ws.Cells(r, 3).Resize(1, 3).ClearContents
        Else
            Application.StatusBar = "Packaging Calculator : ligne " & (r - hdr) & " / " & (lastRow - hdr)
            skus(1) = CStr(ws.Cells(r, 1).Value)
            qtys(1) = ToQty(ws.Cells(r, 2).Value)
            errMsg = BuildItems(skus, qtys, 1, items, n, seasonal, info)
            If Len(errMsg) > 0 Then
                resA = errMsg: resS = errMsg
            Else
                SuggestBoth items, n, seasonal, mA, mS, resA, resS
            End If
            ws.Cells(r, 3).Value = resA
            ws.Cells(r, 4).Value = resS
            ws.Cells(r, 5).Value = info
        End If
    Next r
    ws.Columns(3).Resize(, 3).AutoFit
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If lastRow <= hdr Then MsgBox "Aucune ligne a traiter sous l'en-tete 'Item'.", vbInformation
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

Public Sub RemplirMixedPackHelper()
    Dim ws As Worksheet, wsCalc As Worksheet
    Dim r As Long, lastRow As Long, hdr As Long, outRow As Long
    Dim mA As Double, mS As Double
    Dim nOrders As Long, orderIds() As String, orderFirst() As Long
    Dim i As Long, k As Long, cnt As Long
    Dim skus() As String, qtys() As Double
    Dim resA As String, resS As String, info As String
    Dim items() As TItem, n As Long, seasonal As Boolean, errMsg As String
    Dim oid As String

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SHEET_MIX)
    Set wsCalc = ThisWorkbook.Worksheets(SHEET_CALC)
    If Not ReadMargins(wsCalc, mA, mS) Then Exit Sub
    hdr = FindRowStartingWith(ws, "Order ID", 1, 30)
    If hdr = 0 Then
        MsgBox "En-tete 'Order ID' introuvable en colonne A de la feuille '" & SHEET_MIX & "'. Lancez Packaging_Setup.", vbExclamation
        Exit Sub
    End If
    LoadData
    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' liste des commandes (ordre d'apparition)
    nOrders = 0
    ReDim orderIds(1 To 1): ReDim orderFirst(1 To 1)
    For r = hdr + 1 To lastRow
        oid = Trim(CStr(ws.Cells(r, 1).Value))
        If Len(oid) > 0 Then
            k = 0
            For i = 1 To nOrders
                If StrComp(orderIds(i), oid, vbTextCompare) = 0 Then k = i: Exit For
            Next i
            If k = 0 Then
                nOrders = nOrders + 1
                ReDim Preserve orderIds(1 To nOrders): ReDim Preserve orderFirst(1 To nOrders)
                orderIds(nOrders) = oid: orderFirst(nOrders) = r
            End If
        End If
    Next r

    ' zone resultats : colonnes E:H a partir de la ligne d'en-tete
    ws.Range(ws.Cells(hdr + 1, 5), ws.Cells(ws.Rows.Count, 8)).ClearContents
    outRow = hdr
    For i = 1 To nOrders
        Application.StatusBar = "Mixed Pack Helper : commande " & i & " / " & nOrders
        cnt = 0
        ReDim skus(1 To 1): ReDim qtys(1 To 1)
        For r = hdr + 1 To lastRow
            If StrComp(Trim(CStr(ws.Cells(r, 1).Value)), orderIds(i), vbTextCompare) = 0 Then
                If Len(Trim(CStr(ws.Cells(r, 2).Value))) > 0 Then
                    cnt = cnt + 1
                    ReDim Preserve skus(1 To cnt): ReDim Preserve qtys(1 To cnt)
                    skus(cnt) = CStr(ws.Cells(r, 2).Value)
                    qtys(cnt) = ToQty(ws.Cells(r, 3).Value)
                End If
            End If
        Next r
        If cnt = 0 Then
            resA = "Aucun article": resS = resA: info = ""
        Else
            errMsg = BuildItems(skus, qtys, cnt, items, n, seasonal, info)
            If Len(errMsg) > 0 Then
                resA = errMsg: resS = errMsg
            Else
                SuggestBoth items, n, seasonal, mA, mS, resA, resS
            End If
        End If
        outRow = outRow + 1
        ws.Cells(outRow, 5).Value = orderIds(i)
        ws.Cells(outRow, 6).Value = resA
        ws.Cells(outRow, 7).Value = resS
        ws.Cells(outRow, 8).Value = info
    Next i
    ws.Columns(5).Resize(, 4).AutoFit
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If nOrders = 0 Then MsgBox "Aucune commande trouvee sous l'en-tete 'Order ID'.", vbInformation
    Exit Sub
Fail:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Description, vbCritical
End Sub

' Cree / complete les feuilles et ajoute les boutons "Remplir". A lancer une fois.
Public Sub Packaging_Setup()
    Dim ws As Worksheet

    Set ws = GetOrCreateSheet(SHEET_CALC)
    If Len(CStr(ws.Range("A1").Value)) = 0 Then
        ws.Range("A1").Value = "Marge vide All Seasons": ws.Range("B1").Value = 0.05
        ws.Range("A2").Value = "Marge vide Summer": ws.Range("B2").Value = 0.2
        ws.Range("B1:B2").NumberFormat = "0%"
        ws.Range("B1:B2").Interior.Color = RGB(255, 255, 204)
        ws.Range("A4:E4").Value = Array("Item", "Quantity", "Suggested Package All Seasons", "Suggested Package Summer", "Info")
        ws.Range("A4:E4").Font.Bold = True
        ws.Range("A1:A2").Font.Bold = True
        ws.Columns("A:E").ColumnWidth = 28
        ws.Columns("B").ColumnWidth = 12
    End If
    AddButton ws, "btnRemplirCalc", "Remplir", "RemplirPackagingCalculator", ws.Range("D1:D2")

    Set ws = GetOrCreateSheet(SHEET_MIX)
    If Len(CStr(ws.Range("A1").Value)) = 0 Then
        ws.Range("A1:C1").Value = Array("Order ID", "Item", "Quantity")
        ws.Range("E1:H1").Value = Array("Order ID", "Suggested Package All Seasons", "Suggested Package Summer", "Info")
        ws.Range("A1:H1").Font.Bold = True
        ws.Columns("A:H").ColumnWidth = 18
        ws.Columns("F:G").ColumnWidth = 32
    End If
    AddButton ws, "btnRemplirMix", "Remplir", "RemplirMixedPackHelper", ws.Range("J1:J2")

    MsgBox "Feuilles et boutons 'Remplir' prets.", vbInformation
End Sub

' =====================================================================
'                     LECTURE DES DONNEES
' =====================================================================

Private Sub LoadData()
    LoadProducts
    LoadCaisses
End Sub

Private Sub LoadProducts()
    Dim ws As Worksheet, lastRow As Long, r As Long, k As Long
    Dim cSku As Long, cName As Long, cSeason As Long
    Dim cSub(1 To 3) As Long, cQty(1 To 3) As Long, cL(1 To 3) As Long, cW(1 To 3) As Long, cH(1 To 3) As Long
    Dim v As Variant

    Set ws = ThisWorkbook.Worksheets(SHEET_PRODUITS)
    cSku = FindCol(ws, 1, "SKU"): cName = FindCol(ws, 1, "Produit"): cSeason = FindCol(ws, 1, "Par Saison")
    If cSku = 0 Or cName = 0 Then Err.Raise vbObjectError + 1, , "Feuille Produits : colonnes 'SKU' / 'Produit' introuvables (ligne 1)."
    For k = 1 To 3
        cSub(k) = FindCol(ws, 1, "sub_Item_" & k)
        cQty(k) = FindCol(ws, 1, "quantite_" & k)
        cL(k) = FindCol(ws, 1, "Longueur_" & k)
        cW(k) = FindCol(ws, 1, "Largeur_" & k)
        cH(k) = FindCol(ws, 1, "Hauteur_" & k)
    Next k
    If cSub(1) = 0 Or cL(1) = 0 Or cW(1) = 0 Or cH(1) = 0 Then Err.Raise vbObjectError + 2, , "Feuille Produits : colonnes sub_Item_1 / Longueur_1 / Largeur_1 / Hauteur_1 introuvables."

    lastRow = ws.Cells(ws.Rows.Count, cSku).End(xlUp).Row
    mNProd = 0
    ReDim mProd(1 To IIf(lastRow > 1, lastRow - 1, 1))
    For r = 2 To lastRow
        If Len(Trim(CStr(ws.Cells(r, cSku).Value))) > 0 Then
            mNProd = mNProd + 1
            With mProd(mNProd)
                .Key = NormKey(ws.Cells(r, cSku).Value)
                .ProdName = Trim(CStr(ws.Cells(r, cName).Value))
                .NameKey = NormKey(.ProdName)
                If cSeason > 0 Then .Seasonal = (LCase(Trim(CStr(ws.Cells(r, cSeason).Value))) = "oui") Else .Seasonal = False
                .NSub = 0
                For k = 1 To 3
                    If cSub(k) > 0 Then
                        v = ws.Cells(r, cSub(k)).Value
                        If IsError(v) Then v = ""
                        If Len(Trim(CStr(v))) > 0 Then
                            .NSub = .NSub + 1
                            .SubName(.NSub) = Trim(CStr(v))
                            .SubQty(.NSub) = 1
                            If cQty(k) > 0 Then .SubQty(.NSub) = ToQty(ws.Cells(r, cQty(k)).Value)
                            If .SubQty(.NSub) <= 0 Then .SubQty(.NSub) = 1
                            .SubOK(.NSub) = False
                            If cL(k) > 0 And cW(k) > 0 And cH(k) > 0 Then
                                If IsPosNum(ws.Cells(r, cL(k)).Value) Then
                                    If IsPosNum(ws.Cells(r, cW(k)).Value) Then
                                        If IsPosNum(ws.Cells(r, cH(k)).Value) Then .SubOK(.NSub) = True
                                    End If
                                End If
                            End If
                            If .SubOK(.NSub) Then
                                .SubL(.NSub) = CDbl(ws.Cells(r, cL(k)).Value)
                                .SubW(.NSub) = CDbl(ws.Cells(r, cW(k)).Value)
                                .SubH(.NSub) = CDbl(ws.Cells(r, cH(k)).Value)
                            End If
                        End If
                    End If
                Next k
            End With
        End If
    Next r
End Sub

Private Sub LoadCaisses()
    Dim ws As Worksheet, lastRow As Long, r As Long
    Dim cName As Long, cSeason As Long, cL As Long, cW As Long, cH As Long
    Dim s As String

    Set ws = ThisWorkbook.Worksheets(SHEET_CAISSES)
    cName = FindCol(ws, 1, "Nom"): cSeason = FindCol(ws, 1, "Saison")
    cL = FindCol(ws, 1, "Longueur"): cW = FindCol(ws, 1, "Largeur"): cH = FindCol(ws, 1, "Hauteur")
    If cName = 0 Or cSeason = 0 Or cL = 0 Or cW = 0 Or cH = 0 Then Err.Raise vbObjectError + 3, , "Feuille Caisses : colonnes Nom / Saison / Longueur / Largeur / Hauteur introuvables (ligne 1)."

    lastRow = ws.Cells(ws.Rows.Count, cName).End(xlUp).Row
    mNCaisses = 0
    ReDim mCaisses(1 To IIf(lastRow > 1, lastRow - 1, 1))
    For r = 2 To lastRow
        If Len(Trim(CStr(ws.Cells(r, cName).Value))) > 0 And IsPosNum(ws.Cells(r, cL).Value) And IsPosNum(ws.Cells(r, cW).Value) And IsPosNum(ws.Cells(r, cH).Value) Then
            ' (IsPosNum est sur a toute valeur, pas de court-circuit necessaire)
            mNCaisses = mNCaisses + 1
            With mCaisses(mNCaisses)
                .Idx = mNCaisses
                .CaisseName = Trim(CStr(ws.Cells(r, cName).Value))
                s = LCase(Trim(CStr(ws.Cells(r, cSeason).Value)))
                .IsSummer = (InStr(s, "summer") > 0 Or InStr(s, "ete") > 0 Or InStr(s, "été") > 0)
                .L = CDbl(ws.Cells(r, cL).Value)
                .W = CDbl(ws.Cells(r, cW).Value)
                .H = CDbl(ws.Cells(r, cH).Value)
            End With
        End If
    Next r
    If mNCaisses = 0 Then Err.Raise vbObjectError + 4, , "Feuille Caisses : aucune caisse valide."
End Sub

Private Function FindProduct(ByVal key As String) As Long
    Dim i As Long, k As String
    k = NormKey(key)
    FindProduct = 0
    If Len(k) = 0 Then Exit Function
    For i = 1 To mNProd
        If mProd(i).Key = k Then FindProduct = i: Exit Function
    Next i
    For i = 1 To mNProd
        If mProd(i).NameKey = k Then FindProduct = i: Exit Function
    Next i
End Function

' Construit la liste d'articles (boites) d'une commande. Retourne "" ou un message d'erreur.
Private Function BuildItems(ByRef skus() As String, ByRef qtys() As Double, ByVal nLines As Long, _
                            ByRef items() As TItem, ByRef n As Long, ByRef seasonal As Boolean, ByRef info As String) As String
    Dim i As Long, p As Long, k As Long, c As Long, cnt As Long, tot As Long
    Dim totVol As Double, names As String

    n = 0: seasonal = False: info = "": totVol = 0: tot = 0
    ReDim items(1 To 1)
    For i = 1 To nLines
        If qtys(i) <= 0 Then
            BuildItems = "Quantite invalide pour '" & Trim(skus(i)) & "'": Exit Function
        End If
        p = FindProduct(skus(i))
        If p = 0 Then
            BuildItems = "Produit inconnu : '" & Trim(skus(i)) & "'": Exit Function
        End If
        With mProd(p)
            If .NSub = 0 Then BuildItems = "Aucun sous-article pour " & .ProdName: Exit Function
            If .Seasonal Then seasonal = True
            If Len(names) > 0 Then names = names & " ; "
            names = names & qtys(i) & " x " & .ProdName
            For k = 1 To .NSub
                If Not .SubOK(k) Then
                    BuildItems = "Dimensions manquantes : " & .ProdName & " / " & .SubName(k): Exit Function
                End If
                cnt = CLng(Round(.SubQty(k) * qtys(i), 0))
                For c = 1 To cnt
                    n = n + 1
                    If n > MAX_ITEMS Then BuildItems = "Trop d'articles (> " & MAX_ITEMS & ")": Exit Function
                    If n > UBound(items) Then ReDim Preserve items(1 To UBound(items) * 2)
                    items(n).L = .SubL(k): items(n).W = .SubW(k): items(n).H = .SubH(k)
                    items(n).Vol = .SubL(k) * .SubW(k) * .SubH(k)
                    items(n).MaxDim = Max3(.SubL(k), .SubW(k), .SubH(k))
                    items(n).Ref = .ProdName & " / " & .SubName(k)
                    items(n).Done = False: items(n).Placed = False
                    totVol = totVol + items(n).Vol
                Next c
            Next k
        End With
    Next i
    If n = 0 Then BuildItems = "Aucun article": Exit Function
    BuildOrders items, n
    info = names & " | " & n & " colis | " & Format(totVol / 1000, "0.0") & " L | Saisonnier: " & IIf(seasonal, "Oui", "Non")
    BuildItems = ""
End Function

' ordres de tri : volume decroissant, et plus grande dimension decroissante
Private Sub BuildOrders(ByRef items() As TItem, ByVal n As Long)
    Dim i As Long
    Dim keyVol() As Double, keyDim() As Double
    ReDim mOrdVol(1 To n): ReDim mOrdDim(1 To n)
    ReDim keyVol(1 To n): ReDim keyDim(1 To n)
    For i = 1 To n
        mOrdVol(i) = i: mOrdDim(i) = i
        keyVol(i) = items(i).Vol
        keyDim(i) = items(i).MaxDim * 1000000# + items(i).Vol * 0.000001
    Next i
    SortDesc mOrdVol, keyVol, 1, n
    SortDesc mOrdDim, keyDim, 1, n
End Sub

' tri rapide (descendant) d'un tableau d'index selon des cles
Private Sub SortDesc(ByRef idx() As Long, ByRef key() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Double, t As Long
    If lo >= hi Then Exit Sub
    i = lo: j = hi
    pivot = key(idx((lo + hi) \ 2))
    Do While i <= j
        Do While key(idx(i)) > pivot
            i = i + 1
        Loop
        Do While key(idx(j)) < pivot
            j = j - 1
        Loop
        If i <= j Then
            t = idx(i): idx(i) = idx(j): idx(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortDesc idx, key, lo, j
    If i < hi Then SortDesc idx, key, i, hi
End Sub

' =====================================================================
'                  SUGGESTION (All Seasons + Summer)
' =====================================================================

Private Sub SuggestBoth(ByRef items() As TItem, ByVal n As Long, ByVal seasonal As Boolean, _
                        ByVal mA As Double, ByVal mS As Double, ByRef resA As String, ByRef resS As String)
    resA = Suggest(items, n, False, mA)
    If seasonal Then
        resS = Suggest(items, n, True, mS)
    Else
        resS = resA   ' produit non saisonnier : emballage All Seasons toute l'annee
    End If
End Sub

' Retourne la combinaison de caisses ("1:Caisse V2_1:Caisse V5") ou un message d'erreur.
Private Function Suggest(ByRef items() As TItem, ByVal n As Long, ByVal summer As Boolean, ByVal margin As Double) As String
    Dim cs() As Long, nc As Long, j As Long, i As Long
    Dim uL() As Double, uW() As Double, uH() As Double, uV() As Double, ord() As Long
    Dim counts() As Long
    Dim remVol As Double, remaining As Long, chosen As Long, pv As Double
    Dim bestJ As Long, bestV As Double, bestFlags() As Boolean
    Dim guard As Long, res As String, t As Long, a As Long, b As Long

    ' caisses de la saison
    nc = 0
    ReDim cs(1 To mNCaisses)
    For j = 1 To mNCaisses
        If mCaisses(j).IsSummer = summer Then nc = nc + 1: cs(nc) = j
    Next j
    If nc = 0 Then
        Suggest = "Aucune caisse '" & IIf(summer, "Summer season", "All seasons") & "' dans la feuille Caisses"
        Exit Function
    End If
    ReDim uL(1 To nc): ReDim uW(1 To nc): ReDim uH(1 To nc): ReDim uV(1 To nc): ReDim ord(1 To nc): ReDim counts(1 To nc)
    For j = 1 To nc
        uL(j) = mCaisses(cs(j)).L
        uW(j) = mCaisses(cs(j)).W
        uH(j) = mCaisses(cs(j)).H * (1 - margin)
        uV(j) = uL(j) * uW(j) * uH(j)
        ord(j) = j
    Next j
    ' tri des caisses par volume utile croissant (tri par insertion, peu de caisses)
    For a = 2 To nc
        t = ord(a): b = a - 1
        Do While b >= 1
            If uV(ord(b)) <= uV(t) Then Exit Do
            ord(b + 1) = ord(b): b = b - 1
        Loop
        ord(b + 1) = t
    Next a

    For i = 1 To n: items(i).Done = False: Next i
    ReDim bestFlags(1 To n)
    guard = 0
    Do
        remaining = 0: remVol = 0
        For i = 1 To n
            If Not items(i).Done Then remaining = remaining + 1: remVol = remVol + items(i).Vol
        Next i
        If remaining = 0 Then Exit Do
        guard = guard + 1
        If guard > MAX_CAISSES Then Suggest = "Trop de caisses (> " & MAX_CAISSES & ")": Exit Function

        ' 1) tout le reste tient-il dans une seule caisse ? -> la plus petite qui convient
        chosen = 0
        For a = 1 To nc
            j = ord(a)
            If remVol <= uV(j) + EPS Then
                If PackBest(items, n, uL(j), uW(j), uH(j), pv) Then chosen = j: Exit For
            End If
        Next a
        If chosen > 0 Then
            counts(chosen) = counts(chosen) + 1
            For i = 1 To n
                If items(i).Placed Then items(i).Done = True
            Next i
            Exit Do
        End If

        ' 2) sinon : la caisse qui place le plus de volume (a egalite, la plus petite)
        bestJ = 0: bestV = -1
        For a = 1 To nc
            j = ord(a)
            PackBest items, n, uL(j), uW(j), uH(j), pv
            If pv > bestV + EPS Then
                bestV = pv: bestJ = j
                For i = 1 To n: bestFlags(i) = items(i).Placed: Next i
            End If
        Next a
        If bestV <= EPS Then
            For i = 1 To n
                If Not items(i).Done Then Suggest = "Article trop grand pour toutes les caisses : " & items(i).Ref: Exit Function
            Next i
        End If
        counts(bestJ) = counts(bestJ) + 1
        For i = 1 To n
            If bestFlags(i) Then items(i).Done = True
        Next i
    Loop

    res = ""
    For j = 1 To nc   ' ordre de la feuille Caisses
        If counts(j) > 0 Then
            If Len(res) > 0 Then res = res & "_"
            res = res & counts(j) & ":" & mCaisses(cs(j)).CaisseName
        End If
    Next j
    Suggest = res
End Function

' =====================================================================
'                   PLACEMENT 3D (espaces maximaux)
' =====================================================================

' Essaie plusieurs strategies, garde la meilleure (items(i).Placed), retourne True si tout est place.
Private Function PackBest(ByRef items() As TItem, ByVal n As Long, ByVal boxL As Double, ByVal boxW As Double, ByVal boxH As Double, _
                          ByRef placedVol As Double) As Boolean
    Dim st As Long, sortMode As Long, orientPref As Long, keyMode As Long
    Dim allP As Boolean, pv As Double, bestV As Double, bestAll As Boolean
    Dim bestFlags() As Boolean, i As Long, remaining As Long
    Static strat(1 To 6, 1 To 3) As Long
    Static stratInit As Boolean

    If Not stratInit Then
        ' (tri, orientation preferee, choix de l'espace) : 0=volume desc / 1=plus grande dim desc ; 0=plus bas / 1=plus petite empreinte ; 0=premier espace (z,y,x) / 1=meilleur ajustement
        strat(1, 1) = 0: strat(1, 2) = 0: strat(1, 3) = 0
        strat(2, 1) = 0: strat(2, 2) = 1: strat(2, 3) = 0
        strat(3, 1) = 1: strat(3, 2) = 0: strat(3, 3) = 0
        strat(4, 1) = 1: strat(4, 2) = 1: strat(4, 3) = 0
        strat(5, 1) = 0: strat(5, 2) = 0: strat(5, 3) = 1
        strat(6, 1) = 0: strat(6, 2) = 1: strat(6, 3) = 1
        stratInit = True
    End If

    remaining = 0
    For i = 1 To n
        If Not items(i).Done Then remaining = remaining + 1
    Next i
    ReDim bestFlags(1 To n)
    bestV = -1: bestAll = False
    For st = 1 To 6
        sortMode = strat(st, 1): orientPref = strat(st, 2): keyMode = strat(st, 3)
        If Not (remaining > BIG_N And keyMode = 1) Then
            If sortMode = 0 Then
                allP = PackOnce(items, n, mOrdVol, boxL, boxW, boxH, orientPref, keyMode, pv)
            Else
                allP = PackOnce(items, n, mOrdDim, boxL, boxW, boxH, orientPref, keyMode, pv)
            End If
            If pv > bestV + EPS Then
                bestV = pv: bestAll = allP
                For i = 1 To n: bestFlags(i) = items(i).Placed: Next i
            End If
            If allP Then Exit For
        End If
    Next st
    For i = 1 To n: items(i).Placed = bestFlags(i): Next i
    placedVol = bestV
    PackBest = bestAll
End Function

' Une passe de placement. Retourne True si tous les articles restants sont places.
Private Function PackOnce(ByRef items() As TItem, ByVal n As Long, ByRef ord() As Long, _
                          ByVal boxL As Double, ByVal boxW As Double, ByVal boxH As Double, _
                          ByVal orientPref As Long, ByVal keyMode As Long, ByRef placedVol As Double) As Boolean
    Dim sp() As TSpace, nSp As Long
    Dim ns() As TSpace, nNs As Long
    Dim keep() As Boolean
    Dim allPlaced As Boolean
    Dim k As Long, i As Long, s As Long, o As Long, nOr As Long
    Dim ol As Double, ow As Double, oh As Double, op As Double
    Dim fits As Boolean, sBestOp As Double, sL As Double, sW As Double, sH As Double
    Dim found As Boolean, bestS As Long, bestKey As Double, bL As Double, bW As Double, bH As Double
    Dim px As Double, py As Double, pz As Double, pl As Double, pw As Double, ph As Double
    Dim sKey As Double, a As Long, b As Long, tmp As TSpace

    If ALLOW_TIP_OVER Then nOr = 6 Else nOr = 2
    ReDim sp(1 To 1): nSp = 1
    sp(1).X = 0: sp(1).Y = 0: sp(1).Z = 0: sp(1).L = boxL: sp(1).W = boxW: sp(1).H = boxH
    allPlaced = True: placedVol = 0
    For i = 1 To n: items(i).Placed = False: Next i

    For k = 1 To n
        i = ord(k)
        If Not items(i).Done Then
            ' ---- choix de l'espace et de l'orientation ----
            found = False: bestS = 0: bestKey = 0
            For s = 1 To nSp
                If found And keyMode = 0 Then Exit For          ' espaces tries (z,y,x) : le premier qui convient est le meilleur
                sKey = sp(s).L * sp(s).W * sp(s).H
                If Not (found And keyMode = 1 And sKey >= bestKey - EPS) Then   ' best fit : seulement les espaces plus petits
                    fits = False
                    For o = 1 To nOr
                        Orient items(i), o, ol, ow, oh
                        If ol <= sp(s).L + EPS And ow <= sp(s).W + EPS And oh <= sp(s).H + EPS Then
                            If orientPref = 0 Then op = oh Else op = ol * ow
                            If (Not fits) Or op < sBestOp - EPS Then
                                fits = True: sBestOp = op: sL = ol: sW = ow: sH = oh
                            End If
                        End If
                    Next o
                    If fits Then
                        found = True: bestS = s: bestKey = sKey: bL = sL: bW = sW: bH = sH
                    End If
                End If
            Next s

            If Not found Then
                allPlaced = False
            Else
                px = sp(bestS).X: py = sp(bestS).Y: pz = sp(bestS).Z
                pl = bL: pw = bW: ph = bH
                items(i).Placed = True
                placedVol = placedVol + items(i).Vol

                ' ---- decoupe des espaces touches ----
                ReDim ns(1 To nSp * 6 + 1): nNs = 0
                For s = 1 To nSp
                    With sp(s)
                        If px >= .X + .L - EPS Or px + pl <= .X + EPS Or py >= .Y + .W - EPS Or py + pw <= .Y + EPS _
                           Or pz >= .Z + .H - EPS Or pz + ph <= .Z + EPS Then
                            AddSpace ns, nNs, .X, .Y, .Z, .L, .W, .H
                        Else
                            If px - .X > EPS Then AddSpace ns, nNs, .X, .Y, .Z, px - .X, .W, .H
                            If .X + .L - (px + pl) > EPS Then AddSpace ns, nNs, px + pl, .Y, .Z, .X + .L - (px + pl), .W, .H
                            If py - .Y > EPS Then AddSpace ns, nNs, .X, .Y, .Z, .L, py - .Y, .H
                            If .Y + .W - (py + pw) > EPS Then AddSpace ns, nNs, .X, py + pw, .Z, .L, .Y + .W - (py + pw), .H
                            If pz - .Z > EPS Then AddSpace ns, nNs, .X, .Y, .Z, .L, .W, pz - .Z
                            If .Z + .H - (pz + ph) > EPS Then AddSpace ns, nNs, .X, .Y, pz + ph, .L, .W, .Z + .H - (pz + ph)
                        End If
                    End With
                Next s

                ' ---- suppression des espaces inclus dans un autre ----
                ReDim keep(1 To nNs)
                For a = 1 To nNs
                    keep(a) = True
                    For b = 1 To nNs
                        If b <> a Then
                            If Contains(ns(b), ns(a)) Then
                                If SameSpace(ns(a), ns(b)) Then
                                    If b < a Then keep(a) = False: Exit For
                                Else
                                    keep(a) = False: Exit For
                                End If
                            End If
                        End If
                    Next b
                Next a
                nSp = 0
                ReDim sp(1 To nNs)
                For a = 1 To nNs
                    If keep(a) Then nSp = nSp + 1: sp(nSp) = ns(a)
                Next a
                ' ---- tri (z, y, x) ----
                For a = 2 To nSp
                    tmp = sp(a): b = a - 1
                    Do While b >= 1
                        If Not SpaceBefore(tmp, sp(b)) Then Exit Do
                        sp(b + 1) = sp(b): b = b - 1
                    Loop
                    sp(b + 1) = tmp
                Next a
            End If
        End If
    Next k
    PackOnce = allPlaced
End Function

Private Sub Orient(ByRef it As TItem, ByVal o As Long, ByRef ol As Double, ByRef ow As Double, ByRef oh As Double)
    Select Case o
        Case 1: ol = it.L: ow = it.W: oh = it.H
        Case 2: ol = it.W: ow = it.L: oh = it.H
        Case 3: ol = it.L: ow = it.H: oh = it.W
        Case 4: ol = it.H: ow = it.L: oh = it.W
        Case 5: ol = it.W: ow = it.H: oh = it.L
        Case Else: ol = it.H: ow = it.W: oh = it.L
    End Select
End Sub

Private Sub AddSpace(ByRef ns() As TSpace, ByRef nNs As Long, ByVal X As Double, ByVal Y As Double, ByVal Z As Double, _
                     ByVal L As Double, ByVal W As Double, ByVal H As Double)
    nNs = nNs + 1
    If nNs > UBound(ns) Then ReDim Preserve ns(1 To UBound(ns) * 2)
    ns(nNs).X = X: ns(nNs).Y = Y: ns(nNs).Z = Z
    ns(nNs).L = L: ns(nNs).W = W: ns(nNs).H = H
End Sub

' a est inclus dans b ?
Private Function Contains(ByRef b As TSpace, ByRef a As TSpace) As Boolean
    Contains = (a.X >= b.X - EPS And a.Y >= b.Y - EPS And a.Z >= b.Z - EPS And _
                a.X + a.L <= b.X + b.L + EPS And a.Y + a.W <= b.Y + b.W + EPS And a.Z + a.H <= b.Z + b.H + EPS)
End Function

Private Function SameSpace(ByRef a As TSpace, ByRef b As TSpace) As Boolean
    SameSpace = (Abs(a.X - b.X) <= EPS And Abs(a.Y - b.Y) <= EPS And Abs(a.Z - b.Z) <= EPS And _
                 Abs(a.L - b.L) <= EPS And Abs(a.W - b.W) <= EPS And Abs(a.H - b.H) <= EPS)
End Function

' a avant b dans l'ordre (z, y, x) ?
Private Function SpaceBefore(ByRef a As TSpace, ByRef b As TSpace) As Boolean
    If a.Z < b.Z - EPS Then SpaceBefore = True: Exit Function
    If a.Z > b.Z + EPS Then SpaceBefore = False: Exit Function
    If a.Y < b.Y - EPS Then SpaceBefore = True: Exit Function
    If a.Y > b.Y + EPS Then SpaceBefore = False: Exit Function
    SpaceBefore = (a.X < b.X - EPS)
End Function

' =====================================================================
'                           UTILITAIRES
' =====================================================================

Private Function ReadMargins(ByVal ws As Worksheet, ByRef mA As Double, ByRef mS As Double) As Boolean
    Dim rA As Long, rS As Long
    ReadMargins = False
    rA = FindRowStartingWith(ws, "Marge vide All", 1, 10)
    rS = FindRowStartingWith(ws, "Marge vide Summer", 1, 10)
    If rA = 0 Or rS = 0 Then
        MsgBox "Cellules 'Marge vide All Seasons' / 'Marge vide Summer' introuvables (colonne A, feuille '" & SHEET_CALC & "'). Lancez Packaging_Setup.", vbExclamation
        Exit Function
    End If
    mA = ToMargin(ws.Cells(rA, 2).Value)
    mS = ToMargin(ws.Cells(rS, 2).Value)
    If mA < 0 Or mA >= 1 Or mS < 0 Or mS >= 1 Then
        MsgBox "Les marges doivent etre entre 0% et 99%.", vbExclamation
        Exit Function
    End If
    ReadMargins = True
End Function

' accepte 0.05, 5 (=5%) ou "5%"
Private Function ToMargin(ByVal v As Variant) As Double
    Dim s As String, m As Double
    If IsError(v) Then ToMargin = -1: Exit Function
    If IsNumeric(v) Then
        m = CDbl(v)
    Else
        s = Replace(Trim(CStr(v)), "%", "")
        s = Replace(s, ",", ".")
        If IsNumeric(s) Then
            m = Val(s) / 100
        Else
            ToMargin = -1: Exit Function
        End If
    End If
    If m > 1 Then m = m / 100
    ToMargin = m
End Function

Private Function ToQty(ByVal v As Variant) As Double
    ToQty = 0
    If IsError(v) Then Exit Function
    If IsNumeric(v) Then ToQty = CDbl(v)
End Function

Private Function IsPosNum(ByVal v As Variant) As Boolean
    IsPosNum = False
    If IsError(v) Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    If Len(CStr(v)) = 0 Then Exit Function
    IsPosNum = (CDbl(v) > 0)
End Function

Private Function Max3(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    Dim m As Double
    m = a
    If b > m Then m = b
    If c > m Then m = c
    Max3 = m
End Function

' cle de comparaison : nombres -> "9004", textes -> minuscules sans espaces superflus
Private Function NormKey(ByVal v As Variant) As String
    Dim s As String
    s = Trim(CStr(v))
    If IsNumeric(s) Then
        NormKey = CStr(CDbl(s))
    Else
        NormKey = LCase(s)
    End If
End Function

' premiere colonne (ligne hdrRow) dont l'en-tete commence par prefix
Private Function FindCol(ByVal ws As Worksheet, ByVal hdrRow As Long, ByVal prefix As String) As Long
    Dim c As Long, lastCol As Long, s As String
    lastCol = ws.Cells(hdrRow, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        s = LCase(Trim(CStr(ws.Cells(hdrRow, c).Value)))
        If Left(s, Len(prefix)) = LCase(prefix) Then FindCol = c: Exit Function
    Next c
    FindCol = 0
End Function

Private Function FindRowStartingWith(ByVal ws As Worksheet, ByVal prefix As String, ByVal col As Long, ByVal maxRow As Long) As Long
    Dim r As Long, s As String
    For r = 1 To maxRow
        s = LCase(Trim(CStr(ws.Cells(r, col).Value)))
        If Left(s, Len(prefix)) = LCase(prefix) Then FindRowStartingWith = r: Exit Function
    Next r
    FindRowStartingWith = 0
End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Sub AddButton(ByVal ws As Worksheet, ByVal btnName As String, ByVal caption As String, ByVal macro As String, ByVal anchor As Range)
    Dim btn As Object
    On Error Resume Next
    ws.Buttons(btnName).Delete
    On Error GoTo 0
    Set btn = ws.Buttons.Add(anchor.Left, anchor.Top, anchor.Width, anchor.Height)
    btn.Name = btnName
    btn.Caption = caption
    btn.OnAction = macro
End Sub
