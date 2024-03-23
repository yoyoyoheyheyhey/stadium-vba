Private PRIMARY_CATEGORY_NAMES As Variant
Private DRINK_SECONDARY_CATEGORY_NAMES As Variant
Private REQUEST_SECONDARY_CATEGORY_NAMES As Variant
Private OUTSIDE_SECONDARY_SALE_CATEGORY_NAMES As Variant
Private OTHER_SECONDARY_ALLOWANCE_CATEGORY_NAMES As Variant
Private Sub InitializeCategoryArrays()
    PRIMARY_CATEGORY_NAMES = Array("ドリンク", "リクエスト", "外販", "その他手当")
    ' ドリンク
    DRINK_SECONDARY_CATEGORY_NAMES = Array("ドリンク", "ドリンク調整", "シャンパン", "系列ドリンク")
    ' リクエスト
    REQUEST_SECONDARY_CATEGORY_NAMES = Array("リクエスト", "系列リクエスト")
    ' 外販
    OUTSIDE_SECONDARY_SALE_CATEGORY_NAMES = Array("外販手当")
    ' その他手当
    OTHER_SECONDARY_ALLOWANCE_CATEGORY_NAMES = Array("同伴本指名手当", "その他", "交通費")
End Sub

Function DesiredOrderColumnStructure() As Collection
    Dim categories As New Collection
    Dim primary As PrimaryCategory
    Dim secondary As SecondaryCategory
    Dim secondaryCategoryNameArrays As Variant
    Dim secondaryCategoryName As Variant

    Call InitializeCategoryArrays

    secondaryCategoryNameArrays = Array( _
        DRINK_SECONDARY_CATEGORY_NAMES, _
        REQUEST_SECONDARY_CATEGORY_NAMES, _
        OUTSIDE_SECONDARY_SALE_CATEGORY_NAMES, _
        OTHER_SECONDARY_ALLOWANCE_CATEGORY_NAMES _
    )

    For i = LBound(PRIMARY_CATEGORY_NAMES) To UBound(PRIMARY_CATEGORY_NAMES)
        Set primary = New PrimaryCategory
        primary.Name = PRIMARY_CATEGORY_NAMES(i)
        
        For Each secondaryCategoryName In secondaryCategoryNameArrays(i)
            Set secondary = New SecondaryCategory
            secondary.Name = secondaryCategoryName
            primary.AddChild secondary
        Next secondaryCategoryName
        
        categories.Add primary
    Next i
    
    Set DesiredOrderColumnStructure = categories
End Function

Function DesiredOrderSecondaryColumns() As Variant
    Dim secondaryCategoryNameArrays As Variant

    Call InitializeCategoryArrays

    secondaryCategoryNameArrays = Array( _
        DRINK_SECONDARY_CATEGORY_NAMES, _
        REQUEST_SECONDARY_CATEGORY_NAMES, _
        OUTSIDE_SECONDARY_SALE_CATEGORY_NAMES, _
        OTHER_SECONDARY_ALLOWANCE_CATEGORY_NAMES _
    )

    DesiredOrderSecondaryColumns = merge1DArray(secondaryCategoryNameArrays)
End Function

Private Function merge1DArray(arr As Variant) As Variant
    Dim mergedArray() As Variant
    Dim totalSize As Long
    Dim i As Long, j As Long, index As Long: index = 1 ' 1ベースのインデックスで開始

    ' 合計サイズを計算
    For i = LBound(arr) To UBound(arr)
        totalSize = totalSize + (UBound(arr(i)) - LBound(arr(i)) + 1)
    Next i

    ' mergedArrayを合計サイズに基づいて初期化
    ReDim mergedArray(1 To totalSize)

    ' arrの各配列をmergedArrayにマージ
    For i = LBound(arr) To UBound(arr)
        For j = LBound(arr(i)) To UBound(arr(i))
            mergedArray(index) = arr(i)(j)
            index = index + 1
        Next j
    Next i

    merge1DArray = mergedArray
End Function
