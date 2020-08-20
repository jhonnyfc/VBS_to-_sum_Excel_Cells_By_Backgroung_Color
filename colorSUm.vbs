'################################################################### 
'##         Script Sum Cells By Back Ground                       ## 
'##         Date: 10-08-2020                                      ##
'##         modified by: Jhonny F. Chicaiza -> github: jhonnyfc   ## 
'###################################################################

Function Sumarcolor(RangoEval As Range, Celdacolor As Range, RangoSuma As Range) As Double
    Dim celSum As Range
	Dim ind As Integer
	ind = 1

    For Each celSum In RangoSuma
        If RangoEval.Cells(ind, 1).Interior.ColorIndex = Celdacolor.Cells(1, 1).Interior.ColorIndex Then
            Sumarcolor = Sumarcolor + celSum
        End If
		ind = ind + 1
    Next celSum
    Set celSum = Nothing
End Function