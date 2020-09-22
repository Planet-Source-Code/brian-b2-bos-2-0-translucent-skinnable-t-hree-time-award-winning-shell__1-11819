Attribute VB_Name = "modQuickSort"
' Thanks to Colin Woor for this function
' His website: http://www.woor.co.uk/
' His PSC bubble sort function: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=5799&lngWId=1
'
' Note: This code is NOT subject to the terms and conditions of the GPL

Public Function FastSort(tmparray)

    Dim SortedArray As Boolean
    Dim start, Finish As Integer
    SortedArray = True
    start = LBound(tmparray)
    Finish = UBound(tmparray)


    Do
        SortedArray = True


        For loopcount = start To Finish - 1


            If UCase(tmparray(loopcount)) > UCase(tmparray(loopcount + 1)) Then
                SortedArray = False
                Call swap(tmparray, loopcount, loopcount + 1)
            End If
        Next loopcount
        'start = start + 1
    Loop Until SortedArray = True
    FastSort = tmparray
End Function


Sub swap(swparray, fpos, spos)
    Dim temp As Variant
    temp = swparray(fpos)
    swparray(fpos) = swparray(spos)
    swparray(spos) = temp
End Sub

