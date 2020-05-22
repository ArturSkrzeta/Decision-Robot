Option Explicit

Dim scriptApp As Excel.Application
Dim scriptwb As Workbook
Dim robotWs As Worksheet

Dim questions() As Variant
Dim categories() As Variant

Dim answersForCat1() As Variant
Dim answersForCat2() As Variant
Dim answersForCat3() As Variant
Dim answersForCat4() As Variant
Dim answersForCat5() As Variant

Dim firstAnswerFromUser As String
Dim secondAnswerFromUser As String
Dim thirdAnswerFromUser As String

Dim scoreForCat(4) As Integer

Sub ShowQuestionsForm()
    frmQuest.Show
End Sub

Public Sub MainForRobot()

    Set scriptApp = Excel.Application
    Set scriptwb = scriptApp.ThisWorkbook
    Set robotWs = scriptwb.Sheets("ROBOT")
    
    Call CleanUp(robotWs)
    
    ' - - - 5 categories // 3 questions // 3 answers - - -
    ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    
    Let questions = Array("Question 1", "Question 2", "Question 3") ' - - - not is use
    Let categories = Array("Category 1", "Category 2", "Category 3", "Category 4", "Category 5")
    
    Let answersForCat1 = Array("a", "a", "a")
    Let answersForCat2 = Array("b", "b", "b")
    Let answersForCat3 = Array("c", "c", "c")
    Let answersForCat4 = Array("a", "b", "c")
    Let answersForCat5 = Array("c", "b", "a")
    
    Let firstAnswerFromUser = robotWs.Range("H4").value
    Let secondAnswerFromUser = robotWs.Range("H5").value
    Let thirdAnswerFromUser = robotWs.Range("H6").value
    
    Dim i As Integer
    Dim answer As String
    
    On Error GoTo ErrorHandler
    
    ' - - - Taking answers and scoring categories
    
    For i = 1 To 3
        
        ' - - = as there are 3 questions
        If i = 1 Then answer = firstAnswerFromUser
        If i = 2 Then answer = secondAnswerFromUser
        If i = 3 Then answer = thirdAnswerFromUser
        
        ' - - - scoring for 5 categories
        If answer = answersForCat1(i - 1) Then scoreForCat(0) = scoreForCat(0) + 1
        If answer = answersForCat2(i - 1) Then scoreForCat(1) = scoreForCat(1) + 1
        If answer = answersForCat3(i - 1) Then scoreForCat(2) = scoreForCat(2) + 1
        If answer = answersForCat4(i - 1) Then scoreForCat(3) = scoreForCat(3) + 1
        If answer = answersForCat5(i - 1) Then scoreForCat(4) = scoreForCat(4) + 1
    
    Next i
    
    ' - - - Answers to Worksheet for a Chart Range
    
    For i = 0 To 4
        
        ' - - - 5 categories
        If i = 0 Then answer = scoreForCat(0)
        If i = 1 Then answer = scoreForCat(1)
        If i = 2 Then answer = scoreForCat(2)
        If i = 3 Then answer = scoreForCat(3)
        If i = 4 Then answer = scoreForCat(4)
        
        robotWs.Range("B3").Offset(0, i).value = answer
    
    Next i
    
    ' - - - Creating Dictionary Key:="Category X", Item:=ScoreForCatX and soriting it out
    
    Dim scoreDict As Object
    Set scoreDict = CreateObject("Scripting.Dictionary")
    
    With scoreDict
        ' - - - 5 categories - 5 scores
        .Add categories(0), scoreForCat(0)
        .Add categories(1), scoreForCat(1)
        .Add categories(2), scoreForCat(2)
        .Add categories(3), scoreForCat(3)
        .Add categories(4), scoreForCat(4)
    End With

    Dim sortedDict As Object
    Set sortedDict = SortDictionaryByValue(scoreDict)
    
    ' - - - Loop through dictionary's keys
    
    Dim scoreRng As Range
    Set scoreRng = robotWs.Range("B2").CurrentRegion
    Set scoreRng = scoreRng.Resize(1, scoreRng.Columns.Count).Offset(1, 0)
    
    Dim msgStr As String
    
    Dim k As Variant
    i = 1
    
    For Each k In sortedDict.Keys
    
        If i = 1 And scriptApp.WorksheetFunction.CountIf(scoreRng, sortedDict(k)) = 1 Then
            msgStr = "The best recommended category is: " & k
        ElseIf i = 1 And scriptApp.WorksheetFunction.CountIf(scoreRng, sortedDict(k)) >= 2 Then
            msgStr = "There are more than one best recommended category. Please check the outcome on chart and table next to it."
        End If
        
        ' - - - Outcome to worksheet
        
        With robotWs.Range("AE6")
            .Offset(i, 0).value = k & ":"
            robotWs.Range("AE6").Offset(i, 1).value = sortedDict(k)
        End With
        
        i = i + 1
        
    Next
    
    Call CreateRadarChart(robotWs)
    
    MsgBox msgStr, vbInformation
    
ErrorHandler:
    
    Erase scoreForCat
    Erase questions
    Erase categories
    Erase answersForCat1
    Erase answersForCat2
    Erase answersForCat3
    Erase answersForCat4
    Erase answersForCat5
    
    If Err <> 0 Then MsgBox "Error occured", vbExclamation
    
End Sub

Public Sub CleanUp(robotWs As Worksheet)

    Dim sh As Shape
    
    For Each sh In robotWs.Shapes
        If InStr(1, sh.name, "Chart") <> 0 Then
            sh.Delete
        End If
    Next sh
    
    With robotWs
        .Range("W7").CurrentRegion.Clear
        .Range("AE7").CurrentRegion.Clear
    End With

End Sub

Private Sub CreateRadarChart(robotWs As Worksheet)

    Dim dataRng As Range
    
    Set dataRng = robotWs.Range("B2").CurrentRegion
    Set dataRng = dataRng.Resize(2, dataRng.Columns.Count)
    dataRng.Copy
    robotWs.Range("W7").PasteSpecial xlPasteValues
    
    robotWs.Shapes.AddChart2(317, xlRadar).Select
    ActiveChart.SetSourceData Source:=robotWs.Range("W7").CurrentRegion
    
    With robotWs.ChartObjects(1)
        .Top = robotWs.Range("W7").Top
        .Left = robotWs.Range("W7").Left
        .Height = robotWs.Shapes("robotPicture").Height
    End With
    
    ActiveChart.ChartStyle = 320
    ActiveChart.ChartTitle.Text = "Buying Channel Decision Matrix"
'    ChartTitle.TextFrame2.TextRange.Font.Size = 15

End Sub

Public Function SortDictionaryByValue(dict As Object, Optional sortorder As XlSortOrder = xlDescending) As Object
    
    On Error GoTo eh
    
    Dim arrayList As Object
    Set arrayList = CreateObject("System.Collections.ArrayList")
    
    Dim dictTemp As Object
    Set dictTemp = CreateObject("Scripting.Dictionary")
   
    ' Put values in ArrayList and sort
    ' Store values in tempDict with their keys as a collection
    Dim key As Variant, value As Variant, coll As Collection
    For Each key In dict
    
        value = dict(key)
        
        ' if the value doesn't exist in dict then add
        If dictTemp.exists(value) = False Then
            ' create collection to hold keys
            ' - needed for duplicate values
            Set coll = New Collection
            dictTemp.Add value, coll
            
            ' Add the value
            arrayList.Add value
            
        End If
        
        ' Add the current key to the collection
        dictTemp(value).Add key
    
    Next key
    
    ' Sort the value
    arrayList.Sort
    
    ' Reverse if descending
    If sortorder = xlDescending Then
        arrayList.Reverse
    End If
    
    dict.RemoveAll
    
    ' Read through the ArrayList and add the values and corresponding
    ' keys from the dictTemp
    Dim item As Variant
    For Each value In arrayList
        Set coll = dictTemp(value)
        For Each item In coll
            dict.Add item, value
        Next item
    Next value
    
    Set arrayList = Nothing
    
    ' Return the new dictionary
    Set SortDictionaryByValue = dict
        
Done:
    Exit Function
eh:
    If Err.Number = 450 Then
        Err.Raise vbObjectError + 100, "SortDictionaryByValue" _
                , "Cannot sort the dictionary if the value is an object"
    End If
    
End Function
