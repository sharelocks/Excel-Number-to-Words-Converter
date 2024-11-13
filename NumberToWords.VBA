Function NumberToWords(ByVal MyNumber)
    Dim Units As Variant, Tens As Variant
    Dim Temp As String, DecimalPlace As Integer, Count As Integer
    Dim CentsPart As String
    
    Units = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
    Tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
    
    MyNumber = Trim(CStr(MyNumber))
    DecimalPlace = InStr(MyNumber, ".")
    If DecimalPlace > 0 Then
        CentsPart = Mid(MyNumber, DecimalPlace + 1)
        If Len(CentsPart) = 1 Then CentsPart = CentsPart & "0" ' Ensure two decimal places for cents
        CentsPart = " and " & ChunkToWords(CentsPart) & " Cents"
        MyNumber = Left(MyNumber, DecimalPlace - 1)
    Else
        CentsPart = ""
    End If
    
    Dim Result As String
    Result = ""
    
    Dim Place As Variant
    Place = Array("", " Thousand", " Million", " Billion")
    
    Dim PlaceIndex As Integer
    PlaceIndex = 0
    
    Do While MyNumber <> ""
        Dim NumberChunk As String
        If Len(MyNumber) > 3 Then
            NumberChunk = Right(MyNumber, 3)
            MyNumber = Left(MyNumber, Len(MyNumber) - 3)
        Else
            NumberChunk = MyNumber
            MyNumber = ""
        End If
        
        If Val(NumberChunk) > 0 Then
            Result = ChunkToWords(NumberChunk) & Place(PlaceIndex) & " " & Result
        End If
        PlaceIndex = PlaceIndex + 1
    Loop
    
    NumberToWords = Trim(Result) & " Dollar" & IIf(Val(Result) > 1, "s", "") & CentsPart
End Function

Private Function ChunkToWords(ByVal Chunk)
    Dim Units As Variant, Tens As Variant, Teens As Variant
    Units = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
    Tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
    Teens = Array("Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    
    Dim Words As String
    Words = ""
    
    If Len(Chunk) = 3 Then
        If Val(Left(Chunk, 1)) > 0 Then
            Words = Units(Val(Left(Chunk, 1))) & " Hundred "
        End If
        Chunk = Mid(Chunk, 2)
    End If
    
    If Val(Chunk) >= 10 And Val(Chunk) <= 19 Then
        Words = Words & Teens(Val(Right(Chunk, 1)))
    Else
        If Val(Left(Chunk, 1)) > 0 Then
            Words = Words & Tens(Val(Left(Chunk, 1))) & " "
        End If
        If Val(Right(Chunk, 1)) > 0 Then
            Words = Words & Units(Val(Right(Chunk, 1)))
        End If
    End If
    
    ChunkToWords = Trim(Words)
End Function
