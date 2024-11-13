
# Excel Number to Words Converter Add-In

This project provides an Excel VBA function to convert numerical values into English words, including handling of decimal values. For example, `10,548` becomes "Ten Thousand Five Hundred Forty-Eight Dollars."

## Features

- Converts numbers into words in English.
- Handles values up to billions.
- Supports decimal values by converting digits after the decimal point individually.
- Appends "Dollar" or "Dollars" to the output based on the number provided.

## Getting Started

This guide will help you set up the VBA function as an Excel Add-In, so it is available in all workbooks on your PC.

### Prerequisites

- **Microsoft Excel** with VBA support (Excel for Windows or Mac).

## Installation

### 1. Create an Excel Add-In

1. Open a new Excel workbook.
2. **Press `ALT + F11`** to open the VBA editor.
3. **Go to `Insert > Module`** to add a new module.
4. **Paste the VBA code** (provided below) for `NumberToWords` and `ChunkToWords` into this module.
5. **Save the workbook** as an Excel Add-In:
   - Go to **File > Save As**.
   - Choose **Excel Add-In (.xlam)** as the file type.
   - Name it something like `NumberToWordsConverter` and save it in the default location (`C:\Users\%username%\AppData\Roaming\Microsoft\AddIns`).

### 2. Load the Add-In in Excel

1. In Excel, go to **File > Options**.
2. In the **Excel Options** window, select **Add-Ins** on the left.
3. At the bottom, where it says **Manage**, select **Excel Add-ins** from the dropdown and click **Go...**.
4. In the **Add-Ins** dialog, click **Browse…** and locate the `.xlam` file you saved.
5. Select the add-in and click **OK**. Ensure the checkbox next to your add-in name is checked.

### 3. Use the Function in Any Workbook

After loading the add-in, the `NumberToWords` function is available in any Excel workbook on your PC. You can use it as follows:

```excel
=NumberToWords(A1)
```

### 4. (Optional) Make the Add-In Available Automatically on Startup

If you want this add-in to load automatically every time you open Excel, it should already be configured once you check the box in the **Add-Ins** dialog. Verify by reopening Excel to ensure the function is accessible.

## VBA Code

Copy the following VBA code into the module when creating the Add-In:

```vba
Function NumberToWords(ByVal MyNumber)
    Dim Units As Variant, Tens As Variant
    Dim Temp As String, DecimalPlace As Integer, Count As Integer
    Dim DecimalWords As String
    
    Units = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
    Tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
    
    MyNumber = Trim(CStr(MyNumber))
    DecimalPlace = InStr(MyNumber, ".")
    If DecimalPlace > 0 Then
        DecimalWords = " Point"
        For Count = DecimalPlace + 1 To Len(MyNumber)
            DecimalWords = DecimalWords & " " & Units(Val(Mid(MyNumber, Count, 1)))
        Next Count
        MyNumber = Left(MyNumber, DecimalPlace - 1)
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
    
    NumberToWords = Trim(Result) & " Dollar" & DecimalWords
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
```

## Example

| Number  | Text Conversion                       |
|---------|---------------------------------------|
| 10548   | Ten Thousand Five Hundred Forty-Eight Dollar |
| 123.45  | One Hundred Twenty-Three Dollar Point Four Five |

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! If you have suggestions or improvements, feel free to fork this repository and submit a pull request.
