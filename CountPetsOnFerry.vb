' The code below counts total number of pets(cats and dogs) travelled in the ferry and shows the output in Message Box 
Public Class CountPetsOnFerry
    Public Shared Sub Main()
        Const ForReading = 1
        Const InputDir = "C:\Temp\"
        Const InputFile = "PetsOnFerry.txt"

        Dim retstring

        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim theFile = fso.OpenTextFile(InputDir & InputFile, ForReading, False)
        Dim nCats = 0
        Dim nDogs = 0
        'Dim nOthers = 0

        Do While theFile.AtEndOfStream <> True
            Dim Instring As Object
            Instring = theFile.ReadLine
            'msgbox UCASE(Left(Instring,3))
            Select Case UCase(Left(Instring, 3))

                Case "CAT"
                    nCats = nCats + 1
                Case "DOG"
                    nDogs = nDogs + 1
                    'Case Else
                    ' nOthers = nOthers + 1
            End Select
        Loop
        theFile.Close
        Dim ReadEntireFile As Object

        ReadEntireFile = retstring

        MsgBox(
        "Total Dogs traveled on ferry=" & nDogs & vbCrLf &
        "Total Cats traveled on ferry = " & nCats & vbCrLf)

        MsgBox("Total Pets traveled on ferry=" & nCats + nDogs)
        '+ nOthers)

    End Sub
End Class
