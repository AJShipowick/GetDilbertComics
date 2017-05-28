Imports System.IO
Imports System.Net

'http://www.dotnetperls.com/console-write-vbnet

Module GetDilbert

    Sub Main()
        GetComicDetails()
        CreateSaveLocation()
        GetDilbert("1990", 33039)
    End Sub

    Property comicStrip As String = "Dilbert"
    Property saveLocation As String
    Private Sub GetComicDetails()

        saveLocation = "D:\Comics\" & comicStrip

        Console.WriteLine("Save location will be here:" & saveLocation)
        Console.WriteLine("Go....")

    End Sub

    Private Sub CreateSaveLocation()

        If Directory.Exists(saveLocation) Then
            Exit Sub
        Else
            Directory.CreateDirectory(saveLocation)
        End If

    End Sub

    Public client As New WebClient
    Property goodPicCount As Integer = 0
    Property badPicCount As Integer = 0
    Property weekdayPic As String = ".strip.gif"
    Property newComicFile As String

    Private Sub GetDilbert(year As String, uniqueStartingNum As Integer)

        For i As Integer = 0 To 99  '364 = a year!

            newComicFile = String.Format("{0}\Dilbert_Year{1}_{2}.jpg", saveLocation, year, i)

            Try

                Dim dilbertURL As String = String.Format("http://dilbert.com/dyn/str_strip/000000000/00000000/0000000/000000/30000/3000/000/{0}/{0}{1}", uniqueStartingNum, weekdayPic)

                client.DownloadFile(dilbertURL, newComicFile)

                If ValidComicSize(newComicFile) Then 'Good picture
                    uniqueStartingNum += 1
                    Continue For
                Else

                    If SundayComic(dilbertURL, uniqueStartingNum, "000") Then
                        uniqueStartingNum += 1
                        Continue For
                    End If

                    If HundredDigitPasses(dilbertURL, uniqueStartingNum) Then
                        uniqueStartingNum += 1
                        Continue For
                    End If

                End If

                'Console.WriteLine("****YOU WIN****")
            Catch ex As Exception
                'Console.WriteLine("Failed")
                badPicCount += 1
            End Try

            uniqueStartingNum += 1

        Next

        Console.WriteLine("Good Picture Count = " & goodPicCount)
        Console.WriteLine("Bad Picture Count = " & badPicCount)
        Console.ReadLine()

    End Sub

    Property sundayPic As String = ".strip.sunday.gif"
    Private Function SundayComic(dilbertURL As String, uniqueStartingNum As Integer, three As String) As Boolean

        dilbertURL = String.Format("http://dilbert.com/dyn/str_strip/000000000/00000000/0000000/000000/30000/3000/{2}/{0}/{0}{1}", uniqueStartingNum, sundayPic, three)

        client.DownloadFile(dilbertURL, newComicFile)
        If ValidComicSize(newComicFile) Then Return True Else Return False

    End Function

    Private Function HundredDigitPasses(dilbertURL As String, uniqueStartingNum As Integer) As Boolean

        Dim threeDigitNum As Integer = 100

        For i As Integer = 0 To 9
            dilbertURL = String.Format("http://dilbert.com/dyn/str_strip/000000000/00000000/0000000/000000/30000/3000/{2}/{0}/{0}{1}", uniqueStartingNum, weekdayPic, threeDigitNum)

            client.DownloadFile(dilbertURL, newComicFile)
            If ValidComicSize(newComicFile) Then Return True

            If SundayComic(dilbertURL, uniqueStartingNum, threeDigitNum) Then Return True

            threeDigitNum += 100
        Next

        Return False
    End Function

    Function ValidComicSize(newComicFile As String) As Boolean
        Dim info As New FileInfo(newComicFile)

        If info.Length > 10000 Then 'Good picture
            goodPicCount += 1
            Return True
        Else
            IO.File.Delete(newComicFile)
            badPicCount += 1
            Return False
        End If

    End Function

End Module
