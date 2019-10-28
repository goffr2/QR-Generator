Imports MessagingToolkit.QRCode.Codec.Data
Imports Word = Microsoft.Office.Interop.Word
Public Class Form1
    Private Sub Browse_Click(sender As Object, e As EventArgs) Handles Browse.Click

        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String
        fd.Title = "Find CSV file"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            TextBox1.Text = strFileName
        End If
        If TextBox1.Text = "" Then
            Form2.RichTextBox1.Text = "You must choose a file."
            Exit Sub
        End If


    End Sub

    Private Sub Generate_Click(sender As Object, e As EventArgs) Handles Generate.Click
        Dim strFileName As String
        Dim Type As String
        Dim Compare As String = "csv"
        strFileName = TextBox1.Text
        Type = Microsoft.VisualBasic.Right(strFileName, 3)
        If strFileName = "" Then
            Form2.RichTextBox1.Text = "You must choose a file."
            Exit Sub
        End If
        If Type <> "csv" Then
            MsgBox("File must be a CSV file")
            Exit Sub
        End If

        Dim path As String = Application.StartupPath & "\Test\"
        Dim Combined_path As String = Application.StartupPath & "\Combined\"

        If My.Computer.FileSystem.DirectoryExists(path) = False Then
            My.Computer.FileSystem.CreateDirectory(path)
        End If

        If My.Computer.FileSystem.DirectoryExists(Combined_path) = False Then
            My.Computer.FileSystem.CreateDirectory(Combined_path)
        End If



        Dim fileReader As System.IO.StreamReader
        fileReader =
        My.Computer.FileSystem.OpenTextFileReader(strFileName)
        Dim stringReader As String
        Dim Place As Integer = 0
        While fileReader.EndOfStream = False
            Place = Place + 1
            stringReader = fileReader.ReadLine()
            Generate_Description(stringReader, Place, path)
            Generate_QR_Image(stringReader, Place, path)
            Combine_Images(path, Combined_path, Place)
        End While
        Generate_Word_Document(Place, Combined_path)
        System.IO.Directory.Delete(path, True)
        System.IO.Directory.Delete(Combined_path, True)


    End Sub

    Private Sub Generate_QR_Image(ByRef stringReader As String, ByRef Place As Integer, ByRef path As String)

        Dim enc As New MessagingToolkit.QRCode.Codec.QRCodeEncoder()
        Dim bm As Bitmap = enc.Encode(stringReader)
        bm.Save(path & "QRCode" & Place & ".png", System.Drawing.Imaging.ImageFormat.Png)
        bm.Dispose()

    End Sub

    Private Sub Generate_Description(ByRef stringReader As String, ByRef Place As Integer, ByRef path As String)

        Dim Description As String = ""
        Description = stringReader.Split(","c)(0)
        If Description = "" Then


            Form2.RichTextBox1.Text = "I find your lack of Descriptions Disturbing. You are missing a Description in the" & Place & " th line. Check it and start over."
            Form2.PictureBox1.Image = My.Resources.vader
            Form2.ShowDialog()
            System.IO.Directory.Delete(path, True)
            TextBox1.Text = ""
            Application.Exit()
            End
        End If
        Generate_Description_Image(Description, Place, path)
    End Sub

    Private Sub Generate_Description_Image(ByRef Description As String, ByRef Place As Integer, ByRef path As String)


        Dim FontColor As Color = Color.Black
        Dim BackColor As Color = Color.White
        Dim FontName As String = "Times New Roman"
        Dim FontSize As Integer = 14
        Dim Height As Integer = 40
        Dim Width As Integer = 181
        Dim objBitmap As New Bitmap(Width, Height)
        Dim objGraphics As Graphics = Graphics.FromImage(objBitmap)
        Dim objColor As Color
        Dim objFont As New Font(FontName, FontSize)

        Dim objPoint As New PointF(5.0F, 5.0F)
        Dim objBrushForeColor As New SolidBrush(FontColor)
        Dim objBrushBackColor As New SolidBrush(BackColor)

        objGraphics.FillRectangle(objBrushBackColor, 0, 0, Width, Height)
        objGraphics.DrawString(Description, objFont, objBrushForeColor, objPoint)
        objBitmap.Save(path & "Description" & Place & ".png", Imaging.ImageFormat.Png)
        objBitmap.Dispose()

    End Sub

    Private Sub Combine_Images(ByRef path As String, ByRef Combined_path As String, ByRef Place As Integer)
        Dim Img2 As New System.Drawing.Bitmap(path & "QRCode" & Place & ".png")
        Dim Img1 As New System.Drawing.Bitmap(path & "Description" & Place & ".png")

        Dim bmp As New Bitmap(Math.Max(Img1.Width, Img2.Width), Img1.Height + Img2.Height)
        Dim g As Graphics = Graphics.FromImage(bmp)

        g.DrawImage(Img1, 0, 0, Img1.Width, Img1.Height)
        g.DrawImage(Img2, 0, Img1.Height, Img2.Width, Img2.Width)
        bmp.Save(Combined_path & Place & ".png")
        '  My.Computer.FileSystem.DeleteFile(path & "QRCode" & Place & ".png")
        ' My.Computer.FileSystem.DeleteFile(path & "Description" & Place & ".png")
        g.Dispose()
        Img1.Dispose()
        Img2.Dispose()

    End Sub
    Private Sub Generate_Word_Document(ByRef Total As Integer, ByRef Combined_path As String)
        Dim template As String = Application.StartupPath
        Dim oWord As New Word.Application
        Dim oDoc As New Word.Document
        Dim oTable As Word.Table

        'Start Word and open the document template.
        oWord = CreateObject("Word.Application")
        oWord.Visible = False
        oDoc = oWord.Documents.Add

        oDoc.PageSetup.TopMargin = oWord.InchesToPoints(0.63)
        oDoc.PageSetup.BottomMargin = oWord.InchesToPoints(0.57)
        oDoc.PageSetup.LeftMargin = oWord.InchesToPoints(0.71)
        oDoc.PageSetup.RightMargin = oWord.InchesToPoints(0.31)

        Dim r As Integer = 1
        Dim c As Integer = 1
        Dim J As Integer = 1
        Dim P As Integer = 1

        Dim Num_Of_Pages As Decimal = Total / 12

        Num_Of_Pages = Math.Ceiling(Num_Of_Pages)

        Dim Image_Rows As Decimal = Total / 3
        Image_Rows = Math.Ceiling(Image_Rows)

        Dim rows As Decimal = Image_Rows + (Num_Of_Pages * 3)
        Math.Ceiling(rows)


        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, rows + 1, 5)
        oTable.Borders.Enable = False
        While P < Total

            While c <= 6

                oTable.Cell(r, c).Range.InlineShapes.AddPicture(Application.StartupPath & "\Combined\" & P & ".png")
                oTable.Cell(r, c).Height = oWord.InchesToPoints(2)
                oTable.Cell(r, c).Width = oWord.InchesToPoints(2)
                If P = Total Then
                    Exit While
                End If
                J = J + 1
                P = P + 1
                c = c + 1



                oTable.Cell(r, c).Height = oWord.InchesToPoints(2)
                oTable.Cell(r, c).Width = oWord.InchesToPoints(0.63)


                c = c + 1

                oTable.Cell(r, c).Range.InlineShapes.AddPicture(Application.StartupPath & "\Combined\" & P & ".png")
                oTable.Cell(r, c).Height = oWord.InchesToPoints(2)
                oTable.Cell(r, c).Width = oWord.InchesToPoints(2)
                If P = Total Then
                    Exit While
                End If
                J = J + 1
                P = P + 1
                c = c + 1


                oTable.Cell(r, c).Height = oWord.InchesToPoints(2)
                oTable.Cell(r, c).Width = oWord.InchesToPoints(0.63)


                c = c + 1

                oTable.Cell(r, c).Range.InlineShapes.AddPicture(Application.StartupPath & "\Combined\" & P & ".png")
                oTable.Cell(r, c).Height = oWord.InchesToPoints(2)
                oTable.Cell(r, c).Width = oWord.InchesToPoints(2)
                If P = Total Then
                    Exit While
                End If
                P = P + 1
                J = J + 1
                c = c + 1
                r = r + 1


                If c = 6 And J < 12 Then

                    c = 1


                    oTable.Cell(r, c).Height = oWord.InchesToPoints(0.58)
                    oTable.Cell(r, c).Width = oWord.InchesToPoints(2)


                    oTable.Cell(r, c + 1).Height = oWord.InchesToPoints(0.58)
                    oTable.Cell(r, c + 1).Width = oWord.InchesToPoints(2)

                    oTable.Cell(r, c + 2).Height = oWord.InchesToPoints(0.58)
                    oTable.Cell(r, c + 2).Width = oWord.InchesToPoints(2)

                    oTable.Cell(r, c + 3).Height = oWord.InchesToPoints(0.58)
                    oTable.Cell(r, c + 3).Width = oWord.InchesToPoints(2)

                    oTable.Cell(r, c + 4).Height = oWord.InchesToPoints(0.58)
                    oTable.Cell(r, c + 4).Width = oWord.InchesToPoints(2)
                    r = r + 1
                    c = 1

                End If

                If J >= 12 Then
                    J = 1
                    c = 1

                    oTable.Cell(r, c).Range.InlineShapes.AddPicture(Application.StartupPath & "\Combined\" & P & ".png")
                    oTable.Cell(r, c).Height = oWord.InchesToPoints(2)
                    oTable.Cell(r, c).Width = oWord.InchesToPoints(2)
                    If P = Total Then
                        Exit While
                    End If
                    J = J + 1
                    P = P + 1
                    c = c + 1



                    oTable.Cell(r, c).Height = oWord.InchesToPoints(2)
                    oTable.Cell(r, c).Width = oWord.InchesToPoints(0.63)


                    c = c + 1

                    oTable.Cell(r, c).Range.InlineShapes.AddPicture(Application.StartupPath & "\Combined\" & P & ".png")
                    oTable.Cell(r, c).Height = oWord.InchesToPoints(2)
                    oTable.Cell(r, c).Width = oWord.InchesToPoints(2)
                    If P = Total Then
                        Exit While
                    End If
                    J = J + 1
                    P = P + 1
                    c = c + 1


                    oTable.Cell(r, c).Height = oWord.InchesToPoints(2)
                    oTable.Cell(r, c).Width = oWord.InchesToPoints(0.63)


                    c = c + 1

                    oTable.Cell(r, c).Range.InlineShapes.AddPicture(Application.StartupPath & "\Combined\" & P & ".png")
                    oTable.Cell(r, c).Height = oWord.InchesToPoints(2)
                    oTable.Cell(r, c).Width = oWord.InchesToPoints(2)
                    If P = Total Then
                        Exit While
                    End If
                    P = P + 1
                    J = J + 1
                    c = c + 1
                    r = r + 1



                    c = 1

                    oTable.Cell(r, c).Height = oWord.InchesToPoints(0.58)
                    oTable.Cell(r, c).Width = oWord.InchesToPoints(2)


                    oTable.Cell(r, c + 1).Height = oWord.InchesToPoints(0.58)
                    oTable.Cell(r, c + 1).Width = oWord.InchesToPoints(2)

                    oTable.Cell(r, c + 2).Height = oWord.InchesToPoints(0.58)
                    oTable.Cell(r, c + 2).Width = oWord.InchesToPoints(2)

                    oTable.Cell(r, c + 3).Height = oWord.InchesToPoints(0.58)
                    oTable.Cell(r, c + 3).Width = oWord.InchesToPoints(2)

                    oTable.Cell(r, c + 4).Height = oWord.InchesToPoints(0.58)
                    oTable.Cell(r, c + 4).Width = oWord.InchesToPoints(2)
                    r = r + 1
                    c = 1
                    Exit While
                End If



            End While

        End While
        oWord.Visible = False
        Dim fd As FolderBrowserDialog = New FolderBrowserDialog()
        Dim SaveFileName As String
        fd.Description = "Select a folder to save the file in."
        fd.ShowNewFolderButton = True

        If fd.ShowDialog = DialogResult.OK Then
            SaveFileName = fd.SelectedPath & "\"
            oDoc.SaveAs2(SaveFileName & "QRCODES.doc")
        End If


        MsgBox("done")

        'All done. Close this form.
        oDoc.Close()
        oWord.Application.Quit()
        oWord = Nothing
        Me.Close()
        Shell("Taskkill /IM WINWORD.exe /F")

    End Sub
End Class
