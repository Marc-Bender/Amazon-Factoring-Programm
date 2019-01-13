Imports System.IO
Imports Microsoft.Win32

Class MainWindow
    Dim outputFileName As String
    Dim filenamesChoosen As String()
    Sub endProgramMenuItem_onClick() Handles endProgramMenuItem.Click
        End
    End Sub
    Sub aboutMenu_onClick() Handles aboutMenue.Click
        MsgBox("PDF-Merkmale extrahieren" + Chr(10) + "entwickelt: Marc Bender, 2019", MsgBoxStyle.Information, "Über")
    End Sub
    Sub chooseOutputFile_onClick() Handles chooseOutputFile.Click
        Dim chooseOutputFileDialog = New OpenFileDialog
        chooseOutputFileDialog.Title = "Ausgabedatei wählen"
        chooseOutputFileDialog.Multiselect = False
        chooseOutputFileDialog.ShowDialog()
        outputFileName = chooseOutputFileDialog.FileName
        statusMessages.AppendText(Chr(13) + "Ausgabe nach: " + outputFileName)
    End Sub
    Sub chooseFileMenuItem_onClick() Handles chooseFilesMenuItem.Click
        Dim chooseFilesDialog = New OpenFileDialog
        chooseFilesDialog.Title = "Datei(-en) wählen"
        chooseFilesDialog.Multiselect = True
        chooseFilesDialog.ShowDialog()
        filenamesChoosen = chooseFilesDialog.FileNames
        If filenamesChoosen.Length > 0 Then
            statusMessages.AppendText(Chr(13) + "Ausgewählt " + filenamesChoosen.Length.ToString + " Dateien")
        End If
    End Sub
    Sub go_onClick() Handles go.Click
        findCharacteristicsInFile(filenamesChoosen, filenamesChoosen.Length)
    End Sub
    Private Sub findCharacteristicsInFile(filenamesChoosen As String(), numOfFilesSelected As Integer)
        Dim amazonOrdernumberWasZeroOnceOrMoreOften As Boolean
        For i = 0 To numOfFilesSelected - 1
            Dim oldFilename = filenamesChoosen(i)
            Dim newFilename = oldFilename.Replace(" ", "")
            Rename(oldFilename, newFilename)
            Threading.Thread.Sleep(100)
            Dim pdftotextProcess = Process.Start("pdftotext", newFilename)
            pdftotextProcess.WaitForExit()
            Dim textFilename = newFilename.Replace(".pdf", ".txt")
            Dim myfile = File.OpenRead(textFilename)
            Dim invoicenumber = ""
            Dim amazonordernumber = ""
            Dim myline As String
            For j = 0 To myfile.Length - 1
                Dim mybyte = Chr(myfile.ReadByte).ToString
                If Not (mybyte.Equals(vbCrLf)) AndAlso Not (mybyte.Equals(vbCr)) AndAlso Not (mybyte.Equals(vbLf)) Then
                    myline += mybyte
                Else
                    If myline.Contains("Rechnungsnr. ") Then
                        invoicenumber = myline.Remove(0, "Rechnungsnr. ".Length)
                        Dim sampleinvoicenumber = "184183"
                        invoicenumber = invoicenumber.Remove(sampleinvoicenumber.Length, invoicenumber.Length - sampleinvoicenumber.Length)
                    End If
                    If myline.Contains("Amazon Order ") Then
                        amazonordernumber = myline.Remove(0, "Amazon Order ".Length)
                        Dim sampleamazonordernumber = "302-6384101-1861149"
                        amazonordernumber = amazonordernumber.Remove(sampleamazonordernumber.Length, amazonordernumber.Length - sampleamazonordernumber.Length)
                    End If
                    myline = ""
                End If
            Next
            If i = 0 Then
                My.Computer.FileSystem.WriteAllText(outputFileName, "Rechnungsnummer" + ";" + "Amazon-Order-Nummer" + vbCrLf, True)
            End If

            If IsNothing(invoicenumber) Then
                MsgBox("Error: Keine Rechnungsnummer", MsgBoxStyle.Exclamation, "Error")
            ElseIf IsNothing(amazonordernumber) Then
                amazonOrdernumberWasZeroOnceOrMoreOften = True
                amazonordernumber = 0
            End If
            My.Computer.FileSystem.WriteAllText(outputFileName, invoicenumber.ToString + ";" + amazonordernumber.ToString + vbCrLf, True)
            Rename(newFilename, oldFilename)
            myfile.Close()
            My.Computer.FileSystem.DeleteFile(textFilename)
        Next
        statusMessages.AppendText(Chr(13) + "Fertig mit der Verarbeitung von " + numOfFilesSelected.ToString + " Dateien")
    End Sub
End Class
