Imports System.IO
Imports Microsoft.Win32

Class MainWindow
    Sub endProgramMenuItem_onClick() Handles endProgramMenuItem.Click
        End
    End Sub
    Sub aboutMenu_onClick() Handles aboutMenue.Click
        MsgBox("PDF-Merkmale extrahieren" + Chr(10) + "entwickelt: Marc Bender, 2019", MsgBoxStyle.Information, "Über")
    End Sub
    Sub chooseFileMenuItem_onClick() Handles chooseFilesMenuItem.Click
        Dim chooseFilesDialog = New OpenFileDialog
        chooseFilesDialog.Title = "Datei(-en) wählen"
        chooseFilesDialog.Multiselect = True
        chooseFilesDialog.ShowDialog()
        Dim numOfFilesSelected As Integer = chooseFilesDialog.FileNames.Length
        If numOfFilesSelected > 0 Then
            statusMessages.AppendText(Chr(13) + "Ausgewählt " + numOfFilesSelected.ToString + " Dateien")
            For i = 0 To numOfFilesSelected - 1
                Dim oldFilename = chooseFilesDialog.FileNames(i)
                Dim newFilename = oldFilename.Replace(" ", "")
                Rename(oldFilename, newFilename)
                Threading.Thread.Sleep(100)
                Dim pdftotextProcess = Process.Start("pdftotext", newFilename)
                pdftotextProcess.WaitForExit()
                Dim textFilename = newFilename.Replace(".pdf", ".txt")
                Dim myfile = File.OpenRead(textFilename)
                Dim invoicenumber As String
                Dim amazonordernumber As String
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
                MsgBox("rechnungsnummer " + invoicenumber.ToString)
                MsgBox("ordernummer " + amazonordernumber.ToString)
                Rename(newFilename, oldFilename)
                myfile.Close()
                My.Computer.FileSystem.DeleteFile(textFilename)
            Next
        End If
    End Sub
End Class
