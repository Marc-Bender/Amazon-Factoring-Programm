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
        For i = 0 To chooseFilesDialog.FileNames.Length - 1
            Process.Start("pdftotext", chooseFilesDialog.FileNames(i))
        Next
    End Sub
End Class
