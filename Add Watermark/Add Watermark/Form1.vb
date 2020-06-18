Imports System.IO
Imports System.Threading
Public Class Main
    ' Declaring constants such as the delimeter to be used to parse the data, the files to be searched for, and some others.
    Const delimeter As String = "[%]"
    Const paramPath As String = "params.out"
    Const waterMarkPath As String = "watermark.out"
    Const abortPath As String = "abort.ex"

    ' If a user clicks the Browse Folder button on the form, it will open a dialog to search for a folder you want to add the watermark to.
    ' It will then save the selection into tbFolder.Text which is a string.
    Private Sub btnFolder_Click(sender As Object, e As EventArgs) Handles btnFolder.Click
        Dim foldersearch As New FolderBrowserDialog
        If foldersearch.ShowDialog = Windows.Forms.DialogResult.OK Then
            tbFolder.Text = foldersearch.SelectedPath
        End If
    End Sub
    ' If a user clicks the Browse Drawing button on the form, it will open a dialog to search for a drawing containing the watermark.
    ' It will then save the selection into tbDrawing.Text which is a string.
    Private Sub btnDrawing_Click(sender As Object, e As EventArgs) Handles btnDrawing.Click
        Dim filesearch As New OpenFileDialog
        filesearch.Filter = "AutoCAD Drawing (.dwg)|*.dwg"
        'tbDrawing.Text = "C:\temp\watermark.dwg"
        If filesearch.ShowDialog = Windows.Forms.DialogResult.OK Then
            tbDrawing.Text = filesearch.FileName
        End If
    End Sub
    ' If the text changed and the box is yellow, return it to the original dialog color
    ' The box is usually turned to yellow when the box is empty and Run is pressed as well as if the folder is not found.
    Private Sub tbFolder_TextChanged(sender As Object, e As EventArgs) Handles tbFolder.TextChanged
        If tbFolder.BackColor = Color.Yellow Then
            tbFolder.BackColor = SystemColors.Window
        End If
    End Sub
    ' If the text changed and the box is yellow, return it to the original dialog color
    ' The box is usually turned to yellow when the box is empty and Run is pressed as well as if the drawing is not found.
    Private Sub tbDrawing_TextChanged(sender As Object, e As EventArgs) Handles tbDrawing.TextChanged
        If tbDrawing.BackColor = Color.Yellow Then
            tbDrawing.BackColor = SystemColors.Window
        End If
    End Sub
    ' If a user clicks the Run button on the form, it will first delete an abort file in case it was existing from before.
    ' The abort file is generated when the cancel button is pressed and the VBA embedded drawing will stop running if it finds it in the directory.
    ' It then checks that all conditions are met (the form isn't missing information) and proceeds based on the checkbox the user selected.
    ' It then gets all the files in the selected directory, and if the drawing has a DWG extension, it would create a file called params.out 
    ' containing all the file names that we wish to convert. Another file, watermark.out keeps the path to the watermark file.
    ' Finally, it opens the drawing with the embedded VBA code.
    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        DeleteFile(abortPath)
        Dim proceed As Boolean : proceed = checkConditions()
        If proceed = True Then
            Dim fileListing() As String : fileListing = Directory.GetFiles(tbFolder.Text)
            Dim paramOut As String : paramOut = ""

            For Each fileString In fileListing
                If LCase(fileString).EndsWith(".dwg") Then
                    If paramOut <> "" Then
                        paramOut = paramOut & vbNewLine
                    End If
                    paramOut = paramOut & fileString
                End If
            Next
            writeToFile(paramPath, paramOut, False)
            writeToFile(waterMarkPath, tbDrawing.Text, False)
            Process.Start("WatermarkAdder.dwg")
        End If
    End Sub
    ' Checks to see if there are errors and displays a textbox and highlights all errors.
    ' Errors checked: file/folder does not exist, text box is empty, pages not selected. It then returns a boolean called proceed
    Public Function checkConditions() As Boolean
        Dim proceed As Boolean : proceed = True
        Dim errorMessage As String = "Please correct the errors listed below, then try again." & vbNewLine

        If tbDrawing.Text <> "" And Not File.Exists(tbDrawing.Text) Then
            errorMessage += "The specified file does not exist" & vbNewLine
            tbDrawing.BackColor = Color.Yellow
            proceed = False
        End If
        If tbDrawing.Text = "" Then
            errorMessage += "Please enter a valid file" & vbNewLine
            tbDrawing.BackColor = Color.Yellow
            proceed = False
        End If
        If tbFolder.Text <> "" And Not Directory.Exists(tbFolder.Text) Then
            errorMessage += "The specified directory does not exist" & vbNewLine
            tbFolder.BackColor = Color.Yellow
            proceed = False
        End If
        If tbFolder.Text = "" Then
            errorMessage += "Please enter a valid directory" & vbNewLine
            tbFolder.BackColor = Color.Yellow
            proceed = False
        End If
        If proceed = False Then
            MessageBox.Show(errorMessage)
        End If
        Return proceed
    End Function
    ' Creates a file called abort.ex to tell the VBA program, I want to cancel execution.
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        writeToFile(abortPath, "", False)
    End Sub
End Class