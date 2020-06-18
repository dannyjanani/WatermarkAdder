<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnDrawing = New System.Windows.Forms.Button()
        Me.tbFolder = New System.Windows.Forms.TextBox()
        Me.tbDrawing = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnFolder = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btnDrawing
        '
        Me.btnDrawing.Location = New System.Drawing.Point(335, 100)
        Me.btnDrawing.Name = "btnDrawing"
        Me.btnDrawing.Size = New System.Drawing.Size(111, 23)
        Me.btnDrawing.TabIndex = 1
        Me.btnDrawing.Text = "Browse Drawing..."
        Me.btnDrawing.UseVisualStyleBackColor = True
        '
        'tbFolder
        '
        Me.tbFolder.Location = New System.Drawing.Point(35, 48)
        Me.tbFolder.Name = "tbFolder"
        Me.tbFolder.Size = New System.Drawing.Size(281, 20)
        Me.tbFolder.TabIndex = 2
        '
        'tbDrawing
        '
        Me.tbDrawing.Location = New System.Drawing.Point(35, 100)
        Me.tbDrawing.Name = "tbDrawing"
        Me.tbDrawing.Size = New System.Drawing.Size(281, 20)
        Me.tbDrawing.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(32, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(263, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Select the folder you would like to add a watermark to:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(32, 84)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(270, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Select the file containing the watermark you wish to add"
        '
        'btnFolder
        '
        Me.btnFolder.Location = New System.Drawing.Point(335, 48)
        Me.btnFolder.Name = "btnFolder"
        Me.btnFolder.Size = New System.Drawing.Size(111, 23)
        Me.btnFolder.TabIndex = 6
        Me.btnFolder.Text = "Browse Folder..."
        Me.btnFolder.UseVisualStyleBackColor = True
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(92, 132)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(134, 50)
        Me.btnRun.TabIndex = 7
        Me.btnRun.Text = "Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(246, 132)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(134, 50)
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(467, 194)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.btnFolder)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbDrawing)
        Me.Controls.Add(Me.tbFolder)
        Me.Controls.Add(Me.btnDrawing)
        Me.Name = "Main"
        Me.Text = "Add Watermark to multiple drawings"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnDrawing As System.Windows.Forms.Button
    Friend WithEvents tbFolder As System.Windows.Forms.TextBox
    Friend WithEvents tbDrawing As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnFolder As System.Windows.Forms.Button
    Friend WithEvents btnRun As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button

End Class
