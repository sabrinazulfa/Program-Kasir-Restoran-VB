<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btnAboutUs = New System.Windows.Forms.Button()
        Me.btnMainMenu = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(0, -2)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(515, 454)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'btnAboutUs
        '
        Me.btnAboutUs.Location = New System.Drawing.Point(37, 33)
        Me.btnAboutUs.Name = "btnAboutUs"
        Me.btnAboutUs.Size = New System.Drawing.Size(94, 29)
        Me.btnAboutUs.TabIndex = 1
        Me.btnAboutUs.Text = "About Us"
        Me.btnAboutUs.UseVisualStyleBackColor = True
        '
        'btnMainMenu
        '
        Me.btnMainMenu.Location = New System.Drawing.Point(196, 33)
        Me.btnMainMenu.Name = "btnMainMenu"
        Me.btnMainMenu.Size = New System.Drawing.Size(94, 29)
        Me.btnMainMenu.TabIndex = 2
        Me.btnMainMenu.Text = "Main Menu"
        Me.btnMainMenu.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(349, 33)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(94, 29)
        Me.btnExit.TabIndex = 3
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(500, 450)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnMainMenu)
        Me.Controls.Add(Me.btnAboutUs)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents btnAboutUs As Button
    Friend WithEvents btnMainMenu As Button
    Friend WithEvents btnExit As Button
End Class
