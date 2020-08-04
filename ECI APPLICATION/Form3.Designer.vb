<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
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
        Me.angka = New System.Windows.Forms.TextBox()
        Me.tambah = New System.Windows.Forms.Button()
        Me.kurang = New System.Windows.Forms.Button()
        Me.kali = New System.Windows.Forms.Button()
        Me.hasil = New System.Windows.Forms.TextBox()
        Me.bagi = New System.Windows.Forms.Button()
        Me.samadengan = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'angka
        '
        Me.angka.Location = New System.Drawing.Point(9, 30)
        Me.angka.Multiline = True
        Me.angka.Name = "angka"
        Me.angka.Size = New System.Drawing.Size(386, 91)
        Me.angka.TabIndex = 0
        '
        'tambah
        '
        Me.tambah.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tambah.Location = New System.Drawing.Point(9, 136)
        Me.tambah.Name = "tambah"
        Me.tambah.Size = New System.Drawing.Size(190, 138)
        Me.tambah.TabIndex = 1
        Me.tambah.Text = "+"
        Me.tambah.UseVisualStyleBackColor = True
        '
        'kurang
        '
        Me.kurang.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.kurang.Location = New System.Drawing.Point(205, 136)
        Me.kurang.Name = "kurang"
        Me.kurang.Size = New System.Drawing.Size(208, 138)
        Me.kurang.TabIndex = 2
        Me.kurang.Text = "-"
        Me.kurang.UseVisualStyleBackColor = True
        '
        'kali
        '
        Me.kali.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.kali.Location = New System.Drawing.Point(9, 280)
        Me.kali.Name = "kali"
        Me.kali.Size = New System.Drawing.Size(190, 138)
        Me.kali.TabIndex = 3
        Me.kali.Text = "X"
        Me.kali.UseVisualStyleBackColor = True
        '
        'hasil
        '
        Me.hasil.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.hasil.Location = New System.Drawing.Point(401, 7)
        Me.hasil.Multiline = True
        Me.hasil.Name = "hasil"
        Me.hasil.Size = New System.Drawing.Size(224, 114)
        Me.hasil.TabIndex = 5
        '
        'bagi
        '
        Me.bagi.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bagi.Location = New System.Drawing.Point(205, 280)
        Me.bagi.Name = "bagi"
        Me.bagi.Size = New System.Drawing.Size(208, 138)
        Me.bagi.TabIndex = 6
        Me.bagi.Text = ":"
        Me.bagi.UseVisualStyleBackColor = True
        '
        'samadengan
        '
        Me.samadengan.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.samadengan.Location = New System.Drawing.Point(419, 136)
        Me.samadengan.Name = "samadengan"
        Me.samadengan.Size = New System.Drawing.Size(208, 282)
        Me.samadengan.TabIndex = 7
        Me.samadengan.Text = "="
        Me.samadengan.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(5, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 24)
        Me.Label1.TabIndex = 8
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(636, 427)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.samadengan)
        Me.Controls.Add(Me.bagi)
        Me.Controls.Add(Me.hasil)
        Me.Controls.Add(Me.kali)
        Me.Controls.Add(Me.kurang)
        Me.Controls.Add(Me.tambah)
        Me.Controls.Add(Me.angka)
        Me.Name = "Form3"
        Me.Text = "CALCULATOR"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents angka As TextBox
    Friend WithEvents tambah As Button
    Friend WithEvents kurang As Button
    Friend WithEvents kali As Button
    Friend WithEvents hasil As TextBox
    Friend WithEvents bagi As Button
    Friend WithEvents samadengan As Button
    Friend WithEvents Label1 As Label
End Class
