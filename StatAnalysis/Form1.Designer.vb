<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.btnCalcRunCharts = New System.Windows.Forms.Button()
        Me.btnGetModelNo = New System.Windows.Forms.Button()
        Me.btnGetImpellers = New System.Windows.Forms.Button()
        Me.cmbSN = New System.Windows.Forms.ComboBox()
        Me.cmbImpellers = New System.Windows.Forms.ComboBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtDesignFlow = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.lblDesignFlow = New System.Windows.Forms.Label()
        Me.txtDesignHead = New System.Windows.Forms.TextBox()
        Me.lblDesignHead = New System.Windows.Forms.Label()
        Me.lbSN = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'btnCalcRunCharts
        '
        Me.btnCalcRunCharts.Enabled = False
        Me.btnCalcRunCharts.Location = New System.Drawing.Point(566, 506)
        Me.btnCalcRunCharts.Name = "btnCalcRunCharts"
        Me.btnCalcRunCharts.Size = New System.Drawing.Size(138, 49)
        Me.btnCalcRunCharts.TabIndex = 0
        Me.btnCalcRunCharts.Text = "Calculate Run Charts"
        Me.btnCalcRunCharts.UseVisualStyleBackColor = True
        '
        'btnGetModelNo
        '
        Me.btnGetModelNo.Location = New System.Drawing.Point(12, 27)
        Me.btnGetModelNo.Name = "btnGetModelNo"
        Me.btnGetModelNo.Size = New System.Drawing.Size(108, 49)
        Me.btnGetModelNo.TabIndex = 1
        Me.btnGetModelNo.Text = "Get Model Numbers"
        Me.btnGetModelNo.UseVisualStyleBackColor = True
        Me.btnGetModelNo.Visible = False
        '
        'btnGetImpellers
        '
        Me.btnGetImpellers.Location = New System.Drawing.Point(12, 119)
        Me.btnGetImpellers.Name = "btnGetImpellers"
        Me.btnGetImpellers.Size = New System.Drawing.Size(108, 49)
        Me.btnGetImpellers.TabIndex = 2
        Me.btnGetImpellers.Text = "Get Impeller Sizes"
        Me.btnGetImpellers.UseVisualStyleBackColor = True
        Me.btnGetImpellers.Visible = False
        '
        'cmbSN
        '
        Me.cmbSN.FormattingEnabled = True
        Me.cmbSN.Location = New System.Drawing.Point(181, 52)
        Me.cmbSN.Name = "cmbSN"
        Me.cmbSN.Size = New System.Drawing.Size(422, 24)
        Me.cmbSN.TabIndex = 3
        Me.cmbSN.Visible = False
        '
        'cmbImpellers
        '
        Me.cmbImpellers.FormattingEnabled = True
        Me.cmbImpellers.Location = New System.Drawing.Point(411, 371)
        Me.cmbImpellers.Name = "cmbImpellers"
        Me.cmbImpellers.Size = New System.Drawing.Size(250, 24)
        Me.cmbImpellers.TabIndex = 4
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(35, 373)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(269, 22)
        Me.TextBox1.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(178, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(425, 17)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Model Number (Quantity in Polar Database) * is Supermaret Pump"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(374, 478)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(200, 22)
        Me.TextBox2.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(413, 325)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(291, 45)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Impeller Size for Selected Model Number (Quantity in Polar Database)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(111, 317)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 17)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Label3"
        Me.Label3.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(32, 353)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(159, 17)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Selected Model Number"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(371, 458)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(147, 17)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Selected Impeller Size"
        '
        'txtDesignFlow
        '
        Me.txtDesignFlow.Location = New System.Drawing.Point(35, 428)
        Me.txtDesignFlow.Name = "txtDesignFlow"
        Me.txtDesignFlow.Size = New System.Drawing.Size(200, 22)
        Me.txtDesignFlow.TabIndex = 12
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(32, 408)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(84, 17)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Design Flow"
        '
        'lblDesignFlow
        '
        Me.lblDesignFlow.AutoSize = True
        Me.lblDesignFlow.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDesignFlow.Location = New System.Drawing.Point(241, 428)
        Me.lblDesignFlow.Name = "lblDesignFlow"
        Me.lblDesignFlow.Size = New System.Drawing.Size(39, 17)
        Me.lblDesignFlow.TabIndex = 14
        Me.lblDesignFlow.Text = "asdf"
        Me.lblDesignFlow.Visible = False
        '
        'txtDesignHead
        '
        Me.txtDesignHead.Location = New System.Drawing.Point(35, 478)
        Me.txtDesignHead.Name = "txtDesignHead"
        Me.txtDesignHead.Size = New System.Drawing.Size(200, 22)
        Me.txtDesignHead.TabIndex = 15
        Me.txtDesignHead.Visible = False
        '
        'lblDesignHead
        '
        Me.lblDesignHead.AutoSize = True
        Me.lblDesignHead.Location = New System.Drawing.Point(32, 458)
        Me.lblDesignHead.Name = "lblDesignHead"
        Me.lblDesignHead.Size = New System.Drawing.Size(90, 17)
        Me.lblDesignHead.TabIndex = 16
        Me.lblDesignHead.Text = "Design Head"
        Me.lblDesignHead.Visible = False
        '
        'lbSN
        '
        Me.lbSN.FormattingEnabled = True
        Me.lbSN.ItemHeight = 16
        Me.lbSN.Location = New System.Drawing.Point(35, 27)
        Me.lbSN.Name = "lbSN"
        Me.lbSN.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lbSN.Size = New System.Drawing.Size(422, 260)
        Me.lbSN.TabIndex = 17
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(726, 563)
        Me.Controls.Add(Me.lbSN)
        Me.Controls.Add(Me.lblDesignHead)
        Me.Controls.Add(Me.txtDesignHead)
        Me.Controls.Add(Me.lblDesignFlow)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtDesignFlow)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.cmbImpellers)
        Me.Controls.Add(Me.cmbSN)
        Me.Controls.Add(Me.btnGetImpellers)
        Me.Controls.Add(Me.btnGetModelNo)
        Me.Controls.Add(Me.btnCalcRunCharts)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnCalcRunCharts As Button
    Friend WithEvents btnGetModelNo As Button
    Friend WithEvents btnGetImpellers As Button
    Friend WithEvents cmbSN As ComboBox
    Friend WithEvents cmbImpellers As ComboBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents txtDesignFlow As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents lblDesignFlow As Label
    Friend WithEvents txtDesignHead As TextBox
    Friend WithEvents lblDesignHead As Label
    Friend WithEvents lbSN As ListBox
End Class
