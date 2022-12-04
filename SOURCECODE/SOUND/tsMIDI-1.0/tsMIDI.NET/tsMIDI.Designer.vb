<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class tsMain
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			Static fTerminateCalled As Boolean
			If Not fTerminateCalled Then
				Form_Terminate_renamed()
				fTerminateCalled = True
			End If
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents cmdPlay As System.Windows.Forms.Button
	Public WithEvents cmdStop As System.Windows.Forms.Button
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(tsMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdPlay = New System.Windows.Forms.Button
		Me.cmdStop = New System.Windows.Forms.Button
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
		Me.Text = "tsMIDI-1.0"
		Me.ClientSize = New System.Drawing.Size(134, 34)
		Me.Location = New System.Drawing.Point(3, 19)
		Me.Icon = CType(resources.GetObject("tsMain.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "tsMain"
		Me.cmdPlay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdPlay.Text = "&Play"
		Me.cmdPlay.Size = New System.Drawing.Size(67, 34)
		Me.cmdPlay.Location = New System.Drawing.Point(66, 0)
		Me.cmdPlay.TabIndex = 1
		Me.cmdPlay.BackColor = System.Drawing.SystemColors.Control
		Me.cmdPlay.CausesValidation = True
		Me.cmdPlay.Enabled = True
		Me.cmdPlay.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdPlay.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdPlay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdPlay.TabStop = True
		Me.cmdPlay.Name = "cmdPlay"
		Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdStop.Text = "&Stop"
		Me.cmdStop.Size = New System.Drawing.Size(67, 34)
		Me.cmdStop.Location = New System.Drawing.Point(0, 0)
		Me.cmdStop.TabIndex = 0
		Me.cmdStop.BackColor = System.Drawing.SystemColors.Control
		Me.cmdStop.CausesValidation = True
		Me.cmdStop.Enabled = True
		Me.cmdStop.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdStop.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdStop.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdStop.TabStop = True
		Me.cmdStop.Name = "cmdStop"
		Me.Controls.Add(cmdPlay)
		Me.Controls.Add(cmdStop)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class