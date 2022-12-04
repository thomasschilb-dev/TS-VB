Option Strict Off
Option Explicit On
Friend Class tsMain
	Inherits System.Windows.Forms.Form
	Private Sub cmdPlay_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPlay.Click
		PlayMID("music")
	End Sub
	
	Private Sub cmdStop_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStop.Click
		StopMID("music")
	End Sub
	
	Private Sub tsMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim MidiFile As String
		MidiFile = "tsMIDI.mid"
		OpenMID(MidiFile, "music")
		PlayMID("music")
	End Sub
	
	'UPGRADE_NOTE: Form_Terminate was upgraded to Form_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	'UPGRADE_WARNING: tsMain event Form.Terminate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub Form_Terminate_Renamed()
		CloseMID("music")
	End Sub
	
	Private Sub tsMain_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		CloseMID("music")
	End Sub
End Class