Option Explicit
Const regKey = "kc-monitor"
Dim MonitorPath : MonitorPath = ""
Dim MonitorPlayFlags : MonitorPlayFlags = ""
Dim MonitorQueueFlags : MonitorQueueFlags = "/add"

Sub MonitorSelectExe(obj)
	Dim dlg : Set dlg = SDB.CommonDialog
	dlg.Title = "Select Audio Monitor Program"
	dlg.Filter = "Programs (*.exe)|*.exe|All files (*.*)|*.*"
	dlg.FilterIndex = 1
	dlg.Flags = cdlOFNPathMustExist + cdlOFNFileMustExist
	dlg.InitDir = "C:\Program Files (x86)\"
	dlg.ShowOpen
	 
	If Not dlg.Ok Then Exit Sub
	 
	Dim Form1 : Set Form1 = obj.Common.Parent
	If Not (Form1 Is Nothing) Then Form1.Common.ChildControl("txtMonitorPath").Text = dlg.FileName
End Sub

Sub MonitorReadSettings()
	Dim Reg : Set Reg = SDB.Registry
	If Reg.OpenKey(regKey, True) Then

		If Reg.ValueExists("MonitorPath") Then
		  MonitorPath = Reg.StringValue("MonitorPath")
		End If

		If Reg.ValueExists("MonitorPlayFlags") Then
		  MonitorPlayFlags = Reg.StringValue("MonitorPlayFlags")
		End If

		If Reg.ValueExists("MonitorQueueFlags") Then
		  MonitorQueueFlags = Reg.StringValue("MonitorQueueFlags")
		End If

		Reg.CloseKey
	End If
End Sub

Sub MonitorSaveSettings()
	Dim Reg : Set Reg = SDB.Registry

	Dim Form1 : Set Form1 = SDB.Objects("MonitorSetupForm")
	If Not (Form1 Is Nothing) Then
		Script.UnregisterEvents Form1

		Dim FormCommon : Set FormCommon = Form1.Common
		MonitorPath = FormCommon.ChildControl("txtMonitorPath").Text
		MonitorPlayFlags = FormCommon.ChildControl("txtPlayFlags").Text
		MonitorQueueFlags = FormCommon.ChildControl("txtQueueFlags").Text

		Set Form1 = Nothing
		SDB.Objects("MonitorSetupForm") = Nothing
	End If

	If Reg.OpenKey(regKey, True) Then
		Reg.StringValue("MonitorPath") = MonitorPath
		Reg.StringValue("MonitorPlayFlags") = MonitorPlayFlags
		Reg.StringValue("MonitorQueueFlags") = MonitorQueueFlags

		Reg.CloseKey
	End If
End Sub


' Installation starts here
MonitorReadSettings
  
'*******************************************************************'
'* Form produced by MMVBS Form Creator (http://trixmoto.net/mmvbs) *'
'*******************************************************************'
Dim UI : Set UI = SDB.UI

Dim Form1 : Set Form1 = UI.NewForm
Form1.BorderStyle = 4
Form1.Caption = "Audio Monitor Settings"
Form1.FormPosition = 4
Form1.Common.SetRect 0,0,380,240

Dim Edit1 : Set Edit1 = UI.NewEdit(Form1)
Edit1.Common.SetRect 20,31,320,20
Edit1.Common.ControlName = "txtMonitorPath"
Edit1.Common.Hint = "Enter complete path to audio monitor program."
Edit1.Text = MonitorPath

Dim Edit2 : Set Edit2 = UI.NewEdit(Form1)
Edit2.Common.SetRect 20,84,121,21
Edit2.Common.ControlName = "txtPlayFlags"
Edit2.Text = MonitorPlayFlags

Dim Edit3 : Set Edit3 = UI.NewEdit(Form1)
Edit3.Common.SetRect 20,141,121,21
Edit3.Common.ControlName = "txtQueueFlags"
Edit3.Text = MonitorQueueFlags

Dim Label1 : Set Label1 = UI.NewLabel(Form1)
Label1.Common.SetRect 20,16,65,17
Label1.Caption = "Audio Monitor Program Path"

Dim Label2 : Set Label2 = UI.NewLabel(Form1)
Label2.Common.SetRect 20,69,65,17
Label2.Caption = "Play Command Flags"

Dim Label3 : Set Label3 = UI.NewLabel(Form1)
Label3.Common.SetRect 20,127,65,17
Label3.Caption = "Queue Command Flags"

Dim Button1 : Set Button1 = UI.NewButton(Form1)
Button1.Caption = "..."
Button1.Common.SetRect 340,30,16,19
Button1.Common.ControlName = "btnBrowse"
Button1.UseScript = Script.ScriptPath
Button1.OnClickFunc = "MonitorSelectExe"

Dim Button2 : Set Button2 = UI.NewButton(Form1)
Button2.Caption = "Save"
Button2.ModalResult = 1
Button2.Default = True
Button2.Common.SetRect 280,171,75,25
Button2.UseScript = Script.ScriptPath

'*******************************************************************'
'* End of form                              Richard Lewis (c) 2007 *'
'*******************************************************************'	

SDB.Objects("MonitorSetupForm") = Form1

MonitorSelectExe(Button1)
If Form1.ShowModal=1 Then MonitorSaveSettings

Script.Reload Script.ScriptPath & "monitor.vbs"