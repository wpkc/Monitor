'
'	MediaMonkey Script
'	Name: Monitor
'	[monitor]
'	Filename=monitor.vbs
'	Description=Adds context menus for sending sound files to a external audio monitoring program
'	Language=VBScript
'	ScriptType=0 Auto Script

'	monitor.vbs is copied to scripts\auto
'	Options are set upon installation. Reinstall to change options.
'	--------------------------------------------------------------------	
'	This program is free software: you can redistribute it and/or modify
'	it under the terms of the GNU General Public License as published by
'	the Free Software Foundation, either version 3 of the License, or
'	(at your option) any later version.
'
'	This program is distributed in the hope that it will be useful,
'	but WITHOUT ANY WARRANTY; without even the implied warranty of
'	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'	GNU General Public License for more details.
'
'	You should have received a copy of the GNU General Public License
'	along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
Option Explicit
Const regKey = "kc-monitor"
Dim MonitorPath : MonitorPath=""
Dim MonitorPlayFlags : MonitorPlayFlags=""
Dim MonitorQueueFlags : MonitorQueueFlags=""

Sub OnStartup
	MonitorReadSettings
	If SDB.Tools.FileSystem.FileExists(MonitorPath)<>True Then Exit Sub
	
	Dim PlayCaption : PlayCaption = "Play in Monitor"
	Dim QueueCaption : QueueCaption = "Queue to Monitor"

	Dim UI : Set UI = SDB.UI
	' Add context menu to track list
	Dim puTrackListPlay : Set puTrackListPlay = UI.AddMenuItem(UI.Menu_Pop_TrackList,0,-3)
	puTrackListPlay.Caption = PlayCaption
	puTrackListPlay.OnClickFunc = "Monitor_Play"
	puTrackListPlay.UseScript = Script.ScriptPath
	puTrackListPlay.IconIndex = 14
	puTrackListPlay.Enabled = True

	
	Dim puTracklistQueue : Set puTracklistQueue = UI.AddMenuItem(UI.Menu_Pop_TrackList,0,-3)
	puTracklistQueue.Caption = QueueCaption
	puTracklistQueue.OnClickFunc = "Monitor_Queue"
	puTracklistQueue.UseScript = Script.ScriptPath
	puTracklistQueue.IconIndex = 16
	puTracklistQueue.Enabled = True
	
	' Add context menu to Now Playing list
	Dim puNowListPlay : Set puNowListPlay = UI.AddMenuItem(UI.Menu_Pop_NP,0,-3)
	puNowListPlay.Caption = PlayCaption
	puNowListPlay.OnClickFunc = "Monitor_Play"
	puNowListPlay.UseScript = Script.ScriptPath
	puNowListPlay.IconIndex = 14
	puNowListPlay.Enabled = True
	
	Dim puNowListQueue : Set puNowListQueue = UI.AddMenuItem(UI.Menu_Pop_NP,0,-3)
	puNowListQueue.Caption = QueueCaption
	puNowListQueue.OnClickFunc = "Monitor_Queue"
	puNowListQueue.UseScript = Script.ScriptPath
	puNowListQueue.IconIndex = 16
	puNowListQueue.Enabled = True

End Sub

Sub Monitor_Play(arg)

	Dim i, WShell, ShellCmd, SongCount, Song, Result

	MonitorReadSettings
	If SDB.Tools.FileSystem.FileExists(MonitorPath)<>True Then Exit Sub

	SongCount = SDB.SelectedSongList.Count
	If Not SongCount > 0 Then Exit Sub

	ShellCmd = Chr(34) & MonitorPath & Chr(34) & " " & MonitorPlayFlags	
	SongCount = SongCount - 1 ' Adjust for zero index
	For i=0 To SongCount

		Set Song = SDB.SelectedSongList.Item(i)
		If TypeName(Song)="SDBSongData" Then
	
			If Song.Cached = 1 Then
				ShellCmd = ShellCmd & " " & Chr(34) & Song.CachedPath & Chr(34)
			Else
				If (Left(Song.Path,1) <> "?") Then
					ShellCmd = ShellCmd & " " & Chr(34) & Song.Path & Chr(34)
				End If
			End If
		End If
	Next

	If Len(ShellCmd)>0 Then
		Set WShell = CreateObject("WScript.Shell")
		WShell.Run ShellCmd, 8, False
		Set WShell = Nothing
	End If
End Sub

Sub Monitor_Queue(arg)
	Dim i, WShell, ShellCmd, SongCount, Song, Result

	MonitorReadSettings
	If SDB.Tools.FileSystem.FileExists(MonitorPath)<>True Then Exit Sub

	SongCount = SDB.SelectedSongList.Count
	If Not SongCount > 0 Then Exit Sub

	ShellCmd = Chr(34) & MonitorPath & Chr(34) & " " & MonitorQueueFlags
	SongCount = SongCount - 1 ' Adjust for zero index
		
	For i=0 To SongCount
		Set Song = SDB.SelectedSongList.item(i)
		If TypeName(Song)="SDBSongData" Then
			If Song.Cached = 1 Then
				ShellCmd = ShellCmd & " " & Chr(34) & Song.CachedPath & Chr(34)
			Else
				If (Left(Song.Path,1) <> "?") Then
					ShellCmd = ShellCmd & " " & Chr(34) & Song.Path & Chr(34)
				End If
			End If
		End If
	Next
	
	If Len(ShellCmd)>0 Then
		Set WShell = CreateObject("WScript.Shell")
		WShell.Run ShellCmd, 8, False
		Set WShell = Nothing
	End If	
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
	Set Reg = Nothing
End Sub
