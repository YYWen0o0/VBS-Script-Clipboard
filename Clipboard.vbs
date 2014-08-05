'------------------------------------------------------------------------
'循环检查剪贴板，变化了，弹窗提示
'------------------------------------------------------------------------
Option Explicit
Dim WS
Set WS = WScript.CreateObject("Wscript.Shell")
Dim oHtml
Set oHtml=CreateObject("htmlfile")
Dim oldClipboardText,Changed,ClipboardText
ClipboardText=oHtml.ParentWindow.ClipboardData.GetData("text")
oldClipboardText=ClipboardText
Do
	ClipboardText=oHtml.ParentWindow.ClipboardData.GetData("text")
	Changed=(ClipboardText<>oldClipboardText)
	oldClipboardText=ClipboardText
	If Changed Then
		WS.Popup "Ctrl+C",1,"提示:",4096
	Else
		wscript.sleep 100
	End If
	ClipboardText=""
Loop
Set ws=Nothing
Set oHtml=Nothing
