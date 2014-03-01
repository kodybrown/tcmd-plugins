'
' Exports the properties of the specified shortcut (.lnk) including the
' target, arguments, directory, icon, description, hotkey, and windowstyle.
'
' Copyright (C) 2013-2014 Kody Brown (@wasatchwizard)
' Released under the MIT License.
'
' shortcuts.wlx is Copyright (C) Ivan aka Atlanoff (atlanoff@yandex.ru)
' originally named wlx_vbscript.wlx, part of wlx_vbscript_0_5_1.zip
'

font_name = "Consolas"
font_size = 9
view_end = "false"
view_wrap = "true"
view_scroll = "both"
'view_backgroundcolor = "silver"
'view_textcolor = "blue"

result_text = ""

' Objects
Set wso = CreateObject("Wscript.Shell")

Dim vbCrLf : vbCrLf = Chr(13) & Chr(10)
Dim link
Dim WindowStyles(3)

WindowStyles(1) = "Normal"
WindowStyles(2) = "Minimized"
WindowStyles(3) = "Maximized"

On Error Resume Next
Set link = wso.CreateShortcut(file_name)
On Error Goto 0

If Err.Number <> 0 Then
    result_text = "Could not open shortcut file." & vbCrLf & vbCrLf & Err.Description
Else
    ' result_text = vbCrLf _
    '             & "Target Path: " & vbCrLf & link.TargetPath & vbCrLf & vbCrLf _
    '             & "Arguments: " & vbCrLf & link.Arguments  & vbCrLf& vbCrLf _
    '             & "Working Directory: " & vbCrLf & link.WorkingDirectory  & vbCrLf & vbCrLf _
    '             & "Hotkey: " & link.Hotkey & vbCrLf _
    '             & "Window Style: " & WindowStyles(link.WindowStyle) & vbCrLf & vbCrLf _
    '             & "Icon Location: " & vbCrLf & link.IconLocation & vbCrLf & vbCrLf _
    '             & "Description: " & vbCrLf & " " & link.Description
    result_text = vbCrLf _
                & " Target Path       = " & link.TargetPath & vbCrLf _
                & " Arguments         = " & link.Arguments & vbCrLf _
                & vbCrLf _
                & " Working Directory = " & link.WorkingDirectory & vbCrLf _
                & vbCrLf _
                & " Hotkey            = " & link.Hotkey & vbCrLf _
                & " Window Style      = " & WindowStyles(link.WindowStyle) & vbCrLf _
                & vbCrLf _
                & " Icon Location     = " & link.IconLocation & vbCrLf _
                & vbCrLf _
                & " Description: " & vbCrLf & " " & link.Description & vbCrLf
End If

Set link = Nothing
Set wso = Nothing
