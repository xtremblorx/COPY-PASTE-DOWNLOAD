Set re = New RegExp 
dim x 
dim lenth
Dim text : text = CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text") 
lenth = IsNull(text)

if lenth = True Then
MsgBox "Please copy a URL string."
Wscript.Quit
End if
dim hell 
hell=text 
 With re 
           .Pattern    = "[h][t][t][ps]*://www.youtube.com/watch" 
           .IgnoreCase = False 
           .Global     = False 
        End With 
  If re.Test( hell ) Then 
          x=0 
        Else 
          x=1 
        End If 
 With re 
           .Pattern    = "[h][t][t][ps]*://youtube.com/watch" 
           .IgnoreCase = False 
           .Global     = False 
        End With 
  If x = 0 OR re.Test( hell ) Then 
          x=0 
        Else 
          x=1 
        End If 
if x = 0 Then 
Set WshShell = CreateObject("WScript.Shell") 
MsgBox "Downloading " & text 
WshShell.Run "cmd /c title Download & color 20 & mode 30,40 & echo Downloading now... & echo off & youtube-dl -f bestaudio " & text , 1 
Set WshShell = Nothing 
Else 
MsgBox "URL copied appears to be Invalid." 
End If 
