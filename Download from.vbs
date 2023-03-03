' Hello,
'------------------------------------------------------------------------------------
' This VB SCRIPT DOWNLOADS A YOUTUBE VIDEO After Its Url is Copied Into the ClipboaRD
'---------------------------------------------------------------------------------------
' Please make sure youtube-dl is in the PATH
' You may also use yt-dlp, if you like but change youtube-dl below to it in the code.
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
'REPLACE c:\users\my\downloads\youtube-dl.exe with your PATH
'WshShell.Run "cmd /c title Download & set path=%path%;c:\users\my\downloads\&color 20 & mode 30,40 & echo Checking YOUTUBE-DL EXISTS... & echo off & youtube-dl --version 2>NUL&&echo FOUNDexe || echo NOT FOUNDexe & timeout 20 & exit"
WshShell.Run "cmd /c title Download & set path=%path%;c:\users\my\downloads\youtube-dl.exe& color 20 & mode 30,40 & echo Downloading now... & echo off & youtube-dl -f bestaudio " & text , 1 
Set WshShell = Nothing 
Else 
MsgBox "URL copied appears to be Invalid." 
End If 
