' ytdown.vbs - YouTube Video Downloader
' Inspired by Peteris Krumins' VBScript YouTube Downloader
' For educational purposes only

Option Explicit

Dim videoURL, videoID, videoTitle, downloadURL, savePath, http, stream, fso

' Prompt user for YouTube video URL
videoURL = InputBox("Enter the YouTube video URL:", "YouTube Video Downloader")

If videoURL = "" Then
    WScript.Echo "No URL entered. Exiting."
    WScript.Quit
End If

' Extract Video ID from URL
videoID = ExtractVideoID(videoURL)
If videoID = "" Then
    WScript.Echo "Invalid YouTube URL. Exiting."
    WScript.Quit
End If

' Fetch Video Information
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "https://www.youtube.com/get_video_info?video_id=" & videoID, False
http.Send

If http.Status <> 200 Then
    WScript.Echo "Failed to retrieve video information. Exiting."
    WScript.Quit
End If

' Parse Video Information
Dim videoInfo, args, i, arg
videoInfo = http.ResponseText
Set args = CreateObject("Scripting.Dictionary")
For Each arg In Split(videoInfo, "&")
    i = InStr(arg, "=")
    If i > 0 Then
        args(AddPercentEncoding(Left(arg, i - 1))) = AddPercentEncoding(Mid(arg, i + 1))
    End If
Next

' Check if video is playable
If args.Exists("status") And args("status") = "fail" Then
    WScript.Echo "Error: " & args("reason")
    WScript.Quit
End If

' Get video title
If args.Exists("title") Then
    videoTitle = args("title")
Else
    videoTitle = "Unknown_Title"
End If

' Get download URL
If args.Exists("url_encoded_fmt_stream_map") Then
    Dim fmtStreamMap, streams, streamInfo, streamURL
    fmtStreamMap = args("url_encoded_fmt_stream_map")
    streams = Split(fmtStreamMap, ",")
    For Each streamInfo In streams
        Set args = CreateObject("Scripting.Dictionary")
        For Each arg In Split(streamInfo, "&")
            i = InStr(arg, "=")
            If i > 0 Then
                args(AddPercentEncoding(Left(arg, i - 1))) = AddPercentEncoding(Mid(arg, i + 1))
            End If
        Next
        If args.Exists("url") Then
            streamURL = args("url")
            If args.Exists("quality") Then
                streamURL = streamURL & "&quality=" & args("quality")
            End If
            downloadURL = streamURL
            Exit For
        End If
    Next
Else
    WScript.Echo "No downloadable streams found. Exiting."
    WScript.Quit
End If

' Define save path
Set fso = CreateObject("Scripting.FileSystemObject")
savePath = fso.GetAbsolutePathName(".") & "\" & videoTitle & ".mp4"

' Download the video
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", downloadURL, False
http.Send

If http.Status = 200 Then
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile savePath, 2 ' Overwrite if exists
    stream.Close
    WScript.Echo "Video downloaded successfully: " & savePath
Else
    WScript.Echo "Failed to download video. HTTP Status: " & http.Status
End If

' Function to extract video ID from URL
Function ExtractVideoID(url)
    Dim regEx, matches
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "v=([^&]+)"
    regEx.IgnoreCase = True
    regEx.Global = False
    Set matches = regEx.Execute(url)
    If matches.Count > 0 Then
        ExtractVideoID = matches(0).SubMatches(0)
    Else
        ExtractVideoID = ""
    End If
End Function

' Function to decode percent-encoded strings
Function AddPercentEncoding(str)
    Dim i, ch, result
    result = ""
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If ch = "+" Then
            result = result & " "
        ElseIf ch = "%" Then
            result = result & Chr(CLng("&H" & Mid(str, i + 1, 2)))
            i = i + 2
        Else
            result = result & ch
        End If
    Next
    AddPercentEncoding = result
End Function
