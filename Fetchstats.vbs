' Modified from OpenDNS Stats Fetch for Windows  Brad Hodge <brad.h.hodge@gmail.com>
' 	Based on original fetchstats script from Richard Crowley
' 	Brian Hartvigsen <brian.hartvigsen@opendns.com>


Dim CurPath
	CurPath = Replace(WScriptFullName, WScript.ScriptName, "")
Dim strK
	strK="#$UnEqu1v0cal!?"
Dim strNetwork
Dim UserName
Dim Password,strP
Dim BDate,EDate, DateRange
Dim objHTTP
	Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
Dim URL
	URL = "https://dashboard.opendns.com"
Dim strEmail
Dim objPassword
Dim regEx
Dim data

Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

Function GetUrlData(strUrl, strMethod, strData)
        objHTTP.Open strMethod, strUrl
        If strMethod = "POST" Then
                objHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        End If
        objHTTP.Option(6) = False
        objHTTP.Send(strData)

        If Err.Number <> 0 Then
                GetUrlData = "ERROR: " & Err.Description & vbCrLf & Err.Source & " (" & Err.Nmber & ")"
        Else
                GetUrlData = objHTTP.ResponseText
        End If
End Function

URL="https://dashboard.opendns.com"


Username = Wscript.Arguments.Item(0)
If Len(Username) = 0 Then: Usage
Network = Wscript.Arguments.Item(1)
If Len(Network) = 0 Then: Usage
BDate = Wscript.Arguments.Item(2)
If Len(BDate) = 0 Then: Usage
EDate = Wscript.Arguments.Item(3)
If Len(EDate) = 0 Then: Usage
strP = Wscript.Arguments.Item(4)

Set regDate = new RegExp
regDate.Pattern = "^\d{4}-\d{2}-\d{2}$"

DateRange = BDate & "to" & EDate

' Are they running Vista or 7?
On Error Resume Next
Set objPassword = CreateObject("ScriptPW.Password")
	If strP="" Then strP= Inputbox("Please enter your password.")
    Password = strP

Wscript.StdErr.Write vbCrLf
On Error GoTo 0

Set regEx = New RegExp
regEx.IgnoreCase = true
regEx.Pattern = ".*name=""formtoken"" value=""([0-9a-f]*)"".*"

data = GetUrlData(URL & "/signin", "GET", "")

Set Matches = regEx.Execute(data)
token = Matches(0).SubMatches(0)

data = GetUrlData(URL & "/signin", "POST", "formtoken=" & token & "&username=" & Escape(Username) & "&password=" & Escape(Password) & "&sign_in_submit=foo")
If Len(data) <> 0 Then
        Wscript.StdErr.Write "Login Failed. Check username and password" & vbCrLf
        WScript.Quit 1
End If

page=1
Do While True
        data = GetUrlData(URL & "/stats/" & Network & "/topdomains/" & DateRange & "/page" & page & ".csv", "GET", "")
        If page = 1 Then
                If LenB(data) = 0 Then
                        WScript.StdErr.Write "You can not access " & Network & vbCrLf
                        WScript.Quit 2
                ElseIf InStr(data, "<!DOCTYPE") Then
                 Wscript.StdErr.Write "Error retrieving data. Date range may be outside of available data."
                 Wscript.Quit 2
                End If
        Else
                ' First line is always header
                data=Mid(data,InStr(data,vbLf)+1)
        End If
        If LenB(Trim(data)) = 0 Then
                Exit Do
        End If
        Wscript.StdOut.Write data
        page = page + 1
Loop
'--------------------------------------------------------
'DISPLAY AND FORMAT THE DATA IN EXCEL
Dim xlApp, xlBook
Dim CurrentPath
	CurrentPath=Replace(WScript.ScriptFullName, WScript.ScriptName, "")
Set xlApp = CreateObject("Excel.Application")
Set xlBook = xlApp.Workbooks.Open(CurrentPath & "\StatTemplate.xltm")
xlApp.Run "Process",(CurrentPath)
'--------------------------------------------------------

Function Usage()
	MsgBox("Username, OpenDNS Network, Begin Date, and Ending Date must be passed from 001-CallFetch.vbs")
End Function
