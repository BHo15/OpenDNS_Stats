'Created to allow for a simple one-click method of downloading and viewing OpenDNS stats
'On first run, prompts for OpenDNS username and network, and then asks if you want to store password.
	'If yes, It runs password through RC4 encryption, and stores it in the Check.ini file
	'If no, it asks for your password, but does not store it in the ini file.
'Also on the first run, it asks for your preferences on monitoring OpenDNS URL categories
	'It stores these prefs in the ini file, and uses them to shrink your output to just what you want to see
'The StatTemplate.xlt imports the csv DNLD from OpenDNS, and then will clean it up for simple viewing (VBA code)
'Brad Hodge <brad.h.hodge@gmail.com>

Dim strEmail
Dim strNetwork
Dim strP
Dim BDate
Dim EDate
Dim oShell
Dim network
Dim strK
	strK="#$UnEqu1v0cal!?"
Dim CurrentPath
	CurrentPath=Replace(WScript.ScriptFullName, WScript.ScriptName, "")

Call Check

BDate=InputBox("What is the beginning date?",,Year(date) & "-" & Right(String(2,"0") & Month(date),2) & "-" & Right(String(2,"0") & Day(date),2))
EDate=InputBox("What is the ending date?",,Year(date) & "-" & Right(String(2,"0") & Month(date),2) & "-" & Right(String(2,"0") & Day(date),2))

Set oShell = CreateObject("WScript.Shell")
Set network = CreateObject("WScript.Network")


If Not WScript.FullName = CurrentPath & "cscript.exe" Then
	oShell.Run "cmd.exe /c" & WScript.Path & "\cscript.exe //NOLOGO " & Chr(34) & "TweakedFetch.vbs" & Chr(34) & " " & Chr(34) & strEmail _
	& Chr(34) & " " & Chr(34) & strNetwork & Chr(34) & " " & Chr(34) & BDate & Chr(34) & " " & Chr(34) & EDate & Chr(34) & " " _
	& Chr(34) & strP & Chr(34) & " >DNLD.csv"
    WScript.Quit 0
End If

Function Check()
	Dim objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTextFile
	Dim vbsFile
	Dim strRC
	Dim strNextLine
	Dim intLineFinder
	Dim strNewFile
	Dim blFirst
		blFirst=False
	Dim EncP
	Const ForReading = 1
	Const ForWriting = 2


	Set objTextFile = objFSO.OpenTextFile(CurrentPath & "Check.ini", ForReading)
	
	Set vbsFile= objFSO.OpenTextFile(CurrentPath & "RC4.vbs",1,False)
	strRC =  vbsFile.ReadAll
	vbsFile.Close
	Set vbsFile = Nothing
	ExecuteGlobal strRC
	
	Do Until objTextFile.AtEndOfStream
		strNextLine=objTextFile.Readline
		
		intLineFinder=InStr(strNextLine, "use=")
		If intLineFinder<>"0" Then
			If strNextLine="use=0" Then'First use of app
				blFirst=True
				strEmail=Inputbox("Please type email", "First time use setup...")
				strNetwork = Inputbox ("Please type your OpenDNS network name.", "First time use setup...")
				If MsgBox("Do you want to store your password?", vbYesNo, "First time use setup...") = vbYes Then
					strP = Inputbox("Please type in your OpenDNS password.", "First time use setup...")
					If strP<>"" Then
						EncP=fCrypt(strP,strK)
					Else
						EncP=""
						strP=""
					End If
				End If
				strNextLine="use=1"
			End If
		End If
		
		intLineFinder=InStr(strNextLine, "Email=")
		If intLineFinder<>"0" Then
			If blFirst= True Then
				strNextLine="Email=" & strEmail
			Else strEmail = Trim(Mid(strNextLine,Instr(strNextLine,"=")+1,50))
			End If
		End If
		
		intLineFinder=InStr(strNextLine, "Network=")
		If intLineFinder<>"0" Then
			If blFirst= True Then
				strNextLine="Network=" & strNetwork
			Else
				strNetwork=Trim(Mid(strNextLine,Instr(strNextLine,"=")+1,50))
			End If
		End If
		
		intLineFinder=InStr(strNextLine, "appPW=")
		 If intLineFinder<>"0" Then
			If EncP<>"" Then
				strNextLine="appPW=" & EncP
			Else
				EncP=Trim(Mid(strNextLine,Instr(strNextLine,"PW=")+3,50))
				strP=fCrypt(EncP,strK)
			End If
		 End If
				
        If blFirst = True Then 'Checks for first use... Does not look at categories unless it's first use
            intLineFinder = InStr(strNextLine, "Blacklisted=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Blacklisted'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Blacklisted=1"
                Else
                    strNextLine = "Blacklisted=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Blocked by Category=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Blocked by Category'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Blocked by Category=1"
                Else
                    strNextLine = "Blocked by Category=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Blocked as Botnet=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Blocked as Botnet'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Blocked as Botnet=1"
                Else
                    strNextLine = "Blocked as Botnet=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Blocked as Malware=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Blocked as Malware'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Blocked as Malware=1"
                Else
                    strNextLine = "Blocked as Malware=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Blocked as Phishing=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Blocked as Phishing'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Blocked as Phishing=1"
                Else
                    strNextLine = "Blocked as Phishing=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Resolved by SmartCache=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Resolved by SmartCache'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Resolved by SmartCache=1"
                Else
                    strNextLine = "Resolved by SmartCache=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Academic Fraud=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Academic Fraud'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Academic Fraud=1"
                Else
                    strNextLine = "Academic Fraud=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Adult Themes=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Adult Themes'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Adult Themes=1"
                Else
                    strNextLine = "Adult Themes=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Adware=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Adware'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Adware=1"
                Else
                    strNextLine = "Adware=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Alcohol=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Alcohol'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Alcohol=1"
                Else
                    strNextLine = "Alcohol=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Anime/Manga/Webcomic=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Anime/Manga/Webcomic'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Anime/Manga/Webcomic=1"
                Else
                    strNextLine = "Anime/Manga/Webcomic=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Auctions=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Auctions'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Auctions=1"
                Else
                    strNextLine = "Auctions=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Automotive=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Automotive'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Automotive=1"
                Else
                    strNextLine = "Automotive=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Blogs=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Blogs'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Blogs=1"
                Else
                    strNextLine = "Blogs=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Business Services=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Business Services'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Business Services=1"
                Else
                    strNextLine = "Business Services=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Chat=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Chat'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Chat=1"
                Else
                    strNextLine = "Chat=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Classifieds=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Classifieds'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Classifieds=1"
                Else
                    strNextLine = "Classifieds=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Dating=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Dating'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Dating=1"
                Else
                    strNextLine = "Dating=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Drugs=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Drugs'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Drugs=1"
                Else
                    strNextLine = "Drugs=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Ecommerce/Shopping=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Ecommerce/Shopping'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Ecommerce/Shopping=1"
                Else
                    strNextLine = "Ecommerce/Shopping=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Educational Institutions=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Educational Institutions'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Educational Institutions=1"
                Else
                    strNextLine = "Educational Institutions=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "File Storage=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'File Storage'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "File Storage=1"
                Else
                    strNextLine = "File Storage=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Financial Institutions=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Financial Institutions'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Financial Institutions=1"
                Else
                    strNextLine = "Financial Institutions=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Forums/Message boards=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Forums/Message boards'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Forums/Message boards=1"
                Else
                    strNextLine = "Forums/Message boards=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Gambling=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Gambling'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Gambling=1"
                Else
                    strNextLine = "Gambling=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Games=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Games'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Games=1"
                Else
                    strNextLine = "Games=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "German Youth Protection=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'German Youth Protection'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "German Youth Protection=1"
                Else
                    strNextLine = "German Youth Protection=0"
                End If
            End If
			
			intLineFinder = InStr(strNextLine, "Government=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Government'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Government=1"
                Else
                    strNextLine = "Government=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Hate/Discrimination=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Hate/Discrimination'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Hate/Discrimination=1"
                Else
                    strNextLine = "Hate/Discrimination=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Health and Fitness=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Health and Fitness'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Health and Fitness=1"
                Else
                    strNextLine = "Health and Fitness=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Humor=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Humor'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Humor=1"
                Else
                    strNextLine = "Humor=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Instant Messaging=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Instant Messaging'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Instant Messaging=1"
                Else
                    strNextLine = "Instant Messaging=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Jobs/Employment=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Jobs/Employment'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Jobs/Employment=1"
                Else
                    strNextLine = "Jobs/Employment=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Lingerie/Bikini=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Lingerie/Bikini'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Lingerie/Bikini=1"
                Else
                    strNextLine = "Lingerie/Bikini=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Movies=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Movies'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Movies=1"
                Else
                    strNextLine = "Movies=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Music=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Music'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Music=1"
                Else
                    strNextLine = "Music=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "News/Media=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'News/Media'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "News/Media=1"
                Else
                    strNextLine = "News/Media=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Non-Profits=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Non-Profits'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Non-Profits=1"
                Else
                    strNextLine = "Non-Profits=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Nudity=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Nudity'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Nudity=1"
                Else
                    strNextLine = "Nudity=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "P2P/File sharing=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'P2P/File sharing'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "P2P/File sharing=1"
                Else
                    strNextLine = "P2P/File sharing=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Parked Domains=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Parked Domains'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Parked Domains=1"
                Else
                    strNextLine = "Parked Domains=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Photo Sharing=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Photo Sharing'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Photo Sharing=1"
                Else
                    strNextLine = "Photo Sharing=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Podcasts=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Podcasts'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Podcasts=1"
                Else
                    strNextLine = "Podcasts=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Politics=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Politics'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Politics=1"
                Else
                    strNextLine = "Politics=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Pornography=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Pornography'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Pornography=1"
                Else
                    strNextLine = "Pornography=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Portals=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Portals'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Portals=1"
                Else
                    strNextLine = "Portals=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Proxy/Anonymizer=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Proxy/Anonymizer'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Proxy/Anonymizer=1"
                Else
                    strNextLine = "Proxy/Anonymizer=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Radio=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Radio'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Radio=1"
                Else
                    strNextLine = "Radio=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Religious=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Religious'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Religious=1"
                Else
                    strNextLine = "Religious=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Research/Reference=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Research/Reference'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Research/Reference=1"
                Else
                    strNextLine = "Research/Reference=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Search Engines=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Search Engines'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Search Engines=1"
                Else
                    strNextLine = "Search Engines=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Sexuality=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Sexuality'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Sexuality=1"
                Else
                    strNextLine = "Sexuality=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Social Networking=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Social Networking'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Social Networking=1"
                Else
                    strNextLine = "Social Networking=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Software/Technology=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Software/Technology'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Software/Technology=1"
                Else
                    strNextLine = "Software/Technology=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Sports=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Sports'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Sports=1"
                Else
                    strNextLine = "Sports=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Tasteless=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Tasteless'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Tasteless=1"
                Else
                    strNextLine = "Tasteless=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Television=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Television'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Television=1"
                Else
                    strNextLine = "Television=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Tobacco=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Tobacco'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Tobacco=1"
                Else
                    strNextLine = "Tobacco=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Travel=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Travel'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Travel=1"
                Else
                    strNextLine = "Travel=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Typo Squatting=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Typo Squatting'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Typo Squatting=1"
                Else
                    strNextLine = "Typo Squatting=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Video Sharing=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Video Sharing'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Video Sharing=1"
                Else
                    strNextLine = "Video Sharing=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Visual Search Engines=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Visual Search Engines'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Visual Search Engines=1"
                Else
                    strNextLine = "Visual Search Engines=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Weapons=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Weapons'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Weapons=1"
                Else
                    strNextLine = "Weapons=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Web Spam=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Web Spam'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Web Spam=1"
                Else
                    strNextLine = "Web Spam=0"
                End If
            End If
            
            intLineFinder = InStr(strNextLine, "Webmail=")
            If intLineFinder <> "0" Then
                If MsgBox("Do you want to monitor 'Webmail'?", vbYesNo, "First time use setup - categories...") = vbYes Then
                    strNextLine = "Webmail=1"
                Else
                    strNextLine = "Webmail=0"
                End If
            End If
        End If 'Ends check of blFirst

		strNewFile = strNewFile & strNextLine & vbCrLf

	Loop

	objTextFile.Close
	
	Set objTextFile=objFSO.OpenTextFile("Check.ini",ForWriting)
	
	objTextFile.WriteLine strNewFile
	objTextFile.Close
End Function
