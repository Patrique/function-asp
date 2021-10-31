<%
Private Function OSType ()

	Dim strUserAgent	'Holds info on the users browser and os
	Dim strOS		'Holds the users OS

	'Get the users HTTP user agent (web browser)
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
	

	'Get users OS
	
	'entireweb search engine robot (idfied as windows if not done now)
	If inStr(1, strUserAgent, "Speedy+Spider", 1) OR inStr(1, strUserAgent, "Speedy Spider", 1) Then
		strOS = "Search Robot"
	
	'Windows 10 Mobile
	ElseIf inStr(1, strUserAgent, "Windows NT 10", 1) AND inStr(1, strUserAgent, "ARM", 1) Then
		strOS = "Windows 10 Mobile" 
	'Windows 10 
	ElseIf inStr(1, strUserAgent, "Windows NT 10", 1) OR inStr(1, strUserAgent, "Windows 10", 1) Then
		strOS = "Windows 10"  'Only show windows 10 even if Windows Server Next
	'Windows RT 8.1 (both NT6.3)
	ElseIf inStr(1, strUserAgent, "NT 6.3", 1) AND inStr(1, strUserAgent, "ARM", 1) Then
		strOS = "Windows RT 8.1" 
	'Windows 8.1 and Windows Server 2012 R2 (both NT6.3)
	ElseIf inStr(1, strUserAgent, "Windows 8.1", 1) OR inStr(1, strUserAgent, "NT 6.3", 1) Then
		strOS = "Windows 8.1"  'Only show windows 8.1 even if Windows Server 2012 R2
	'Windows RT 8 (both NT6.2)
	ElseIf inStr(1, strUserAgent, "NT 6.3", 1) AND inStr(1, strUserAgent, "ARM", 1) Then
		strOS = "Windows RT 8"
	'Windows 8 and Windows Server 2012 (both NT6.2)
	ElseIf inStr(1, strUserAgent, "Windows 8", 1) OR inStr(1, strUserAgent, "NT 6.2", 1) Then
		strOS = "Windows 8"  'Only show windows 8 even if Windows Server 2012
	'Windows 7 and Windows 2008 R2 (both NT6.1)
	ElseIf inStr(1, strUserAgent, "Windows 7", 1) OR inStr(1, strUserAgent, "NT 6.1", 1) Then
		strOS = "Windows 7"  'Only show windows 7 even if Windows 2008 R2
	'Windows Vista and Windows Server 2008 (both NT6.0)
	ElseIf inStr(1, strUserAgent, "Windows Vista", 1) OR inStr(1, strUserAgent, "NT 6.0", 1) Then
		strOS = "Windows Vista"  
	ElseIf inStr(1, strUserAgent, "Windows 2003", 1) OR inStr(1, strUserAgent, "NT 5.2", 1) Then
		strOS = "Windows 2003"
	ElseIf inStr(1, strUserAgent, "Windows XP", 1) OR inStr(1, strUserAgent, "NT 5.1", 1) Then
		strOS = "Windows XP"
	ElseIf inStr(1, strUserAgent, "NT 5.01", 1) Then
		strOS = "Windows 2000 SP1"
	ElseIf inStr(1, strUserAgent, "Windows 2000", 1) OR inStr(1, strUserAgent, "NT 5", 1) Then
		strOS = "Windows 2000"
	ElseIf inStr(1, strUserAgent, "Windows NT", 1) OR inStr(1, strUserAgent, "WinNT", 1) Then
		strOS = "Windows  NT 4"
	ElseIf inStr(1, strUserAgent, "Windows 95", 1) OR inStr(1, strUserAgent, "Win95", 1) Then
		strOS = "Windows 95"
	ElseIf inStr(1, strUserAgent, "Windows ME", 1) OR inStr(1, strUserAgent, "Win 9x 4.90", 1) Then
		strOS = "Windows ME"
	ElseIf inStr(1, strUserAgent, "Windows 98", 1) OR inStr(1, strUserAgent, "Win98", 1) Then
		strOS = "Windows 98"
	ElseIf Instr(1, strUserAgent, "Windows CE", 1) Then
		strOS = "Windows CE"
	ElseIf Instr(1, strUserAgent, "Windows Phone OS 7.0", 1) Then
		strOS = "Windows Phone 7"

	'Android 
	ElseIf inStr(1, strUserAgent, "Android", 1) Then
		strOS = "Android"
	
	'PalmOS
	ElseIf inStr(1, strUserAgent, "PalmOS", 1) Then
		strOS = "Palm OS"

	'PalmPilot
	ElseIf inStr(1, strUserAgent, "Elaine", 1) Then
		strOS = "PalmPilot"

	'Nokia
	ElseIf inStr(1, strUserAgent, "Nokia", 1) Then
		strOS = "Nokia"
		 
	'Ubuntu
	ElseIf inStr(1, strUserAgent, "Ubuntu", 1) Then
		strOS = "Ubuntu"

	'Amiga
	ElseIf inStr(1, strUserAgent, "Amiga", 1) Then
		strOS = "Amiga"

	'Solaris
	ElseIf inStr(1, strUserAgent, "Solaris", 1) Then
		strOS = "Solaris"

	'SunOS
	ElseIf inStr(1, strUserAgent, "SunOS", 1) Then
		strOS = "Sun OS"

	'BSD
	ElseIf inStr(1, strUserAgent, "BSD", 1) or inStr(1, strUserAgent, "FreeBSD", 1) Then
		strOS = "Free BSD"

	'Unix
	ElseIf inStr(1, strUserAgent, "Unix", 1) OR inStr(1, strUserAgent, "X11", 1) Then
		strOS = "Unix"

	'AOL webTV
	ElseIf inStr(1, strUserAgent, "AOLTV", 1) OR inStr(1, strUserAgent, "AOL_TV", 1) Then
		strOS = "AOL TV"
	ElseIf inStr(1, strUserAgent, "WebTV", 1) Then
		strOS = "Web TV"
		
		
	'iPad
	ElseIf inStr(1, strUserAgent, "iPad", 1) Then
		strOS = "iPad"
		
	'iPhone
	ElseIf inStr(1, strUserAgent, "iPhone", 1) Then
		strOS = "iPhone"
		
	'iPod
	ElseIf inStr(1, strUserAgent, "iPod", 1) Then
		strOS = "iPod"	
	
	'Linux
	ElseIf inStr(1, strUserAgent, "Linux", 1) Then
		strOS = "Linux"
		
	'VigLink
	ElseIf inStr(1, strUserAgent, "VigLink", 1) Then
		strOS = "Robot"


	'Machintosh
	ElseIf inStr(1, strUserAgent, "Mac OS X", 1) Then
		strOS = "Mac OS X"
	ElseIf inStr(1, strUserAgent, "Mac_PowerPC", 1) or Instr(1, strUserAgent, "PPC", 1) Then
		strOS = "Mac PowerPC"
	ElseIf inStr(1, strUserAgent, "Mac", 1) or inStr(1, strUserAgent, "apple", 1) Then
		strOS = "Macintosh"

	'OS/2
	ElseIf inStr(1, strUserAgent, "OS/2", 1) Then
		strOS = "OS/2"


	'Content Scraper
	ElseIf  inStr(1, strUserAgent, "nameprotect", 1) OR inStr(1, strUserAgent, "magpie-crawler", 1) OR inStr(1, strUserAgent, "AhrefsBot", 1) OR inStr(1, strUserAgent, "80legs.com", 1) Then
		strOS = "Content Scraper"

	'Search Robot
	ElseIf inStr(1, strUserAgent, "Googlebot", 1) OR inStr(1, strUserAgent, "LinkedInBot", 1) OR inStr(1, strUserAgent, "PostRank", 1) OR inStr(1, strUserAgent, "Mediapartners-Google", 1) OR inStr(1, strUserAgent, "ZyBorg", 1) OR inStr(1, strUserAgent, "slurp", 1) OR inStr(1, strUserAgent, "Scooter", 1) OR inStr(1, strUserAgent, "Robozilla", 1) OR inStr(1, strUserAgent, "Jeeves", 1) OR inStr(1, strUserAgent, "lycos", 1) OR inStr(1, strUserAgent, "ArchitextSpider", 1) OR inStr(1, strUserAgent, "Gulliver", 1) OR inStr(1, strUserAgent, "crawler@fast", 1) OR inStr(1, strUserAgent, "TurnitinBot", 1) OR inStr(1, strUserAgent, "internetseer", 1) OR inStr(1, strUserAgent, "nameprotect", 1) OR inStr(1, strUserAgent, "PhpDig", 1) OR inStr(1, strUserAgent, "StackRambler", 1) OR inStr(1, strUserAgent, "UbiCrawler", 1) OR inStr(1, strUserAgent, "Spider", 1) OR inStr(1, strUserAgent, "ia_archiver", 1) OR inStr(1, strUserAgent, "bingbot", 1) OR inStr(1, strUserAgent, "msnbot", 1) OR inStr(1, strUserAgent, "arianna.libero.it", 1) OR inStr(1, strUserAgent, "y2bot", 1) OR inStr(1, strUserAgent, "Twiceler", 1) OR inStr(1, strUserAgent, "Baiduspider", 1) OR inStr(1, strUserAgent, "YandexBot", 1) OR inStr(1, strUserAgent, "facebook", 1) OR inStr(1, strUserAgent, "Exabot", 1) OR inStr(1, strUserAgent, "DotBot", 1) OR inStr(1, strUserAgent, "MJ12bot", 1) Then
		strOS = "Search Robot"

	Else
		strOS = "Unknown"
	End If
	
	'Teseting
	'strOS = "Search Robot"

	'Return function
	OSType = strOS
End Function
%>
