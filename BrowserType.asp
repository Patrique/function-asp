<%
Private Function BrowserType()

	Dim strUserAgent	'Holds info on the users browser and os
	Dim strBrowserUserType	'Holds the users browser type

	'Get the users HTTP user agent (web browser)
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

	'Get the uesrs web browser

	'AOL
	If inStr(1, strUserAgent, "AOL", 1) Then
		strBrowserUserType = "AOL"
		
	'Edge
	ElseIf inStr(1, strUserAgent, "Edge", 1) Then
		strBrowserUserType = "Edge"
	
	'Konqueror
	ElseIf inStr(1, strUserAgent, "Konqueror", 1) Then
		strBrowserUserType = "Konqueror"

	'EudoraWeb
	ElseIf inStr(1, strUserAgent, "EudoraWeb", 1) Then
		strBrowserUserType = "EudoraWeb"

	'Dreamcast
	ElseIf inStr(1, strUserAgent, "Dreamcast", 1) Then
		strBrowserUserType = "Dreamcast"
		
	'Lynx
	ElseIf inStr(1, strUserAgent, "Lynx", 1) Then
		strBrowserUserType = "Lynx"

	'iCab
	ElseIf inStr(1, strUserAgent, "iCab", 1) Then
		strBrowserUserType = "iCab"

	'HotJava
	ElseIf inStr(1, strUserAgent, "Sun", 1) AND inStr(1, strUserAgent, "Mozilla/3", 1) Then
		strBrowserUserType = "HotJava"

	'Galeon
	ElseIf inStr(1, strUserAgent, "Galeon", 1) Then
		strBrowserUserType = "Galeon"

	'Epiphany
	ElseIf inStr(1, strUserAgent, "Epiphany", 1) Then
		strBrowserUserType = "Epiphany"

	'DocZilla
	ElseIf inStr(1, strUserAgent, "DocZilla", 1) Then
		strBrowserUserType = "DocZilla"

	'Camino
	ElseIf inStr(1, strUserAgent, "Chimera", 1) OR inStr(1, strUserAgent, "Camino", 1) Then
		strBrowserUserType = "Camino"

	'Dillo
	ElseIf inStr(1, strUserAgent, "Dillo", 1) Then
		strBrowserUserType = "Dillo"

	'amaya
	ElseIf inStr(1, strUserAgent, "amaya", 1) Then
		strBrowserUserType = "Amaya"

	'NetCaptor
	ElseIf inStr(1, strUserAgent, "NetCaptor", 1) Then
		strBrowserUserType = "NetCaptor"

	'Twiceler
	ElseIf inStr(1, strUserAgent, "Twiceler", 1) Then
		strBrowserUserType = "Twiceler"
		
	'ICE
	ElseIf inStr(1, strUserAgent, "ICE", 1) Then
		strBrowserUserType = "ICE"

	'LookSmart search engine robot
	ElseIf inStr(1, strUserAgent, "ZyBorg", 1) Then
		strBrowserUserType = "LookSmart"
	
	'Googlebot-Mobile search engine robot
	ElseIf inStr(1, strUserAgent, "Googlebot-Mobile", 1) Then
		strBrowserUserType = "Google/Mobile"

	'Googlebot search engine robot
	ElseIf inStr(1, strUserAgent, "Googlebot", 1) Then
		strBrowserUserType = "Google"

	 'Google/AdSense search engine robot
    	ElseIf inStr(1, strUserAgent, "Mediapartners-Google", 1) Then
        	strBrowserUserType = "Google/AdSense"

	'MSN  search engine robot
	ElseIf inStr(1, strUserAgent, "msnbot", 1) Then
		strBrowserUserType = "Bing"
	
	'Bing  search engine robot
	ElseIf inStr(1, strUserAgent, "bingbot", 1) Then
		strBrowserUserType = "Bing"

	'inktomi search engine robot
	ElseIf inStr(1, strUserAgent, "slurp", 1) Then
		strBrowserUserType = "Yahoo"
	
	'YahooSeeker search engine robot
	ElseIf inStr(1, strUserAgent, "YahooSeeker/M1A1-R2D2", 1) Then
		strBrowserUserType = "Yahoo/Mobile"

	'AltaVista search engine robot
	ElseIf inStr(1, strUserAgent, "Scooter", 1) Then
		strBrowserUserType = "AltaVista"

	'DMOZ search engine robot
	ElseIf inStr(1, strUserAgent, "Robozilla", 1) Then
		strBrowserUserType = "DMOZ"

	'Ask Jeeves search engine robot
	ElseIf inStr(1, strUserAgent, "Ask Jeeves", 1) OR inStr(1, strUserAgent, "Ask+Jeeves", 1) Then
		strBrowserUserType = "Ask Jeeves"

	'Lycos search engine robot
	ElseIf inStr(1, strUserAgent, "lycos", 1) Then
		strBrowserUserType = "Lycos"

	'Excite search engine robot
	ElseIf inStr(1, strUserAgent, "ArchitextSpider", 1) Then
		strBrowserUserType = "Excite"

	'Facebook bot
	ElseIf inStr(1, strUserAgent, "facebook", 1) Then
		strBrowserUserType = "Facebook"
		
	'LinkedInBot search engine robot
	ElseIf inStr(1, strUserAgent, "LinkedInBot", 1) Then
		strBrowserUserType = "LinkedIn"

	'Northernlight search engine robot
	ElseIf inStr(1, strUserAgent, "Gulliver", 1) Then
		strBrowserUserType = "Northernlight"

	'AllTheWeb search engine robot
	ElseIf inStr(1, strUserAgent, "crawler@fast", 1) Then
		strBrowserUserType = "AllTheWeb"

	'Turnitin search engine robot
	ElseIf inStr(1, strUserAgent, "TurnitinBot", 1) Then
		strBrowserUserType = "Turnitin"
	
	'PostRank search engine robot
	ElseIf inStr(1, strUserAgent, "PostRank", 1) Then
		strBrowserUserType = "PostRank"

	'InternetSeer search engine robot
	ElseIf inStr(1, strUserAgent, "internetseer", 1) Then
		strBrowserUserType = "InternetSeer"

	'NameProtect Inc. search engine robot
	ElseIf inStr(1, strUserAgent, "nameprotect", 1) Then
		strBrowserUserType = "NameProtect"

	'PhpDig search engine robot
	ElseIf inStr(1, strUserAgent, "PhpDig", 1) Then
		strBrowserUserType = "PhpDig"

	'Rambler search engine robot
	ElseIf inStr(1, strUserAgent, "StackRambler", 1) Then
		strBrowserUserType = "Rambler"

	'UbiCrawler search engine robot
	ElseIf inStr(1, strUserAgent, "UbiCrawler", 1) Then
		strBrowserUserType = "UbiCrawler"

	'entireweb search engine robot
	ElseIf inStr(1, strUserAgent, "Speedy+Spider", 1) OR inStr(1, strUserAgent, "Speedy Spider", 1) Then
		strBrowserUserType = "Entireweb"

	'Alexa.com search engine robot
	ElseIf inStr(1, strUserAgent, "ia_archiver", 1) Then
		strBrowserUserType = "Alexa"

	'Arianna/Libero search engine robot
	ElseIf inStr(1, strUserAgent, "arianna.libero.it", 1) Then
		strBrowserUserType = "Arianna/Libero"

	'y2bot/1.0 (+http://bot.y2crack4.com) search engine robot
	ElseIf inStr(1, strUserAgent, "y2bot", 1) Then
		strBrowserUserType = "y2bot"
		
	'Baiduspider search engine robot (Chineese Search Engine)
	ElseIf inStr(1, strUserAgent, "Baiduspider", 1) Then
		strBrowserUserType = "Baidu"
		
	'Sogou search engine robot (Chineese Search Engine)
	ElseIf inStr(1, strUserAgent, "Sogou", 1) Then
		strBrowserUserType = "Sogou"
		
	'Exabot search engine robot
	ElseIf inStr(1, strUserAgent, "Exabot", 1) Then
		strBrowserUserType = "Exabot"
		
	'YandexBot search engine robot
	ElseIf inStr(1, strUserAgent, "YandexBot", 1) Then
		strBrowserUserType = "Yandex"
		
	'Amazon robot checking their affiliate sites 
	ElseIf inStr(1, strUserAgent, "aranhabot", 1) Then
		strBrowserUserType = "Amazon.com"
	
	'Brandwatch robot best off being blocked! 
	ElseIf inStr(1, strUserAgent, "magpie-crawler", 1) Then
		strBrowserUserType = "Brandwatch"
		
	'AhrefsBot - SEO Spybots
	ElseIf inStr(1, strUserAgent, "AhrefsBot", 1) Then
		strBrowserUserType = "AhrefsBot"
		
	'008 - 80legs
	ElseIf inStr(1, strUserAgent, "80legs", 1) Then
		strBrowserUserType = "008"
		
	'DotBot - moz.com
	ElseIf inStr(1, strUserAgent, "DotBot", 1) Then
		strBrowserUserType = "DotBot - MOZ"
	
	'MJ12bot
	ElseIf inStr(1, strUserAgent, "MJ12bot", 1) Then
		strBrowserUserType = "Majestic"
		
	'VigLink
	ElseIf inStr(1, strUserAgent, "VigLink", 1) Then
		strBrowserUserType = "VigLink" 


	'Opera Mini
	ElseIf InStr(1, strUserAgent, "Opera Mini", 1) Then
		strBrowserUserType = "Opera Mini"
	
	'Opera
	ElseIf InStr(1, strUserAgent, "Opera", 1) Then
		strBrowserUserType = "Opera"
	
	'Google Chrome
	ElseIf inStr(1, strUserAgent, "Chrome", 1) Then
		strBrowserUserType = "Chrome"
	
	'Google Chrome Mobilee
	ElseIf inStr(1, strUserAgent, "CrMo", 1) Then
		strBrowserUserType = "Chrome Mobile"
	
	'Mobile Safari
	ElseIf inStr(1, strUserAgent, "Mobile Safari", 1) Then
		strBrowserUserType = "Mobile Safari"
		
	'Safari
	ElseIf inStr(1, strUserAgent, "Safari", 1) Then
		strBrowserUserType = "Safari"
		

	'Internet Explorer
	ElseIf inStr(1, strUserAgent, "MSIE", 1) OR inStr(1, strUserAgent, "Trident", 1) Then
		strBrowserUserType = "IE"


	'Pocket Internet Explorer
	ElseIf inStr(1, strUserAgent, "MSPIE", 1) Then
		strBrowserUserType = "Pocket IE"

	'Firefox
	ElseIf inStr(1, strUserAgent, "Firefox", 1) Then
		strBrowserUserType = "Firefox"
		
	
	'Netscape
	ElseIf inStr(1, strUserAgent, "Netscape/9", 1) Then
		strBrowserUserType = "Netscape 9"
	ElseIf inStr(1, strUserAgent, "Netscape/8", 1) Then
		strBrowserUserType = "Netscape 8"
	ElseIf inStr(1, strUserAgent, "Netscape/7", 1) Then
		strBrowserUserType = "Netscape 7"
	ElseIf inStr(1, strUserAgent, "Netscape6", 1) Then
		strBrowserUserType = "Netscape 6"
	ElseIf inStr(1, strUserAgent, "Mozilla/4", 1) Then
		strBrowserUserType = "Netscape 4"
	
	
	'Mozilla
	ElseIf inStr(1, strUserAgent, "Gecko", 1) Then
		strBrowserUserType = "Mozilla"


	'Else unknown or robot
	Else
		strBrowserUserType = "Unknown"
	End If

	'Return function
	BrowserType = strBrowserUserType
End Function
%>
