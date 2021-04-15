function FindProxyForURL(url, host)
{
    // Internal domains - do not use any proxy
    if (shExpMatch(host,"*.test.com")||
        shExpMatch(host,"*.gsxclients.com")||
        shExpMatch(host,"*.gsx.com"))

        return "PROXY 127.0.0.1:8888";

        // Microsoft login pages 
        if (dnsDomainIs(host, "login.microsoftonline.com") ||
            dnsDomainIs(host, "login.microsoft.com") ||
            dnsDomainIs(host, "login.windows.net") ||
            dnsDomainIs(host, ".microsoftonline-p.com"))
		    {		
            return "PROXY 127.0.0.1:8888";
		    }            
            
            return "DIRECT";
 
}
