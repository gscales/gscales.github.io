function FindProxyForURL(url, host)
{
        // Microsoft login pages 
        if (dnsDomainIs(host, "login.microsoftonline.com") ||
            dnsDomainIs(host, "login.microsoft.com") ||
            dnsDomainIs(host, "login.windows.net") ||
            dnsDomainIs(host, ".microsoftonline-p.com"))
		    {		
            return "PROXY 127.0.0.1:8866";
		    }            
            
            return "DIRECT";;
 
}