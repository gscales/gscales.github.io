function FindProxyForURL(url, host)
{
    if(host == "fs.gsxclients.com"){
	return "DIRECT";
    }else{
	return "PROXY 127.0.0.1:8888";
    }
 
}