Set WshNetwork = WScript.CreateObject("WScript.Network")

On Error Resume Next

WshNetwork.MapNetworkDrive "G:","\\Csi\General",0,"Administrator","5492!"

WshNetwork.MapNetworkDrive "S:","\\Csi\Data",0,"Administrator","5492!"


