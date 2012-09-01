On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration",,48)
For Each objItem in colItems
    wscript.echo "############################################################################"
	Wscript.Echo "ArpAlwaysSourceRoute: " & objItem.ArpAlwaysSourceRoute
    Wscript.Echo "ArpUseEtherSNAP: " & objItem.ArpUseEtherSNAP
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "DatabasePath: " & objItem.DatabasePath
    Wscript.Echo "DeadGWDetectEnabled: " & objItem.DeadGWDetectEnabled
    Wscript.Echo "DefaultIPGateway: " & objItem.DefaultIPGateway
    Wscript.Echo "DefaultTOS: " & objItem.DefaultTOS
    Wscript.Echo "DefaultTTL: " & objItem.DefaultTTL
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DHCPEnabled: " & objItem.DHCPEnabled
    Wscript.Echo "DHCPLeaseExpires: " & objItem.DHCPLeaseExpires
    Wscript.Echo "DHCPLeaseObtained: " & objItem.DHCPLeaseObtained
    Wscript.Echo "DHCPServer: " & objItem.DHCPServer
    Wscript.Echo "DNSDomain: " & objItem.DNSDomain
    Wscript.Echo "DNSDomainSuffixSearchOrder: " & objItem.DNSDomainSuffixSearchOrder
    Wscript.Echo "DNSEnabledForWINSResolution: " & objItem.DNSEnabledForWINSResolution
    Wscript.Echo "DNSHostName: " & objItem.DNSHostName
    Wscript.Echo "DNSServerSearchOrder: " & objItem.DNSServerSearchOrder
    Wscript.Echo "DomainDNSRegistrationEnabled: " & objItem.DomainDNSRegistrationEnabled
    Wscript.Echo "ForwardBufferMemory: " & objItem.ForwardBufferMemory
    Wscript.Echo "FullDNSRegistrationEnabled: " & objItem.FullDNSRegistrationEnabled
    Wscript.Echo "GatewayCostMetric: " & objItem.GatewayCostMetric
    Wscript.Echo "IGMPLevel: " & objItem.IGMPLevel
    Wscript.Echo "Index: " & objItem.Index
    Wscript.Echo "IPAddress: " & objItem.IPAddress
    Wscript.Echo "IPConnectionMetric: " & objItem.IPConnectionMetric
    Wscript.Echo "IPEnabled: " & objItem.IPEnabled
    Wscript.Echo "IPFilterSecurityEnabled: " & objItem.IPFilterSecurityEnabled
    Wscript.Echo "IPPortSecurityEnabled: " & objItem.IPPortSecurityEnabled
    Wscript.Echo "IPSecPermitIPProtocols: " & objItem.IPSecPermitIPProtocols
    Wscript.Echo "IPSecPermitTCPPorts: " & objItem.IPSecPermitTCPPorts
    Wscript.Echo "IPSecPermitUDPPorts: " & objItem.IPSecPermitUDPPorts
    Wscript.Echo "IPSubnet: " & objItem.IPSubnet
    Wscript.Echo "IPUseZeroBroadcast: " & objItem.IPUseZeroBroadcast
    Wscript.Echo "IPXAddress: " & objItem.IPXAddress
    Wscript.Echo "IPXEnabled: " & objItem.IPXEnabled
    Wscript.Echo "IPXFrameType: " & objItem.IPXFrameType
    Wscript.Echo "IPXMediaType: " & objItem.IPXMediaType
    Wscript.Echo "IPXNetworkNumber: " & objItem.IPXNetworkNumber
    Wscript.Echo "IPXVirtualNetNumber: " & objItem.IPXVirtualNetNumber
    Wscript.Echo "KeepAliveInterval: " & objItem.KeepAliveInterval
    Wscript.Echo "KeepAliveTime: " & objItem.KeepAliveTime
    Wscript.Echo "MACAddress: " & objItem.MACAddress
    Wscript.Echo "MTU: " & objItem.MTU
    Wscript.Echo "NumForwardPackets: " & objItem.NumForwardPackets
    Wscript.Echo "PMTUBHDetectEnabled: " & objItem.PMTUBHDetectEnabled
    Wscript.Echo "PMTUDiscoveryEnabled: " & objItem.PMTUDiscoveryEnabled
    Wscript.Echo "ServiceName: " & objItem.ServiceName
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "TcpipNetbiosOptions: " & objItem.TcpipNetbiosOptions
    Wscript.Echo "TcpMaxConnectRetransmissions: " & objItem.TcpMaxConnectRetransmissions
    Wscript.Echo "TcpMaxDataRetransmissions: " & objItem.TcpMaxDataRetransmissions
    Wscript.Echo "TcpNumConnections: " & objItem.TcpNumConnections
    Wscript.Echo "TcpUseRFC1122UrgentPointer: " & objItem.TcpUseRFC1122UrgentPointer
    Wscript.Echo "TcpWindowSize: " & objItem.TcpWindowSize
    Wscript.Echo "WINSEnableLMHostsLookup: " & objItem.WINSEnableLMHostsLookup
    Wscript.Echo "WINSHostLookupFile: " & objItem.WINSHostLookupFile
    Wscript.Echo "WINSPrimaryServer: " & objItem.WINSPrimaryServer
    Wscript.Echo "WINSScopeID: " & objItem.WINSScopeID
    Wscript.Echo "WINSSecondaryServer: " & objItem.WINSSecondaryServer
Next