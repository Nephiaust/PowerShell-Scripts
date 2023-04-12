#Requires -Version 5.0

# Referenced https://stackoverflow.com/a/17354800
#    For class based system

# Referenced https://www.reddit.com/r/usefulscripts/comments/7xjwlc/comment/du9lxha/?utm_source=share&utm_medium=web2x&context=3
#    For new way of splitting string up

$Settings = @{
    Servers = @("SERVER.DOMAIN.LOCAL","SERVER2.DOMAIN.LOCAL")
    RootHints = "http://www.internic.net/domain/named.cache"
    TrustSSLCert = $true
    ClearExisting = $true
    IPv4 = $true
    IPv6 = $true
}

Write-Debug 'Creating DNS Record class'
Class DNSRecord {
    [string]$HostName
    [int]$ttl
    [string]$RecordType
    [string]$IPAddress;
}

Write-Debug 'Getting the current date & timestamp'
$CurrentTimestamp = Get-Date

$TempRootHints = $env:TEMP + '\RootHint.zone_' + $CurrentTimestamp.ToString('yyyyMMddTHHmmss')

If (!($Settings.TrustSSLCert)) {
    Write-Debug 'Creating new type to disable SSL/TLS checking'
    add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true} ;
}

Write-Debug 'Creating System.Net.WebClient object'
$WebClient = new-object System.Net.WebClient
$WebClient.Headers['User-Agent'] = "My PowerShell Downloader 1.0"

Write-Debug 'Downloading the root hints file'
$WebClient.Downloadfile($Settings.RootHints, $TempRootHints)

Write-Debug 'Loading the root hints file'
$RootHintsRaw = Get-Content -Path $TempRootHints | Select-String -Pattern '^[A-M].ROOT' | Sort-Object

Write-Debug 'Processing the ROOT name hint file'
Remove-Variable NewRootHints -ErrorAction SilentlyContinue
$NewRootHints = New-Object System.Collections.ArrayList
ForEach ($RootHintRaw in $RootHintsRaw) {
    $NewRootHint = [DNSRecord]::new()
    $NewRootHint.HostName=($RootHintRaw -split "\s+")[0].Substring(0,($RootHintRaw -split "\s+")[0].Length-1)
    $NewRootHint.ttl=($RootHintRaw -split "\s+")[1]
    $NewRootHint.RecordType=($RootHintRaw -split "\s+")[2]
    $NewRootHint.IPAddress=($RootHintRaw -split "\s+")[3]
    $NewRootHints.Add($NewRootHint) | Out-Null
    Remove-Variable NewRootHint
}

Write-Debug 'Processing server list'
$UniqueRootHints = $NewRootHints.Hostname | Select-Object -Unique
ForEach ($Server in $Settings.Servers) {
    Write-Host "Working on $Server"
    Foreach ($RootHint in $UniqueRootHints) {
        if ($Settings.ClearExisting) {Remove-DnsServerRootHint -Computer $Server -Force -NameServer $RootHint}
        if ($Settings.IPv4) {
            $Current = $NewRootHints | Where-Object {($_.Hostname -eq $RootHint) -and ($_.RecordType -eq 'A')}
            Add-DnsServerRootHint -Computer $Server -NameServer $Current.HostName -IPAddress $Current.IPAddress
        }
        if ($Settings.IPv6) {
            $Current = $NewRootHints | Where-Object {($_.Hostname -eq $RootHint) -and ($_.RecordType -eq 'AAAA')}
            Add-DnsServerRootHint -Computer $Server -NameServer $Current.HostName -IPAddress $Current.IPAddress
        }
    }
    
}
