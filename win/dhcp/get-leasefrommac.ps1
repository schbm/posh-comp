#Author: mas
#Date: 29.11.21
function Get-LeaseInfoFromMac {
  <#
  .SYNOPSIS
      Finds an IP lease on the DHCP servers for the supplied MAC address.
  
  .DESCRIPTION
      This function searches for an IP lease on the DHCP servers based on the supplied MAC address. It returns a PSCustomObject with the following fields: Lease (the found lease object), ServerDns (DNS name of the DHCP server where the lease resides), and ServerIp (IP address of the DHCP server).
  
  .PARAMETER MacAddress
      [String] The MAC address to search for an IP lease. Format: "xx-xx-xx-xx-xx-xx"
  
  .OUTPUTS
      [PSCustomObject[]]
      Returns an array of PSCustomObjects containing the lease information.
  
  .EXAMPLE
      $leases = Get-LeaseInfoFromMac -MacAddress 'f4-92-bf-9a-a1-a7'
      $leases | % {$_.Lease}
  
  .NOTES 
      Author: Marcel Schubert
      LastEdit: 30.11.2021
  #>
  [CmdletBinding()]
  param(
      [Parameter(Mandatory = $true)]
      [ValidatePattern("^([0-9A-Fa-f]{2}-){5}[0-9A-Fa-f]{2}$")]
      [string]$MacAddress
  )

  $dhcpServers = Get-DhcpServerInDC
  $results = @()

  foreach ($dhcpServer in $dhcpServers) {
      $dhcpServerToSkip = @("dhcpServerName1", "dhcpServerName2")  # Specify DHCP servers to skip if needed

      if ($dhcpServer.DnsName -in $dhcpServerToSkip) {
          continue
      }

      $scopes = Get-DhcpServerv4Scope -ComputerName $dhcpServer.DnsName

      foreach ($scope in $scopes) {
          $leases = Get-DhcpServerv4Lease -ComputerName $dhcpServer.DnsName -ScopeId $scope.ScopeId -AllLeases
          $result = $leases | Where-Object { $_.ClientId -match $MacAddress }

          if ($result) {
              $results += [PSCustomObject]@{
                  Lease = $result
                  ServerDns = $dhcpServer.DnsName
                  ServerIp = $dhcpServer.IPAddress
              }
          }
      }
  }
  
  return $results
}