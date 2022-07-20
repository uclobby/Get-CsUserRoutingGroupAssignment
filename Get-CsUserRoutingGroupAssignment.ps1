<#
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
	SOFTWARE.
#>


<#
.DESCRIPTION
	This script returns the current routing group count from Lync/Skype for Business Front Ends.

.NOTES
  Version      	   		: 1.0
  Author    			: David Paulino https://uclobby.com
  
#>

[CmdletBinding()]
param(
[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
    [string] $PoolFqdn,
[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
	[switch] $ExcludeSBA,
[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
	[switch] $Detailed
)

$startTime=Get-Date;
#Query to get all Routing Groups in RTCLOCAL
$RGAQuery = "SELECT DISTINCT f.Fqdn, r.RoutingGroupName FROM dbo.RoutingGroupAssignment r, dbo.FrontEnd f WHERE f.FrontEndId = r.FrontEndId"

#Checking if the Lync/Skype for Business Module is available
if(!(Get-Module -ListAvailable -Name Lync,SkypeforBusiness)){
    Write-Warning "Could not find Lync/Skype for Business PowerShell Module."
    
    return
}

$ServerFqdn = [System.Net.Dns]::GetHostByName((hostname)).HostName

#If the PoolFQDN is missing we will try to use the current computer.
if($PoolFqdn){
    $ComputersInPool = (Get-CsComputer -Pool $PoolFqdn -ErrorAction SilentlyContinue)
    
} else {
    $ComputersInPool = (Get-CsComputer -Identity $ServerFqdn -ErrorAction SilentlyContinue)
    $PoolFqdn = $ComputersInPool.Pool
}

if($ComputersInPool){
    Write-Host "Pool FQDN:" $PoolFqdn -ForegroundColor Green

    #Push/Pop so avoid the SQLSERV "drive".
    Push-Location
    foreach($Computer in $ComputersInPool){
        try{
            $ServerInstance = $Computer.fqdn + "\RTCLOCAL"
            $RGAList = Invoke-Sqlcmd -query $RGAQuery -ServerInstance $ServerInstance -Database RTC -ErrorAction SilentlyContinue
            break
        } catch {
            Write-Warning "Failed to connect to: $ServerInstance"
        }
    }
    Pop-Location
    
    $RGAOutput = New-Object System.Collections.ArrayList
    if($ExcludeSBA){
        $RegistrarFQDNs = $PoolFqdn
    } else {
        $WebServerFQDN = "WebServer:"+$PoolFqdn
        $RegistrarFQDNs = Get-CsService -Registrar | Where-Object {$_.WebServer -eq $WebServerFQDN } | Select-Object PoolFqdn 
    }

    foreach($RegistrarFQDN in $RegistrarFQDNs ){
        if($ExcludeSBA){
            $fqdn = $RegistrarFQDN
        } else {
            $fqdn = $RegistrarFQDN.PoolFqdn
        }

        if($Detailed) {
            $RGs = Get-CsUser -Filter {RegistrarPool -eq $FQDN}| Select-Object UserRoutingGroupId, SipAddress

            foreach ($RG in $RGs) {
                if($RGAList) {
                    $FEfqdn = ($RGAList | Where-Object {$_.RoutingGroupName -eq $RG.UserRoutingGroupId}).Fqdn
                } else {
                    $FEfqdn = $null
                }
                $RGInfo = New-Object PSObject -Property @{            
                            RegistrarPool  = $fqdn
                            Pool           = $PoolFqdn
                            FrontEnd       = $FEfqdn
                            RoutingGroup   = $RG.UserRoutingGroupId
                            UserSipAddress = $RG.SipAddress
                          }
                [void]$RGAOutput.Add($RGInfo)
            }
        } else {
            $RGs = Get-CsUser -Filter {RegistrarPool -eq $FQDN}| Group-Object -Property UserRoutingGroupId | Select-Object Name, Count 
            foreach ($RG in $RGs) {
                if($RGAList) {
                    $FEfqdn = ($RGAList | Where-Object {$_.RoutingGroupName -eq $RG.Name}).Fqdn
                } else {
                    $FEfqdn = $null
                }
                $RGInfo = New-Object PSObject -Property @{            
                            RegistrarPool = $fqdn
                            FrontEnd      = $FEfqdn
                            RoutingGroup  = $RG.Name
                            UserCount     = $RG.Count
                          }
                [void]$RGAOutput.Add($RGInfo)
            }

            $RGATemp = $RGAOutput | Select RoutingGroup



            #Adding empty routing groups:
            foreach ($RG in $RGAList){
                if ($RGATemp -notmatch $RG.RoutingGroupName) {
                    $RGInfo = New-Object PSObject -Property @{            
                            RegistrarPool = $fqdn
                            FrontEnd      = $RG.Fqdn
                            RoutingGroup  = $RG.RoutingGroupName
                            UserCount     = 0
                          }
                   [void]$RGAOutput.Add($RGInfo)
                }
            }
        }
    }
    $endTime = Get-Date
    $totalTime= [math]::round(($endTime - $startTime).TotalSeconds,2)
    Write-Host "Date:" (Get-Date -format g) -ForegroundColor Yellow
    Write-Host "Execution time:" $totalTime "seconds" -ForegroundColor Cyan
    if($Detailed){
        $RGAOutput | Select-Object RegistrarPool, Pool, FrontEnd, RoutingGroup, UserSipAddress
    } else {
        $RGAOutput | Group-Object FrontEnd | Select-Object @{N="FrontEnd";E={$_.Name}},@{N="RGCount";E={$_.Count}},@{N="UserCount";E={($_.Group | measure-object UserCount -sum).sum}} | Sort-Object FrontEnd
    }
} else {
    Write-Warning "Invalid/unknown Pool FQDN."
}