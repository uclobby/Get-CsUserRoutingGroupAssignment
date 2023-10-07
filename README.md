# Get-CsUserRoutingGroupAssignment

Returns the routing group user count from Lync/Skype for Business Front Ends. 

<b>Parameters</b>
<ul>
    <li>PoolFqdn - Specify the Pool FQDN that we want to get the user per routing group count.</li>
    <li>ExcludeSBA - Switch if we want to exclude users that are homed on a Survivable Branch Appliance that is associated with the Front End Pool.</li>
    <li>Detailed - It will output the users SIP addresses.</li>
</ul>
<b>Release Notes</b>
<ul>
    <li>Version 1.0: 2019/07/16 - Initial release.</li>
    <li>Version 1.1: 2023/10/07 - Updated to publish in PowerShell Gallery.</li>
</ul>