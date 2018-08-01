function Get-Services-List([string]$targetServer = $(throw "targetServer is required!"))
{
    $service = Get-WmiObject -Namespace "root\cimv2" -Class "Win32_Service" -ComputerName $targetServer -Impersonation 3    
    return $service
}

function Get-Service([string]$serviceName = $(throw "serviceName is required!"), [string]$targetServer = $(throw "targetServer is required!"))
{
    $service = Get-WmiObject -Namespace "root\cimv2" -Class "Win32_Service" -ComputerName $targetServer -Filter "Name='$serviceName'" -Impersonation 3    
    return $service
}


function Install-Service([string]$serviceName = $(throw "serviceName is required!"), [string]$targetServer = $(throw "targetServer is required!"),[string]$displayName = $(throw "displayName is required!"),[string]$physicalPath = $(throw "physicalPath is required!"),[string]$userName = $(throw "userName is required!"),[string]$password = "",[string]$startMode = "Automatic",[string]$description = "",[bool]$interactWithDesktop = $false)
{
    $params = $serviceName, $displayName, $physicalPath, 16, 1, $startMode, $interactWithDesktop, $userName, $password, $null, $null, $null           
    
	$scope = new-object System.Management.ManagementScope("\\$targetServer\root\cimv2", (new-object System.Management.ConnectionOptions))
    "Connecting to $targetServer"
    $scope.Connect()
    $mgt = new-object System.Management.ManagementClass($scope, (new-object System.Management.ManagementPath("Win32_Service")), (new-object System.Management.ObjectGetOptions))
      
    $op = "Service $serviceName ($physicalPath) on $targetServer"    
    "Installing $op..."
    $result = $mgt.InvokeMethod("Create", $params)    
    Test-ServiceResult -operation "Install $op" -result $result
    "Installed $op!"
      
    "Setting $serviceName description to '$description'..."
    Set-Service -ComputerName $targetServer -Name $serviceName -Description $description
    "Service install complete!"
}

function Start-Service([string]$serviceName = $(throw "serviceName is required!"), [string]$targetServer = $(throw "targetServer is required!"))
{
    "Getting service $serviceName on server $targetServer"
    $service = Get-Service $serviceName $targetServer
    if (!($service.Started))
    {
        "Starting service $serviceName on server $targetServer..."
        $result = $service.StartService()
        Test-ServiceResult -operation "Starting service $serviceName on $targetServer" -result $result   
    }
}

function Stop-Service([string]$serviceName = $(throw "serviceName is required!"), [string]$targetServer = $(throw "targetServer is required!"))
{
    "Getting service $serviceName on server $targetServer..."
    $service = Get-Service $serviceName $targetServer
    if (($service.Started))
    {
        "Stopping service $serviceName on server $targetServer..."
        $result = $service.StopService()
        Test-ServiceResult -operation "Stopping service $serviceName on $targetServer" -result $result   
    }
}

function Uninstall-Service([string]$serviceName = $(throw "serviceName is required!"), [string]$targetServer = $(throw "targetServer is required!"))
{
    $service = Get-Service $serviceName $targetServer
     
    if (!($service))
    { 
        Write-Warning "Failed to find service $serviceName on $targetServer!"
        return
    }     
    "Found service $serviceName on $targetServer; checking status..."
             
    if ($service.Started)
    {
        "Stopping service $serviceName on $targetServer..."
        #could also use Set-Service, net stop, SC, psservice, psexec etc.
        $result = $service.StopService()
        Test-ServiceResult -operation "Stop service $serviceName on $targetServer" -result $result
    }
     
    "Attempting to uninstall service $serviceName on $targetServer..."
    $result = $service.Delete()
    Test-ServiceResult -operation "Delete service $serviceName on $targetServer" -result $result   
}




function Test-ServiceResult([string]$operation = $(throw "Operation is required!"), [object]$result = $(throw "Result is required!"), [switch]$continueOnError = $false)
{
    $retVal = -1
    if ($result.GetType().Name -eq "UInt32") { $retVal = $result } else {$retVal = $result.ReturnValue}         
    if ($retVal -eq 0) {return}     
    $errorcode = 'Success,Not Supported,Access Denied,Dependent Services Running,Invalid Service Control,Service Cannot Accept Control,Service Not Active,Service Request Timeout,Unknown Failure,Path Not Found,Service Already Running,Service Database Locked,Service Dependency Deleted,Service Dependency Failure,Service Disabled,Service Logon Failure,Service Marked for Deletion,Service No Thread,Status Circular Dependency,Status Duplicate Name,Status Invalid Name,Status Invalid Parameter,Status Invalid Service Account,Status Service Exists,Service Already Paused'
    $desc = $errorcode.Split(',')[$retVal]     
    $msg = ("Operation {0} failed with code {1}:{2}" -f $operation, $retVal, $desc)     
    if (!$continueOnError) { Write-Error $msg } else { Write-Warning $msg }        
}


