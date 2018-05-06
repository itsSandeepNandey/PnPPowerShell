
 # Read the configuration file
 [xml]$XmlDocument = Get-Content -Path ".\PSConfig.xml";

function ConnectSPOnline()
{
    # Get ConnectSPOnline node properties
    $Connectionconfig = $XmlDocument.Configurations.ConnectSPOnline;
    $siteUrl= $Connectionconfig.SiteUrl;
    $userName= $Connectionconfig.UserName;
    $password= $Connectionconfig.Password;

    $credentials;
    if($userName -ne '' -and $password -ne ''){
        # Convert password to secure string    
        $secureStringPwd = ConvertTo-SecureString -AsPlainText $password -Force  

        # Create new credential object 
        $credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName,$secureStringPwd  
    }
    else {
        # If credential is blank in PSConfig file, get credential at run time 
        $credentials = (Get-Credential);
    }

    Try
    {
        write-host "Info: Connecting to site: " $siteURL  "..."

        # Connect to SP online site  
        Connect-PnPOnline -Url $siteURL -Credentials $credentials  
        
        Write-Host -ForegroundColor Green "Info: Connected !"
    }
    Catch
    {
        Write-Host -ForegroundColor Red "Error: Unable to connect !"
        Write-Host -ForegroundColor Red $_.Exception.Message
        Break
    }
     
}

<#***************************************************************************************#>
function StartScript()
{
    write-host -ForegroundColor Yellow "Info: Script execution started"   

    # Connect to SPOnline 
    ConnectSPOnline;

    write-host -ForegroundColor Yellow "Info: Execution Completed"   
}

# Start point of the script

# Clear the screen
Clear-Host
# Create log file 
$date= Get-Date -format MMddyyyyHHmmss  
start-transcript -path .\Logs\Log_$date.doc   

# Start the script executio
StartScript;