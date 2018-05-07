
# Read the configuration file
[xml]$XmlDocument = Get-Content -Path ".\PSConfig.xml";


function CreateListItems() {
    write-host ""
    write-host "Info: Creating list items..."

    # List Name 
    $listName = 'Contacts'   
    Try {
        # Field Values
        $item_1_Values = @{"Title" = "Sandeep Nandey"; "Contact_x0020_Number" = "29480323"}  
     
        # Create new list item just by providing List Name & Field Values
        $item = Add-PnPListItem -List $listName -Values $item_1_Values 

        Write-Host -ForegroundColor Green "Item created with title: " $item["Title"]
    }
    Catch {
        Write-Host -ForegroundColor Red "Error at CreateListItems(): "$_.Exception.Message
    }   
}

function ReadListItems() {
    write-host ""
    write-host "Info: Reading list items.."

    # List Name 
    $listName = 'Contacts'
    # Field Names To Retrive          
    $fields = "Title", "Contact_x0020_Number", "Address"
    Try {
        # Get list items just by providing List Name & Field Names
        $listItems = Get-PnPListItem -List $listName -Fields $fields
        # Iterate through each list items
        foreach ($listItem in $listItems) {  
            Write-Host $listItem["ID"]:  $listItem["Title"]  ','  $listItem["Address"] ',' $listItem["Contact_x0020_Number"]  
        }  
    }
    Catch {
        Write-Host -ForegroundColor Red "Error at ReadListItems(): "$_.Exception.Message
    }   
}

function UpdateListItems () {
    write-host "";
    write-host "Info: Updating list items...";

    # List Name 
    $listName = 'Contacts'   
    Try {
        # Field Values
        
        # Get the list-item using CAML Query   
        $listItems = Get-PnPListItem -List  $listName -Query "<View><Query><Where><Eq><FieldRef Name='Contact_x0020_Number'/><Value Type='Text'>29480323</Value></Eq></Where></Query></View>"  
        Write-Host "Total items found: "  $listItems.count
        Write-Host ""
        foreach ($listItem in $listItems) {  

            TRY {
                Write-Host "Updating value for item, ID: " $listItem["ID"]
    
                # Prepare the object with item values
                $item_Values = @{"Title" = "Sandeep"; "Contact_x0020_Number" = "10101010"}  
            
                # Update the list-item
                $updatedItem = Set-PnPListItem -List $listName -Identity $listItem -Values $item_Values  
        
                Write-Host -ForegroundColor Green "Value updated for item, ID: " $updatedItem["ID"]
            } 
            Catch {
                Write-Host -ForegroundColor Red "Error while updating item with ID:" $listItem["ID"] +  $_.Exception.Message
                Write-Host -ForegroundColor Red $_.Exception.Message
            }   
        }  
    }
    Catch {
        Write-Host -ForegroundColor Red "Error at UpdateListItems(): "$_.Exception.Message
    }   
}

function DeleteListItems () {
    write-host "";
    write-host "Info: Deleting list items...";

    # List Name 
    $listName = 'Contacts'   
    Try {
        # Field Values
        
        # Get the list-item using CAML Query   
        $listItems = Get-PnPListItem -List  $listName -Query "<View><Query><Where><Eq><FieldRef Name='Contact_x0020_Number'/><Value Type='Text'>10101010</Value></Eq></Where></Query></View>"  
        Write-Host "Total items found: "  $listItems.count
        Write-Host ""
        foreach ($listItem in $listItems) {  

            TRY {
                Write-Host "Deleting item, ID: " $listItem["ID"]

                # Delete the list-item
                Remove-PnPListItem -List $listName -Identity $listItem -Force
        
                Write-Host -ForegroundColor Green "Item deleted"
            } 
            Catch {
                Write-Host -ForegroundColor Red "Error while deleting item with ID:" $listItem["ID"]
                Write-Host -ForegroundColor Red $_.Exception.Message
            }   
        }  
    }
    Catch {
        Write-Host -ForegroundColor Red "Error at DeleteListItems(): "$_.Exception.Message
    }   
}

function ConnectSPOnline() {
    # Get ConnectSPOnline node properties
    $Connectionconfig = $XmlDocument.Configurations.ConnectSPOnline;
    $siteUrl = $Connectionconfig.SiteUrl;
    $userName = $Connectionconfig.UserName;
    $password = $Connectionconfig.Password;

    $credentials;
    if ($userName -ne '' -and $password -ne '') {
        # Convert password to secure string    
        $secureStringPwd = ConvertTo-SecureString -AsPlainText $password -Force  

        # Create new credential object 
        $credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $userName, $secureStringPwd  
    }
    else {
        # If UserName & Password are blank in PSConfig file, get credential at run time 
        $credentials = (Get-Credential);
    }

    Try {
        write-host "Info: Connecting to site: " $siteURL  "..."

        # Connect to SP online site  
        Connect-PnPOnline -Url $siteURL -Credentials $credentials  
        
        Write-Host -ForegroundColor Green "Info: Connected !"
    }
    Catch {
        Write-Host -ForegroundColor Red "Error: Unable to connect !"
        Write-Host -ForegroundColor Red $_.Exception.Message
        Break
    }
     
}

function DisconnectOnline() {
    Disconnect-PnPOnline -Connection
}


function StartScript() {
    write-host -ForegroundColor Yellow "Info: Script execution started"   
    write-host ""

    # Connect to SPOnline 
    ConnectSPOnline;

    # 1- Create List Item 
    CreateListItems
    # 2- Read List Items
    ReadListItems;
    # 3- Update List Items
    UpdateListItems; 
    # 4- Delete List Items
    DeleteListItems;

    write-host ""
    write-host -ForegroundColor Yellow "Info: Execution Completed"   

    # Disconnect 
    Disconnect-PnPOnline 
}

<#***************************************************************************************#>
# Start point of the script

# Clear the screen
Clear-Host
# Create log file 
$date = Get-Date -format MMddyyyyHHmmss  
start-transcript -path .\Logs\Log_$date.doc   

# Start the script execution
StartScript;