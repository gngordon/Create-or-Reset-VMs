<#
.SYNOPSIS
Script to select VMs to create or reset, and then add them to Microsoft Deployment Toolkit, power on and open console.
Requires PowerCLI.
Settings.ini = Contains the variables for vCenter and the MDT SQL database. Change as required.
List of VMs in comma separated file. (defaults to vmlist.csv)

.USAGE
     .\createresetvm.ps1 [vmlist.csv] [vCenterUser] [vCenterPassword] [MDTUser] [MDTPassword]
     
     WHERE
         vmlist.csv       = Comma delimited file with a VM per row. Fields are: Name,TaskSeq,Datastore,Network,Folder,Disk,Mem,vCPU,Displays,VideoMem,HWVersion,GuestId,vGPU
         vCenterUser      = Username for vCenter Server.
         vCenterPassword  = Password for vCenter Server user.
         MDTUser          = Username with rights to the MDT SQL database.
         MDTPassword      = Password for MDT database user

.EXAMPLES
     .\createresetvm.ps1
     .\createresetvm.ps1 mylist.csv
     .\createresetvm.ps1 mylist.csv administrator@vsphere.local VMware1!
     .\createresetvm.ps1 mylist.csv administrator@vsphere.local VMware1! mdtuser sqlpassword
       
.ACTIONS
    *Select VMS
    *Select actions
    *Create - New VM using spec defined
    *Reset existing VM - Power Off, remove any snapshots, delete hard disk from VMs, add new hard disk
    *Optionally add computer object to Microsoft Deployment Toolkit (MDT) database
    *Optionally power on VM 
    *Optionally open remote console
 
 .NOTES
    Version:        2.2
    Author:         Graeme Gordon - ggordon@vmware.com
    Creation Date:  2023/03/15
    Purpose/Change: Create or reset virtual machines
  
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
    VMWARE,INC. BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
    IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
    CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 #>

param([string]$vmListFile = �vmlist.csv�, [string] $vCenterUser, [string] $vCenterPassword, [string] $SQLUser, [string] $SQLPassword )

#region variables
################################################################################
#                                    Variables                                 #
################################################################################
$SettingsFile			= "settings.ini"

$global:SQLConnected    = $False
$global:vcConnected		= $false
$MacAddress             = "00:00:00:00:00:00"
#endregion variables

function ImportIni
{
################################################################################
# Function to parse token values from a .ini configuration file                #
################################################################################
	param ($file)

	$ini = @{}
	switch -regex -file $file
	{
            "^\s*#" {
                continue
            }
    		"^\[(.+)\]$" {
        		$section = $matches[1]
        		$ini[$section] = @{}
    		}
    		"([A-Za-z0-9#_]+)=(.+)" {
        		$name,$value = $matches[1..2]
        		$ini[$section][$name] = $value.Trim()
    		}
	}
	$ini
}

function Initialize_Env ($vCenterServer)
{
################################################################################
#               Function Initialize_Env                                        #
################################################################################
    # --- Initialize PowerCLI Modules ---
    #Get-Module -ListAvailable VMware* | Import-Module
    Import-Module VMware.VimAutomation.Core
	Set-PowerCLIConfiguration -Scope User -ParticipateInCeip $false -InvalidCertificateAction ignore -DefaultVIServerMode Multiple -Confirm:$false

    # --- Connect to the vCenter server ---
	$attempt = 0
    Do {
        Write-Output "", "Connecting To vCenter Server:"
		Write-Host ("Connecting To vCenter Server: " + $vCenterServer) -ForegroundColor Yellow
        If (!$vCenterUser)
        {
            $vc = Connect-VIServer -Server $vCenterServer -ErrorAction SilentlyContinue
        }
        elseif (!$vCenterPassword)
        {
            $vc = Connect-VIServer -Server $vCenterServer -User $vCenterUser -ErrorAction SilentlyContinue
        }
        else
        {
             $vc = Connect-VIServer -Server $vCenterServer -User $vCenterUser -Password $vCenterPassword -Force -ErrorAction SilentlyContinue
        }
        If (!$vc.IsConnected)
		{
			$attempt += 1
			Write-Host ("Failed to connect to vCenter Server. Attempt " + $attempt + " of 3")  -ForegroundColor Red
		}
    } Until ($vc.IsConnected -or $attempt -ge 3)
	If ($vc.IsConnected) { $global:vcConnected = $true }
}

function CreateVM ($vm, $scsiControllerType, $StorageFormat)
{
################################################################################
#               Function CreateVM                                              #
################################################################################
    #Determine if the portgroup is on a distributed or standard switch
    $pg = Get-VirtualPortGroup -Name $vm.Network
	
	#Calculate sockets per core
	$vCPU = [int]$vm.vCPU
	$rem = $vCPU % 2
	if ($rem -eq 0) #even
		{ $corespersocket = (0.5 * $vCPU) }
	else #odd
		{ $corespersocket = $vCPU }
	If (($vm.GuestId -like '*srv*') -or ($vm.GuestId -like '*Server*')) { $corespersocket = 1 } #Server OS so set cores per socket to 1
	
	#Create VM
	If ($pg.ExtensionData.Config) #Portgroup is on Distributed virtual switch
    {
		New-VM -Name $vm.Name -ResourcePool $ResourcePool -HardwareVersion $vm.HWVersion -GuestId $vm.GuestId -DiskGB $vm.Disk -DiskStorageFormat $StorageFormat -NumCpu $vCPU -CoresPerSocket $corespersocket -MemoryGB $vm.Mem -Datastore $vm.Datastore -Location $vm.Folder -Portgroup $pg
    }
    else #Portgroup is on Standard virtual switch
    {
		New-VM -Name $vm.Name -ResourcePool $ResourcePool -HardwareVersion $vm.HWVersion -GuestId $vm.GuestId -DiskGB $vm.Disk -DiskStorageFormat $StorageFormat -NumCpu $vCPU -CoresPerSocket $corespersocket -MemoryGB $vm.Mem -Datastore $vm.Datastore -Location $vm.Folder -NetworkName $vm.Network
    }
	$vmobj = Get-VM -Name $vm.Name
  
    #Reserve Memory
    Write-Host ("Reserving Memory") -ForegroundColor Yellow
    $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
    $spec.memoryReservationLockedToMax = $true
    $vmobj.ExtensionData.ReconfigVM_Task($spec)

    #Change scsi Controller to VMware Para Virtual
    Write-Host ("Setting SCSI Controller to " + $scsiControllerType) -ForegroundColor Yellow
    $scsiController = Get-HardDisk -VM $vm.Name | Select -First 1 | Get-ScsiController
    Set-ScsiController -ScsiController $scsiController -Type $scsiControllerType

    #Change Network Adapter to vmxnet3
    Write-Host ("Set Network Adapter to vmxnet3") -ForegroundColor Yellow
    $vmobj | Get-NetworkAdapter | Set-NetworkAdapter -Type Vmxnet3 -Confirm:$false

    #Set Video Displays and Memory and deselect Secure Boot
    Write-Host ("Configure Video to " + $vm.Displays + " displays and " + $vm.VideoMem + " video RAM") -ForegroundColor Yellow
    $vmobj | %{
        $vid = $_.ExtensionData.Config.Hardware.Device | ?{$_.GetType().Name -eq "VirtualMachineVideoCard"}
        $spec = New-Object VMware.Vim.VirtualMachineConfigSpec
        $devChange = New-Object VMware.Vim.VirtualDeviceConfigSpec
        $devChange.Operation = 'edit'
        $vid.NumDisplays = $vm.Displays
        $vid.VideoRamSizeInKB = $vm.VideoMem/1KB
        $devChange.Device += $vid
        $spec.DeviceChange += $devChange

        #Deselect Secure Boot 
        $spec.Firmware = [VMware.Vim.GuestOsDescriptorFirmwareType]::efi
        $boot = New-Object VMware.Vim.VirtualMachineBootOptions
        $boot.EfiSecureBootEnabled = $false
        $spec.BootOptions = $boot

		Write-Host ("vGPU VM value: " + $vm.vGPU) -ForegroundColor Yellow
		#Add vGPU if specified
		if ($vm.vGPU -ne "false" -and $vm.vGPU -ne $null)
		{
			Write-Host ("Add vGPU with profile: " + $vm.vGPU) -ForegroundColor Yellow
			$spec.deviceChange = New-Object VMware.Vim.VirtualDeviceConfigSpec[] (1)
			$spec.deviceChange[0] = New-Object VMware.Vim.VirtualDeviceConfigSpec
			$spec.deviceChange[0].operation = 'add'
			$spec.deviceChange[0].device = New-Object VMware.Vim.VirtualPCIPassthrough
			$spec.deviceChange[0].device.deviceInfo = New-Object VMware.Vim.Description
			$spec.deviceChange[0].device.deviceInfo.summary = ''
			$spec.deviceChange[0].device.deviceInfo.label = 'New PCI device'
			$spec.deviceChange[0].device.backing = New-Object VMware.Vim.VirtualPCIPassthroughVmiopBackingInfo
			$spec.deviceChange[0].device.backing.vgpu = $vm.vGPU
		}
        $_.ExtensionData.ReconfigVM($spec)
    }

    #Set Advanced Configuration Parameters
    Write-Host ("Set Advanced Configuration Parameters: devices.hotplug = false") -ForegroundColor Yellow
    $vmobj | New-AdvancedSetting -Name devices.hotplug -Value FALSE -Confirm:$False
}

function ResetVM ($vm, $scsiControllerType, $StorageFormat )
{
################################################################################
#               Function ResetVM                                               #
################################################################################
    $vmobj = Get-VM -Name $vm.Name
    if ($vmobj.powerstate -eq "PoweredOn") {Stop-VM -vm $vmobj -Confirm:$false} #if VM is powered on, power off
        
    #Delete any snaphots
    $snaps = Get-Snapshot -VM $vmobj
    ForEach ($snap in $snaps)
    {
        Write-Host ("Deleting snapshot") -ForegroundColor Yellow
        Remove-Snapshot -Snapshot $snap -RemoveChildren -Confirm:$false
    }

    #Delete the old hard disk and create a new hard disk
    Write-Host ("Deleting old hard disk") -ForegroundColor Yellow
    Get-HardDisk -vm $vmobj | Remove-HardDisk -DeletePermanently:$true -Confirm:$false #Remove old hard disk
    
    Write-Host ("Adding new hard disk") -ForegroundColor Yellow
    New-HardDisk -VM $vmobj -CapacityGB $vm.Disk -StorageFormat $StorageFormat | New-ScsiController -Type $scsiControllerType #Add new hard disk
}

function Connect-MDTDatabase ($server, $port, $db, $user, $password) 
{
################################################################################
#               Function Connect-MDTDatabase                                   #
################################################################################
    Do {     
        If (!$user -And !$SQLIntegratedAuth) #Need username and password or last  attempt failed.
        {
            if($SQLCred = $host.ui.PromptForCredential("SQL credentials", "Enter credentials for MDT database","", "")){}else{return}
            $user = $SQLCred.UserName
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SQLCred.Password) #Convert Password
            $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        }
        
        #Form connection string and open a connection
        If ($SQLIntegratedAuth)
        {
            $connString = "Server = $server, $port; Database = $db; Integrated Security = True"
        }
        else
        {
            $connString = "Server = $server, $port; Database = $db; Integrated Security = False; User ID=$user; Password=$password"
        }
		$global:mdtSQLConnection = New-Object System.Data.SqlClient.SqlConnection
        $mdtSQLConnection.ConnectionString = $connString
        $mdtSQLConnection.Open()
        
        #Check to see if we managed to connect
        If ($mdtSQLConnection.State -eq "Open")
        {
            $global:SQLConnected = $True
        }
        else
        {
            Write-Host ("Failed to connect to MDT SQL Server. Let's try that again.")  -ForegroundColor Red
            $global:SQLConnected = $False
            $user = $null
        }
    } Until ($global:SQLConnected)
    Write-Host ("Connected to MDT SQL Server.")  -ForegroundColor Green
}

function Get-MDTComputer ( $macAddress )
{
################################################################################
#               Function Get-MDTComputer                                       #
################################################################################
    # Specified the initial command
    $sql = "SELECT * FROM ComputerSettings WHERE "

    if ($macAddress -ne "")
    {
        $sql = "$sql MacAddress='$macAddress'"
    }
    
    $selectAdapter = New-Object System.Data.SqlClient.SqlDataAdapter($sql, $mdtSQLConnection)
    $selectDataset = New-Object System.Data.Dataset
    $null = $selectAdapter.Fill($selectDataset, "ComputerSettings")
    $selectDataset.Tables[0].Rows
}

function New-MDTComputer
{
################################################################################
#               Function New-MDTComputer                                       #
################################################################################
    [CmdletBinding()]
    PARAM
    (
        [Parameter(ValueFromPipelineByPropertyName=$true)] $MacAddress,
        [Parameter(ValueFromPipelineByPropertyName=$true)] $Description,
        [Parameter(ValueFromPipelineByPropertyName=$true, Mandatory=$true)] $Settings
    ) 
 
    # Insert a new computer row into the ComputerIdentity table
    $sql = "INSERT INTO ComputerIdentity (Description, MacAddress) VALUES ('$Description', '$MacAddress') SELECT @@IDENTITY"
        
    Write-Verbose "About to execute command: $sql"
    $identityCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
    $identity = $identityCmd.ExecuteScalar()
        
    Write-Verbose "Added computer identity record"
    
    # Insert the settings row, adding the values as specified in the hash table
    $settingsColumns = $Settings.Keys -join ","
    $settingsValues = $Settings.Values -join "','"
    $sql = "INSERT INTO Settings (Type, ID, $settingsColumns) VALUES ('C', $identity, '$settingsValues')"
        
    Write-Verbose "About to execute command: $sql"
    $settingsCmd = New-Object System.Data.SqlClient.SqlCommand($sql, $mdtSQLConnection)
    $null = $settingsCmd.ExecuteScalar()
}

function Define_GUI
{
################################################################################
#              Function Define_GUI                                             #
################################################################################
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $global:form                     = New-Object System.Windows.Forms.Form
    $form.Text                       = 'Create or Reset Virtual Machines'
    $form.Size                       = New-Object System.Drawing.Size(500,400)
    #$form.Autosize                   = $true
    $form.StartPosition              = 'CenterScreen'
    $form.Topmost                    = $true

    #OK button
    $OKButton                        = New-Object System.Windows.Forms.Button
    $OKButton.Location               = New-Object System.Drawing.Point(300,320)
    $OKButton.Size                   = New-Object System.Drawing.Size(75,23)
    $OKButton.Text                   = 'OK'
    $OKButton.DialogResult           = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton               = $OKButton
    $form.Controls.Add($OKButton)

    #Cancel button
    $CancelButton                    = New-Object System.Windows.Forms.Button
    $CancelButton.Location           = New-Object System.Drawing.Point(400,320)
    $CancelButton.Size               = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text               = 'Cancel'
    $CancelButton.DialogResult       = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    #Checkbox to allow pause for adding applicatons to MDT computer
    $global:AppsSelect               = New-Object System.Windows.Forms.CheckBox
    $AppsSelect.Location             = New-Object System.Drawing.Point(300,40)
    $AppsSelect.Size                 = New-Object System.Drawing.Size(200,23)
    $AppsSelect.Text                 = 'Pause to add apps to MDT'    
    $AppsSelect.Checked              = $PauseforApps
    $AppsSelect.Add_CheckStateChanged({ $global:PauseforApps = $AppsSelect.Checked })
    $form.Controls.Add($AppsSelect)

    #Checkbox on whether to try adding to MDT
    $global:MDTSelect                = New-Object System.Windows.Forms.CheckBox
    $MDTSelect.Location              = New-Object System.Drawing.Point(300,80)
    $MDTSelect.Size                  = New-Object System.Drawing.Size(200,23)
    $MDTSelect.Text                  = 'Add to MDT database'    
    $MDTSelect.Checked               = $AddtoMDT
    $MDTSelect.Add_CheckStateChanged({ $global:AddtoMDT = $MDTSelect.Checked })
    $form.Controls.Add($MDTSelect) 

    #Checkbox on whether to power on VM
    $global:PowerSelect              = New-Object System.Windows.Forms.CheckBox
    $PowerSelect.Location            = New-Object System.Drawing.Point(300,120)
    $PowerSelect.Size                = New-Object System.Drawing.Size(200,23)
    $PowerSelect.Text                = 'Power on'    
    $PowerSelect.Checked             = $StartVM
    $PowerSelect.Add_CheckStateChanged({ $global:StartVM = $PowerSelect.Checked })
    $form.Controls.Add($PowerSelect)

    #Checkbox on whether to open remote console to VM
    $global:ConsoleSelect            = New-Object System.Windows.Forms.CheckBox
    $ConsoleSelect.Location          = New-Object System.Drawing.Point(300,160)
    $ConsoleSelect.Size              = New-Object System.Drawing.Size(200,23)
    $ConsoleSelect.Text              = 'Remote console'    
    $ConsoleSelect.Checked           = $OpenConsole 
    $ConsoleSelect.Add_CheckStateChanged({ $global:OpenConsole = $ConsoleSelect.Checked })
    $form.Controls.Add($ConsoleSelect)

    #Checkbox on whether to demo actions
    $global:DemoSelect                = New-Object System.Windows.Forms.CheckBox
    $DemoSelect.Location              = New-Object System.Drawing.Point(300,200)
    $DemoSelect.Size                  = New-Object System.Drawing.Size(200,23)
    $DemoSelect.Text                  = 'Demo'    
    $DemoSelect.Checked               = $Demo
    $DemoSelect.Add_CheckStateChanged({ $global:Demo = $DemoSelect.Checked })
    $form.Controls.Add($DemoSelect) 

    #Text above list box of VMs
    $label                            = New-Object System.Windows.Forms.Label
    $label.Location                   = New-Object System.Drawing.Point(10,20)
    $label.Size                       = New-Object System.Drawing.Size(280,20)
    $label.Text                       = 'Select VMs from the list below:'
    $form.Controls.Add($label)

    #List box for selection of VMs
    $global:listBox                   = New-Object System.Windows.Forms.Listbox
    $listBox.Location                 = New-Object System.Drawing.Point(10,40)
    $listBox.Size                     = New-Object System.Drawing.Size(260,250)
    $listBox.SelectionMode            = 'MultiExtended'
    ForEach ($vm in $vmlist)
    {
        [void] $listBox.Items.Add($vm.Name)
    }
    $listBox.Height = 250
    $form.Controls.Add($listBox)  
}

#region main
################################################################################
#              Main
################################################################################
Clear-Host

#Check the settings file exists
if (!(Test-path $SettingsFile)) {
	WriteErrorString "Error: Configuration file ($SettingsFile) not found."
	Exit
}
#Import settings variables
$global:vars = ImportIni $SettingsFile
If ($vars.Controls.Demo -like "No") { $global:Demo = $False } Else { $global:Demo = $True }
If ($vars.Controls.AddtoMDT -like "Yes") { $global:AddtoMDT = $True } Else { $global:AddtoMDT = $False }
If ($vars.Controls.PauseforApps -like "Yes") { $global:PauseforApps = $True } Else { $global:PauseforApps = $False }
If ($vars.Controls.StartVM -like "Yes") { $global:StartVM = $True } Else { $global:StartVM = $False }
If ($vars.Controls.OpenConsole -like "Yes") { $global:OpenConsole = $True } Else { $global:OpenConsole = $False }
If ($vars.SQL.SQLIntegratedAuth -like "Yes") { $global:SQLIntegratedAuth = $True } Else { $global:SQLIntegratedAuth = $False }

#Check the VM list file exists
if (!(Test-path $vmListFile)) {
	WriteErrorString "Error: VM list file ($vmListFile) not found."
	Exit
}
$global:vmlist = Import-Csv $vmListFile #Import the list of VMs

Define_GUI
$result = $form.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    #Write-Host ("Button Pressed") -ForegroundColor Green
    $selection = $listBox.SelectedItems   


    If ($selection)
    {
        Write-Host ("Selected VMs   : " + $selection) -ForegroundColor Yellow
        Write-Host ("Pause for Apps : " + $PauseforApps) -ForegroundColor Green
        Write-Host ("Add to MDT     : " + $AddtoMDT) -ForegroundColor Green
        Write-Host ("Start VM       : " + $StartVM) -ForegroundColor Green
        Write-Host ("Open Console   : " + $OpenConsole) -ForegroundColor Green
        Write-Host ("Demo           : " + $Demo) -ForegroundColor Green

		Initialize_Env $vars.vSphere.vCenterServer
		If (!$vcConnected) { Exit }
        Add-Type -AssemblyName 'PresentationFramework'
    
        If ($vars.vSphere.ResourcePoolName -ne $vars.vSphere.ClusterName) { $ResourcePool = Get-ResourcePool -Name $vars.vSphere.ResourcePoolName }
        Else { $ResourcePool = Get-Cluster -Name $vars.vSphere.ClusterName }
                
        ForEach ($vm in $vmlist)
        {
            If ($selection.Contains($vm.Name))
            {
                $VMExists = get-vm -name $vm.name -ErrorAction SilentlyContinue
                If ($VMExists)
                {
                    #VM already exists, so reset the VM and replace its hard drive with a new blank disk
                    Write-Host ("Resetting VM: " + $vm.Name) -ForegroundColor Yellow
                    If (!$Demo) { ResetVM $vm $vars.vSphere.scsiControllerType $vars.vSphere.StorageFormat }
                }
                Else
                { 
                    #Create a new VM
                    Write-Host ("Creating new VM: " + $vm.Name) -ForegroundColor Yellow
					If (!$Demo) { CreateVM -VM $vm $vars.vSphere.scsiControllerType $vars.vSphere.StorageFormat }
                } 
                            
                If ($AddtoMDT -and !$VMExists)
                {
                    If (!$SQLConnected) { Connect-MDTDatabase $vars.SQL.SQLServer $vars.SQL.SQLPort $vars.SQL.SQLDatabase $SQLUser $SQLPassword } #Connect to MDT Database
                    If (!$Demo)
                    {
                        #Get the MAC address of the VMs nework adapter so that it can be added to the computer object in the MDT database.
						$nic = Get-NetworkAdapter -vm $vm.name
                        $MacAddress = ($nic.MacAddress).ToUpper()
                    }
                    If ($SQLConnected)
                    {
                        $mdtentry = Get-MDTComputer $MacAddress #Check to see if there is already an entry in the MDT databse for this VM
                        If (-Not $mdtentry)
                        {
                            #Add a new Computer entry to the MDT database
                            Write-Host ("Add new entry to MDT database: " + $vm.Name + "; " + $MacAddress) -ForegroundColor Yellow
                            If (!$Demo) { New-MDTComputer -Description $vm.Name -MacAddress $MacAddress -Settings @{TaskSequenceID=$vm.TaskSeq; OSInstall="YES"; SkipApplications="yes"; SkipTaskSequence="yes"} -verbose }
                        }
                    }
                }

                If ($PauseforApps)
                {
                    #Prompt to add applications to computer in MDT
                    [void] [System.Windows.MessageBox]::Show( "Add applications to the computer in MDT. Press OK when complete.", $vm.Name, "OK", "Information" )
                }
                    
                If ($StartVM)
                {
                    #Power on VM
                    Write-Host ("Power on VM: " + $vm.Name) -ForegroundColor Yellow
                    If (!$Demo) { Get-VM -Name $vm.Name | Start-VM -Confirm:$false }
        
                    If ($OpenConsole)
                    {
                        #Open Remote Console to VM
                        Write-Host ("Open Remote Console to VM: " + $vm.Name) -ForegroundColor Yellow
                        If (!$Demo)  { Get-VM -Name $vm.Name | Open-VMConsoleWindow }
                    }
                }
            }
        }
    }
    else
    {
        Write-Host ("No VMs Selected") -ForegroundColor Yellow
    }
}
else
{
    #Write-Host ("Cancel Button Pressed") -ForegroundColor Red
}
#endregion main