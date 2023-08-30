# Create or Reset VMs

Helps with creating new VMs or resetting existing VMs. This is intended to be used with Microsoft Deployment Toolkit (MDT) to automate the creation of VMs that then PXE boot and follow the give MDT task sequence. MDT can then use the OS Optimization Tool and the MDT plugin to automate the installation of the agents and the optimization of Windows.

	* New VMs - Created using the given specification.
	* Existing VMs  - Any snapshots are deleted, the existing hard disk is deleted and a new hard disk is added to the VM.
	* MDT Database - Optionally add new VMs to the MDT database specifying a task sequence for MDT to use.
	* Power On - Optionally power on the VM after created or reset.
	* Remote Console - Open a vSphere remote console to the VM.
	* GUI that allows selection of VMs from a list in a comma separated file that contains the VM names and specifications.

## Usage
.\createresetvm.ps1 [vmlist.csv] [vCenterUser] [vCenterPassword] [MDTUser] [MDTPassword]

### Where
* vmlist.csv       = Comma delimited file with a VM per row. Fields are: Name,TaskSeq,Datastore,Network,Folder,Disk,Mem,vCPU,Displays,VideoMem,HWVersion,GuestId
* vCenterUser      = Username for vCenter Server.
* vCenterPassword  = Password for vCenter Server user.
* MDTUser          = Username with rights to the MDT SQL database.
* MDTPassword      = Password for MDT database user

### Examples
* .\createresetvm.ps1
* .\createresetvm.ps1 mylist.csv
* .\createresetvm.ps1 mylist.csv administrator@vsphere.local VMware1!
* .\createresetvm.ps1 mylist.csv administrator@vsphere.local VMware1! mdtuser sqlpassword

## Requirements
* Requires PowerCLI.
* Change variables for vCenter and the MDT SQL database.
* List of VMs in comma separated file.