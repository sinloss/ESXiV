param(
    [Parameter(Mandatory=$False)]
    [Alias("Bundle")]
    [ValidateScript({
        $ErrorView = "CategoryView"
        if (-Not ($_ | Test-Path -PathType Leaf)) {
            throw "It does not seem to be a valid path for a ESXi offline bundle file"
        }
        return $True
    })]
    [System.IO.FileInfo]$ESXiBundle,
	[switch]$Online = $False,
	[Parameter(Mandatory=$False)]
    [System.IO.FileInfo]$VMHome,
	[String]$Pswd = "1234560",
	[String]$VMName = "ESXi"
	
)

$principal = New-Object System.Security.Principal.WindowsPrincipal(
	[System.Security.Principal.WindowsIdentity]::GetCurrent()
)

# Check to see if we are currently running as an administrator
if ($principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $Host.UI.RawUI.WindowTitle = $MyInvocation.MyCommand.Definition + "(Elevated)"
} else {
	# Elevate
    $elev = New-Object System.Diagnostics.ProcessStartInfo "PowerShell"
    $elev.Arguments = "& '" + $script:MyInvocation.MyCommand.Path + "'"
    $elev.Verb = "runas"
    [System.Diagnostics.Process]::Start($elev) | Out-Null
    Exit
}

$wd = & Split-Path -Parent $MyInvocation.MyCommand.Definition

function CyanPut($head, $item, $tail, $fin = "") {
	Write-Host -NoNewline $head
	Write-Host -NoNewline -Foreground "Cyan" $item
	Write-Host -NoNewline $tail
	if ($fin -ne "--") {
		Write-Host -Foreground "Green" $fin
	}
}

function CountDown($i, $step = 1) {
	$len = ([String]$i).length
	$con = $host.UI.RawUI
	$y = $con.CursorPosition.Y
	$x = $con.CursorPosition.X
	do {
		Write-Host -NoNewLine -Foreground "Green" (([String]$i).PadLeft($len, '0'))
		[Console]::SetCursorPosition($x, $y)
		Sleep $step
		$i -= $step
	} while ($i -gt 0)
	Write-Host -Foreground "Green" "ok".PadLeft($len, '-')
}

# Create ISO file if needed
function CreateISO($bundleHome, $imageProfile) {
	[System.Net.WebRequest]::DefaultWebProxy = [System.Net.WebRequest]::GetSystemWebProxy()
	[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
		
	if (-not (Get-Module -ListAvailable -Name "VMware.PowerCLI" -ErrorAction SilentlyContinue)) {
		if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
			CyanPut "Could not install module " "VMware.PowerCLI " "witout administrator privileges"
			Exit
		}
		Install-Module -Name VMware.PowerCLI -Scope AllUsers -AllowClobber -Force
	}

	if (-not $?) {
		Exit
	}

	if ($Online) {
		CyanPut "Connecting the " "VMware ESXi Online Depot " "... " "--"
		Add-EsxSoftwareDepot "https://hostupdate.vmware.com/software/VUM/PRODUCTION/main/vmw-depot-index.xml" | Out-Null
		# It is the latest version currently known to work with Hyper-V legacy network drivers
		$esxiProfile = "ESXi-6.0.0-20170604001-standard"
		Write-Host -Foreground "Green" "done"
	} else {
		Add-EsxSoftwareDepot $ESXiBundle
		$esxiProfile = ((Get-EsxImageProfile "ESXi-6.0*" | Where {$_.Name -Match "\d+-standard" })[0]).Name
		if ($esxiProfile.Split("-")[2] -gt 20170604001) {
			Write-Host -NoNewline -Foreground "Red" "Invalid bundle: "
			Write-Host -NoNewline "The bundle version "
			Write-Host -NoNewline -Foreground "Cyan" "$esxiProfile "
			Write-Host "would not work with Hyper-V legacy "
			Write-Host "network adapters, it should be or older than ESXi-6.0.0-20170604001"
			Write-Host " "
			Exit
		}
	}
	
	# Download and add net-tulip
	$tulip = -join($bundleHome, "\Tulip-driver.zip")
	if (-not (Test-Path $tulip)) {
		CyanPut "Downloading the " "tulip driver " "... " "--"
		(New-Object System.Net.WebClient).DownloadFile("http://vibsdepot.v-front.de/depot/bundles/net-tulip-1.1.15-1-offline_bundle.zip", $tulip)
		if ($?) {
			Write-Host -Foreground "Green" "done"
		}
	}
	Add-EsxSoftwareDepot $tulip

	# Create image profile
	$ErrorActionPreference = "SilentlyContinue"
	New-EsxImageProfile -CloneProfile $esxiProfile -Name $imageProfile -Vendor "${env:COMPUTERNAME}" | Out-Null
	Add-EsxSoftwarePackage -ImageProfile $imageProfile -SoftwarePackage net-tulip -Force | Out-Null
	Set-EsxImageProfile -AcceptanceLevel CommunitySupported -ImageProfile $imageProfile | Out-Null
	$ErrorActionPreference = ""
	# Export
	$iso = (Join-Path $bundleHome "$imageProfile.iso")
	CyanPut "Exporting " "$iso " "... " "--"
	Export-EsxImageProfile -ImageProfile $imageProfile -FilePath $iso -ExportToIso -Force
	Write-Host -Foreground "Green" "done "
}

# Prepare the ISO file
function PrepareISO {
	$imageProfile = "$VMName-${env:COMPUTERNAME}"
	if ($Online) {
		$bundleHome = $wd
	} else {
		if (-not $ESXiBundle) {
			# Try to locate the fallback iso file around this very script
			$fallback = & Join-Path $wd "$imageProfile.iso"
			if (-not (Test-Path $fallback)) {
				# No fallback, no further
				return
			} else {
				$ESXiBundle = [System.IO.FileInfo]$fallback
			}
		}
		
		# Try to use the given existing iso file
		if ($ESXiBundle.Name.EndsWith(".iso", "CurrentCultureIgnoreCase")) {
			$script:iso = $ESXiBundle.FullName
			CyanPut "Use ISO file " $iso " "
			return
		}
		# Try to locate or create the iso file beside the given bundle
		$bundleHome = (Split-Path -Parent $ESXiBundle)
	}
	
	$script:iso = (Join-Path $bundleHome "$imageProfile.iso")
	if (Test-Path $iso) {
		CyanPut "ISO file " "$iso " "exists, " "skip"
	} else {
		# Create a new one
		CreateISO $bundleHome $imageProfile
	}

}

function CreateVM {
	if (-not (Hyper-V\Get-VM -Name $VMName -ErrorAction SilentlyContinue)) {
		if (-not $VMHome) {
			CyanPut "The " "VMHome " "is needed in order to create the VM"
			Exit
		} else {
			if (-not $ESXiBundle) {
				CyanPut "The " "ESXiBundle " "is needed in order to create the VM"
				Exit
			}
		}
		
		if (-Not ($VMHome | Test-Path -PathType Container)) {
			CyanPut "The given VM Home " "$VMHome " "seems to be invalid"
			Exit
		}
		
		if (Test-Path "$VMHome\$VMName.vhdx") {
			CyanPut "The " "$VMName.vhdx " "has already existed in the $VMHome, please choose another folder"
			Exit
		}
		
		$vSwitch = "$($VMName)Exterior"
		$nic = (Get-NetAdapter -Physical | Where { -not ($_.InterfaceDescription -Like '*Loopback*') } | Select -First 1)
		$existing = (Get-VMSwitch -SwitchType External)
		if (-not $existing) {
			# Create a new external VMSwitch
			Hyper-V\New-VMSwitch -Name $vSwitch -NetAdapterName $nic.Name -AllowManagementOS $true 
		} else {
			$vSwitch = $existing.Name
			CyanPut "The external VMSwitch " "$vSwitch " "exists, " "will use it"
		}
		
		# Mew VM
		Hyper-V\New-VM -Name $VMName -MemoryStartupBytes 4GB -NewVHDPath "$VMHome\$VMName.vhdx" -NewVHDSizeBytes 75GB -Path $VMHome -Generation 1 -BootDevice CD 
		# Mount the iso image
		Hyper-V\Set-VMDvdDrive -VMName $VMName -Path $iso
		# Add processor
		Hyper-V\Set-VMProcessor -VMName $VMName -Count 3
		# Remove default network adapter
		Hyper-V\Remove-VMNetworkAdapter -VMName $VMName -VMNetworkAdapterName "Network Adapter"
		# Add legacy network adapter
		Hyper-V\Add-VMNetworkAdapter -IsLegacy $true -VMName $VMName -Name "LegacyAdapter" -SwitchName $vSwitch
		# Enable MAC address spoofing
		Hyper-V\Set-VMNetworkAdapter -VMName $VMName -Name "LegacyAdapter" -MacAddressSpoofing On
		# Enable nested virtualization
		Hyper-V\Set-VMProcessor -VMName $VMName -ExposeVirtualizationExtensions $true
	}
}

function Combo([Int32[]]$keys) {
	$len = $keys.Length
	if ($len -lt 0) { return }
	if ($len -eq 1) {
		Sleep -Milliseconds 50
		$keyboard.TypeKey($keys[0]) | out-null
		return
	} else {
		Sleep -Milliseconds 50
		$keyboard.PressKey($keys[0]) | out-null
		Combo $keys[1..($len - 1)]
		Sleep -Milliseconds 50
		$keyboard.ReleaseKey($keys[0]) | out-null
	}
}

function Stroke($key, $sleep = 150) {
	Sleep -Milliseconds $sleep
	Switch ($key.GetType().Name) {
		"Int32" { $keyboard.TypeKey($key) | out-null; break }
		"String" { $keyboard.TypeText($key) | out-null; break }
		"Object[]" { Combo $key; break }
	}
	
}

function ApplyInstallation {
	if (-not (Hyper-V\Get-VMDvdDrive -VMName $VMName | where {$_.Path -eq $iso})) {
		return
	}
	
	$VMCS = Get-WmiObject -Namespace root\virtualization\v2 -Class Msvm_ComputerSystem -Filter "ElementName='$($VMName)'"
	$keyboard = $VMCS.GetRelated("Msvm_Keyboard")
	
	Stroke 0x09 3900                      #Tab:       edit boot options
	Stroke " ignoreHeadless=TRUE"         #           edit
	Stroke 0x0D                           #Enter:     apply the editted options
	Stroke 0x0D                           #Enter:     boot

	Write-Host " "
	Write-Warning "After the installer is compeltely loaded, CONTINUE ME!!!" -WarningAction Inquire
	if (!$?) { Exit }
	Write-Host -Foreground "Green" "Please wait for the intallation to be started, you don't need to do anything except waiting"

	Stroke 0x0D                           #Enter:     answer to the welcome message
	Stroke 0x7A                           #F11:       accept the EULA
	Write-Host -NoNewLine -Foreground "Gray" "* wait for the disk scanning to be compeleted - "
	CountDown 15
	$keyboard.TypeKey(0x0D) | out-null    #Enter:     use the default disk
	Stroke 0x0D                           #Enter:     use the default keyboard layout
	Stroke $Pswd                          #           password
	Stroke 0x09 1500                      #Tab:       next line
	Stroke $Pswd                          #           confirm password
	Stroke 0x0D                           #Enter:     continue
	Stroke 0x0D                           #Enter:     continue if something poped up
	Write-Host -NoNewLine -Foreground "Gray" "* wait for the pre-installation scanning to be compeleted - "
	CountDown 15
	$keyboard.TypeKey(0x7A) | out-null    #F11:       install

	Write-Host " "
	Write-Warning "After the installation is compeleted, CONTINUE ME!!!" -WarningAction Inquire
	if (!$?) { Exit }
	Write-Host -Foreground "Green" "Please wait for the rebooting to be compeleted, you don't need to do anything except waiting"

	Stroke 0x0D                           #Enter:     reboot
	# We might not be able to figure out a perfect timing to perform the SHIFT + O trick.
	# So we just shut it down, and restart the VM under our control
	$conn.Kill()
	Hyper-V\Stop-VM -Name $VMName -Force
	# Eject the iso image
	Hyper-V\Set-VMDvdDrive -VMName $VMName -Path $null
	Write-Host -NoNewLine -Foreground "DarkGreen" "* In a moment, let me take a little breath - "
	CountDown 3
	# Start another clean process to do the following configurations
    $proc = New-Object System.Diagnostics.ProcessStartInfo "PowerShell"
    $proc.Arguments = "& '" + $script:MyInvocation.MyCommand.Path + "'"
    $proc.Verb = "runas"
    [System.Diagnostics.Process]::Start($proc) | Out-Null
    Exit
}

$iso = ""
PrepareISO

CreateVM

Hyper-V\Start-VM -Name $VMName

$conn = Start-Process VMConnect -ArgumentList $env:COMPUTERNAME,$VMName -PassThru

ApplyInstallation

$VMCS = Get-WmiObject -Namespace root\virtualization\v2 -Class Msvm_ComputerSystem -Filter "ElementName='$($VMName)'"
$keyboard = $VMCS.GetRelated("Msvm_Keyboard")

Stroke @(0x10, 0x4F) 3500       #Shift + O: try the first time
Stroke 0x1B
Stroke @(0x10, 0x4F) 3500       #Shift + O: try the second time
Stroke 0x1B
Stroke @(0x10, 0x4F) 3500       #Shift + O: try the last time
Stroke " ignoreHeadless=TRUE"   #           edit
Stroke 0x0D                     #           apply the editted options

Write-Host " "
Write-Warning "After the starting is compeleted, CONTINUE ME!!!" -WarningAction Inquire
if (!$?) { Exit }
Write-Host -Foreground "Green" "Please wait for the configuration to be compeleted, you don't need to do anything except waiting"
Stroke 0x71                     #F2:        configur
Stroke 0x09                     #Tab:       next line
Stroke $Pswd                    #           password
Stroke 0x0D  

Stroke 0x28 5000                #           ↓
Stroke 0x28                     #           ↓
Stroke 0x28                     #           ↓
Stroke 0x28                     #           ↓
Stroke 0x28                     #           ↓
Stroke 0x28                     #           ↓
Stroke 0x0D                     #Enter:     enter Troubleshooting Options
Stroke 0x0D                     #Enter:     enable ESXi Shell
Stroke 0x28 1500                #           ↓
Stroke 0x0D                     #Enter:     enable ESXi SSH
Stroke @(0x12, 0x70) 300        #Alt + F1:  enter ESXi Shell
Stroke "root" 1500              #           username
Stroke 0x0D 
Stroke $Pswd                    #           password
Stroke 0x0D 

# execute 'esxcfg-advcfg -k TRUE ignoreHeadless' on the shell and exit
Stroke "esxcfg-advcfg -k TRUE ignoreHeadless"
Stroke 0x0D
Stroke "exit"
Stroke 0x0D

Stroke @(0x12, 0x71) 300        #Alt + F2:  leave ESXi Shell
						     
Stroke 0x1B 1500                #ESC:       exit to the customizing screen
Stroke 0x1B 1500                #ESC:       exit to the main screen
Stroke 0x7B 1500                #F12:       restart
Stroke 0x09                     #Tab:       next line
Stroke $Pswd                    #           password
Stroke 0x0D
Stroke 0x7A                     #F11:       restart

Write-Host " "
Write-Host -Foreground "Green" "Congratulations! The ESXi has been successfully installed and configured on your Hyper-V!"
Write-Host -Foreground "Green" "Just enjoy after it is restarted"
Write-Host " "
Read-Host -Prompt "Press Enter to exit"