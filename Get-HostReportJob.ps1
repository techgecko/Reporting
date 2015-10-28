<#
.SYNOPSIS
    Retrieves hosts from VMware vCenter(s) and exports to Excel & CSV files.

.DESCRIPTION
    Retrieves hosts from VMware vCenter and exports details to Excel & CSV files. Parallel threads are used to increase performance.

.PARAMETER vCenter
    An FQDN or hostname for a vCenter

.PARAMETER -test
    Append "-test" to create a timestamped file, otherwise it will be named _Hosts_Current.xlsx & .csv

.EXAMPLE
    .\Get-HostReportJob.ps1 vCenter1,vCenter2,vCenter3
       This returns all hosts from all entered comma-separated vCenters.

.EXAMPLE
    .\Get-HostReportJob.ps1 vCenter1.company.com
       This returns all hosts from a single vCenter "vCenter1.company.com".

.EXAMPLE
    .\Get-HostReportJob.ps1 vCenter1.company.com -test
       This returns all hosts from a single vCenter "vCenter1.company.com" and names the files with a timestamp.

.EXAMPLE
    .\Get-HostReportJob.ps1
       You'll be prompted to enter one vCenter per line, followed by a blank line.
	   This returns hosts from all entered vCenters.

.NOTES
    Author: Greg Hatch, greg.hatch@raymondjames.com
	10/28/2015
	
	Output defaults to C:\Temp... edit the script to change the ExportPath
	
#>


param (
	[Parameter(Mandatory=$true, HelpMessage="List of vCenter servers")]
	[ValidateNotNull()]
	[string[]]
	$vCenters,
	[switch]$Test
)

if(${env:ComputerName} -eq "YourScriptServerName"){
	$ExportPath = "F:\Reports"
	$MaxThreads = 9
}
else{
	$ExportPath = "C:\Temp"
	$MaxThreads = 7
}

$AdminReportDay = "Sunday"	# Timestamped weekly report created this day
$DateNow = Get-Date
$VCCounter = 0
$HostNICs = @()
$NumVCenters = $vCenters.count
$TempPath = "C:\Temp"


Get-Date -format yyyy\-MM\-dd\-HH:mm:ss

function Export-Xlsx {
	<#
		.SYNOPSIS
			Exports data to an Excel workbook, http://www.lucd.info/2013/01/03/export-xls-the-sequel-and-ordered-data/
		.DESCRIPTION
			Exports data to an Excel workbook and applies cosmetics. Optionally add a title, autofilter, autofit and a chart.
			Allows for export to .xls and .xlsx format. If .xlsx is specified but not available (Excel 2003) the data will be exported to .xls.
		.PARAMETER InputData
			The data to be exported to Excel
		.PARAMETER Path
			The path of the Excel file. Defaults to %HomeDrive%\Export.xlsx.
		.PARAMETER WorksheetName
			The name of the worksheet. Defaults to filename in $Path without extension.
		.PARAMETER ChartType
			Name of an Excel chart to be added.
		.PARAMETER Title
			Adds a title to the worksheet.
		.PARAMETER SheetPosition
			Adds the worksheet either to the 'begin' or 'end' of the Excel file.
			This parameter is ignored when creating a new Excel file.
		.PARAMETER FreezeRow
			Freeze the row # when scrolling up/down -- 1 = first row
		.PARAMETER FreezeColumn
			Freeze the column # when scrolling right/left
		.PARAMETER ChartOnNewSheet
			Adds a chart to a new worksheet instead of to the worksheet containing data.
			The Chart will be placed after the sheet containing data.
			Only works when parameter ChartType is used.
		.PARAMETER AppendWorksheet
			Appends a worksheet to an existing Excel file.
			This parameter is ignored when creating a new Excel file.
		.PARAMETER Borders
			Adds borders to all cells. Defaults to True.
		.PARAMETER HeaderColor
			Applies background color to the header row. Defaults to True.
		.PARAMETER GreenBar
			Adds 'greenbar' shading every 3 rows. Defaults to True
		.PARAMETER WrapText
			Wraps text in cells. Defaults to False
		.PARAMETER AutoFit
			Apply autofit to columns. Defaults to True.
		.PARAMETER AutoFilter
			Apply autofilter. Defaults to True.
		.PARAMETER PassThrough
			When enabled returns file object of the generated file.
		.PARAMETER Force
			Overwrites existing Excel sheet. When this switch is not used but the Excel file already exists, a new file with datestamp will be generated.
			This switch is ignored when using the AppendWorksheet switch.
		.EXAMPLE
			Get-Process | Export-Xlsx D:\Data\ProcessList.xlsx
			Exports a list of running processes to Excel
		.EXAMPLE
			Get-ADuser -Filter {enabled -ne $True} | Select-Object Name,Surname,GivenName,DistinguishedName | Export-Xlsx -Path 'D:\Data\Disabled Users.xlsx' -Title 'Disabled users of Contoso.com'
			Export all disabled AD users to Excel with optional title
		.EXAMPLE
			Get-Process | Sort-Object CPU -Descending | Export-Xlsx -Path D:\Data\Processes_by_CPU.xlsx
			Export a sorted processlist to Excel
		.EXAMPLE
			Export-Xlsx (Get-Process) -AutoFilter:$False -PassThrough | Invoke-Item
			Export a processlist to %HomeDrive%\Export.xlsx with AutoFilter disabled, and open the Excel file
		.NOTES
			Shamelessly borrowed ideas and code from http://www.lucd.info/2010/05/29/beyond-export-csv-export-xls/
			4/15/14 Greg Hatch: Added GreenBar and FreezeHeader features
	#>
	[CmdletBinding()]
	param (
		[Parameter(Position=0,Mandatory=$True,ValueFromPipeline=$True)]
		[ValidateNotNullOrEmpty()]
		$InputData,
		[Parameter(Position=1)]
		[ValidateScript({
			$ReqExt = [System.IO.Path]::GetExtension($_)
			($ReqExt -eq ".xls") -or
			($ReqExt -eq ".xlsx")
		})]
		$Path=(Join-Path $env:HomeDrive "Export.xlsx"),
		[Parameter(Position=2)] $WorksheetName = [System.IO.Path]::GetFileNameWithoutExtension($Path),
		[Parameter(Position=3)]
		[ValidateSet("xl3DArea","xl3DAreaStacked","xl3DAreaStacked100","xl3DBarClustered",
			"xl3DBarStacked","xl3DBarStacked100","xl3DColumn","xl3DColumnClustered",
			"xl3DColumnStacked","xl3DColumnStacked100","xl3DLine","xl3DPie",
			"xl3DPieExploded","xlArea","xlAreaStacked","xlAreaStacked100",
			"xlBarClustered","xlBarOfPie","xlBarStacked","xlBarStacked100",
			"xlBubble","xlBubble3DEffect","xlColumnClustered","xlColumnStacked",
			"xlColumnStacked100","xlConeBarClustered","xlConeBarStacked","xlConeBarStacked100",
			"xlConeCol","xlConeColClustered","xlConeColStacked","xlConeColStacked100",
			"xlCylinderBarClustered","xlCylinderBarStacked","xlCylinderBarStacked100","xlCylinderCol",
			"xlCylinderColClustered","xlCylinderColStacked","xlCylinderColStacked100","xlDoughnut",
			"xlDoughnutExploded","xlLine","xlLineMarkers","xlLineMarkersStacked",
			"xlLineMarkersStacked100","xlLineStacked","xlLineStacked100","xlPie",
			"xlPieExploded","xlPieOfPie","xlPyramidBarClustered","xlPyramidBarStacked",
			"xlPyramidBarStacked100","xlPyramidCol","xlPyramidColClustered","xlPyramidColStacked",
			"xlPyramidColStacked100","xlRadar","xlRadarFilled","xlRadarMarkers",
			"xlStockHLC","xlStockOHLC","xlStockVHLC","xlStockVOHLC",
			"xlSurface","xlSurfaceTopView","xlSurfaceTopViewWireframe","xlSurfaceWireframe",
			"xlXYScatter","xlXYScatterLines","xlXYScatterLinesNoMarkers","xlXYScatterSmooth",
			"xlXYScatterSmoothNoMarkers")]
		[PSObject] $ChartType,
		[Parameter(Position=4)] $Title,
		[Parameter(Position=5)] [ValidateSet("begin","end")] $SheetPosition="begin",
		[Parameter(Position=6)] $FreezeRow=0,
		[Parameter(Position=7)] $FreezeColumn=0,
		[switch] $ChartOnNewSheet,
		[switch] $AppendWorksheet,
		[switch] $Borders=$True,
		[switch] $HeaderColor=$True,
		[switch] $GreenBar=$True,
		[switch] $WrapText=$False,
		[switch] $AutoFit=$True,
		[switch] $AutoFilter=$True,
		[switch] $PassThrough,
		[switch] $Force
	)
	begin {
		function Convert-NumberToA1 {
		param([parameter(Mandatory=$true)] [int]$number)
		$a1Value = $null
		while ($number -gt 0) {
			$multiplier = [int][system.math]::Floor(($number / 26))
			$charNumber = $number - ($multiplier * 26)
			if ($charNumber -eq 0) { $multiplier-- ; $charNumber = 26 }
			$a1Value = [char]($charNumber + 96) + $a1Value
			$number = $multiplier
		}
		return $a1Value
		}

		$Script:WorkingData = @()
	}
	process {
		$Script:WorkingData += $InputData
	}
	end {
		$Props = $Script:WorkingData[0].PSObject.properties | % { $_.Name }
		$Rows = $Script:WorkingData.Count+1
		$Cols = $Props.Count
		$A1Cols = Convert-NumberToA1 $Cols
		$Array = New-Object 'object[,]' $Rows,$Cols

		$Col = 0
		$Props | % {
			$Array[0,$Col] = $_.ToString()
			$Col++
		}

		$Row = 1
		$Script:WorkingData | % {
			$Item = $_
			$Col = 0
			$Props | % {
				if ($Item.($_) -eq $Null) {
					$Array[$Row,$Col] = ""
				} else {
					$Array[$Row,$Col] = $Item.($_).ToString()
				}
				$Col++
			}
			$Row++
		}

		$xl = New-Object -ComObject Excel.Application
		$xl.DisplayAlerts = $False
		$xlFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlWorkbookNormal

		if ([System.IO.Path]::GetExtension($Path) -eq '.xlsx') {
			if ($xl.Version -lt 12) {
				$Path = $Path.Replace(".xlsx",".xls")
			} else {
				$xlFixedFormat = [Microsoft.Office.Interop.Excel.XLFileFormat]::xlWorkbookDefault
			}
		}

		if (Test-Path -Path $Path -PathType "Leaf") {
			if ($AppendWorkSheet) {
				$wb = $xl.Workbooks.Open($Path)
				if ($SheetPosition -eq "end") {
					$wb.Worksheets.Add([System.Reflection.Missing]::Value,$wb.Sheets.Item($wb.Sheets.Count)) | Out-Null
				} else {
					$wb.Worksheets.Add($wb.Worksheets.Item(1)) | Out-Null
				}
			} else {
				if (!($Force)) {
					$Path = $Path.Insert($Path.LastIndexOf(".")," - $(Get-Date -Format "ddMMyyyy-HHmm")")
				}
				$wb = $xl.Workbooks.Add()
				while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(1).Delete() }
			}
		} else {
			$wb = $xl.Workbooks.Add()
			while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(1).Delete() }
		}

		$ws = $wb.ActiveSheet
		try { $ws.Name = $WorksheetName }
		catch { }

		if ($Title) {
			$ws.Cells.Item(1,1) = $Title
			$TitleRange = $ws.Range("a1","$($A1Cols)2")
			$TitleRange.Font.Size = 18
			$TitleRange.Font.Bold=$True
			$TitleRange.Font.Name = "Cambria"
			$TitleRange.Font.ThemeFont = 1
			$TitleRange.Font.ThemeColor = 4
			$TitleRange.Font.ColorIndex = 55
			$TitleRange.Font.Color = 8210719
			$TitleRange.Merge()
			$TitleRange.VerticalAlignment = -4160
			$usedRange = $ws.Range("a3","$($A1Cols)$($Rows + 2)")
			if ($HeaderColor) {
				$ws.Range("a3","$($A1Cols)3").Interior.ColorIndex = 48
				$ws.Range("a3","$($A1Cols)3").Font.Bold = $True
			}
		} else {
			$usedRange = $ws.Range("a1","$($A1Cols)$($Rows)")
			if ($HeaderColor) {
				$ws.Range("a1","$($A1Cols)1").Interior.ColorIndex = 48
				$ws.Range("a1","$($A1Cols)1").Font.Bold = $True
			}
		}
		$usedRange.Value2 = $Array
		if ($Borders) {
			$usedRange.Borders.LineStyle = 1
			$usedRange.Borders.Weight = 2
		}
		$Selection = $ws.UsedRange.EntireColumn
		if($Greenbar){
			$Formula = '=MOD(ROW()-2,3*2)+1<=3'
			$Selection.FormatConditions.Add(2, 0, $Formula) | Out-Null
			$Selection.FormatConditions.Item(1).SetFirstPriority()
			$Selection.FormatConditions.Item(1).Interior.Color = 15071206
			#$Selection.FormatConditions.Item(1).Interior.Color = 13430991	#darker green
		}
		if($WrapText){$Selection.WrapText = $True}
		else{$Selection.WrapText = $False}
		if($FreezeColumn -gt 0 -or $FreezeRow -gt 0){
			if($FreezeColumn -gt 0){$usedRange.application.activewindow.splitcolumn = $FreezeColumn}
			if($FreezeRow -gt 0){$usedRange.application.activewindow.splitrow = $FreezeRow}
			$usedRange.application.activewindow.freezepanes = $true
		}
		if ($AutoFilter) { $usedRange.AutoFilter() | Out-Null }
		if ($AutoFit) { $ws.UsedRange.EntireColumn.AutoFit() | Out-Null }
		if ($ChartType) {
			[Microsoft.Office.Interop.Excel.XlChartType]$ChartType = $ChartType
			if ($ChartOnNewSheet) {
				$wb.Charts.Add().ChartType = $ChartType
				$wb.ActiveChart.setSourceData($usedRange)
				try { $wb.ActiveChart.Name = "$($WorksheetName) - Chart" }
				catch { }
				$wb.ActiveChart.Move([System.Reflection.Missing]::Value,$wb.Sheets.Item($ws.Name))
			} else {
				$ws.Shapes.AddChart($ChartType).Chart.setSourceData($usedRange) | Out-Null
			}
		}
		$wb.SaveAs($Path,$xlFixedFormat)
		$wb.Close()
		$xl.Quit()

		while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($usedRange)) {}
		while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)) {}
		if ($Title) { while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($TitleRange)) {} }
		while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)) {}
		while ([System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)) {}
		[GC]::Collect()

		if ($PassThrough) { return Get-Item $Path }
	}
}


function Confirm-PowerCLI58 {
	Write-Host "Checking System Prerequisites for PowerCLI 5.8+..."
	$powercli = Get-PSSnapin VMware.VimAutomation.Core -Registered
	try {
		switch ($powercli.Version.Major){
			{$_ -ge 6} {
				Write-Host  -nonewline "  Importing PowerCLI 6+ module..."
				Import-Module VMware.VimAutomation.Core -ErrorAction Stop
				Write-Host "  imported."
			} 5 {
				if ($powercli.Version.Minor -lt 8) {
					Write-Warning "VMWare PowerCLI 5.8 or greater is not installed!"
					Send-PopUp "Please Install the latest version of VMware PowerCLI."
					Write-Warning "Please Install the latest version of VMware PowerCLI." -WarningAction Stop
				} else {
					if (!(get-pssnapin -Name Vmware.VimAutomation.Core -ea "silentlycontinue")) {
						Add-PSSnapin VMware.VimAutomation.Core -ErrorAction Stop
						Write-Host "  PowerCLI 5 snapin added; recommend upgrading your PowerCLI version"
					}
				}
			} default {
				throw "This script requires PowerCLI version 5.8 or later"
				Send-PopUp "Please Install the latest version of VMware PowerCLI."
				Write-Warning "Please Install the latest version of VMware PowerCLI." -WarningAction Stop
			}
		}
	} catch {
		Send-PopUp "Please Install the latest version of VMware PowerCLI."
		throw "Could not load the required VMware.VimAutomation.Core cmdlets"
		Write-Warning "Please Install the latest version of PowerCLI." -WarningAction Stop
	}
}

#Ensure minimum version of PowerCLI loaded
Confirm-PowerCLI58


$Scriptblock = {
	param (
		[string]$vCenter
	)
	$WarningPreference = "SilentlyContinue"
	# Add PowerCLI
	Add-PSSnapin -Name Vmware.VimAutomation.Core -erroraction 'silentlycontinue'
	
	$Null = Connect-VIServer -Server $vCenter -WarningAction SilentlyContinue
	
	$VCHostCounter = 0    # Initialize counter for progress bar
	$VCCounter++
	$VCHosts = Get-VMHost | Sort Name
	$VCHostReport = @()
	$NumHosts = $VCHosts.count
	$CustomFieldLUNs = "LUNs"
	$CustomFieldLocation = "Location"
	$CustomFieldRemoteMgt = "RemoteMgt"
	$SI = Get-View ServiceInstance
	$CFM = Get-View $SI.Content.CustomFieldsManager
	$myCustomFieldLUNs = $CFM.Field | where {$_.Name -eq $CustomFieldLUNs}
	$myCustomFieldLocation = $CFM.Field | where {$_.Name -eq $CustomFieldLocation}
	$myCustomFieldRemoteMgt = $CFM.Field | where {$_.Name -eq $CustomFieldRemoteMgt}

	function Get-VMHostWSManInstance {
		param (
		[Parameter(Mandatory=$TRUE,HelpMessage="VMHosts to probe")]
		[VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl[]]
		$VMHost,

		[Parameter(Mandatory=$TRUE,HelpMessage="Class Name")]
		[string]
		$class,

		[switch]
		$ignoreCertFailures,
	
		[System.Management.Automation.PSCredential]
		$credential=$null
		)

		$omcBase = "http://schema.omc-project.org/wbem/wscim/1/cim-schema/2/"
		$dmtfBase = "http://schemas.dmtf.org/wbem/wscim/1/cim-schema/2/"
		$vmwareBase = "http://schemas.vmware.com/wbem/wscim/1/cim-schema/2/"
	
		if ($ignoreCertFailures) {
			$option = New-WSManSessionOption -SkipCACheck -SkipCNCheck -SkipRevocationCheck
		} else {
			$option = New-WSManSessionOption
		}
		foreach ($H in $VMHost) {
			if ($credential -eq $null) {
				$hView = $H | Get-View -property Value
				$Error.clear()	# begin error checking since permissions issues may occur querying older hardware, like HP BL460 G6
				try { $ticket = $hView.AcquireCimServicesTicket() }
				catch {	write-host "Error occurred connecting to CIM for $H" }
				if (!$Error) { 	#No Error Occured
					try { $password = convertto-securestring $ticket.SessionId -asplaintext -force }
					catch { write-host "Error occurred converting password" }
					if (!$Error) { 	#No Error Occured
						$credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $ticket.SessionId, $password
					}
				}
			}
			$uri = "https`://" + $h.Name + "/wsman"
			if ($class -cmatch "^CIM") {
				$baseUrl = $dmtfBase
			} elseif ($class -cmatch "^OMC") {
				$baseUrl = $omcBase
			} elseif ($class -cmatch "^VMware") {
				$baseUrl = $vmwareBase
			} else {
				Throw "Unrecognized class"
			}
		$ErrorActionPreference = "SilentlyContinue"
		if (!$Error) {$Result = Get-WSManInstance -Authentication basic -ConnectionURI $uri -Credential $credential -Enumerate -Port 443 -UseSSL -SessionOption $option -ResourceURI "$baseUrl/$class" }
		if($Result -ne $null){$Result}
		$ErrorActionPreference = "Continue"
		}
	}
	
	foreach($VCHost in $VCHosts){
		$VCHostCounter++
		#Write-Progress -Activity "Gathering host information from vCenter ($VCCounter/$NumVCenters): $vCenter" -Status "Processing Host ($VCHostCounter/$NumHosts): $VCHost" -PercentComplete (100*($VCHostCounter/$NumHosts))    # Display progress bar
		$VCHostView = $VCHost | Get-View
		$Report = "" | select vCenter, Hostname, Version, Build, Cluster, Manu, Model, MemGB, CPU, NumCPUs, NumCores, HyperthreadingActive, BIOSver, BIOSdate, ServiceTag, NICDriver, NICDescription, NICFirmware, NICDriverVersion,  NIC2Driver, NIC2Description, NIC2Firmware, NIC2DriverVersion, HBADriver, HBADescription, HBADriverVersion, HPSmartArrayFW, HPSmartArrayDrv, iLOFirmware, VMKIP0, VMKIP1, VMKIP2, VMKIP3, License, DomainMembership, VAAI_MoveInitLockUnmap, Syslog, NTP, BootTime, State, PowerPath, LUNs, Location, RemoteMgt
		$Report.vCenter = $vCenter
		$Report.Hostname = $VCHost.Name 
		$Report.Version = $VCHostView.Config.Product.Version 
		$Report.Build = $VCHostView.Config.Product.Build 
		$Report.Cluster = $VCHost.Parent
		$Report.Manu = $VCHostView.Hardware.SystemInfo.Vendor 
		$Report.Model = $VCHostView.Hardware.SystemInfo.Model 
		$Report.MemGB = [int]$VCHost.MemoryTotalGB
		$Report.CPU = (($VCHost.ProcessorType).Replace("Intel(R) Xeon(R) CPU           ","")).Replace("Intel(R) Xeon(R) CPU ","")
		$Report.NumCPUs = $VCHostView.Hardware.CpuInfo.NumCpuPackages
		$Report.NumCores = $VCHostView.Hardware.CpuInfo.NumCpuCores
		$Report.HyperthreadingActive = $VCHost.HyperthreadingActive
		$Report.BIOSver = $VCHostView.Hardware.BiosInfo.BiosVersion
		$Report.BIOSdate = $VCHostView.Hardware.BiosInfo.ReleaseDate
		
		if($Report.Manu -eq "Cisco Systems Inc"){
			$Report.ServiceTag = (Get-VMHostWSManInstance -VMHost $VCHost -class CIM_PhysicalPackage -ignoreCertFailures | Where-Object {$_.ElementName -eq "Chassis" -and $_.SerialNumber -like "FCH*"}).SerialNumber
		}
		else{
			$Report.ServiceTag = (Get-VMHostWSManInstance -VMHost $VCHost -class CIM_PhysicalPackage -ignoreCertFailures | Where-Object {$_.ElementName -eq "Chassis"}).SerialNumber
		}
		if($Report.Manu -eq "Dell Inc." -or $Report.Manu -eq "HP"){
			$Report.ServiceTag = (Get-VMHostWSManInstance -VMHost $VCHost -class CIM_PhysicalPackage -ignoreCertFailures | Where-Object {$_.ElementName -eq "Chassis" -and $_.OEMSpecificStrings}).SerialNumber
		}
		if($Report.ServiceTag.Length -lt 1){
			$Report.ServiceTag = ($VCHostView.Hardware.SystemInfo.OtherIdentifyingInfo | where {$_.IdentifierType.Key -eq "ServiceTag"}).IdentifierValue
		}
#		$Report.ServiceTag = $VCHostView | select @{Name="ServiceTag"; Expression={($_.Hardware.SystemInfo.OtherIdentifyingInfo | where {$_.IdentifierType.Key -eq "ServiceTag"}).IdentifierValue}} 
#		if($Report.Manu -eq "HP"){
#			$Report.ServiceTag = ($VCHostView.Hardware.SystemInfo.OtherIdentifyingInfo | where {$_.IdentifierType.Key -eq "ServiceTag"}).IdentifierValue
#			if($Report.ServiceTag.Length -lt 1){
#				$ReportTemp = ($VCHostView.Hardware.SystemInfo.OtherIdentifyingInfo | where {$_.IdentifierValue -like "USE*"}).IdentifierValue  # HP serial #s start with 'USE'
#			}
#		}
#		else{
#			$Report.ServiceTag = ($VCHostView.Hardware.SystemInfo.OtherIdentifyingInfo | where {$_.IdentifierValue -like "FCH*"}).IdentifierValue   # UCS serial #s start with 'FCH'
#		}

		$esxcli = get-esxcli -VMHost $VCHost
		$NIC0Desc = $esxcli.network.nic.list() | where-object {$_.name -eq "vmnic0"} | Select-Object Driver,Description
		#$Report.NICDriver = $esxcli.network.nic.get("vmnic0").DriverInfo.Driver
		$Report.NICDriver = $NIC0Desc.Driver
		$Report.NICDescription = $NIC0Desc.Description
		$Report.NICFirmware = $esxcli.network.nic.get("vmnic0").DriverInfo.FirmwareVersion
		$Report.NICDriverVersion = $esxcli.network.nic.get("vmnic0").DriverInfo.Version
		$NIC2Desc = $esxcli.network.nic.list() | where-object {$_.name -eq "vmnic2"} | Select-Object Driver,Description
		if($NIC2Desc.Description -ne $NIC0Desc.Description){
			$Report.NIC2Driver = $NIC2Desc.Driver
			$Report.NIC2Description = $NIC2Desc.Description
			$Report.NIC2Firmware = $esxcli.network.nic.get("vmnic2").DriverInfo.FirmwareVersion
			$Report.NIC2DriverVersion = $esxcli.network.nic.get("vmnic2").DriverInfo.Version
		}
		else{
			$Report.NIC2Driver = ""
			$Report.NIC2Description = ""
			$Report.NIC2Firmware = ""
			$Report.NIC2DriverVersion = ""
		}
		
		$HBADriver = $VCHost | Get-VMHostHba -Type FibreChannel | Select -Expand Driver | Get-Unique
		$HBADescription = $VCHost | Get-VMHostHba -Type FibreChannel | Select Model | Get-Unique
		$Report.HBADriverVersion = $Null
		$HBADriverVersion = $esxcli.software.vib.list() | ? {$_.Name -match "$HBADriver"} | Select -Expand Version
		$Report.HBADriver = $HBADriver
		$Report.HBADescription = $HBADescription.Model
		$Report.HBADriverVersion = ($HBADriverVersion -split "-1OEM")[0]	# trim off "-1OEM.500.0.0.472560" from Cisco entries
		
		$arrNumericSensorInfo = @($VCHostView.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo)
		if($Report.Manu -eq "HP"){
			$nsiArrayCtrlr = ($arrNumericSensorInfo | ? {$_.Name -like "HP Smart Array Controller*"})[0]
			$nsiILO = $arrNumericSensorInfo | ? {$_.Name -like "Hewlett-Packard BMC Firmware*"}
			if($nsiArrayCtrlr.Name -ne $null){$Report.HPSmartArrayFW = ($nsiArrayCtrlr.Name.ToString()).SubString(41)}		# trim off "HP Smart Array Controller HPSA1 Firmware " prefix
			else{$Report.HPSmartArrayFW = $null}
			if($nsiILO.Name -ne $null){$Report.iLOFirmware = ($nsiILO.Name.ToString()).SubString(47)}  # trim off "Hewlett-Packard BMC Firmware (node 0) 46:10000 " prefix
			else{$Report.iLOFirmware = $null}
			$nsiSADrv = $arrNumericSensorInfo | ? {$_.Name -like "Hewlett-Packard scsi-hpsa *"}
			if($nsiSADrv.Name -ne $null){$Report.HPSmartArrayDrv = (($nsiSADrv.Name.ToString() -split "OEM.500.0.0.472560")[0] -split "Hewlett-Packard scsi-hpsa ")[1]}  # trim off unnecessary data from line
			else{$Report.HPSmartArrayDrv = $null}
		}
		else{
			$Report.iLOFirmware = $null
			$Report.HPSmartArrayFW = $null
			$Report.HPSmartArrayDrv = $null
		}

		$Report.VMKIP0 = ($VCHostView.config.vmotion.netconfig.CandidateVnic | where-object {$_.Portgroup -eq "Management Network"}).Spec.Ip.IpAddress
		$Report.VMKIP1 = ($VCHostView.config.vmotion.netconfig.CandidateVnic | where-object {$_.Device -eq "vmk1"}).Spec.Ip.IpAddress
		$Report.VMKIP2 = ($VCHostView.config.vmotion.netconfig.CandidateVnic | where-object {$_.Device -eq "vmk2"}).Spec.Ip.IpAddress
		$Report.VMKIP3 = ($VCHostView.config.vmotion.netconfig.CandidateVnic | where-object {$_.Device -eq "vmk3"}).Spec.Ip.IpAddress

#		$Report.VMKIP0 = ($VCHost | Get-VMHostNetwork).VirtualNic[0].IP
#		$Report.VMKIP1 = ($VCHost | Get-VMHostNetwork).VirtualNic[1].IP
#		$Report.VMKIP2 = ($VCHost | Get-VMHostNetwork).VirtualNic[2].IP
#		$Report.VMKIP3 = ($VCHost | Get-VMHostNetwork).VirtualNic[3].IP
		$Report.License = $VCHost.LicenseKey
		$DomainMembership = ($VCHost | Get-VMHostAuthentication).DomainMembershipStatus
		$Report.DomainMembership = $DomainMembership
		$VAAIMove = ($VCHost | Get-AdvancedSetting -Name DataMover.HardwareAcceleratedMove).value
		$VAAIInit = ($VCHost | Get-AdvancedSetting -Name DataMover.HardwareAcceleratedInit).value
		$VAAILock = ($VCHost | Get-AdvancedSetting -Name VMFS3.HardwareAcceleratedLocking).value
		$VAAIUnmap = ($VCHost | Get-AdvancedSetting -Name VMFS3.EnableBlockDelete).value
		$Report.VAAI_MoveInitLockUnmap = "$VAAIMove|$VAAIInit|$VAAILock|$VAAIUnmap"
		$SyslogHost = ($VCHost | Select @{N="SyslogHost";E={(Get-AdvancedSetting -Entity $_ -Name "Syslog.global.logHost").Value}}).SyslogHost
		$Report.Syslog = $SyslogHost
		$NTPHost = [system.string]::Join(",",($VCHost | Select Name, @{N="NTP";E={Get-VMHostNtpServer $_}}).NTP)
		$Report.NTP = $NTPHost
		$Report.BootTime = $VCHostView.Summary.Runtime.BootTime
		$Report.State = $VCHost.ConnectionState
		$Report.PowerPath = (($Esxcli.software.vib.list() | ?{$_.Name -like "powerpath*"}).Version)[0]
		
		$Report.LUNs =		($VCHostView.CustomValue | ?{$_.Key -eq $myCustomFieldLUNs.Key}).Value
		$Report.Location =	($VCHostView.CustomValue | ?{$_.Key -eq $myCustomFieldLocation.Key}).Value
		$Report.RemoteMgt =	($VCHostView.CustomValue | ?{$_.Key -eq $myCustomFieldRemoteMgt.Key}).Value
		
		$VCHostReport += $Report
	}
	$Null = Disconnect-VIServer -Server $vCenter -confirm:$False -WarningAction SilentlyContinue
	$VCHostReport
}


#Cleanup old jobs
foreach ($Job in (Get-Job | where { $_.Name -like "Host-*" })) {Remove-Job $Job -Force}

$Submitted = 0
$StopWatch = New-Object system.Diagnostics.Stopwatch
# Start the stop watch
$StopWatch.Start()
foreach ($vCenter in $vCenters) {
	Write-Progress -Activity "vCenter Inventory" -Status "Submitting threads: $($NumVCenters - $Submitted)  of $($NumVCenters)" -PercentComplete ($Submitted / $NumVCenters * 100)
	Write-Host "Next: $vCenter"
	$running = @(Get-Job | Where-Object { $_.State -eq 'Running' })
	
	while (@(Get-Job | where { $_.State -eq "Running" }).Count -ge $MaxThreads){
		$Elapsed = $StopWatch.Elapsed
		$ElapsedTime = [system.String]::Format("{0:00}:{1:00}:{2:00}", $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds );
		Write-Progress -Activity "vCenter Inventory" -Status "Submitting threads (Elapsed $($ElapsedTime)): $($NumVCenters - $Submitted) of $($NumVCenters)" -PercentComplete ($Submitted / $NumVCenters * 100)
		Start-Sleep -Seconds 3
	}
	$null = Start-Job -ScriptBlock $Scriptblock -ArgumentList $vCenter -Name Host-$vCenter -PSVersion 3.0 -RunAs32
	$Submitted ++
	Get-Job
	Write-Host "________________________________________________________________________________"
}
Write-Host "********************************************************************************"

Write-Host "Collecting Host NIC details..."
foreach ($vCenter in $vCenters) {
	if($vCenter -match "Thrasher*"){
		If(${env:ComputerName} -eq "GAR7"){	$CredPath = "F:\Credentials\vreporter05.xml" }
		else{ $CredPath = "C:\Temp\Thrasher.xml" }
		$Creds = Get-MyCredential ($CredPath)
		$Null = Connect-VIServer -Server $vCenter -WarningAction SilentlyContinue -Credential $Creds
	} else{ $Null = Connect-VIServer -Server $vCenter -WarningAction SilentlyContinue }

	Write-Host ""
	$VCHostCounter = 0    # Initialize counter for progress bar
	$VCCounter++
	$VCHosts = Get-VMHost | Sort Name
	$NumHosts = $VCHosts.count
	foreach ($VCHost in $VCHosts){
		$VCHostCounter++
		Write-Progress -Activity "Gathering host information from vCenter ($VCCounter/$NumVCenters): $vCenter" -Status "Processing Host ($VCHostCounter/$NumHosts): $VCHost" -PercentComplete (100*($VCHostCounter/$NumHosts))    # Display progress bar
		$HostNICs += $VCHost | Get-VMHostNetworkAdapter | select  @{Name="vCenter"; Expression={$vCenter}},VMhost,Name,IP,SubnetMask,Mac,PortGroupName,vMotionEnabled,mtu,FullDuplex,BitRatePerSec
	}
	Disconnect-VIServer -Server $vCenter -confirm:$False
}


$PreviousJobCount = $NumVCenters
while (@(Get-Job | where { $_.State -eq "Running" }).Count -ne 0){
	$JobCount = @(Get-Job | where {$_.State -ne "Running"}).Count
	$Elapsed = $StopWatch.Elapsed
	$ElapsedTime = [system.String]::Format("{0:00}:{1:00}:{2:00}", $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds );
	Write-Progress -Activity "vCenter Inventory" -Status "Waiting for background jobs to finish (Elapsed $($ElapsedTime)): $($NumVCenters - $JobCount) of $($NumVCenters)" -PercentComplete ($JobCount / $NumVCenters * 100)
	if ($PreviousJobCount -ne $JobCount){
		Get-Job
		Write-Host "________________________________________________________________________________"
		$PreviousJobCount = $JobCount
	}
	Start-Sleep -Seconds 3
}
Get-Job

$JobData = @()
Write-Progress -Activity "vCenter Inventory" -Status "All background jobs completed!" -PercentComplete 100
$Elapsed = $StopWatch.Elapsed
$ElapsedTime = [system.String]::Format("{0:00}:{1:00}:{2:00}", $Elapsed.Hours, $Elapsed.Minutes, $Elapsed.Seconds );
write-host "Elapsed job time: $ElapsedTime"
$StopWatch.Reset()
foreach ($Job in (Get-Job)){
	$JobData += Receive-Job $Job | Select-Object -Property * -ExcludeProperty RunspaceID,PSComputerName,PSShowComputerName
	Remove-Job $Job -WarningAction SilentlyContinue
}
$JobData = $JobData | Sort vCenter,Hostname

Get-Date -format yyyy\-MM\-dd\-HH:mm:ss

$Timestamp = Get-Date -format yyyy\-MM\-dd\-HHmm
if($NumVCenters -gt 1){
	$VCNames = [system.string]::Join("~",$vCenters)      # Concatenate the vCenter names for the filename
	if($VCNames.Length -gt 30) {    # unless it's too long with too many names
		$VCNames = [STRING]$NumVCenters + "vCenters"
	}
}
else{		# Only a single vCenter name was entered
	$VCNames = $vCenter
}
$ExportFile1 = "_Hosts_Current.csv"
$ExportFile2 = "_Hosts_Current.xlsx"
$ExportFile3 = "Hosts_" + $VCNames + "_" + $Timestamp + ".xlsx"
$ExportFile4 = "_HostNICs.csv"


Write-Progress -Activity "Exporting host report" -Status "Writing $ExportPath\$ExportFile1"    # Display progress bar
New-Item -ItemType Directory -Force -Path $ExportPath | Out-Null          #Create directory for the report

Remove-Item $ExportPath\$ExportFile2

$JobData | Export-CSV $ExportPath\$ExportFile1 –NoTypeInformation
$JobData | Export-Xlsx -Path $ExportPath\$ExportFile2 -FreezeRow 1 -FreezeColumn 2 -Force
$HostNICs | Export-CSV $ExportPath\$ExportFile4 –NoTypeInformation

if(($DateNow.DayOfWeek -eq $AdminReportDay) -or ($Test -eq $True)){
	New-Item -ItemType Directory -Force -Path $ExportPath\Archive | Out-Null          #Create directory for the report
	copy-File "C:\Temp\$ExportFile2" "$ExportPath\Archive\$ExportFile3"
}

Write-Host ""
Write-Host -foregroundcolor green $MyInvocation.InvocationName "script completed."
