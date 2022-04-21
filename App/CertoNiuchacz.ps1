Add-Type -AssemblyName PresentationFramework
[xml]$xaml =
@"
<Window
		xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
		xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
		Title="CertoNiuchacz v2.0.1.0" Height="300" Width="500" MinWidth="500" MinHeight="300"  WindowStyle="ToolWindow" SizeToContent="WidthAndHeight">
	<Grid>
		<Grid.RowDefinitions>
			<RowDefinition Height="70"/>
			<RowDefinition/>
			<RowDefinition Height="30"/>
		</Grid.RowDefinitions>
		<Grid.ColumnDefinitions>
			<ColumnDefinition/>
			<ColumnDefinition Width="70"/>
			<ColumnDefinition Width="70"/>
			<ColumnDefinition Width="70"/>
		</Grid.ColumnDefinitions>

		<Label Content="Filters:" Height="22" Margin="12,37,170,0" VerticalAlignment="Top" FontSize="9"/>
		<ComboBox x:Name="CA" Margin="117,10,10,0" VerticalAlignment="Top" Height="22" IsEditable="True" Grid.ColumnSpan="3"/>

		<Label Content="Certification Authority:" Height="22" Margin="10,10,172,0" VerticalAlignment="Top" FontSize="9"/>
		<ComboBox x:Name="Filter" Margin="117,37,10,0" VerticalAlignment="Top" Height="22" IsEditable="True" Grid.ColumnSpan="3"/>
		<Button x:Name="Connect" Content="Connect" Height="22" Margin="0,10,10,0" VerticalAlignment="Top" Width="60" HorizontalAlignment="Right" Grid.Column="3"/>
		<Button x:Name="Search" Content="Search" Height="22" Margin="0,37,10,0" VerticalAlignment="Top" Width="60" HorizontalAlignment="Right" Grid.Column="3"/>

		<ListView x:Name="Result" Margin="10" Grid.ColumnSpan="4" SelectionMode="Extended" Grid.Row="1" FontSize="10">
			<ListView.View>
				<GridView>
					<GridViewColumn Header="UPN" DisplayMemberBinding="{Binding UPN}"/>
					<GridViewColumn Header="SerialNumber" DisplayMemberBinding="{Binding SerialNumber}"/>
					<GridViewColumn Header="RequestID" DisplayMemberBinding="{Binding RequestID}"/>
					<GridViewColumn Header="NotBefore" DisplayMemberBinding="{Binding NotBefore}"/>
					<GridViewColumn Header="NotAfter" DisplayMemberBinding="{Binding NotAfter}"/>
					<GridViewColumn Header="Template" DisplayMemberBinding="{Binding Template}"/>
					<GridViewColumn Header="CertificateHash" DisplayMemberBinding="{Binding CertificateHash}"/>
				</GridView>
			</ListView.View>
		</ListView>
		<Label x:Name="Status" Content="" Height="22" Margin="10,0,80,0" VerticalAlignment="Top" FontSize="9" Grid.Row="2"/>
		<Button x:Name="GetScript" Content="GetKey Script" Margin="0,0,10,10" Grid.Row="2" Width="60" HorizontalAlignment="Right" FontSize="8" IsEnabled="False"/>
		<Button x:Name="GetRec" Content="GetKey (.rec)" Margin="0,0,10,10" IsEnabled="False" RenderTransformOrigin="0.533,0.561" Grid.Column="1" Grid.Row="2" FontSize="8"/>
		<Button x:Name="RecKey" Content="RecKey (.p12)" Margin="0,0,10,10" IsEnabled="False" RenderTransformOrigin="0.533,0.561" Grid.Column="2" Grid.Row="2" FontSize="8"/>
		<Button x:Name="Merge" Content="Merge (.pfx)" Margin="0,0,10,10" IsEnabled="False" RenderTransformOrigin="0.533,0.561" Grid.Column="3" Grid.Row="2" FontSize="8"/>

	</Grid>
</Window>
"@


Function Set-SearchButtonStatus
{
	if($ComboBox_Filter.SelectedItems.Count -and $ComboBox_CA.SelectedItems.Count)
	{
		$Button_Search.IsEnabled = $true
	}
	else
	{
		$Button_Search.IsEnabled = $false
	}
}


Function Set-ButtonStatus
{
	if($ListView_Result.SelectedItems.Count)
	{
		$Button_GetScript.IsEnabled = $true
		$Button_GetRec.IsEnabled = $true
		$Button_RecKey.IsEnabled = $true
		$Button_Merge.IsEnabled = $true
	}
	else
	{
		$Button_GetScript.IsEnabled = $false
		$Button_GetRec.IsEnabled = $false
		$Button_RecKey.IsEnabled = $false
		$Button_Merge.IsEnabled = $true
	}
}

Function Format-Certutil
{
	param($Result)

	$CertObjects = @()
	foreach($Row in $Result)
	{
		switch($Row)
		{
				{$_ -match 'Row \d+:(.*)'} {
					$CertObject = New-Object PSObject
					break
				}

				{$_ -match '(.*):(.*)'} {
					if($CertObject){
						$CertObject | Add-Member -MemberType NoteProperty -Name ($Row -Split ": ")[0].Replace(" ","").Trim() -Value (($Row -Split ": ")[1] -replace "$([char]96)|$([char]34)|$([char]39)","").Trim() -Force
					}
					break
				}

				{$_.Trim() -eq ""} {
					if($CertObject){
						$CertObjects += $CertObject
						$CertObject = $null
					}
					break
				}
			 }
	}

	Return $CertObjects
}

Function New-RemoteSession
{
	if($null -eq $Script:Session)
	{
		$CAText = $ComboBox_CA.SelectedItem
		if($null -eq $CAText) { $CAText = $ComboBox_CA.Text }
		if($null -eq $CAText)
		{
			$Label_Status.Content = "Missing CA"
			$Window.Title = $OrgText
			return
		}

		$Credential = Get-Credential -Message "Credential for WinRM remote session to $CAText"
		$ComputerName = ($CAText -split "\\")[0]
		$OrgText = $Window.Title
		$Window.Title += " - Connecting..."

		try
		{
			$Script:Session = New-PSSession -ComputerName $ComputerName -Credential $Credential -Authentication Negotiate -ErrorAction Stop
			$Button_Connect.Content = "Disconnect"
			$Label_Status.Content = "Connected remote session to $CAText"
		}
		catch
		{
			$Label_Status.Content = "Cant establish remote session to $CAText"

			$Button = [System.Windows.MessageBoxButton]::OK
			$Icon = [System.Windows.MessageBoxImage]::Error
			$DefaultButton = [System.Windows.MessageBoxResult]::None

			$Result = [System.Windows.MessageBox]::Show($_.Exception,"Error",$Button,$Icon,$DefaultButton)
		}

		$Window.Title = $OrgText
	} else {
		Remove-PSSession -Session $Script:Session
		$Label_Status.Content = "Disconnected remote session to $CAText"
		$Script:Session = $null
		$Button_Connect.Content = "Connect"
	}
}

Function Get-Certs
{
	$OrgText = $Window.Title
	$Window.Title += " - Search..."
	$Label_Status.Content = ""

	$SearchText = $ComboBox_Filter.SelectedItem
	if($null -eq $SearchText) { $SearchText = $ComboBox_Filter.Text }

	$CAText = $ComboBox_CA.SelectedItem
	if($null -eq $CAText) { $CAText = $ComboBox_CA.Text }

	if($null -eq $SearchText)
	{
		$Label_Status.Content = "Missing filter"
		$Window.Title = $OrgText
		return
	}
	if($null -eq $CAText)
	{
		$Label_Status.Content = "Missing CA"
		$Window.Title = $OrgText
		return
	}

	$Command = 'certutil.exe -config "' + $CAText + '" -view -restrict "' + $SearchText + '" -out "UPN,SerialNumber,RequestID,NotBefore,NotAfter,CertificateTemplate,CertificateHash"'
	Write-Host $Command
	if($null -eq $Script:Session)
	{
		$Result = Invoke-Expression $Command
	} else {
		$Result = Invoke-Command -Session $Script:Session -ScriptBlock { param($Command) Invoke-Expression $Command } -ArgumentList $Command
	}

	$Certs = Format-Certutil -Result $Result

	Foreach($Cert in $Certs)
	{
		if(($ListView_Result.Items | Select-Object -exp SerialNumber) -notcontains $Cert.SerialNumber)
		{
			$Row = New-Object -TypeName PSCustomObject -Property @{
				UPN = $Cert.UserPrincipalName
				SerialNumber = $Cert.SerialNumber
				RequestID = $Cert.IssuedRequestID
				NotBefore = $Cert.CertificateEffectiveDate
				NotAfter = $Cert.CertificateExpirationDate
				Template = $Cert.CertificateTemplate
				CertificateHash = $Cert.CertificateHash
			}

			$ListView_Result.Items.Add($Row) | Out-Null
		}
	}

	$Window.Title = $OrgText
}

Function Invoke-GetScript
{
	$OrgText = $Window.Title
	$Window.Title += " - GetScript..."

	$CAText = $ComboBox_CA.SelectedItem
	if($null -eq $CAText) { $CAText = $ComboBox_CA.Text }

	foreach ($SelectedItem in $ListView_Result.SelectedItems)
	{
		$SearchText = $SelectedItem.SerialNumber
		$Command = 'certutil.exe -config "' + $CAText + '" -f -getkey ' + $SearchText + ' script ' + $SearchText + '.log'
		Write-Host $Command
		if($null -eq $Script:Session)
		{
			$Result = Invoke-Expression $Command
		} else {
			$Result = Invoke-Command -Session $Script:Session -ScriptBlock { param($Command) Invoke-Expression $Command } -ArgumentList $Command
			$ScriptFile = ([String]($Result | Select-String ".log")).Trim()
			$ScriptFilePath = Invoke-Command -Session $Session -ScriptBlock { param($ScriptFile) (Get-Item $ScriptFile).FullName } -ArgumentList $ScriptFile
			Copy-Item -FromSession $Session -Path $ScriptFilePath -Destination ".\" -Force
		}

		if($Result -match "Error")
		{
			$Button = [System.Windows.MessageBoxButton]::OK
			$Icon = [System.Windows.MessageBoxImage]::Error
			$DefaultButton = [System.Windows.MessageBoxResult]::None

			$Result = [System.Windows.MessageBox]::Show($Result,"Error",$Button,$Icon,$DefaultButton)
		}
	}

	$Window.Title = $OrgText
}

Function Invoke-GetRec
{
	$OrgText = $Window.Title
	$Window.Title += " - GetRec..."

	foreach ($SelectedItem in $ListView_Result.SelectedItems)
	{
		$SearchText = $SelectedItem.SerialNumber

		$ScriptFile = $SearchText + ".log"
		if(Test-Path -Path $ScriptFile)
		{
			$Command = Get-Content $($SearchText + ".log") | Select-String "CertUtil -config "
			Write-Host $Command
			if($null -eq $Script:Session)
			{
				$Result = Invoke-Expression $Command
			} else {
				$Result = Invoke-Command -Session $Script:Session -ScriptBlock { param($Command) Invoke-Expression $Command } -ArgumentList $Command
				$ScriptFile = ([String]($Result | Select-String ".rec")).Trim() -replace "Error saving key data: ",""
				$ScriptFilePath = Invoke-Command -Session $Session -ScriptBlock { param($ScriptFile) (Get-Item $ScriptFile).FullName } -ArgumentList $ScriptFile
				Copy-Item -FromSession $Session -Path $ScriptFilePath -Destination ".\" -Force
			}

			if($Result -match "Error")
			{
				$Button = [System.Windows.MessageBoxButton]::OK
				$Icon = [System.Windows.MessageBoxImage]::Error
				$DefaultButton = [System.Windows.MessageBoxResult]::None

				$Result = [System.Windows.MessageBox]::Show($Result,"Error",$Button,$Icon,$DefaultButton)
			}
		} else {
			$Button = [System.Windows.MessageBoxButton]::OK
			$Icon = [System.Windows.MessageBoxImage]::Error
			$DefaultButton = [System.Windows.MessageBoxResult]::None

			$Result = [System.Windows.MessageBox]::Show("Panie, nie widza skryptu, jak tu żyć...","Error",$Button,$Icon,$DefaultButton)
		}
	}

	$Window.Title = $OrgText
}

Function Invoke-RecKey
{
	$OrgText = $Window.Title
	$Window.Title += " - RecKey..."

	foreach ($SelectedItem in $ListView_Result.SelectedItems)
	{
		$SearchText = $SelectedItem.SerialNumber

		$ScriptFile = $SearchText + ".log"
		if(Test-Path -Path $ScriptFile)
		{
			$Command = Get-Content $($SearchText + ".log") | Select-String " -recoverkey -user "
			Write-Host $Command
			if($null -eq $Script:Session)
			{
				$Result = Invoke-Expression $Command
			} else {
				$Result = Invoke-Command -Session $Script:Session -ScriptBlock { param($Command) Invoke-Expression $Command } -ArgumentList $Command
				$ScriptFile = ([String]($Result | Select-String ".p12")).Trim() -replace "Error saving key data: ",""
				$ScriptFilePath = Invoke-Command -Session $Session -ScriptBlock { param($ScriptFile) (Get-Item $ScriptFile).FullName } -ArgumentList $ScriptFile
				Copy-Item -FromSession $Session -Path $ScriptFilePath -Destination ".\" -Force
			}

			if($Result -match "Error")
			{
				$Button = [System.Windows.MessageBoxButton]::OK
				$Icon = [System.Windows.MessageBoxImage]::Error
				$DefaultButton = [System.Windows.MessageBoxResult]::None

				$Result = [System.Windows.MessageBox]::Show($Result,"Error",$Button,$Icon,$DefaultButton)
			}
		} else {
			$Button = [System.Windows.MessageBoxButton]::OK
			$Icon = [System.Windows.MessageBoxImage]::Error
			$DefaultButton = [System.Windows.MessageBoxResult]::None

			$Result = [System.Windows.MessageBox]::Show("Panie, nie widza skryptu, jak tu żyć...","Error",$Button,$Icon,$DefaultButton)
		}
	}

	$Window.Title = $OrgText
}

Function Invoke-Merge
{
	$OrgText = $Window.Title
	$Window.Title += " - Merge..."

	foreach ($SelectedItem in $ListView_Result.SelectedItems)
	{
		$SearchText = $SelectedItem.SerialNumber

		$ScriptFile = $SearchText + ".log"
		if(Test-Path -Path $ScriptFile)
		{
			$Command = Get-Content $($SearchText + ".log") | Select-String " -MergePFX -user "
			Write-Host $Command
			if($null -eq $Script:Session)
			{
				$Result = Invoke-Expression $Command
			} else {
				$Result = Invoke-Command -Session $Script:Session -ScriptBlock { param($Command) Invoke-Expression $Command } -ArgumentList $Command
				if($Result -eq "CertUtil: -MergePFX command completed successfully.")
				{
					$Command -match "(?'Cmd'CertUtil -p.*-user)\s""(?'Key'.*\.p12)""\s""(?'Pfx'.*\.p12)"""
					$ScriptFile = $Matches['Pfx']
					$ScriptFilePath = Invoke-Command -Session $Session -ScriptBlock { param($ScriptFile) (Get-Item $ScriptFile).FullName } -ArgumentList $ScriptFile
					Copy-Item -FromSession $Session -Path $ScriptFilePath -Destination ".\" -Force
				}
			}

			if($Result -match "Error")
			{
				$Button = [System.Windows.MessageBoxButton]::OK
				$Icon = [System.Windows.MessageBoxImage]::Error
				$DefaultButton = [System.Windows.MessageBoxResult]::None

				$Result = [System.Windows.MessageBox]::Show($Result,"Error",$Button,$Icon,$DefaultButton)
			}
		} else {
			$Button = [System.Windows.MessageBoxButton]::OK
			$Icon = [System.Windows.MessageBoxImage]::Error
			$DefaultButton = [System.Windows.MessageBoxResult]::None

			$Result = [System.Windows.MessageBox]::Show("Panie, nie widza skryptu, jak tu żyć...","Error",$Button,$Icon,$DefaultButton)
		}
	}

	$Window.Title = $OrgText
}


#Init window
$XmlNodeReader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window = [Windows.Markup.XamlReader]::Load($XmlNodeReader)
$ComboBox_CA = $Window.FindName("CA")
$ComboBox_Filter = $Window.FindName("Filter")
$Button_Connect = $Window.FindName("Connect")
$Button_Search = $Window.FindName("Search")
$ListView_Result = $Window.FindName("Result")
$Button_GetScript= $Window.FindName("GetScript")
$Button_GetRec = $Window.FindName("GetRec")
$Button_RecKey = $Window.FindName("RecKey")
$Button_Merge = $Window.FindName("Merge")
$Label_Status = $Window.FindName("Status")

#Change script location
Set-Location -Path $([System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition))

#Check admin role
$User = [Security.Principal.WindowsIdentity]::GetCurrent()
$Role = (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if(!$Role)
{
	$Label_Status.Content = "Admin rights required !!!"

	$Button_Connect.IsEnabled = $false
	$Button_Search.IsEnabled = $false
	$ComboBox_CA.IsEnabled = $false
	$ComboBox_Filter.IsEnabled = $false
} else {
	#Get current CA
	$Command = "certutil.exe -dump"
	Write-Host $Command
	$Results = Invoke-Expression $Command | Select-String "Config:"
	Foreach($Result in $Results)
	{
		$CAName = (($Result -Split ": ")[1] -replace "$([char]96)|$([char]34)|$([char]39)","").Trim()
		$ComboBox_CA.Items.Add($CAName) | Out-Null
		if($ComboBox_CA.SelectedItem.count -eq 0)
		{
			$ComboBox_CA.SelectedItem = $CAName
		}
	}

	#Load default filters
	$Filters = Get-Content "CertoNiuchacz.cfg" -ErrorAction SilentlyContinue
	Foreach($Filter in $Filters)
	{
		$ComboBox_Filter.Items.Add($Filter) | Out-Null
	}


	$Method_Connect = $Button_Connect.add_click
	$Method_Connect.Invoke({New-RemoteSession})

	$Method_Search = $Button_Search.add_click
	$Method_Search.Invoke({Get-Certs})

	$Method_GetScript = $Button_GetScript.add_click
	$Method_GetScript.Invoke({Invoke-GetScript})

	$Method_GetRec = $Button_GetRec.add_click
	$Method_GetRec.Invoke({Invoke-GetRec})

	$Method_RecKey = $Button_RecKey.add_click
	$Method_RecKey.Invoke({Invoke-RecKey})

	$Method_Merge = $Button_Merge.add_click
	$Method_Merge.Invoke({Invoke-Merge})

	$EventResult = $ListView_Result.Add_SelectionChanged
	$EventResult.Invoke({Set-ButtonStatus})
}

$Window.ShowDialog() | Out-Null

