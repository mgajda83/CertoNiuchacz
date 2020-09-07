Add-Type -AssemblyName PresentationFramework
[xml]$xaml = 
@"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="CertoNiuchacz v2.0.0.2" Height="300" Width="500" MinWidth="500" MinHeight="300"  WindowStyle="ToolWindow" SizeToContent="WidthAndHeight">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="81"/>
            <RowDefinition/>
            <RowDefinition Height="30"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="70"/>
            <ColumnDefinition Width="70"/>
            <ColumnDefinition Width="70"/>
        </Grid.ColumnDefinitions>

        <Label Content="Predefined filters:" Height="22" Margin="10,10,10,0" VerticalAlignment="Top" FontSize="9"/>
        <ComboBox x:Name="CA" Margin="3,35,10,0" VerticalAlignment="Top" Height="19" Grid.ColumnSpan="3" Grid.Column="1"/>

        <Label Content="Certification Authority:" Height="22" Margin="3,10,10,0" VerticalAlignment="Top" FontSize="9" Grid.ColumnSpan="3" Grid.Column="1"/>
        <ComboBox x:Name="Filter" Margin="10,35,10,0" VerticalAlignment="Top" Height="19"/>
        <TextBox x:Name="Phrase" Height="19" Margin="10,60,10,0" TextWrapping="Wrap" VerticalAlignment="Top" AutomationProperties.HelpText="Nazwa użytkownika, numer seryjny lub odcisk palca certyfikatu"/>
        <Button x:Name="Search" Content="Search" Height="19" Margin="0,60,7,0" VerticalAlignment="Top" Width="60" HorizontalAlignment="Right" Grid.Column="1"/>

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
        <Label x:Name="Status" Content="" Height="19" Margin="10,0,80,0" VerticalAlignment="Top" FontSize="9" Grid.Row="2"/>
        <Button x:Name="GetScript" Content="GetKey Script" Margin="202,0,10,10" Grid.Row="2" Width="60" HorizontalAlignment="Right" FontSize="8" IsEnabled="False"/>
        <Button x:Name="GetRec" Content="GetKey (.rec)" Margin="0,0,10,10" IsEnabled="False" RenderTransformOrigin="0.533,0.561" Grid.Column="1" Grid.Row="2" FontSize="8"/>
        <Button x:Name="RecKey" Content="RecKey (.p12)" Margin="0,0,10,10" IsEnabled="False" RenderTransformOrigin="0.533,0.561" Grid.Column="2" Grid.Row="2" FontSize="8"/>
        <Button x:Name="Merge" Content="Merge (.pfx)" Margin="0,0,10,10" IsEnabled="False" RenderTransformOrigin="0.533,0.561" Grid.Column="3" Grid.Row="2" FontSize="8"/>

    </Grid>
</Window>
"@

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

Function Set-Filter
{
    $TextBox_Phrase.Text = $ComboBox_Filter.SelectedItem
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


Function Get-Certs
{
    $OrgText = $Window.Title
    $Window.Title += " - Search..."

    $SearchText = $TextBox_Phrase.Text
    $CAText = $ComboBox_CA.SelectedItem
    $Command = 'certutil.exe -config "' + $CAText + '" -view -restrict "' + $SearchText + '" -out "UPN,SerialNumber,RequestID,NotBefore,NotAfter,CertificateTemplate,CertificateHash"'
    Write-Host $Command
    $Result = Invoke-Expression $Command

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

    foreach ($SelectedItem in $ListView_Result.SelectedItems) 
    {
        $SearchText = $SelectedItem.SerialNumber
        $Command = 'certutil.exe -config "' + $CAText + '" -f -getkey ' + $SearchText + ' script ' + $SearchText + '.log'
        Write-Host $Command
        $Result = Invoke-Expression $Command

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
            $Result = Invoke-Expression $Command

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
            $Result = Invoke-Expression $Command

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
            $Result = Invoke-Expression $Command

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
$TextBox_Phrase = $Window.FindName("Phrase")
$Button_Search = $Window.FindName("Search")
$ListView_Result = $Window.FindName("Result")
$Button_GetScript= $Window.FindName("GetScript")
$Button_GetRec = $Window.FindName("GetRec")
$Button_RecKey = $Window.FindName("RecKey")
$Button_Merge = $Window.FindName("Merge")

#Change script location
Set-Location -Path $([System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition))

#Check admin role
$User = [Security.Principal.WindowsIdentity]::GetCurrent()
$Role = (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if(!$Role)
{
    $TextBox_Phrase.Text = "Brak uprawnień !!!"

    $Button_Search.IsEnabled = $false
    $TextBox_Phrase.IsEnabled = $false
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
    $Filters = Get-Content "CertoNiuchacz.cfg"
    Foreach($Filter in $Filters)
    {
        $ComboBox_Filter.Items.Add($Filter) | Out-Null
    }



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

    $TextBox_Phrase.Add_KeyDown({
        if ($_.Key -eq "Enter") {
            Get-Certs
        }
    })

    $EventFilter = $ComboBox_Filter.Add_SelectionChanged
    $EventFilter.Invoke({Set-Filter})


    $EventResult = $ListView_Result.Add_SelectionChanged
    $EventResult.Invoke({Set-ButtonStatus})
}

$Window.ShowDialog() | Out-Null

