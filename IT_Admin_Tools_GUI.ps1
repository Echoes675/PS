$version = "1.0.0"

#ERASE ALL THIS AND PUT XAML BELOW between the @" "@
$inputXML = @"
<Window x:Name="Credentials" x:Class="IT_Admin_Tool.Window1"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:IT_Admin_Tool"
        mc:Ignorable="d"
        Title="Credentials" Height="256.658" Width="476.259">
    <Grid Background="#FFE5E5E5">
        <TextBox x:Name="textbox_cred_username" HorizontalAlignment="Left" Height="26" Margin="113,75,0,0" TextWrapping="Wrap" Text="terexlocal\" VerticalAlignment="Top" Width="250"/>
        <Label x:Name="label_cred_username" Content="Username:" HorizontalAlignment="Left" Margin="42,75,0,0" VerticalAlignment="Top" Height="26"/>
        <Label x:Name="label_cred_password" Content="Password:" HorizontalAlignment="Left" Margin="45,121,0,0" VerticalAlignment="Top" Height="26"/>
        <TextBlock x:Name="txtblock_cred_text" HorizontalAlignment="Left" Margin="66,36,0,0" TextWrapping="Wrap" Text="Please enter your A1/A2/A4/A5 username and password" VerticalAlignment="Top"/>
        <PasswordBox x:Name="textbox_cred_password" HorizontalAlignment="Left" Margin="113,121,0,0" VerticalAlignment="Top" Height="26" Width="250"/>
        <Button x:Name="btn_cred_OK" Content="OK" HorizontalAlignment="Left" Margin="312,176,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
    </Grid>
</Window>
"@       
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
Get-FormVariables
 
#===========================================================================
# Actually make the objects work
#===========================================================================
 function Run-Preload{
    
    Check-Username
    Check-Password
    
    if ((Test-Credential) -eq $True) 
        { 
            Load-MainWindow
        }

    if ((Test-Credential) -eq $False) 
        {
            Set-Credential
            $username = cat "C:\scripts\1\$env:USERNAME.txt"
            $password = Get-Content "C:\scripts\1\oc.txt" | ConvertTo-SecureString
            $credential = new-object System.Management.Automation.PSCredential $username, $password
            Run-Preload
            
        }
}

function Check-Username{
    
    if ((Test-Path "C:\scripts\1\$env:USERNAME.txt") -eq $False)
        {
        Set-Credential
        } ELSE {
                $username = cat "C:\scripts\1\$env:USERNAME.txt"
                If ($username -eq $null){
                                        Set-Credential
                                        }
                }

}

function Check-Password{
    
    if ((Test-Path "C:\scripts\1\oc.txt") -eq $False)
        {
        Set-Credential
        } ELSE {
                $password = Get-Content "C:\scripts\1\oc.txt" | ConvertTo-SecureString
                If ($password -eq $null)
                    {
                    Set-Credential
                    }
                }
}

function Test-Credential {

    <#
    .SYNOPSIS
        Takes a PSCredential object and validates it against the domain (or local machine, or ADAM instance).

    .PARAMETER cred
        A PScredential object with the username/password you wish to test. Typically this is generated using the Get-Credential cmdlet. Accepts pipeline input.

    .PARAMETER context
        An optional parameter specifying what type of credential this is. Possible values are 'Domain','Machine',and 'ApplicationDirectory.' The default is 'Domain.'

    .OUTPUTS
        A boolean, indicating whether the credentials were successfully validated.

    #>

    param(
        [parameter()][validateset('Domain','Machine','ApplicationDirectory')]
        [string]$context = 'Domain'
    )
    begin {
        Add-Type -assemblyname system.DirectoryServices.accountmanagement
        $DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext([System.DirectoryServices.AccountManagement.ContextType]::$context) 
    }
    process {
        $username = cat "C:\scripts\1\$env:USERNAME.txt"
        $password = Get-Content "C:\scripts\1\oc.txt" | ConvertTo-SecureString
        $credential = new-object System.Management.Automation.PSCredential $username, $password
        $DS.ValidateCredentials($credential.UserName, $credential.GetNetworkCredential().password)
    }
}

function Set-Credential{
    $Form.ShowDialog() | out-null
}

$WPFbtn_cred_OK.Add_Click({
    
    $WPFtextbox_cred_username.Text | out-file "C:\scripts\1\$env:USERNAME.txt"
    $WPFtextbox_cred_password.Password| ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | out-file "C:\scripts\1\oc.txt"
    
    $Form.Hide() | out-null
    #Run-Preload
    
})

function Load-MainWindow{

        $username = cat "C:\scripts\1\$env:USERNAME.txt"
        $password = Get-Content "C:\scripts\1\oc.txt" | ConvertTo-SecureString
        $credential = new-object System.Management.Automation.PSCredential $username, $password

        Start-Process Powershell -WorkingDirectory "C:\scripts\" -ArgumentList ("-file c:\scripts\IT_Admin_Tool\std_admin_commands_GUI_current.ps1") -Credential $credential
                
}

#Sample entry of how to add data to a field
 
#$vmpicklistView.items.Add([pscustomobject]@{'VMName'=($_).Name;Status=$_.Status;Other="Yes"})
 
#===========================================================================
# Shows the form
#===========================================================================
CLS
Write-Host -ForegroundColor Cyan "IT Admin Tool launcher $Version"
Write-Host -ForegroundColor Red "Do NOT close this window. Please wait for the GUI to load."
Run-Preload
