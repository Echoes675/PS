$Global:syncHash = [hashtable]::Synchronized(@{})
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)

# Load WPF assembly if necessary
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

$psCmd = [PowerShell]::Create().AddScript({
    [xml]$xaml = @"
<Window 
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="IT Admin Tool" Height="468.109" Width="525" ResizeMode="NoResize">
    <Grid>
        <TabControl Margin="0,0,0.4,0.4">
            <TabItem x:Name="tab_launcher" Header="Launcher">
                <Grid Background="#FFE5E5E5
                      ">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="btn_ADUaC" Content="AD Users &amp; Computers" HorizontalAlignment="Left" Margin="97,51,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Button x:Name="btn_PowershellISE" Content="Powershell ISE" HorizontalAlignment="Left" Margin="97,146,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Button x:Name="btn_CMD" Content="CMD" HorizontalAlignment="Left" Margin="362,51,0,0" Width="129" Height="26" VerticalAlignment="Top"/>
                    <Button x:Name="btn_compmgmt" Content="Computer Mgt" HorizontalAlignment="Left" Margin="362,146,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Button x:Name="btn_pulseadmin" Content="Pulse Admin" HorizontalAlignment="Left" Margin="97,247,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Button x:Name="btn_IE" Content="Internet Explorer" HorizontalAlignment="Left" Margin="362,344,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Button x:Name="btn_lockout" Content="Lockout Status" HorizontalAlignment="Left" Margin="97,343,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Button x:Name="btn_PulseLicense" Content="Pulse License" HorizontalAlignment="Left" Margin="362,247,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_ADUaC" HorizontalAlignment="Left" Height="60" Margin="22,33,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\ADUaC.png"/>
                    <Image x:Name="img_CMD" HorizontalAlignment="Left" Height="60" Margin="274,28,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\cmd.png"/>
                    <Image x:Name="img_PowershellISE" HorizontalAlignment="Left" Height="60" Margin="22,126,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\powershell_ISE.png"/>
                    <Image x:Name="img_compmgmt" HorizontalAlignment="Left" Height="60" Margin="274,126,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\compmgmt.png"/>
                    <Image x:Name="img_lockout" HorizontalAlignment="Left" Height="60" Margin="22,328,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Account_Lockout.png"/>
                    <Image x:Name="img_IE" HorizontalAlignment="Left" Height="60" Margin="274,328,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\IE.png"/>
                    <Image x:Name="img_pulseadmin" HorizontalAlignment="Left" Height="60" Margin="22,231,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\pulse-admin.png"/>
                    <Image x:Name="img_pulselicense" HorizontalAlignment="Left" Height="60" Margin="274,231,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\pulse-license.png"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="tab_ADtools" Header="AD commands">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <GroupBox x:Name="grpbox_ADCmds_computers" Header="Computers" HorizontalAlignment="Left" Height="124" Margin="10,151,0,0" VerticalAlignment="Top" Width="483"/>
                    <GroupBox x:Name="grpbox_ADCmds_Users" Header="Users" HorizontalAlignment="Left" Height="110" Margin="10,10,0,0" VerticalAlignment="Top" Width="483"/>
                    <TextBox x:Name="txtbox_ADCmds_username" Height="26" Margin="113,33,0,0" TextWrapping="Wrap" VerticalAlignment="Top" HorizontalAlignment="Left" Width="369"/>
                    <Label x:Name="label_Username" Content="Username:" HorizontalAlignment="Left" Margin="42,33,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="btn_unlock" Content="Unlock account" HorizontalAlignment="Left" Margin="113,74,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Button x:Name="btn_resetPWD" Content="Password Reset" HorizontalAlignment="Left" Margin="249,74,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <TextBox x:Name="txtbox_ADCmds_computername" HorizontalAlignment="Left" Height="26" Margin="76,189,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="168"/>
                    <Label x:Name="label_computername" Content="Computer&#xA;name:" HorizontalAlignment="Left" Margin="10,175,0,0" VerticalAlignment="Top" Height="40" Width="66"/>
                    <Button x:Name="btn_exportGroups" Content="Export Groups" HorizontalAlignment="Left" Margin="382,74,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <ComboBox x:Name="combobox_ADCmds_site" HorizontalAlignment="Left" Margin="65,234,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Label x:Name="label_Site" Content="Site:" HorizontalAlignment="Left" Margin="28,234,0,0" VerticalAlignment="Top" Height="26" RenderTransformOrigin="0.911,0.674"/>
                    <Button x:Name="btn_computeracc" Content="Create computer account" HorizontalAlignment="Left" Margin="332,234,0,0" VerticalAlignment="Top" Width="150" Height="26" IsEnabled="False"/>
                    <TextBox x:Name="txtbox_ADCmds_group" HorizontalAlignment="Left" Height="26" Margin="113,333,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="244"/>
                    <Label x:Name="label_group" Content="Group name:" HorizontalAlignment="Left" Margin="28,333,0,0" VerticalAlignment="Top" Height="26" Width="80"/>
                    <Button x:Name="btn_exportMembers" Content="Export Members" HorizontalAlignment="Left" Margin="382,333,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Button x:Name="btn_loggedOn" Content="Who is logged on" HorizontalAlignment="Left" Margin="249,189,0,0" VerticalAlignment="Top" Width="117" Height="26"/>
                    <Button x:Name="btn_lastBoot" Content="Last Boot" HorizontalAlignment="Left" Margin="382,189,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <ComboBox x:Name="combobox_Type" HorizontalAlignment="Left" Margin="218,234,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Label x:Name="label_type" Content="Type:" HorizontalAlignment="Left" Margin="175,234,0,0" VerticalAlignment="Top"/>
                    <GroupBox x:Name="grpbox_ADCmds_groups" Header="Groups" HorizontalAlignment="Left" Height="81" Margin="10,299,0,0" VerticalAlignment="Top" Width="483"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="tab_NewADUser" Header="New AD User">
                <Grid Background="#FFE5E5E5">
                    <ComboBox x:Name="combo_ADNU_Site" HorizontalAlignment="Left" Margin="71,30,0,0" VerticalAlignment="Top" Width="120" Height="26"/>
                    <Label Content="Site:" HorizontalAlignment="Left" Margin="28,30,0,0" VerticalAlignment="Top" Width="43"/>
                    <TextBox x:Name="txtbox_ADNU_firstname" HorizontalAlignment="Left" Height="26" Margin="74,82,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <Label x:Name="label_firstname" Content="First Name:" HorizontalAlignment="Left" Margin="1,81,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtbox_ADNU_lastname" HorizontalAlignment="Left" Height="26" Margin="71,146,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <Label Content="Last Name:" HorizontalAlignment="Left" Margin="2,146,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtbox_ADNU_initial" HorizontalAlignment="Left" Height="26" Margin="71,201,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
                    <Label x:Name="label_ADNU_initial" Content="Initial:" HorizontalAlignment="Left" Margin="28,201,0,0" VerticalAlignment="Top"/>
                    <GroupBox x:Name="grpbox_ADNU_Vaults" Header="Pulse Vaults" HorizontalAlignment="Left" Height="80" Margin="14,260,0,0" VerticalAlignment="Top" Width="180">
                        <CheckBox x:Name="chkbox_ADNU_Site1City" Content="Site1 Vault" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="80"/>
                    </GroupBox>
                    <CheckBox x:Name="chkbox_ADNU_Site2City" Content="Site2 VAult" HorizontalAlignment="Left" Margin="30,310,0,0" VerticalAlignment="Top" Width="80"/>
                    <GroupBox x:Name="grpbox_ADNU_expiry" Header="Expiry" HorizontalAlignment="Left" Height="49" Margin="259,22,0,0" VerticalAlignment="Top" Width="228">
                        <CheckBox x:Name="chkbox_ADNU_temp" Content="Temp User" HorizontalAlignment="Right" Margin="0,5,132.6,0" VerticalAlignment="Top"/>
                    </GroupBox>
                    <DatePicker x:Name="datepkr_ADNU_expiry" HorizontalAlignment="Left" Margin="363,39,0,0" VerticalAlignment="Top" Width="110" IsEnabled="False"/>
                    <Label x:Name="label_ADNU_logonname" Content="Logon Name:" HorizontalAlignment="Left" Margin="209,82,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtbox_ADNU_logonname" HorizontalAlignment="Left" Height="26" Margin="289,82,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200"/>
                    <Label x:Name="label_ADNU_displayname" Content="Display Name:" HorizontalAlignment="Left" Margin="204,146,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtbox_ADNU_displayname" HorizontalAlignment="Left" Height="26" Margin="289,147,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200"/>
                    <TextBox x:Name="txtbox_ADNU_manager" HorizontalAlignment="Left" Height="26" Margin="289,202,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200"/>
                    <TextBox x:Name="txtbox_ADNU_title" HorizontalAlignment="Left" Height="26" Margin="289,262,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200"/>
                    <Label x:Name="label_ADNU_title" Content="Job Title:" HorizontalAlignment="Left" Margin="227,261,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtbox_ADNU_dept" HorizontalAlignment="Left" Height="26" Margin="289,318,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200"/>
                    <Label x:Name="label_ADNU_dept" Content="Department:" HorizontalAlignment="Left" Margin="208,318,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="label_ADNI_manager" Content="Manager Name:" HorizontalAlignment="Left" Margin="196,202,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.504,0.465"/>
                    <Button x:Name="btn_ADNU_resetform" Content="Reset" HorizontalAlignment="Left" Margin="4,363,0,0" VerticalAlignment="Top" Width="110" Height="26"/>
                    <Button x:Name="btn_ADNU_validate" Content="Validate" HorizontalAlignment="Left" Margin="128,363,0,0" VerticalAlignment="Top" Width="110" Height="26"/>
                    <Button x:Name="btn_ADNU_edit" Content="Edit" HorizontalAlignment="Left" Margin="253,363,0,0" VerticalAlignment="Top" Width="110" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_ADNU_Create" Content="Create user" HorizontalAlignment="Left" Margin="379,363,0,0" VerticalAlignment="Top" Width="110" Height="26" IsEnabled="False"/>
                    <TextBlock x:Name="txtblock_ADNU_errorexpiry" HorizontalAlignment="Left" Margin="475,43,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_ADNU_errorsite" HorizontalAlignment="Left" Margin="195,35,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_ADNU_errorfirst" HorizontalAlignment="Left" Margin="195,87,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" Width="5" FontWeight="Bold" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_ADNU_errorlast" HorizontalAlignment="Left" Margin="195,151,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_ADNU_errorlogonunique" HorizontalAlignment="Left" Margin="289,108,0,0" TextWrapping="Wrap" Text="Logon name already exists in AD. Please enter a unique logon name" VerticalAlignment="Top" Width="198" Foreground="Red" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_ADNU_errormgrunique" HorizontalAlignment="Left" Margin="289,237,0,0" TextWrapping="Wrap" Text="Manager does not exist in AD. Please enter valid manager" VerticalAlignment="Top" Width="200" Foreground="Red" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_ADNU_errorlogon" HorizontalAlignment="Left" Margin="493,86,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_ADNU_errordisp" HorizontalAlignment="Left" Margin="493,151,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_ADNU_errormgr" HorizontalAlignment="Left" Margin="493,207,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_ADNU_errortitle" HorizontalAlignment="Left" Margin="493,267,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_ADNU_errordept" HorizontalAlignment="Left" Margin="493,324,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="tab_ExportUserInfo" Header="Export AD User Info">
                <Grid Background="#FFE5E5E5">
                    <GroupBox x:Name="grpbox_EADUI_Scope" Header="Scope" HorizontalAlignment="Left" Height="126" Margin="10,10,0,0" VerticalAlignment="Top" Width="487"/>
                    <RadioButton x:Name="radio_EADUI_User" Content="Single user" HorizontalAlignment="Left" Margin="29,36,0,0" VerticalAlignment="Top" IsChecked="True"/>
                    <TextBox x:Name="txtbox_EADUI_username" Height="26" Margin="275,31,0,0" TextWrapping="Wrap" VerticalAlignment="Top" HorizontalAlignment="Left" Width="200"/>
                    <TextBlock x:Name="txtblock_EADUI_errorusername" HorizontalAlignment="Left" Margin="480,36,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_EADUI_erroruserunique" HorizontalAlignment="Left" Margin="279,56,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200" Foreground="Red" Visibility="Hidden"><Run Text="User does not exist in AD. Please enter valid "/><Run Text="user"/><Run Text="name"/></TextBlock>
                    <RadioButton x:Name="radio_EADUI_OU" Content="Users in OU" HorizontalAlignment="Left" Margin="29,103,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="label_EADUI_username" Content="Username:" HorizontalAlignment="Left" Height="25" Margin="204,31,0,0" VerticalAlignment="Top" Width="66"/>
                    <Label x:Name="label_EADUI_username_Copy" Content="Site:" HorizontalAlignment="Left" Height="25" Margin="237,93,0,0" VerticalAlignment="Top" Width="33"/>
                    <CheckBox x:Name="chkbox_EADUI_AllAtts" Content="All Attributes" HorizontalAlignment="Left" Margin="17,141,0,0" VerticalAlignment="Top" IsChecked="True"/>
                    <CheckBox x:Name="chkbox_EADUI_common" Content="Common Attributes" HorizontalAlignment="Left" Margin="138,141,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_AdditionalAtts" Content="Additional Attributes" HorizontalAlignment="Left" Margin="279,141,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                    <GroupBox x:Name="grpbox_EADUI_common" Header="Common Attributes" HorizontalAlignment="Left" Height="155" Margin="10,162,0,0" VerticalAlignment="Top" Width="229"/>
                    <GroupBox x:Name="grpbox_EADUI_additionalatts" Header="Additional Attributes" HorizontalAlignment="Left" Height="195" Margin="258,162,0,0" VerticalAlignment="Top" Width="238"/>
                    <ComboBox x:Name="combo_EADUI_site" HorizontalAlignment="Left" Margin="275,93,0,0" VerticalAlignment="Top" Width="120" Height="26" IsEnabled="False"/>
                    <TextBlock x:Name="txtblock_EADUI_errorsite" HorizontalAlignment="Left" Margin="400,98,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                    <CheckBox x:Name="chkbox_EADUI_userPrincipalName" Content="Logon Name" HorizontalAlignment="Left" Margin="17,184,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_samaccname" Content="Logon Pre 2000" HorizontalAlignment="Left" Margin="17,205,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_display" Content="Display Name" HorizontalAlignment="Left" Margin="17,226,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_givenName" Content="First Name" HorizontalAlignment="Left" Margin="17,247,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_initials" Content="Initials" HorizontalAlignment="Left" Margin="17,268,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_sn" Content="Last Name" HorizontalAlignment="Left" Margin="17,289,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_email" Content="Email" HorizontalAlignment="Left" Margin="138,184,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_title" Content="Title" HorizontalAlignment="Left" Margin="138,205,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_dept" Content="Department" HorizontalAlignment="Left" Margin="138,226,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_desc" Content="Description" HorizontalAlignment="Left" Margin="138,247,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_mgr" Content="Manager" HorizontalAlignment="Left" Margin="138,268,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_alias" Content="Mail alias" HorizontalAlignment="Left" Margin="279,184,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_expiry" Content="Account Expires" HorizontalAlignment="Left" Margin="279,205,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_cn" Content="Canonical Name" HorizontalAlignment="Left" Margin="279,226,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_loginscpt" Content="Login Script" HorizontalAlignment="Left" Margin="279,247,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_hmefolder" Content="Home Folder" HorizontalAlignment="Left" Margin="279,268,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_hmedrv" Content="Home Drive" HorizontalAlignment="Left" Margin="279,289,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_logonto" Content="Log on to" HorizontalAlignment="Left" Margin="279,310,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_company" Content="Company" HorizontalAlignment="Left" Margin="389,184,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_Country" Content="Country" HorizontalAlignment="Left" Margin="389,205,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_office" Content="Office" HorizontalAlignment="Left" Margin="389,226,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_telephone" Content="Telephone" HorizontalAlignment="Left" Margin="389,247,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_street" Content="Street" HorizontalAlignment="Left" Margin="389,268,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_POBox" Content="Post Code" HorizontalAlignment="Left" Margin="389,289,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_City" Content="City" HorizontalAlignment="Left" Margin="389,310,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_State" Content="State/Province" HorizontalAlignment="Left" Margin="389,331,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <Button x:Name="btn_EADUI_export" Content="Export to CSV" HorizontalAlignment="Left" Margin="374,366,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@


    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    
    $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )

    [xml]$XAML = $xaml
        $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
        #Find all of the form types and add them as members to the synchash
        $syncHash.Add($_.Name,$syncHash.Window.FindName($_.Name) )        

    }

    $Script:JobCleanup = [hashtable]::Synchronized(@{})
    $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

    #region Background runspace to clean up jobs
    $jobCleanup.Flag = $True
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"          
    $newRunspace.Open()        
    $newRunspace.SessionStateProxy.SetVariable("jobCleanup",$jobCleanup)     
    $newRunspace.SessionStateProxy.SetVariable("jobs",$jobs) 
    $jobCleanup.PowerShell = [PowerShell]::Create().AddScript({
        #Routine to handle completed runspaces
        Do {    
            Foreach($runspace in $jobs) {            
                If ($runspace.Runspace.isCompleted) {
                    [void]$runspace.powershell.EndInvoke($runspace.Runspace)
                    $runspace.powershell.dispose()
                    $runspace.Runspace = $null
                    $runspace.powershell = $null               
                } 
            }
            #Clean out unused runspace jobs
            $temphash = $jobs.clone()
            $temphash | Where {
                $_.runspace -eq $Null
            } | ForEach {
                $jobs.remove($_)
            }        
            Start-Sleep -Seconds 1     
        } while ($jobCleanup.Flag)
    })
    $jobCleanup.PowerShell.Runspace = $newRunspace
    $jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()  
    
    #endregion Background runspace to clean up jobs

#######################################################################################################################################################################################
#
#       Stuff Happens here
#
#######################################################################################################################################################################################


### Load options into combo boxes ###

$Sitecode_str = " " , "Site1" , "Site2" , "Site3" , "Site4" , "Site5" , "Site6"

## AD Tools Tab
$Sitecode_str | ForEach-object {$syncHash.combobox_ADCmds_site.AddChild($_)}
" " , "Workstation" , "Notebook" | ForEach-object {$syncHash.comboBox_type.AddChild($_)}

## Create new AD user tab
$Sitecode_str | ForEach-object {$syncHash.combo_ADNU_site.AddChild($_)}

## Export AD User Info tab
$Sitecode_str | ForEach-object {$syncHash.combo_EADUI_site.AddChild($_)}


#=======================#
# Main Launcher Buttons #
#=======================#


### AD Users & Computers ###
$syncHash.btn_ADUac.Add_Click({ 
    Start-Process dsa /32
})
 

### Command Prompt ###
$syncHash.btn_CMD.Add_Click({
    Start-Process cmd
}) 


### Computer Management ###
$syncHash.btn_compmgmt.Add_Click({ 
    Start-Process compmgmt.msc
}) 


### Internet Explorer ###
$syncHash.btn_IE.Add_Click({
    Start-process iexplore.exe
}) 


### Account lockout status ###
$syncHash.btn_lockout.Add_Click({ 
    Start-Process lockoutstatus.exe
}) 


### Powershell ISE ###
$syncHash.btn_PowershellISE.Add_Click({ 
    Start-Process PowerShell_ISE.exe
})


### Pulse Administration ###
$syncHash.btn_pulseadmin.Add_Click({ 
    DO {
        $reg = $null
        $SID = Get-ADUser -Identity $env:USERNAME | select SID
        $Sidft = $SID.SID
        $regquery = "HKU\" + "$SIDft" + "\Software\VB and VBA Program Settings\Pulse Data Manager"

        IF (!(REG QUERY $regquery)) 
            { 
            regedit.exe /s "C:\scripts\IT_Admin_Tool\pulse_vaults_A2.reg" 
            $reg = "m"
            }
    }
    While ($reg -eq "m") 

    Start-Process "C:\Program Files (x86)\Pulse PLM\Administration.exe"
})


### Pulse License upload ###
$syncHash.btn_pulselicense.Add_Click({
    Start-Process "C:\Program Files (x86)\Pulse PLM\PLM Load Server License.exe"
})

#======================#
# AD Tools             #
#======================#

### Unlock user account ###
$syncHash.btn_unlock.Add_Click({
    #Start-Job -Name Sleeping -ScriptBlock {start-sleep 5}
    #while ((Get-Job Sleeping).State -eq 'Running'){
        $x+= "."
    #region Boe's Additions
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"          
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash) 
    $PowerShell = [PowerShell]::Create().AddScript({

    function Read-Textbox {
    param($arg)
    $Normal = [System.Windows.Threading.DispatcherPriority]::Normal
    $SyncHash.Window.Dispatcher.Invoke($Normal,[action]{$Global:Return = $SyncHash.$arg.Text})
    $Global:Return
    remove-variable -Name Return -scope Global
    }

    $username = Read-Textbox -Arg txtbox_ADCmds_username | Out-String

    #$syncHash.txtbox_ADCmds_username = $syncHash.Window.FindName("txtbox_ADCmds_username")
    #$username = $syncHash.Window.FindName("txtbox_ADCmds_username")

    If ($Username.text -eq "")
        {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "ERROR: Invalid user name"
        $Messageboxbody = "Username cannot be blank. Please enter a valid username"
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
        } 
        Else {
             $username = $username.trim()
             $username =$username.replace(' ','.')
             $unlockerror = $null
             $attempt = 0
                If (dsquery user -samid $username)
                        {
                        Do {
                            If ((Get-ADUser  $username -Properties LockedOut | Select-Object LockedOut) -eq $true)
                                {
                                    Unlock-ADAccount -identity $username
                                    If ((Get-ADUser  $username -Properties LockedOut | Select-Object LockedOut) -eq $false)
                                    {
                                    $attempt = 5
                                    $ButtonType = [System.Windows.MessageBoxButton]::OK
                                    $MessageboxTitle = "Unlock user account"
                                    $Messageboxbody = "$Username" + "'s account has been unlocked!"
                                    $MessageIcon = [System.Windows.MessageBoxImage]::Information
                                    [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                                    }
                                    Else {
                                        $attempt = $attempt++
                                        If ($attempt -ge 4)
                                            {
                                            $ButtonType = [System.Windows.MessageBoxButton]::OK
                                            $MessageboxTitle = "Unlock user account"
                                            $Messageboxbody = "$Username" + "'s account cannot be unlocked!"
                                            $MessageIcon = [System.Windows.MessageBoxImage]::Error
                                            [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                                            }
                                        }
                                    }
                            Else {
                                $attempt = 5
                                $ButtonType = [System.Windows.MessageBoxButton]::OK
                                $MessageboxTitle = "Unlock user account"
                                $Messageboxbody = "$Username" + "'s account is not locked. No action taken!"
                                $MessageIcon = [System.Windows.MessageBoxImage]::Information
                                [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                                }
                                }
                        While ($attempt -lt 4)
                                
                        } ELSE
                        {
                        $ButtonType = [System.Windows.MessageBoxButton]::OK
                        $MessageboxTitle = "ERROR: Invalid user name"
                        $Messageboxbody = "$Username does not exist in AD. Please enter a valid username"
                        $MessageIcon = [System.Windows.MessageBoxImage]::Error
                        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                        }
    Get-module | Remove-Module        
            }
                                                    })
        $PowerShell.Runspace = $newRunspace
        [void]$Jobs.Add((
            [pscustomobject]@{
                PowerShell = $PowerShell
                Runspace = $PowerShell.BeginInvoke()
            }
                        ))
})

### Password reset ###
$syncHash.btn_resetPWD.Add_Click({
    #Start-Job -Name Sleeping -ScriptBlock {start-sleep 5}
    #while ((Get-Job Sleeping).State -eq 'Running'){
        $x+= "."
    #region Boe's Additions
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"          
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash) 
    $PowerShell = [PowerShell]::Create().AddScript({

    $username = $syncHash.txtbox_ADCmds_username.Text

    If ($Username -eq "")
        {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "ERROR: Invalid user name"
        $Messageboxbody = "Username cannot be blank. Please enter a valid username"
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
        } 
        Else {
             $username = $username.trim()
             $username =$username.replace(' ','.')
             $attempt = 0
                If (dsquery user -samid $username)
                        {
                        $Paswd = "CompanyName2017"
                        $SecPaswd= ConvertTo-SecureString –String "$Paswd" –AsPlainText –Force
                        Set-ADAccountPassword -Reset -NewPassword $SecPaswd –Identity $username
                        Unlock-ADAccount -identity $username
                        Set-ADUser –Identity $username –ChangePasswordAtLogon $true
                        $ButtonType = [System.Windows.MessageBoxButton]::OK
                        $MessageboxTitle = "Unlock user account"
                        $Messageboxbody = "$Username" + "'s password has been reset to: $Paswd"
                        $MessageIcon = [System.Windows.MessageBoxImage]::Information
                        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                                
                        } ELSE
                        {
                        $ButtonType = [System.Windows.MessageBoxButton]::OK
                        $MessageboxTitle = "ERROR: Invalid user name"
                        $Messageboxbody = "$Username does not exist in AD. Please enter a valid username"
                        $MessageIcon = [System.Windows.MessageBoxImage]::Error
                        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                        }
    Get-module | Remove-Module        
            }

                                })
    $PowerShell.Runspace = $newRunspace
    [void]$Jobs.Add((
        [pscustomobject]@{
            PowerShell = $PowerShell
            Runspace = $PowerShell.BeginInvoke()
        }
    ))
})

### Export user's group membership to CSV ###
$syncHash.btn_exportGroups.Add_Click({
    #Start-Job -Name Sleeping -ScriptBlock {start-sleep 5}
    #while ((Get-Job Sleeping).State -eq 'Running'){
        $x+= "."
    #region Boe's Additions
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"          
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash) 
    $PowerShell = [PowerShell]::Create().AddScript({    

    $Username = $syncHash.txtbox_ADCmds_username.Text
    $username = $username.trim()
    $username =$username.replace(' ','.')
    
    If ($Username -eq "")
        {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "ERROR: Invalid user name"
        $Messageboxbody = "Username cannot be blank. Please enter a valid username"
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
        }
        ELSE {
                IF (dsquery user -samid $username)
                    {
                    Get-ADPrincipalGroupMembership $Username | select name | Export-csv -path "C:\Member of Groups - $Username.csv" -NoTypeInformation
                    $ButtonType = [System.Windows.MessageBoxButton]::OK
                    $MessageboxTitle = "Export a user's group membership"
                    $Messageboxbody = "Group membership for $username has been exported to C:\Member of Groups - $Username.csv"
                    $MessageIcon = [System.Windows.MessageBoxImage]::Information
                    [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)

                    } ELSE
                    {
                    $ButtonType = [System.Windows.MessageBoxButton]::OK
                    $MessageboxTitle = "ERROR: Invalid user name"
                    $Messageboxbody = "$Username does not exist in AD. Please enter a valid username"
                    $MessageIcon = [System.Windows.MessageBoxImage]::Error
                    [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                    }
              }
    Get-module | Remove-Module 
    })
    $PowerShell.Runspace = $newRunspace
    [void]$Jobs.Add((
        [pscustomobject]@{
            PowerShell = $PowerShell
            Runspace = $PowerShell.BeginInvoke()
        }
    ))
})

#######################################################################################################################################################################################
    #region Window Close 
    $syncHash.Window.Add_Closed({
        Write-Verbose 'Halt runspace cleanup job processing'
        $jobCleanup.Flag = $False

        #Stop all runspaces
        $jobCleanup.PowerShell.Dispose()      
    })
    #endregion Window Close 
    #endregion Boe's Additions

    #$x.Host.Runspace.Events.GenerateEvent( "TestClicked", $x.test, $null, "test event")

    #$syncHash.Window.Activate()
    $syncHash.Window.ShowDialog() | Out-Null
    $syncHash.Error = $Error
})
$psCmd.Runspace = $newRunspace
$data = $psCmd.BeginInvoke()