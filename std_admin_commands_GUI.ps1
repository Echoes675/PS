$Global:version = "Version 1.2.1"
$Global:installpath = $PSScriptRoot

Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

#ERASE ALL THIS AND PUT XAML BELOW between the @" "@
$inputXML = @"
<Window x:Name="IT_Admin_Tool" x:Class="IT_Admin_Tool.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:IT_Admin_Tool"
        mc:Ignorable="d"
        Title="IT Admin Tool" Height="699.609" Width="800" ResizeMode="NoResize"
        SizeToContent="WidthAndHeight" WindowStartupLocation="CenterScreen">
    <Grid Background="#FFE5E5E5">
        <TabControl Margin="0,34,240,32">
            <TabItem x:Name="tab_launcher" Header="Launcher">
                <Grid Background="#FFE5E5E5                       ">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="btn_A1" Content="Unassigned" HorizontalAlignment="Left" Margin="97,61,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_A1" HorizontalAlignment="Left" Height="60" Margin="22,43,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>
                    <Button x:Name="btn_A2" Content="Unassigned" HorizontalAlignment="Left" Margin="97,146,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_A2" HorizontalAlignment="Left" Height="60" Margin="22,126,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>
                    <Button x:Name="btn_A3" Content="Unassigned" HorizontalAlignment="Left" Margin="97,235,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_A3" HorizontalAlignment="Left" Height="60" Margin="22,219,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>
                    <Button x:Name="btn_A4" Content="Unassigned" HorizontalAlignment="Left" Margin="97,323,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_A4" HorizontalAlignment="Left" Height="60" Margin="22,308,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>
                    <Button x:Name="btn_A5" Content="Unassigned" HorizontalAlignment="Left" Margin="97,412,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_A5" HorizontalAlignment="Left" Height="60" Margin="22,400,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>
                    <Button x:Name="btn_A6" Content="Unassigned" HorizontalAlignment="Left" Margin="97,500,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_A6" HorizontalAlignment="Left" Height="60" Margin="22,487,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>

                    <Button x:Name="btn_B1" Content="Unassigned" HorizontalAlignment="Left" Margin="362,61,0,0" Width="129" Height="26" VerticalAlignment="Top"/>
                    <Image x:Name="img_B1" HorizontalAlignment="Left" Height="60" Margin="274,38,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>
                    <Button x:Name="btn_B2" Content="Unassigned" HorizontalAlignment="Left" Margin="362,146,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_B2" HorizontalAlignment="Left" Height="60" Margin="274,126,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>
                    <Button x:Name="btn_B3" Content="Unassigned" HorizontalAlignment="Left" Margin="362,235,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_B3" HorizontalAlignment="Left" Height="60" Margin="274,219,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>
                    <Button x:Name="btn_B4" Content="Unassigned" HorizontalAlignment="Left" Margin="362,324,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_B4" HorizontalAlignment="Left" Height="60" Margin="274,308,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>
                    <Button x:Name="btn_B5" Content="Unassigned" HorizontalAlignment="Left" Margin="362,412,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_B5" HorizontalAlignment="Left" Height="60" Margin="274,400,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>
                    <Button x:Name="btn_B6" Content="Unassigned" HorizontalAlignment="Left" Margin="362,500,0,0" VerticalAlignment="Top" Width="129" Height="26"/>
                    <Image x:Name="img_B6" HorizontalAlignment="Left" Height="60" Margin="274,487,0,0" VerticalAlignment="Top" Width="60" Source="C:\Scripts\IT_Admin_Tool\icons\Default.png"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="tab_ADtools" Header="AD commands">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <GroupBox x:Name="grpbox_ADCmds_groups" Header="Groups" HorizontalAlignment="Left" Height="67" Margin="10,393,0,0" VerticalAlignment="Top" Width="483"/>
                    <GroupBox x:Name="grpbox_ADCmds_computers" Header="Computers" HorizontalAlignment="Left" Height="124" Margin="10,264,0,0" VerticalAlignment="Top" Width="483"/>
                    <GroupBox x:Name="grpbox_ADCmds_Users" Header="Users" HorizontalAlignment="Left" Height="115" Margin="10,144,0,0" VerticalAlignment="Top" Width="483"/>
                    
                    <TextBox x:Name="txtbox_ADCmds_username" Height="26" Margin="113,15,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" HorizontalAlignment="Left" Width="203" VerticalScrollBarVisibility="Disabled"/>
                    <Label x:Name="label_Username" Content="Username:" HorizontalAlignment="Left" Margin="42,15,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtbox_ADCmds_computername" HorizontalAlignment="Left" Height="26" Margin="113,41,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="203" VerticalScrollBarVisibility="Disabled"/>
                    <Label x:Name="label_computername" Content="Computer name:" HorizontalAlignment="Left" Margin="10,41,0,0" VerticalAlignment="Top" Height="26" Width="98"/>
                    <TextBox x:Name="txtbox_ADCmds_group" HorizontalAlignment="Left" Height="26" Margin="113,67,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="203" VerticalScrollBarVisibility="Disabled"/>
                    <Label x:Name="label_group" Content="Group name:" HorizontalAlignment="Left" Margin="28,67,0,0" VerticalAlignment="Top" Height="26" Width="80"/>
                    <TextBox x:Name="txtbox_ADCmds_dirpath" HorizontalAlignment="Left" Height="26" Margin="113,93,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="379" VerticalScrollBarVisibility="Disabled"/>
                    <Label x:Name="label_dirpath" Content="Directory Path:" HorizontalAlignment="Left" Margin="19,93,0,0" VerticalAlignment="Top" Height="26" Width="89"/>

                    <Label x:Name="label_Site" Content="Site:" HorizontalAlignment="Left" Margin="360,15,0,0" VerticalAlignment="Top" Height="26" RenderTransformOrigin="0.911,0.674"/>
                    <ComboBox x:Name="combobox_ADCmds_site" HorizontalAlignment="Left" Margin="392,15,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <ComboBox x:Name="combobox_kittype" HorizontalAlignment="Left" Margin="392,41,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Label x:Name="label_kittype" Content="Kit Type:" HorizontalAlignment="Left" Margin="337,41,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="combobox_GrpType" HorizontalAlignment="Left" Margin="392,67,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Label x:Name="label_Grptype" Content="Group Type:" HorizontalAlignment="Left" Margin="317,67,0,0" VerticalAlignment="Top"/>

                    <Button x:Name="btn_unlock" Content="Unlock account" HorizontalAlignment="Left" Margin="28,173,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_resetPWD" Content="Password Reset" HorizontalAlignment="Left" Margin="135,173,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_exportGroups" Content="Export Groups" HorizontalAlignment="Left" Margin="240,173,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_addlocalAdmin" Content="Add Local Admin" HorizontalAlignment="Left" Margin="28,219,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_remlocalAdmin" Content="Rmv Loc Admin" HorizontalAlignment="Left" Margin="135,219,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_addUsertoGroup" Content="Add to Group" HorizontalAlignment="Left" Margin="240,219,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_rmvUserfromGroup" Content="Remove from Grp" HorizontalAlignment="Left" Margin="345,219,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                                        
                    <Button x:Name="btn_loggedOn" Content="Logged on User" HorizontalAlignment="Left" Margin="28,302,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_lastBoot" Content="Last Boot" HorizontalAlignment="Left" Margin="135,302,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_discomputeracc" Content="Disable Account" HorizontalAlignment="Left" Margin="240,302,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_encomputeracc" Content="Enable Account" HorizontalAlignment="Left" Margin="345,302,0,0" VerticalAlignment="Top" Width="105" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_nextavailacc" Content="Next Avail Acc" HorizontalAlignment="Left" Margin="28,348,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_computeracc" Content="Create Account" HorizontalAlignment="Left" Margin="135,348,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_chknetconf" Content="Network Config" HorizontalAlignment="Left" Margin="240,348,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>

                    <Button x:Name="btn_exportMembers" Content="Export Members" HorizontalAlignment="Left" Margin="28,419,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_createGroup" Content="Create Group" HorizontalAlignment="Left" Margin="135,419,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_addGrpAuthor" Content="Add as Author" HorizontalAlignment="Left" Margin="240,419,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_addread" Content="Add as Read Only" HorizontalAlignment="Left" Margin="345,419,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                                        
                </Grid>
            </TabItem>
            <TabItem x:Name="tab_NewADUser" Header="New AD User">
                <Grid Background="#FFE5E5E5">
                    <GroupBox Header="Single User" HorizontalAlignment="Left" Height="346" Margin="10,63,0,0" VerticalAlignment="Top" Width="507"/>
                    <GroupBox Header="Multiple Users" HorizontalAlignment="Left" Height="67" Margin="10,435,0,0" VerticalAlignment="Top" Width="507"/>
                    <ComboBox x:Name="combo_ADNU_Site" HorizontalAlignment="Left" Margin="81,91,0,0" VerticalAlignment="Top" Width="120" Height="26"/>
                    <Label Content="Site:" HorizontalAlignment="Left" Margin="38,91,0,0" VerticalAlignment="Top" Width="43"/>
                    <TextBox x:Name="txtbox_ADNU_firstname" HorizontalAlignment="Left" Height="26" Margin="84,143,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" VerticalScrollBarVisibility="Disabled"/>
                    <Label x:Name="label_firstname" Content="First Name:" HorizontalAlignment="Left" Margin="11,142,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtbox_ADNU_lastname" HorizontalAlignment="Left" Height="26" Margin="81,207,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" VerticalScrollBarVisibility="Disabled"/>
                    <Label Content="Last Name:" HorizontalAlignment="Left" Margin="12,207,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtbox_ADNU_initial" HorizontalAlignment="Left" Height="26" Margin="81,262,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="120" VerticalScrollBarVisibility="Disabled"/>
                    <Label x:Name="label_ADNU_initial" Content="Initial:" HorizontalAlignment="Left" Margin="38,262,0,0" VerticalAlignment="Top"/>
                    <GroupBox x:Name="grpbox_ADNU_Vaults" Header="Pulse Vaults" HorizontalAlignment="Left" Height="80" Margin="24,314,0,0" VerticalAlignment="Top" Width="180">
                        <CheckBox x:Name="chkbox_ADNU_site1" Content="OMA Vault" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="80"/>
                    </GroupBox>
                    <CheckBox x:Name="chkbox_ADNU_site2" Content="DGN VAult" HorizontalAlignment="Left" Margin="40,364,0,0" VerticalAlignment="Top" Width="80"/>
                    <GroupBox x:Name="grpbox_ADNU_expiry" Header="Expiry" HorizontalAlignment="Left" Height="49" Margin="269,83,0,0" VerticalAlignment="Top" Width="228">
                        <CheckBox x:Name="chkbox_ADNU_temp" Content="Temp User" HorizontalAlignment="Right" Margin="0,5,132.6,0" VerticalAlignment="Top"/>
                    </GroupBox>
                    <DatePicker x:Name="datepkr_ADNU_expiry" HorizontalAlignment="Left" Margin="373,100,0,0" VerticalAlignment="Top" Width="110" IsEnabled="False"/>
                    <Label x:Name="label_ADNU_logonname" Content="Logon Name:" HorizontalAlignment="Left" Margin="219,143,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtbox_ADNU_logonname" HorizontalAlignment="Left" Height="26" Margin="299,143,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="200" VerticalScrollBarVisibility="Disabled"/>
                    <Label x:Name="label_ADNU_displayname" Content="Display Name:" HorizontalAlignment="Left" Margin="214,207,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtbox_ADNU_displayname" HorizontalAlignment="Left" Height="26" Margin="299,208,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="200" VerticalScrollBarVisibility="Disabled"/>
                    <TextBox x:Name="txtbox_ADNU_manager" HorizontalAlignment="Left" Height="26" Margin="299,263,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="200" VerticalScrollBarVisibility="Disabled"/>
                    <TextBox x:Name="txtbox_ADNU_title" HorizontalAlignment="Left" Height="26" Margin="299,321,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="200" VerticalScrollBarVisibility="Disabled"/>
                    <Label x:Name="label_ADNU_title" Content="Job Title:" HorizontalAlignment="Left" Margin="237,320,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="txtbox_ADNU_dept" HorizontalAlignment="Left" Height="26" Margin="299,372,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="200" VerticalScrollBarVisibility="Disabled"/>
                    <Label x:Name="label_ADNU_dept" Content="Department:" HorizontalAlignment="Left" Margin="218,372,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="label_ADNI_manager" Content="Manager Name:" HorizontalAlignment="Left" Margin="206,263,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.504,0.465"/>
                    <Button x:Name="btn_ADNU_resetform" Content="Reset" HorizontalAlignment="Left" Margin="14,529,0,0" VerticalAlignment="Top" Width="110" Height="26"/>
                    <Button x:Name="btn_ADNU_validate" Content="Validate" HorizontalAlignment="Left" Margin="138,529,0,0" VerticalAlignment="Top" Width="110" Height="26"/>
                    <Button x:Name="btn_ADNU_edit" Content="Edit" HorizontalAlignment="Left" Margin="263,529,0,0" VerticalAlignment="Top" Width="110" Height="26" IsEnabled="False"/>
                    <Button x:Name="btn_ADNU_Create" Content="Create user" HorizontalAlignment="Left" Margin="389,529,0,0" VerticalAlignment="Top" Width="110" Height="26" IsEnabled="False"/>
                    <TextBlock x:Name="txtblock_ADNU_errorexpiry" HorizontalAlignment="Left" Margin="488,101,0,0" TextWrapping="NoWrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Visible"/>
                    <TextBlock x:Name="txtblock_ADNU_errorsite" HorizontalAlignment="Left" Margin="206,96,0,0" TextWrapping="NoWrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Visible"/>
                    <TextBlock x:Name="txtblock_ADNU_errorfirst" HorizontalAlignment="Left" Margin="209,148,0,0" TextWrapping="NoWrap" Text="!" VerticalAlignment="Top" Foreground="Red" Width="5" FontWeight="Bold" Visibility="Visible"/>
                    <TextBlock x:Name="txtblock_ADNU_errorlast" HorizontalAlignment="Left" Margin="206,212,0,0" TextWrapping="NoWrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Visible"/>
                    <TextBlock x:Name="txtblock_ADNU_errorlogonunique" HorizontalAlignment="Left" Margin="300,169,0,0" TextWrapping="Wrap" Text="Logon name already exists in AD. Please enter a unique logon name" VerticalAlignment="Top" Width="198" Foreground="Red" Visibility="Visible"/>
                    <TextBlock x:Name="txtblock_ADNU_errormgrunique" HorizontalAlignment="Left" Margin="301,289,0,0" TextWrapping="Wrap" Text="Manager does not exist in AD. Please enter valid manager" VerticalAlignment="Top" Width="198" Foreground="Red" Visibility="Visible" Height="32"/>
                    <TextBlock x:Name="txtblock_ADNU_errorlogon" HorizontalAlignment="Left" Margin="504,148,0,0" TextWrapping="NoWrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Visible"/>
                    <TextBlock x:Name="txtblock_ADNU_errordisp" HorizontalAlignment="Left" Margin="504,215,0,0" TextWrapping="NoWrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Visible"/>
                    <TextBlock x:Name="txtblock_ADNU_errormgr" HorizontalAlignment="Left" Margin="504,268,0,0" TextWrapping="NoWrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Visible"/>
                    <TextBlock x:Name="txtblock_ADNU_errortitle" HorizontalAlignment="Left" Margin="504,325,0,0" TextWrapping="NoWrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Visible"/>
                    <TextBlock x:Name="txtblock_ADNU_errordept" HorizontalAlignment="Left" Margin="504,377,0,0" TextWrapping="NoWrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Visible"/>
                    <Label x:Name="label_ADNU_CSVPath" Content="CSV Path:" HorizontalAlignment="Left" Margin="22,459,0,0" VerticalAlignment="Top" Width="62" Height="26"/>
                    <TextBox x:Name="txtbox_ADNU_filebrowse" HorizontalAlignment="Left" Height="26" Margin="84,459,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" Width="289" IsEnabled="False" VerticalScrollBarVisibility="Disabled"/>
                    <Button x:Name="btn_ADNU_importCSV" Content="Browse" HorizontalAlignment="Left" Margin="389,459,0,0" VerticalAlignment="Top" Width="108" Height="26" IsEnabled="False"/>
                    <RadioButton x:Name="radio_ADNU_Single" Content="Single User Creation" HorizontalAlignment="Left" Margin="78,24,0,0" VerticalAlignment="Top" IsChecked="True"/>
                    <RadioButton x:Name="radio_ADNU_multi" Content="Multi User Creation" HorizontalAlignment="Left" Margin="301,24,0,0" VerticalAlignment="Top"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="tab_ExportUserInfo" Header="Export AD User Info">
                <Grid Background="#FFE5E5E5">
                    <GroupBox x:Name="grpbox_EADUI_Scope" Header="Scope" HorizontalAlignment="Left" Height="126" Margin="10,10,0,0" VerticalAlignment="Top" Width="487"/>
                    <RadioButton x:Name="radio_EADUI_User" Content="Single user" HorizontalAlignment="Left" Margin="29,36,0,0" VerticalAlignment="Top" IsChecked="True"/>
                    <TextBox x:Name="txtbox_EADUI_username" Height="26" Margin="275,31,0,0" TextWrapping="NoWrap" VerticalAlignment="Top" HorizontalAlignment="Left" Width="200" VerticalScrollBarVisibility="Disabled"/>
                    <TextBlock x:Name="txtblock_EADUI_errorusername" HorizontalAlignment="Left" Margin="480,36,0,0" TextWrapping="Wrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                    <TextBlock x:Name="txtblock_EADUI_erroruserunique" HorizontalAlignment="Left" Margin="279,56,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="200" Foreground="Red" Visibility="Hidden"><Run Text="User does not exist in AD. Please enter valid "/><Run Text="user"/><Run Text="name"/></TextBlock>
                    <RadioButton x:Name="radio_EADUI_OU" Content="Users in OU" HorizontalAlignment="Left" Margin="29,103,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="label_EADUI_username" Content="Username:" HorizontalAlignment="Left" Height="25" Margin="204,31,0,0" VerticalAlignment="Top" Width="66"/>
                    <Label x:Name="label_EADUI_username_Copy" Content="Site:" HorizontalAlignment="Left" Height="25" Margin="237,93,0,0" VerticalAlignment="Top" Width="33"/>
                    <CheckBox x:Name="chkbox_EADUI_AllAtts" Content="All Attributes" HorizontalAlignment="Left" Margin="17,156,0,0" VerticalAlignment="Top" IsChecked="True"/>
                    <CheckBox x:Name="chkbox_EADUI_common" Content="Common Attributes" HorizontalAlignment="Left" Margin="138,156,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_AdditionalAtts" Content="Additional Attributes" HorizontalAlignment="Left" Margin="279,156,0,0" VerticalAlignment="Top" IsEnabled="False"/>
                    <GroupBox x:Name="grpbox_EADUI_common" Header="Common Attributes" HorizontalAlignment="Left" Height="155" Margin="10,192,0,0" VerticalAlignment="Top" Width="229"/>
                    <GroupBox x:Name="grpbox_EADUI_additionalatts" Header="Additional Attributes" HorizontalAlignment="Left" Height="195" Margin="258,192,0,0" VerticalAlignment="Top" Width="238"/>
                    <ComboBox x:Name="combo_EADUI_site" HorizontalAlignment="Left" Margin="275,93,0,0" VerticalAlignment="Top" Width="120" Height="26" IsEnabled="False"/>
                    <TextBlock x:Name="txtblock_EADUI_errorsite" HorizontalAlignment="Left" Margin="400,98,0,0" TextWrapping="NoWrap" Text="!" VerticalAlignment="Top" Foreground="Red" FontWeight="Bold" Visibility="Hidden"/>
                    <CheckBox x:Name="chkbox_EADUI_userPrincipalName" Content="Logon Name" HorizontalAlignment="Left" Margin="17,214,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_samaccname" Content="Logon Pre 2000" HorizontalAlignment="Left" Margin="17,235,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_display" Content="Display Name" HorizontalAlignment="Left" Margin="17,256,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_givenName" Content="First Name" HorizontalAlignment="Left" Margin="17,277,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_initials" Content="Initials" HorizontalAlignment="Left" Margin="17,298,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_sn" Content="Last Name" HorizontalAlignment="Left" Margin="17,319,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_email" Content="Email" HorizontalAlignment="Left" Margin="138,214,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_title" Content="Title" HorizontalAlignment="Left" Margin="138,235,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_dept" Content="Department" HorizontalAlignment="Left" Margin="138,256,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_desc" Content="Description" HorizontalAlignment="Left" Margin="138,277,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_mgr" Content="Manager" HorizontalAlignment="Left" Margin="138,298,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_alias" Content="Mail alias" HorizontalAlignment="Left" Margin="279,214,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_expiry" Content="Account Expires" HorizontalAlignment="Left" Margin="279,235,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_cn" Content="Canonical Name" HorizontalAlignment="Left" Margin="279,256,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_loginscpt" Content="Login Script" HorizontalAlignment="Left" Margin="279,277,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_hmefolder" Content="Home Folder" HorizontalAlignment="Left" Margin="279,298,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_hmedrv" Content="Home Drive" HorizontalAlignment="Left" Margin="279,319,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_logonto" Content="Log on to" HorizontalAlignment="Left" Margin="279,340,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_company" Content="Company" HorizontalAlignment="Left" Margin="389,214,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_Country" Content="Country" HorizontalAlignment="Left" Margin="389,235,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_office" Content="Office" HorizontalAlignment="Left" Margin="389,256,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_telephone" Content="Telephone" HorizontalAlignment="Left" Margin="389,277,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_street" Content="Street" HorizontalAlignment="Left" Margin="389,298,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_POBox" Content="Post Code" HorizontalAlignment="Left" Margin="389,319,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_City" Content="City" HorizontalAlignment="Left" Margin="389,340,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <CheckBox x:Name="chkbox_EADUI_State" Content="State/Province" HorizontalAlignment="Left" Margin="389,361,0,0" VerticalAlignment="Top" IsChecked="True" IsEnabled="False"/>
                    <Button x:Name="btn_EADUI_export" Content="Export to CSV" HorizontalAlignment="Left" Margin="374,406,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                </Grid>
            </TabItem>
        </TabControl>
        <TextBlock x:Name="txtblock_main_cred" HorizontalAlignment="Left" Margin="10,640,0,0" TextWrapping="NoWrap" Text="Credentials: " Width="232" Height="17" VerticalAlignment="Top"/>
        <TextBlock x:Name="txtblock_main_vsn" HorizontalAlignment="Left" Margin="684,640,0,0" TextWrapping="NoWrap" Text="Version: " Width="100" Height="17" VerticalAlignment="Top"/>
        <ProgressBar x:Name="progressbar_main" HorizontalAlignment="Left" Height="17" Margin="247,640,0,0" VerticalAlignment="Top" Width="175" Visibility="Hidden"/>
        <Menu x:Name="menu_Main" HorizontalAlignment="Left" Height="29" Margin="10,0,0,0" VerticalAlignment="Top" Width="517" IsEnabled="False" Visibility="Hidden">
            <MenuItem x:Name="menu_main_file" Header="File" Height="217" Width="117"/>
            <MenuItem x:Name="menu_main_help" Header="Help" Height="100" Width="100"/>
        </Menu>
        <TextBox x:Name="txtbox_main_log" HorizontalAlignment="Left" Height="584" Margin="554,56,0,0" TextWrapping="Wrap" Text="Logs" VerticalAlignment="Top" Width="230" VerticalScrollBarVisibility="Auto"/>
        <TextBlock x:Name="txtblock_main_logtitle" HorizontalAlignment="Left" Margin="554,34,0,0" TextWrapping="Wrap" Text="Log Title" VerticalAlignment="Top" Height="25" Width="230"/>
    </Grid>
</Window>

"@       
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window' -replace
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
#===========================================================================
# Load XAML Objects In PowerShell
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

#=======================#
#    Some Functions     #
#=======================#
function Reset-ADNUErrors{
    $WPFtxtblock_ADNU_errorexpiry.Visibility = "hidden"
    $WPFtxtblock_ADNU_errorsite.Visibility = "hidden"
    $WPFtxtblock_ADNU_errorfirst.Visibility = "hidden"
    $WPFtxtblock_ADNU_errorlast.Visibility = "hidden"
    $WPFtxtblock_ADNU_errorlogonunique.Visibility = "hidden"
    $WPFtxtblock_ADNU_errormgrunique.Visibility = "hidden"
    $WPFtxtblock_ADNU_errorlogon.Visibility = "hidden"
    $WPFtxtblock_ADNU_errordisp.Visibility = "hidden"
    $WPFtxtblock_ADNU_errormgr.Visibility = "hidden"
    $WPFtxtblock_ADNU_errordept.Visibility = "hidden"
    $WPFtxtblock_ADNU_errortitle.Visibility = "hidden"
}

function Reset-UserInputVars{
    $Global:Grptype = $null
    $Global:Grpname = $null
    $Global:kit = $null
    $Global:Groups = $null
    $Global:homeDirectory = $null
    $Global:homeDrive = $null
    $Global:company = $null
    $Global:L_city = $null
    $Global:physicalDeliveryOfficeName = $null
    $Global:postalCode = $null
    $Global:scriptPath = $null
    $Global:st = $null
    $Global:streetAddress = $null
    $Global:userPrincipalName = $null
    $Global:SiteOU = $null
    $Global:errorcount = $null
}

function Reset-ADNUFormValues{
    $WPFcombo_ADNU_Site.SelectedValue = $null
    $WPFtxtbox_ADNU_firstname.Text = $null
    $WPFtxtbox_ADNU_lastname.Text = $null
    $WPFtxtbox_ADNU_initial.Text = $null
    $WPFtxtbox_ADNU_logonname.Text = $null
    $WPFtxtbox_ADNU_displayname.Text = $null
    $WPFtxtbox_ADNU_manager.Text = $null
    $WPFtxtbox_ADNU_title.Text = $null
    $WPFtxtbox_ADNU_dept.Text = $null
    $WPFtxtbox_ADNU_filebrowse.Text = $null
}

function Reset-ADNUForm{
    $WPFcombo_ADNU_Site.IsEnabled = $true
    $WPFtxtbox_ADNU_firstname.IsEnabled = $true
    $WPFtxtbox_ADNU_lastname.IsEnabled = $true
    $WPFtxtbox_ADNU_initial.IsEnabled = $true
    $WPFchkbox_ADNU_Site1.IsEnabled = $true
    $WPFchkbox_ADNU_Site2.IsEnabled = $true
    $WPFchkbox_ADNU_temp.IsEnabled = $true
    $WPFtxtbox_ADNU_logonname.IsEnabled = $true
    $WPFtxtbox_ADNU_displayname.IsEnabled = $true
    $WPFtxtbox_ADNU_manager.IsEnabled = $true
    $WPFtxtbox_ADNU_title.IsEnabled = $true
    $WPFtxtbox_ADNU_dept.IsEnabled = $true
    $WPFbtn_ADNU_validate.IsEnabled = $true
    $WPFbtn_ADNU_edit.IsEnabled = $false
    $WPFbtn_ADNU_Create.IsEnabled = $false
    $WPFchkbox_ADNU_Site1.IsChecked = $false
    $WPFchkbox_ADNU_Site2.IsChecked = $false
    $WPFchkbox_ADNU_temp.IsChecked = $false
    $WPFradio_ADNU_Single.IsChecked = $true
    $WPFradio_ADNU_multi.IsChecked = $false
    $WPFtxtbox_ADNU_filebrowse.IsEnabled = $false
    $WPFbtn_ADNU_importCSV.IsEnabled = $false
}

function Display-Error{
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageIcon = [System.Windows.MessageBoxImage]::Error
    [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
    $Global:LogStatus = "ERROR"
    $Global:LogMSG = $Messageboxbody
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
}

function Display-Warning{
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageIcon = [System.Windows.MessageBoxImage]::Warning
    [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
    $Global:LogStatus = "WARNING"
    $Global:LogMSG = $Messageboxbody
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
}

function Display-Info{
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageIcon = [System.Windows.MessageBoxImage]::Information
    [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $Messageboxbody
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
}

function Display-Success{
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageIcon = [System.Windows.MessageBoxImage]::Information
    [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
    $Global:LogStatus = "SUCCESS"
    $Global:LogMSG = $Messageboxbody
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
}

function Load-CSVData{

    $Global:AllDefaultAtts = Import-Csv "C:\scripts\IT_Admin_Tool\DefaultAttsCSV.csv"
    $Global:LauncherApps = Import-Csv "C:\scripts\IT_Admin_Tool\LauncherApplicationsCSV.csv"

}

function Show-CredAndVsn{
    $WPFtxtblock_main_cred.Text =  "Credentials: " + $env:USERNAME
    $WPFtxtblock_main_vsn.Text = $version
}

function Load-LauncherButtons{
    # Load App1 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "A1"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_A1.Visibility = "Hidden"
        $WPFimg_A1.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_A1.Content = $Appdetails.AppName
                $WPFimg_A1.Source = $Appdetails.IconPath
                $Global:cmd_A1 = $Appdetails.AppPathCommand
                $Global:prerun_A1 = $Appdetails.PreRunCmd
                }

    # Load App2 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "A2"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_A2.Visibility = "Hidden"
        $WPFimg_A2.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_A2.Content = $Appdetails.AppName
                $WPFimg_A2.Source = $Appdetails.IconPath
                $Global:cmd_A2 = $Appdetails.AppPathCommand
                $Global:prerun_A2 = $Appdetails.PreRunCmd
                }

    # Load App3 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "A3"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_A3.Visibility = "Hidden"
        $WPFimg_A3.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_A3.Content = $Appdetails.AppName
                $WPFimg_A3.Source = $Appdetails.IconPath
                $Global:cmd_A3 = $Appdetails.AppPathCommand
                $Global:prerun_A3 = $Appdetails.PreRunCmd
                }

    # Load App4 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "A4"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_A4.Visibility = "Hidden"
        $WPFimg_A4.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_A4.Content = $Appdetails.AppName
                $WPFimg_A4.Source = $Appdetails.IconPath
                $Global:cmd_A4 = $Appdetails.AppPathCommand
                $Global:prerun_A4 = $Appdetails.PreRunCmd
                }

    # Load App5 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "A5"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_A5.Visibility = "Hidden"
        $WPFimg_A5.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_A5.Content = $Appdetails.AppName
                $WPFimg_A5.Source = $Appdetails.IconPath
                $Global:cmd_A5 = $Appdetails.AppPathCommand
                $Global:prerun_A5 = $Appdetails.PreRunCmd
                }

    # Load App5 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "A6"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_A6.Visibility = "Hidden"
        $WPFimg_A6.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_A6.Content = $Appdetails.AppName
                $WPFimg_A6.Source = $Appdetails.IconPath
                $Global:cmd_A6 = $Appdetails.AppPathCommand
                $Global:prerun_A6 = $Appdetails.PreRunCmd
                }

    # Load App7 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "B1"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_B1.Visibility = "Hidden"
        $WPFimg_B1.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_B1.Content = $Appdetails.AppName
                $WPFimg_B1.Source = $Appdetails.IconPath
                $Global:cmd_B1 = $Appdetails.AppPathCommand
                $Global:prerun_B1 = $Appdetails.PreRunCmd
                }

    # Load App8 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "B2"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_B2.Visibility = "Hidden"
        $WPFimg_B2.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_B2.Content = $Appdetails.AppName
                $WPFimg_B2.Source = $Appdetails.IconPath
                $Global:cmd_B2 = $Appdetails.AppPathCommand
                $Global:prerun_B2 = $Appdetails.PreRunCmd
                }

    # Load App9 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "B3"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_B3.Visibility = "Hidden"
        $WPFimg_B3.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_B3.Content = $Appdetails.AppName
                $WPFimg_B3.Source = $Appdetails.IconPath
                $Global:cmd_B3 = $Appdetails.AppPathCommand
                $Global:prerun_B3 = $Appdetails.PreRunCmd
                }

    # Load App10 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "B4"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_B4.Visibility = "Hidden"
        $WPFimg_B4.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_B4.Content = $Appdetails.AppName
                $WPFimg_B4.Source = $Appdetails.IconPath
                $Global:cmd_B4 = $Appdetails.AppPathCommand
                $Global:prerun_B4 = $Appdetails.PreRunCmd
                }

    # Load App11 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "B5"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_B5.Visibility = "Hidden"
        $WPFimg_B5.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_B5.Content = $Appdetails.AppName
                $WPFimg_B5.Source = $Appdetails.IconPath
                $Global:cmd_B5 = $Appdetails.AppPathCommand
                $Global:prerun_B5 = $Appdetails.PreRunCmd
                }

    # Load App12 details and hide button if none
    $Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "B6"}
    If ($Appdetails.AppName -eq ""){
        $WPFbtn_B6.Visibility = "Hidden"
        $WPFimg_B6.Visibility = "Hidden"
        } ELSE {
                $WPFbtn_B6.Content = $Appdetails.AppName
                $WPFimg_B6.Source = $Appdetails.IconPath
                $Global:cmd_B6 = $Appdetails.AppPathCommand
                $Global:prerun_B6 = $Appdetails.PreRunCmd
                }
}

function Load-ComboBoxes{

    $Global:ComboSiteCodes = @() 
    $Global:ComboSiteCodes += ""
    $Global:ComboSiteCodes += $Global:AllDefaultAtts.SiteCode

    $Global:GroupTypesCombo = @()
    $Global:GroupTypesCombo += "" 
    $Global:GroupTypesCombo += "-Data_"
    $Global:GroupTypesCombo += "-Data_Dept_"
    $Global:GroupTypesCombo += "-DST_"
    $Global:GroupTypesCombo += "-MBX-"
     
    ## AD Tools Tab
    $Global:ComboSiteCodes | ForEach-object {$WPFcombobox_ADCmds_site.AddChild($_)}
    "" , "Workstation" , "Notebook" | ForEach-object {$WPFcombobox_kittype.AddChild($_)}
    $Global:GroupTypesCombo| ForEach-object {$WPFcomboBox_Grptype.AddChild($_)}

    ## Create new AD user tab
    $Global:ComboSiteCodes | ForEach-object {$WPFcombo_ADNU_site.AddChild($_)}

    ## Export AD User Info tab
    $Global:ComboSiteCodes | ForEach-object {$WPFcombo_EADUI_site.AddChild($_)}
}

function Load-Logs{
    $Date = Get-Date -UFormat "%d-%m-%Y"
    $DateTime = Get-Date
    $Time = Get-Date -UFormat "%H:%M:%S"

    $Global:logfilename = $env:USERNAME + "_" + $Date + ".log"
    $Global:logpath = $installpath + "\logs\"
    $Global:logfullpath = $logpath + $logfilename
    $logloaderror = $null

    Do {

        If (Test-Path $logpath)
            {
            If (Test-Path $logfullpath)
                    {
                    $WPFtxtbox_main_log.Text = (Get-Content $logfullpath) -join "`n"
                    $WPFtxtblock_main_logtitle.Text = "Logs: " + $Date
                    $logloaderror = $null
                    } Else {
                        New-Item $logfullpath -ItemType file
                        $Global:LogStatus = "INFO"
                        $Global:LogModule = "MAIN"
                        $Global:LogMSG = "New Log file generated"
                        Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
                        $logloaderror++
                        }
            } Else {
                New-Item $logpath -ItemType directory
                $logloaderror++
                }


        If ($logloaderror -eq 6)
            {
            $WPFtxtbox_main_log.Text = "Error loading logfile. Please ensure you have the necessary write permissions on the location of the logs: `n" + $logpath
            }
        }
    Until (($logloaderror -eq $null) -or ($logloaderror -eq 6))

}

function Update-LogFile ($Global:LogStatus,$Global:LogModule,$Global:LogMSG){
    $Date = Get-Date -UFormat "%d-%m-%Y"
    $DateTime = Get-Date
    $Time = Get-Date -UFormat "%H:%M:%S"

    $Global:logfilename = $env:USERNAME + "_" + $Date + ".log"
    $Global:logpath = $installpath + "\logs\"
    $Global:logfullpath = $logpath + $logfilename
    $LogCurrentContent = Get-Content $logfullpath
    $NewLogContent = $Global:LogStatus + " | " + $Time + " | " + $LogModule + " | " + $Global:LogMSG
    
    Set-Content $logfullpath -value $null
    Add-Content $logfullpath $NewLogContent
    Add-Content $logfullpath $LogCurrentContent

    Load-Logs
}

function Get-DefaultAtts{
    $Site = $WPFcombo_ADNU_Site.SelectedValue
    $Global:SiteDefaultAtts = $Global:AllDefaultAtts | Where-Object{$_.SiteCode -eq $Site}

    $Global:Groups = $Global:SiteDefaultAtts.DefaultGroups
    $Global:homeDirectory = $Global:SiteDefaultAtts.HomeDirectory + $logonname + "\"
    $Global:homeDrive = $Global:SiteDefaultAtts.HomeDrive
    $Global:company = $Global:SiteDefaultAtts.CompanyAtt
    $Global:L_city = $Global:SiteDefaultAtts.SiteCity
    $Global:physicalDeliveryOfficeName = $Global:SiteDefaultAtts.SiteCity
    $Global:postalCode = $Global:SiteDefaultAtts.PostCode
    $Global:scriptPath = $Global:SiteDefaultAtts.ScriptPath
    $Global:st = $Global:SiteDefaultAtts.County
    $Global:streetAddress = $Global:SiteDefaultAtts.streetAddress
    $Global:userPrincipalName = "$logonname" + "@" + $Global:SiteDefaultAtts.CompanyName + ".com"
    $Global:SiteOU = "OU=Users," + $Global:SiteDefaultAtts.OU
}

function Uncheck-CommonAtts{
    # Uncheck the common atts
    $WPFchkbox_EADUI_userPrincipalName.IsChecked = $false
    $WPFchkbox_EADUI_samaccname.IsChecked = $false
    $WPFchkbox_EADUI_display.IsChecked = $false
    $WPFchkbox_EADUI_givenName.IsChecked = $false
    $WPFchkbox_EADUI_initials.IsChecked = $false
    $WPFchkbox_EADUI_sn.IsChecked = $false
    $WPFchkbox_EADUI_email.IsChecked = $false
    $WPFchkbox_EADUI_title.IsChecked = $false
    $WPFchkbox_EADUI_dept.IsChecked = $false
    $WPFchkbox_EADUI_desc.IsChecked = $false
    $WPFchkbox_EADUI_mgr.IsChecked = $false
}

function Check-CommonAtts{
    # Check the common atts
    $WPFchkbox_EADUI_userPrincipalName.IsChecked = $true
    $WPFchkbox_EADUI_samaccname.IsChecked = $true
    $WPFchkbox_EADUI_display.IsChecked = $true
    $WPFchkbox_EADUI_givenName.IsChecked = $true
    $WPFchkbox_EADUI_initials.IsChecked = $true
    $WPFchkbox_EADUI_sn.IsChecked = $true
    $WPFchkbox_EADUI_email.IsChecked = $true
    $WPFchkbox_EADUI_title.IsChecked = $true
    $WPFchkbox_EADUI_dept.IsChecked = $true
    $WPFchkbox_EADUI_desc.IsChecked = $true
    $WPFchkbox_EADUI_mgr.IsChecked = $true
}

function Uncheck-AdditionalAtts{
    # Uncheck the additional atts
    $WPFchkbox_EADUI_alias.IsChecked = $false
    $WPFchkbox_EADUI_City.IsChecked = $false
    $WPFchkbox_EADUI_company.IsChecked = $false
    $WPFchkbox_EADUI_Country.IsChecked = $false
    $WPFchkbox_EADUI_expiry.IsChecked = $false
    $WPFchkbox_EADUI_hmedrv.IsChecked = $false
    $WPFchkbox_EADUI_hmefolder.IsChecked = $false
    $WPFchkbox_EADUI_loginscpt.IsChecked = $false
    $WPFchkbox_EADUI_logonto.IsChecked = $false
    $WPFchkbox_EADUI_office.IsChecked = $false
    $WPFchkbox_EADUI_POBox.IsChecked = $false
    $WPFchkbox_EADUI_cn.IsChecked = $false
    $WPFchkbox_EADUI_State.IsChecked = $false
    $WPFchkbox_EADUI_street.IsChecked = $false
    $WPFchkbox_EADUI_telephone.IsChecked = $false
}

function Check-AdditionalAtts{
    # Check the additional atts
    $WPFchkbox_EADUI_alias.IsChecked = $true
    $WPFchkbox_EADUI_City.IsChecked = $true
    $WPFchkbox_EADUI_company.IsChecked = $true
    $WPFchkbox_EADUI_Country.IsChecked = $true
    $WPFchkbox_EADUI_expiry.IsChecked = $true
    $WPFchkbox_EADUI_hmedrv.IsChecked = $true
    $WPFchkbox_EADUI_hmefolder.IsChecked = $true
    $WPFchkbox_EADUI_loginscpt.IsChecked = $true
    $WPFchkbox_EADUI_logonto.IsChecked = $true
    $WPFchkbox_EADUI_office.IsChecked = $true
    $WPFchkbox_EADUI_POBox.IsChecked = $true
    $WPFchkbox_EADUI_cn.IsChecked = $true
    $WPFchkbox_EADUI_State.IsChecked = $true
    $WPFchkbox_EADUI_street.IsChecked = $true
    $WPFchkbox_EADUI_telephone.IsChecked = $true
}

function Disable-CommonAtts{
    # Disable the common atts
    $WPFchkbox_EADUI_userPrincipalName.IsEnabled = $false
    $WPFchkbox_EADUI_samaccname.IsEnabled = $false
    $WPFchkbox_EADUI_display.IsEnabled = $false
    $WPFchkbox_EADUI_givenName.IsEnabled = $false
    $WPFchkbox_EADUI_initials.IsEnabled = $false
    $WPFchkbox_EADUI_sn.IsEnabled = $false
    $WPFchkbox_EADUI_email.IsEnabled = $false
    $WPFchkbox_EADUI_title.IsEnabled = $false
    $WPFchkbox_EADUI_dept.IsEnabled = $false
    $WPFchkbox_EADUI_desc.IsEnabled = $false
    $WPFchkbox_EADUI_mgr.IsEnabled = $false
}

function Disable-AdditionalAtts{
    # Disable the additional atts
    $WPFchkbox_EADUI_alias.IsEnabled = $false
    $WPFchkbox_EADUI_City.IsEnabled = $false
    $WPFchkbox_EADUI_company.IsEnabled = $false
    $WPFchkbox_EADUI_Country.IsEnabled = $false
    $WPFchkbox_EADUI_expiry.IsEnabled = $false
    $WPFchkbox_EADUI_hmedrv.IsEnabled = $false
    $WPFchkbox_EADUI_hmefolder.IsEnabled = $false
    $WPFchkbox_EADUI_loginscpt.IsEnabled = $false
    $WPFchkbox_EADUI_logonto.IsEnabled = $false
    $WPFchkbox_EADUI_office.IsEnabled = $false
    $WPFchkbox_EADUI_POBox.IsEnabled = $false
    $WPFchkbox_EADUI_cn.IsEnabled = $false
    $WPFchkbox_EADUI_State.IsEnabled = $false
    $WPFchkbox_EADUI_street.IsEnabled = $false
    $WPFchkbox_EADUI_telephone.IsEnabled = $false
}

function Enable-CommonAtts{
    # Enable the common atts
    $WPFchkbox_EADUI_userPrincipalName.IsEnabled = $true
    $WPFchkbox_EADUI_samaccname.IsEnabled = $true
    $WPFchkbox_EADUI_display.IsEnabled = $true
    $WPFchkbox_EADUI_givenName.IsEnabled = $true
    $WPFchkbox_EADUI_initials.IsEnabled = $true
    $WPFchkbox_EADUI_sn.IsEnabled = $true
    $WPFchkbox_EADUI_email.IsEnabled = $true
    $WPFchkbox_EADUI_title.IsEnabled = $true
    $WPFchkbox_EADUI_dept.IsEnabled = $true
    $WPFchkbox_EADUI_desc.IsEnabled = $true
    $WPFchkbox_EADUI_mgr.IsEnabled = $true
}

function Enable-AdditionalAtts{
    # Enable the additional atts
    $WPFchkbox_EADUI_alias.IsEnabled = $true
    $WPFchkbox_EADUI_City.IsEnabled = $true
    $WPFchkbox_EADUI_company.IsEnabled = $true
    $WPFchkbox_EADUI_Country.IsEnabled = $true
    $WPFchkbox_EADUI_expiry.IsEnabled = $true
    $WPFchkbox_EADUI_hmedrv.IsEnabled = $true
    $WPFchkbox_EADUI_hmefolder.IsEnabled = $true
    $WPFchkbox_EADUI_loginscpt.IsEnabled = $true
    $WPFchkbox_EADUI_logonto.IsEnabled = $true
    $WPFchkbox_EADUI_office.IsEnabled = $true
    $WPFchkbox_EADUI_POBox.IsEnabled = $true
    $WPFchkbox_EADUI_cn.IsEnabled = $true
    $WPFchkbox_EADUI_State.IsEnabled = $true
    $WPFchkbox_EADUI_street.IsEnabled = $true
    $WPFchkbox_EADUI_telephone.IsEnabled = $true
}

function Set-GroupNTFSAccess ($Account,$Path,$AccessRights){
    $Global:LogModule = "GroupNTFSAccess"
    # If Group and Dir fields are not blank
    If (($Account -ne " ") -and ($account -ne "") -and ($account -ne $null) -and ($Path -ne " ") -and ($Path -ne "") -and ($Path -ne $null))
        {
        # If group exists
        If (dsquery group -name $account)
            {
            # If Directory exists
            If (Test-Path $Path){
                # Set permissions on folder
                Add-NTFSAccess -Path $Path -Account $Account -AccessRights $AccessRights
                # Confirm the permissions are set
                $permission = (Get-Acl $Path).Access | ?{$_.IdentityReference -match $Account} | Select IdentityReference,FileSystemRights
                If ($permission){
                    $Messageboxbody = "Folder created and $AccessRights permissions set."
                    Display-Success
                    }
                    Else {
                        $Messageboxbody = "Operation failed. $AccessRights permissions could not be set."
                        Display-Error
                        }
                } Else {

                    $MessageboxTitle = "QUESTION"
                    $Messageboxbody = "Folder path does not exist. Do you want to create folder? `n$path"
                    $Global:LogMSG = "Folder does not exist: $Path"
                    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG

                    $msgBoxInput =  [System.Windows.MessageBox]::Show($Messageboxbody,$MessageboxTitle,'YesNo','Question')

                    switch  ($msgBoxInput) {
                        'Yes' {
                            $Global:LogMSG = "Attempting to create folder..."
                            Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
                            # Create the new folder
                            New-Item $Path -type directory
                            If (Test-Path $Path){
                                $Global:LogMSG = "Folder created."
                                Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
                                # Set the folder permissions
                                $Global:LogMSG = "Setting $AccessRights access rights on folder."
                                Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
                                Add-NTFSAccess -Path $Path -Account $Account -AccessRights $AccessRights

                                # Confirm the permissions are set
                                $permission = (Get-Acl $Path).Access | ?{$_.IdentityReference -match $Account} | Select IdentityReference,FileSystemRights
                                If ($permission){
                                    $Messageboxbody = "Folder created and $AccessRights permissions set."
                                    Display-Success 
                                    }
                                    Else {
                                        $Messageboxbody = "Operation failed. Folder was created but $AccessRights permissions could not be set."
                                        Display-Error
                                        }
                                } Else {
                                $Messageboxbody = "Operation failed. Aborting!!"
                                Display-Error
                                }
                            }

                        'No' {
                            $Messageboxbody = "Operation Cancelled"
                            Display-Info
                            }
                        }
                    }
            } Else {
                $Messageboxbody = "Group name $Account cannot be validated in AD. Plase enter a valid group name to continue"
                Display-Error
                }
        }
    Get-module | Remove-Module
}

function Reset-ADCmdsForm{
    $WPFcombobox_kittype.SelectedValue = $null
    $WPFcombobox_GrpType.SelectedValue = $null
    $WPFcombobox_kittype.IsEnabled = $false
    $WPFcombobox_GrpType.IsEnabled = $false
    $WPFcombobox_ADCmds_site.SelectedValue = $null
    $WPFtxtbox_ADCmds_username.Text = $null
    $WPFtxtbox_ADCmds_computername.Text = $null
    $WPFtxtbox_ADCmds_group.Text = $null
    $WPFtxtbox_ADCmds_dirpath.Text = $null
}

function Create-SingleADUser{
    $Global:LogModule = "SingleADNU"
    Reset-UserInputVars

    # Collect user input into variables and clean the data
    $Site = $WPFcombo_ADNU_Site.SelectedValue

    $firstname = $WPFtxtbox_ADNU_firstname.Text
    If (($firstname -ne "") -and ($firstname -ne $null))
        {
        $firstname = $firstname.Replace(' ','')
        $firstname = $firstname.Replace("'",'')
        }

    $lastname = $WPFtxtbox_ADNU_lastname.Text
    If (($lastname -ne "") -and ($lastname -ne $null))
        {
        $lastname = $lastname.Replace(' ','')
        $lastname = $lastname.Replace("'",'')
        }

    $initial = $WPFtxtbox_ADNU_initial.Text
    If (($initial -ne "") -and ($initial -ne $null))
        {
        $initial = $initial.Replace(' ','')
        }

    $logonname = $WPFtxtbox_ADNU_logonname.Text
    If (($logonname -ne "") -and ($logonname -ne $null))
        {
        $logonname = $logonname.Replace(' ','')
        }

    $displayname = $WPFtxtbox_ADNU_displayname.Text
    If (($displayname -ne "") -and ($displayname -ne $null))
        {
        $displayname = $displayname.Trim()
        }

    $manager = $WPFtxtbox_ADNU_manager.Text
    If (($manager -ne "") -and ($manager -ne $null))
        {
        $manager = $manager.Trim()
        $manager = $manager.Replace(' ','.')
        }

    $job = $WPFtxtbox_ADNU_title.Text
    If(($job -ne "") -and ($job -ne $null))
        {
        $job = $job.Trim()
        }

    $dept = $WPFtxtbox_ADNU_dept.Text
    If (($dept -ne "") -and ($dept -ne $null))
        {
        $dept = $dept.Trim()
        }
    If ($WPFchkbox_ADNU_Site1.IsChecked -eq $true)
        {
        $omavault = $true
        }

    If ($WPFchkbox_ADNU_Site2.IsChecked -eq $true)
        {
        $dgnvault = $true
        }

    If ($WPFchkbox_ADNU_temp.IsChecked -eq $true)
        {
        $newDateTime = [DateTime]::Parse($WPFdatepkr_ADNU_expiry.Text) 
        $ExpireDate = $newDateTime.AddDays(1)
        }




### Set default attributes and pulse vaults if applicable ###

    Get-DefaultAtts

    
    ## Add the Pulse vault security groups if Pulse vaults checked. (OMA vault is added by default to OMA users so no need to add twice)
    IF (($omavault -eq $true) -and ($Site -ne "OMA"))
        {
        $Groups = "$Groups" + "," + "OMA-Data_PDM_A"
        }                    
    
    IF ($dgnvault -eq $true)
        {   
        $Groups = "$Groups" + "," + "DGNApp_Pulse"
        }   


### This section creates the user account per the details entered by the technician. ###    

    # Create User account
    $Global:LogMSG = "Attempting to create AD account for $logonname"
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
    New-ADUser -SamAccountName $logonname -AccountExpirationDate $ExpireDate -userPrincipalName $userPrincipalName -GivenName $firstname -Surname $lastname -initials $initial -displayName $displayname -department $dept -Name $displayname -Path $SiteOU -Description $job -Title $job -enable $True -AccountPassword (ConvertTo-SecureString -AsPlainText 'TempPwd2017' -Force) -homeDirectory $homeDirectory -homeDrive $homeDrive -company $company -city $L_city -office $physicalDeliveryOfficeName -postalCode $postalCode -scriptPath $scriptPath -state $st -streetAddress $streetAddress -manager $manager
    Set-ADUser –Identity $logonname –ChangePasswordAtLogon $true

    # Check that user account was created successfully before adding to groups
    IF (dsquery user -samid $logonname)
        {
        $Groups = $Groups.split(",")
        foreach ($Group in $Groups)
            {
            Add-ADGroupMember -Identity $Group $logonname
            }  
        
        # Display success message
        $Messageboxbody = "AD account for $logonname has been created successfully!"
        Display-Success

        # Reset form for next use
        Reset-ADNUErrors
        Reset-ADNUFormValues
        Reset-ADNUForm         

        } ELSE 
              {
               # An error has occured.
               $Messageboxbody = "Failed to create user account $logonname. Job aborted!"
               Display-Error 
               }
    Get-module | Remove-Module
}

function Create-MultiADUser{
    Reset-UserInputVars
    ## Import CSV
    $multiuserCSV = Import-Csv $WPFtxtbox_ADNU_filebrowse.text

    foreach ($newUserdata in $multiuserCSV){

        ## Validate the CSV data and export errors to a log for analysis.
        Reset-ADNUErrors
        $Global:LogModule = "MultiADNUValidate"
        $LogMsg = @()
        $LogErrors = @()
    
        #Collect user input into variables and clean the data
        $Site = $newUserdata.Site
        If (($site -ne "") -and ($site -ne $null))
            {
            $site = $site.Trim()
            }

        $firstname = $newUserdata.FirstName
        If (($firstname -ne "") -and ($firstname -ne $null))
            {
            $firstname = $firstname.Replace(' ','')
            $firstname = $firstname.Replace("'",'')
            }

        $lastname = $newUserdata.LastName
        If (($lastname -ne "") -and ($lastname -ne $null))
            {
            $lastname = $lastname.Replace(' ','')
            $lastname = $lastname.Replace("'",'')
            }

        $initial = $newUserdata.Initial
        If (($initial -ne "") -and ($initial -ne $null))
            {
            $initial = $initial.Replace(' ','')
            }

        $logonname = $newUserdata.LogonName
        If (($logonname -ne "") -and ($logonname -ne $null))
            {
            $logonname = $logonname.Replace(' ','')
            }

        $displayname = $newUserdata.DisplayName
        If (($displayname -ne "") -and ($displayname -ne $null))
            {
            $displayname = $displayname.Trim()
            }

        $manager = $newUserdata.Manager
        If (($manager -ne "") -and ($manager -ne $null))
            {
            $manager = $manager.Trim()
            $manager = $manager.Replace(' ','.')
            }

        $job = $newUserdata.JobTitle
        If(($job -ne "") -and ($job -ne $null))
            {
            $job = $job.Trim()
            }

        $dept = $newUserdata.Department
        If (($dept -ne "") -and ($dept -ne $null))
            {
            $dept = $dept.Trim()
            }
        
        $Site1App = $newUserdata.OMA_Vault
        If (($Site1App -eq "Y") -or ($Site1App -eq "Yes"))
            {
            $omavault = $true
            }

        $Site2App = $newUserdata.DGN_Vault
        If (($Site2App -eq "Y") -or ($Site2App -eq "Yes"))
            {
            $dgnvault = $true
            }

        $tempuser = $newUserdata.Temp_User
        If (($tempuser -eq "Y") -or ($tempuser -eq "Yes"))
            {
            $tempuser = $true
            }

        # If Temp User is checked but no expiry date entered, error.
        If (($tempuser -eq $true) -and ($newUserdata.ExpiryDate -eq $null))
            {
            $errorcount++
            $LogErrors += "Temp User is checked but no expiry date entered; "
            } ELSE {
                If (($tempuser -eq $true) -and ($newUserdata.ExpiryDate -ne $null))
                    {
                    $newDateTime = [DateTime]::Parse($WPFdatepkr_ADNU_expiry.Text) 
                    $ExpireDate = $newDateTime.AddDays(1)
                    }  
                }    

        # If the site has not been provided.
        If (($Site -eq "") -or ($site -eq $null))
            {
            $errorcount++
            $LogErrors += "Site code not selected; "
            }

        # If the firstname is left blank, error
        If (($firstname -eq "") -or ($firstname -eq $null))
            {
            $errorcount++
            $LogErrors += "First name is blank; "
            }

        # If the lastname is left blank, error
        If (($lastname -eq "") -or ($lastname -eq $null))
            {
            $errorcount++
            $LogErrors += "Last name is blank; "
            }
        
        # If logon name is blank, error
        If (($logonname -eq "") -or ($logonname -eq $null))
            {
            $errorcount++
            $LogErrors += "Logon name is blank; "
            } ELSE {
                    # Checks that the logon name is unique
                    If (dsquery user -samid $logonname)
                        {
                        $errorcount++
                        $LogErrors += "$logonname is not unique; "
                        }
                    }

        # If display name is blank, error
        If (($displayname -eq "") -or ($displayname -eq $null))
            {
            $errorcount++
            $LogErrors += "Display name is blank; "
            }

        # If manager name is blank, error
        If (($manager -eq "") -or ($manager -eq $null))
            {
            $errorcount++
            $LogErrors += "Manager name is blank; "
            } ELSE {
                    # Check the inputted manager exists in AD
                    If (!(dsquery user -samid $manager))
                        {
                        $errorcount++
                        $LogErrors += "Entered manager does not exist in AD; "
                        }
                    }
        # If Job Title is blank, error
            If (($job -eq "") -or ($job -eq $null))
            {
            $errorcount++
            $LogErrors += "Job title is blank; "
            }
        # If Dept is blank, error.
            If (($dept -eq "") -or ($dept -eq $null))
            {
            $errorcount++
            $LogErrors += "Department is blank"
            }

        ## If errors are present, highlight issues.
        If ($errorcount -ne 0)
            {
            $LogMSG += "There are $errorcount error(s). Cannot create user $logonname`n"
            $LogMSG += $LogErrors
            $LogStatus = "WARNING"
            Update-LogFile $LogStatus $LogModule $LogMSG
            }

        ## If no errors, create user.
        If ($errorcount -eq 0)
            {


            ### Set default attributes and pulse vaults if applicable ###

            Get-DefaultAtts

    
            ## Add the Pulse vault security groups if Pulse vaults checked. (OMA vault is added by default to OMA users so no need to add twice)
            IF (($omavault -eq $true) -and ($Site -ne "OMA"))
                {
                $Groups = "$Groups" + "," + "OMA-Data_PDM_A"
                }                    
    
            IF ($dgnvault -eq $true)
                {   
                $Groups = "$Groups" + "," + "DGNApp_Pulse"
                }   


        ### This section creates the user account per the details in the CSV. ###    

            # Create User account
            $logStatus = "INFO"
            $Global:LogMSG = "Attempting to create AD account for $logonname"
            Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
            New-ADUser -SamAccountName $logonname -AccountExpirationDate $ExpireDate -userPrincipalName $userPrincipalName -GivenName $firstname -Surname $lastname -initials $initial -displayName $displayname -department $dept -Name $displayname -Path $SiteOU -Description $job -Title $job -enable $True -AccountPassword (ConvertTo-SecureString -AsPlainText 'TempPwd2017' -Force) -homeDirectory $homeDirectory -homeDrive $homeDrive -company $company -city $L_city -office $physicalDeliveryOfficeName -postalCode $postalCode -scriptPath $scriptPath -state $st -streetAddress $streetAddress -manager $manager
            Set-ADUser –Identity $logonname –ChangePasswordAtLogon $true

            # Check that user account was created successfully before adding to groups
            IF (dsquery user -samid $logonname)
                {
                $Groups = $Groups.split(",")
                foreach ($Group in $Groups)
                    {
                    Add-ADGroupMember -Identity $Group $logonname
                    }  
        
                # Display success message
                $logMSG = "AD account for $logonname has been created successfully!"
                $LogStatus = "SUCCESS"
                Update-LogFile $LogStatus $LogModule $LogMSG
                } ELSE 
                      {
                       # An error has occured.
                       $LogMSG = "Failed to create user account $logonname. Job aborted!"
                       $LogStatus = "ERROR"
                        Update-LogFile $LogStatus $LogModule $LogMSG
                       }


            }

        } # end foreach NewUserData in multiuserCSV
    
    # Finished processing CSV.
    $LogMSG = "Finished processing CSV: " + $WPFtxtbox_ADNU_filebrowse.text
    $LogStatus = "INFO"
    Update-LogFile $LogStatus $LogModule $LogMSG
    Get-module | Remove-Module
}

function Get-NetworkDetails { 
    param ($strComputer) 
    # PowerShell cmdlet to interrogate the Network Adapter

    $colItems = Get-WmiObject -class "Win32_NetworkAdapterConfiguration" `
        -computername $strComputer | Where {($_.description -match "Ethernet") -or ($_.description -match "Wireless") -or ($_.description -match "Intel") -or ($_.description -match "Broadcom")} | Select Description , MACAddress , IPAddress , IPEnabled , DNSServerSearchOrder
    $messageboxbody = "$strComputer Network Adapters"

    foreach ($objItem in $colItems) {
            $messageboxbody = $messageboxbody + "`n"
            $messageboxbody = $messageboxbody + "`nDescription: " + $objItem.Description
            $messageboxbody = $messageboxbody + "`nMACAddress: " + $objItem.MACAddress
            $messageboxbody = $messageboxbody + "`nIPAddress: " + $objItem.IPAddress
            $messageboxbody = $messageboxbody + "`nIPEnabled: " + $objItem.IPEnabled
            $messageboxbody = $messageboxbody + "`nDNSServers: " + $objItem.DNSServerSearchOrder
        } 
    Display-Info
} 

#=======================#
# Main Launcher Buttons #
#=======================#

# When clicked the buttons will execute any pre-run commands saved in the CSV file before executing the application.

### Button A1 ###
$WPFbtn_A1.Add_Click({ 
$Global:LogModule = "AppLauncher"
If (($prerun_A1 -ne $null) -and ($prerun_A1 -ne ""))
    {
    Invoke-Expression $prerun_A1
    }
    Start-Process $cmd_A1
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_A1.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
})
 
### Button A2 ###
$WPFbtn_A2.Add_Click({ 
$Global:LogModule = "AppLauncher"
If (($prerun_A2 -ne $null) -and ($prerun_A2 -ne ""))
    {
    Invoke-Expression $prerun_A2
    }
    Start-Process $cmd_A2
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_A2.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
})

### Button A3 ###
$WPFbtn_A3.Add_Click({ 
$Global:LogModule = "AppLauncher"
If (($prerun_A3 -ne $null) -and ($prerun_A3 -ne ""))
    {
    Invoke-Expression $prerun_A3
    }
    Start-Process $cmd_A3
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_A3.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
})

### Button A4 ###
$WPFbtn_A4.Add_Click({ 
$Global:LogModule = "AppLauncher"
If (($prerun_A4 -ne $null) -and ($prerun_A4 -ne ""))
    {
    Invoke-Expression $prerun_A4
    }
    Start-Process $cmd_A4
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_A4.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
}) 

### Button A5 ###
$WPFbtn_A5.Add_Click({
$Global:LogModule = "AppLauncher"
If (($prerun_A5 -ne $null) -and ($prerun_A5 -ne ""))
    {
    Invoke-Expression $prerun_A5
    }
    Start-Process $cmd_A5
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_A5.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
})

### Button A5 ###
$WPFbtn_A6.Add_Click({
$Global:LogModule = "AppLauncher"
If (($prerun_A6 -ne $null) -and ($prerun_A6 -ne ""))
    {
    Invoke-Expression $prerun_A6
    }
    Start-Process $cmd_A6
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_A6.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
})

### Button B1 ###
$WPFbtn_B1.Add_Click({
$Global:LogModule = "AppLauncher"
If (($prerun_B1 -ne $null) -and ($prerun_B1 -ne ""))
    {
    Invoke-Expression $prerun_B1
    }
    Start-Process $cmd_B1
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_B1.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
}) 

### Button B2 ###
$WPFbtn_B2.Add_Click({ 
$Global:LogModule = "AppLauncher"
If (($prerun_B2 -ne $null) -and ($prerun_B2 -ne ""))
    {
    Invoke-Expression $prerun_B2
    }
    Start-Process $cmd_B2
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_B2.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
}) 

### Button B3 ###
$WPFbtn_B3.Add_Click({
$Global:LogModule = "AppLauncher"
If (($prerun_B3 -ne $null) -and ($prerun_B3 -ne ""))
    {
    Invoke-Expression $prerun_B3
    }
    Start-Process $cmd_B3
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_B3.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
})

### Button B4 ###
$WPFbtn_B4.Add_Click({
$Global:LogModule = "AppLauncher"
If (($prerun_B4 -ne $null) -and ($prerun_B4 -ne ""))
    {
    Invoke-Expression $prerun_B4
    }
    Start-Process $cmd_B4
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_B4.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
}) 

### Button B5 ###
$WPFbtn_B5.Add_Click({
$Global:LogModule = "AppLauncher"
If (($prerun_B5 -ne $null) -and ($prerun_B5 -ne ""))
    {
    Invoke-Expression $prerun_B5
    }
    Start-Process $cmd_B5
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_B5.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
})

### Button B6 ###
$WPFbtn_B5.Add_Click({
$Global:LogModule = "AppLauncher"
If (($prerun_B6 -ne $null) -and ($prerun_B6 -ne ""))
    {
    Invoke-Expression $prerun_B6
    }
    Start-Process $cmd_B6
    $Global:LogStatus = "INFO"
    $Global:LogMSG = $WPFbtn_B6.Content + " started."
    Update-Logfile $Global:LogStatus $Global:LogModule $Global:LogMSG
})

#======================#
# AD Tools             #
#======================#

#===============#
### Listeners ###
#===============#

# When a selection is made in the Site combobox, enable the Kit & Group Type comboboxes. When selection is removed, disable Computer & Group Type comboboxes and btn_computeracc. Update Computer/group name fields automatically with correct computer/group name prefix.
$WPFcombobox_ADCmds_site.Add_SelectionChanged({
    $Site = $WPFcombobox_ADCmds_site.SelectedValue

    # If site has selection and Kit/Grp type comboboxes do not, enable Kit/Grp type comboboxes
    If ( (($Site -ne " ") -and ($Site -ne "") -and ($Site -ne $null)) -and (($WPFcombobox_kittype.selectedvalue -eq " ") -or ($WPFcombobox_kittype.selectedvalue -eq $null)) -and (($WPFcombobox_Grptype.selectedvalue -eq " ") -or ($WPFcombobox_Grptype.selectedvalue -eq $null)) )
        {
        $WPFcombobox_kittype.IsEnabled = $true
        $WPFcombobox_Grptype.IsEnabled = $true
        } 
   
    # If computername field has text 
    If (($WPFtxtbox_ADCmds_computername.Text -ne " ") -and ($WPFtxtbox_ADCmds_computername.Text -ne $null) -and ($WPFtxtbox_ADCmds_computername.Text -ne ""))
        {
        # Update site code in string contained in computername field
        $WPFtxtbox_ADCmds_computername.Text = $site + $Global:kit + $IDno
        # If a complete computername is present (eg SITEW0123) extract the ID number to the $IDno variable
        $IDno = [regex]::Matches($WPFtxtbox_ADCmds_computername.Text, "\d+(?!.*\d+)").value
        }

    # If group name field has text, update site code in string contained in computername field
    If (($WPFtxtbox_ADCmds_group.Text -ne " ") -and ($WPFtxtbox_ADCmds_group.Text -ne $null) -and ($WPFtxtbox_ADCmds_group.Text -ne ""))
        {
        $WPFtxtbox_ADCmds_group.Text = $site + $Global:Grptype + $Global:Grpname
        $Global:PrevGrptype = $site  + $Global:Grptype
        }


    # If site combobox is reset, disable kit/group type comboboxes, clear values in computer/group name fields, disable create group button and clear and disable Dir path field.
    If (($Site -eq $null) -or ($Site -eq "") -or ($Site -eq " "))
        {
        $WPFcombobox_kittype.IsEnabled = $false
        $WPFcombobox_Grptype.IsEnabled = $false
        $WPFcombobox_kittype.SelectedValue = $null
        $WPFcombobox_Grptype.SelectedValue = $null
        $WPFtxtbox_ADCmds_computername.Text = $null
        $WPFtxtbox_ADCmds_group.Text = $null
        $Global:kit = $null
        $Global:Grpname = $null
        $WPFbtn_computeracc.IsEnabled = $false
        $WPFtxtbox_ADCmds_dirpath.Text = $null
        $WPFbtn_createGroup.IsEnabled = $false
        }
})

# When a selection is made in the Kit type combobox, enable btn_computeracc. When selection is removed, disable btn_computeracc. Update Computer name field automatically with correct computer name prefix.
$WPFcombobox_kittype.Add_SelectionChanged({
    $Site = $WPFcombobox_ADCmds_site.SelectedValue
    $kittype = $WPFcombobox_kittype.SelectedValue
    If (($WPFtxtbox_ADCmds_computername.Text -ne " ") -and ($WPFtxtbox_ADCmds_computername.Text -ne $null) -and ($WPFtxtbox_ADCmds_computername.Text -ne ""))
        {
        $IDno = [regex]::Matches($WPFtxtbox_ADCmds_computername.Text, "\d+(?!.*\d+)").value
        }
    If($kittype -eq "Notebook")
        {
        $Global:kit = "NB"
        $WPFbtn_computeracc.IsEnabled = $true
        } 
        
    If($kittype -eq "Workstation")
        {
        $Global:Kit = "W"
        $WPFbtn_computeracc.IsEnabled = $true
        } 
                
    If(($Kittype -eq " ") -or ($Kittype -eq "") -or ($Kittype -eq $null))
        {
        $WPFbtn_computeracc.IsEnabled = $false
        $Global:Kit = $null
        }
                      
            
    # If the kittype combobox has a selection, disable the grptype combobox and dir path text box and reset their values to null
    If (($kittype -ne "") -and ($kittype -ne " ") -and ($kittype -ne $null))
        {
        $WPFbtn_nextavailacc.IsEnabled = $true
        $WPFcombobox_Grptype.IsEnabled = $false
        $WPFcombobox_Grptype.SelectedValue = $null
        $Global:Grpname = $null
        $WPFtxtbox_ADCmds_dirpath.Text = $null
        $WPFbtn_createGroup.IsEnabled = $false
        $WPFtxtbox_ADCmds_group.Text = $null
        $WPFbtn_exportMembers.IsEnabled = $false
        $WPFbtn_createGroup.IsEnabled = $false
        $WPFbtn_addGrpAuthor.IsEnabled = $false
        $WPFbtn_addread.IsEnabled = $false
        $WPFtxtbox_ADCmds_group.Text = $null
        $WPFtxtbox_ADCmds_computername.Text = $site + $Global:kit + $IDno
        }
    # If the kittype combobox is reset, 
    If (($kittype -eq "") -or ($kittype -eq " ") -or ($kittype -eq $null))
        {
        $WPFcombobox_Grptype.IsEnabled = $true
        $WPFtxtbox_ADCmds_computername.Text = $null
        $WPFbtn_nextavailacc.IsEnabled = $false
        }
})

# When a selection is made in the Grptype combobox, enable btn_createGroup. When selection is removed, disable btn_computeracc. Update Computer name field automatically with correct computer name prefix.
$WPFcombobox_Grptype.Add_SelectionChanged({
    $Site = $WPFcombobox_ADCmds_site.SelectedValue
    $Global:Grptype = $WPFcombobox_grptype.SelectedValue

    If(($Global:grptype -eq "") -or ($Global:grptype -eq " ") -or ($Global:grptype -eq $null))
        {
        $WPFbtn_createGroup.IsEnabled = $false
        $WPFcombobox_kittype.IsEnabled = $true
        $WPFtxtbox_ADCmds_group.text = $null
        }
    # If the grptype combobox has a selection, disable the grptype combobox and reset it's value to null
    If (($Global:Grptype -ne "") -and ($Global:Grptype -ne " ") -and ($Global:Grptype -ne $null))
        {
        $WPFcombobox_kittype.IsEnabled = $false
        $WPFcombobox_kittype.SelectedValue = $null
        $WPFtxtbox_ADCmds_dirpath.IsEnabled = $true
        $WPFtxtbox_ADCmds_computername.text = $null

        If (($WPFtxtbox_ADCmds_group.Text -ne "") -and ($WPFtxtbox_ADCmds_group.Text -ne " ") -and ($WPFtxtbox_ADCmds_group.Text -ne $null)){
            $Global:Grpname = $WPFtxtbox_ADCmds_group.Text
            $Global:Grpname = $Grpname.Replace($Global:PrevGrptype,'')
            $Global:Grpname = $Grpname.TrimEnd("_A")
            $Global:Grpname = $Grpname.TrimEnd("_R")

            If (($Global:grpname -ne "") -and ($Global:grpname -ne " ") -and ($Global:grpname -ne $null))
                {
                $WPFtxtbox_ADCmds_dirpath.IsEnabled = $true
                $WPFbtn_createGroup.IsEnabled = $false
                }
            }
        $WPFtxtbox_ADCmds_group.Text = $site + $Global:Grptype + $Global:Grpname
        }

    $Global:PrevGrptype = $site  + $Global:Grptype

})

# Entry in Username field
$WPFtxtbox_ADCmds_username.Add_TextChanged({

    # If text is entered in the username field
    If (($WPFtxtbox_ADCmds_username.text -ne " ") -and ($WPFtxtbox_ADCmds_username.text -ne "") -and ($WPFtxtbox_ADCmds_username.text -ne $null)){
        $WPFbtn_unlock.IsEnabled = $true
        $WPFbtn_resetPWD.IsEnabled = $true
        $WPFbtn_exportGroups.IsEnabled = $true

        #If text in username and computername field
        If (($WPFtxtbox_ADCmds_computername.text -ne " ") -and ($WPFtxtbox_ADCmds_computername.text -ne "") -and ($WPFtxtbox_ADCmds_computername.text -ne $null)){
            $WPFbtn_addlocaladmin.IsEnabled = $true
            $WPFbtn_remlocalAdmin.IsEnabled = $true
            }
        #If text in username and group field
        If (($WPFtxtbox_ADCmds_group.text -ne " ") -and ($WPFtxtbox_ADCmds_group.text -ne "") -and ($WPFtxtbox_ADCmds_group.text -ne $null)){
            $WPFbtn_addusertoGroup.IsEnabled = $true
            $WPFbtn_rmvUserfromGroup.IsEnabled = $true
            }
        }

    # If text is removed from the username field
    If (($WPFtxtbox_ADCmds_username.text -eq " ") -or ($WPFtxtbox_ADCmds_username.text -eq "") -or ($WPFtxtbox_ADCmds_username.text -eq $null)){
        $WPFbtn_unlock.IsEnabled = $false
        $WPFbtn_resetPWD.IsEnabled = $false
        $WPFbtn_exportGroups.IsEnabled = $false    
        $WPFbtn_addlocaladmin.IsEnabled = $false
        $WPFbtn_remlocalAdmin.IsEnabled = $false   
        $WPFbtn_addusertoGroup.IsEnabled = $false
        $WPFbtn_rmvUserfromGroup.IsEnabled = $false                     
        }
})

# Entry in Computername field
$WPFtxtbox_ADCmds_computername.Add_TextChanged({
    
    # If text is entered in the computername field
    If (($WPFtxtbox_ADCmds_computername.text -ne " ") -and ($WPFtxtbox_ADCmds_computername.text -ne "") -and ($WPFtxtbox_ADCmds_computername.text -ne $null)){
        $WPFbtn_loggedOn.IsEnabled = $True
        $WPFbtn_lastBoot.IsEnabled = $True
        $WPFbtn_discomputeracc.IsEnabled = $True
        $WPFbtn_encomputeracc.IsEnabled = $True
        $WPFbtn_chknetconf.IsEnabled = $True

        # If text is entered in Username and Computername fields
        If (($WPFtxtbox_ADCmds_username.text -ne " ") -and ($WPFtxtbox_ADCmds_username.text -ne "") -and ($WPFtxtbox_ADCmds_username.text -ne $null)){
            $WPFbtn_addlocaladmin.IsEnabled = $true
            $WPFbtn_remlocalAdmin.IsEnabled = $true
            }

        }

    # If text is removed from the username field
    If (($WPFtxtbox_ADCmds_computername.text -eq " ") -or ($WPFtxtbox_ADCmds_computername.text -eq "") -or ($WPFtxtbox_ADCmds_computername.text -eq $null)){
        $WPFbtn_addlocaladmin.IsEnabled = $false
        $WPFbtn_remlocalAdmin.IsEnabled = $false
        $WPFbtn_loggedOn.IsEnabled = $false
        $WPFbtn_lastBoot.IsEnabled = $false
        $WPFbtn_discomputeracc.IsEnabled = $false
        $WPFbtn_encomputeracc.IsEnabled = $false
        $WPFbtn_nextavailacc.IsEnabled = $false  
        $WPFbtn_chknetconf.IsEnabled = $false      
        }
})

# When an entry is made in the group field.
$WPFtxtbox_ADCmds_group.Add_TextChanged({
    
    # If text is entered in group name field
    If (($WPFtxtbox_ADCmds_group.text -ne " ") -and ($WPFtxtbox_ADCmds_group.text -ne "") -and ($WPFtxtbox_ADCmds_group.text -ne $null)){
        $WPFbtn_createGroup.IsEnabled = $true
        $WPFbtn_exportMembers.IsEnabled = $true

        # If Dirpath and Group fields have entries, enable the Add Group as Author and Add Group as Read Only buttons
        If (($WPFtxtbox_ADCmds_dirpath.text -ne "") -and ($WPFtxtbox_ADCmds_dirpath.text -ne " ") -and ($WPFtxtbox_ADCmds_dirpath.text -ne $null)){
            $WPFbtn_addGrpAuthor.IsEnabled = $true
            $WPFbtn_addread.IsEnabled = $true
            }
        # If Username and Group fields have entries, enable the Add Group as Author and Add Group as Read Only buttons
        If(($WPFtxtbox_ADCmds_username.text -ne " ") -and ($WPFtxtbox_ADCmds_username.text -ne "") -and ($WPFtxtbox_ADCmds_username.text -ne $null))
            {
            $WPFbtn_addUsertoGroup.IsEnabled = $true
            $WPFbtn_rmvUserfromGroup.IsEnabled = $true
            }
        }

    # If no text is entered in group name field
    If (($WPFtxtbox_ADCmds_group.text -eq " ") -or ($WPFtxtbox_ADCmds_group.text -eq "") -or ($WPFtxtbox_ADCmds_group.text -eq $null)){
        $WPFbtn_createGroup.IsEnabled = $false
        $WPFbtn_addGrpAuthor.IsEnabled = $false
        $WPFbtn_addread.IsEnabled = $false
        $WPFbtn_addUsertoGroup.IsEnabled = $false
        $WPFbtn_rmvUserfromGroup.IsEnabled = $false
        $WPFbtn_exportMembers.IsEnabled = $false
        }

})

# Entry in Dir Path field
$WPFtxtbox_ADCmds_dirpath.Add_TextChanged({

    # If Dirpath and Group fields have entries, enable the Add Group as Author and Add Group as Read Only buttons
    If (($WPFtxtbox_ADCmds_dirpath.text -ne "") -and ($WPFtxtbox_ADCmds_dirpath.text -ne " ") -and ($WPFtxtbox_ADCmds_dirpath.text -ne $null) -and ($WPFtxtbox_ADCmds_group.text -ne "") -and ($WPFtxtbox_ADCmds_group.text -ne " ") -and ($WPFtxtbox_ADCmds_group.text -ne $null)){
        $WPFbtn_addGrpAuthor.IsEnabled = $true
        $WPFbtn_addread.IsEnabled = $true
        }
    # If Dirpath or Group fields have no entry, disable the Add Group as Author and Add Group as Read Only buttons
    If (($WPFtxtbox_ADCmds_dirpath.text -eq "") -or ($WPFtxtbox_ADCmds_dirpath.text -eq " ") -or ($WPFtxtbox_ADCmds_dirpath.text -eq $null) -or ($WPFtxtbox_ADCmds_group.text -eq "") -or ($WPFtxtbox_ADCmds_group.text -eq " ") -or ($WPFtxtbox_ADCmds_group.text -eq $null)){
        $WPFbtn_addGrpAuthor.IsEnabled = $false
        $WPFbtn_addread.IsEnabled = $false
        }

})

#===============#
###  Actions  ###
#===============#

########## USERS ##########

### Unlock user account ###
$WPFbtn_unlock.Add_Click({
    $Global:LogModule = "UnlockAccount"
    $username = $WPFtxtbox_ADCmds_username.Text

    If (($Username -eq "") -or ($Username -eq " ") -or ($Username -eq $null))
        {
        $Messageboxbody = "Username cannot be blank."
        Display-Warning
        } 
        Else {
             $username = $username.trim()
             $username =$username.replace(' ','.')
             $unlockerror = $null
             $attempt = 0
                If (dsquery user -samid $username)
                        {
                        Do {
                            $Lockchk = (Get-ADUser -identity $username -Properties LockedOut | Select-Object LockedOut)
                            If ($Lockchk.LockedOut -eq $true)
                                {
                                    Unlock-ADAccount -identity $username
                                    $Lockchk = (Get-ADUser -identity $username -Properties LockedOut | Select-Object LockedOut)
                                    If ($Lockchk.LockedOut -eq $false)
                                    {
                                    $attempt = 5
                                    $Messageboxbody = "$Username" + "'s account has been unlocked!"
                                    Display-Success
                                    }
                                    Else {
                                        $attempt = $attempt++
                                        If ($attempt -ge 4)
                                            {
                                            $Messageboxbody = "AD account $Username cannot be unlocked!"
                                            Display-Error
                                            }
                                        }
                                    }
                            Else {
                                $attempt = 5
                                $Messageboxbody = "$Username" + "'s account is not locked. No action taken!"
                                Display-Info
                                }
                                }
                        While ($attempt -lt 4)
                                
                        } ELSE
                        {
                        $Messageboxbody = "$Username does not exist in AD. Please enter a valid username"
                        Display-Error
                        }
    Reset-ADCmdsForm
    Get-module | Remove-Module   
            }
})


### Password reset ###
$WPFbtn_resetPWD.Add_Click({
    $Global:LogModule = "PasswordReset"
    $username = $WPFtxtbox_ADCmds_username.Text

    If (($Username -eq "") -or ($Username -eq " ") -or ($Username -eq $null))
        {
        $Messageboxbody = "Username cannot be blank. Please enter a valid username"
        Display-WARNING
        } 
        Else {
             $username = $username.trim()
             $username =$username.replace(' ','.')
             $attempt = 0
                If (dsquery user -samid $username)
                        {
                        $CustomPaswd = $null
                        $Paswd = "TempPwd2017"
                        $CustomPaswd = Read-Host -prompt "Enter a new password. (or leave blank to use default: '$Paswd')"
                        If (($CustomPaswd -eq $null) -or ($CustomPaswd -eq ""))
                            {
                            $Paswd = $CustomPaswd
                            }
                        $SecPaswd= ConvertTo-SecureString –String "$Paswd" –AsPlainText –Force
                        Set-ADAccountPassword -Reset -NewPassword $SecPaswd –Identity $username
                        Unlock-ADAccount -identity $username
                        Set-ADUser –Identity $username –ChangePasswordAtLogon $true
                        $Messageboxbody = "$Username" + " password reset"
                        Display-Success
                        } ELSE
                        {
                        $Messageboxbody = "$Username does not exist in AD. Please enter a valid username"
                        Display-Error
                        }
    Reset-ADCmdsForm
    Get-module | Remove-Module 
            }

})

### Export user's group membership to CSV ###
$WPFbtn_exportGroups.Add_Click({
    $Global:LogModule = "ExportUserMembership"
    $Username = $WPFtxtbox_ADCmds_username.Text
    $username = $username.trim()
    $username =$username.replace(' ','.')

    If (($Username -eq "") -or ($Username -eq " ") -or ($Username -eq $null))
        {
        $Messageboxbody = "Cannot export user group membership to CSV. Username cannot be blank. Please enter a valid username"
        Display-WARNING
        }
        ELSE {
                IF (dsquery user -samid $username)
                    {
                    Get-ADPrincipalGroupMembership $Username | select name | Export-csv -path "C:\Member of Groups - $Username.csv" -NoTypeInformation
                    $Messageboxbody = "Group membership for $username has been exported to C:\Member of Groups - $Username.csv"
                    Display-Success

                    } ELSE
                    {
                    $Messageboxbody = "Cannot export user group membership to CSV. $Username does not exist in AD. Please enter a valid username"
                    Display-Warning
                    }
              }
    Reset-ADCmdsForm
    Get-module | Remove-Module 
})

### Add user to local admin group of remote machine ###
$WPFbtn_addlocalAdmin.Add_Click({
    $LogModule = "AddLocalAdmin"
    $ComputerName = $WPFtxtbox_ADCmds_computername.Text
    $UserName = $WPFtxtbox_ADCmds_username.text
    $username = $username.trim()
    $username =$username.replace(' ','.')

    #Check the machine is online.
    IF (-not (Test-Connection -comp $computername -quiet))
        {
        $Messageboxbody = "Unable to add $Username to local admin group. $computername is offline or unreachable"
        Display-Error
        }ELSE
            {
            # Check the username is valid
            IF (dsquery user -samid $username)
                {
                    # Check the computername is valid
                    IF (dsquery computer -name $computername)
                    {
                    $AdminGroup = [ADSI]"WinNT://$ComputerName/Administrators,group"
                    $User = [ADSI]"WinNT://$env:USERDOMAIN/$UserName,user"
                    $AdminGroup.add($User.Path)
                    $Messageboxbody = "$UserName added to local Administrators group on $computername"
                    Display-Success
                    }ELSE{
                        # Computername not valid
                        $Messageboxbody = "Unable to add $Username to local admin group. $computername does not exist in AD. Please enter a valid computer name"
                        Display-Error
                        }
                } ELSE 
                        {
                        # Username not valid
                        $Messageboxbody = "$username does not exist in AD. Please enter a valid user name to add to local admin group on $computername"
                        Display-Error
                        }
        }
    Reset-ADCmdsForm
    Get-Module | Remove-Module
})

### Remove user from local admin group of remote machine ###
$WPFbtn_remlocalAdmin.Add_Click({
    $Global:LogModule = "RemoveLocalAdmin"
    $ComputerName = $WPFtxtbox_ADCmds_computername.Text
    $UserName = $WPFtxtbox_ADCmds_username.text
    $username = $username.trim()
    $username =$username.replace(' ','.')
    #Check the machine is online.
    IF (-not (Test-Connection -comp $computername -quiet))
        {
        $Messageboxbody = "Unable to remove $Username from local admin group. $computername is offline or unreachable"
        Display-Error
        }ELSE{
            # Check the username is valid
            IF (dsquery user -samid $username)
                {
                    # Check the computername is valid
                    IF (dsquery computer -name $computername)
                    {
                    $AdminGroup = [ADSI]"WinNT://$ComputerName/Administrators,group"
                    $User = [ADSI]"WinNT://$env:USERDOMAIN/$UserName,user"
                    $AdminGroup.remove($User.Path)
                    $Messageboxbody = "$UserName removed from local Administrators group on $computername"
                    Display-Success
                    }ELSE{
                        # Computername not valid
                        $Messageboxbody = "Unable to remove $Username from local admin group. $computername does not exist in AD. Please enter a valid computer name"
                        Display-Error
                        }
                } ELSE {
                        # Username not valid
                        $Messageboxbody = "$username does not exist in AD. Please enter a valid user name to be removed from local admin group on $computername"
                        Display-Error
                        }
        }
    Reset-ADCmdsForm
    Get-module | Remove-Module 
})

### Add user to group ###
$WPFbtn_addUsertoGroup.Add_Click({
    $Global:LogModule = "AddUserToGrp"
    $logonname = $WPFtxtbox_ADCmds_username.text
    $logonname = $logonname.trim()
    $logonname = $logonname.replace(' ','.')
    $Groups = $WPFtxtbox_ADCmds_group.text
    $Groups = $Groups.Replace(' ',',')
    $Groups = $Groups.Replace(';',',')

    IF (dsquery user -samid $logonname)
        {
        $Groups = $Groups.split(",")
        foreach ($Group in $Groups)
            {
            Add-ADGroupMember -Identity $Group $logonname
            }  
        
        # Display success message
        $Messageboxbody = "All tasks have been completed successfully!"
        Display-Success
        } ELSE 
                  {
                   # An error has occured.
                   $Messageboxbody = "Failed to add $username to one or more group(s). Aborting job!"
                   Display-Error 
                   }
    Reset-ADCmdsForm
    Get-module | Remove-Module 
})

### Remove user from group ###
$WPFbtn_rmvUserfromGroup.Add_Click({
    $Global:LogModule = "RemoveUserFromGrp"
    $logonname = $WPFtxtbox_ADCmds_username.text
    $logonname = $logonname.trim()
    $logonname = $logonname.replace(' ','.')
    $Groups = $WPFtxtbox_ADCmds_group.text
    $Groups = $Groups.Replace(' ',',')
    $Groups = $Groups.Replace(';',',')

    IF (dsquery user -samid $logonname)
        {
        $Groups = $Groups.split(",")
        foreach ($Group in $Groups)
            {
            Remove-ADGroupMember -Identity $Group $logonname
            }  
        
        # Display success message
        $Messageboxbody = "$logonname removed from $Group"
        Display-Success
        } ELSE 
                  {
                   # An error has occured.
                   $Messageboxbody = "Failed to remove $username from one or more group(s). Aborting job!"
                   Display-Error 
                   }
    Reset-ADCmdsForm
    Get-module | Remove-Module
})

########## COMPUTERS ##########

### Who is logged onto computer ###
$WPFbtn_loggedOn.Add_Click({
    $Global:LogModule = "LoggedOnUser"
    $computername = $WPFtxtbox_ADCmds_computername.Text
    $computername = $computername.trim()
    
    If (($computername -eq "") -or ($computername -eq " ") -or ($computername -eq $null))
        {
        $Messageboxbody = "Cannot check logged on user as Computer name cannot be blank."
        Display-WARNING
        }
        ELSE {
             #Check the machine is online.
             IF (-not (Test-Connection -comp $computername -quiet))
                {
                $Messageboxbody = "Cannot check logged on user as $computername is offline or unreachable"
                Display-Error
                }ELSE{
                    IF (dsquery computer -name $computername)
                        {
                        $output = wmic.exe /node:$computername computersystem get username
                        $output = $Output.replace(' ','')
                        $output = $Output.replace('UserName','')
                        $output = $Output.trim()
                        $Messageboxbody = "User logged onto $computername is: `n" + $output
                        Display-Success
                        }ELSE{
                            $Messageboxbody = "Cannot check logged on user as $computername does not exist in AD."
                            Display-Error
                            }
                    }
            }
    Reset-ADCmdsForm
    Get-module | Remove-Module 
})


### Last boot of computer ###
$WPFbtn_lastBoot.Add_Click({
    $Global:LogModule = "LastBoot"
    $computername = $WPFtxtbox_ADCmds_computername.Text
    $computername = $computername.trim()
    
    If (($computername -eq "") -or ($computername -eq " ") -or ($computername -eq $null))
        {
        $Messageboxbody = "Cannot check last boot as Computer name is blank."
        Display-Warning
        }ELSE{
             #Check the machine is online.
             IF (-not (Test-Connection -comp $computername -quiet))
                {
                $Messageboxbody = "Cannot check last boot as $computername is offline or unreachable"
                Display-Error
                }ELSE{
                     IF (dsquery computer -name $computername)
                        {
                        $Booted = Get-WmiObject -Class Win32_OperatingSystem -Computer $Computername
                        $Output = $Booted.ConvertToDateTime($Booted.LastBootUpTime) | Out-String
                        $Messageboxbody = "Last boot of $computername is: $output"
                        Display-Success

                        }ELSE
                            {
                            $Messageboxbody = "Cannot check last boot as $computername does not exist in AD."
                            Display-Error
                            }
                      }
            }
    Reset-ADCmdsForm
    Get-module | Remove-Module 
})


### Disable computer account ###
$WPFbtn_discomputeracc.Add_Click({
    $Global:LogModule = "DisableComputer" 
    $computername = $WPFtxtbox_ADCmds_computername.Text
    $computername = $computername.trim()
    $samaccountname = $Computername + "$"
    If (dsquery computer -name $computername){
    Disable-ADAccount -Identity $samaccountname
    Reset-ADCmdsForm
    Get-module | Remove-Module
    $Messageboxbody = "$Computername disabled"
    Display-Success
    } Else {
        $Messageboxbody = "$Computername computer account doesn't exist in AD"
        Display-Error
        }
})

### Enable computer account ###
$WPFbtn_encomputeracc.Add_Click({
    $Global:LogModule = "EnableComputer"
    $computername = $WPFtxtbox_ADCmds_computername.Text
    $computername = $computername.trim()
    $samaccountname = $Computername + "$"
    If (dsquery computer -name $computername){
    Enable-ADAccount -Identity $samaccountname
    Reset-ADCmdsForm
    Get-module | Remove-Module
    $Messageboxbody = "$Computername enabled"
    Display-Success
    } Else {
        $Messageboxbody = "$Computername computer account doesn't exist in AD"
        Display-Error
        }
})

### Check for next available unused computer account before create computer account ###
$WPFbtn_nextavailacc.Add_Click({
    
    $Site = $WPFcombobox_ADCmds_site.SelectedValue
    $KitType = $WPFcombobox_kittype.SelectedValue
    $Global:SiteDefaultAtts = $Global:AllDefaultAtts | Where-Object{$_.SiteCode -eq $Site}
    If($kittype -eq "Notebook")
        {
        $OU = "OU=Notebooks," + $Global:SiteDefaultAtts.OU
        $kit = "NB"
        }
    If($kittype -eq "Workstation")
        {
        $OU = "OU=Workstations," + $Global:SiteDefaultAtts.OU
        $Kit = "W"
        }

    $activenumbers = @()
    $IDNumber = 001
    $ActiveComputers = Get-ADComputer -Filter * -Property Name -SearchBase $OU | Sort-Object -Property Name

    foreach($activecomputer in $ActiveComputers){
        $activenumbers += [regex]::Matches($ActiveComputer.name, "\d+(?!.*\d+)").value
    }

    $Activenumbers = $Activenumbers | Sort-Object
    $Activenumbers = $Activenumbers | select -Unique

    foreach ($activenumber in $activenumbers){
        If ($IDNumber -eq $activenumber)
            {
            $IDNumber++
            If ($kit -eq "NB")
                {
                $IDNo = "{0:D3}" -f $IDNumber
                }
            If ($kit -eq "W")
                {
                $IDNo = "{0:D4}" -f $IDNumber
                }

            $computersam = $Site + $kit + $IDNo
            If (dsquery computer -samid $computersam)
                {
                $IDNumber++
                }

            }
    }


    If ($kit -eq "NB")
        {
        $IDNo = "{0:D3}" -f $IDNumber
        }
    If ($kit -eq "W")
        {
        $IDNo = "{0:D4}" -f $IDNumber
        }

    $WPFtxtbox_ADCmds_computername.text = $Site + $kit + $IDNo
    Get-module | Remove-Module 
})

### Create computer account ###
# When btn_computeracc is clicked, validate data and create the computer account
$WPFbtn_computeracc.Add_Click({
   $Global:LogModule = "CreateComputerAcc"
   $computername = $WPFtxtbox_ADCmds_computername.Text
   $Site = $WPFcombobox_ADCmds_site.SelectedValue
   $KitType = $WPFcombobox_kittype.SelectedValue
   $computername = $computername.trim()
    
    If ($computername -eq "")
        {
        $Messageboxbody = "Cannot create computer account as Computer name cannot be blank."
        Display-Warning
        }
        ELSE {
            If (($kittype -ne "Notebook") -and ($kittype -ne "Workstation"))
                {
                $Messageboxbody = "Cannot create computer account as Type of kit cannot be blank."
                Display-Warning
                }
                ELSE {
                     If ($Site -notin $Global:AllDefaultAtts.SiteCode)
                        {
                        $Messageboxbody = "Cannot create computer account as Site code cannot be blank."
                        Display-Warning
                        }
                        ELSE {
                             $Global:SiteDefaultAtts = $Global:AllDefaultAtts | Where-Object{$_.SiteCode -eq $Site}
                             If($kittype -eq "Notebook")
                                {
                                $OU = "OU=Notebooks," + $Global:SiteDefaultAtts.OU
                                $kit = "NB"
                                }
                            If($kittype -eq "Workstation")
                                {
                                $OU = "OU=Workstations," + $Global:SiteDefaultAtts.OU
                                $Kit = "W"
                                }
                            # Check requested computer name doesn't already exist #
                            $computersam = "$computername" + "$"
                            IF (dsquery computer -samid $computersam)
                                {
                                $Messageboxbody = "Requested computer name already exists in AD. Please enter a unique computer name"
                                Display-Error
                                } Else {
                                    New-ADComputer -Name $computername -SAMAccountName $computersam -Path $OU
                                    IF (dsquery computer -samid $computersam)
                                        {
                                        $Messageboxbody = "$computername computer account created!"
                                        Display-Success
                                        } Else {
                                        $Messageboxbody = "An error has occured. $computername computer account not created. Aborting!"
                                        Display-Error
                                        }
                                    $WPFcombobox_ADCmds_site.text = ""
                                    $WPFcombobox_kittype.text = ""
                                    $WPFcombobox_kittype.IsEnabled = $false
                                    $WPFbtn_computeracc.IsEnabled = $false
                                    $WPFtxtbox_ADCmds_computername.Text = ""
                                    }
                            $WPFtxtbox_ADCmds_computername.Text = $null
                             }
                     }
        }
    Reset-ADCmdsForm
    Get-module | Remove-Module 

})

### Check Network details of computer ###
$WPFbtn_chknetconf.Add_Click({
    
    $Global:LogModule = "ComputerNetDetails"
    $computername = $WPFtxtbox_ADCmds_computername.Text
    $computername = $computername.trim()

    If (($computername -eq "") -or ($computername -eq " ") -or ($computername -eq $null))
        {
        $Messageboxbody = "Cannot check last boot as Computer name is blank."
        Display-Warning
        }ELSE{
             #Check the machine is online.
             IF (-not (Test-Connection -comp $computername -quiet))
                {
                $Messageboxbody = "Cannot check last boot as $computername is offline or unreachable"
                Display-Error
                }ELSE{
                     IF (dsquery computer -name $computername)
                        {
                        #$strComputer = $computername
                        Get-NetworkDetails $computername
                        }ELSE
                            {
                            $Messageboxbody = "Cannot check last boot as $computername does not exist in AD."
                            Display-Error
                            }
                      }
            }
    Reset-ADCmdsForm
    Get-module | Remove-Module 


})


########## GROUPS ##########

### Export the members of a group ###
$WPFbtn_exportMembers.Add_Click({
    $Global:LogModule = "ExportGrpMembers"
    $Group = $WPFtxtbox_ADCmds_Group.Text
    $Group = $Group.trim()
    
    If ($Group -eq "")
        {
        $Messageboxbody = "Cannot export members of group as Group name cannot be blank."
        Display-Warning
        }
        ELSE {
            IF (dsquery group -samid $Group)
                {
                Get-ADGroupMember -identity $Group | select name | Export-csv -path "C:\Members of group - $Group.csv" -NoTypeInformation
                $Messageboxbody = "Group membership of $group has been exported to C:\Members of group - $Group.csv"
                Display-Success
                } ELSE
                    {
                    $Messageboxbody = "Cannot export members of group as $Group does not exist in AD."
                    Display-Error
                    }
            }
    Reset-ADCmdsForm
    Get-module | Remove-Module 
})

### Create new Group ###
$WPFbtn_createGroup.Add_Click({
    $Global:LogModule = "NewGroup"
    $Account = $WPFtxtbox_ADCmds_group.Text
    $Account = $Account.Trim()
    $Account = $Account.Replace(' ','_')
    $Path = $WPFtxtbox_ADCmds_dirpath.Text
    $Path = $Path.Trim()

    # Check for blanks
    If (($Account -eq " ") -or ($account -eq "") -or ($account -eq $null))
        {
        $Messageboxbody = "Cannot create new group as Group name is blank."
        Display-Warning
        } ELSE {
        $Global:SiteDefaultAtts = $Global:AllDefaultAtts | Where-Object{$_.SiteCode -eq $WPFcombobox_ADCmds_site.SelectedValue}
        $GrpPath = "OU=Groups," + $Global:SiteDefaultAtts.OU
        If (($Path -eq " ") -or ($Path -eq "") -or ($Path -eq $null))
            {
            If (($WPFcombobox_GrpType.SelectedItem -eq "-DST_") -or ($WPFcombobox_GrpType.SelectedItem -eq "-MBX-")){
                        $Account= $Account.Trim()
                        $Account= $Account.TrimEnd("_A")
                        $Account= $Account.TrimEnd("_R")
                        New-ADGroup -name $Account -GroupScope Universal -GroupCategory Security -DisplayName $Account -Path $GrpPath
                        If (dsquery group -name $Account)
                            {
                            $MessageboxBody = "$Account created successfully"
                            Display-Success
                            } ELSE {
                                $MessageboxBody = "$Account not created. Aborting operation"
                                Display-Error
                                }
                    }
            } ELSE {

                If (($WPFcombobox_GrpType.SelectedItem -eq "-Data_") -or ($WPFcombobox_GrpType.SelectedItem -eq "-Data_Dept_")){
                        $Account= $Account.Trim()
                        $Account= $Account.TrimEnd("_A")
                        $Account= $Account.TrimEnd("_R")
                        $AccessRights = "FullControl"
                        $Account = $Account + "_A"
                        $GrpDesc = $AccessRights + "access to " + $Path
                        New-ADGroup -name $Account -GroupScope Global -GroupCategory Security -DisplayName $Account -Path $GrpPath
                        If (dsquery group -name $Account)
                            {
                            $MessageboxBody = "$Account created successfully"
                            Display-Success
                            Set-GroupNTFSAccess $Account $Path $AccessRights

                            $AccessRights = "ReadAndExecute"
                            $Account= $Account.TrimEnd("_A")
                            $Account = $Account + "_R"
                            $GrpDesc = $AccessRights + "access to " + $Path
                            New-ADGroup -name $Account -GroupScope Global -GroupCategory Security -DisplayName $Account -Path $GrpPath
                            
                            If (dsquery group -name $Account)
                                {
                            
                                Set-GroupNTFSAccess $Account $Path $AccessRights
                                } ELSE {
                                    $MessageboxBody = "$Account not created. Aborting operation"
                                    Display-Error
                                    }

                            } ELSE {
                                $MessageboxBody = "$Account not created. Aborting operation"
                                Display-Error
                                }
                        }
                    }
        }
    Reset-ADCmdsForm
    Get-module | Remove-Module 
})

### Add Folder Author Access to Group ###
$WPFbtn_addGrpAuthor.Add_Click({
    $Global:LogModule = "GroupFolderAuthor"
    $Account = $WPFtxtbox_ADCmds_group.Text
    $Account = $Account.Trim()
    $Path = $WPFtxtbox_ADCmds_dirpath.Text
    $Path = $Path.Trim()
    $AccessRights = "FullControl"

    # Check for blanks
    If (($Account -eq " ") -or ($account -eq "") -or ($account -eq $null))
        {
        $Messageboxbody = "Cannot add group Author access as Group name is blank."
        Display-Warning
        } ELSE {
        If (($Path -eq " ") -or ($Path -eq "") -or ($Path -eq $null))
            {
            $Messageboxbody = "Cannot add group Author access as Directory path is blank."
            Display-Warning
            } ELSE {
                    Set-GroupNTFSAccess $Account $Path $AccessRights
                    }
        }
    Reset-ADCmdsForm
    Get-module | Remove-Module 
})

### Add Folder Read only Access to Group ###
$WPFbtn_addread.Add_Click({
    $Global:LogModule = "GroupFolderReadExec"
    $Account = $WPFtxtbox_ADCmds_group.Text
    $Account = $Account.Trim()
    $Path = $WPFtxtbox_ADCmds_dirpath.Text
    $Path = $Path.Trim()
    $AccessRights = "ReadAndExecute"

    # Check for blanks
    If (($Account -eq " ") -or ($account -eq "") -or ($account -eq $null))
        {
        $Messageboxbody = "Cannot add group Read access as Group name is blank."
        Display-Warning

        If (($Path -eq " ") -or ($Path -eq "") -or ($Path -eq $null))
            {
            $Messageboxbody = "Cannot add group Author access as Directory path is blank."
            Display-Warning
            } ELSE {
                    Set-GroupNTFSAccess $Account $Path $AccessRights
                    }
        }
    Reset-ADCmdsForm
    Get-module | Remove-Module 
})

#======================#
# Create new AD user   #
#======================#

#===============#
### Listeners ###
#===============#

## Single user radio button ##
$WPFradio_ADNU_Single.Add_Checked({

    Reset-UserInputVars
    Reset-ADNUErrors
    Reset-ADNUForm
    Reset-ADNUFormValues

})

## Multi user radio button ##
$WPFradio_ADNU_multi.Add_Checked({
    Reset-UserInputVars
    Reset-ADNUErrors
    Reset-ADNUFormValues

    $WPFradio_ADNU_multi.IsChecked = $true
    $WPFtxtbox_ADNU_filebrowse.IsEnabled = $true
    $WPFbtn_ADNU_importCSV.IsEnabled = $true

    $WPFcombo_ADNU_Site.IsEnabled = $false
    $WPFtxtbox_ADNU_firstname.IsEnabled = $false
    $WPFtxtbox_ADNU_lastname.IsEnabled = $false
    $WPFtxtbox_ADNU_initial.IsEnabled = $false
    $WPFchkbox_ADNU_site1.IsEnabled = $false
    $WPFchkbox_ADNU_site2.IsEnabled = $false
    $WPFchkbox_ADNU_temp.IsEnabled = $false
    $WPFtxtbox_ADNU_logonname.IsEnabled = $false
    $WPFtxtbox_ADNU_displayname.IsEnabled = $false
    $WPFtxtbox_ADNU_manager.IsEnabled = $false
    $WPFtxtbox_ADNU_title.IsEnabled = $false
    $WPFtxtbox_ADNU_dept.IsEnabled = $false
    $WPFbtn_ADNU_validate.IsEnabled = $false


})

## Enable Create User when CSV path is present ##
$WPFtxtbox_ADNU_filebrowse.Add_TextChanged({
    If (($WPFtxtbox_ADNU_filebrowse.text -eq "") -or ($WPFtxtbox_ADNU_filebrowse.text -eq " ") -or ($WPFtxtbox_ADNU_filebrowse.text -eq $null))
        {
        $WPFbtn_ADNU_Create.IsEnabled = $false
        }
    If (($WPFtxtbox_ADNU_filebrowse.text -ne "") -and ($WPFtxtbox_ADNU_filebrowse.text -ne " ") -and ($WPFtxtbox_ADNU_filebrowse.text -ne $null))
        {
        $WPFbtn_ADNU_Create.IsEnabled = $true
        } 
})

## Update logon name and display name when data is entered
$WPFtxtbox_ADNU_firstname.Add_TextChanged({
    $firstname = $WPFtxtbox_ADNU_firstname.Text
    $lastname = $WPFtxtbox_ADNU_lastname.Text
    $initial = $WPFtxtbox_ADNU_initial.Text
    $firstname = $firstname.Replace(' ','')
    $lastname = $lastname.Replace(' ','')
    $initial = $initial.Replace(' ','')
    $firstname = $firstname.Replace("'",'')
    $lastname = $lastname.Replace("'",'')

    if ($initial -eq "")
        {
        $displayname = $lastname + ", " + $firstname
        $logonname = $firstname + "." + $lastname
        }
    if($initial -ne "")
        {
        $displayname = $lastname + ", " + $firstname + " " + $initial
        $logonname = $firstname + "." + $initial + "." + $lastname
        }
    if (($initial -eq "") -and ($firstname -eq "") -and ($lastname -eq ""))
        {
        $displayname = ""
        $logonname = ""
        }
    $WPFtxtbox_ADNU_displayname.Text = $displayname
    $WPFtxtbox_ADNU_logonname.Text = $logonname

})

## Update logon name and display name when data is entered
$WPFtxtbox_ADNU_lastname.Add_TextChanged({
    $firstname = $WPFtxtbox_ADNU_firstname.Text
    $lastname = $WPFtxtbox_ADNU_lastname.Text
    $initial = $WPFtxtbox_ADNU_initial.Text
    $firstname = $firstname.Replace(' ','')
    $lastname = $lastname.Replace(' ','')
    $initial = $initial.Replace(' ','')
    $firstname = $firstname.Replace("'",'')
    $lastname = $lastname.Replace("'",'')

    if ($initial -eq "")
        {
        $displayname = $lastname + ", " + $firstname
        $logonname = $firstname + "." + $lastname
        }
    if($initial -ne "")
        {
        $displayname = $lastname + ", " + $firstname + " " + $initial
        $logonname = $firstname + "." + $initial + "." + $lastname
        }
    if (($initial -eq "") -and ($firstname -eq "") -and ($lastname -eq ""))
        {
        $displayname = ""
        $logonname = ""
        }
    $WPFtxtbox_ADNU_displayname.Text = $displayname
    $WPFtxtbox_ADNU_logonname.Text = $logonname

})

## Update logon name and display name when data is entered
$WPFtxtbox_ADNU_initial.Add_TextChanged({
    $firstname = $WPFtxtbox_ADNU_firstname.Text
    $lastname = $WPFtxtbox_ADNU_lastname.Text
    $initial = $WPFtxtbox_ADNU_initial.Text
    $firstname = $firstname.Replace(' ','')
    $lastname = $lastname.Replace(' ','')
    $initial = $initial.Replace(' ','')
    $firstname = $firstname.Replace("'",'')
    $lastname = $lastname.Replace("'",'')

    if ($initial -eq "")
        {
        $displayname = $lastname + ", " + $firstname
        $logonname = $firstname + "." + $lastname
        }
    if($initial -ne "")
        {
        $displayname = $lastname + ", " + $firstname + " " + $initial
        $logonname = $firstname + "." + $initial + "." + $lastname
        }
    if (($initial -eq "") -and ($firstname -eq "") -and ($lastname -eq ""))
        {
        $displayname = ""
        $logonname = ""
        }
    $WPFtxtbox_ADNU_displayname.Text = $displayname
    $WPFtxtbox_ADNU_logonname.Text = $logonname

})

## Enable expiry date picker if Temp is checked
$WPFchkbox_ADNU_temp.Add_Checked({
    $WPFdatepkr_ADNU_expiry.IsEnabled = $true
})

## Disable expiry date picker and reset to null if Temp is unchecked
$WPFchkbox_ADNU_temp.Add_UnChecked({
    $WPFdatepkr_ADNU_expiry.IsEnabled = $false
    $WPFdatepkr_ADNU_expiry.Text = $null
    $WPFtxtblock_ADNU_errorexpiry.Visibility = "Hidden"
})

#===============#
###  Actions  ###
#===============#

## Browse for CSV of multiple new users to be imported ##
$WPFbtn_ADNU_importCSV.Add_Click({
    $Global:LogModule = "LoadCSVMultiADNU"

    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Multiselect = $false # Multiple files can be chosen
	    Filter = 'CSV (*.csv)|*.csv' # Specified file types
        }
 
    [void]$FileBrowser.ShowDialog()

    $file = $FileBrowser.FileName;

    If($FileBrowser.FileNames -like "*\*") {

	    # Do something 
	    $FileBrowser.FileName #Lists selected files (optional)
	
        }

    $WPFtxtbox_ADNU_filebrowse.text = $file
    $LogStatus = "INFO"
    $LogMSG = "Imported CSV: $file"
    Update-LogFile $LogStatus $Global:LogModule $LogMSG
})

## Validate entered data on Validate button click
$WPFbtn_ADNU_validate.Add_Click({
    Reset-UserInputVars
    $errorcount = 0
    # Reset any highlighted errors from a previous validation
    Reset-ADNUErrors
    $Global:LogModule = "ADNUValidate"
    $LogMsg = @()
    
    #Collect user input into variables and clean the data
    $Site = $WPFcombo_ADNU_Site.SelectedValue

    $firstname = $WPFtxtbox_ADNU_firstname.Text
    If (($firstname -ne "") -and ($firstname -ne $null))
        {
        $firstname = $firstname.Replace(' ','')
        $firstname = $firstname.Replace("'",'')
        }

    $lastname = $WPFtxtbox_ADNU_lastname.Text
    If (($lastname -ne "") -and ($lastname -ne $null))
        {
        $lastname = $lastname.Replace(' ','')
        $lastname = $lastname.Replace("'",'')
        }

    $initial = $WPFtxtbox_ADNU_initial.Text
    If (($initial -ne "") -and ($initial -ne $null))
        {
        $initial = $initial.Replace(' ','')
        }

    $logonname = $WPFtxtbox_ADNU_logonname.Text
    If (($logonname -ne "") -and ($logonname -ne $null))
        {
        $logonname = $logonname.Replace(' ','')
        }

    $displayname = $WPFtxtbox_ADNU_displayname.Text
    If (($displayname -ne "") -and ($displayname -ne $null))
        {
        $displayname = $displayname.Trim()
        }

    $manager = $WPFtxtbox_ADNU_manager.Text
    If (($manager -ne "") -and ($manager -ne $null))
        {
        $manager = $manager.Trim()
        $manager = $manager.Replace(' ','.')
        }

    $job = $WPFtxtbox_ADNU_title.Text
    If(($job -ne "") -and ($job -ne $null))
        {
        $job = $job.Trim()
        }

    $dept = $WPFtxtbox_ADNU_dept.Text
    If (($dept -ne "") -and ($dept -ne $null))
        {
        $dept = $dept.Trim()
        }

    If ($WPFchkbox_ADNU_Site1.IsChecked -eq $true)
        {
        $omavault = $true
        }

    If ($WPFchkbox_ADNU_Site2.IsChecked -eq $true)
        {
        $dgnvault = $true
        }

    If ($WPFchkbox_ADNU_temp.IsChecked -eq $true)
        {
        $tempuser = $true
        }

    # If Temp User is checked but no expiry date entered, error.
    If (($WPFchkbox_ADNU_temp.IsChecked -eq $true) -and ($WPFdatepkr_ADNU_expiry.SelectedDate -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errorexpiry.Visibility = "Visible"
        $LogMSG += "Temp User is checked but no expiry date entered; "
        }

    If ($tempuser -eq $true)
        {
        $newDateTime = [DateTime]::Parse($WPFdatepkr_ADNU_expiry.Text) 
        $ExpireDate = $newDateTime.AddDays(1)
        }      

    # If the site has not been selected from the combobox, error.
    If (($Site -eq "") -or ($site -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errorsite.Visibility = "Visible"
        $LogMSG += "Site code not selected; "
        }

    # If the firstname is left blank, error
    If (($firstname -eq "") -or ($firstname -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errorfirst.Visibility = "Visible"
        $LogMSG += "First name is blank; "
        }

    # If the lastname is left blank, error
    If (($lastname -eq "") -or ($lastname -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errorlast.Visibility = "Visible"
        $LogMSG += "Last name is blank; "
        }
        
    # If logon name is blank, error
    If (($logonname -eq "") -or ($logonname -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errorlogon.Visibility = "Visible"
        $LogMSG += "Logon name is blank; "
        } ELSE {
                # Checks that the logon name is unique
                If (dsquery user -samid $logonname)
                    {
                    $errorcount++
                    $WPFtxtblock_ADNU_errorlogonunique.Visibility = "Visible"
                    $LogMSG += "$logonname is not unique; "
                    }
                }

    # If display name is blank, error
    If (($displayname -eq "") -or ($displayname -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errordisp.Visibility = "Visible"
        $LogMSG += "Display name is blank; "
        }

    # If manager name is blank, error
    If (($manager -eq "") -or ($manager -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errormgr.Visibility = "Visible"
        $LogMSG += "Manager name is blank; "
        } ELSE {
                # Check the inputted manager exists in AD
                If (!(dsquery user -samid $manager))
                    {
                    $errorcount++
                    $WPFtxtblock_ADNU_errormgrunique.Visibility = "Visible"
                    $LogMSG += "Entered manager does not exist in AD; "
                    }
                }
    # If Job Title is blank, error
        If (($job -eq "") -or ($job -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errortitle.Visibility = "Visible"
        $LogMSG += "Job title is blank; "
        }
    # If Dept is blank, error.
        If (($dept -eq "") -or ($dept -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errordept.Visibility = "Visible"
        $LogMSG += "Department is blank"
        }

    ## If errors are present, highlight issues.
    If ($errorcount -ne 0)
        {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "ERROR"
        $Messageboxbody = "There are $errorcount error(s). Please ensure to all reqired fields marked with excalamation marks ! are completed and action any other errors as displayed."
        $MessageIcon = [System.Windows.MessageBoxImage]::ERROR
        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
        $LogStatus = "WARNING"
        Update-LogFile $LogStatus $LogModule $LogMSG
        }

    ## If no errors, disable all fields so data cannot be changed. Then, enable the Edit (in case the user does want to edit the info) and Create User buttons.
    If ($errorcount -eq 0)
        {
        $WPFcombo_ADNU_Site.IsEnabled = $false
        $WPFtxtbox_ADNU_firstname.IsEnabled = $false
        $WPFtxtbox_ADNU_lastname.IsEnabled = $false
        $WPFtxtbox_ADNU_initial.IsEnabled = $false
        $WPFchkbox_ADNU_Site1.IsEnabled = $false
        $WPFchkbox_ADNU_Site2.IsEnabled = $false
        $WPFchkbox_ADNU_temp.IsEnabled = $false
        $WPFdatepkr_ADNU_expiry.IsEnabled = $false
        $WPFtxtbox_ADNU_logonname.IsEnabled = $false
        $WPFtxtbox_ADNU_displayname.IsEnabled = $false
        $WPFtxtbox_ADNU_manager.IsEnabled = $false
        $WPFtxtbox_ADNU_title.IsEnabled = $false
        $WPFtxtbox_ADNU_dept.IsEnabled = $false
        $WPFbtn_ADNU_validate.IsEnabled = $false
        $WPFbtn_ADNU_edit.IsEnabled = $true
        $WPFbtn_ADNU_Create.IsEnabled = $true
        $Messageboxbody = "User info has been validated. To make changes, click Edit. To proceed, click Create User."
        Display-Info
        }
    Get-module | Remove-Module  
})

## Reset the New AD users form
$WPFbtn_ADNU_resetform.Add_Click({
    $Global:LogModule = "ADNUReset"
    Reset-UserInputVars
    Reset-ADNUErrors
    Reset-ADNUForm
    Reset-ADNUFormValues
    $LogStatus = "INFO"
    $LogMSG = "ADNU form reset"
    Update-LogFile $LogStatus $LogModule $LogMSG
})

## If user wants to edit the data before submitting, the edit button will enable the data entry fields and disable the create user button until validated again.
$WPFbtn_ADNU_edit.Add_Click({
    $WPFcombo_ADNU_Site.IsEnabled = $true
    $WPFtxtbox_ADNU_firstname.IsEnabled = $true
    $WPFtxtbox_ADNU_lastname.IsEnabled = $true
    $WPFtxtbox_ADNU_initial.IsEnabled = $true
    $WPFchkbox_ADNU_Site1.IsEnabled = $true
    $WPFchkbox_ADNU_Site2.IsEnabled = $true
    $WPFchkbox_ADNU_temp.IsEnabled = $true
    If ($WPFchkbox_ADNU_temp.IsChecked)
        {
        $WPFdatepkr_ADNU_expiry.IsEnabled = $true
        }
    $WPFtxtbox_ADNU_logonname.IsEnabled = $true
    $WPFtxtbox_ADNU_displayname.IsEnabled = $true
    $WPFtxtbox_ADNU_manager.IsEnabled = $true
    $WPFtxtbox_ADNU_title.IsEnabled = $true
    $WPFtxtbox_ADNU_dept.IsEnabled = $true
    $WPFbtn_ADNU_validate.IsEnabled = $true
    $WPFbtn_ADNU_edit.IsEnabled = $false
    $WPFbtn_ADNU_Create.IsEnabled = $false

})

## Once the data has been validated the user can be created
$WPFbtn_ADNU_Create.Add_Click({

    If ($WPFradio_ADNU_Single.IsChecked -eq $true) {
    Create-SingleADUser
    $Global:LogModule = "SingleADNU"
    $LogStatus = "INFO"
    $LogMSG = "Create Single AD new user"
    Update-LogFile $LogStatus $LogModule $LogMSG
    }

    If ($WPFradio_ADNU_Multi.IsChecked -eq $true) {
    Create-MultiADUser
    $Global:LogModule = "MultiADNU"
    $LogStatus = "INFO"
    $LogMSG = "Create Multiple AD new user"
    Update-LogFile $LogStatus $LogModule $LogMSG
    }

})


#==============================================#
# Export details of AD user or all users in OU #
#==============================================#

## If radio button Single User is selected, username is enabled and site combobox is disabled
$WPFradio_EADUI_User.Add_Checked({
    $WPFtxtbox_EADUI_username.IsEnabled = $true
    $WPFcombo_EADUI_site.IsEnabled = $false
    $WPFcombo_EADUI_site.SelectedValue = $null
})


## If radio button Users in OU is selected, username is disabled and site combobox is enabled
$WPFradio_EADUI_OU.Add_Checked({
    $WPFcombo_EADUI_site.IsEnabled = $true
    $WPFtxtbox_EADUI_username.IsEnabled = $false
    $WPFtxtbox_EADUI_username.text = $null
})


## If All Attributes Check box is unchecked default to only Common attributes checked
$WPFchkbox_EADUI_AllAtts.Add_Unchecked({
    # Check the Common Atts check box when the All atts check box is unchecked
    $WPFchkbox_EADUI_common.IsChecked = $true
    $WPFchkbox_EADUI_common.IsEnabled = $true
    $WPFchkbox_EADUI_AdditionalAtts.IsEnabled = $true

    # Enable the common atts
    Enable-CommonAtts

    # Enable the additional atts
    Enable-AdditionalAtts

    # Check the common atts
    Check-CommonAtts

    # Uncheck the additional atts
    Uncheck-AdditionalAtts


})


## If All Attributes Check box is checked, common and additional atts are unchecked and all attributes are disabled and checked
$WPFchkbox_EADUI_AllAtts.Add_Checked({

    $WPFchkbox_EADUI_AdditionalAtts.IsChecked = $false
    $WPFchkbox_EADUI_common.IsChecked = $false
    $WPFchkbox_EADUI_common.IsEnabled = $false
    $WPFchkbox_EADUI_AdditionalAtts.IsEnabled = $false

    # Disable the common atts
    Disable-CommonAtts

    # Disable the additional atts
    Disable-AdditionalAtts

    # Check the common atts
    Check-CommonAtts

    #Check the additional atts
    Check-AdditionalAtts
})


## If Common Attributes Check box is checked, common atts are checked
$WPFchkbox_EADUI_Common.Add_Checked({

    # Check the common atts
    Check-CommonAtts

})


## If Common Attributes Check box is unchecked, common atts are unchecked
$WPFchkbox_EADUI_Common.Add_UnChecked({

    # Uncheck the common atts
    Uncheck-CommonAtts

})


## If Additional Attributes Check box is checked, additional atts are checked
$WPFchkbox_EADUI_AdditionalAtts.Add_Checked({

    # Check the additional atts
    Check-AdditionalAtts
       
})
 

## If Additional Attributes Check box is unchecked, additional atts are unchecked
$WPFchkbox_EADUI_AdditionalAtts.Add_UnChecked({

    # Uncheck the additional atts
    Uncheck-AdditionalAtts
       
})


## When the Export button is clicked, validate data, set attributes to be exported and then export
$WPFbtn_EADUI_export.Add_Click({
    
    $errorcount = 0
    $WPFtxtblock_EADUI_errorusername.Visibility = "Hidden"
    $WPFtxtblock_EADUI_erroruserunique.Visibility = "Hidden"
    $WPFtxtblock_EADUI_errorsite.Visibility = "Hidden"

    $username = $WPFtxtbox_EADUI_username.Text
    $username = $username.Trim()
    $username = $username.Replace(' ','.')

    $site = $WPFcombo_EADUI_site.SelectedValue

    If ($WPFradio_EADUI_User.IsChecked)
        {
        $Global:LogModule = "ExportADUserDetails"
        If (($username -eq $null) -or ($username -eq ""))
            {
            $errorcount++
            $WPFtxtblock_EADUI_errorusername.Visibility = "Visible"
            $WPFtxtbox_EADUI_username.Text = $username
            $LogStatus = "WARNING"
            $LogMSG = "Username is blank"
            Update-LogFile $LogStatus $LogModule $LogMSG
            } ELSE {
                    If (!(dsquery user -samid $username))
                        {
                        $errorcount++
                        $WPFtxtblock_EADUI_erroruserunique.Visibility = "Visible"
                        $WPFtxtbox_EADUI_username.Text = $username
                        $LogStatus = "WARNING"
                        $LogMSG = "$Username does not exist in AD."
                        Update-LogFile $LogStatus $LogModule $LogMSG
                        }
                     }
        }

    If ($WPFradio_EADUI_OU.IsChecked)
        {
        $Global:LogModule = "ExportADUsersinOUDetails"
        If (($site -eq $null) -or ($site -eq ""))
            {
            $errorcount++
            $WPFtxtblock_EADUI_errorsite.Visibility = "Visible"
            $LogStatus = "WARNING"
            $LogMSG = "Site code is blank"
            Update-LogFile $LogStatus $LogModule $LogMSG
            } 
        }

    If ($errorcount -eq 0)
        {
        $attributes = @()

        # Common Attributes
        If ($WPFchkbox_EADUI_userPrincipalName.IsChecked -eq $true)
            {
            $attributes += "UserPrincipalName"
            }
        
        If ($WPFchkbox_EADUI_samaccname.IsChecked -eq $true)
            {
            $attributes += "SamAccountName"
            }
                    
        If ($WPFchkbox_EADUI_display.IsChecked -eq $true)
            {
            $attributes += "DisplayName"
            }
                    
        If ($WPFchkbox_EADUI_givenName.IsChecked -eq $true)
            {
            $attributes += "givenName"
            }
                    
        If ($WPFchkbox_EADUI_initials.IsChecked -eq $true)
            {
            $attributes += "Initials"
            }
                    
        If ($WPFchkbox_EADUI_sn.IsChecked -eq $true)
            {
            $attributes += "Surname"
            }
                    
        If ($WPFchkbox_EADUI_email.IsChecked -eq $true)
            {
            $attributes += "EmailAddress"
            }
                    
        If ($WPFchkbox_EADUI_title.IsChecked -eq $true)
            {
            $attributes += "Title"
            }
                    
        If ($WPFchkbox_EADUI_dept.IsChecked -eq $true)
            {
            $attributes += "Department"
            }
                    
        If ($WPFchkbox_EADUI_desc.IsChecked -eq $true)
            {
            $attributes += "Description"
            }
                    
        If ($WPFchkbox_EADUI_mgr.IsChecked -eq $true)
            {
            $attributes += "Manager"
            }
            
        # Additional atts
        If ($WPFchkbox_EADUI_alias.IsChecked -eq $true)
            {
            $attributes += "MailNickName"
            }
            
        If ($WPFchkbox_EADUI_City.IsChecked -eq $true)
            {
            $attributes += "City"
            }
            
        If ($WPFchkbox_EADUI_company.IsChecked -eq $true)
            {
            $attributes += "Company"
            }
            
        If ($WPFchkbox_EADUI_Country.IsChecked -eq $true)
            {
            $attributes += "Country"
            }
            
        If ($WPFchkbox_EADUI_expiry.IsChecked -eq $true)
            {
            $attributes += "AccountExpirationDate"
            }
            
        If ($WPFchkbox_EADUI_hmedrv.IsChecked -eq $true)
            {
            $attributes += "HomeDrive"
            }
            
        If ($WPFchkbox_EADUI_hmefolder.IsChecked -eq $true)
            {
            $attributes += "HomeDirectory"
            }
            
        If ($WPFchkbox_EADUI_loginscpt.IsChecked -eq $true)
            {
            $attributes += "ScriptPath"
            }
            
        If ($WPFchkbox_EADUI_logonto.IsChecked -eq $true)
            {
            $attributes += "LogonWorkstations"
            }
            
        If ($WPFchkbox_EADUI_office.IsChecked -eq $true)
            {
            $attributes += "Office"
            }
            
        If ($WPFchkbox_EADUI_POBox.IsChecked -eq $true)
            {
            $attributes += "POBox"
            }
            
        If ($WPFchkbox_EADUI_cn.IsChecked -eq $true)
            {
            $attributes += "CanonicalName"
            }
            
        If ($WPFchkbox_EADUI_State.IsChecked -eq $true)
            {
            $attributes += "State"
            }
            
        If ($WPFchkbox_EADUI_street.IsChecked -eq $true)
            {
            $attributes += "StreetAddress"
            }
            
        If ($WPFchkbox_EADUI_telephone.IsChecked -eq $true)
            {
            $attributes += "OfficePhone"
            }
        
        $Global:SiteDefaultAtts = $AllDefaultAtts | Where-Object{$_.SiteCode -eq $Site}
        $OU = "OU=Users," + $Global:SiteDefaultAtts.OU

        If($WPFradio_EADUI_OU.IsChecked)
            {
            $exportdir = "C:\" + "$Site" + "_Users_export.csv"
            Get-ADUser -Filter * -SearchBase $OU -Properties $attributes | Export-csv $exportdir
            $Messageboxbody = "Users details in $OU exported to: $exportdir"
            Display-Success
            $WPFtxtbox_EADUI_username.text = $null
            }

        If ($WPFradio_EADUI_User.IsChecked)
            {
            $exportdir = "C:\" + "$Username" + "_export.csv"
            Get-ADUser -Identity $username -Properties $attributes | Export-csv $exportdir
            $Messageboxbody = "$username details exported to: $exportdir"
            Display-Success
            $WPFcombo_EADUI_site.SelectedValue = $null
                        }
        }
    Get-module | Remove-Module 
})



#===========================================================================#
# Imports settings and shows the GUI                                        #
#===========================================================================#
Reset-UserInputVars
Reset-ADNUErrors
Show-CredAndVsn
Load-CSVData
Load-LauncherButtons
Load-ComboBoxes
Load-Logs
#CLS
Write-Host -ForegroundColor Cyan "IT Admin Tool $Version"
Write-Host -ForegroundColor Cyan "Running with Credentials: $env:USERNAME"
Write-Host -ForegroundColor Red "Do NOT close this window. Please use the GUI to interact with the IT Admin Tool."
$Form.ShowDialog() | out-null
