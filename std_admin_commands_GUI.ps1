$version = "Version 1.1.4"


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
        Title="IT Admin Tool" Height="544.109" Width="525" ResizeMode="NoResize">
    <Grid>
        <TabControl Margin="0,0,0.4,0.4">
            <TabItem x:Name="tab_launcher" Header="Launcher">
                <Grid Background="#FFE5E5E5
                      ">
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
                </Grid>
            </TabItem>
            <TabItem x:Name="tab_ADtools" Header="AD commands">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <GroupBox x:Name="grpbox_ADCmds_groups" Header="Groups" HorizontalAlignment="Left" Height="67" Margin="10,393,0,0" VerticalAlignment="Top" Width="483"/>
                    <GroupBox x:Name="grpbox_ADCmds_computers" Header="Computers" HorizontalAlignment="Left" Height="124" Margin="10,264,0,0" VerticalAlignment="Top" Width="483"/>
                    <GroupBox x:Name="grpbox_ADCmds_Users" Header="Users" HorizontalAlignment="Left" Height="97" Margin="10,162,0,0" VerticalAlignment="Top" Width="483"/>
                    <TextBox x:Name="txtbox_ADCmds_username" Height="26" Margin="113,15,0,0" TextWrapping="Wrap" VerticalAlignment="Top" HorizontalAlignment="Left" Width="203"/>
                    <Label x:Name="label_Username" Content="Username:" HorizontalAlignment="Left" Margin="42,15,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="btn_unlock" Content="Unlock account" HorizontalAlignment="Left" Margin="28,188,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Button x:Name="btn_resetPWD" Content="Password Reset" HorizontalAlignment="Left" Margin="135,188,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <TextBox x:Name="txtbox_ADCmds_computername" HorizontalAlignment="Left" Height="26" Margin="113,41,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="203"/>
                    <Label x:Name="label_computername" Content="Computer name:" HorizontalAlignment="Left" Margin="10,41,0,0" VerticalAlignment="Top" Height="26" Width="98"/>
                    <Button x:Name="btn_exportGroups" Content="Export Groups" HorizontalAlignment="Left" Margin="240,188,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <ComboBox x:Name="combobox_ADCmds_site" HorizontalAlignment="Left" Margin="392,15,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Label x:Name="label_Site" Content="Site:" HorizontalAlignment="Left" Margin="360,15,0,0" VerticalAlignment="Top" Height="26" RenderTransformOrigin="0.911,0.674"/>
                    <Button x:Name="btn_computeracc" Content="Create Comp Acct" HorizontalAlignment="Left" Margin="36,349,0,0" VerticalAlignment="Top" Width="117" Height="26" IsEnabled="False"/>
                    <TextBox x:Name="txtbox_ADCmds_group" HorizontalAlignment="Left" Height="26" Margin="113,67,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="203"/>
                    <Label x:Name="label_group" Content="Group name:" HorizontalAlignment="Left" Margin="28,67,0,0" VerticalAlignment="Top" Height="26" Width="80"/>
                    <TextBox x:Name="txtbox_ADCmds_dirpath" HorizontalAlignment="Left" Height="26" Margin="113,93,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="203" IsEnabled="False"/>
                    <Label x:Name="label_dirpath" Content="Directory Path:" HorizontalAlignment="Left" Margin="19,93,0,0" VerticalAlignment="Top" Height="26" Width="89"/>

                    <Button x:Name="btn_exportMembers" Content="Export Members" HorizontalAlignment="Left" Margin="42,419,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Button x:Name="btn_loggedOn" Content="Who is logged on" HorizontalAlignment="Left" Margin="36,302,0,0" VerticalAlignment="Top" Width="117" Height="26"/>
                    <Button x:Name="btn_lastBoot" Content="Last Boot" HorizontalAlignment="Left" Margin="158,302,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <ComboBox x:Name="combobox_Type" HorizontalAlignment="Left" Margin="392,41,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Label x:Name="label_kittype" Content="Kit Type:" HorizontalAlignment="Left" Margin="337,41,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="combobox_GrpType" HorizontalAlignment="Left" Margin="392,67,0,0" VerticalAlignment="Top" Width="100" Height="26" IsEnabled="False"/>
                    <Label x:Name="label_Grptype" Content="Group Type:" HorizontalAlignment="Left" Margin="317,67,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="btn_discomputeracc" Content="Disable Comp Acc" HorizontalAlignment="Left" Margin="263,302,0,0" VerticalAlignment="Top" Width="105" Height="26"/>
                    <Button x:Name="btn_encomputeracc" Content="Enble Comp Acc" HorizontalAlignment="Left" Margin="373,302,0,0" VerticalAlignment="Top" Width="105" Height="26"/>
                    <Button x:Name="btn_addlocalAdmin" Content="Add Local Admin" HorizontalAlignment="Left" Margin="28,219,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Button x:Name="btn_remlocalAdmin" Content="Rmv Loc Admin" HorizontalAlignment="Left" Margin="135,219,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Button x:Name="btn_createGroup" Content="Create Group" HorizontalAlignment="Left" Margin="149,419,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Button x:Name="btn_addUsertoGroup" Content="Add to Group" HorizontalAlignment="Left" Margin="240,219,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Button x:Name="btn_rmvUserfromGroup" Content="Remove from Grp" HorizontalAlignment="Left" Margin="345,219,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Button x:Name="btn_addGrpAuthor" Content="Add as Author" HorizontalAlignment="Left" Margin="255,419,0,0" VerticalAlignment="Top" Width="100" Height="26"/>
                    <Button x:Name="btn_addread" Content="Add as Read Only" HorizontalAlignment="Left" Margin="360,419,0,0" VerticalAlignment="Top" Width="100" Height="26"/>

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
                        <CheckBox x:Name="chkbox_ADNU_omagh" Content="OMA Vault" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="80"/>
                    </GroupBox>
                    <CheckBox x:Name="chkbox_ADNU_dungannon" Content="DGN VAult" HorizontalAlignment="Left" Margin="30,310,0,0" VerticalAlignment="Top" Width="80"/>
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

function Load-CSVData{

$Global:AllDefaultAtts = Import-Csv "C:\scripts\IT_Admin_Tool\DefaultAttsCSV.csv"
$Global:LauncherApps = Import-Csv "C:\scripts\IT_Admin_Tool\LauncherApplicationsCSV.csv"

}

function Load-LauncherButtons{

$Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "A1"}
$WPFbtn_A1.Content = $Appdetails.AppName
$WPFimg_A1.Source = $Appdetails.IconPath
$Global:cmd_A1 = $Appdetails.AppPathCommand
$Global:prerun_A1 = $Appdetails.PreRunCmd
If ($Appdetails.AppName -eq ""){
    $WPFbtn_A1.Visibility = "Hidden"
    $WPFimg_A1.Visibility = "Hidden"
}

$Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "A2"}
$WPFbtn_A2.Content = $Appdetails.AppName
$WPFimg_A2.Source = $Appdetails.IconPath
$Global:cmd_A2 = $Appdetails.AppPathCommand
$Global:prerun_A2 = $Appdetails.PreRunCmd
If ($Appdetails.AppName -eq ""){
    $WPFbtn_A2.Visibility = "Hidden"
    $WPFimg_A2.Visibility = "Hidden"
}

$Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "A3"}
$WPFbtn_A3.Content = $Appdetails.AppName
$WPFimg_A3.Source = $Appdetails.IconPath
$Global:cmd_A3 = $Appdetails.AppPathCommand
$Global:prerun_A3 = $Appdetails.PreRunCmd
If ($Appdetails.AppName -eq ""){
    $WPFbtn_A3.Visibility = "Hidden"
    $WPFimg_A3.Visibility = "Hidden"
}

$Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "A4"}
$WPFbtn_A4.Content = $Appdetails.AppName
$WPFimg_A4.Source = $Appdetails.IconPath
$Global:cmd_A4 = $Appdetails.AppPathCommand
$Global:prerun_A4 = $Appdetails.PreRunCmd
If ($Appdetails.AppName -eq ""){
    $WPFbtn_A4.Visibility = "Hidden"
    $WPFimg_A4.Visibility = "Hidden"
}

$Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "A5"}
$WPFbtn_A5.Content = $Appdetails.AppName
$WPFimg_A5.Source = $Appdetails.IconPath
$Global:cmd_A5 = $Appdetails.AppPathCommand
$Global:prerun_A5 = $Appdetails.PreRunCmd
If ($Appdetails.AppName -eq ""){
    $WPFbtn_A5.Visibility = "Hidden"
    $WPFimg_A5.Visibility = "Hidden"
}

$Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "B1"}
$WPFbtn_B1.Content = $Appdetails.AppName
$WPFimg_B1.Source = $Appdetails.IconPath
$Global:cmd_B1 = $Appdetails.AppPathCommand
$Global:prerun_B1 = $Appdetails.PreRunCmd
If ($Appdetails.AppName -eq ""){
    $WPFbtn_B1.Visibility = "Hidden"
    $WPFimg_B1.Visibility = "Hidden"
}

$Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "B2"}
$WPFbtn_B2.Content = $Appdetails.AppName
$WPFimg_B2.Source = $Appdetails.IconPath
$Global:cmd_B2 = $Appdetails.AppPathCommand
$Global:prerun_B2 = $Appdetails.PreRunCmd
If ($Appdetails.AppName -eq ""){
    $WPFbtn_B2.Visibility = "Hidden"
    $WPFimg_B2.Visibility = "Hidden"
}

$Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "B3"}
$WPFbtn_B3.Content = $Appdetails.AppName
$WPFimg_B3.Source = $Appdetails.IconPath
$Global:cmd_B3 = $Appdetails.AppPathCommand
$Global:prerun_B3 = $Appdetails.PreRunCmd
If ($Appdetails.AppName -eq ""){
    $WPFbtn_B3.Visibility = "Hidden"
    $WPFimg_B3.Visibility = "Hidden"
}

$Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "B4"}
$WPFbtn_B4.Content = $Appdetails.AppName
$WPFimg_B4.Source = $Appdetails.IconPath
$Global:cmd_B4 = $Appdetails.AppPathCommand
$Global:prerun_B4 = $Appdetails.PreRunCmd
If ($Appdetails.AppName -eq ""){
    $WPFbtn_B4.Visibility = "Hidden"
    $WPFimg_B4.Visibility = "Hidden"
}

$Appdetails = $Global:LauncherApps | Where-Object{$_.BtnID -eq "B5"}
$WPFbtn_B5.Content = $Appdetails.AppName
$WPFimg_B5.Source = $Appdetails.IconPath
$Global:cmd_B5 = $Appdetails.AppPathCommand
$Global:prerun_B5 = $Appdetails.PreRunCmd
If ($Appdetails.AppName -eq ""){
    $WPFbtn_B5.Visibility = "Hidden"
    $WPFimg_B5.Visibility = "Hidden"
}
}

function Load-ComboBoxes{

    $Global:ComboSiteCodes = @() 
    $Global:ComboSiteCodes += ""
    $Global:ComboSiteCodes += $Global:AllDefaultAtts.SiteCode

    ## AD Tools Tab
    $Global:ComboSiteCodes | ForEach-object {$WPFcombobox_ADCmds_site.AddChild($_)}
    " " , "Workstation" , "Notebook" | ForEach-object {$WPFcomboBox_type.AddChild($_)}
    " " , "-Data_" , "-Data_Dept" , "-DST_" , "-MBX-"| ForEach-object {$WPFcomboBox_Grptype.AddChild($_)}

    ## Create new AD user tab
    $Global:ComboSiteCodes | ForEach-object {$WPFcombo_ADNU_site.AddChild($_)}

    ## Export AD User Info tab
    $Global:ComboSiteCodes | ForEach-object {$WPFcombo_EADUI_site.AddChild($_)}
}

function Get-DefaultAtts{
    $Site = $WPFcombo_ADNU_Site.SelectedValue
    $Global:SiteDefaultAtts = $Global:AllDefaultAtts | Where-Object{$_.SiteCode -eq $Site}

    $Groups = $Global:SiteDefaultAtts.DefaultGroups
    $homeDirectory = $Global:SiteDefaultAtts.HomeDirectory + $logonname + "\"
    $homeDrive = $Global:SiteDefaultAtts.HomeDrive
    $company = $Global:SiteDefaultAtts.CompanyAtt
    $L_city = $Global:SiteDefaultAtts.SiteCity
    $physicalDeliveryOfficeName = $Global:SiteDefaultAtts.SiteCity
    $postalCode = $Global:SiteDefaultAtts.PostCode
    $scriptPath = $Global:SiteDefaultAtts.ScriptPath
    $st = $Global:SiteDefaultAtts.County
    $streetAddress = $Global:SiteDefaultAtts.streetAddress
    $userPrincipalName = "$logonname" + "@" + $Global:SiteDefaultAtts.CompanyName + ".com"
    $SiteOU = "OU=Users," + $Global:SiteDefaultAtts.OU
}

#=======================#
# Main Launcher Buttons #
#=======================#

# When clicked the buttons will execute any pre-run commands saved in the CSV file before executing the 

### Button A1 ###
$WPFbtn_A1.Add_Click({ 

If (($prerun_A1 -ne $null) -and ($prerun_A1 -ne ""))
    {
    Invoke-Expression $prerun_A1
    }
    Start-Process $cmd_A1
})
 
 ### Button A2 ###
$WPFbtn_A2.Add_Click({ 
If (($prerun_A2 -ne $null) -and ($prerun_A2 -ne ""))
    {
    Invoke-Expression $prerun_A2
    }
    Start-Process $cmd_A2
})


### Button A3 ###
$WPFbtn_A3.Add_Click({ 
If (($prerun_A3 -ne $null) -and ($prerun_A3 -ne ""))
    {
    Invoke-Expression $prerun_A3
    }
    Start-Process $cmd_A3
})

### Button A4 ###
$WPFbtn_A4.Add_Click({ 
If (($prerun_A4 -ne $null) -and ($prerun_A4 -ne ""))
    {
    Invoke-Expression $prerun_A4
    }
    Start-Process $cmd_A4
}) 

### Button A5 ###
$WPFbtn_A5.Add_Click({
If (($prerun_A5 -ne $null) -and ($prerun_A5 -ne ""))
    {
    Invoke-Expression $prerun_A5
    }
    Start-Process $cmd_A5
})

### Button B1 ###
$WPFbtn_B1.Add_Click({
If (($prerun_B1 -ne $null) -and ($prerun_B1 -ne ""))
    {
    Invoke-Expression $prerun_B1
    }
    Start-Process $cmd_B1
}) 


### Button B2 ###
$WPFbtn_B2.Add_Click({ 
If (($prerun_B2 -ne $null) -and ($prerun_B2 -ne ""))
    {
    Invoke-Expression $prerun_B2
    }
    Start-Process $cmd_B2
}) 

### Button B3 ###
$WPFbtn_B3.Add_Click({
If (($prerun_B3 -ne $null) -and ($prerun_B3 -ne ""))
    {
    Invoke-Expression $prerun_B3
    }
    Start-Process $cmd_B3
})

### Button B4 ###
$WPFbtn_B4.Add_Click({
If (($prerun_B4 -ne $null) -and ($prerun_B4 -ne ""))
    {
    Invoke-Expression $prerun_B4
    }
    Start-Process $cmd_B4
}) 

### Button B5 ###
$WPFbtn_B5.Add_Click({
If (($prerun_A1 -ne $null) -and ($prerun_B5 -ne ""))
    {
    Invoke-Expression $prerun_B5
    }
    Start-Process $cmd_B5
})


#======================#
# AD Tools             #
#======================#

### Unlock user account ###
$WPFbtn_unlock.Add_Click({
    $username = $WPFtxtbox_ADCmds_username.Text

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
    $WPFtxtbox_ADCmds_username.Text = $null
    Get-module | Remove-Module        
            }
})

### Password reset ###
$WPFbtn_resetPWD.Add_Click({
    $username = $WPFtxtbox_ADCmds_username.Text

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
                        $Paswd = "TempPwd2017"
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
    $WPFtxtbox_ADCmds_username.Text = $null
    Get-module | Remove-Module        
            }

})


### Export user's group membership to CSV ###
$WPFbtn_exportGroups.Add_Click({
    $Username = $WPFtxtbox_ADCmds_username.Text
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
    $WPFtxtbox_ADCmds_username.Text = $null
    Get-module | Remove-Module 
})


### Who is logged onto computer ###
$WPFbtn_loggedOn.Add_Click({
    $computername = $WPFtxtbox_ADCmds_computername.Text
    $computername = $computername.trim()
    
    If ($computername -eq "")
        {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "ERROR: Invalid computer name"
        $Messageboxbody = "Computer name cannot be blank. Please enter a valid computer name"
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
        }
        ELSE {
             #Check the machine is online.
             IF (-not (Test-Connection -comp $computername -quiet))
                {
                $ButtonType = [System.Windows.MessageBoxButton]::OK
                $MessageboxTitle = "Warning!"
                $Messageboxbody = "$computername is offline or unreachable"
                $MessageIcon = [System.Windows.MessageBoxImage]::Warning
                [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                }ELSE{
                    IF (dsquery computer -name $computername)
                        {
                        $output = wmic.exe /node:$computername computersystem get username
                        $output = $Output.replace(' ','')
                        $output = $Output.replace('UserName','')
                        $output = $Output.trim()
                        $ButtonType = [System.Windows.MessageBoxButton]::OK
                        $MessageboxTitle = "User logged onto $computername is..."
                        $Messageboxbody = "User logged onto $computername is: `n" + $output
                        $MessageIcon = [System.Windows.MessageBoxImage]::Information
                        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                        }ELSE{
                            $ButtonType = [System.Windows.MessageBoxButton]::OK
                            $MessageboxTitle = "ERROR: Invalid computer name"
                            $Messageboxbody = "$computername does not exist in AD. Please enter a valid computer name"
                            $MessageIcon = [System.Windows.MessageBoxImage]::Error
                            [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                            }
                    }
            }
    $WPFtxtbox_ADCmds_computername.Text = $null
    Get-module | Remove-Module 
})


### Last boot of computer ###
$WPFbtn_lastBoot.Add_Click({
    $computername = $WPFtxtbox_ADCmds_computername.Text
    $computername = $computername.trim()
    
    If ($computername -eq "")
        {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "ERROR: Invalid computer name"
        $Messageboxbody = "Computer name cannot be blank. Please enter a valid computer name"
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
        }ELSE{
             #Check the machine is online.
             IF (-not (Test-Connection -comp $computername -quiet))
                {
                $ButtonType = [System.Windows.MessageBoxButton]::OK
                $MessageboxTitle = "Warning!"
                $Messageboxbody = "$computername is offline or unreachable"
                $MessageIcon = [System.Windows.MessageBoxImage]::Warning
                [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                }ELSE{
                     IF (dsquery computer -name $computername)
                        {
                        $Booted = Get-WmiObject -Class Win32_OperatingSystem -Computer $Computername
                        $Output = $Booted.ConvertToDateTime($Booted.LastBootUpTime) | Out-String

                        $ButtonType = [System.Windows.MessageBoxButton]::OK
                        $MessageboxTitle = "$computername last boot"
                        $Messageboxbody = $output
                        $MessageIcon = [System.Windows.MessageBoxImage]::Information
                        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)

                        }ELSE
                            {
                            $ButtonType = [System.Windows.MessageBoxButton]::OK
                            $MessageboxTitle = "ERROR: Invalid computer name"
                            $Messageboxbody = "$computername does not exist in AD. Please enter a valid computer name"
                            $MessageIcon = [System.Windows.MessageBoxImage]::Error
                            [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                            }
                      }
            }
    $WPFtxtbox_ADCmds_computername.Text = $null
    Get-module | Remove-Module 
})



# When a selection is made in the Site combobox, enable the Type combobox. When selection is removed, disable Type combobox and btn_computeracc. Update Computer name field automatically with correct computer name prefix.
$WPFcombobox_ADCmds_site.Add_SelectionChanged({
    $Site = $WPFcombobox_ADCmds_site.SelectedValue
    $kittype = $WPFcombobox_type.SelectedValue

    If($WPFtxtbox_ADCmds_computername.Text -ne $null)
        {
        $IDno = [regex]::Matches($WPFtxtbox_ADCmds_computername.Text, "\d+(?!.*\d+)").value
        }
    If($kittype -eq "Notebook")
        {
        $kit = "NB"
        } 
        ELSE {
            If($kittype -eq "Workstation")
                {
            $Kit = "W"
                } 
            }

    If($Site -ne " ")
        {
        $WPFcombobox_type.IsEnabled = $true
        $WPFcombobox_Grptype.IsEnabled = $true
        $WPFtxtbox_ADCmds_computername.Text = $site + $kit + $IDno
        } 
    If (($Site -eq " ") -or ($Site -eq ""))
        {
        $WPFcombobox_type.IsEnabled = $false
        $WPFbtn_computeracc.IsEnabled = $false
        $WPFcombobox_type.SelectedValue = $null
        $WPFtxtbox_ADCmds_computername.Text = $null
        }
})

# When a selection is made in the Site combobox, enable btn_computeracc. When selection is removed, disable btn_computeracc. Update Computer name field automatically with correct computer name prefix.
$WPFcombobox_type.Add_SelectionChanged({
    $Site = $WPFcombobox_ADCmds_site.SelectedValue
    $kittype = $WPFcombobox_type.SelectedValue
    If($WPFtxtbox_ADCmds_computername.Text -ne $null)
        {
        $IDno = [regex]::Matches($WPFtxtbox_ADCmds_computername.Text, "\d+(?!.*\d+)").value
        }
    If($kittype -eq "Notebook")
        {
        $kit = "NB"
        $WPFbtn_computeracc.IsEnabled = $true
        } 
        ELSE {
            If($kittype -eq "Workstation")
                {
            $Kit = "W"
            $WPFbtn_computeracc.IsEnabled = $true
                } 
                ELSE {
                    If($Kittype -eq " ")
                        {
                        $WPFbtn_computeracc.IsEnabled = $false
                        }
                      }
            }
    $WPFtxtbox_ADCmds_computername.Text = $site + $kit + $IDno
})

### Create computer account ###
# When btn_computeracc is clicked, validate data and create the computer account
$WPFbtn_computeracc.Add_Click({
   $computername = $WPFtxtbox_ADCmds_computername.Text
   $Site = $WPFcombobox_ADCmds_site.SelectedValue
   $KitType = $WPFcombobox_Type.SelectedValue
   $computername = $computername.trim()
    
    If ($computername -eq "")
        {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "ERROR: Invalid computer name"
        $Messageboxbody = "Computer name cannot be blank. Please enter a valid computer name"
        $MessageIcon = [System.Windows.MessageBoxImage]::Warning
        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
        }
        ELSE {
            If (($kittype -ne "Notebook") -and ($kittype -ne "Workstation"))
                {
                $ButtonType = [System.Windows.MessageBoxButton]::OK
                $MessageboxTitle = "ERROR: Invalid kit type"
                $Messageboxbody = "Type of kit cannot be blank. Please select Notebook or Workstation"
                $MessageIcon = [System.Windows.MessageBoxImage]::Warning
                [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                }
                ELSE {
                     If ($Site -notin $Global:AllDefaultAtts.SiteCode)
                        {
                        $ButtonType = [System.Windows.MessageBoxButton]::OK
                        $MessageboxTitle = "ERROR: Invalid site"
                        $Messageboxbody = "Please select a site."
                        $MessageIcon = [System.Windows.MessageBoxImage]::Warning
                        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
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
                                $ButtonType = [System.Windows.MessageBoxButton]::OK
                                $MessageboxTitle = "Duplicate computer name"
                                $Messageboxbody = "Requested computer name already exists in AD. Please enter a unique computer name"
                                $MessageIcon = [System.Windows.MessageBoxImage]::Warning
                                [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                                } Else {
                                    New-ADComputer -Name $computername -SAMAccountName $computersam -Path $OU
                                    IF (dsquery computer -samid $computersam)
                                        {
                                        $ButtonType = [System.Windows.MessageBoxButton]::OK
                                        $MessageboxTitle = "Success"
                                        $Messageboxbody = "$computername computer account created!"
                                        $MessageIcon = [System.Windows.MessageBoxImage]::Information
                                        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                                        } Else {
                                        Write-Host "OU: "$OU
                                        $ButtonType = [System.Windows.MessageBoxButton]::OK
                                        $MessageboxTitle = "ERROR"
                                        $Messageboxbody = "An error has occured. $computername computer account not created. Aborting!"
                                        $MessageIcon = [System.Windows.MessageBoxImage]::Error
                                        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                                        }
                                    $WPFcombobox_ADCmds_site.text = ""
                                    $WPFcombobox_Type.text = ""
                                    $WPFcombobox_Type.IsEnabled = $false
                                    $WPFbtn_computeracc.IsEnabled = $false
                                    $WPFtxtbox_ADCmds_computername.Text = ""
                                    }
                            $WPFtxtbox_ADCmds_computername.Text = $null
                             }
                     }
        }
    Get-module | Remove-Module 

})


### Export the members of a group ###
$WPFbtn_exportMembers.Add_Click({
    $Group = $WPFtxtbox_ADCmds_Group.Text
    $Group = $Group.trim()
    
    If ($Group -eq "")
        {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "ERROR: Invalid group name"
        $Messageboxbody = "Group cannot be blank. Please enter a valid Group"
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
        }
        ELSE {
            IF (dsquery group -samid $Group)
                {
                Get-ADGroupMember -identity $Group | select name | Export-csv -path "C:\Members of group - $Group.csv" -NoTypeInformation
                $ButtonType = [System.Windows.MessageBoxButton]::OK
                $MessageboxTitle = "Export a group's members"
                $Messageboxbody = "Group membership of $group has been exported to C:\Members of group - $Group.csv"
                $MessageIcon = [System.Windows.MessageBoxImage]::Information
                [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                } ELSE
                    {
                    $ButtonType = [System.Windows.MessageBoxButton]::OK
                    $MessageboxTitle = "ERROR: Invalid group name"
                    $Messageboxbody = "$Group does not exist in AD. Please enter a valid group name"
                    $MessageIcon = [System.Windows.MessageBoxImage]::Error
                    [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
                    }
            }
    $WPFtxtbox_ADCmds_Group.Text = $null
    Get-module | Remove-Module 
})

#======================#
# Create new AD user   #
#======================#

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

## Validate entered data on Validate button click
$WPFbtn_ADNU_validate.Add_Click({
    # Reset any highlighted errors from a previous validation
    $errorcount = 0
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
        }

    # If the firstname is left blank, error
    If (($firstname -eq "") -or ($firstname -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errorfirst.Visibility = "Visible"
        }

    # If the lastname is left blank, error
    If (($lastname -eq "") -or ($lastname -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errorlast.Visibility = "Visible"
        }
        
    # If logon name is blank, error
    If (($logonname -eq "") -or ($logonname -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errorlogon.Visibility = "Visible"
        } ELSE {
                # Checks that the logon name is unique
                If (dsquery user -samid $logonname)
                    {
                    $errorcount++
                    $WPFtxtblock_ADNU_errorlogonunique.Visibility = "Visible"
                    }
                }

    # If display name is blank, error
    If (($displayname -eq "") -or ($displayname -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errordisp.Visibility = "Visible"
        }

    # If manager name is blank, error
    If (($manager -eq "") -or ($manager -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errormgr.Visibility = "Visible"
        } ELSE {
                # Check the inputted manager exists in AD
                If (!(dsquery user -samid $manager))
                    {
                    $errorcount++
                    $WPFtxtblock_ADNU_errormgrunique.Visibility = "Visible"
                    }
                }
    # If Job Title is blank, error
        If (($job -eq "") -or ($job -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errortitle.Visibility = "Visible"
        }
    # If Dept is blank, error.
        If (($dept -eq "") -or ($dept -eq $null))
        {
        $errorcount++
        $WPFtxtblock_ADNU_errordept.Visibility = "Visible"
        }

    ## If errors are present, highlight issues.
    If ($errorcount -ne 0)
        {
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "Check errors"
        $Messageboxbody = "There are $errorcount error(s). Please ensure to all reqired fields marked with excalamation marks ! are completed and action any other errors as displayed."
        $MessageIcon = [System.Windows.MessageBoxImage]::Warning
        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
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
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "Validate user info"
        $Messageboxbody = "User info has been validated. To make changes, click Edit. To proceed, click Create User."
        $MessageIcon = [System.Windows.MessageBoxImage]::Information
        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)
        }
    Get-module | Remove-Module  
})

## Reset the form
$WPFbtn_ADNU_resetform.Add_Click({
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
    $WPFcombo_ADNU_Site.SelectedValue = $null
    $WPFtxtbox_ADNU_firstname.Text = $null
    $WPFtxtbox_ADNU_lastname.Text = $null
    $WPFtxtbox_ADNU_initial.Text = $null
    $WPFchkbox_ADNU_Site1.IsChecked = $false
    $WPFchkbox_ADNU_Site2.IsChecked = $false
    $WPFchkbox_ADNU_temp.IsChecked = $false
    $WPFtxtbox_ADNU_logonname.Text = $null
    $WPFtxtbox_ADNU_displayname.Text = $null
    $WPFtxtbox_ADNU_manager.Text = $null
    $WPFtxtbox_ADNU_title.Text = $null
    $WPFtxtbox_ADNU_dept.Text = $null
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
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "Complete"
        $Messageboxbody = "All tasks have been completed successfully!"
        $MessageIcon = [System.Windows.MessageBoxImage]::Warning
        [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)  

        # Reset form for next use
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
        $WPFtxtblock_ADNU_errorexpiry.Visibility = "hidden"
        $WPFtxtblock_ADNU_errorsite.Visibility = "hidden"
        $WPFtxtblock_ADNU_errorfirst.Visibility = "hidden"
        $WPFtxtblock_ADNU_errorlast.Visibility = "hidden"
        $WPFtxtblock_ADNU_errorlogonunique.Visibility = "hidden"
        $WPFtxtblock_ADNU_errormgrunique.Visibility = "hidden"
        $WPFtxtblock_ADNU_errorlogon.Visibility = "hidden"
        $WPFtxtblock_ADNU_errordisp.Visibility = "hidden"
        $WPFtxtblock_ADNU_errormgr.Visibility = "hidden"
        $WPFcombo_ADNU_Site.SelectedValue = $null
        $WPFtxtbox_ADNU_firstname.Text = $null
        $WPFtxtbox_ADNU_lastname.Text = $null
        $WPFtxtbox_ADNU_initial.Text = $null
        $WPFchkbox_ADNU_Site1.IsChecked = $false
        $WPFchkbox_ADNU_Site2.IsChecked = $false
        $WPFchkbox_ADNU_temp.IsChecked = $false
        $WPFtxtbox_ADNU_logonname.Text = $null
        $WPFtxtbox_ADNU_displayname.Text = $null
        $WPFtxtbox_ADNU_manager.Text = $null
        $WPFtxtbox_ADNU_title.Text = $null
        $WPFtxtbox_ADNU_dept.Text = $null          
        } ELSE 
              {
               # An error has occured.
               $ButtonType = [System.Windows.MessageBoxButton]::OK
               $MessageboxTitle = "ERROR"
               $Messageboxbody = "Failed to create user account. Aborting job!"
               $MessageIcon = [System.Windows.MessageBoxImage]::Error
               [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)  
               }

    Get-module | Remove-Module  
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


})


## If All Attributes Check box is checked, common and additional atts are unchecked and all attributes are disabled and checked
$WPFchkbox_EADUI_AllAtts.Add_Checked({

    $WPFchkbox_EADUI_AdditionalAtts.IsChecked = $false
    $WPFchkbox_EADUI_common.IsChecked = $false
    $WPFchkbox_EADUI_common.IsEnabled = $false
    $WPFchkbox_EADUI_AdditionalAtts.IsEnabled = $false

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

    #Check the additional atts
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
})


## If Common Attributes Check box is checked, common atts are checked
$WPFchkbox_EADUI_Common.Add_Checked({

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

})


## If Common Attributes Check box is unchecked, common atts are unchecked
$WPFchkbox_EADUI_Common.Add_UnChecked({

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

})


## If Additional Attributes Check box is checked, additional atts are checked
$WPFchkbox_EADUI_AdditionalAtts.Add_Checked({

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
       
})
 

## If Additional Attributes Check box is unchecked, additional atts are unchecked
$WPFchkbox_EADUI_AdditionalAtts.Add_UnChecked({

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

        If (($username -eq $null) -or ($username -eq ""))
            {
            $errorcount++
            $WPFtxtblock_EADUI_errorusername.Visibility = "Visible"
            $WPFtxtbox_EADUI_username.Text = $username
            } ELSE {
                    If (!(dsquery user -samid $username))
                        {
                        $errorcount++
                        $WPFtxtblock_EADUI_erroruserunique.Visibility = "Visible"
                        $WPFtxtbox_EADUI_username.Text = $username
                        }
                     }
        }

    If ($WPFradio_EADUI_OU.IsChecked)
        {
        If (($site -eq $null) -or ($site -eq ""))
            {
            $errorcount++
            $WPFtxtblock_EADUI_errorsite.Visibility = "Visible"
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

            $ButtonType = [System.Windows.MessageBoxButton]::OK
            $MessageboxTitle = "Success"
            $Messageboxbody = "User details exported to: $exportdir"
            $MessageIcon = [System.Windows.MessageBoxImage]::Information
            [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)  
            $WPFtxtbox_EADUI_username.text = $null
            }

        If ($WPFradio_EADUI_User.IsChecked)
            {
            $exportdir = "C:\" + "$Username" + "_export.csv"

            Get-ADUser -Identity $username -Properties $attributes | Export-csv $exportdir

            $ButtonType = [System.Windows.MessageBoxButton]::OK
            $MessageboxTitle = "Success"
            $Messageboxbody = "User details exported to: $exportdir"
            $MessageIcon = [System.Windows.MessageBoxImage]::Information
            [System.Windows.Messagebox]::Show($Messageboxbody,$MessageboxTitle,$ButtonType,$MessageIcon)  
            $WPFcombo_EADUI_site.SelectedValue = $null
                        }
        }
    Get-module | Remove-Module 
})



#===========================================================================#
# Imports settings and shows the GUI                                        #
#===========================================================================#
Load-CSVData
Load-LauncherButtons
Load-ComboBoxes
CLS
Write-Host -ForegroundColor Cyan "IT Admin Tool $Version"
Write-Host -ForegroundColor Red "Do NOT close this window. Please use the GUI to interact with the IT Admin Tool."
$Form.ShowDialog() | out-null
