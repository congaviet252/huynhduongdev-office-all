# You need to have Administrator rights to run this script!
    if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "Bạn cần quyền Administrator để chạy script này!`nVui lòng chạy lại script dưới quyền Admin (Run as Administrator)!"
        Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "iex (irm https://raw.githubusercontent.com/congaviet252/huynhduongdev-office-all/refs/heads/main/office.ps1)"
        break
    }

# Load ddls to the current session.
    Add-Type -AssemblyName PresentationFramework, System.Drawing, PresentationFramework, System.Windows.Forms, WindowsFormsIntegration, PresentationCore
    [System.Windows.Forms.Application]::EnableVisualStyles()

# GIAO DIỆN XAML ĐƯỢC NÂNG CẤP 3D / MATERIAL DESIGN
$xamlInput = @'
<Window x:Class="install.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Huỳnh Dương Developer - Ultimate Office Tool" 
        WindowStartupLocation="CenterScreen" 
        ResizeMode="CanMinimize"
        Width="1200" Height="700"
        Background="#1E1E2E"
        Icon="https://raw.githubusercontent.com/mseduvn/msoffice/refs/heads/main/Files/images.png">
    
    <Window.Resources>
        <!-- COLORS & BRUSHES -->
        <LinearGradientBrush x:Key="MainBackground" StartPoint="0,0" EndPoint="1,1">
            <GradientStop Color="#1a1c2c" Offset="0.0"/>
            <GradientStop Color="#4a192c" Offset="1.0"/>
        </LinearGradientBrush>

        <LinearGradientBrush x:Key="CardBackground" StartPoint="0,0" EndPoint="0,1">
            <GradientStop Color="#2C2F48" Offset="0"/>
            <GradientStop Color="#23263A" Offset="1"/>
        </LinearGradientBrush>

        <LinearGradientBrush x:Key="ButtonGradient" StartPoint="0,0" EndPoint="1,0">
            <GradientStop Color="#00C9FF" Offset="0"/>
            <GradientStop Color="#92FE9D" Offset="1"/>
        </LinearGradientBrush>

        <SolidColorBrush x:Key="TextPrimary" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#B0B5C1"/>
        <SolidColorBrush x:Key="AccentColor" Color="#00C9FF"/>

        <!-- STYLES -->
        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="{StaticResource AccentColor}"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="BorderBrush" Value="#444960"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GroupBox">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Border Grid.Row="0" Grid.RowSpan="2" CornerRadius="10" Background="#25283D" BorderBrush="#444960" BorderThickness="1">
                                <Border.Effect>
                                    <DropShadowEffect Color="Black" BlurRadius="10" ShadowDepth="3" Opacity="0.4"/>
                                </Border.Effect>
                            </Border>
                            <Border Grid.Row="0" Background="Transparent" Padding="10,5,0,0">
                                <ContentPresenter ContentSource="Header" RecognizesAccessKey="True"/>
                            </Border>
                            <Border Grid.Row="1" Padding="10">
                                <ContentPresenter/>
                            </Border>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Margin" Value="0,5,0,5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsChecked" Value="True">
                    <Setter Property="Foreground" Value="White"/>
                    <Setter Property="FontWeight" Value="Bold"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="Button" x:Key="ModernButton">
            <Setter Property="Foreground" Value="#1a1c2c"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" CornerRadius="20" Background="{StaticResource ButtonGradient}" BorderThickness="0">
                            <Border.Effect>
                                <DropShadowEffect Color="#00C9FF" BlurRadius="15" ShadowDepth="0" Opacity="0.5"/>
                            </Border.Effect>
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.9"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="RenderTransform">
                                    <Setter.Value>
                                        <ScaleTransform ScaleX="0.95" ScaleY="0.95"/>
                                    </Setter.Value>
                                </Setter>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

         <Style TargetType="Button" x:Key="DangerButton">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" CornerRadius="8" Background="#FF4B4B" BorderThickness="0">
                             <Border.Effect>
                                <DropShadowEffect Color="Red" BlurRadius="10" ShadowDepth="2" Opacity="0.3"/>
                            </Border.Effect>
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                         <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#FF6B6B"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Custom Tab Item Style -->
        <Style TargetType="TabItem">
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="Border" Padding="15,10" Margin="0,0,5,0" CornerRadius="5,5,0,0" Background="#2C2F48">
                            <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" HorizontalAlignment="Center" ContentSource="Header" Margin="10,2"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="Background" Value="{StaticResource ButtonGradient}"/>
                                <Setter Property="Foreground" Value="#1a1c2c"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="False">
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <!-- MAIN LAYOUT -->
    <Grid Background="{StaticResource MainBackground}">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Header -->
            <RowDefinition Height="*"/>    <!-- Content -->
            <RowDefinition Height="Auto"/> <!-- Footer -->
        </Grid.RowDefinitions>

        <!-- HEADER -->
        <Border Grid.Row="0" Padding="20" Background="#151725">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                 <!-- Logo Placeholder -->
                <Border Width="40" Height="40" CornerRadius="10" Background="{StaticResource ButtonGradient}" Margin="0,0,15,0">
                    <TextBlock Text="HD" HorizontalAlignment="Center" VerticalAlignment="Center" FontWeight="ExtraBold" FontSize="20" Foreground="#1a1c2c"/>
                </Border>
                <StackPanel>
                    <TextBlock Text="HUỲNH DƯƠNG DEVELOPER PRO" Foreground="White" FontSize="24" FontWeight="Bold" FontFamily="Segoe UI">
                        <TextBlock.Effect>
                            <DropShadowEffect Color="#00C9FF" BlurRadius="10" ShadowDepth="0" Opacity="0.6"/>
                        </TextBlock.Effect>
                    </TextBlock>
                    <TextBlock Text="Ultimate Office Deployment Tool - Phiên bản Việt Hóa" Foreground="#B0B5C1" FontSize="12"/>
                </StackPanel>
            </StackPanel>
        </Border>

        <!-- BODY CONTENT -->
        <Grid Grid.Row="1" Margin="20">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="280"/> <!-- Left Sidebar (Settings) -->
                <ColumnDefinition Width="*"/>   <!-- Right Main (Office Select) -->
            </Grid.ColumnDefinitions>

            <!-- LEFT SIDEBAR: Settings -->
            <StackPanel Grid.Column="0" Margin="0,0,20,0">
                <!-- Architecture -->
                <GroupBox x:Name="groupBoxArch" Header="Kiến Trúc (Bit)">
                    <StackPanel>
                        <RadioButton x:Name="radioButtonArch64" Content="x64 (64-bit) - Khuyên dùng" IsChecked="True"/>
                        <RadioButton x:Name="radioButtonArch32" Content="x86 (32-bit)"/>
                    </StackPanel>
                </GroupBox>

                <!-- License -->
                <GroupBox x:Name="groupBoxLicenseType" Header="Loại Giấy Phép">
                    <StackPanel>
                        <RadioButton x:Name="radioButtonVolume" Content="Volume (Doanh nghiệp)" IsChecked="True"/>
                        <RadioButton x:Name="radioButtonRetail" Content="Retail (Bán lẻ)"/>
                    </StackPanel>
                </GroupBox>

                <!-- Mode -->
                <GroupBox x:Name="groupBoxMode" Header="Chế Độ Hoạt Động">
                    <StackPanel>
                        <RadioButton x:Name="radioButtonInstall" Content="Cài Đặt (Install)" IsChecked="True"/>
                        <RadioButton x:Name="radioButtonDownload" Content="Chỉ Tải Về (Download)"/>
                        <TextBlock Text="(*) File tải về sẽ nằm ở Desktop" Foreground="#666" FontSize="10" Margin="20,2,0,0"/>
                    </StackPanel>
                </GroupBox>
                
                <!-- Language -->
                <GroupBox x:Name="groupBoxLanguage" Header="Ngôn Ngữ" Height="200">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel>
                            <RadioButton x:Name="radioButtonEnglish" Content="Tiếng Anh (English)" IsChecked="True"/>
                            <RadioButton x:Name="radioButtonVietnamese" Content="Tiếng Việt (Vietnamese)" Foreground="#00C9FF" FontWeight="Bold"/>
                            <RadioButton x:Name="radioButtonJapanese" Content="Tiếng Nhật"/>
                            <RadioButton x:Name="radioButtonKorean" Content="Tiếng Hàn"/>
                            <RadioButton x:Name="radioButtonChinese" Content="Tiếng Trung"/>
                            <RadioButton x:Name="radioButtonFrench" Content="Tiếng Pháp"/>
                            <RadioButton x:Name="radioButtonSpanish" Content="Tây Ban Nha"/>
                            <RadioButton x:Name="radioButtonGerman" Content="Tiếng Đức"/>
                            <RadioButton x:Name="radioButtonHindi" Content="Tiếng Hindi"/>
                        </StackPanel>
                    </ScrollViewer>
                </GroupBox>
            </StackPanel>

            <!-- RIGHT MAIN: Office Versions (Tabs) -->
            <Border Grid.Column="1" Background="#23263A" CornerRadius="10" Padding="10">
                <Border.Effect>
                    <DropShadowEffect Color="Black" BlurRadius="20" ShadowDepth="5" Opacity="0.3"/>
                </Border.Effect>
                
                <TabControl Background="Transparent" BorderThickness="0">
                    
                    <!-- TAB 1: NEWEST (2024 - 2021 - 365) -->
                    <TabItem Header="Mới Nhất (2021 - 2024 - 365)">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel Margin="10">
                                <!-- Office 2024 -->
                                <GroupBox x:Name="groupBoxMicrosoftOffice" Header="Office 2024 LTSC (Mới nhất)" BorderBrush="#FFE2820E">
                                    <WrapPanel>
                                        <RadioButton x:Name="radioButton2024Pro" Content="Professional Plus" Width="140"/>
                                        <RadioButton x:Name="radioButton2024Std" Content="Standard" Width="140"/>
                                        <RadioButton x:Name="radioButton2024ProjectPro" Content="Project Pro" Width="140"/>
                                        <RadioButton x:Name="radioButton2024VisioPro" Content="Visio Pro" Width="140"/>
                                        <RadioButton x:Name="radioButton2024Word" Content="Word Only" Width="100"/>
                                        <RadioButton x:Name="radioButton2024Excel" Content="Excel Only" Width="100"/>
                                        <RadioButton x:Name="radioButton2024PowerPoint" Content="PowerPoint" Width="100"/>
                                        <RadioButton x:Name="radioButton2024ProjectStd" Content="Project Std" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2024VisioStd" Content="Visio Std" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2024Outlook" Content="Outlook" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2024Access" Content="Access" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2024Publisher" Content="Publisher" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2024HomeStudent" Content="HomeStudent" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2024HomeBusiness" Content="HomeBusiness" Visibility="Collapsed"/>
                                    </WrapPanel>
                                </GroupBox>

                                <!-- Office 2021 -->
                                <GroupBox Header="Office 2021 LTSC" BorderBrush="#FF3C10DE">
                                    <WrapPanel>
                                        <RadioButton x:Name="radioButton2021Pro" Content="Professional Plus" Width="140"/>
                                        <RadioButton x:Name="radioButton2021Std" Content="Standard" Width="140"/>
                                        <RadioButton x:Name="radioButton2021ProjectPro" Content="Project Pro" Width="140"/>
                                        <RadioButton x:Name="radioButton2021VisioPro" Content="Visio Pro" Width="140"/>
                                        <RadioButton x:Name="radioButton2021Word" Content="Word Only" Width="100"/>
                                        <RadioButton x:Name="radioButton2021Excel" Content="Excel Only" Width="100"/>
                                        <RadioButton x:Name="radioButton2021PowerPoint" Content="PowerPoint" Width="100"/>
                                        <!-- Hidden logic mapping -->
                                        <RadioButton x:Name="radioButton2021ProjectStd" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2021VisioStd" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2021Outlook" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2021Access" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2021Publisher" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2021HomeStudent" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2021HomeBusiness" Visibility="Collapsed"/>
                                    </WrapPanel>
                                </GroupBox>

                                <!-- Office 365 -->
                                <GroupBox Header="Microsoft 365 (O365)" BorderBrush="#FFDA2323">
                                    <StackPanel Orientation="Horizontal">
                                        <RadioButton x:Name="radioButton365Home" Content="Home (Cá nhân)" Margin="0,0,15,0"/>
                                        <RadioButton x:Name="radioButton365Business" Content="Business (Kinh doanh)" Margin="0,0,15,0"/>
                                        <RadioButton x:Name="radioButton365Enterprise" Content="Enterprise (Doanh nghiệp)"/>
                                    </StackPanel>
                                </GroupBox>
                            </StackPanel>
                        </ScrollViewer>
                    </TabItem>

                    <!-- TAB 2: OLDER VERSIONS -->
                    <TabItem Header="Phiên Bản Cũ (2019 - 2013)">
                        <ScrollViewer>
                            <StackPanel Margin="10">
                                <!-- 2019 -->
                                <GroupBox Header="Office 2019" BorderBrush="#FF0F8E40">
                                    <WrapPanel>
                                        <RadioButton x:Name="radioButton2019Pro" Content="Professional Plus" Width="140"/>
                                        <RadioButton x:Name="radioButton2019Std" Content="Standard" Width="140"/>
                                        <RadioButton x:Name="radioButton2019ProjectPro" Content="Project Pro" Width="140"/>
                                        <RadioButton x:Name="radioButton2019VisioPro" Content="Visio Pro" Width="140"/>
                                        <RadioButton x:Name="radioButton2019Word" Content="Word" Width="80"/>
                                        <RadioButton x:Name="radioButton2019Excel" Content="Excel" Width="80"/>
                                        <!-- Hidden logic mapping -->
                                        <RadioButton x:Name="radioButton2019ProjectStd" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2019VisioStd" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2019PowerPoint" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2019Outlook" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2019Access" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2019Publisher" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2019HomeStudent" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2019HomeBusiness" Visibility="Collapsed"/>
                                    </WrapPanel>
                                </GroupBox>

                                <!-- 2016 -->
                                <GroupBox Header="Office 2016" BorderBrush="#FFA28210">
                                    <WrapPanel>
                                        <RadioButton x:Name="radioButton2016Pro" Content="Professional Plus" Width="140"/>
                                        <RadioButton x:Name="radioButton2016Std" Content="Standard" Width="140"/>
                                        <RadioButton x:Name="radioButton2016VisioPro" Content="Visio Pro" Width="140"/>
                                        <!-- Hidden logic mapping -->
                                        <RadioButton x:Name="radioButton2016ProjectPro" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2016ProjectStd" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2016VisioStd" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2016Word" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2016Excel" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2016PowerPoint" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2016Outlook" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2016Access" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2016Publisher" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2016OneNote" Visibility="Collapsed"/>
                                    </WrapPanel>
                                </GroupBox>

                                <!-- 2013 -->
                                <GroupBox Header="Office 2013" BorderBrush="#FF1B0F0F">
                                    <WrapPanel>
                                        <RadioButton x:Name="radioButton2013Pro" Content="Professional Plus" Width="140"/>
                                        <RadioButton x:Name="radioButton2013Std" Content="Standard" Width="140"/>
                                        <!-- Hidden logic mapping -->
                                        <RadioButton x:Name="radioButton2013ProjectPro" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2013ProjectStd" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2013VisioPro" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2013VisioStd" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2013Word" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2013Excel" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2013PowerPoint" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2013Outlook" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2013Access" Visibility="Collapsed"/>
                                        <RadioButton x:Name="radioButton2013Publisher" Visibility="Collapsed"/>
                                    </WrapPanel>
                                </GroupBox>
                            </StackPanel>
                        </ScrollViewer>
                    </TabItem>

                    <!-- TAB 3: TOOLS & UNINSTALL -->
                    <TabItem Header="Công Cụ &amp; Gỡ Cài Đặt">
                        <StackPanel Margin="20">
                            <GroupBox Header="Xóa Bỏ Hoàn Toàn Office (Clean Uninstall)" BorderBrush="Red">
                                <StackPanel>
                                    <TextBlock Text="(*) Chức năng này sẽ quét sạch mọi phiên bản Office trên máy." Foreground="#FF6B6B" Margin="0,0,0,10" FontStyle="Italic"/>
                                    
                                    <CheckBox x:Name="radioButtonRemoveAllApp" Content="Tôi hiểu rủi ro và đồng ý xóa" Foreground="White" Margin="0,0,0,10"/>
                                    
                                    <Button x:Name="buttonRemoveAll" Content="GỠ CÀI ĐẶT NGAY" Style="{StaticResource DangerButton}" Width="150" Height="40" HorizontalAlignment="Left"/>
                                    
                                    <!-- Elements required by backend logic but kept hidden or structured -->
                                    <TextBlock x:Name="textBoxRemoveAll" Visibility="Collapsed"/>
                                    <Label x:Name="LabelRemoveAll" Visibility="Collapsed"/>
                                    <Rectangle x:Name="RemoveAll" Visibility="Collapsed"/>
                                </StackPanel>
                            </GroupBox>
                            
                             <!-- Info hidden -->
                            <TextBox x:Name="textBox1" Visibility="Collapsed"/>
                            <TextBox x:Name="textBox2" Visibility="Collapsed"/>
                            <TextBox x:Name="textBox3" Visibility="Collapsed"/>
                            <TextBox x:Name="textBox5" Visibility="Collapsed"/>
                        </StackPanel>
                    </TabItem>
                </TabControl>
            </Border>
        </Grid>

        <!-- FOOTER & STATUS -->
        <Border Grid.Row="2" Background="#151725" Padding="20">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <TextBlock Text="Trạng thái xử lý:" Foreground="#B0B5C1" FontSize="11" Margin="0,0,0,5"/>
                    <Grid>
                        <ProgressBar x:Name="progressbar" Height="15" Background="#2C2F48" BorderThickness="0" Foreground="{StaticResource AccentColor}"/>
                        <TextBox x:Name="textbox" Background="Transparent" BorderThickness="0" Foreground="White" 
                                 FontWeight="Bold" TextAlignment="Center" VerticalAlignment="Center" 
                                 Text="Sẵn sàng" IsReadOnly="True" IsHitTestVisible="False"/>
                    </Grid>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" Margin="20,0,0,0">
                    <Image x:Name="image" Width="30" Height="30" Margin="0,0,15,0" Source="https://raw.githubusercontent.com/mseduvn/msoffice/refs/heads/main/Files/images.png" Visibility="Hidden"/>
                    <Button x:Name="buttonSubmit" Content="BẮT ĐẦU" Style="{StaticResource ModernButton}" Width="150" Height="45"/>
                    <Label x:Name="Link1" Visibility="Collapsed"/> <!-- Hidden logic link -->
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

# Store form objects (variables) in PowerShell

    [xml]$xaml = $xamlInput -replace '^<Window.*', '<Window' -replace 'mc:Ignorable="d"','' -replace "x:Name",'Name'
    $xmlReader = (New-Object System.Xml.XmlNodeReader $xaml)
    $Form = [Windows.Markup.XamlReader]::Load( $xmlReader)

    $xaml.SelectNodes("//*[@Name]") | ForEach-Object -Process {
        Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)
    }

    # Custom logic for Hyperlink since we hid the label
    # We can add a click event to the logo area or similar if needed, 
    # but strictly keeping original logic requires the variable $Link1 to exist.
    if ($Link1) {
        $Link1.Add_PreviewMouseDown({[system.Diagnostics.Process]::start('https://msedu.vn')})
    }

# Download links
    $uri            = "https://github.com/mseduvn/msoffice/raw/refs/heads/main/Files/setup.exe"
    $uri2013        = "https://github.com/mseduvn/msoffice/raw/refs/heads/main/Files/bin2013.exe"
    $uri2016        = "https://github.com/mseduvn/msoffice/raw/refs/heads/main/Files/setup.exe"
    $uninstall      = "https://github.com/mseduvn/msoffice/raw/refs/heads/main/Files/setup.exe"
    $removeAllXML   = 'https://raw.githubusercontent.com/mseduvn/msoffice/refs/heads/main/Files/RemoveAll/configuration.xml'

# Prepiaration for download and install
    function PreparingOffice {
        if ($radioButtonDownload.IsChecked) {
            $workingDir = New-Item -Path $env:userprofile\Desktop\$productName -ItemType Directory -Force
            Set-Location $workingDir
            Invoke-Item $workingDir
        }

        if ($radioButtonInstall.IsChecked) {
            $workingDir = New-Item -Path $env:temp\ClickToRun\$productId -ItemType Directory -Force
            Set-Location $workingDir
        }

        $configurationFile = "configuration-x$arch.xml"
        New-Item $configurationFile -ItemType File -Force
        Add-Content $configurationFile -Value "<Configuration>"
        Add-content $configurationFile -Value "<Add OfficeClientEdition=`"$arch`">"
        Add-content $configurationFile -Value "<Product ID=`"$productId`">"
        Add-content $configurationFile -Value "<Language ID=`"$languageId`"/>"
        Add-Content $configurationFile -Value "</Product>"
        Add-Content $configurationFile -Value "</Add>"
        Add-Content $configurationFile -Value "</Configuration>"

        $batchFile = "Install-$($arch)bit.bat"
        New-Item $batchFile -ItemType File -Force
        Add-content $batchFile -Value "@echo off"
        Add-content $batchFile -Value "cd /d %~dp0"
        Add-content $batchFile -Value "bin.exe /configure $configurationFile"

        (New-Object Net.WebClient).DownloadFile($uri, "$workingDir\bin.exe")

        $sync.configurationFile = $configurationFile
        $sync.workingDir = $workingDir
    }
    
# Creating script block for download and install
    $DownloadInstallOffice = {

        # To referece our elements we use the $sync variable from hashtable.
            $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = "Hidden" })
            $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "$($sync.UIstatus) $($sync.productName) $($sync.arch)-bit ($($sync.language))" })
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.BorderBrush = "#FF707070" })
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })
            $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Visible" })

        Set-Location -Path $($sync.workingDir)

        Start-Process -FilePath .\bin.exe -ArgumentList "$($sync.mode) .\$($sync.configurationFile)" -NoNewWindow -Wait
                
        # Bring back our Button, set the Label and ProgressBar, we're done..
            $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Hidden" })
            $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = 'Visible' })
            $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Content = 'BẮT ĐẦU' })
            $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = 'Hoàn tất thành công!' })
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $false })
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.Value = '100' })
    }

# Share info between runspaces
    $sync = [hashtable]::Synchronized(@{})
    $sync.runspace = $runspace
    $sync.host = $host
    $sync.Form = $Form
    $sync.ProgressBar = $ProgressBar
    $sync.textbox = $textbox
    $sync.image = $image
    $sync.buttonSubmit = $buttonSubmit
    $sync.DebugPreference = $DebugPreference
    $sync.VerbosePreference = $VerbosePreference

# Build a runspace
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()

# Add shared data to the runspace
    $runspace.SessionStateProxy.SetVariable("sync", $sync)

# Create a Powershell instance
    $PSIinstance = [powershell]::Create().AddScript($scriptBlock)
    $PSIinstance.Runspace = $runspace


# Download and install Microsoft Office with a runspace

    $buttonSubmit.Add_Click( {

            $i = 0
            if ($radioButtonArch32.IsChecked) {$arch = '32'}
            if ($radioButtonArch64.IsChecked) {$arch = '64'}

            if ($radioButtonVolume.IsChecked) {$licType = 'Volume'}
            if ($radioButtonRetail.IsChecked) {$licType = 'Retail'}

            if ($radioButtonEnglish.IsChecked) {$languageId="en-US"; $language = 'English'}
            if ($radioButtonJapanese.IsChecked) {$languageId="ja-JP"; $language = 'Japanese'}
            if ($radioButtonKorean.IsChecked) {$languageId="ko-KR"; $language = 'Korean'}
            if ($radioButtonChinese.IsChecked) {$languageId="zh-TW"; $language = 'Chinese'}
            if ($radioButtonFrench.IsChecked) {$languageId="fr-FR"; $language = 'French'}
            if ($radioButtonSpanish.IsChecked) {$languageId="es-ES"; $language = 'Spanish'}
            if ($radioButtonHindi.IsChecked) {$languageId="hi-IN"; $language = 'Hindi'}
            if ($radioButtonGerman.IsChecked) {$languageId="de-DE"; $language = 'German'}
            if ($radioButtonVietnamese.IsChecked) {$languageId="vi-VN"; $language = 'Vietnamese'}

            if ($radioButtonDownload.IsChecked) {$mode = '/download'; $UIstatus = 'Đang tải'}
            if ($radioButtonInstall.IsChecked) {$mode = '/configure'; $UIstatus = 'Đang cài'}

            if ($radioButton365Home.IsChecked -eq $true) {$productId = "O365HomePremRetail"; $productName = 'Microsoft 365 Home'; $i++}
            if ($radioButton365Business.IsChecked -eq $true) {$productId = "O365BusinessRetail"; $productName = 'Microsoft 365 Apps for Business'; $i++}
            if ($radioButton365Enterprise.IsChecked -eq $true) {$productId = "O365ProPlusRetail"; $productName = 'Microsoft 365 Apps for Enterprise'; $i++}

        # For Office 2024
            if ($radioButton2024Pro.IsChecked -eq $true) {$productId = "ProPlus2024$licType"; $productName = 'Office 2024 Professional LTSC'; $i++}
            if ($radioButton2024Std.IsChecked -eq $true) {$productId = "Standard2024$licType"; $productName = 'Office 2024 Standard LTSC'; $i++}
            if ($radioButton2024ProjectPro.IsChecked -eq $true) {$productId = "ProjectPro2024$licType"; $productName = 'Project Pro 2024'; $i++}
            if ($radioButton2024ProjectStd.IsChecked -eq $true) {$productId = "ProjectStd2024$licType"; $productName = 'Project Standard 2024'; $i++}
            if ($radioButton2024VisioPro.IsChecked -eq $true) {$productId = "VisioPro2024$licType"; $productName = 'Visio Pro 2024'; $i++}
            if ($radioButton2024VisioStd.IsChecked -eq $true) {$productId = "VisioStd2024$licType"; $productName = 'Visio Standard 2024'; $i++}
            if ($radioButton2024Word.IsChecked -eq $true) {$productId = "Word2024$licType"; $productName = 'Microsoft Word LTSC 2024'; $i++}
            if ($radioButton2024Excel.IsChecked -eq $true) {$productId = "Excel2024$licType"; $productName = 'Microsoft Excel LTSC 2024'; $i++}
            if ($radioButton2024PowerPoint.IsChecked -eq $true) {$productId = "PowerPoint2024$licType"; $productName = 'Microsoft PowerPoint LTSC 2024'; $i++}
            if ($radioButton2024Outlook.IsChecked -eq $true) {$productId = "Outlook2024$licType"; $productName = 'Microsoft Outlook LTSC 2024'; $i++}
            if ($radioButton2024Publisher.IsChecked -eq $true) {$productId = "Publisher2024$licType"; $productName = 'Microsoft Publisher LTSC 2024'; $i++}
            if ($radioButton2024Access.IsChecked -eq $true) {$productId = "Access2024$licType"; $productName = 'Microsoft Access LTSC 2024'; $i++}
            if ($radioButton2024HomeBusiness.IsChecked -eq $true) {$productId = "HomeBusiness2024Retail"; $productName = 'Office HomeBusiness 2024'; $i++}
            if ($radioButton2024HomeStudent.IsChecked -eq $true) {$productId = "HomeStudent2024Retail"; $productName = 'Office HomeStudent LTSC 2024'; $i++}

        # For Office 2021
            if ($radioButton2021Pro.IsChecked -eq $true) {$productId = "ProPlus2021$licType"; $productName = 'Office 2021 Professional LTSC'; $i++}
            if ($radioButton2021Std.IsChecked -eq $true) {$productId = "Standard2021$licType"; $productName = 'Office 2021 Standard LTSC'; $i++}
            if ($radioButton2021ProjectPro.IsChecked -eq $true) {$productId = "ProjectPro2021$licType"; $productName = 'Project Pro 2021'; $i++}
            if ($radioButton2021ProjectStd.IsChecked -eq $true) {$productId = "ProjectStd2021$licType"; $productName = 'Project Standard 2021'; $i++}
            if ($radioButton2021VisioPro.IsChecked -eq $true) {$productId = "VisioPro2021$licType"; $productName = 'Visio Pro 2021'; $i++}
            if ($radioButton2021VisioStd.IsChecked -eq $true) {$productId = "VisioStd2021$licType"; $productName = 'Visio Standard 2021'; $i++}
            if ($radioButton2021Word.IsChecked -eq $true) {$productId = "Word2021$licType"; $productName = 'Microsoft Word LTSC 2021'; $i++}
            if ($radioButton2021Excel.IsChecked -eq $true) {$productId = "Excel2021$licType"; $productName = 'Microsoft Excel LTSC 2021'; $i++}
            if ($radioButton2021PowerPoint.IsChecked -eq $true) {$productId = "PowerPoint2021$licType"; $productName = 'Microsoft PowerPoint LTSC 2021'; $i++}
            if ($radioButton2021Outlook.IsChecked -eq $true) {$productId = "Outlook2021$licType"; $productName = 'Microsoft Outlook LTSC 2021'; $i++}
            if ($radioButton2021Publisher.IsChecked -eq $true) {$productId = "Publisher2021$licType"; $productName = 'Microsoft Publisher LTSC 2021'; $i++}
            if ($radioButton2021Access.IsChecked -eq $true) {$productId = "Access2021$licType"; $productName = 'Microsoft Access LTSC 2021'; $i++}
            if ($radioButton2021HomeBusiness.IsChecked -eq $true) {$productId = "HomeBusiness2021Retail"; $productName = 'Office HomeBusiness 2021'; $i++}
            if ($radioButton2021HomeStudent.IsChecked -eq $true) {$productId = "HomeStudent2021Retail"; $productName = 'Office HomeStudent LTSC 2021'; $i++}

        # For Office 2019
            if ($radioButton2019Pro.IsChecked -eq $true) {$productId = "ProPlus2019$licType"; $productName = 'Office 2019 Professional Plus'; $i++}
            if ($radioButton2019Std.IsChecked -eq $true) {$productId = "Standard2019$licType"; $productName = 'Office 2019 Standard'; $i++}
            if ($radioButton2019ProjectPro.IsChecked -eq $true) {$productId = "ProjectPro2019$licType"; $productName = 'Project Pro 2019'; $i++}
            if ($radioButton2019ProjectStd.IsChecked -eq $true) {$productId = "ProjectStd2019$licType"; $productName = 'Project Standard 2019'; $i++}
            if ($radioButton2019VisioPro.IsChecked -eq $true) {$productId = "VisioPro2019$licType"; $productName = 'Visio Pro 2019'; $i++}
            if ($radioButton2019VisioStd.IsChecked -eq $true) {$productId = "VisioStd2019$licType"; $productName = 'Visio Standard 2019'; $i++}
            if ($radioButton2019Word.IsChecked -eq $true) {$productId = "Word2019$licType"; $productName = 'Microsoft Word 2019'; $i++}
            if ($radioButton2019Excel.IsChecked -eq $true) {$productId = "Excel2019$licType"; $productName = 'Microsoft Excel 2019'; $i++}
            if ($radioButton2019PowerPoint.IsChecked -eq $true) {$productId = "PowerPoint2019$licType"; $productName = 'Microsoft PowerPoint 201p'; $i++}
            if ($radioButton2019Outlook.IsChecked -eq $true) {$productId = "Outlook2019$licType"; $productName = 'Microsoft Outlook 2019'; $i++}
            if ($radioButton2019Publisher.IsChecked -eq $true) {$productId = "Publisher2019$licType"; $productName = 'Microsoft Publisher 2019'; $i++}
            if ($radioButton2019Access.IsChecked -eq $true) {$productId = "Access2019$licType"; $productName = 'Microsoft Access 2019'; $i++}
            if ($radioButton2019HomeBusiness.IsChecked -eq $true) {$productId = "HomeBusiness2019Retail"; $productName = 'Office HomeBusiness 2019'; $i++}
            if ($radioButton2019HomeStudent.IsChecked -eq $true) {$productId = "HomeStudent2019Retail"; $productName = 'Office HomeStudent 2019'; $i++}

        # For Office 2016
            if ($radioButton2016Pro.IsChecked -eq $true) {$productId = "ProfessionalRetail"; $uri = $uri2016; $productName = 'Office 2016 Professional Plus'; $i++}
            if ($radioButton2016Std.IsChecked -eq $true) {$productId = "StandardRetail"; $uri = $uri2016; $productName = 'Office 2016 Standard'; $i++}
            if ($radioButton2016ProjectPro.IsChecked -eq $true) {$productId = "ProjectProRetail"; $uri = $uri2016; $productName = 'Microsoft Project Pro 2016'; $i++}
            if ($radioButton2016ProjectStd.IsChecked -eq $true) {$productId = "ProjectStdRetail"; $uri = $uri2016; $productName = 'Microsoft Project Standard 2016'; $i++}
            if ($radioButton2016VisioPro.IsChecked -eq $true) {$productId = "VisioProRetail"; $uri = $uri2016; $productName = 'Microsoft Visio Pro 2016'; $i++}
            if ($radioButton2016VisioStd.IsChecked -eq $true) {$productId = "VisioStdRetail"; $uri = $uri2016; $productName = 'Microsoft Visio Standard 2016'; $i++}
            if ($radioButton2016Word.IsChecked -eq $true) {$productId = "WordRetail"; $uri = $uri2016; $productName = 'Microsoft Word 2016'; $i++}
            if ($radioButton2016Excel.IsChecked -eq $true) {$productId = "ExcelRetail"; $uri = $uri2016; $productName = 'Microsoft Excel 2016'; $i++}
            if ($radioButton2016PowerPoint.IsChecked -eq $true) {$productId = "PowerPointRetail"; $uri = $uri2016; $productName = 'Microsoft PowerPoint 2016'; $i++}
            if ($radioButton2016Outlook.IsChecked -eq $true) {$productId = "OutlookRetail"; $uri = $uri2016; $productName = 'Microsoft Outlook 2016'; $i++}
            if ($radioButton2016Publisher.IsChecked -eq $true) {$productId = "PublisherRetail"; $uri = $uri2016; $productName = 'Microsoft Publisher 2016'; $i++}
            if ($radioButton2016Access.IsChecked -eq $true) {$productId = "AccessRetail"; $uri = $uri2016; $productName = 'Microsoft Access 2016'; $i++}
            if ($radioButton2016OneNote.IsChecked -eq $true) {$productId = "OneNoteRetail"; $uri = $uri2016; $productName = 'Microsoft Onenote 2016'; $i++}

        # For Office 2013
            if ($radioButton2013Pro.IsChecked -eq $true) {$productId = "ProfessionalRetail"; $uri = $uri2013; $productName = 'Office 2013 Professional Plus'; $i++}
            if ($radioButton2013Std.IsChecked -eq $true) {$productId = "StandardRetail"; $uri = $uri2013; $productName = 'Office 2013 Standard'; $i++}
            if ($radioButton2013ProjectPro.IsChecked -eq $true) {$productId = "ProjectProRetail"; $uri = $uri2013; $productName = 'Microsoft Project Pro 2013'; $i++}
            if ($radioButton2013ProjectStd.IsChecked -eq $true) {$productId = "ProjectStdRetail"; $uri = $uri2013; $productName = 'Microsoft Project Standard 2013'; $i++}
            if ($radioButton2013VisioPro.IsChecked -eq $true) {$productId = "VisioProRetail"; $uri = $uri2013; $productName = 'Microsoft Visio Pro 2013'; $i++}
            if ($radioButton2013VisioStd.IsChecked -eq $true) {$productId = "VisioStdRetail"; $uri = $uri2013; $productName = 'Microsoft Visio Standard 2013'; $i++}
            if ($radioButton2013Word.IsChecked -eq $true) {$productId = "WordRetail"; $uri = $uri2013; $productName = 'Microsoft Word 2013'; $i++}
            if ($radioButton2013Excel.IsChecked -eq $true) {$productId = "ExcelRetail"; $uri = $uri2013; $productName = 'Microsoft Excel 2013'; $i++}
            if ($radioButton2013PowerPoint.IsChecked -eq $true) {$productId = "PowerPointRetail"; $uri = $uri2013; $productName = 'Microsoft PowerPoint 2013'; $i++}
            if ($radioButton2013Outlook.IsChecked -eq $true) {$productId = "OutlookRetail"; $uri = $uri2013; $productName = 'Microsoft Outlook 2013'; $i++}
            if ($radioButton2013Publisher.IsChecked -eq $true) {$productId = "PublisherRetail"; $uri = $uri2013; $productName = 'Microsoft Publisher 2013'; $i++}
            if ($radioButton2013Access.IsChecked -eq $true) {$productId = "AccessRetail"; $uri = $uri2013; $productName = 'Microsoft Access 2013'; $i++}
        # Update the shared hashtable
            $sync.arch = $arch
            $sync.mode = $mode
            $sync.language = $language
            $sync.UIstatus = $UIstatus
            $sync.productName = $productName

            if ($i -eq '1') {
                PreparingOffice
                $PSIinstance = [powershell]::Create().AddScript($DownloadInstallOffice)
                $PSIinstance.Runspace = $runspace
                $PSIinstance.BeginInvoke()
            } else {
                $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Foreground = "Red" })
                $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.FontWeight = "Bold" })
                $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Vui lòng chọn 1 phiên bản Office!" })
            } 
        }
    )

# Uninstall all installed Microsoft Office apps.
    $UninstallOffice = {

        $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Đang gỡ cài đặt Microsoft Office..." })
        $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = "Hidden" })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.BorderBrush = "#FF707070" })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })
        $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Visible" })
        
        Set-Location -Path $($sync.workingDir)
        Invoke-Item Path $($sync.workingDir)
  
        (New-Object Net.WebClient).DownloadFile($($sync.removeAllXML), "$($sync.workingDir)\configuration.xml")
        (New-Object Net.WebClient).DownloadFile($($sync.uri), "$($sync.workingDir)\bin.exe")

        $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Đang gỡ bằng Office Deployment Tool..." })
        $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = "Hidden" })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.BorderBrush = "#FF707070" })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })
        $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Visible" })

        Start-Process -FilePath .\bin.exe -ArgumentList "/configure .\configuration.xml" -NoNewWindow -Wait

        if (Test-Path -Path "C:\Program Files*\Microsoft Office\Office15\ospp.vbs") {
            (New-Object Net.WebClient).DownloadFile('https://aka.ms/SaRA_EnterpriseVersionFiles', "$($sync.workingDir)\SaRA.zip")
            Expand-Archive -Path .\SaRA.zip -DestinationPath .\SaRA

            $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Đang chạy kịch bản dọn dẹp (OfficeScrub)..." })
            $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = "Hidden" })
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.BorderBrush = "#FF707070" })
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })
            $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Visible" })

            Start-Process -FilePath ".\SaRA\SaRACmd.exe" -ArgumentList "-S OfficeScrubScenario -AcceptEula -OfficeVersion All" -NoNewWindow -Wait
        }

        $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Hidden" })
        $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = 'Visible' })
        $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Content = 'BẮT ĐẦU' })
        $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = 'Hoàn tất gỡ cài đặt' })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $false })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.Value = '100' })

        # Cleanup
        Set-Location ..
        Remove-Item ClickToRunU -Recurse -Force
    }

    if ($buttonRemoveAll) {
        $buttonRemoveAll.Add_Click({
            if ($radioButtonRemoveAllApp.IsChecked) {
                $workingDir = New-Item -Path $env:temp\ClickToRunU -ItemType Directory -Force
                Set-Location $workingDir
                $sync.workingDir = $workingDir
                $sync.uri = $uri
                $sync.removeAllXML = $removeAllXML

                $PSIinstance = [powershell]::Create().AddScript($UninstallOffice)
                $PSIinstance.Runspace = $runspace
                $PSIinstance.BeginInvoke()
            } else {
                 $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Bạn chưa tích vào ô đồng ý!" })
            }
        })
    }

$null = $Form.ShowDialog()
