# You need to have Administrator rights to run this script!
    if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "Bạn cần quyền Administrator để chạy script này!`nVui lòng chạy lại script dưới quyền Admin (Run as Administrator)!"
        Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "irm office.msedu.vn | iex"
        break
    }

# Load ddls to the current session.
    Add-Type -AssemblyName PresentationFramework, System.Drawing, PresentationFramework, System.Windows.Forms, WindowsFormsIntegration, PresentationCore
    [System.Windows.Forms.Application]::EnableVisualStyles()

# GIAO DIỆN XAML MỚI (MODERN DARK THEME + FACEBOOK LINK)
$xamlInput = @'
<Window x:Class="install.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:install"
        mc:Ignorable="d"
        Title="Huỳnh Dương Developer - Office Installer" 
        Height="650" Width="1100" 
        WindowStartupLocation="CenterScreen" 
        ResizeMode="CanMinimize"
        Background="#1E1E1E"
        Foreground="#FFFFFF"
        FontFamily="Segoe UI">

    <Window.Resources>
        <!-- Style cho GroupBox -->
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="BorderBrush" Value="#3E3E42"/>
            <Setter Property="Foreground" Value="#007ACC"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GroupBox">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Border Grid.Row="0" BorderThickness="0,0,0,1" BorderBrush="#007ACC" Margin="0,0,0,5">
                                <ContentPresenter Margin="5" ContentSource="Header" RecognizesAccessKey="True"/>
                            </Border>
                            <Border Grid.Row="1" BorderThickness="1" BorderBrush="#3E3E42" CornerRadius="3" Background="#252526">
                                <ContentPresenter Margin="10" />
                            </Border>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Style cho RadioButton (Product) -->
        <Style x:Key="ProductRadio" TargetType="RadioButton">
            <Setter Property="Foreground" Value="#CCCCCC"/>
            <Setter Property="Margin" Value="0,5,10,5"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="Normal"/>
            <Setter Property="GroupName" Value="OfficeVersion"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsChecked" Value="True">
                    <Setter Property="Foreground" Value="#00A4EF"/>
                    <Setter Property="FontWeight" Value="Bold"/>
                </Trigger>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Foreground" Value="#FFFFFF"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Style cho RadioButton (Option nhỏ bên trái) -->
        <Style x:Key="OptionRadio" TargetType="RadioButton">
            <Setter Property="Foreground" Value="#DDDDDD"/>
            <Setter Property="Margin" Value="0,3,0,3"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>

        <!-- Style cho Button Chính -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#0097FB"/>
                    <Setter Property="Cursor" Value="Hand"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#555555"/>
                    <Setter Property="Foreground" Value="#AAAAAA"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="260"/> <!-- Cột Menu Trái -->
            <ColumnDefinition Width="20"/>  <!-- Khoảng cách -->
            <ColumnDefinition Width="*"/>   <!-- Cột Nội Dung Phải -->
        </Grid.ColumnDefinitions>

        <!-- === CỘT TRÁI: CẤU HÌNH === -->
        <StackPanel Grid.Column="0">
            <!-- Logo / Tiêu đề -->
            <Border Background="#2D2D30" CornerRadius="5" Padding="10" Margin="0,0,0,15">
                <StackPanel>
                    <TextBlock Text="HUỲNH DƯƠNG" Foreground="#00A4EF" FontWeight="Bold" FontSize="18" HorizontalAlignment="Center"/>
                    <TextBlock Text="DEVELOPER" Foreground="White" FontWeight="Light" FontSize="14" HorizontalAlignment="Center" Margin="0,-2,0,0"/>
                    <Image x:Name="image" Height="60" Width="60" Source="https://scontent.fsgn2-9.fna.fbcdn.net/v/t39.30808-6/551384385_122223446444134054_1467993974427749848_n.jpg?_nc_cat=103&ccb=1-7&_nc_sid=6ee11a&_nc_ohc=J5mqzzj9NzIQ7kNvwG0ZRRx&_nc_oc=AdkUOvUdojgnksikilA6jBv2XA0Kg_lmRz3dQmnpzex63LXIOBkRUXxf6B5QAaO84IvBlIJafdCfPtxMEXcivN3Z&_nc_zt=23&_nc_ht=scontent.fsgn2-9.fna&_nc_gid=SYS6WLeOiT-BZPAPeaBU6w&oh=00_AflY8hLhQt8k7IyVGH3MMxa--kteDNhP0Aas1CGmoOKLWw&oe=69505394" Margin="0,10,0,0" Visibility="Hidden"/>
                </StackPanel>
            </Border>

            <!-- Kiến trúc -->
            <GroupBox x:Name="groupBoxArch" Header="1. Kiến Trúc (Architecture)">
                <StackPanel>
                    <RadioButton x:Name="radioButtonArch64" Style="{StaticResource OptionRadio}" Content="x64 (64-bit) - Khuyên dùng" IsChecked="True"/>
                    <RadioButton x:Name="radioButtonArch32" Style="{StaticResource OptionRadio}" Content="x86 (32-bit)"/>
                </StackPanel>
            </GroupBox>

            <!-- Loại Giấy Phép -->
            <GroupBox x:Name="groupBoxLicenseType" Header="2. Loại Giấy Phép (License)">
                <StackPanel>
                    <RadioButton x:Name="radioButtonVolume" Style="{StaticResource OptionRadio}" Content="Volume (VL)" IsChecked="True"/>
                    <RadioButton x:Name="radioButtonRetail" Style="{StaticResource OptionRadio}" Content="Retail (Bán lẻ)"/>
                </StackPanel>
            </GroupBox>

            <!-- Chế Độ -->
            <GroupBox x:Name="groupBoxMode" Header="3. Chế Độ (Mode)">
                <StackPanel>
                    <RadioButton x:Name="radioButtonInstall" Style="{StaticResource OptionRadio}" Content="Install (Cài đặt ngay)" IsChecked="True"/>
                    <RadioButton x:Name="radioButtonDownload" Style="{StaticResource OptionRadio}" Content="Download (Tải bộ cài)"/>
                </StackPanel>
            </GroupBox>

            <!-- Ngôn Ngữ -->
            <GroupBox x:Name="groupBoxLanguage" Header="4. Ngôn Ngữ (Language)">
                <UniformGrid Columns="2">
                    <RadioButton x:Name="radioButtonEnglish" Style="{StaticResource OptionRadio}" Content="English" IsChecked="True"/>
                    <RadioButton x:Name="radioButtonVietnamese" Style="{StaticResource OptionRadio}" Content="Tiếng Việt"/>
                    <RadioButton x:Name="radioButtonJapanese" Style="{StaticResource OptionRadio}" Content="Japanese"/>
                    <RadioButton x:Name="radioButtonKorean" Style="{StaticResource OptionRadio}" Content="Korean"/>
                    <RadioButton x:Name="radioButtonChinese" Style="{StaticResource OptionRadio}" Content="Chinese"/>
                    <RadioButton x:Name="radioButtonFrench" Style="{StaticResource OptionRadio}" Content="French"/>
                    <RadioButton x:Name="radioButtonSpanish" Style="{StaticResource OptionRadio}" Content="Spanish"/>
                    <RadioButton x:Name="radioButtonGerman" Style="{StaticResource OptionRadio}" Content="German"/>
                    <RadioButton x:Name="radioButtonHindi" Style="{StaticResource OptionRadio}" Content="Hindi"/>
                </UniformGrid>
            </GroupBox>

            <!-- Nút Hành Động -->
            <Button x:Name="buttonSubmit" Content="BẮT ĐẦU THỰC HIỆN" Height="40" Margin="0,10,0,0" Background="#107C10"/>
            
            <ProgressBar x:Name="progressbar" Height="5" Margin="0,10,0,0" Background="#333333" BorderThickness="0" Foreground="#00A4EF"/>
            <TextBox x:Name="textbox" Text="Sẵn sàng..." Background="Transparent" Foreground="#AAAAAA" BorderThickness="0" TextWrapping="Wrap" Margin="0,5,0,0" HorizontalContentAlignment="Center" IsReadOnly="True"/>
            
            <!-- Link Web & Facebook -->
            <StackPanel Margin="0,15,0,0">
                <Label x:Name="Link1" HorizontalAlignment="Center" Cursor="Hand" Padding="5,2">
                    <Hyperlink NavigateUri="https://nlacc.site" Foreground="#DDDDDD" TextDecorations="None">
                        <TextBlock Text="🌐 Website: nlacc.site"/>
                    </Hyperlink>
                </Label>
                
                <Label x:Name="LinkFacebook" HorizontalAlignment="Center" Cursor="Hand" Padding="5,2">
                    <Hyperlink NavigateUri="https://www.facebook.com/huynh.duong.389204/" Foreground="#1877F2" TextDecorations="None" FontWeight="Bold">
                        <TextBlock Text="f  FB Của Tao!"/>
                    </Hyperlink>
                </Label>
            </StackPanel>
        </StackPanel>

        <!-- === CỘT PHẢI: DANH SÁCH OFFICE === -->
        <Border Grid.Column="2" Background="#252526" CornerRadius="5" BorderBrush="#3E3E42" BorderThickness="1">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Header Phải -->
                <Border Background="#2D2D30" Padding="15,10" CornerRadius="5,5,0,0">
                    <TextBlock Text="CHỌN PHIÊN BẢN OFFICE CẦN CÀI ĐẶT" FontWeight="Bold" FontSize="15" Foreground="#FFFFFF"/>
                </Border>

                <!-- Danh sách cuộn -->
                <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="15">
                    <StackPanel>
                        
                        <!-- Microsoft 365 -->
                        <GroupBox Header="Microsoft 365 (Office 365)" BorderBrush="#FFDA2323" Foreground="#FF5E5E">
                            <WrapPanel>
                                <RadioButton x:Name="radioButton365Home" Style="{StaticResource ProductRadio}" Content="Home (Cá nhân)"/>
                                <RadioButton x:Name="radioButton365Business" Style="{StaticResource ProductRadio}" Content="Business (Doanh nghiệp nhỏ)"/>
                                <RadioButton x:Name="radioButton365Enterprise" Style="{StaticResource ProductRadio}" Content="Enterprise (Doanh nghiệp lớn)"/>
                            </WrapPanel>
                        </GroupBox>

                        <!-- Office 2024 -->
                        <GroupBox Header="Office 2024 LTSC (Mới nhất)" BorderBrush="#FFE2820E" Foreground="#FFA500">
                            <WrapPanel>
                                <RadioButton x:Name="radioButton2024Pro" Style="{StaticResource ProductRadio}" Content="Professional Plus"/>
                                <RadioButton x:Name="radioButton2024Std" Style="{StaticResource ProductRadio}" Content="Standard"/>
                                <RadioButton x:Name="radioButton2024HomeBusiness" Style="{StaticResource ProductRadio}" Content="Home &amp; Business"/>
                                <RadioButton x:Name="radioButton2024HomeStudent" Style="{StaticResource ProductRadio}" Content="Home &amp; Student"/>
                                <RadioButton x:Name="radioButton2024ProjectPro" Style="{StaticResource ProductRadio}" Content="Project Pro"/>
                                <RadioButton x:Name="radioButton2024VisioPro" Style="{StaticResource ProductRadio}" Content="Visio Pro"/>
                                <RadioButton x:Name="radioButton2024Word" Style="{StaticResource ProductRadio}" Content="Word"/>
                                <RadioButton x:Name="radioButton2024Excel" Style="{StaticResource ProductRadio}" Content="Excel"/>
                                <RadioButton x:Name="radioButton2024PowerPoint" Style="{StaticResource ProductRadio}" Content="PowerPoint"/>
                            </WrapPanel>
                        </GroupBox>

                        <!-- Office 2021 -->
                        <GroupBox Header="Office 2021 LTSC" BorderBrush="#FF3C10DE" Foreground="#8A75FF">
                            <WrapPanel>
                                <RadioButton x:Name="radioButton2021Pro" Style="{StaticResource ProductRadio}" Content="Professional Plus"/>
                                <RadioButton x:Name="radioButton2021Std" Style="{StaticResource ProductRadio}" Content="Standard"/>
                                <RadioButton x:Name="radioButton2021ProjectPro" Style="{StaticResource ProductRadio}" Content="Project Pro"/>
                                <RadioButton x:Name="radioButton2021VisioPro" Style="{StaticResource ProductRadio}" Content="Visio Pro"/>
                                <RadioButton x:Name="radioButton2021Word" Style="{StaticResource ProductRadio}" Content="Word"/>
                                <RadioButton x:Name="radioButton2021Excel" Style="{StaticResource ProductRadio}" Content="Excel"/>
                                <RadioButton x:Name="radioButton2021PowerPoint" Style="{StaticResource ProductRadio}" Content="PowerPoint"/>
                                <RadioButton x:Name="radioButton2021Access" Style="{StaticResource ProductRadio}" Content="Access"/>
                                <RadioButton x:Name="radioButton2021Outlook" Style="{StaticResource ProductRadio}" Content="Outlook"/>
                                <!-- Ẩn bớt các bản ít dùng để gọn, nếu cần có thể thêm lại như code cũ -->
                                <RadioButton x:Name="radioButton2021ProjectStd" Style="{StaticResource ProductRadio}" Content="Project Std" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2021VisioStd" Style="{StaticResource ProductRadio}" Content="Visio Std" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2021HomeStudent" Style="{StaticResource ProductRadio}" Content="Home Student" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2021HomeBusiness" Style="{StaticResource ProductRadio}" Content="Home Business" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2021Publisher" Style="{StaticResource ProductRadio}" Content="Publisher" Visibility="Collapsed"/>
                            </WrapPanel>
                        </GroupBox>

                        <!-- Office 2019 -->
                        <GroupBox Header="Office 2019" BorderBrush="#FF0F8E40" Foreground="#4CAF50">
                            <WrapPanel>
                                <RadioButton x:Name="radioButton2019Pro" Style="{StaticResource ProductRadio}" Content="Professional Plus"/>
                                <RadioButton x:Name="radioButton2019Std" Style="{StaticResource ProductRadio}" Content="Standard"/>
                                <RadioButton x:Name="radioButton2019ProjectPro" Style="{StaticResource ProductRadio}" Content="Project Pro"/>
                                <RadioButton x:Name="radioButton2019VisioPro" Style="{StaticResource ProductRadio}" Content="Visio Pro"/>
                                <RadioButton x:Name="radioButton2019Word" Style="{StaticResource ProductRadio}" Content="Word"/>
                                <RadioButton x:Name="radioButton2019Excel" Style="{StaticResource ProductRadio}" Content="Excel"/>
                                <RadioButton x:Name="radioButton2019PowerPoint" Style="{StaticResource ProductRadio}" Content="PowerPoint"/>
                                
                                <RadioButton x:Name="radioButton2019ProjectStd" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2019VisioStd" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2019Outlook" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2019Access" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2019Publisher" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2019HomeStudent" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2019HomeBusiness" Visibility="Collapsed" GroupName="OfficeVersion"/>
                            </WrapPanel>
                        </GroupBox>
                        
                        <!-- Office 2016 -->
                         <GroupBox Header="Office 2016" BorderBrush="#FFA28210" Foreground="#FFD700">
                            <WrapPanel>
                                <RadioButton x:Name="radioButton2016Pro" Style="{StaticResource ProductRadio}" Content="Professional Plus"/>
                                <RadioButton x:Name="radioButton2016Std" Style="{StaticResource ProductRadio}" Content="Standard"/>
                                <RadioButton x:Name="radioButton2016ProjectPro" Style="{StaticResource ProductRadio}" Content="Project Pro"/>
                                <RadioButton x:Name="radioButton2016VisioPro" Style="{StaticResource ProductRadio}" Content="Visio Pro"/>
                                
                                <RadioButton x:Name="radioButton2016ProjectStd" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2016VisioStd" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2016Word" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2016Excel" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2016PowerPoint" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2016Outlook" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2016Access" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2016Publisher" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2016OneNote" Visibility="Collapsed" GroupName="OfficeVersion"/>
                            </WrapPanel>
                        </GroupBox>

                        <!-- Office 2013 -->
                         <GroupBox Header="Office 2013" BorderBrush="#777777" Foreground="#AAAAAA">
                            <WrapPanel>
                                <RadioButton x:Name="radioButton2013Pro" Style="{StaticResource ProductRadio}" Content="Professional Plus"/>
                                <RadioButton x:Name="radioButton2013Std" Style="{StaticResource ProductRadio}" Content="Standard"/>
                                
                                <RadioButton x:Name="radioButton2013ProjectPro" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2013ProjectStd" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2013VisioPro" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2013VisioStd" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2013Word" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2013Excel" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2013PowerPoint" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2013Outlook" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2013Access" Visibility="Collapsed" GroupName="OfficeVersion"/>
                                <RadioButton x:Name="radioButton2013Publisher" Visibility="Collapsed" GroupName="OfficeVersion"/>
                            </WrapPanel>
                        </GroupBox>
                        
                    </StackPanel>
                </ScrollViewer>

                <!-- Footer Danger Zone -->
                <Border Grid.Row="2" Background="#330000" Padding="10" Margin="10" CornerRadius="4" BorderBrush="#FF4444" BorderThickness="1">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <Label x:Name="LabelRemoveAll" Content="⚠ VÙNG NGUY HIỂM:" FontWeight="Bold" Foreground="#FF4444" VerticalAlignment="Center"/>
                        <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="10,0,0,0">
                            <RadioButton x:Name="radioButtonRemoveAllApp" Content="Tôi đồng ý gỡ bỏ toàn bộ Office" Foreground="White"/>
                            <TextBlock x:Name="textBoxRemoveAll" Text="Hành động này sẽ xóa sạch mọi phiên bản Office trên máy." FontSize="10" Foreground="#FF8888" Margin="18,2,0,0"/>
                        </StackPanel>
                        <Button x:Name="buttonRemoveAll" Grid.Column="2" Content="GỠ CÀI ĐẶT" Background="#CC0000" Foreground="White" Width="100" FontWeight="Bold"/>
                    </Grid>
                </Border>
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

    # --- KHU VỰC CẤU HÌNH LINK FACEBOOK ---
    # Thay link facebook của bạn vào bên dưới:
    $FacebookURL = "https://www.facebook.com/HuynhDuongDeveloperPro" 
    
    $Link1.Add_PreviewMouseDown({[system.Diagnostics.Process]::start('https://msedu.vn')})
    $LinkFacebook.Add_PreviewMouseDown({[system.Diagnostics.Process]::start($FacebookURL)})

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
            $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Content = 'BẮT ĐẦU THỰC HIỆN' })
            $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = 'Hoàn tất' })
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
                $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Vui lòng chọn một phiên bản Office!" })
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
        $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Content = 'BẮT ĐẦU THỰC HIỆN' })
        $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = 'Hoàn tất' })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $false })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.Value = '100' })

        # Cleanup
        Set-Location ..
        Remove-Item ClickToRunU -Recurse -Force
    }

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
        }
    })

$null = $Form.ShowDialog()
