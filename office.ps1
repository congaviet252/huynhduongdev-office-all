# You need to have Administrator rights to run this script!
    if (-not([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning "Bạn cần quyền Administrator để chạy script này!`nVui lòng chạy lại script dưới quyền Admin (Run as Administrator)!"
        Start-Process -Verb runas -FilePath powershell.exe -ArgumentList "irm office.msedu.vn | iex"
        break
    }

# Load ddls to the current session.
    Add-Type -AssemblyName PresentationFramework, System.Drawing, PresentationFramework, System.Windows.Forms, WindowsFormsIntegration, PresentationCore
    [System.Windows.Forms.Application]::EnableVisualStyles()

# --- GIAO DIỆN XAML NÂNG CẤP (MODERN DARK UI) ---
$xamlInput = @'
<Window x:Class="install.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Huỳnh Dương Developer Pro - Office Tool" 
        Height="600" Width="1000" 
        WindowStartupLocation="CenterScreen" 
        ResizeMode="CanMinimize"
        Background="#1E1E1E"
        Icon="https://raw.githubusercontent.com/mseduvn/msoffice/refs/heads/main/Files/images.png">

    <Window.Resources>
        <!-- Màu sắc chủ đạo -->
        <SolidColorBrush x:Key="PrimaryColor" Color="#007ACC"/>
        <SolidColorBrush x:Key="AccentColor" Color="#28C840"/>
        <SolidColorBrush x:Key="DangerColor" Color="#E81123"/>
        <SolidColorBrush x:Key="DarkBg" Color="#252526"/>
        <SolidColorBrush x:Key="LightText" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="GrayText" Color="#CCCCCC"/>
        <SolidColorBrush x:Key="BorderColor" Color="#3E3E42"/>

        <!-- Style cho RadioButton dạng thẻ (Card) -->
        <Style x:Key="CardRadioButton" TargetType="RadioButton">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource LightText}"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="border" BorderBrush="{StaticResource BorderColor}" BorderThickness="1" Background="#2D2D30" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{StaticResource PrimaryColor}"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="{StaticResource PrimaryColor}"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="BorderBrush" Value="{StaticResource PrimaryColor}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Style cho GroupBox -->
        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="{StaticResource PrimaryColor}"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
        </Style>

        <!-- Style cho Button chính -->
        <Style x:Key="MainButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource AccentColor}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" CornerRadius="5">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#32D74B"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#555"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="10">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="220"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- CỘT TRÁI: CẤU HÌNH -->
        <Border Grid.Column="0" Background="{StaticResource DarkBg}" CornerRadius="5" Margin="0,0,10,0" Padding="10">
            <StackPanel>
                <TextBlock Text="CẤU HÌNH" Foreground="{StaticResource GrayText}" FontWeight="Bold" Margin="0,0,0,10" HorizontalAlignment="Center"/>
                
                <!-- Kiến trúc -->
                <GroupBox Header="Kiến trúc (Bit)">
                    <StackPanel>
                        <RadioButton x:Name="radioButtonArch64" Content="x64 (64-bit)" IsChecked="True" Foreground="White" Margin="5"/>
                        <RadioButton x:Name="radioButtonArch32" Content="x86 (32-bit)" Foreground="White" Margin="5"/>
                    </StackPanel>
                </GroupBox>

                <!-- Giấy phép -->
                <GroupBox Header="Loại giấy phép">
                    <StackPanel>
                        <RadioButton x:Name="radioButtonVolume" Content="Volume (VL)" IsChecked="True" Foreground="White" Margin="5"/>
                        <RadioButton x:Name="radioButtonRetail" Content="Retail" Foreground="White" Margin="5"/>
                    </StackPanel>
                </GroupBox>

                <!-- Chế độ -->
                <GroupBox Header="Chế độ hoạt động">
                    <StackPanel>
                        <RadioButton x:Name="radioButtonInstall" Content="Cài đặt (Install)" IsChecked="True" Foreground="White" Margin="5"/>
                        <RadioButton x:Name="radioButtonDownload" Content="Tải bộ cài (Download)" Foreground="White" Margin="5"/>
                    </StackPanel>
                </GroupBox>

                <!-- Ngôn ngữ -->
                <GroupBox Header="Ngôn ngữ">
                    <ComboBox SelectedIndex="0" Height="25">
                        <ComboBoxItem IsSelected="True">
                             <RadioButton x:Name="radioButtonEnglish" Content="English (US)" IsChecked="True" BorderThickness="0"/>
                        </ComboBoxItem>
                         <ComboBoxItem>
                             <RadioButton x:Name="radioButtonVietnamese" Content="Tiếng Việt" BorderThickness="0"/>
                        </ComboBoxItem>
                        <ComboBoxItem>
                             <RadioButton x:Name="radioButtonJapanese" Content="Japanese" BorderThickness="0"/>
                        </ComboBoxItem>
                         <ComboBoxItem>
                             <RadioButton x:Name="radioButtonKorean" Content="Korean" BorderThickness="0"/>
                        </ComboBoxItem>
                         <ComboBoxItem>
                             <RadioButton x:Name="radioButtonChinese" Content="Chinese" BorderThickness="0"/>
                        </ComboBoxItem>
                    </ComboBox>
                     <!-- Các Radio ẩn để giữ logic code cũ -->
                    <StackPanel Visibility="Collapsed">
                        <RadioButton x:Name="radioButtonFrench"/>
                        <RadioButton x:Name="radioButtonSpanish"/>
                        <RadioButton x:Name="radioButtonHindi"/>
                        <RadioButton x:Name="radioButtonGerman"/>
                    </StackPanel>
                </GroupBox>

                <Image x:Name="image" Source="https://raw.githubusercontent.com/mseduvn/msoffice/refs/heads/main/Files/images.png" Height="80" Visibility="Hidden"/>
            </StackPanel>
        </Border>

        <!-- CỘT PHẢI: LỰA CHỌN PHIÊN BẢN -->
        <Grid Grid.Column="1">
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <!-- Tabs chọn phiên bản -->
            <TabControl Grid.Row="0" Background="Transparent" BorderThickness="0">
                <TabControl.Resources>
                    <Style TargetType="TabItem">
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="TabItem">
                                    <Border Name="Border" BorderThickness="0,0,0,2" BorderBrush="Transparent" Margin="0,0,10,0" Padding="10,5">
                                        <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" HorizontalAlignment="Center" ContentSource="Header" Margin="10,2"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter TargetName="Border" Property="BorderBrush" Value="{StaticResource PrimaryColor}"/>
                                            <Setter Property="Foreground" Value="{StaticResource PrimaryColor}"/>
                                            <Setter Property="FontWeight" Value="Bold"/>
                                        </Trigger>
                                        <Trigger Property="IsSelected" Value="False">
                                            <Setter Property="Foreground" Value="{StaticResource GrayText}"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </TabControl.Resources>

                <!-- TAB OFFICE 2024 (MỚI NHẤT) -->
                <TabItem Header="OFFICE 2024">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <WrapPanel Orientation="Horizontal" ItemWidth="160">
                            <RadioButton x:Name="radioButton2024Pro" Content="Pro Plus 2024" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2024Std" Content="Standard 2024" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2024ProjectPro" Content="Project Pro 2024" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2024VisioPro" Content="Visio Pro 2024" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2024Word" Content="Word 2024" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2024Excel" Content="Excel 2024" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2024PowerPoint" Content="PowerPoint 2024" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2024Outlook" Content="Outlook 2024" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2024Access" Content="Access 2024" Style="{StaticResource CardRadioButton}"/>
                            <!-- Ẩn bớt các bản ít dùng để gọn, logic vẫn chạy nếu chọn -->
                            <RadioButton x:Name="radioButton2024ProjectStd" Content="Project Std" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2024VisioStd" Content="Visio Std" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2024Publisher" Content="Publisher" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2024HomeStudent" Content="Home Student" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2024HomeBusiness" Content="Home Business" Visibility="Collapsed"/>
                        </WrapPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- TAB MICROSOFT 365 -->
                <TabItem Header="MICROSOFT 365">
                    <WrapPanel Orientation="Horizontal" ItemWidth="200">
                        <RadioButton x:Name="radioButton365Enterprise" Content="Apps for Enterprise" Style="{StaticResource CardRadioButton}"/>
                        <RadioButton x:Name="radioButton365Business" Content="Apps for Business" Style="{StaticResource CardRadioButton}"/>
                        <RadioButton x:Name="radioButton365Home" Content="Home Premium" Style="{StaticResource CardRadioButton}"/>
                    </WrapPanel>
                </TabItem>

                <!-- TAB OFFICE 2021 -->
                <TabItem Header="OFFICE 2021">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <WrapPanel Orientation="Horizontal" ItemWidth="160">
                            <RadioButton x:Name="radioButton2021Pro" Content="Pro Plus 2021" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2021Std" Content="Standard 2021" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2021ProjectPro" Content="Project Pro 2021" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2021VisioPro" Content="Visio Pro 2021" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2021Word" Content="Word 2021" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2021Excel" Content="Excel 2021" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2021PowerPoint" Content="PowerPoint 2021" Style="{StaticResource CardRadioButton}"/>
                            <!-- Hidden logic fields -->
                            <RadioButton x:Name="radioButton2021ProjectStd" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2021VisioStd" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2021Outlook" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2021Access" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2021Publisher" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2021HomeStudent" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2021HomeBusiness" Visibility="Collapsed"/>
                        </WrapPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- TAB OFFICE 2019 -->
                <TabItem Header="OFFICE 2019">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <WrapPanel Orientation="Horizontal" ItemWidth="160">
                            <RadioButton x:Name="radioButton2019Pro" Content="Pro Plus 2019" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2019Std" Content="Standard 2019" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2019Word" Content="Word 2019" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2019Excel" Content="Excel 2019" Style="{StaticResource CardRadioButton}"/>
                            <RadioButton x:Name="radioButton2019PowerPoint" Content="PowerPoint 2019" Style="{StaticResource CardRadioButton}"/>
                            <!-- Hidden -->
                            <RadioButton x:Name="radioButton2019ProjectPro" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2019ProjectStd" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2019VisioPro" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2019VisioStd" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2019Outlook" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2019Access" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2019Publisher" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2019HomeStudent" Visibility="Collapsed"/>
                            <RadioButton x:Name="radioButton2019HomeBusiness" Visibility="Collapsed"/>
                        </WrapPanel>
                    </ScrollViewer>
                </TabItem>

                <!-- TAB CŨ HƠN (2016/2013) -->
                <TabItem Header="CŨ HƠN (2016/13)">
                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                        <StackPanel>
                            <TextBlock Text="Office 2016" Foreground="{StaticResource GrayText}" FontWeight="Bold" Margin="5"/>
                            <WrapPanel Orientation="Horizontal" ItemWidth="150">
                                <RadioButton x:Name="radioButton2016Pro" Content="Pro Plus 2016" Style="{StaticResource CardRadioButton}"/>
                                <RadioButton x:Name="radioButton2016Std" Content="Standard 2016" Style="{StaticResource CardRadioButton}"/>
                                <!-- Hidden 2016 others to save space -->
                                <RadioButton x:Name="radioButton2016ProjectPro" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2016ProjectStd" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2016VisioPro" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2016VisioStd" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2016Word" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2016Excel" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2016PowerPoint" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2016Outlook" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2016Access" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2016Publisher" Visibility="Collapsed"/>
                                <RadioButton x:Name="radioButton2016OneNote" Visibility="Collapsed"/>
                            </WrapPanel>

                            <TextBlock Text="Office 2013" Foreground="{StaticResource GrayText}" FontWeight="Bold" Margin="5,10,5,5"/>
                            <WrapPanel Orientation="Horizontal" ItemWidth="150">
                                <RadioButton x:Name="radioButton2013Pro" Content="Pro Plus 2013" Style="{StaticResource CardRadioButton}"/>
                                <RadioButton x:Name="radioButton2013Std" Content="Standard 2013" Style="{StaticResource CardRadioButton}"/>
                                <!-- Hidden 2013 others -->
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
                        </StackPanel>
                    </ScrollViewer>
                </TabItem>
                
                <!-- TAB CÔNG CỤ XÓA -->
                <TabItem Header="CÔNG CỤ XÓA">
                    <Border Background="#2D2D30" CornerRadius="5" Padding="20">
                        <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center">
                             <Label x:Name="LabelRemoveAll" Content="GỠ BỎ TOÀN BỘ OFFICE" Foreground="#E81123" FontWeight="Bold" HorizontalAlignment="Center" FontSize="16"/>
                             <TextBlock x:Name="textBoxRemoveAll" TextWrapping="Wrap" Text="Lưu ý: Hành động này sẽ gỡ sạch mọi phiên bản Office trên máy." Foreground="White" Margin="0,10,0,20" HorizontalAlignment="Center"/>
                             
                             <RadioButton x:Name="radioButtonRemoveAllApp" Content="Tôi hiểu và đồng ý xóa" Foreground="White" HorizontalAlignment="Center" Margin="0,0,0,10"/>
                             
                             <Button x:Name="buttonRemoveAll" Content="GỠ CÀI ĐẶT NGAY" Width="200" Height="40" Background="#E81123" Foreground="White" FontWeight="Bold">
                                 <Button.Template>
                                    <ControlTemplate TargetType="Button">
                                        <Border x:Name="bdr" Background="{TemplateBinding Background}" CornerRadius="5">
                                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                        </Border>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsMouseOver" Value="True">
                                                <Setter TargetName="bdr" Property="Background" Value="#FF4C4C"/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                 </Button.Template>
                             </Button>
                        </StackPanel>
                    </Border>
                </TabItem>
            </TabControl>

            <!-- TRẠNG THÁI VÀ NÚT CHẠY -->
            <StackPanel Grid.Row="1" Margin="0,10,0,0">
                <TextBox x:Name="textbox" Text="Sẵn sàng..." Background="Transparent" BorderThickness="0" Foreground="{StaticResource AccentColor}" FontWeight="Bold" FontSize="14" IsReadOnly="True" HorizontalContentAlignment="Right"/>
                <ProgressBar x:Name="progressbar" Height="5" Background="#333" BorderThickness="0" Foreground="{StaticResource AccentColor}"/>
            </StackPanel>

            <Grid Grid.Row="2" Margin="0,15,0,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <Label x:Name="Link1" Content="Trang chủ: msedu.vn" Foreground="{StaticResource GrayText}" Cursor="Hand" VerticalAlignment="Center" FontSize="11">
                    <Label.Style>
                         <Style TargetType="Label">
                            <Style.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Foreground" Value="{StaticResource PrimaryColor}"/>
                                </Trigger>
                            </Style.Triggers>
                         </Style>
                    </Label.Style>
                </Label>

                <Button x:Name="buttonSubmit" Grid.Column="2" Content="BẮT ĐẦU THỰC HIỆN" Width="200" Height="40" Style="{StaticResource MainButton}"/>
            </Grid>
        </Grid>
    </Grid>
</Window>
'@

# Store form objects (variables) in PowerShell

    [xml]$xaml = $xamlInput -replace 'mc:Ignorable="d"','' -replace "x:Name",'Name'
    $xmlReader = (New-Object System.Xml.XmlNodeReader $xaml)
    $Form = [Windows.Markup.XamlReader]::Load( $xmlReader)

    $xaml.SelectNodes("//*[@Name]") | ForEach-Object -Process {
        Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)
    }

    $Link1.Add_PreviewMouseDown({[system.Diagnostics.Process]::start('https://msedu.vn')})

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
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })
            # $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Visible" }) # Image hidden in modern UI design

        Set-Location -Path $($sync.workingDir)

        Start-Process -FilePath .\bin.exe -ArgumentList "$($sync.mode) .\$($sync.configurationFile)" -NoNewWindow -Wait
                
        # Bring back our Button, set the Label and ProgressBar, we're done..
            # $sync.Form.Dispatcher.Invoke([action] { $sync.image.Visibility = "Hidden" })
            $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = 'Visible' })
            $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Content = 'Bắt Đầu' })
            $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = 'Hoàn tất tác vụ!' })
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
                $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Foreground = "#E81123" }) # Red color for warning
                $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Vui lòng chọn 1 phiên bản Office!" })
            } 
        }
    )

# Uninstall all installed Microsoft Office apps.
    $UninstallOffice = {

        $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Đang gỡ cài đặt Microsoft Office..." })
        $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = "Hidden" })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })
        
        Set-Location -Path $($sync.workingDir)
        Invoke-Item Path $($sync.workingDir)
  
        (New-Object Net.WebClient).DownloadFile($($sync.removeAllXML), "$($sync.workingDir)\configuration.xml")
        (New-Object Net.WebClient).DownloadFile($($sync.uri), "$($sync.workingDir)\bin.exe")

        $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Đang gỡ bằng Office Deployment Tool..." })
        $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = "Hidden" })
        $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })

        Start-Process -FilePath .\bin.exe -ArgumentList "/configure .\configuration.xml" -NoNewWindow -Wait

        if (Test-Path -Path "C:\Program Files*\Microsoft Office\Office15\ospp.vbs") {
            (New-Object Net.WebClient).DownloadFile('https://aka.ms/SaRA_EnterpriseVersionFiles', "$($sync.workingDir)\SaRA.zip")
            Expand-Archive -Path .\SaRA.zip -DestinationPath .\SaRA

            $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = "Đang chạy kịch bản dọn dẹp (OfficeScrub)..." })
            $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = "Hidden" })
            $sync.Form.Dispatcher.Invoke([action] { $sync.ProgressBar.IsIndeterminate = $true })

            Start-Process -FilePath ".\SaRA\SaRACmd.exe" -ArgumentList "-S OfficeScrubScenario -AcceptEula -OfficeVersion All" -NoNewWindow -Wait
        }

        $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Visibility = 'Visible' })
        $sync.Form.Dispatcher.Invoke([action] { $sync.buttonSubmit.Content = 'Bắt Đầu' })
        $sync.Form.Dispatcher.Invoke([action] { $sync.textbox.Text = 'Đã gỡ cài đặt hoàn tất!' })
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
