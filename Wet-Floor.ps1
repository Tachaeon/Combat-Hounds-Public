$Version = "1.5"

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework") 

# Create object for the systray 
$Systray_Tool_Icon = New-Object System.Windows.Forms.NotifyIcon

# Text displayed when you pass the mouse over the systray icon
$Systray_Tool_Icon.Text = "Wet Floor"

#Systray icon
$IconBase64 = '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'
$IconBytes = [Convert]::FromBase64String($IconBase64)
$Stream = [System.IO.MemoryStream]::new($IconBytes, 0, $IconBytes.Length)
$Systray_Tool_Icon.Icon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($Stream).GetHIcon()))
$Systray_Tool_Icon.Visible = $true

# Create object for the systray 
$contextmenu = New-Object System.Windows.Forms.ContextMenuStrip

#Menus
$Menu1 = $contextmenu.Items.Add("Active Directory")
$IconBase64 = 'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAER0lEQVRIS91UXWwUVRg9d+d/d/aHdsuWbQWFtpQfSTTypIlINGokNib6ZGJMfMBEw6v6Vl+MEoyBBBMeBI0aBAIPYID4oBASo6TFYmgBS7u73YXdme7ObHd3dnb+7nhLYgIPhRolMd7JzZ3J3HvOd865+Qge8CAPGB//DYKjR1/jsiknY2jG01a5qMBzNISuKVBqd1pUCsHHvcBOi5y8whHUTTTCFVdF6YEdH8yZ91QQhvuk0vXm+kqpsYOX0i+3veDxwJnmQ1ezArfWJtR0aWAL8DtSy5RUXeeqphP/sR2Ie0c/G7u4aP+SBNrs51vmdf2dhdr8812ZLdn+ge2CEkuzIwYonUNIrwHhVXQas9ALK1BpbA4rerkpK3QqDJ3znufnInZw5C4CM3coVbftYddxdzUt7yU5mkj0rYqB+DNo12fht2ugXhPUqcFzW/BoEh3xBWitFHs3sWljPzIr03B8HxOXJkM/6HnxNsG1S59ki3PFbb3Z3lcFUX1Kivb0dPX0IaZGgcBioDpcOw/HmMTCrTFYCx7C2JNoYgPqtoPubh7rBh5CTJFRMwzkZm+iWm34ltv1Ovnh1K63T548t7NmWOv61/THB4eGsHbtIDKZfiiKyjx00F6YYcAXsXDzIhLdj6L7kREU52sI/BY2bFyDbDYD1+lgNjeLM2ensfWJRTLiffXNz6+Qs8fepRNX/iDFUgVlrQ7CCYgnUuhiM8EUxBWCqOQiKlNkenuRXf8sKloOK5IShjc8DIHnoWka5oolmHUKNSZDVeMwai3n+InzI6ScP+SGxBcsqwmtUoVpNOB6lAVJIQo8OI5jpBx4QYIksSkKWJXtQW8mjUargcqtm6x6D0bDh645GBjoQ6Omm+3KxJcT1298RLTCQVdJxIQIYXEQiW0msDse2m1G5DQR4UJWkQJFkhmZAEUWEYQUhXwOzWYTyUQMP10oQpJlbBxeg+qNC5PN8tX3vOqVc2/sgUX04kFPUeM8CAuUJECIgAjhELLH913YloG2XQb1O0yByJQFrFIdXCQCUVIwV2ogEpEgcjCbhfMnKtNX39/56a3qXy2IFKf2X44m06t5JZPkxRhBGLL7vfibUbCV0hA0bDGCKiNcDHIGMrOKBgTzNZfZanm9an3cnBvf19J/+/6t3exy3THI5dOj/aIaXcdLyW2CHH1OTa/cKsdTIsf8D4IAgRvAD1ymymDTQaGQYzlwGL9ch1GZt/ukqQP10szeNz+uFJjJt0u7i+DOj+nT+yRbaA/xEX4knl39TDTZNUAEOcV0qBFSYzE5yDPv8/kFj3O0G7Xfv/sw94t9fPQc/KW68pKtYmzsgKCYrc2cGh0UFHW7EvOHOYl7rJSfco8fOXPY16/t333MuX6/dr+sdj12ajQqiDFVXZ1M1Qu/0sPfflHe8zWs+4Hfs9kt5/By9ixLwXKA/nYG/wR0yVv0b4H+vwj+BByiDCpt1NN/AAAAAElFTkSuQmCC'
$IconBytes = [Convert]::FromBase64String($IconBase64)
$Stream = [System.IO.MemoryStream]::new($IconBytes, 0, $IconBytes.Length)
$Menu1_Picture = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($Stream).GetHIcon()))
$Menu1.Image = $Menu1_Picture

$Menu2 = $contextmenu.Items.Add("Group Policy")
$IconBase64 = 'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAE4UlEQVRIS42VS2wbVRSG/3k448RJWs+MYxIndlxKiASqwGlRChVdIiFgwxJYsEBIiFUFEmLVJQi1XSEQIDZsqkitlLosQNCqFGhCaagIbZpHH0lM6vd7bM+bM2M7OH0lY13d8cyd851z/nvOZdBxTU1NcYcmn3qeMZIfsowxYYMRAaa1ojkzDEuDge38LKv1zt6cNYNDo7GheLnyl1458En7a3fBtWu/i6LP+7kkd7/OcaynE96+t9u2XFjnCgaWzSCZzMJW51EpJ5IW5zm0Zcn8zE/BUHT01G5x9wsPMr7ds2K5gZX5s6gVZ5FIVyo8evZvAdz49ccheU/0rBgQn93O2H3vKZxbdxKY+fmbq9XsbxfXs74rtURwagtgjgDhx6NxURZjhmEilSrCsmxYptWcKT82zc6vfUliH0Sp19Vlcfk2Ls/88dFbb7/3afv9AwGS7I9ZJKCq6nBybrcS3547ve/q4uEMR5DF5VsE+PPRgJE94bgcEAlgtwDkb8vhnQGu7ATgpxQZlKIypcaCaVJ6aHZADphi2gxCFHshUYqaETgpmtsOMEIR7Ip1GtmR4C7gDgGubgcYJkB/TG3oWFrKwDBNGIYTRTuS/1PmgCMRvzucelxcXm0B3n+4yCPRQRdAydiR45uLnAhW1gjwN0WwLcAXc3bQ0lLB9d5oee9E4XSHTrEjkX5Ewv1NDVzAtUcBpodGoqNxWe7eooEDqVHKTBKYYxnwHINcrgCWZV2YIHggyyKWbq7j8uzCdhEMxGVJaBZaugGFIkmkS0gW6mB5LwTORiG5hkuzc4gMh+HhDShKAa+9/CJ6+noxd3n50YBwVI5LkifmVK3aMHHm/G3sP3AQBu3MuqYhsbaGhYUFhCIjCPj74eEs3FhcwV+zM5icnECh2Jg2+b5jJoN1NbWauL+So1JcErlWJVs4fW4V+yYmyRALgaeU0BeapqNYqSCXL6BSUZBMJZG/exe80AOW4y3D0KtURyny6bv7AaN+ArCUIgupjIYzF1bgEwdQyGZQq9UwNj5OKlvo7fW554HzrFSqIJ3KwKSonf+FfN4tSmlXX/oeAIk8GoxLfmyKfPKHf7F33wFU00l8P30Gh195FblMFvlcDpVSGY1qGbzXB4YEZ+kwyqRSpBWPgWAAHsYubwKOHj3KHn5u79D4WHA6IKkxx0vnOnlOwciTE7h+dQ5KPounDx6CSl5yPLVnyn0ll0S3GARPRjfWN6DUFDwxPgaPR0Apm24Clpd/Ccii+KahaSHo+Te6vfpj7QZ36nwB8vCYG7ppmdBVlZqgCk3XyPsaWLWKOivA09WFEqXGaduhcBi6oaOazzUBq7cvvDsY3H2cdPQyTsXYjpTN69vT8zD5fvdkdgG0k3Rdd4XW6L5cKILz8K4e1XIFg6FBGiFnrV3KpBTX0D+Xjr3j9xWO02Hivbc/5Cos0iUeqSLPZCseKA0PYzEeRjfBkA2YVC+6ocGkB/V6zTU+NDwErdFQVm+uNE+0Lz77YEDyb3xMm1CiADqO9SaOAmIsm4NmdzFVtYdTdB9fVfv6VYMN6IYdoBbit2xLaNQb/DClRx4I3M1nMl+t31n72gU4AovU2AVB2LKrHtbtumo1JmeW+EKBF+qW4VMN7NUN9RnTNF8aCg1LPb3CkfSKcvHEiSP1/wDbXdYF8bisrQAAAABJRU5ErkJggg=='
$IconBytes = [Convert]::FromBase64String($IconBase64)
$Stream = [System.IO.MemoryStream]::new($IconBytes, 0, $IconBytes.Length)
$Menu2_Picture = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($Stream).GetHIcon()))
$Menu2.Image = $Menu2_Picture

$Menu3 = $contextmenu.Items.Add("Powershell")
$Menu3_Picture = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe")
$Menu3.Image = $Menu3_Picture

$Menu4 = $contextmenu.Items.Add("Powershell ISE")
$Menu4_Picture = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe")
$Menu4.Image = $Menu4_Picture

$Menu5 = $contextmenu.Items.Add("Command Prompt")
$Menu5_Picture = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\cmd.exe")
$Menu5.Image = $Menu5_Picture

$Menu6 = $contextmenu.Items.Add("Group Policy Update")
$Menu6_Picture = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\cmd.exe")
$Menu6.Image = $Menu6_Picture

$Menu7 = $contextmenu.Items.Add("MMC")
$Menu7_Picture = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\windows\system32\mmc.exe")
$Menu7.Image = $Menu7_Picture

$Menu8 = $contextmenu.Items.Add("Registry")
$Menu8_Picture = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\regedit.exe")
$Menu8.Image = $Menu8_Picture

$Menu9 = $contextmenu.Items.Add("Computer Management")
$Menu9_Picture = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\compmgmt.msc")
$Menu9.Image = $Menu9_Picture

$Menu10 = $contextmenu.Items.Add("Task Scheduler")
$Menu10_Picture = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\taskschd.msc")
$Menu10.Image = $Menu10_Picture

$Menu11 = $contextmenu.Items.Add("Task Manager")
$Menu11_Picture = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\Taskmgr.exe")
$Menu11.Image = $Menu11_Picture

$Menu12 = $contextmenu.Items.Add("System Properties")
$Menu12_Picture = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\sysdm.cpl")
$Menu12.Image = $Menu12_Picture

$Menu13 = $contextmenu.Items.Add("Hyper-V Manager")
$IconBase64 = 'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAADsElEQVRIS7WWaUhUURTH/++90cosAhOyLxkUZH4oUCIq1DSdNoqKSJM2s6yIyIwWbS+qL20QUaEV7djillYUki1YiFGQlablFzO1qHRa1HnzOue9d+eNadMQdOHOneXe/++c/z3nMRL+85B80R9W9X5drQsHeG+ojIr68MFjfDnHe/4ICH7VuKjFqZ3xJhRukz5UhYWEeNvTBRD6sjGnXtVSfIpO4130oqrAVwcQM7LHYPUv066UaPYpCZj9rtknbaguQ7T1G/CzHXA6gaSoPwNSLxdq9ml22GQJNkmCvyTDXtvUFeZiURJs/W6JchYavfBM/gsgO68IaGpAUUkBFIIoFE+lKiGzshZoY9EOsoMi1egHIaqvBHbRuiDGewYJUxOgUAZ96MDUjJ1A9XPsLbmBzOInRibCc3fUprAOoLkozjeAsKmCgu2vyFh3gwCeVvAHjpiFRfQMWBLvHZBd+Qp4+xoF507o98AAf1lGpg4wfRaW8MqiPPk9V1LqZN8zmLn9IMGqsSv7GLbdfmoCTDERvQC4SJyrKm2a74DyDg29yR49g8Jyywrhtx45TY587OhupZ0yyK//qeDgNncnc5kGBg3E4fOXET9iKCJWrrIA+Y+6+x0Z7r1fPrcCLV+AOWPJbBp6H1AVyVRFDztceH73PkqLizF+VBgeBYeZGZieU/RX502EP9Wxn6RgiuiXL9R4H78C7dR4XBRsXXK0BRBlyoBeZI3bomtlxmbhOfmdmxSHp1U12J+VhX1F+djMpSwqzfNuFsZagEkJcUjcsg9yfQ02XDxnAXJLTXETQoAJdeVYv2snWqjpSp+9wKWGH2ZFmVny3XBTptg9LKJHxYN2Vb/YgTKwKXM3UEelO3f5bwAVDxJj8SkgAMsOZSPU1omKIRHWHiHuJAhVlnUHHgB/vYLomcRVdOGWYRGXor6qkKufISdrLer6DYBEnu++Q6WsC5vT2Wk8AFfPsgBKbKzuPYt2AZwtdgsbECeeLJ2BS6qMQJsNezZuB8bFG4JsC68CkD7PABxudJStvZAXtWbxnO6A0wVG+hwdZ0Ajb/lszFqRji3HjqAXdfLWa/e6incSBHQmY74B8BxBx69rnxzfsCMtEZF+MqafuG6Ic2PxoENX2tq18k4X+toUBCgKNp+/6RE52SNGTwBPWODRXM1BF4/o4X0RGUnPbGOsaHZoASTuBnCWwiIPcXcne2/Lnn/d+KbpVk7ZY3tW0nSkn7xq2EcR/77bp38V/xKAOPMLKdEgN0nrvccAAAAASUVORK5CYII='
$IconBytes = [Convert]::FromBase64String($IconBase64)
$Stream = [System.IO.MemoryStream]::new($IconBytes, 0, $IconBytes.Length)
$Menu13_Picture = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($Stream).GetHIcon()))
$Menu13.Image = $Menu13_Picture

$Menu14 = $contextmenu.Items.Add("Exchange/AzureAD Connect")
$Menu14_Picture = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe")
$Menu14.Image = $Menu14_Picture

$Menu_Exit = $contextmenu.Items.Add("Exit")
$IconBase64 = 'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAApklEQVRIS2NkQAPGxsb/0cVI4Z89e5YRWT0Kh1LDYQYjW4LVAnRXEOsDmANHLcAZYlQPIpCB6PGFLkZRJMNc/PHjR447d+78xOY1qlgANfg30Dds6JYQZQEp+YOojIasiBTDycpo5FgAsgjmSIJBNCgtILmoIMUXJEcyvoIOzWLykykuS+iS0WhaVOCqDUmO5NEKBx4CpKR3fMGGMw5Amii1BD1VAQBT08IZ4lLZJwAAAABJRU5ErkJggg=='
$IconBytes = [Convert]::FromBase64String($IconBase64)
$Stream = [System.IO.MemoryStream]::new($IconBytes, 0, $IconBytes.Length)
$Menu_Exit_Picture = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::new($Stream).GetHIcon()))
$Menu_Exit.Image = $Menu_Exit_Picture

$Systray_Tool_Icon.ContextMenuStrip = $contextmenu

# Action after clicking on the systray icon 
$Systray_Tool_Icon.Add_Click({
    If ($_.Button -eq [Windows.Forms.MouseButtons]::Middle) {
        # Setup the XAML
        [xml]$xaml = @"
    <Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    ResizeMode="NoResize"
    Width="250"
    SizeToContent="Height"
    Title="Wet-Floor Information"
    Topmost="True">
    <Grid Margin="10,10,10,10" ShowGridLines="False">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Column="0" Grid.Row="1" Margin="5">Version:
        </TextBlock>
        <TextBlock Grid.Column="1" Grid.Row="1" Margin="5">$Version
        </TextBlock>
        <TextBlock Grid.Column="0" Grid.Row="2" Margin="5">Author:
        </TextBlock>
        <TextBlock Grid.Column="1" Grid.Row="2" Margin="5">James Orner
        </TextBlock>
    </Grid>
</Window>
"@
        # Create the form and set variables
        $window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
        $xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $window.FindName($_.Name) -Scope Script }

        # here's the base64 string of the image
        $base64 = "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"

        # Create a streaming image by streaming the base64 string to a bitmap streamsource
        $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
        $bitmap.BeginInit()
        $bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64)
        $bitmap.EndInit()
        $bitmap.Freeze()

        # This is the pic in the middle
        #$image.source = $bitmap

        # This is the icon in the upper left hand corner of the app
        $window.Icon = $bitmap

        # This is the toolbar icon and description
        $window.TaskbarItemInfo.Overlay = $bitmap
        $window.TaskbarItemInfo.Description = $window.Title

        # Allow input to window for TextBoxes, etc
        [System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($window)

        # Running this without $appContext and ::Run would actually cause a really poor response.
        $window.Show()

        # This makes it pop up
        $window.Activate()

        # Create an application context for it to all run within. 
        # This helps with responsiveness and threading.
        $appContext = New-Object System.Windows.Forms.ApplicationContext 
        [void][System.Windows.Forms.Application]::Run($appContext)
    }
})

# When Exit is clicked, close everything and kill the PowerShell process
$Menu_Exit.add_Click({
        $Systray_Tool_Icon.Visible = $false
        $window.Close()
        # $window_Config.Close() 
        Stop-Process $pid -ErrorAction 'SilentlyContinue'
    })

# Action after clicking on the Menu 1 - Submenu 1
$Menu1.add_Click({
        $Filename = "C:\windows\system32\dsa.msc"
        If (Test-Path -Path $FileName) {
            Start-Process $Filename
        }
        else {
            $Continue = [System.Windows.MessageBox]::Show("Do you want to install Active Directory Users and Computers?", "Active Directory Users and Computers", 'YesNo', 'Warning')
            if ($Continue -eq 'Yes') {
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 0
                Restart-Service wuauserv
                Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 1
                Restart-Service wuauserv
            }
            Start-Process $Filename
        }
    })

# Action after clicking on the Menu 1 - Submenu 2
$Menu2.add_Click({
        $Filename = "C:\windows\system32\gpmc.msc"
        If (Test-Path -Path $FileName) {
            Start-Process C:\windows\system32\mmc.exe $Filename
        }
        else {
            $Continue = [System.Windows.MessageBox]::Show("Do you want to install Group Policy Management?", "Group Policy Management", 'YesNo', 'Warning')
            if ($Continue -eq 'Yes') {
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 0
                Restart-Service wuauserv
                Add-WindowsCapability -Online -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0
                Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 1
                Restart-Service wuauserv
            }
            Start-Process C:\windows\system32\mmc.exe $Filename
        }
    })

# Action after clicking on the Menu 1 - Submenu 3
$Menu3.add_Click({ 
        Start-Process powershell.exe
    })

# Action after clicking on the Menu 1 - Submenu 4
$Menu4.add_Click({ 
        Start-Process "C:\Windows\System32\WindowsPowerShell\v1.0\powershell_ise.exe"
    })

# Action after clicking on the Menu 1 - Submenu 5
$Menu5.add_Click({ 
        Start-Process cmd.exe
    })

# Action after clicking on the Menu 1 - Submenu 6
$Menu6.add_Click({
        gpupdate /force
        If ($LastExitCode -eq "0") {
            [System.Windows.Forms.MessageBox]::Show("Computer Policy update has completed successfully.`nUser Policy update has completed successfully.", 'Update Group Policy', 'OK', 'Information')
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Could not flush the DNS Resolver Cache: Function failed during execution.", 'Update Group Policy', 'OK', 'Information')
        }

    })

# Action after clicking on the Menu 1 - Submenu 7
$Menu7.add_Click({
        Start-Process C:\windows\system32\mmc.exe
    })

# Action after clicking on the Menu 1 - Submenu 8
$Menu8.add_Click({ 
        Start-Process regedit
    })

# Action after clicking on the Menu 1 - Submenu 9
$Menu9.add_Click({ 
        Start-Process compmgmt.msc
    })

# Action after clicking on the Menu 1 - Submenu 10
$Menu10.add_Click({
        Start-Process C:\Windows\System32\taskschd.msc -ArgumentList "/s"
    })

# Action after clicking on the Menu 1 - Submenu 11
$Menu11.add_Click({ 
        Start-Process C:\Windows\System32\Taskmgr.exe
    })

# Action after clicking on the Menu 1 - Submenu 12
$Menu12.add_Click({ 
        Start-Process C:\Windows\System32\sysdm.cpl
    })

$Menu13.add_Click({ 
        $Filename = "C:\Windows\System32\virtmgmt.msc"
        If (Test-Path -Path $FileName) {
            $startParams = @{
                FilePath     = 'C:\Windows\System32\mmc.exe'
                ArgumentList = 'C:\Windows\System32\virtmgmt.msc'
                PassThru     = $true
            }
            Start-Process @startParams    
        }
        else {
            $Continue = [System.Windows.MessageBox]::Show("Do you want to install Hyper-V?", "Hyper-V", 'YesNo', 'Warning')
            if ($Continue -eq 'Yes') {
                $startParams = @{
                    FilePath     = 'powershell.exe'
                    ArgumentList = '-NoExit', '-Command', 'Enable-WindowsOptionalFeature', '-FeatureName "Microsoft-Hyper-V-All"', '-All', '-Online'
                    PassThru     = $true
                    Wait         = $true
                }
                Start-Process @startParams -Verb Runas
            }
            $startParams = @{
                FilePath     = 'C:\Windows\System32\mmc.exe'
                ArgumentList = 'C:\Windows\System32\virtmgmt.msc'
                PassThru     = $true
            }
            Start-Process @startParams
        }
    })

# Action after clicking on the Menu 1 - Submenu 13
$Menu14.add_Click({
        $Filename = "C:\Program Files\WindowsPowerShell\Modules\ExchangeOnlineManagement"
        If (Test-Path -Path $FileName) {
            Import-Module -Name ExchangeOnlineManagement
        }
        else {
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
            Import-Module -Name ExchangeOnlineManagement
        }
        $Filename = "C:\Program Files\WindowsPowerShell\Modules\AzureAD"
        If (Test-Path -Path $FileName) {
            Import-Module -Name AzureAD
        }
        else {
            Install-Module -Name AzureAD -Scope CurrentUser -Force
            Import-Module -Name AzureAD
        }
        $startParams = @{
            FilePath     = 'powershell.exe'
            ArgumentList = '-NoExit', '-Command', { Clear-Host; Write-Warning 'Connecting to Exchange Online.'; Connect-ExchangeOnline -ErrorAction SilentlyContinue; Clear-Host; Write-Warning 'Connecting to AzureAD.'; Connect-AzureAD -ErrorAction SilentlyContinue; Clear-Host }
            PassThru     = $true
        }
        Start-Process @startParams
    })

# Make PowerShell Disappear
$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$asyncwindow = Add-Type -MemberDefinition $windowcode -Name Win32ShowWindowAsync -Namespace Win32Functions -PassThru
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
 
# Force garbage collection just to start slightly lower RAM usage.
[void][System.GC]::Collect()
 
# Create an application context for it to all run within.
# This helps with responsiveness, especially when clicking Exit.
$appContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($appContext)