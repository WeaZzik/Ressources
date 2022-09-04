Add-Type -assembly System.Windows.Forms					#Add a Library

$Main = New-Object System.Windows.Forms.Form			#Create the Window
$Main.Text ='Bel � Computer Preparation � v1.1' 		#Window's Title
$Main.Width = 600										#Window's Width
$Main.Height = 400										#Window's Height
$Main.AutoSize = $true									#Resize Window if the screen is too small
$Main.StartPosition = 'CenterScreen'                    #Choose Window's Natural Position
$Main.MaximizeBox = $False                              #Disable Maximize Button
$Main.MinimizeBox = $False                              #Disable Minimize Button
$Main.FormBorderStyle = 'FixedDialog'                   #Disable Resize Possibility

#Configuration Profile Text
$Profile_Label = New-Object System.Windows.Forms.Label
$Profile_Label.Location = New-Object System.Drawing.Point(20,30)
$Profile_Label.Size = New-Object System.Drawing.Size(80,20)
$Profile_Label.Text = 'Config Profile:'
$Main.Controls.Add($Profile_Label)

#Username Text
$Username_Label = New-Object System.Windows.Forms.Label
$Username_Label.Location = New-Object System.Drawing.Point(30,70)
$Username_Label.Size = New-Object System.Drawing.Size(70,20)
$Username_Label.Text = 'Username:'
$Main.Controls.Add($Username_Label)

#Mail Text
$Mail_Label = New-Object System.Windows.Forms.Label
$Mail_Label.Location = New-Object System.Drawing.Point(60,100)
$Mail_Label.Size = New-Object System.Drawing.Size(40,20)
$Mail_Label.Text = 'Mail:'
$Main.Controls.Add($Mail_Label)

#HPsupport Text
$HPsupport_Label = New-Object System.Windows.Forms.Label
$HPsupport_Label.Location = New-Object System.Drawing.Point(40,133)
$HPsupport_Label.Size = New-Object System.Drawing.Size(70,20)
$HPsupport_Label.Text = 'HP Support'
$Main.Controls.Add($HPsupport_Label)

#Epson iProjection Text
$Epson1_Label = New-Object System.Windows.Forms.Label
$Epson1_Label.Location = New-Object System.Drawing.Point(10,153)
$Epson1_Label.Size = New-Object System.Drawing.Size(100,20)
$Epson1_Label.Text = 'Epson iProjection'
$Main.Controls.Add($Epson1_Label)

#Epson EasyMP Text
$Epson2_Label = New-Object System.Windows.Forms.Label
$Epson2_Label.Location = New-Object System.Drawing.Point(20,173)
$Epson2_Label.Size = New-Object System.Drawing.Size(90,20)
$Epson2_Label.Text = 'Epson EasyMP'
$Main.Controls.Add($Epson2_Label)

#KeyPass Text
$KeyPass_Label = New-Object System.Windows.Forms.Label
$KeyPass_Label.Location = New-Object System.Drawing.Point(50,193)
$KeyPass_Label.Size = New-Object System.Drawing.Size(50,20)
$KeyPass_Label.Text = 'KeePass'
$Main.Controls.Add($KeyPass_Label)

#Windows Update Text
$WindowsUpdate_Label = New-Object System.Windows.Forms.Label
$WindowsUpdate_Label.Location = New-Object System.Drawing.Point(10,213)
$WindowsUpdate_Label.Size = New-Object System.Drawing.Size(90,20)
$WindowsUpdate_Label.Text = 'Windows Update'
$Main.Controls.Add($WindowsUpdate_Label)

#Username Box
$Username_TextBox = New-Object System.Windows.Forms.TextBox
$Username_TextBox.Location = New-Object System.Drawing.Point(110,67)
$Username_TextBox.Size = New-Object System.Drawing.Size(100,20)
$Main.Controls.Add($Username_TextBox)

#Mail Box
$Mail_TextBox = New-Object System.Windows.Forms.TextBox
$Mail_TextBox.Location = New-Object System.Drawing.Point(110,97)
$Mail_TextBox.Size = New-Object System.Drawing.Size(200,20)
$Mail_Click = {$Mail_TextBox.Text = ($Username_TextBox.Text + '@groupe-bel.com')}
$Mail_TextBox.Add_GotFocus($Mail_Click)
$Main.Controls.Add($Mail_textBox)


#HPsupport CheckBox
$HPsupport_CheckBox = New-Object System.Windows.Forms.CheckBox
$HPsupport_CheckBox.Location = New-Object System.Drawing.Point(110,130)
$HPsupport_CheckBox.Size = New-Object System.Drawing.Size(20,20)
$HPsupport_CheckBox.Checked =$true
$Main.Controls.Add($HPsupport_CheckBox)

#Epson iProjection CheckBox
$Epson1_CheckBox = New-Object System.Windows.Forms.CheckBox
$Epson1_CheckBox.Location = New-Object System.Drawing.Point(110,150)
$Epson1_CheckBox.Size = New-Object System.Drawing.Size(20,20)
$Epson1_CheckBox.Checked =$true
$Main.Controls.Add($Epson1_CheckBox)

#Epson EasyMP CheckBox
$Epson2_CheckBox = New-Object System.Windows.Forms.CheckBox
$Epson2_CheckBox.Location = New-Object System.Drawing.Point(110,170)
$Epson2_CheckBox.Size = New-Object System.Drawing.Size(20,20)
$Epson2_CheckBox.Checked =$true
$Main.Controls.Add($Epson2_CheckBox)

#KeePass CheckBox
$KeePass_CheckBox = New-Object System.Windows.Forms.CheckBox
$KeePass_CheckBox.Location = New-Object System.Drawing.Point(110,190)
$KeePass_CheckBox.Size = New-Object System.Drawing.Size(20,20)
$Main.Controls.Add($KeePass_CheckBox)

#Windows Update CheckBox
$WindowsUpdate_CheckBox = New-Object System.Windows.Forms.CheckBox
$WindowsUpdate_CheckBox.Location = New-Object System.Drawing.Point(110,210)
$WindowsUpdate_CheckBox.Size = New-Object System.Drawing.Size(20,20)
$WindowsUpdate_CheckBox.Checked =$true
$Main.Controls.Add($WindowsUpdate_CheckBox)

#Configuration Profile ComboBox
$Profile_ComboBox = New-Object System.Windows.Forms.ComboBox
$Profile_ComboBox.Location = New-Object System.Drawing.Point(110,27)
$Profile_ComboBox.Size = New-Object System.Drawing.Size(150,20)
$Profile_ComboBox.DropDownStyle = 'DropDownList'
$Profile_ComboBox.Items.Add('Generic Laptop')
$Profile_ComboBox.Items.Add('Generic Desktop')
$Main.Controls.Add($Profile_ComboBox)






# ---------------------------------------------------------------- ENDSCREEN -------------------------------------------------------------------


$EndScreen = New-Object System.Windows.Forms.Form			#Create the Window
$EndScreen.Text ='Bel � Computer Preparation � v1.1' 		#Window's Title
$EndScreen.Width = 300										#Window's Width
$EndScreen.Height = 200										#Window's Height
$EndScreen.AutoSize = $true									#Resize Window if the screen is too small
$EndScreen.StartPosition = 'CenterScreen'                    #Choose Window's Natural Position
$EndScreen.MaximizeBox = $False                              #Disable Maximize Button
$EndScreen.MinimizeBox = $False                              #Disable Minimize Button
$EndScreen.ControlBox = $False
$EndScreen.FormBorderStyle = 'FixedDialog'                   #Disable Resize Possibility

#End ListView
$ListView = New-Object System.Windows.Forms.ListView
$ListView.Location = New-Object System.Drawing.Point(10,20)
$ListView.Size = New-Object System.Drawing.Size(275,150)
$ListView.View = 'Details'
$ListView.Columns.Add('Application',175)
$ListView.Columns.Add('Status',95)
$EndScreen.Controls.Add($ListView)

#ProgressBar
$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(10,400)
$ProgressBar.Size = New-Object System.Drawing.Size(470,30)
$Main.Controls.Add($ProgressBar)

#Support Button
$Support_Button = New-Object System.Windows.Forms.Button
$Support_Button.Location = New-Object System.Drawing.Point(500,10)
$Support_Button.Size = New-Object System.Drawing.Size(80,20)
$Support_Button.Text = 'Support'
$Support_Button_Click = {[System.Windows.Forms.MessageBox]::Show('luka.laurentoff@gmail.com')}
$Support_Button.Add_Click($Support_Button_Click)
$Main.Controls.Add($Support_Button)

#Start Button
$Start_Button = New-Object System.Windows.Forms.Button
$Start_Button.Location = New-Object System.Drawing.Point(485,400)
$Start_Button.Size = New-Object System.Drawing.Size(100,30)
$Start_Button.Text = 'START'

$Start_Button_Click = {
    Start-Process companyportal:
    $ProgressBar.Value = '100'
    If ($HPsupport_CheckBox.Checked -eq $true)
    {
        $ListView_Item1 =  $ListView.Items.Add('HP Support')
        $Temp = Test-Path -Path 'C:\Program Files'
        If ($Temp -eq $true)
        {
        $ListView_Item1.SubItems.Add('Installed')
        & 'C:\Program Files\Notepad++\notepad++.exe'
        }
        If ($Temp -eq $false){$ListView_Item1.SubItems.Add('Not Found')}
    }
    If ($Epson1_CheckBox.Checked -eq $true)
    {
        $ListView_Item2 =  $ListView.Items.Add('Epson iProjection')
        $Temp = Test-Path -Path 'C:\Program Files\Epson1'
        If ($Temp -eq $true){$ListView_Item2.SubItems.Add('Installed')}
        If ($Temp -eq $false){$ListView_Item2.SubItems.Add('Not Found')}
    }
    If ($Epson2_CheckBox.Checked -eq $true)
    {
        $ListView_Item3 =  $ListView.Items.Add('Epson EasyMP')
        $Temp = Test-Path -Path 'C:\Program Files\Epson2'
        If ($Temp -eq $true){$ListView_Item3.SubItems.Add('Installed')}
        If ($Temp -eq $false){$ListView_Item3.SubItems.Add('Not Found')}
    }
    If ($KeePass_CheckBox.Checked -eq $true)
    {
        $ListView_Item4 =  $ListView.Items.Add('KeePass')
        $Temp = Test-Path -Path 'C:\Program Files\KeePass'
        If ($Temp -eq $true){$ListView_Item4.SubItems.Add('Installed')}
        If ($Temp -eq $false){$ListView_Item4.SubItems.Add('Not Found')}
    }
    If ($WindowsUpdate_CheckBox.Checked -eq $true){C:\Windows\System32\control.exe /name Microsoft.WindowsUpdate}
    $EndScreen.ShowDialog()
}
$Start_Button.Add_Click($Start_Button_Click)
$Main.Controls.Add($Start_Button)

#EndScreenOK Button
$EndScreenOK_Button = New-Object System.Windows.Forms.Button
$EndScreenOK_Button.Location = New-Object System.Drawing.Point(110,200)
$EndScreenOK_Button.Size = New-Object System.Drawing.Size(80,30)
$EndScreenOK_Button.Text = 'OK'
$EndScreenOK_Click = 
{
    $EndScreen.Close()              #Close the End Window
    $Main.Close()                   #Close the Main Window
}
$EndScreenOK_Button.Add_Click($EndScreenOK_Click)
$EndScreen.Controls.Add($EndScreenOK_Button)

$Main.ShowDialog()										#Show the Main Window