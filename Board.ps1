Add-Type -assembly System.Windows.Forms					#Add a Library



$Main = New-Object System.Windows.Forms.Form			#Create the Window
$Main.Text ='Bel · IT Board' 		#Window's Title
$Main.Width = 1200										#Window's Width
$Main.Height = 800										#Window's Height
$Main.AutoSize = $true									#Resize Window if the screen is too small
$Main.StartPosition = 'CenterScreen'                    #Choose Window's Natural Position
$Main.MaximizeBox = $False                              #Disable Maximize Button
$Main.MinimizeBox = $False                              #Disable Minimize Button
$Main.FormBorderStyle = 'FixedDialog'                   #Disable Resize Possibility

#---------------------------------- NEW WINDOW ------------------------------------

$NewWindow = New-Object System.Windows.Forms.Form			#Create the Window
$NewWindow.Text ='Bel · New Item' 		                    #Window's Title
$NewWindow.Width = 450										#Window's Width
$NewWindow.Height = 50										#Window's Height
$NewWindow.AutoSize = $true									#Resize Window if the screen is too small
$NewWindow.StartPosition = 'CenterScreen'                    #Choose Window's Natural Position
$NewWindow.MaximizeBox = $False                              #Disable Maximize Button
$NewWindow.MinimizeBox = $False                              #Disable Minimize Button
$NewWindow.FormBorderStyle = 'FixedDialog'                   #Disable Resize Possibility

# ------------------------------------ MENU ---------------------------------------

$Menu = New-Object System.Windows.Forms.TabControl
$Menu.Location = New-Object System.Drawing.Point(50,30)
$Menu.Size = New-Object System.Drawing.Size(1100,650)
$Menu.Name = 'MenuPage'
$Menu.SelectedIndex = 0


#------------------------------------- PAGES ---------------------------------------

$Page1 = New-Object System.Windows.Forms.TabPage
$Page1.Padding = '5,5,5,5'
$Page1.TabIndex = 1
$Page1.Text = 'Logistique'
$Page1.UseVisualStyleBackColor = $True

# ----------------------------------- LISTS -----------------------------------------

$List = New-Object System.Windows.Forms.ListView
$List.View = 'Details'
$List.Size = '1090,623'
$List.Columns.Add('Name',150)
$List.Columns.Add('Number',150)
$List.Columns.Add('Status',100)
$List.FullRowSelect = $True
$List.MultiSelect = $Fals
$Main.Controls.Add($List)

#------------------------------------ BUTTONS ---------------------------------------

#SAVE BUTTON
$Save = New-Object System.Windows.Forms.Button
$Save.Location = New-Object System.Drawing.Point(1000,700)
$Save.Size = New-Object System.Drawing.Size(100,30)
$Save.text = 'SAVE'
$SaveClick = {
    foreach($item in $List){                                                  #Export
        $item | Export-CSV -Path "output.csv" -NoTypeInformation              #Export
        }                                                                     #Export
    }                                                                         #Export
$Save.Add_Click($SaveClick)

# NEW BUTTON
$New = New-Object System.Windows.Forms.Button
$New.Location = New-Object System.Drawing.Point(50,700)
$New.Size = New-Object System.Drawing.Size(100,30)
$New.text = 'NEW'
$NewClick = {
    $NewWindow.Controls.Add($New_Box1)
    $NewWindow.Controls.Add($New_Box2)
    $NewWindow.Controls.Add($New_Box3)
    $NewWindow.Controls.Add($Add)
    $NewWindow.ShowDialog()
    }
$New.Add_Click($NewClick)

#EDIT BUTTON
$Edit = New-Object System.Windows.Forms.Button
$Edit.Location = New-Object System.Drawing.Point(170,700)
$Edit.Size = New-Object System.Drawing.Size(100,30)
$Edit.text = 'EDIT'
$EditClick = {
    foreach($item in $List.SelectedItems){
        If ($item.index -ne $null){
            $New_Box1.text = $item.Text
            $New_Box2.text = $item.SubItems[1].Text
            $New_Box3.text = $item.SubItems[2].Text
            $List.Items.RemoveAt($item.Index)
            $NewWindow.Controls.Add($New_Box1)
            $NewWindow.Controls.Add($New_Box2)
            $NewWindow.Controls.Add($New_Box3)
            $NewWindow.Controls.Add($Confirm)
            $NewWindow.ControlBox = $False
            $NewWindow.ShowDialog()
            }
        }
    }
$Edit.Add_Click($EditClick)

#DELETE BUTTON
$Delete = New-Object System.Windows.Forms.Button
$Delete.Location = New-Object System.Drawing.Point(290,700)
$Delete.Size = New-Object System.Drawing.Size(100,30)
$Delete.text = 'DELETE'
$DeleteClick = {
    foreach($item in $List.SelectedItems){$List.Items.RemoveAt($item.Index)}
    }
$Delete.Add_Click($DeleteClick)

# ADD BUTTON
$Add = New-Object System.Windows.Forms.Button
$Add.Location = New-Object System.Drawing.Point(340,20)
$Add.Size = New-Object System.Drawing.Size(100,30)
$Add.text = 'ADD'
$AddClick = {
    $Item = $List.Items.Add($New_Box1.Text)
    $Item.SubItems.Add($New_Box2.Text)
    $Item.SubItems.Add($New_Box3.Text)
    $New_Box1.Text = ''
    $New_Box2.Text = ''
    $New_Box3.Text = ''
    $NewWindow.Close()
    }
$Add.Add_Click($AddClick)

# CONFIRM BUTTON
$Confirm = New-Object System.Windows.Forms.Button
$Confirm.Location = New-Object System.Drawing.Point(340,20)
$Confirm.Size = New-Object System.Drawing.Size(100,30)
$Confirm.text = 'CONFIRM'
$ConfirmClick = {
    $Item = $List.Items.Add($New_Box1.Text)
    $Item.SubItems.Add($New_Box2.Text)
    $Item.SubItems.Add($New_Box3.Text)
    $New_Box1.Text = ''
    $New_Box2.Text = ''
    $New_Box3.Text = ''
    $NewWindow.Close()
    }
$Confirm.Add_Click($ConfirmClick)

# ---------------------------------- TEXTBOX ---------------------------------------

$New_Box1 = New-Object System.Windows.Forms.TextBox
$New_Box1.Location = New-Object System.Drawing.Point(10,25)
$New_Box1.Size = New-Object System.Drawing.Size(100,20)

$New_Box2 = New-Object System.Windows.Forms.TextBox
$New_Box2.Location = New-Object System.Drawing.Point(120,25)
$New_Box2.Size = New-Object System.Drawing.Size(100,20)

$New_Box3 = New-Object System.Windows.Forms.TextBox
$New_Box3.Location = New-Object System.Drawing.Point(230,25)
$New_Box3.Size = New-Object System.Drawing.Size(100,20)

#----------------------------------- EXECUTION -------------------------------------

$Page1.Controls.AddRange(@($List))
$Menu.Controls.AddRange(@($Page1))
$Main.Controls.Add($Menu)
$Main.Controls.Add($New)
$Main.Controls.Add($Edit)
$Main.Controls.Add($Delete)
$Main.Controls.Add($Save)
$Main.ShowDialog()