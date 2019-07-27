
function GeneratePopUp {
#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#endregion

#region Generated Form Objects
$form1 = New-Object System.Windows.Forms.Form
$label1 = New-Object System.Windows.Forms.Label
$PopUpOk = New-Object System.Windows.Forms.Button
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

$B_PopUpOK= 
{
#TODO: Place custom script here
	$form1.close()
}
$OnLoadForm_StateCorrection=
{
	$form1.WindowState = $InitialFormWindowState
}
#region Generated Form Code
$form1.Name = 'form1'
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 246
$System_Drawing_Size.Height = 95
$form1.ClientSize = $System_Drawing_Size
$form1.FormBorderStyle = 5

$label1.TabIndex = 1
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 203
$System_Drawing_Size.Height = 53
$label1.Size = $System_Drawing_Size
$label1.Text = 'NEM végezted el a beállítást, Vagy rossz oszlopban tetted meg! Kérlek ellenõrizd!'
$label1.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",9,0,3,0)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 22
$System_Drawing_Point.Y = 4
$label1.Location = $System_Drawing_Point
$label1.DataBindings.DefaultDataSourceUpdateMode = 0
$label1.Name = 'label1'

$form1.Controls.Add($label1)

$PopUpOk.TabIndex = 0
$PopUpOk.Name = 'PopUpOk'
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 23
$PopUpOk.Size = $System_Drawing_Size
$PopUpOk.UseVisualStyleBackColor = $True

$PopUpOk.Text = 'OK'

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 76
$System_Drawing_Point.Y = 60
$PopUpOk.Location = $System_Drawing_Point
$PopUpOk.DataBindings.DefaultDataSourceUpdateMode = 0
$PopUpOk.add_Click($B_PopUpOK)

$form1.Controls.Add($PopUpOk)

#endregion Generated Form Code
$InitialFormWindowState = $form1.WindowState
$form1.add_Load($OnLoadForm_StateCorrection)
$form1.ShowDialog()| Out-Null

} 

GeneratePopUp
