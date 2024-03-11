Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Preview a Project from Cricut Design Space'
$form.Size = New-Object System.Drawing.Size(500, 500)
$form.StartPosition = 'CenterScreen'

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(480, 60)
$label.Text = 'This is a list of your locally saved Cricut Design Space projects. If the project you are looking for is not here, open it in Cricut Design Space, and then click "Save" > "Save As", name the project accordingly, and re-run this program to see it here.'
$form.Controls.Add($label)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10, 80)
$label2.Size = New-Object System.Drawing.Size(490, 20)
$label2.Text = 'Select a project and then click "Open Preview" to open an image of it.'
$form.Controls.Add($label2)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(20, 110)
$listBox.Size = New-Object System.Drawing.Size(440, 300)

$openPreviewButton = New-Object System.Windows.Forms.Button
$openPreviewButton.Location = New-Object System.Drawing.Point(20, 420)
$openPreviewButton.Size = New-Object System.Drawing.Size(100, 23)
$openPreviewButton.Text = 'Open Preview'
$form.Controls.Add($openPreviewButton)

$closeButton = New-Object System.Windows.Forms.Button
$closeButton.Location = New-Object System.Drawing.Point(130, 420)
$closeButton.Size = New-Object System.Drawing.Size(75, 23)
$closeButton.Text = 'Close'
$closeButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $closeButton
$form.Controls.Add($closeButton)

$listBoxAssociatedPaths = New-Object System.Collections.ArrayList
Get-ChildItem $env:HOMEPATH\.cricut-design-space\LocalData -Directory | ForEach-Object {
	$canvasPath = "$env:HOMEPATH\.cricut-design-space\LocalData\$_\Canvas"
  Get-ChildItem $canvasPath\*\Details | ForEach-Object {
    $details = $(Get-Content $_ | ConvertFrom-Json)
    if (Get-Member -InputObject $details -Name "name" -MemberType Properties) {
      $listBoxAssociatedPaths.Add("$canvasPath\$($details.canvasID)")
      $listBox.Items.Add($details.name) | Out-Null
    }
  }
}

$openPreviewButton.Add_Click({
	$index = $listBox.SelectedIndex
	Invoke-Item "$($listBoxAssociatedPaths[$index])\Preview.png" | Out-Null
})

$form.Controls.Add($listBox)
$form.ShowDialog()
