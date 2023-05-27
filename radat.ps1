Add-Type -AssemblyName System.Windows.Forms

# Create a new form
$form = New-Object System.Windows.Forms.Form
$form.Text = "MEMCM Remove and Deploy Applications Tool v1.0.3 - May 2023"
$form.Size = New-Object System.Drawing.Size(600, 315)
$form.StartPosition = "CenterScreen"

# Create controls
$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(10, 20)
$label1.Size = New-Object System.Drawing.Size(200, 40)
$label1.Text = "Application Name to Remove (Use RZ Code for RuckZuck App):"
$form.Controls.Add($label1)

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(220, 20)
$textBox1.Size = New-Object System.Drawing.Size(350, 40)
$form.Controls.Add($textBox1)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10, 60)
$label2.Size = New-Object System.Drawing.Size(200, 40)
$label2.Text = "Application Name to Deploy:"
$form.Controls.Add($label2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(220, 60)
$textBox2.Size = New-Object System.Drawing.Size(350, 40)
$form.Controls.Add($textBox2)

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(10, 100)
$label3.Size = New-Object System.Drawing.Size(200, 20)
$label3.Text = "Install Collection Name:"
$form.Controls.Add($label3)

$installDropDownBox = New-Object System.Windows.Forms.ComboBox
$installDropDownBox.Location = New-Object System.Drawing.Point(220, 100)
$installDropDownBox.Size = New-Object System.Drawing.Size(350, 20)
$installDropDownBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$installDropDownBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems

# Query to populate the install collection drop-down box
$installCollectionQuery = Get-CMCollection | Select-Object -ExpandProperty Name

# Add the query results to the install collection drop-down box
$installDropDownBox.Items.AddRange($installCollectionQuery)

$form.Controls.Add($installDropDownBox)

$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(10, 140)
$label4.Size = New-Object System.Drawing.Size(200, 20)
$label4.Text = "Upgrade Collection Name:"
$form.Controls.Add($label4)

$upgradeDropDownBox = New-Object System.Windows.Forms.ComboBox
$upgradeDropDownBox.Location = New-Object System.Drawing.Point(220, 140)
$upgradeDropDownBox.Size = New-Object System.Drawing.Size(350, 20)
$upgradeDropDownBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::SuggestAppend
$upgradeDropDownBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::ListItems

# Query to populate the upgrade collection drop-down box
$upgradeCollectionQuery = Get-CMCollection | Select-Object -ExpandProperty Name

# Add the query results to the upgrade collection drop-down box
$upgradeDropDownBox.Items.AddRange($upgradeCollectionQuery)

$form.Controls.Add($upgradeDropDownBox)

$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(220, 180)
$button.Size = New-Object System.Drawing.Size(75, 23)
$button.Text = "Execute"
$button.Add_Click({
    $AppRemoveDisplayName = '*' + $textBox1.Text + '*'
    $AppDeployDisplayName = $textBox2.Text + '*'
    $InstallCollectionName = $installDropDownBox.Text
    $UpgradeCollectionName = $upgradeDropDownBox.Text

    # Your PowerShell script code here
    $totalSteps = 5
    for ($step = 1; $step -le $totalSteps; $step++) {
        # Update status label
        $labelStatus.Text = "Executing Step $step of $totalSteps..."

        # Perform the specific operations for each step
        switch ($step) {
            1 {
                (Get-CMApplication | Where-Object { ($_.LocalizedDisplayName -like $AppRemoveDisplayName) } | Select-Object LocalizedDisplayName | Sort-Object) |
                    ForEach-Object {
                        Remove-CMApplicationDeployment -CollectionName $InstallCollectionName -Name "$($_.LocalizedDisplayName)" -Force
                    } -Confirm:$False
            }
            2 {
                (Get-CMApplication | Where-Object { ($_.LocalizedDisplayName -like $AppRemoveDisplayName) } | Select-Object LocalizedDisplayName | Sort-Object) |
                    ForEach-Object {
                        Remove-CMApplicationDeployment -CollectionName $UpgradeCollectionName -Name "$($_.LocalizedDisplayName)" -Force
                    } -Confirm:$False
            }
            3 {
                Remove-CMApplication -Name $AppRemoveDisplayName -Confirm:$False -Force
            }
            4 {
                (Get-CMApplication | Where-Object { ($_.LocalizedDisplayName -like $AppDeployDisplayName) } | Select-Object LocalizedDisplayName | Sort-Object) |
                    ForEach-Object {
                        New-CMApplicationDeployment -CollectionName $InstallCollectionName -Name "$($_.LocalizedDisplayName)" -DeployAction Install -DeployPurpose Required -UserNotification DisplaySoftwareCenterOnly -EnableMomAlert $False -RaiseMomAlertsOnFailure $False -PersistOnWriteFilterDevice $True
                    } -Confirm:$False
            }
            5 {
                (Get-CMApplication | Where-Object { ($_.LocalizedDisplayName -like $AppDeployDisplayName) } | Select-Object LocalizedDisplayName | Sort-Object) |
                    ForEach-Object {
                        New-CMApplicationDeployment -CollectionName $UpgradeCollectionName -Name "$($_.LocalizedDisplayName)" -DeployAction Install -DeployPurpose Required -UserNotification DisplaySoftwareCenterOnly -EnableMomAlert $False -RaiseMomAlertsOnFailure $False -PersistOnWriteFilterDevice $True
                    } -Confirm:$False
            }
        }

        # Update progress bar
        $progress = ($step / $totalSteps) * 100
        $progressBar.Value = $progress

        # Sleep to simulate work being done
        Start-Sleep -Milliseconds 500
    }

    # Reset status label and progress bar
    $labelStatus.Text = "Execution complete."
    $progressBar.Value = 0
})

$form.Controls.Add($button)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 220)
$progressBar.Size = New-Object System.Drawing.Size(570, 23)
$progressBar.Style = "Continuous"
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$form.Controls.Add($progressBar)

# Status Label
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(10, 250)
$labelStatus.Size = New-Object System.Drawing.Size(570, 20)
$form.Controls.Add($labelStatus)

# Show the form
$form.ShowDialog()
