Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "Textaenderungs-Tool"
$form.Width = 700
$form.Height = 400

$button = New-Object System.Windows.Forms.Button
$button.Text = "Click Me"
$button.Width = 100
$button.Height = 30
$button.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Button clicked!")
})

$form.Controls.Add($button)

$form.ShowDialog()