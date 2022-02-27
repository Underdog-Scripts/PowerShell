# Load the System.Windows.Forms assembly into PowerShell
Add-Type -AssemblyName System.Windows.Forms

# Create a new Form object and assign to the variable $Form
$Form = New-Object System.Windows.Forms.Form

# Add description to the Form's titlebar
$Form.Text = "The Titlebar - Form"

# Create a new Label object
$Label = New-Object System.Windows.Forms.Label

# Add some text to label
$Label.Text = "Some random text to display"

# Use AutoSize to guarantee label is correct size to contain text we add
$Label.AutoSize = $true

# Add Label to form
$Form.Controls.Add($Label)

#Initialize Form so it can be seen
$Form.showDialog()