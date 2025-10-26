Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Password Generator'
$form.Size = New-Object System.Drawing.Size(500, 250)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog  # Prevent resizing

# Create the description label
$descriptionLabel = New-Object System.Windows.Forms.Label
$descriptionLabel.Text = "Generate a secure 14 character password - with capitals, symbols, and numbers in an easy to read format"
$descriptionLabel.Size = New-Object System.Drawing.Size(450, 40)
$descriptionLabel.Location = New-Object System.Drawing.Point(25, 20)
$form.Controls.Add($descriptionLabel)

# Create the button to generate password
$button = New-Object System.Windows.Forms.Button
$button.Text = 'Generate'
$button.Size = New-Object System.Drawing.Size(200, 40)
$button.Location = New-Object System.Drawing.Point(150, 70)
$form.Controls.Add($button)

# Create the text box to display the generated password with larger text
$passwordBox = New-Object System.Windows.Forms.TextBox
$passwordBox.Size = New-Object System.Drawing.Size(260, 30)
$passwordBox.Location = New-Object System.Drawing.Point(70, 120)
$passwordBox.ReadOnly = $true  # Prevents user input
$passwordBox.Enabled = $false  # Ensures the user cannot interact with the text box
$passwordBox.Font = New-Object System.Drawing.Font('Arial', 14)  # Increase text size
$form.Controls.Add($passwordBox)

# Create the "Copy" button
$copyButton = New-Object System.Windows.Forms.Button
$copyButton.Text = 'Copy'
$copyButton.Size = New-Object System.Drawing.Size(100, 30)
$copyButton.Location = New-Object System.Drawing.Point(350, 120)
$form.Controls.Add($copyButton)

# Define file paths for word lists
$4letter = "v:\scripts\saved scripts\files\4letter.csv"
$5letter = "v:\scripts\saved scripts\files\5letter.csv"
$6letter = "v:\scripts\saved scripts\files\6letter.csv"
$7letter = "v:\scripts\saved scripts\files\7letter.csv"

# Import word lists and clean up the words (trim spaces)
$list4 = (import-csv $4letter | ForEach-Object { $_.Word.Trim() })
$list5 = (import-csv $5letter | ForEach-Object { $_.Word.Trim() })
$list6 = (import-csv $6letter | ForEach-Object { $_.Word.Trim() })
$list7 = (import-csv $7letter | ForEach-Object { $_.Word.Trim() })

# Validate that all lists contain words
if ($list4.Count -eq 0 -or $list5.Count -eq 0 -or $list6.Count -eq 0 -or $list7.Count -eq 0) {
    Write-Output "Error: One or more CSV files are empty or missing."
    exit
}

# Event handler for the "Generate" button click
$button.Add_Click({
    # Begin password generation
    Write-Output "Generating secure 14 character password..."

    # Randomly choose between the 4+7 combination or 5+6 combination
    $randomCombination = Get-Random -Minimum 1 -Maximum 3

    if ($randomCombination -eq 1) {
        # 4+7 Combination
        $word1 = $list4 | Get-Random
        $word2 = $list7 | Get-Random
    } else {
        # 5+6 Combination
        $word1 = $list5 | Get-Random
        $word2 = $list6 | Get-Random
    }

    # Capitalize the first letter of both words
    $word1 = $word1.Substring(0, 1).ToUpper() + $word1.Substring(1).ToLower()
    $word2 = $word2.Substring(0, 1).ToUpper() + $word2.Substring(1).ToLower()

    # Combine the two words
    $combinedWords = "$word1$word2"

    # Generate 2 random numbers (10-99)
    $randomNumber1 = Get-Random -Minimum 10 -Maximum 100
    $randomNumber2 = Get-Random -Minimum 10 -Maximum 100
    $randomNumbers = "$randomNumber1$randomNumber2"

    # Choose a random symbol from the given set
    $symbols = "!@#$%^&*-+!~_?"
    $randomSymbol = $symbols[(Get-Random -Minimum 0 -Maximum $symbols.Length)]

    # Randomly determine whether to place the symbol before or after the numbers (50/50 chance)
    $symbolPlacement = Get-Random -Minimum 1 -Maximum 3

    if ($symbolPlacement -eq 1) {
        # Place the symbol before the numbers
        $password = "$combinedWords$randomSymbol$randomNumbers"
    } else {
        # Place the symbol after the numbers
        $password = "$combinedWords$randomNumbers$randomSymbol"
    }

    # Display the generated password in the textbox
    $passwordBox.Text = $password
})

# Event handler for the "Copy" button click
$copyButton.Add_Click({
    # Copy the password to clipboard
    Set-Clipboard -Value $passwordBox.Text

    # Show the pop-up message
    [System.Windows.Forms.MessageBox]::Show("Password has been copied to the clipboard!", "Password Copied")
})

# Show the form
$form.ShowDialog()
