function Generate-Password {
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
        return
    }

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
    $randomNumber1 = Get-Random -Minimum 0 -Maximum 10
    $randomNumber2 = Get-Random -Minimum 0 -Maximum 10
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

    # Return the generated password
    return $password
}

# Call the function to generate a password
$password = Generate-Password
