#Author: mas
#Date: 29.11.21
function Get-UIChoicePrompt {
  <#
  .SYNOPSIS
      Prompts a closed question with the ability to decide.
  
  .DESCRIPTION
      This function prompts a closed question with the ability to decide by providing yes or no choices. It returns $false for yes and $true for no. If nothing is supplied, the default result is no. If wrong inputs are supplied, it repeats until a valid input or no input is given.
  
  .PARAMETER Title
      [String] Specifies the title of the question.
  
  .PARAMETER Content
      [String] Specifies the closed yes or no question.
  
  .OUTPUTS
      [bool]
      Returns $false for yes and $true for no.
  
  .EXAMPLE
      Get-UIChoicePrompt -Title "Test" -Content "This or that"
  
  .NOTES 
      Author: Marcel Schubert
      LastEdit: 30.11.2021
  #>
  [CmdletBinding()]
  param(
      [string]$Title = "Title",
      [string]$Content = "Content"
  )

  $Choices = @("&Yes", "&No")

  if ($Choices.Count -ne 2) {
      Throw "Get-UIChoicePrompt: Exactly 2 choices should be supplied."
  }

  $validInput = $false
  $decision = $null

  do {
      $decision = $Host.UI.PromptForChoice($Title, $Content, $Choices, 1)
      
      switch ($decision) {
          0 { $validInput = $true; return $false }
          1 { $validInput = $true; return $true }
          default { Write-Warning "Invalid choice. Please try again." }
      }
  } while (-not $validInput)
}