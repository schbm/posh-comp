Function Get-TUIChoicePrompt{
    <#
      .Synopsis
        Prompts a closed question with the ability to decide.
      .Description
        Prompt a closed question with the ability to decide. Returns $false for yes and $true for no. If nothing is supplied,
        no is the default result. If wrong inputs are supplied, it is repeated until a valid input or no input is given.
        Ctrl-C interupt not implemented yet.
      .Parameter Title
        [String] Define a title of the question.
      .Parameter Content
        [String] Define a closed yes or no question.
      .Outputs
        [bool]
      .Example
        Get-UIChoicePrompt -Title "Test" -Content "This or that"
      .Notes 
        Name: Marcel Schubert 
        Author: Marcel Schubert
        LastEdit: 30.11.2021
    #>
    [CmdletBinding()]
    param(
        [string]$Title = "Title",
        [string]$Content = "Content"
    )
    $Choices = @("&yes", "&no")
    if(($Choices | Measure-Object).Count -gt 2){
        Throw('Get-UIChoicePrompt: More than 2 choices supplied.')
    }
    $decision = $Host.UI.PromptForChoice($Title, $Content, $Choices, 1)
    return [bool]$decision
}

Function Update-TProgress{
    <#
      .Synopsis
        Write-Progress wrapper
      .Notes 
        Name: Marcel Schubert 
        Author: Marcel Schubert
        LastEdit: 30.11.2021
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [int]$current,
        [Parameter(Mandatory=$True)]
        [int]$total,
        [Parameter(Mandatory=$True)]
        [string]$loopName,
        [Parameter(Mandatory=$True)]
        [int]$id
    )
    $percentComplete = ($current / $total) * 100 
    Write-Progress -Activity 'Writing number blocks' -Status "$current/$total" -PercentComplete $percentComplete -CurrentOperation $loopName -Id $id
  }
  
  Function Finish-TProgress{
      <#
        .Synopsis
          Write-Progress wrapper
        .Notes 
          Name: Marcel Schubert 
          Author: Marcel Schubert
          LastEdit: 30.11.2021
      #>
      [CmdletBinding()]
      param(
        [Parameter(Mandatory=$True)]
        [string]$loopName,
        [Parameter(Mandatory=$True)]
        [int]$id
      )
    Write-Progress -Activity "Writing number blocks" -Status "Ready" -Completed -CurrentOperation $loopName -Id $id
  }
