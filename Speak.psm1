<#
.Synopsis
   Text to Speech
.DESCRIPTION
   Windows computer will speak given text.  
.EXAMPLE
   Speak "You've been pwned" -Drunk
.EXAMPLE
   Speak You have mail 
#>
function Speak {
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Text,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $Text1,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        $Text2,

        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        $Text3,

        [int]$Rate = 1,

        [switch]
        $Drunk
    )

    $object = New-Object -ComObject SAPI.SpVoice
    if ($drunk) { $object.Rate = -10 }
    else {$object.Rate = $Rate}
    If ($Text1 -like "")
    {
        $object.Speak("$text") | Out-Null
    }
    Else
    {
        If ($Text2 -like "")
        {
            $object.Speak("$text $Text1") | Out-Null
        }
        Else
        {
            $object.Speak("$text $Text1 $Text2 $Text3") | Out-Null
        }
    }
    

}
New-Alias -name text2speech -Value Speak
New-Alias -name Say -Value Speak
Export-ModuleMember -Alias * -Function Speak