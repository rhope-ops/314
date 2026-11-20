# PSScriptAnalyzer settings, read by tests/Lint.ps1. Errors and warnings fail the lint.
# The key rule: syntax must work on Windows PowerShell 5.1 (the gov-device version).
# Turned off, with the reason:
#   PSUseSingularNouns - names such as Find-SsnHits read better; style only.
#   PSUseShouldProcessForStateChangingFunctions - New-PiiHit, New-ReportRow and New-ZipBytes only build
#     objects in memory; they change no system state.
#   PSReviewUnusedParameter - the test fakes must declare the real PnP cmdlet parameters to accept the calls.
#   PSAvoidUsingWriteHost - the console summary is for a person to read.
@{
    Severity     = @('Error', 'Warning')
    Rules        = @{
        PSUseCompatibleSyntax = @{ Enable = $true; TargetVersions = @('5.1') }
    }
    ExcludeRules = @('PSAvoidUsingWriteHost', 'PSUseSingularNouns', 'PSUseShouldProcessForStateChangingFunctions', 'PSReviewUnusedParameter')
}
