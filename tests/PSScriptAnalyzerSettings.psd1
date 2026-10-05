# Application and maintained top-level test entry points use security/defect rules.
# Pester DSL test bodies, legacy reproduction tools and vendor code are outside
# this scope; the mandatory Pester gate exercises test discovery and execution.
# These checks complement actual Windows 5.1/7 execution; they cannot prove API,
# codec, filesystem or desktop compatibility. All findings fail the analyzer gate.
@{
    IncludeDefaultRules = $false
    IncludeRules = @(
        'PSAvoidAssignmentToAutomaticVariable'
        'PSAvoidUsingInvokeExpression'
        'PSAvoidUsingPlainTextForPassword'
        'PSAvoidUsingConvertToSecureStringWithPlainText'
        'PSAvoidUsingAllowUnencryptedAuthentication'
        'PSAvoidUsingBrokenHashAlgorithms'
        'PSAvoidUsingUsernameAndPasswordParams'
        'PSAvoidUsingComputerNameHardcoded'
        'PSUsePSCredentialType'
        'PSAvoidUsingWMICmdlet'
        'PSReservedParams'
        'PSMisleadingBacktick'
        'PSUseCmdletCorrectly'
        'PSPossibleIncorrectComparisonWithNull'
        'PSUseCompatibleSyntax'
    )
    Rules = @{
        PSUseCompatibleSyntax = @{
            Enable = $true
            TargetVersions = @('5.1', '7.0')
        }
    }
}
