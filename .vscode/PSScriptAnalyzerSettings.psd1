@{
    # Enable all rules by default
    IncludeDefaultRules = $true

    # Severity levels to include
    Severity = @('Error', 'Warning', 'Information')

    # Custom rules configuration
    Rules = @{
        # Enforce consistent indentation
        PSUseConsistentIndentation = @{
            Enable = $true
            IndentationSize = 4
            PipelineIndentation = 'IncreaseIndentationForFirstPipeline'
            Kind = 'space'
        }

        # Enforce consistent whitespace
        PSUseConsistentWhitespace = @{
            Enable = $true
            CheckInnerBrace = $true
            CheckOpenBrace = $true
            CheckOpenParen = $true
            CheckOperator = $false # Disabled to allow vertical alignment
            CheckPipe = $true
            CheckPipeForRedundantWhitespace = $false
            CheckSeparator = $false # Disabled to allow vertical alignment
            CheckParameter = $false
        }

        # Enforce proper case for cmdlets
        PSUseCorrectCasing = @{
            Enable = $true
        }

        # Require approved verbs
        PSUseApprovedVerbs = @{
            Enable = $true
        }

        # Avoid using Write-Host
        PSAvoidUsingWriteHost = @{
            Enable = $true
        }

        # Require full cmdlet names (no aliases)
        PSAvoidUsingCmdletAliases = @{
            Enable = $true
        }

        # Require named parameters
        PSAvoidUsingPositionalParameters = @{
            Enable = $true
            CommandAllowList = @()
        }

        # Security rules
        PSAvoidUsingPlainTextForPassword = @{
            Enable = $true
        }

        PSAvoidUsingConvertToSecureStringWithPlainText = @{
            Enable = $true
        }

        # Performance rules
        PSUseDeclaredVarsMoreThanAssignments = @{
            Enable = $true
        }

        # Style rules
        PSPlaceOpenBrace = @{
            Enable = $true
            OnSameLine = $true
            NewLineAfter = $true
            IgnoreOneLineBlock = $true
        }

        PSPlaceCloseBrace = @{
            Enable = $true
            NewLineAfter = $false
            IgnoreOneLineBlock = $true
            NoEmptyLineBefore = $false
        }
    }

    # Exclude specific rules if needed
    ExcludeRules = @(
        # Add any rules you want to exclude here
    )
}
