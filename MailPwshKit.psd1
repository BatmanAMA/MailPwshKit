@{
    GUID = "4ae9fd46-338a-459c-8186-07f910774cb8"
    Author = "Batmanama"
    CompanyName = "Batmanama"
    Copyright = "(C) Batmanama. No rights reserved."
    HelpInfoUri = "https://GitHub.com/batmanama/mailpwshkit"
    ModuleVersion = "0.5.0"
    PowerShellVersion = "3.0"
    ClrVersion = "4.0"
    RootModule = "mailpwshkit.dll"
    Description = 'A replacement for "Send-MailMessage" using Mailkit.'

    CmdletsToExport = @(
        'send-mailkitmessage'
    )

    FormatsToProcess  = @()

    PrivateData = @{
        PSData = @{
            Tags = @('Mail', 'Mailkit','send-mailmessage', 'PSEdition_Core', 'PSEdition_Desktop', 'Windows', 'Linux', 'macOS')
            ProjectUri = "https://GitHub.com/batmanama/mailpwshkit"
        }
    }
}