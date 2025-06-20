Connect-ExchangeOnline

Function Add-Mailbox-Users {
    param(
        [ValidateSet("FullAccess", "SendAs", "Both")]
        [string]$PermissionType = "Both"
    )

    # Lista de caixas compartilhadas
    $sharedMailboxes = @(
        
    )

    # Lista de usuarios que devem ser adicionados em todas as caixas
    $usersToAdd = @(
       
    )

    # Inicializar log
    $log = @()

    foreach ($mailbox in $sharedMailboxes) {
        Write-Host ""
        Write-Host "Processando caixa: $mailbox" -ForegroundColor Cyan

        foreach ($user in $usersToAdd) {
            Write-Host "Verificando usuario: $user" -ForegroundColor Yellow

            try {
                if ($PermissionType -eq "FullAccess" -or $PermissionType -eq "Both") {
                    $hasFullAccess = Get-MailboxPermission -Identity $mailbox -User $user -ErrorAction SilentlyContinue |
                        Where-Object { $_.AccessRights -contains "FullAccess" -and $_.IsInherited -eq $false }

                    if (-not $hasFullAccess) {
                        Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -AutoMapping:$false -ErrorAction Stop
                        Write-Host "FullAccess adicionado." -ForegroundColor Green
                    }
                    else {
                        Write-Host "FullAccess ja existe." -ForegroundColor Yellow
                    }
                }

                if ($PermissionType -eq "SendAs" -or $PermissionType -eq "Both") {
                    $hasSendAs = Get-RecipientPermission -Identity $mailbox -Trustee $user -ErrorAction SilentlyContinue |
                        Where-Object { $_.AccessRights -contains "SendAs" }

                    if (-not $hasSendAs) {
                        Add-RecipientPermission -Identity $mailbox -Trustee $user -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                        Write-Host "SendAs adicionado." -ForegroundColor Green
                    }
                    else {
                        Write-Host "SendAs ja existe." -ForegroundColor Yellow
                    }
                }

                $log += [PSCustomObject]@{
                    Mailbox = $mailbox
                    User = $user
                    Status = "Sucesso ou ja existente"
                }

            }
            catch {
                $log += [PSCustomObject]@{
                    Mailbox = $mailbox
                    User = $user
                    Status = "Falha: $($_.Exception.Message)"
                }

                Write-Host "Erro: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    # Exibir o resumo final
    Write-Host ""
    Write-Host "Resumo das operacoes:" -ForegroundColor Cyan
    $log | Format-Table -AutoSize

    # Exportar log opcional
    # $log | Export-Csv -Path "log_mailbox_permissions.csv" -NoTypeInformation -Encoding UTF8
}

# Exemplo de uso:
# Add-Mailbox-Users -PermissionType Both
# Add-Mailbox-Users -PermissionType FullAccess
# Add-Mailbox-Users -PermissionType SendAs

Add-Mailbox-Users -PermissionType Both
