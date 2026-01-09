#Requires -Modules ActiveDirectory

# === ABILITA TLS 1.2 ===
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# === CONFIGURAZIONE MICROSOFT 365 ===
$TenantId = "xxxxxxxxxxx"
$ClientId = "xxxxxxxxxxx"
$ClientSecret = "xxxxxxxxxxx"
$MailSender = "info@tuodominio.it"
$PasswordExpiryCC = "cc@tuodominio.it"

# === CONFIGURAZIONE NOTIFICHE PASSWORD ===
$WarningDays = @(20, 15, 10, 5, 4, 3, 2, 1)

# === CONFIGURAZIONE NOTIFICA SCADENZA CLIENT SECRET ===
$AdminEmail = "admin@tuodominio.it"
$ClientSecretExpiryDate = [datetime]::ParseExact("18/11/2027", "dd/MM/yyyy", $null)
$SecretWarningDays = @(60, 30, 15, 10, 5, 3, 1)  # Avvisi piu anticipati per il secret

# Funzione per ottenere il token di accesso
function Get-GraphAccessToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $TokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    $Body = @{
        client_id     = $ClientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $ClientSecret
        grant_type    = "client_credentials"
    }

    try {
        $Response = Invoke-RestMethod -Uri $TokenEndpoint -Method Post -Body $Body
        return $Response.access_token
    }
    catch {
        Write-Error "Errore nell'ottenere il token: $_"
        Write-Error "Dettagli: $($_.Exception.Message)"
        return $null
    }
}

# Funzione per inviare email tramite Graph API
function Send-GraphMail {
    param(
        [string]$AccessToken,
        [string]$From,
        [string]$To,
        [string]$Subject,
        [string]$Body,
        [string]$CC = $null  # Parametro opzionale per CC
    )
    
    $Headers = @{
        "Authorization" = "Bearer $AccessToken"
        "Content-Type"  = "application/json"
    }
    
    # Costruisci il messaggio base
    $MailMessage = @{
        message = @{
            subject = $Subject
            body = @{
                contentType = "HTML"
                content = $Body
            }
            from = @{
                emailAddress = @{
                    address = $From
                }
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $To
                    }
                }
            )
        }
        saveToSentItems = "true"
    }
    
    # Aggiungi CC solo se specificato E diverso dal destinatario principale
    if (-not [string]::IsNullOrEmpty($CC) -and $To -ne $CC) {
        $MailMessage.message.ccRecipients = @(
            @{
                emailAddress = @{
                    address = $CC
                }
            }
        )
    }
    
    $BodyJson = $MailMessage | ConvertTo-Json -Depth 10
    
    $GraphEndpoint = "https://graph.microsoft.com/v1.0/users/$From/sendMail"
    
    try {
        Invoke-RestMethod -Uri $GraphEndpoint -Method Post -Headers $Headers -Body $BodyJson
        return $true
    }
    catch {
        Write-Error "Errore nell'invio email: $_"
        return $false
    }
}

# === SCRIPT PRINCIPALE ===
Import-Module ActiveDirectory

# Ottieni token di accesso
$AccessToken = Get-GraphAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

if (-not $AccessToken) {
    Write-Error "Impossibile ottenere il token di accesso. Uscita dallo script."
    exit 1
}

Write-Host "=== CONTROLLO SCADENZA CLIENT SECRET AZURE AD ===" -ForegroundColor Cyan

# Controlla scadenza del Client Secret
$Today = Get-Date
$DaysUntilSecretExpiry = ($ClientSecretExpiryDate - $Today).Days

if ($DaysUntilSecretExpiry -le 0) {
    Write-Host "[CRITICO] Il Client Secret e SCADUTO!" -ForegroundColor Red
}
elseif ($SecretWarningDays -contains $DaysUntilSecretExpiry) {
    Write-Host "[ALERT] Client Secret in scadenza tra $DaysUntilSecretExpiry giorni" -ForegroundColor Yellow

    # Template email per l'amministratore
    $AdminEmailBody = @"
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
        <div style="background-color: #dc3545; color: white; padding: 15px; border-radius: 5px 5px 0 0;">
            <h2 style="margin: 0;">ALERT: Scadenza Client Secret Azure AD</h2>
        </div>
        <div style="background-color: #f8f9fa; padding: 20px; border: 1px solid #dc3545; border-top: none;">
            <p><strong>Attenzione Amministratore,</strong></p>
            <div style="background-color: #fff3cd; padding: 15px; border-left: 4px solid #ffc107; margin: 20px 0;">
                <p style="margin: 0; font-size: 16px;">
                    <strong>Il Client Secret dell'applicazione Azure AD sta per scadere!</strong>
                </p>
            </div>
            <table style="width: 100%; margin: 20px 0; border-collapse: collapse;">
                <tr style="background-color: #e9ecef;">
                    <td style="padding: 10px; border: 1px solid #dee2e6; font-weight: bold;">Applicazione</td>
                    <td style="padding: 10px; border: 1px solid #dee2e6;">Password Expiry Notification Service</td>
                </tr>
                <tr>
                    <td style="padding: 10px; border: 1px solid #dee2e6; font-weight: bold;">Application ID</td>
                    <td style="padding: 10px; border: 1px solid #dee2e6; font-family: monospace;">$ClientId</td>
                </tr>
                <tr style="background-color: #e9ecef;">
                    <td style="padding: 10px; border: 1px solid #dee2e6; font-weight: bold;">Data di scadenza</td>
                    <td style="padding: 10px; border: 1px solid #dee2e6;">$($ClientSecretExpiryDate.ToString("dd/MM/yyyy"))</td>
                </tr>
                <tr style="background-color: #fff3cd;">
                    <td style="padding: 10px; border: 1px solid #dee2e6; font-weight: bold;">Giorni rimanenti</td>
                    <td style="padding: 10px; border: 1px solid #dee2e6; color: #dc3545; font-size: 18px; font-weight: bold;">$DaysUntilSecretExpiry giorni</td>
                </tr>
            </table>

            <h3 style="color: #0066cc; margin-top: 30px;">Azioni da intraprendere:</h3>
            <ol style="line-height: 1.8;">
                <li>Accedi al <strong>Portale Azure</strong> (https://portal.azure.com)</li>
                <li>Vai su <strong>Azure Active Directory</strong> &gt; <strong>Registrazioni app</strong></li>
                <li>Seleziona l'applicazione: <strong>Password Expiry Notification Service</strong></li>
                <li>Vai alla scheda <strong>Certificati e segreti</strong></li>
                <li>Crea un <strong>nuovo Client Secret</strong></li>
                <li>Copia il <strong>Value</strong> del nuovo secret (NON il Secret ID!)</li>
                <li>Aggiorna lo script PowerShell con il nuovo secret:
                    <br><code style="background-color: #f1f1f1; padding: 5px; display: block; margin-top: 5px;">C:\automa\Send-PasswordExpiryNotification.ps1</code>
                </li>
                <li>Testa lo script dopo l'aggiornamento</li>
                <li>Una volta verificato il funzionamento, elimina il vecchio secret scaduto</li>
            </ol>

            <div style="background-color: #d1ecf1; padding: 15px; border-left: 4px solid #0c5460; margin-top: 20px;">
                <p style="margin: 0; font-size: 14px; color: #0c5460;">
                    <strong>Importante:</strong> Se il secret scade senza essere rinnovato, lo script non potra piu inviare 
                    le notifiche di scadenza password agli utenti.
                </p>
            </div>

            <hr style="margin: 30px 0; border: none; border-top: 1px solid #ddd;">
            <p style="font-size: 12px; color: #666;">
                Questo e un messaggio automatico generato dallo script di notifica password.<br>
                Server: $env:COMPUTERNAME<br>
                Data/Ora: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss")
            </p>
        </div>
    </div>
</body>
</html>
"@

    $Subject = "[URGENTE] Client Secret Azure AD in scadenza tra $DaysUntilSecretExpiry giorni"

    # Invia notifica all'amministratore
    $Result = Send-GraphMail -AccessToken $AccessToken -From $MailSender -To $AdminEmail -Subject $Subject -Body $AdminEmailBody

    if ($Result) {
        Write-Host "[OK] Email di alert inviata all'amministratore: $AdminEmail" -ForegroundColor Green
    }
    else {
        Write-Host "[ERRORE] Impossibile inviare l'email all'amministratore" -ForegroundColor Red
    }
}
else {
    Write-Host "[OK] Client Secret valido per altri $DaysUntilSecretExpiry giorni" -ForegroundColor Green
}

Write-Host ""
Write-Host "=== CONTROLLO SCADENZA PASSWORD UTENTI ===" -ForegroundColor Cyan

# Template email per utenti
$UserEmailTemplate = @"
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
        <h2 style="color: #d9534f; border-bottom: 2px solid #d9534f; padding-bottom: 10px;">
            Avviso Scadenza Password
        </h2>
        <p>Gentile <strong>{NAME}</strong>,</p>
        <div style="background-color: #f8f9fa; padding: 15px; border-left: 4px solid #d9534f; margin: 20px 0;">
            <p style="margin: 0;">
                La tua password Windows scadra tra <strong style="color: #d9534f; font-size: 18px;">{DAYS} giorni</strong>
            </p>
            <p style="margin: 10px 0 0 0;">
                Data di scadenza: <strong>{EXPIRY_DATE}</strong>
            </p>
        </div>
        <h3 style="color: #0066cc;">Come cambiare la password:</h3>
        <ul>
            <li><strong>Su PC:</strong> Premi CTRL+ALT+CANC e seleziona "Cambia password"</li>
            <li><strong>Da remoto:</strong> Accedi al portale aziendale</li>
        </ul>
        <div style="background-color: #fff3cd; padding: 10px; border-radius: 5px; margin-top: 20px;">
            <p style="margin: 0; font-size: 14px;">
                Ti preghiamo di cambiare la password prima della scadenza per evitare interruzioni nell'accesso ai servizi.
            </p>
        </div>
        <hr style="margin: 20px 0; border: none; border-top: 1px solid #ddd;">
        <p style="font-size: 12px; color: #666;">
            Questo e un messaggio automatico. Non rispondere a questa email.<br>
            Per assistenza contatta l'helpdesk IT.
        </p>
    </div>
</body>
</html>
"@

# Trova utenti con password in scadenza
$Users = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordLastSet -gt 0} -Properties "Name", "EmailAddress", "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property "Name", "EmailAddress", @{Name = "PasswordExpiry"; Expression = {[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}

$EmailsSent = 0
$EmailsFailed = 0

# Elabora ogni utente
foreach ($User in $Users) {
    if ([string]::IsNullOrEmpty($User.EmailAddress)) {
        continue
    }

    $DaysToExpiry = ($User.PasswordExpiry - $Today).Days

    if ($WarningDays -contains $DaysToExpiry) {
        # Prepara corpo email
        $EmailBody = $UserEmailTemplate -replace '{NAME}', $User.Name
        $EmailBody = $EmailBody -replace '{DAYS}', $DaysToExpiry
        $EmailBody = $EmailBody -replace '{EXPIRY_DATE}', $User.PasswordExpiry.ToLongDateString()

        $Subject = "Avviso: Password in scadenza tra $DaysToExpiry giorni"

        # Invia email
        $Result = Send-GraphMail -AccessToken $AccessToken -From $MailSender -To $User.EmailAddress -Subject $Subject -Body $EmailBody -CC $PasswordExpiryCC

        if ($Result) {
            Write-Host "[OK] Email inviata a $($User.Name) ($($User.EmailAddress)) - Scadenza tra $DaysToExpiry giorni" -ForegroundColor Green
            $EmailsSent++
        }
        else {
            Write-Host "[ERRORE] Invio email fallito a $($User.Name)" -ForegroundColor Red
            $EmailsFailed++
        }

        # Pausa per evitare throttling
        Start-Sleep -Milliseconds 500
    }
}

# Riepilogo finale
Write-Host ""
Write-Host "=== RIEPILOGO ===" -ForegroundColor Cyan
Write-Host "Email utenti inviate: $EmailsSent" -ForegroundColor Green

if ($EmailsFailed -eq 0) {
    Write-Host "Email fallite: $EmailsFailed" -ForegroundColor Green
}
else {
    Write-Host "Email fallite: $EmailsFailed" -ForegroundColor Red
}
