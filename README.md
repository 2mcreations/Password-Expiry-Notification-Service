# Password Expiry Notification Service

Script PowerShell automatico per notificare gli utenti sulla scadenza delle password Active Directory e monitorare la validità del Client Secret di Azure AD utilizzato per l'invio delle email tramite Microsoft Graph API.

## Descrizione

Questo script esegue due funzioni principali [web:11][web:13]:

1. **Monitoraggio Client Secret Azure AD**: Controlla la data di scadenza del Client Secret e invia notifiche all'amministratore quando si avvicina la scadenza
2. **Notifiche scadenza password utenti**: Invia email automatiche agli utenti Active Directory quando la loro password sta per scadere

## Prerequisiti

- Windows Server con modulo Active Directory installato
- Applicazione registrata in Azure AD con permessi per Microsoft Graph API
- Permessi `Mail.Send` configurati nell'applicazione Azure AD
- PowerShell 5.1 o superiore
- Accesso in lettura ad Active Directory

## Configurazione

### 1. Registrazione applicazione Azure AD

1. Accedi al [Portale Azure](https://portal.azure.com)
2. Vai su **Azure Active Directory** > **Registrazioni app** > **Nuova registrazione**
3. Assegna un nome (es. "Password Expiry Notification Service")
4. Crea un nuovo **Client Secret** in **Certificati e segreti**
5. Configura i permessi API:
   - `Mail.Send` (Application permission)
   - Concedi il consenso amministratore

### 2. Variabili di configurazione

Crea un file di configurazione o definisci le seguenti variabili nello script:

```powershell
$TenantId = "your-tenant-id"
$ClientId = "your-client-id"
$ClientSecret = "your-client-secret"
$ClientSecretExpiryDate = [datetime]"2026-12-31"
$MailSender = "noreply@tuodominio.com"
$AdminEmail = "admin@tuodominio.com"
$PasswordExpiryCC = "helpdesk@tuodominio.com"
$WarningDays = @(1, 3, 7, 14)  # Giorni prima della scadenza per inviare notifiche
$SecretWarningDays = @(30, 15, 7, 3, 1)  # Giorni prima della scadenza del secret
