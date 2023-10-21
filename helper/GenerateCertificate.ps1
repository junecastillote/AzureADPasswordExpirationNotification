# Generate a self-signed certificate

$certificateName = 'Azure AD Password Notification'

$certSplat = @{
    Subject           = $certificateName
    NotBefore         = ((Get-Date).AddDays(-1))
    NotAfter          = ((Get-Date).AddYears(5))
    CertStoreLocation = "Cert:\CurrentUser\My"
    Provider          = "Microsoft Enhanced RSA and AES Cryptographic Provider"
    HashAlgorithm     = "SHA256"
    KeySpec           = "KeyExchange"
}
$selfSignedCertificate = New-SelfSignedCertificate @certSplat

# Display the certificate details
$selfSignedCertificate | Format-List PSParentPath, ThumbPrint, Subject, NotAfter

# Export the private certificate to PFX.
$selfSignedCertificate | Export-PfxCertificate -FilePath "$($certificateName).pfx" -Password $(ConvertTo-SecureString -String "Trickily2#Activity#Whisking#Emcee" -AsPlainText -Force)

# Export the public certificate to CER.
$selfSignedCertificate | Export-Certificate -FilePath "$($certificateName).cer"