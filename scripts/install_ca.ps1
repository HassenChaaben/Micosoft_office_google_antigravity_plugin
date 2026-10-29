$certPath = "C:\Users\user\.office-addin-dev-certs\ca.crt"
if (Test-Path $certPath) {
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store([System.Security.Cryptography.X509Certificates.StoreName]::Root, [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    $store.Add($cert)
    $store.Close()
    Write-Host "Certificate added to CurrentUser\Root!"
} else {
    Write-Host "ca.crt not found!"
}
