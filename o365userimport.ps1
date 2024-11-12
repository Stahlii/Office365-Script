# Excel dosyasının yolunu belirtin
$excelPath = "Path:\to\file.xlsx"
$users = Import-Excel -Path $excelPath -StartRow 1

foreach ($user in $users) {
    # Şifre ve diğer bilgileri hazırlayın
    $passwordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $passwordProfile.Password = $user.'Initial password [passwordProfile] Required'

    $mailNickname = ($user.'User name [userPrincipalName] Required').Split("@")[0]
    $accountEnabled = if ($user.'Block sign in (Yes/No) [accountEnabled] Required' -eq "No") { $true } else { $false }

    # Telefon numarasını uluslararası formatta temizleyin
    $cleanedMobilePhone = $user.'Mobile phone [mobile]' -replace '\s', ''  # Boşlukları kaldır
    if ($cleanedMobilePhone -notmatch '^\+90\d{10}$') {
        Write-Host "Geçersiz telefon numarası formatı: $($user.'Mobile phone [mobile]') - Düzeltilmiş: +90 $cleanedMobilePhone"
        $cleanedMobilePhone = ""
    }

    # Kullanıcı var mı kontrol edin
    try {
        $existingUser = Get-AzureADUser -ObjectId $user.'User name [userPrincipalName] Required' -ErrorAction Stop
    } catch {
        $existingUser = $null
    }

    if ($existingUser) {
        Write-Host "Mevcut kullanıcı güncelleniyor: $($user.'User name [userPrincipalName] Required')"
        
        # Mevcut kullanıcıyı güncelleyin
        Set-AzureADUser -ObjectId $existingUser.ObjectId -UsageLocation "TR"
        Set-AzureADUser -ObjectId $existingUser.ObjectId -AccountEnabled $accountEnabled

        # Başlık ve diğer bilgileri güncelleyin
        try {
            Update-MgUser -UserId $existingUser.ObjectId `
                -JobTitle $user.'Job title [jobTitle]' `
                -Department $user.'Department [department]' `
                -GivenName $user.'First name [givenName]' `
                -Surname $user.'Last name [surname]' `
                -OfficeLocation $user.'Office [physicalDeliveryOfficeName]' `
                -City $user.'City [city]' `
                -Country $user.'Country or region [country]'
        } catch {
            Write-Host "Başlık veya adres güncelleme hatası: $_"
        }

        # Telefon numarasını güncelleyin
        if ($cleanedMobilePhone -ne "") {
            try {
                Update-MgUser -UserId $existingUser.ObjectId -MobilePhone $cleanedMobilePhone
            } catch {
                Write-Host "Telefon güncelleme hatası: $_"
            }
        }
    }
    else {
        Write-Host "Yeni kullanıcı oluşturuluyor: $($user.'User name [userPrincipalName] Required')"
        
        # Yeni kullanıcı oluşturun
        $newUser = New-AzureADUser `
            -DisplayName $user.'Name [displayName] Required' `
            -UserPrincipalName $user.'User name [userPrincipalName] Required' `
            -PasswordProfile $passwordProfile `
            -MailNickName $mailNickname `
            -AccountEnabled $accountEnabled `
            -GivenName $user.'First name [givenName]' `
            -Surname $user.'Last name [surname]' `
            -JobTitle $user.'Job title [jobTitle]' `
            -Department $user.'Department [department]' `
            -UsageLocation "TR"

        # Başlık ve diğer bilgileri ekleyin
        try {
            Update-MgUser -UserId $newUser.ObjectId `
                -OfficeLocation $user.'Office [physicalDeliveryOfficeName]' `
                -City $user.'City [city]' `
                -Country $user.'Country or region [country]'
        } catch {
            Write-Host "Başlık veya adres ekleme hatası: $_"
        }

        # Telefon numarasını ekleyin
        if ($cleanedMobilePhone -ne "") {
            try {
                Update-MgUser -UserId $newUser.ObjectId -MobilePhone $cleanedMobilePhone
            } catch {
                Write-Host "Telefon ekleme hatası: $_"
            }
        }
    }
}
