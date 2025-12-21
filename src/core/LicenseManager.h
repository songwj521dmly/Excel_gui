#pragma once

#include <string>
#include <vector>
#include <QString>
#include <QDateTime>
#include <memory>

class LicenseManager {
public:
    static LicenseManager& instance();

    // Check if software is expired
    bool isExpired();

    // Get current expiry date
    QDateTime getExpiryDate();

    // Activate with code
    bool activateWithCode(const QString& code);

    // Generate activation code (Time + "songwj" -> Encrypt)
    QString generateActivationCode();

    // Validate activation code
    bool validateActivationCode(const QString& code);

private:
    LicenseManager();
    ~LicenseManager();
    LicenseManager(const LicenseManager&) = delete;
    LicenseManager& operator=(const LicenseManager&) = delete;

    void saveExpiryDate(const QDateTime& date);
    QString encryptString(const QString& plainText);
    QString decryptString(const QString& cipherText);

    // AES Implementation Helpers
    QByteArray aesEncrypt(const QByteArray& data, const QByteArray& key, const QByteArray& iv);
    QByteArray aesDecrypt(const QByteArray& data, const QByteArray& key, const QByteArray& iv);

    const QString REGISTRY_KEY = "HKEY_CURRENT_USER\\Software\\ExcelCL";
    const QString EXPIRY_VALUE_NAME = "ExpiryDate";
    const QString FIRST_RUN_VALUE_NAME = "FirstRunDate";
    
    // Key source from C# reference: "CadCad2024License!@#$%^&*()_+123"
    // Key (32 bytes) and IV (16 bytes) are derived from this.
    const std::string ENCRYPTION_KEY_SOURCE = "CadCad2024License!@#$%^&*()_+123";
};
