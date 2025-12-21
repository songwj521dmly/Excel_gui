#include "LicenseManager.h"
#include <QSettings>
#include <QDebug>
#include <windows.h>
#include <wincrypt.h>
#include <vector>
#include <cstring>

// Ensure we link against the Windows API library
#pragma comment(lib, "advapi32.lib")

LicenseManager& LicenseManager::instance() {
    static LicenseManager instance;
    return instance;
}

LicenseManager::LicenseManager() {}
LicenseManager::~LicenseManager() {}

bool LicenseManager::isExpired() {
    QDateTime expiry = getExpiryDate();
    return QDateTime::currentDateTime() > expiry;
}

QDateTime LicenseManager::getExpiryDate() {
    // Access Registry
    QSettings settings(REGISTRY_KEY, QSettings::NativeFormat);
    QString encryptedDate = settings.value(EXPIRY_VALUE_NAME).toString();

    if (!encryptedDate.isEmpty()) {
        QString decryptedDate = decryptString(encryptedDate);
        // Format from C# code: yyyy-MM-dd HH:mm:ss
        QDateTime expiry = QDateTime::fromString(decryptedDate, "yyyy-MM-dd HH:mm:ss");
        if (expiry.isValid()) {
            return expiry;
        }
    }

    // Expiry invalid or missing -> Check First Run Date
    QString encryptedFirstRun = settings.value(FIRST_RUN_VALUE_NAME).toString();
    QDateTime firstRun;
    
    if (!encryptedFirstRun.isEmpty()) {
         QString decryptedFirstRun = decryptString(encryptedFirstRun);
         firstRun = QDateTime::fromString(decryptedFirstRun, "yyyy-MM-dd HH:mm:ss");
    }
    
    // If First Run is invalid (e.g. first time ever), set it to NOW
    if (!firstRun.isValid()) {
        firstRun = QDateTime::currentDateTime();
        // Save First Run Date
        settings.setValue(FIRST_RUN_VALUE_NAME, encryptString(firstRun.toString("yyyy-MM-dd HH:mm:ss")));
    }
    
    // Default Trial Policy: 1 Year from First Run
    QDateTime trialExpiry = firstRun.addYears(1);
    
    // Save this calculated expiry so we don't calculate it every time (and to allow manual extension later)
    saveExpiryDate(trialExpiry);
    
    return trialExpiry;
}

void LicenseManager::saveExpiryDate(const QDateTime& date) {
    QSettings settings(REGISTRY_KEY, QSettings::NativeFormat);
    QString encryptedDate = encryptString(date.toString("yyyy-MM-dd HH:mm:ss"));
    settings.setValue(EXPIRY_VALUE_NAME, encryptedDate);
}

bool LicenseManager::activateWithCode(const QString& code) {
    if (code.trimmed().isEmpty()) return false;
    
    if (validateActivationCode(code)) {
        // Extend for 1 year from NOW
        QDateTime newExpiry = QDateTime::currentDateTime().addYears(1);
        saveExpiryDate(newExpiry);
        return true;
    }
    return false;
}

QString LicenseManager::generateActivationCode() {
    // Format: yyyyMMddHHmm
    QString timestamp = QDateTime::currentDateTime().toString("yyyyMMddHHmm");
    QString dataToEncrypt = timestamp + "songwj";
    return encryptString(dataToEncrypt);
}

bool LicenseManager::validateActivationCode(const QString& code) {
    if (code.isEmpty()) return false;
    
    QString decrypted = decryptString(code);
    
    // Check basic format: length >= 18 and ends with "songwj"
    // 12 (timestamp) + 6 ("songwj") = 18
    if (decrypted.isEmpty() || decrypted.length() < 18 || !decrypted.endsWith("songwj")) {
        return false;
    }
    
    // Extract timestamp
    QString timestampStr = decrypted.left(12);
    QDateTime codeTime = QDateTime::fromString(timestampStr, "yyyyMMddHHmm");
    
    if (!codeTime.isValid()) return false;
    
    // Check 10 mins validity (Activation code must be used within 10 mins of generation)
    // C# logic: DateTime.Now - codeTime
    qint64 minsDiff = codeTime.secsTo(QDateTime::currentDateTime()) / 60;
    
    // If diff is > 10 mins (code is old) or < 0 (code is in future), invalid
    if (minsDiff > 10 || minsDiff < 0) {
        return false;
    }
    
    QString expected = timestampStr + "songwj";
    return decrypted == expected;
}

QString LicenseManager::encryptString(const QString& plainText) {
    if (plainText.isEmpty()) return "";
    
    // Key derivation: First 32 bytes for Key, First 16 bytes for IV
    // Source: "CadCad2024License!@#$%^&*()_+123"
    QByteArray keyBytes = QByteArray::fromRawData(ENCRYPTION_KEY_SOURCE.c_str(), 32);
    QByteArray ivBytes = QByteArray::fromRawData(ENCRYPTION_KEY_SOURCE.c_str(), 16);
    
    QByteArray encrypted = aesEncrypt(plainText.toUtf8(), keyBytes, ivBytes);
    return QString::fromLatin1(encrypted.toBase64());
}

QString LicenseManager::decryptString(const QString& cipherText) {
    if (cipherText.isEmpty()) return "";
    
    QByteArray keyBytes = QByteArray::fromRawData(ENCRYPTION_KEY_SOURCE.c_str(), 32);
    QByteArray ivBytes = QByteArray::fromRawData(ENCRYPTION_KEY_SOURCE.c_str(), 16);
    
    QByteArray encryptedBytes = QByteArray::fromBase64(cipherText.toLatin1());
    QByteArray decrypted = aesDecrypt(encryptedBytes, keyBytes, ivBytes);
    return QString::fromUtf8(decrypted);
}

// ==========================================
// AES Implementation using Windows CAPI
// ==========================================

// Custom struct for Plaintext Key Import
struct PlainTextKeyBlob {
    BLOBHEADER hdr;
    DWORD cbKeySize;
    BYTE rgbKeyData[32]; // 32 bytes for AES-256
};

QByteArray LicenseManager::aesEncrypt(const QByteArray& data, const QByteArray& key, const QByteArray& iv) {
    HCRYPTPROV hProv = 0;
    HCRYPTKEY hKey = 0;
    QByteArray result;
    
    // Acquire Context
    if (!CryptAcquireContext(&hProv, NULL, MS_ENH_RSA_AES_PROV, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)) {
        // If it fails, try to create a new key container (though VERIFYCONTEXT avoids needing one)
        // Usually VERIFYCONTEXT is enough for ephemeral keys.
        return result;
    }
    
    // Import Key
    PlainTextKeyBlob keyBlob;
    keyBlob.hdr.bType = PLAINTEXTKEYBLOB;
    keyBlob.hdr.bVersion = CUR_BLOB_VERSION;
    keyBlob.hdr.reserved = 0;
    keyBlob.hdr.aiKeyAlg = CALG_AES_256;
    keyBlob.cbKeySize = 32;
    memcpy(keyBlob.rgbKeyData, key.constData(), 32);
    
    if (CryptImportKey(hProv, (BYTE*)&keyBlob, sizeof(keyBlob), 0, 0, &hKey)) {
        // Set IV
        CryptSetKeyParam(hKey, KP_IV, (BYTE*)iv.constData(), 0);
        
        // Set Mode CBC
        DWORD mode = CRYPT_MODE_CBC;
        CryptSetKeyParam(hKey, KP_MODE, (BYTE*)&mode, 0);
        
        // Prepare Buffer
        // CryptEncrypt encrypts in-place. We need enough buffer for padding.
        DWORD dataLen = data.size();
        DWORD bufLen = dataLen + 16; // AES block size is 16. Padding adds at most 1 block.
        std::vector<BYTE> buffer(bufLen);
        memcpy(buffer.data(), data.constData(), dataLen);
        
        // Encrypt
        if (CryptEncrypt(hKey, 0, TRUE, 0, buffer.data(), &dataLen, bufLen)) {
            result = QByteArray((char*)buffer.data(), dataLen);
        }
        
        CryptDestroyKey(hKey);
    }
    
    CryptReleaseContext(hProv, 0);
    return result;
}

QByteArray LicenseManager::aesDecrypt(const QByteArray& data, const QByteArray& key, const QByteArray& iv) {
    HCRYPTPROV hProv = 0;
    HCRYPTKEY hKey = 0;
    QByteArray result;
    
    if (!CryptAcquireContext(&hProv, NULL, MS_ENH_RSA_AES_PROV, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)) {
        return result;
    }
    
    PlainTextKeyBlob keyBlob;
    keyBlob.hdr.bType = PLAINTEXTKEYBLOB;
    keyBlob.hdr.bVersion = CUR_BLOB_VERSION;
    keyBlob.hdr.reserved = 0;
    keyBlob.hdr.aiKeyAlg = CALG_AES_256;
    keyBlob.cbKeySize = 32;
    memcpy(keyBlob.rgbKeyData, key.constData(), 32);
    
    if (CryptImportKey(hProv, (BYTE*)&keyBlob, sizeof(keyBlob), 0, 0, &hKey)) {
        // Set IV
        CryptSetKeyParam(hKey, KP_IV, (BYTE*)iv.constData(), 0);
        
        // Set Mode CBC
        DWORD mode = CRYPT_MODE_CBC;
        CryptSetKeyParam(hKey, KP_MODE, (BYTE*)&mode, 0);
        
        // Prepare Buffer
        DWORD dataLen = data.size();
        std::vector<BYTE> buffer(dataLen);
        memcpy(buffer.data(), data.constData(), dataLen);
        
        // Decrypt
        if (CryptDecrypt(hKey, 0, TRUE, 0, buffer.data(), &dataLen)) {
            result = QByteArray((char*)buffer.data(), dataLen);
        }
        
        CryptDestroyKey(hKey);
    }
    
    CryptReleaseContext(hProv, 0);
    return result;
}