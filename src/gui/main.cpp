#include <QApplication>
#include <QInputDialog>
#include <QMessageBox>
#include "MainWindow.h"
#include "../core/LicenseManager.h"

int main(int argc, char *argv[]) {
    QApplication app(argc, argv);

    // Check License
    if (LicenseManager::instance().isExpired()) {
        bool ok;
        while (true) {
            QString code = QInputDialog::getText(nullptr, 
                QString::fromUtf8("\xE8\xBD\xAF\xE4\xBB\xB6\xE8\xBF\x87\xE6\x9C\x9F"), // 软件过期
                QString::fromUtf8("\xE8\xBD\xAF\xE4\xBB\xB6\xE5\xB7\xB2\xE5\x88\xB0\xE6\x9C\x9F\xEF\xBC\x8C\xE8\xAF\xB7\xE8\xBE\x93\xE5\x85\xA5\xE6\xBF\x80\xE6\xB4\xBB\xE7\xA0\x81\xE7\xBB\xAD\xE6\x9C\x9F\xEF\xBC\x9A"), // 软件已到期，请输入激活码续期：
                QLineEdit::Normal,
                "", &ok);
            
            if (!ok) {
                return 0; // User cancelled
            }

            if (LicenseManager::instance().activateWithCode(code)) {
                QMessageBox::information(nullptr, 
                    QString::fromUtf8("\xE6\xBF\x80\xE6\xB4\xBB\xE6\x88\x90\xE5\x8A\x9F"), // 激活成功
                    QString::fromUtf8("\xE8\xBD\xAF\xE4\xBB\xB6\xE5\xB7\xB2\xE6\x88\x90\xE5\x8A\x9F\xE7\xBB\xAD\xE6\x9C\x9F\xE4\xB8\x80\xE5\xB9\xB4\xEF\xBC\x81")); // 软件已成功续期一年！
                break;
            } else {
                QMessageBox::critical(nullptr, 
                    QString::fromUtf8("\xE6\xBF\x80\xE6\xB4\xBB\xE5\xA4\xB1\xE8\xB4\xA5"), // 激活失败
                    QString::fromUtf8("\xE6\xBF\x80\xE6\xB4\xBB\xE7\xA0\x81\xE6\x97\xA0\xE6\x95\x88\xE6\x88\x96\xE5\xB7\xB2\xE8\xBF\x87\xE6\x9C\x9F\xEF\xBC\x8C\xE8\xAF\xB7\xE9\x87\x8D\xE8\xAF\x95\xE3\x80\x82")); // 激活码无效或已过期，请重试。
            }
        }
    }

    ExcelProcessorGUI window;
    window.show();

    return app.exec();
}
