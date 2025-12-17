#include <QCoreApplication>
#include <QTextStream>
#include <QDebug>
#include "../core/LicenseManager.h"
#include <iostream>

int main(int argc, char *argv[]) {
    QCoreApplication app(argc, argv);
    QTextStream out(stdout);

    out << "==================================================" << endl;
    out << "   Excel Processor Activation Code Generator" << endl;
    out << "==================================================" << endl;

    QString code = LicenseManager::instance().generateActivationCode();
    
    out << "\n[NEW CODE]: " << code << endl;
    out << "\nCopy the line above to activate the software." << endl;
    out << "This code is valid for 10 minutes from now." << endl;
    out << "==================================================" << endl;

    system("pause");
    return 0;
}
