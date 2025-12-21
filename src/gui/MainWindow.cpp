#ifdef _WIN32
#define NOMINMAX
#endif
#include "MainWindow.h"
#include <QStringList>
#include "RuleEditDialog.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QPushButton>
#include <QTableWidget>
#include <QListWidget>
#include <QListWidgetItem>
#include <QTreeWidget>
#include <QTreeWidgetItem>
#include <QTextEdit>
#include <QLabel>
#include <QLineEdit>
#include <QComboBox>
#include <QCheckBox>
#include <QProgressBar>
#include <QMessageBox>
#include <QFileDialog>
#include <QSplitter>
#include <QTabWidget>
#include <QGroupBox>
#include <QInputDialog>
#include <QStatusBar>
#include <QHeaderView>
#include <QTimer>
#include <QApplication>
#include <QDir>
#include <QFileInfo>
#include <QDialog>
#include <QDialogButtonBox>
#include <QMenu>
#include <QAction>
#include <QRadioButton>
#include <QButtonGroup>
#include <QScrollArea>
#include <iostream>
#include <sstream>
#include <iomanip>
#include <chrono>
#include <thread>
#ifdef _WIN32
#include <objbase.h>
#undef DELETE
#endif

#include <QSettings>
#include <QDateTime>
#include <QFileInfo>

ExcelProcessorGUI::ExcelProcessorGUI(QWidget *parent) : QMainWindow(parent) {
    processor_ = std::make_unique<ExcelProcessorCore>();
    
    // Setup Callbacks
    processor_->setLogger([this](const std::string& msg) {
        QMetaObject::invokeMethod(this, [this, msg]() {
            if (logTextEdit_) {
                QString timestamp = QDateTime::currentDateTime().toString("[yyyy-MM-dd hh:mm:ss] ");
                // Explicitly use fromUtf8 to ensure correct decoding of UTF-8 log messages
                logTextEdit_->append(timestamp + QString::fromUtf8(msg.c_str()));
            }
        });
    });

    processor_->setProgressCallback([this](int progress, const std::string& msg) {
        QMetaObject::invokeMethod(this, [this, progress, msg]() {
            if (progressBar_) {
                // For indeterminate progress, we don't set value
                // progressBar_->setValue(progress); 
            }
            
            // Log progress for large files (every chunk)
            if (logTextEdit_) {
                QString timestamp = QDateTime::currentDateTime().toString("[yyyy-MM-dd hh:mm:ss] ");
                logTextEdit_->append(timestamp + QString::fromUtf8(msg.c_str()));
            }
        });
    });

    setupUI();
    setupConnections();
    loadSettings();
}

ExcelProcessorGUI::~ExcelProcessorGUI() {
    saveSettings();
}

void ExcelProcessorGUI::loadSettings() {
    QSettings settings("MyCompany", "ExcelProcessor");
    QString lastConfig = settings.value("LastConfigFile").toString();

    // Auto load last config if it exists
    if (!lastConfig.isEmpty() && QFile::exists(lastConfig)) {
        if (processor_->loadConfiguration(lastConfig.toStdString())) {
            currentConfigFile_ = lastConfig;
            updateRuleTable();
            updateTaskTree(); // Changed from updateTaskTable
            statusLabel_->setText(QString::fromUtf8("\xE5\xB7\xB2\xE8\x87\xAA\xE5\x8A\xA0\xE8\xBD\xBD\xE4\xB8\x8A\xE6\xAC\xA1\xE9\x85\x8D\xE7\xBD\xAE"));
            
            if (configNameLabel_) {
                QString configName = QFileInfo(currentConfigFile_).fileName();
                configNameLabel_->setText(QString::fromUtf8("\xE5\xBD\x93\xE5\x89\x8D\xE9\x85\x8D\xE7\xBD\xAE: ") + configName);
                configNameLabel_->setStyleSheet("color: black; font-weight: bold; margin-right: 10px;");
            }

            // Log startup info
            if (logTextEdit_) {
                QString timestamp = QDateTime::currentDateTime().toString("[yyyy-MM-dd hh:mm:ss] ");
                int ruleCount = processor_->getRules().size();
                QString msg = QString::fromUtf8("\xE7\xB3\xBB\xE7\xBB\x9F\xE5\x90\xAF\xE5\x8A\xA8: \xE8\x87\xAA\xE5\x8A\xA0\xE8\xBD\xBD\xE9\x85\x8D\xE7\xBD\xAE\xE6\x96\x87\xE4\xBB\xB6 '%1' (\xE5\x85\xB1 %2 \xE4\xB8\xAA\xE8\xA7\x84\xE5\x88\x99)")
                              .arg(QFileInfo(currentConfigFile_).fileName())
                              .arg(ruleCount);
                logTextEdit_->append(timestamp + msg);
            }
        }
    }
}

void ExcelProcessorGUI::saveSettings() {
    QSettings settings("MyCompany", "ExcelProcessor");
    if (!currentConfigFile_.isEmpty()) {
        settings.setValue("LastConfigFile", currentConfigFile_);
    }
}

void ExcelProcessorGUI::loadConfiguration() {
    QString filename = QFileDialog::getOpenFileName(
        this, QString::fromUtf8("\xE5\x8A\xA0\xE8\xBD\xBD\xE9\x85\x8D\xE7\xBD\xAE"), "", QString::fromUtf8("\xE9\x85\x8D\xE7\xBD\xAE\xE6\x96\x87\xE4\xBB\xB6 (*.cfg);;\xE6\x89\x80\xE6\x9C\x89\xE6\x96\x87\xE4\xBB\xB6 (*.*)")
    );

    if (!filename.isEmpty()) {
        if (processor_->loadConfiguration(filename.toStdString())) {
            currentConfigFile_ = filename;
            updateRuleTable();
            // setupTaskTree(); // Refresh Task Table Headers - Not needed if setup once
            updateTaskTree(); // Refresh Task Data
            statusLabel_->setText(QString::fromUtf8("\xE9\x85\x8D\xE7\xBD\xAE\xE5\x8A\xA0\xE8\xBD\xBD\xE6\x88\x90\xE5\x8A\x9F"));
            
            if (configNameLabel_) {
                QString configName = QFileInfo(currentConfigFile_).fileName();
                configNameLabel_->setText(QString::fromUtf8("\xE5\xBD\x93\xE5\x89\x8D\xE9\x85\x8D\xE7\xBD\xAE: ") + configName);
                configNameLabel_->setStyleSheet("color: black; font-weight: bold; margin-right: 10px;");
            }
        } else {
            QMessageBox::warning(this, QString::fromUtf8("\xE9\x94\x99\xE8\xAF\xAF"), QString::fromUtf8("\xE9\x85\x8D\xE7\xBD\xAE\xE6\x96\x87\xE4\xBB\xB6\xE5\x8A\xA0\xE8\xBD\xBD\xE5\xA4\xB1\xE8\xB4\xA5"));
            statusLabel_->setText(QString::fromUtf8("\xE9\x85\x8D\xE7\xBD\xAE\xE5\x8A\xA0\xE8\xBD\xBD\xE5\xA4\xB1\xE8\xB4\xA5"));
        }
    }
}

void ExcelProcessorGUI::saveConfiguration() {
    QString filename = QFileDialog::getSaveFileName(
        this, QString::fromUtf8("\xE4\xBF\x9D\xE5\xAD\x98\xE9\x85\x8D\xE7\xBD\xAE"), "", QString::fromUtf8("\xE9\x85\x8D\xE7\xBD\xAE\xE6\x96\x87\xE4\xBB\xB6 (*.cfg);;\xE6\x89\x80\xE6\x9C\x89\xE6\x96\x87\xE4\xBB\xB6 (*.*)")
    );

    if (!filename.isEmpty()) {
        if (processor_->saveConfiguration(filename.toStdString())) {
            currentConfigFile_ = filename;
            statusLabel_->setText(QString::fromUtf8("\xE9\x85\x8D\xE7\xBD\xAE\xE4\xBF\x9D\xE5\xAD\x98\xE6\x88\x90\xE5\x8A\x9F")); // Success
            
            if (configNameLabel_) {
                QString configName = QFileInfo(currentConfigFile_).fileName();
                configNameLabel_->setText(QString::fromUtf8("\xE5\xBD\x93\xE5\x89\x8D\xE9\x85\x8D\xE7\xBD\xAE: ") + configName);
                configNameLabel_->setStyleSheet("color: black; font-weight: bold; margin-right: 10px;");
            }
        } else {
            QMessageBox::warning(this, QString::fromUtf8("\xE9\x94\x99\xE8\xAF\xAF"), QString::fromUtf8("\xE9\x85\x8D\xE7\xBD\xAE\xE6\x96\x87\xE4\xBB\xB6\xE4\xBF\x9D\xE5\xAD\x98\xE5\xA4\xB1\xE8\xB4\xA5"));
            statusLabel_->setText(QString::fromUtf8("\xE9\x85\x8D\xE7\xBD\xAE\xE4\xBF\x9D\xE5\xAD\x98\xE5\xA4\xB1\xE8\xB4\xA5"));
        }
    }
}



void ExcelProcessorGUI::addRule() {
    Rule rule;
    rule.id = generateRuleId();
    rule.name = QString::fromUtf8("\xE6\x96\xB0\xE8\xA7\x84\xE5\x88\x99 ").toStdString() + std::to_string(rule.id);
    rule.type = RuleType::FILTER;
    rule.enabled = true;

    // Add default condition
    RuleCondition condition;
    condition.column = 1;
    condition.oper = Operator::EQUAL;
    condition.value = std::string("");
    rule.conditions.push_back(condition);

    processor_->addRule(rule);
    updateRuleTable();
    statusLabel_->setText(QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99\xE6\xB7\xBB\xE5\x8A\xA0\xE6\x88\x90\xE5\x8A\x9F"));
}

void ExcelProcessorGUI::deleteRule() {
    int currentRow = ruleTable_->currentRow();
    if (currentRow >= 0) {
        int ruleId = ruleTable_->item(currentRow, 0)->text().toInt();
        processor_->removeRule(ruleId);
        updateRuleTable();
        statusLabel_->setText(QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99\xE5\x88\xA0\xE9\x99\xA4\xE6\x88\x90\xE5\x8A\x9F"));
    }
}

void ExcelProcessorGUI::editRule() {
    int currentRow = ruleTable_->currentRow();
    if (currentRow >= 0) {
        int ruleId = ruleTable_->item(currentRow, 0)->text().toInt();
        showRuleEditDialog(ruleId);
    }
}



void ExcelProcessorGUI::processFile() {
    auto tasks = processor_->getTasks();
    bool hasTasks = !tasks.empty();

    if (!hasTasks) {
         QMessageBox::warning(this, QString::fromUtf8("\xE4\xBB\xBB\xE5\x8A\xA1\xE5\x88\x97\xE8\xA1\xA8\xE4\xB8\xBA\xE7\xA9\xBA"), 
                              QString::fromUtf8("\xE4\xBB\xBB\xE5\x8A\xA1\xE5\x88\x97\xE8\xA1\xA8\xE4\xB8\xBA\xE7\xA9\xBA\xEF\xBC\x8C\xE4\xB8\x8D\xE6\x89\xA7\xE8\xA1\x8C\xE6\x95\xB0\xE6\x8D\xAE\xE5\xA4\x84\xE7\x90\x86\xE3\x80\x82\nPlease add tasks to the task list."));
         return;
    }

    // Disable process button, show progress
    processButton_->setEnabled(false);
    progressBar_->setVisible(true);
    progressBar_->setRange(0, 0); // Indeterminate mode for large files
    if (logTextEdit_) logTextEdit_->clear();
    
    // Clear previous errors
    processor_->clearErrors();

    // Process in background thread
    std::thread([this]() {
#ifdef _WIN32
        CoInitialize(nullptr);
#endif
        auto startTime = std::chrono::high_resolution_clock::now(); // Start timer
        
        ProcessingResult totalResult;
        std::vector<ProcessingResult> results;
        
        results = processor_->processTasks("", "", "");
        
        auto endTime = std::chrono::high_resolution_clock::now(); // End timer
        double totalDuration = std::chrono::duration<double, std::milli>(endTime - startTime).count() / 1000.0;

        // Aggregate results for simple display
        if (!results.empty()) {
            totalResult.totalRows = results[0].totalRows;
            for (const auto& r : results) {
                totalResult.processedRows += r.processedRows;
                totalResult.matchedRows += r.matchedRows;
                // totalResult.processingTime += r.processingTime; // Don't sum task times
                totalResult.errors.insert(totalResult.errors.end(), r.errors.begin(), r.errors.end());
            }
        }
        totalResult.processingTime = totalDuration; // Use real wall-clock time

        // Check for any global errors (e.g. file read failures)
        auto globalErrors = processor_->getErrors();
        if (!globalErrors.empty()) {
             for (const auto& err : globalErrors) {
                 bool exists = false;
                 for (const auto& e : totalResult.errors) {
                     if (e == err) { exists = true; break; }
                 }
                 if (!exists) totalResult.errors.push_back(err);
             }
        }

#ifdef _WIN32
        CoUninitialize();
#endif

        // Update UI in main thread
        QMetaObject::invokeMethod(this, [this, totalResult, results]() {
            processButton_->setEnabled(true);
            progressBar_->setVisible(false);
            progressBar_->setRange(0, 100); // Reset range

            std::ostringstream ss;
            // Multi-task processing complete!
            ss << "\xE5\xA4\x9A\xE4\xBB\xBB\xE5\x8A\xA1\xE5\xA4\x84\xE7\x90\x86\xE5\xAE\x8C\xE6\x88\x90\xEF\xBC\x81\n"
               // Total time:
               << "\xE6\x80\xBB\xE8\x80\x97\xE6\x97\xB6: " << std::fixed << std::setprecision(3) << totalResult.processingTime << "\xE7\xA7\x92\n" // seconds
               // Executed tasks:
               << "\xE6\x89\xA7\xE8\xA1\x8C\xE4\xBB\xBB\xE5\x8A\xA1\xE6\x95\xB0: " << results.size() << "\n"
               // Matched rows:
               << "\xE7\xAC\xA6\xE5\x90\x88\xE8\xA7\x84\xE5\x88\x99\xE8\xA1\x8C\xE6\x95\xB0: " << totalResult.matchedRows;

            // Log to text edit
            if (logTextEdit_) {
                QString timestamp = QDateTime::currentDateTime().toString("[yyyy-MM-dd hh:mm:ss] ");
                logTextEdit_->append(timestamp + QString::fromStdString(ss.str()));
            }

            if (!totalResult.errors.empty()) {
                // Errors occurred during processing:
                QString errorList = QString::fromUtf8("\xE5\xA4\x84\xE7\x90\x86\xE4\xB8\xAD\xE5\x8F\x91\xE7\x94\x9F\xE9\x94\x99\xE8\xAF\xAF\xEF\xBC\x9A\n");
                for (const auto& error : totalResult.errors) {
                    errorList += QString::fromUtf8("\xE2\x80\xA2 ") + QString::fromStdString(error) + "\n";
                }
                // Processing Errors
                QMessageBox::warning(this, QString::fromUtf8("\xE5\xA4\x84\xE7\x90\x86\xE9\x94\x99\xE8\xAF\xAF"), errorList);
            } else {
                 QString msg = QString::fromUtf8("\xE5\xA4\x84\xE7\x90\x86\xE5\xAE\x8C\xE6\x88\x90\xEF\xBC\x81\n\xE6\x80\xBB\xE8\x80\x97\xE6\x97\xB6: %1 \xE7\xA7\x92\n\xE5\x85\xB1\xE5\xA4\x84\xE7\x90\x86 %2 \xE4\xB8\xAA\xE4\xBB\xBB\xE5\x8A\xA1")
                    .arg(totalResult.processingTime, 0, 'f', 2)
                    .arg(results.size());
                 QMessageBox::information(this, QString::fromUtf8("\xE5\xA4\x84\xE7\x90\x86\xE5\xAE\x8C\xE6\x88\x90"), msg);
            }
        });
    }).detach();
}

/*
void ExcelProcessorGUI::createRuleCombination() {
    if (getSelectedRuleIds().empty()) {
        QMessageBox::information(this, QString::fromUtf8("\xE6\x8F\x90\xE7\xA4\xBA"), QString::fromUtf8("\xE8\xAF\xB7\xE5\x85\x88\xE9\x80\x89\xE6\x8B\xA9\xE8\xA6\x81\xE7\xBB\x84\xE5\x90\x88\xE7\x9A\x84\xE8\xA7\x84\xE5\x88\x99"));
        return;
    }

    bool ok;
    int comboId = QInputDialog::getInt(
        this, QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99\xE7\xBB\x84\xE5\x90\x88"), QString::fromUtf8("\xE8\xAF\xB7\xE8\xBE\x93\xE5\x85\xA5\xE7\xBB\x84\xE5\x90\x88ID:"), generateComboId(), 1, 9999, 1, &ok
    );

    if (ok) {
        processor_->createRuleCombination(
            comboId,
            getSelectedRuleIds(),
            CombinationStrategy::AND
        );
        updateCombinationTable();
        statusLabel_->setText(QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99\xE7\xBB\x84\xE5\x90\x88\xE5\x88\x9B\xE5\xBB\xBA\xE6\x88\x90\x9F"));
    }
}
*/

void ExcelProcessorGUI::validateRules() {
    auto rules = processor_->getRules();
    std::vector<std::string> allErrors;

    for (const auto& rule : rules) {
        auto errors = processor_->getValidationMessages();
        allErrors.insert(allErrors.end(), errors.begin(), errors.end());
    }

    if (allErrors.empty()) {
        QMessageBox::information(this, QString::fromUtf8("\xE9\xAA\x8C\xE8\xAF\x81\xE7\xBB\x93\xE6\x9E\x9C"), QString::fromUtf8("\xE6\x89\x80\xE6\x9C\x89\xE8\xA7\x84\xE5\x88\x99\xE9\xAA\x8C\xE8\xAF\x81\xE9\x80\x9A\xE8\xBF\x87"));
    } else {
        QString errorList = QString::fromUtf8("\xE5\x8F\x91\xE7\x8E\xB0\xE4\xBB\xA5\xE4\xB8\x8B\xE9\x97\xAE\xE9\xA2\x98\xEF\xBC\x9A\n");
        for (const auto& error : allErrors) {
            errorList += QString::fromUtf8("\xE2\x80\xA2 ") + QString::fromStdString(error) + "\n";
        }
        QMessageBox::warning(this, QString::fromUtf8("\xE9\xAA\x8C\xE8\xAF\x81\xE7\xBB\x93\xE6\x9E\x9C"), errorList);
    }
}

void ExcelProcessorGUI::tabChanged(int index) {
    if (index == 0) { // Task Config Tab
        updateTaskTree();
    }
    // else if (index == 3) { // Rule Combination Tab REMOVED
    //    updateCombinationTable();
    // }
}

/*
void ExcelProcessorGUI::updateProgress() {
    if (!processButton_->isEnabled()) {
        int value = (progressBar_->value() + 10) % 100;
        progressBar_->setValue(value);

        if (progressBar_->isVisible()) {
            QTimer::singleShot(100, this, &ExcelProcessorGUI::updateProgress);
        }
    }
}
*/

void ExcelProcessorGUI::setupUI() {
    setWindowTitle(QString::fromUtf8("\x45\x78\x63\x65\x6C\xE6\x95\xB0\xE6\x8D\xAE\xE5\xA4\x84\xE7\x90\x86\xE5\xB7\xA5\xE5\x85\xB7\x2D\xE6\xA8\xAA\xE6\xA2\x81\xE4\xB8\x8B\xE6\x96\x99\xE4\xB8\x93\xE4\xB8\x9A\xE7\xBB\x84")); // Excel Data Processing Tool - Crossbeam Material Group
    setMinimumSize(1200, 800);

    auto centralWidget = new QWidget(this);
    setCentralWidget(centralWidget);

    // Main Layout
    auto mainLayout = new QVBoxLayout(centralWidget);

    // File Selection Group REMOVED
    
    // Create Tab Widget
    tabWidget_ = new QTabWidget();
    mainLayout->addWidget(tabWidget_);

    // Task Config Tab (Moved to First Tab)
    auto taskTab = createTaskTab();
    // Multi-workbook Tasks
    tabWidget_->addTab(taskTab, QString::fromUtf8("\xE5\xA4\x9A\xE5\xB7\xA5\xE4\xBD\x9C\xE7\xB0\xBF\xE4\xBB\xBB\xE5\x8A\xA1"));

    // Rule Config
    auto ruleTab = createRuleTab();
    tabWidget_->addTab(ruleTab, QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99\xE9\x85\x8D\xE7\xBD\xAE"));

    // Data Preview Tab
    auto previewTab = createPreviewTab();
    // Data Preview
    tabWidget_->addTab(previewTab, QString::fromUtf8("\xE6\x95\xB0\xE6\x8D\xAE\xE9\xA2\x84\xE8\xA7\x88"));

    // Rule Combination Tab REMOVED
    /*
    auto combinationTab = createCombinationTab();
    tabWidget_->addTab(combinationTab, QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99\xE7\xBB\x84\xE5\x90\x88"));
    */

    // About Tab
    auto aboutTab = createAboutTab();
    tabWidget_->addTab(aboutTab, QString::fromUtf8("\xE5\x85\xB3\xE4\xBA\x8E"));

    // Log Output
    auto logGroup = new QGroupBox(QString::fromUtf8("\xE6\x97\xA5\xE5\xBF\x97\xE8\xBE\x93\xE5\x87\xBA")); // Log Output
    logGroup->setMaximumHeight(120); // Limit group box height
    auto logLayout = new QVBoxLayout(logGroup);
    logTextEdit_ = new QTextEdit();
    logTextEdit_->setReadOnly(true);
    // logTextEdit_->setMaximumHeight(80); // Controlled by parent layout now
    logLayout->addWidget(logTextEdit_);
    mainLayout->addWidget(logGroup);

    // Operation Button Group
    auto buttonGroup = createButtonGroup();
    mainLayout->addWidget(buttonGroup);

    // Status Bar
    statusBar_ = new QStatusBar();
    setStatusBar(statusBar_);
    // Ready
    statusLabel_ = new QLabel(QString::fromUtf8("\xE5\xB0\xB1\xE7\xBB\xAA"));
    progressBar_ = new QProgressBar();
    progressBar_->setVisible(false);
    statusBar_->addWidget(statusLabel_);
    statusBar_->addWidget(progressBar_);

    // Connect tab change signal
    connect(tabWidget_, &QTabWidget::currentChanged, this, &ExcelProcessorGUI::tabChanged);

    updateRuleTable();
}

// Removed createFileSelectionGroup implementation

QWidget* ExcelProcessorGUI::createAboutTab() {
    auto tab = new QWidget();
    auto layout = new QVBoxLayout(tab);

    // Title
    auto titleLabel = new QLabel(QString::fromUtf8("\x45\x78\x63\x65\x6C\xE6\x95\xB0\xE6\x8D\xAE\xE5\xA4\x84\xE7\x90\x86\xE5\xB7\xA5\xE5\x85\xB7\x2D\xE6\xA8\xAA\xE6\xA2\x81\xE4\xB8\x8B\xE6\x96\x99\xE4\xB8\x93\xE4\xB8\x9A\xE7\xBB\x84"));
    titleLabel->setStyleSheet("font-size: 24px; font-weight: bold; color: #333333; margin-top: 20px;");
    titleLabel->setAlignment(Qt::AlignCenter);
    layout->addWidget(titleLabel);

    // Version
    auto versionLabel = new QLabel(QString::fromUtf8("\xE7\x89\x88\xE6\x9C\xAC: v1.0.0")); // Version: v1.0.0
    versionLabel->setStyleSheet("font-size: 16px; color: #666666; margin-top: 10px;");
    versionLabel->setAlignment(Qt::AlignCenter);
    layout->addWidget(versionLabel);

    // Description
    auto descLabel = new QLabel(QString::fromUtf8("\x45\x78\x63\x65\x6C\xE6\x95\xB0\xE6\x8D\xAE\xE5\xBF\xAB\xE9\x80\x9F\xE5\xA4\x84\xE7\x90\x86\xEF\xBC\x8C\xE8\x87\xAA\xE5\xAE\x9A\xE4\xB9\x89\xE5\xA4\x8D\xE6\x9D\x82\xE8\xA7\x84\xE5\x88\x99\xE8\xBE\x93\xE5\x87\xBA\xE6\x95\xB0\xE6\x8D\xAE\xE3\x80\x82"));
    descLabel->setStyleSheet("font-size: 14px; color: #555555; margin-top: 20px;");
    descLabel->setAlignment(Qt::AlignCenter);
    layout->addWidget(descLabel);

    // User Guide Group
    auto guideGroup = new QGroupBox(QString::fromUtf8("\xE7\xAE\x80\xE6\x98\x93\xE4\xBD\xBF\xE7\x94\xA8\xE8\xAF\xB4\xE6\x98\x8E")); // Simple User Guide
    guideGroup->setStyleSheet("QGroupBox { font-weight: bold; border: 1px solid #cccccc; border-radius: 5px; margin-top: 20px; padding-top: 15px; } QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top center; padding: 0 5px; color: #333333; }");
    auto guideLayout = new QVBoxLayout(guideGroup);

    auto guideLabel = new QLabel(QString::fromUtf8(
        "1. \xE9\xA6\x96\xE5\x85\x88\xE5\x9C\xA8 <b>\xE8\xA7\x84\xE5\x88\x99\xE9\x85\x8D\xE7\xBD\xAE</b> \xE7\x95\x8C\xE9\x9D\xA2\xE5\x88\x9B\xE5\xBB\xBA\xE8\xA7\x84\xE5\x88\x99\xE3\x80\x82<br>"
        "2. \xE7\x84\xB6\xE5\x90\x8E\xE5\x9C\xA8 <b>\xE5\xA4\x9A\xE5\xB7\xA5\xE4\xBD\x9C\xE7\xB0\xBF\xE4\xBB\xBB\xE5\x8A\xA1</b> \xE7\x95\x8C\xE9\x9D\xA2\xE4\xBD\xBF\xE7\x94\xA8\xE8\xBF\x99\xE4\xBA\x9B\xE8\xA7\x84\xE5\x88\x99\xE3\x80\x82<br>"
        "3. \xE5\x8F\xAF\xE4\xBB\xA5\xE4\xBF\x9D\xE5\xAD\x98\xE5\xA4\x9A\xE4\xB8\xAA\xE9\x85\x8D\xE7\xBD\xAE\xE6\x96\x87\xE4\xBB\xB6\xE4\xBB\xA5\xE5\xBA\x94\xE5\xAF\xB9\xE4\xB8\x8D\xE5\x90\x8C\xE7\x9A\x84\xE9\x9C\x80\xE6\xB1\x82\xE3\x80\x82"
    ));
    guideLabel->setStyleSheet("font-size: 14px; color: #333333; font-weight: normal;");
    guideLabel->setAlignment(Qt::AlignCenter);
    guideLayout->addWidget(guideLabel);

    // Center the group box horizontally
    auto hLayout = new QHBoxLayout();
    hLayout->addStretch();
    hLayout->addWidget(guideGroup);
    hLayout->addStretch();
    layout->addLayout(hLayout);

    layout->addStretch();

    // Developer Info
    auto devLabel = new QLabel(QString::fromUtf8("\xE5\xBC\x80\xE5\x8F\x91\xE8\x80\x85\xE4\xBF\xA1\xE6\x81\xAF")); // Developer Info
    devLabel->setStyleSheet("font-size: 18px; font-weight: bold; color: #333333;");
    devLabel->setAlignment(Qt::AlignCenter);
    layout->addWidget(devLabel);

    auto emailLabel = new QLabel("Email: songwj@yutong.com");
    emailLabel->setStyleSheet("font-size: 16px; color: #0066cc; margin-bottom: 40px;");
    emailLabel->setAlignment(Qt::AlignCenter);
    emailLabel->setTextInteractionFlags(Qt::TextSelectableByMouse); // Allow copying email
    layout->addWidget(emailLabel);

    return tab;
}

QWidget* ExcelProcessorGUI::createRuleTab() {
    auto widget = new QWidget();
    auto layout = new QVBoxLayout(widget);

    // Rule Operation Buttons
    auto ruleButtonLayout = new QHBoxLayout();

    auto addRuleBtn = new QPushButton(QString::fromUtf8("\xE6\xB7\xBB\xE5\x8A\xA0\xE8\xA7\x84\xE5\x88\x99")); // Add Rule
    auto editRuleBtn = new QPushButton(QString::fromUtf8("\xE7\xBC\x96\xE8\xBE\x91\xE8\xA7\x84\xE5\x88\x99")); // Edit Rule
    auto deleteRuleBtn = new QPushButton(QString::fromUtf8("\xE5\x88\xA0\xE9\x99\xA4\xE8\xA7\x84\xE5\x88\x99")); // Delete Rule
    auto validateBtn = new QPushButton(QString::fromUtf8("\xE9\xAA\x8C\xE8\xAF\x81\xE8\xA7\x84\xE5\x88\x99")); // Validate Rules
    auto loadConfigBtn = new QPushButton(QString::fromUtf8("\xE5\x8A\xA0\xE8\xBD\xBD\xE9\x85\x8D\xE7\xBD\xAE")); // Load Config
    auto saveConfigBtn = new QPushButton(QString::fromUtf8("\xE4\xBF\x9D\xE5\xAD\x98\xE9\x85\x8D\xE7\xBD\xAE")); // Save Config

    // Apply consistent style (Height 35, Add button green/bold)
    int btnHeight = 35;
    addRuleBtn->setMinimumHeight(btnHeight);
    editRuleBtn->setMinimumHeight(btnHeight);
    deleteRuleBtn->setMinimumHeight(btnHeight);
    validateBtn->setMinimumHeight(btnHeight);
    loadConfigBtn->setMinimumHeight(btnHeight);
    saveConfigBtn->setMinimumHeight(btnHeight);

    addRuleBtn->setStyleSheet("font-weight: bold; color: green;");

    configNameLabel_ = new QLabel(QString::fromUtf8("\xE6\x9C\xAA\xE5\x8A\xA0\xE8\xBD\xBD\xE9\x85\x8D\xE7\xBD\xAE")); // Unloaded Config
    configNameLabel_->setStyleSheet("color: gray; font-style: italic; margin-right: 10px;");

    ruleButtonLayout->addWidget(addRuleBtn);
    ruleButtonLayout->addWidget(editRuleBtn);
    ruleButtonLayout->addWidget(deleteRuleBtn);
    ruleButtonLayout->addWidget(validateBtn);
    ruleButtonLayout->addStretch();
    ruleButtonLayout->addWidget(configNameLabel_);
    ruleButtonLayout->addWidget(loadConfigBtn);
    ruleButtonLayout->addWidget(saveConfigBtn);

    layout->addLayout(ruleButtonLayout);

    // Rule Table
    ruleTable_ = new QTableWidget();
    setupRuleTable();
    layout->addWidget(ruleTable_);

    // Connect signals
    connect(addRuleBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::addRule);
    connect(editRuleBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::editRule);
    connect(deleteRuleBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::deleteRule);
    connect(validateBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::validateRules);
    connect(loadConfigBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::loadConfiguration);
    connect(saveConfigBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::saveConfiguration);

    return widget;
}

QWidget* ExcelProcessorGUI::createPreviewTab() {
    auto widget = new QWidget();
    auto layout = new QVBoxLayout(widget);

    // Info Labels
    auto infoLayout = new QHBoxLayout();
    previewInputLabel_ = new QLabel(QString::fromUtf8("\xE8\xBE\x93\xE5\x85\xA5\xE6\xBA\x90: \xE6\x9C\xAA\xE9\x80\x89\xE6\x8B\xA9")); // Input Source: Not Selected
    previewOutputLabel_ = new QLabel(QString::fromUtf8("\xE8\xBE\x93\xE5\x87\xBA\xE8\xA7\x84\xE5\x88\x99: \xE6\x9C\xAA\xE9\x80\x89\xE6\x8B\xA9")); // Output Rule: Not Selected
    
    previewInputLabel_->setStyleSheet("font-weight: bold; color: #333; margin-right: 20px;");
    previewOutputLabel_->setStyleSheet("font-weight: bold; color: #333;");

    infoLayout->addWidget(previewInputLabel_);
    infoLayout->addWidget(previewOutputLabel_);
    infoLayout->addStretch();
    layout->addLayout(infoLayout);

    auto buttonLayout = new QHBoxLayout();
    auto refreshBtn = new QPushButton(QString::fromUtf8("\xE5\x88\xB7\xE6\x96\xB0\xE9\xA2\x84\xE8\xA7\x88")); // Refresh Preview
    auto previewRulesBtn = new QPushButton(QString::fromUtf8("\xE9\xA2\x84\xE8\xA7\x88\xE8\xA7\x84\xE5\x88\x99\xE7\xBB\x93\xE6\x9E\x9C")); // Preview Rules Result
    buttonLayout->addWidget(refreshBtn);
    buttonLayout->addWidget(previewRulesBtn);
    buttonLayout->addStretch();
    layout->addLayout(buttonLayout);

    previewTable_ = new QTableWidget();
    layout->addWidget(previewTable_);

    connect(refreshBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::previewData);
    connect(previewRulesBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::previewProcessedData);

    return widget;
}

/*
QWidget* ExcelProcessorGUI::createCombinationTab() {
    auto widget = new QWidget();
    auto layout = new QVBoxLayout(widget);

    auto buttonLayout = new QHBoxLayout();
    auto createComboBtn = new QPushButton(QString::fromUtf8("\xE5\x88\x9B\xE5\xBB\xBA\xE8\xA7\x84\xE5\x88\x99\xE7\xBB\x84\xE5\x90\x88")); // Create Rule Combination
    buttonLayout->addWidget(createComboBtn);
    buttonLayout->addStretch();
    layout->addLayout(buttonLayout);

    combinationTable_ = new QTableWidget();
    setupCombinationTable();
    layout->addWidget(combinationTable_);

    connect(createComboBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::createRuleCombination);

    return widget;
}
*/

QGroupBox* ExcelProcessorGUI::createButtonGroup() {
    auto group = new QGroupBox(QString::fromUtf8("\xE6\x93\x8D\xE4\xBD\x9C")); // Operation
    auto layout = new QHBoxLayout(group);

    // Only "Start Processing" button remains, "Preview Data" removed as requested
    auto processBtn = new QPushButton(QString::fromUtf8("\xE5\xBC\x80\xE5\xA7\x8B\xE5\xA4\x84\xE7\x90\x86")); // Start Processing

    processBtn->setMinimumHeight(50);
    // Adjusted style: slightly larger font since it's the only button now
    processBtn->setStyleSheet("font-size: 16px; font-weight: bold; background-color: #4CAF50; color: white;");

    processButton_ = processBtn; // Save member variable

    layout->addWidget(processBtn);

    connect(processBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::processFile);

    return group;
}

void ExcelProcessorGUI::setupConnections() {
    // Connect signals and slots
}

void ExcelProcessorGUI::setupRuleTable() {
    QStringList headers;
    headers << "ID" 
            << QString::fromUtf8("\xE5\x90\x8D\xE7\xA7\xB0") // Name
            << QString::fromUtf8("\xE7\xB1\xBB\xE5\x9E\x8B") // Type
            << QString::fromUtf8("\xE6\x9D\xA1\xE4\xBB\xB6\xE6\x95\xB0") // Condition Count
            // << QString::fromUtf8("\xE7\x9B\xAE\xE6\xA0\x87\xE8\xA1\xA8") // Target Sheet REMOVED
            << QString::fromUtf8("\xE5\x90\xAF\xE7\x94\xA8") // Enabled
            << QString::fromUtf8("\xE6\x8F\x8F\xE8\xBF\xB0") // Description
            << QString::fromUtf8("\xE4\xBC\x98\xE5\x85\x88\xE7\xBA\xA7"); // Priority
    ruleTable_->setColumnCount(headers.size());
    ruleTable_->setHorizontalHeaderLabels(headers);
    ruleTable_->setSelectionBehavior(QAbstractItemView::SelectRows);
    ruleTable_->setAlternatingRowColors(true);
    ruleTable_->setSortingEnabled(true);

    // Optimize column widths
    QHeaderView* header = ruleTable_->horizontalHeader();
    header->setSectionResizeMode(0, QHeaderView::ResizeToContents); // ID
    header->setSectionResizeMode(1, QHeaderView::Stretch);          // Name
    header->setSectionResizeMode(2, QHeaderView::ResizeToContents); // Type
    header->setSectionResizeMode(3, QHeaderView::ResizeToContents); // Condition Count
    header->setSectionResizeMode(4, QHeaderView::ResizeToContents); // Enabled
    header->setSectionResizeMode(5, QHeaderView::Stretch);          // Description
    header->setSectionResizeMode(6, QHeaderView::ResizeToContents); // Priority
}

/*
void ExcelProcessorGUI::setupCombinationTable() {
    QStringList headers;
    headers << QString::fromUtf8("\xE7\xBB\x84\xE5\x90\x88ID") // Combination ID
            << QString::fromUtf8("\xE5\x8C\x85\xE5\x90\xAB\xE8\xA7\x84\xE5\x88\x99") // Include Rule
            << QString::fromUtf8("\xE7\xAD\x96\xE7\x95\xA5"); // Strategy
    combinationTable_->setColumnCount(headers.size());
    combinationTable_->setHorizontalHeaderLabels(headers);
    combinationTable_->setSelectionBehavior(QAbstractItemView::SelectRows);
    combinationTable_->setAlternatingRowColors(true);
}
*/

void ExcelProcessorGUI::updateRuleTable() {
    auto rules = processor_->getRules();
    ruleTable_->setRowCount(rules.size());

    for (size_t i = 0; i < rules.size(); ++i) {
        const auto& rule = rules[i];

        ruleTable_->setItem(i, 0, new QTableWidgetItem(QString::number(rule.id)));
        ruleTable_->setItem(i, 1, new QTableWidgetItem(QString::fromStdString(rule.name)));
        ruleTable_->setItem(i, 2, new QTableWidgetItem(QString::fromStdString(getRuleTypeString(rule.type))));
        
        QTableWidgetItem* countItem = new QTableWidgetItem(QString::number(rule.conditions.size()));
        countItem->setTextAlignment(Qt::AlignCenter);
        ruleTable_->setItem(i, 3, countItem);

        // ruleTable_->setItem(i, 4, new QTableWidgetItem(QString::fromStdString(rule.targetSheet))); // REMOVED
        
        QTableWidgetItem* enabledItem = new QTableWidgetItem(rule.enabled ? QString::fromUtf8("\xE6\x98\xAF") : QString::fromUtf8("\xE5\x90\xA6"));
        enabledItem->setTextAlignment(Qt::AlignCenter);
        ruleTable_->setItem(i, 4, enabledItem);
        
        // Generate Description automatically
        QString desc = generateRuleDescription(rule);
        ruleTable_->setItem(i, 5, new QTableWidgetItem(desc));
        ruleTable_->setItem(i, 6, new QTableWidgetItem(QString::number(rule.priority)));
    }
}

void ExcelProcessorGUI::updatePreviewTable(const std::vector<DataRow>& data, const QStringList& customHeaders) {
    if (data.empty()) {
        previewTable_->setRowCount(0);
        previewTable_->setColumnCount(0);
        return;
    }

    size_t columnCount = data[0].data.size();
    previewTable_->setRowCount(data.size());
    previewTable_->setColumnCount(columnCount);

    // Set column headers
    if (!customHeaders.empty()) {
         previewTable_->setHorizontalHeaderLabels(customHeaders);
    } else {
        QStringList headers;
        for (size_t i = 0; i < columnCount; ++i) {
            headers << QString::fromUtf8("\xE5\x88\x97%1").arg(i + 1);
        }
        previewTable_->setHorizontalHeaderLabels(headers);
    }

    // Populate data
    for (size_t i = 0; i < data.size(); ++i) {
        for (size_t j = 0; j < columnCount; ++j) {
            std::string cellValue;
            if (j < data[i].data.size()) { // Safety check
                std::visit([&cellValue](const auto& val) {
                    using ValType = std::decay_t<decltype(val)>;
                    if constexpr (std::is_same_v<ValType, std::string>) {
                        cellValue = val;
                    } else if constexpr (std::is_arithmetic_v<ValType>) {
                        // Remove trailing zeros for display if double
                        if constexpr (std::is_floating_point_v<ValType>) {
                             std::string s = std::to_string(val);
                             s.erase(s.find_last_not_of('0') + 1, std::string::npos);
                             if (s.back() == '.') s.pop_back();
                             cellValue = s;
                        } else {
                             cellValue = std::to_string(val);
                        }
                    } else if constexpr (std::is_same_v<ValType, bool>) {
                        cellValue = val ? "true" : "false";
                    }
                }, data[i].data[j]);
            }

            previewTable_->setItem(i, j, new QTableWidgetItem(QString::fromStdString(cellValue)));
        }
    }

    previewTable_->resizeColumnsToContents();
}

/*
void ExcelProcessorGUI::updateCombinationTable() {
    auto combinations = processor_->getRuleCombinations();
    combinationTable_->setRowCount(combinations.size());

    int row = 0;
    for (const auto& [comboId, ruleIds] : combinations) {
        combinationTable_->setItem(row, 0, new QTableWidgetItem(QString::number(comboId)));

        QStringList ruleIdStrings;
        for (int ruleId : ruleIds) {
            ruleIdStrings << QString::number(ruleId);
        }
        combinationTable_->setItem(row, 1, new QTableWidgetItem(ruleIdStrings.join(", ")));

        combinationTable_->setItem(row, 2, new QTableWidgetItem("AND"));

        ++row;
    }

    combinationTable_->resizeColumnsToContents();
}
*/

std::vector<int> ExcelProcessorGUI::getSelectedRuleIds() {
    std::vector<int> selectedIds;
    for (int i = 0; i < ruleTable_->rowCount(); ++i) {
        if (ruleTable_->selectionModel()->isRowSelected(i, QModelIndex())) {
            selectedIds.push_back(ruleTable_->item(i, 0)->text().toInt());
        }
    }
    return selectedIds;
}

int ExcelProcessorGUI::generateRuleId() {
    auto rules = processor_->getRules();
    int maxId = 0;
    for (const auto& rule : rules) {
        if (rule.id > maxId) maxId = rule.id;
    }
    return maxId + 1;
}

/*
int ExcelProcessorGUI::generateComboId() {
    auto combinations = processor_->getRuleCombinations();
    int maxId = 0;
    for (const auto& [comboId, ruleIds] : combinations) {
        if (comboId > maxId) maxId = comboId;
    }
    return maxId + 1;
}
*/

void ExcelProcessorGUI::showRuleEditDialog(int ruleId) {
    Rule rule = processor_->getRule(ruleId);
    auto allRules = processor_->getRules();
    RuleEditDialog dialog(rule, allRules, this);
    if (dialog.exec() == QDialog::Accepted) {
        Rule updatedRule = dialog.getRule();
        processor_->updateRule(updatedRule);
        updateRuleTable();
    }
}

// Helper function to generate rule description
QString ExcelProcessorGUI::generateRuleDescription(const Rule& rule) {
    if (rule.conditions.empty()) return QString::fromUtf8("\xE6\x97\xA0\xE6\x9D\xA1\xE4\xBB\xB6"); // "No condition"

    QStringList parts;
    for (const auto& cond : rule.conditions) {
        QString part = QString::fromUtf8("\xE5\x88\x97") + QString::number(cond.column) + " "; // "Col" + col
        
        // Add Operator
        switch (cond.oper) {
            case Operator::EQUAL: part += "="; break;
            case Operator::NOT_EQUAL: part += "<>"; break;
            case Operator::GREATER: part += ">"; break;
            case Operator::LESS: part += "<"; break;
            case Operator::GREATER_EQUAL: part += ">="; break;
            case Operator::LESS_EQUAL: part += "<="; break;
            case Operator::CONTAINS: part += QString::fromUtf8("\xE5\x8C\x85\xE5\x90\xAB"); break; // "Contains"
            case Operator::NOT_CONTAINS: part += QString::fromUtf8("\xE4\xB8\x8D\xE5\x8C\x85\xE5\x90\xAB"); break; // "Not Contains"
            case Operator::EMPTY: part += QString::fromUtf8("\xE4\xB8\xBA\xE7\xA9\xBA"); break; // "Is Empty"
            case Operator::NOT_EMPTY: part += QString::fromUtf8("\xE4\xB8\x8D\xE4\xB8\xBA\xE7\xA9\xBA"); break; // "Not Empty"
            case Operator::STARTS_WITH: part += QString::fromUtf8("\xE5\xBC\x80\xE5\xA4\xB4\xE6\x98\xAF"); break; // "Starts With"
            case Operator::ENDS_WITH: part += QString::fromUtf8("\xE7\xBB\x93\xE5\xB0\xBE\xE6\x98\xAF"); break; // "Ends With"
            case Operator::REGEX: part += "Regex"; break;
            default: part += "?"; break;
        }
        
        // Add Value (if not unary operator)
        if (cond.oper != Operator::EMPTY && cond.oper != Operator::NOT_EMPTY) {
            std::string valStr;
            std::visit([&valStr](const auto& val) {
                using ValType = std::decay_t<decltype(val)>;
                if constexpr (std::is_same_v<ValType, std::string>) {
                    valStr = "'" + val + "'";
                } else if constexpr (std::is_arithmetic_v<ValType>) {
                    valStr = std::to_string(val);
                } else if constexpr (std::is_same_v<ValType, bool>) {
                    valStr = val ? "True" : "False";
                }
            }, cond.value);
            part += " " + QString::fromStdString(valStr);
        }
        parts << part;
    }

    QString logicStr = (rule.logic == RuleLogic::AND) ? " AND " : " OR ";
    return parts.join(logicStr);
}

std::string ExcelProcessorGUI::getRuleTypeString(RuleType type) {
    switch (type) {
        case RuleType::FILTER: return "\xE7\xAD\x9B\xE9\x80\x89"; // Filter
        case RuleType::TRANSFORM: return "\xE8\xBD\xAC\xE6\x8D\xA2"; // Transform
        case RuleType::DELETE_ROW: return "\xE5\x88\xA0\xE9\x99\xA4"; // Delete
        case RuleType::SPLIT: return "\xE6\x8B\x86\xE5\x88\x86"; // Split
        default: return "\xE6\x9C\xAA\xE7\x9F\xA5"; // Unknown
    }
}

void ExcelProcessorGUI::addInputGroup() {
    QString pattern, sheet;
    if (showInputGroupDialog(pattern, sheet)) {
        showTaskEditDialog(-1, pattern, sheet);
    }
}

void ExcelProcessorGUI::addTaskToGroup() {
    QTreeWidgetItem* item = taskTree_->currentItem();
    if (!item) {
        QMessageBox::warning(this, QString::fromUtf8("\xE6\x8F\x90\xE7\xA4\xBA"), QString::fromUtf8("\xE8\xAF\xB7\xE5\x85\x88\xE9\x80\x89\xE6\x8B\xA9\xE4\xB8\x80\xE4\xB8\xAA\xE8\xBE\x93\xE5\x85\xA5\xE6\xBA\x90\xE6\x88\x96\xE4\xBB\xBB\xE5\x8A\xA1"));
        return;
    }

    QString pattern, sheet;
    if (item->parent() == nullptr) {
        pattern = item->data(0, Qt::UserRole).toString();
        sheet = item->data(0, Qt::UserRole + 1).toString();
    } else {
        QTreeWidgetItem* parent = item->parent();
        pattern = parent->data(0, Qt::UserRole).toString();
        sheet = parent->data(0, Qt::UserRole + 1).toString();
    }
    showTaskEditDialog(-1, pattern, sheet);
}

void ExcelProcessorGUI::previewData() {
    // Check if we are refreshing a specific task preview
    if (currentPreviewTaskId_ != -1) {
        previewTask(currentPreviewTaskId_);
        return;
    }

    // 1. Determine Input File
    QString inputPattern;
    QString inputSheet;
    
    QTreeWidgetItem* item = taskTree_->currentItem();
    if (item) {
        if (item->parent()) {
            // Child (Task)
            inputPattern = item->parent()->data(0, Qt::UserRole).toString();
            inputSheet = item->parent()->data(0, Qt::UserRole + 1).toString();
        } else {
            // Root (Input Group)
            inputPattern = item->data(0, Qt::UserRole).toString();
            inputSheet = item->data(0, Qt::UserRole + 1).toString();
        }
    } else {
        // Try to get first item if nothing selected
        if (taskTree_->topLevelItemCount() > 0) {
            item = taskTree_->topLevelItem(0);
            inputPattern = item->data(0, Qt::UserRole).toString();
            inputSheet = item->data(0, Qt::UserRole + 1).toString();
        } else {
             QMessageBox::warning(this, QString::fromUtf8("\xE6\x8F\x90\xE7\xA4\xBA"), QString::fromUtf8("\xE8\xAF\xB7\xE5\x85\x88\xE6\xB7\xBB\xE5\x8A\xA0\xE8\xBE\x93\xE5\x85\xA5\xE6\xBA\x90")); // Please add input source first
             return;
        }
    }

    // Resolve Pattern to File
    QString actualFile = inputPattern;
    if (actualFile.contains('*') || actualFile.contains('?')) {
        QFileInfo fi(actualFile);
        QDir dir = fi.absoluteDir();
        QStringList filters;
        filters << fi.fileName();
        dir.setNameFilters(filters);
        QStringList files = dir.entryList(QDir::Files);
        if (files.empty()) {
             QMessageBox::warning(this, QString::fromUtf8("\xE9\x94\x99\xE8\xAF\xAF"), QString::fromUtf8("\xE6\x9C\xAA\xE6\x89\xBE\xE5\x88\xB0\xE5\x8C\xB9\xE9\x85\x8D\xE7\x9A\x84\xE6\x96\x87\xE4\xBB\xB6")); // No matching file found
             return;
        }
        actualFile = dir.absoluteFilePath(files.first());
    }
    
    // Call Processor
    statusBar_->showMessage(QString::fromUtf8("\xE6\xAD\xA3\xE5\x9C\xA8\xE5\x8A\xA0\xE8\xBD\xBD\xE9\xA2\x84\xE8\xA7\x88\xE6\x95\xB0\xE6\x8D\xAE...")); // Loading preview data...
    QApplication::setOverrideCursor(Qt::WaitCursor);
    
    if (processor_->previewResults(actualFile.toStdString(), inputSheet.toStdString(), 5000)) {
         auto data = processor_->getPreviewData();
         updatePreviewTable(data);
         
         if (previewInputLabel_) previewInputLabel_->setText(QString::fromUtf8("\xE8\xBE\x93\xE5\x85\xA5\xE6\xBA\x90: ") + actualFile + " / " + inputSheet);
         if (previewOutputLabel_) previewOutputLabel_->setText(QString::fromUtf8("\xE8\xBE\x93\xE5\x87\xBA\xE8\xA7\x84\xE5\x88\x99: \xE6\x97\xA0 (\xE5\x8E\x9F\xE5\xA7\x8B\xE6\x95\xB0\xE6\x8D\xAE)"));
         
         statusBar_->showMessage(QString::fromUtf8("\xE9\xA2\x84\xE8\xA7\x88\xE6\x95\xB0\xE6\x8D\xAE\xE5\xB7\xB2\xE5\x8A\xA0\xE8\xBD\xBD")); // Preview data loaded
    } else {
         QMessageBox::warning(this, QString::fromUtf8("\xE9\x94\x99\xE8\xAF\xAF"), QString::fromUtf8("\xE6\x97\xA0\xE6\xB3\x95\xE8\xAF\xBB\xE5\x8F\x96\xE6\x96\x87\xE4\xBB\xB6")); // Unable to read file
         statusBar_->clearMessage();
    }
    QApplication::restoreOverrideCursor();
}

void ExcelProcessorGUI::previewProcessedData() {
    // Get selected rule IDs
    std::vector<int> selectedRules = getSelectedRuleIds();
    
    statusBar_->showMessage(QString::fromUtf8("\xE6\xAD\xA3\xE5\x9C\xA8\xE5\xA4\x84\xE7\x90\x86\xE9\xA2\x84\xE8\xA7\x88\xE6\x95\xB0\xE6\x8D\xAE..."));
    QApplication::setOverrideCursor(Qt::WaitCursor);
    
    auto data = processor_->getProcessedPreviewData(selectedRules);
    updatePreviewTable(data);
    
    statusBar_->showMessage(QString::fromUtf8("\xE5\xB7\xB2\xE5\xBA\x94\xE7\x94\xA8 %1 \xE6\x9D\xA1\xE8\xA7\x84\xE5\x88\x99").arg(selectedRules.size()));
    QApplication::restoreOverrideCursor();
}

void ExcelProcessorGUI::previewTask(int taskId) {
    auto tasks = processor_->getTasks();
    ProcessingTask task;
    bool found = false;
    for (const auto& t : tasks) {
        if (t.id == taskId) {
            task = t;
            found = true;
            break;
        }
    }
    if (!found) return;

    // Update current preview task ID
    currentPreviewTaskId_ = taskId;

    statusBar_->showMessage(QString::fromUtf8("\xE6\xAD\xA3\xE5\x9C\xA8\xE5\x8A\xA0\xE8\xBD\xBD\xE4\xBB\xBB\xE5\x8A\xA1\xE9\xA2\x84\xE8\xA7\x88...")); // Loading task preview...
    QApplication::setOverrideCursor(Qt::WaitCursor);

    // Resolve Input File (Assume pattern is a file or use first match)
    QString actualFile = QString::fromStdString(task.inputFilenamePattern);
    if (actualFile.contains('*') || actualFile.contains('?')) {
        QFileInfo fi(actualFile);
        QDir dir = fi.absoluteDir();
        QStringList filters;
        filters << fi.fileName();
        dir.setNameFilters(filters);
        QStringList files = dir.entryList(QDir::Files);
        if (!files.empty()) {
             actualFile = dir.absoluteFilePath(files.first());
        } else {
             QMessageBox::warning(this, QString::fromUtf8("\xE9\x94\x99\xE8\xAF\xAF"), QString::fromUtf8("\xE6\x9C\xAA\xE6\x89\xBE\xE5\x88\xB0\xE8\xBE\x93\xE5\x85\xA5\xE6\x96\x87\xE4\xBB\xB6: ") + actualFile);
             QApplication::restoreOverrideCursor();
             return;
        }
    }

    // Load Data if needed (Preview Load) 5000 is the max preview rows
    if (processor_->previewResults(actualFile.toStdString(), task.inputSheetName, 5000)) {
         auto data = processor_->getTaskPreviewData(taskId);
         
         // Handle Headers
         QStringList headers;
         if (task.useHeader && !data.empty()) {
             // Extract header from first row
             const auto& headerRow = data[0];
             for (const auto& cell : headerRow.data) {
                 std::string cellStr;
                 std::visit([&cellStr](const auto& val) {
                    using ValType = std::decay_t<decltype(val)>;
                    if constexpr (std::is_same_v<ValType, std::string>) {
                        cellStr = val;
                    } else if constexpr (std::is_arithmetic_v<ValType>) {
                        cellStr = std::to_string(val);
                    } else if constexpr (std::is_same_v<ValType, bool>) {
                        cellStr = val ? "true" : "false";
                    }
                 }, cell);
                 headers << QString::fromStdString(cellStr);
             }
             // Remove header row from data to display
             data.erase(data.begin());
         }

         updatePreviewTable(data, headers);
         
         if (previewInputLabel_) previewInputLabel_->setText(QString::fromUtf8("\xE8\xBE\x93\xE5\x85\xA5\xE6\xBA\x90: ") + actualFile + " / " + QString::fromStdString(task.inputSheetName));
         if (previewOutputLabel_) previewOutputLabel_->setText(QString::fromUtf8("\xE8\xBE\x93\xE5\x87\xBA\xE8\xA7\x84\xE5\x88\x99: ") + QString::fromStdString(task.taskName));

         // Switch to Preview Tab (Index 2: Task=0, Rule=1, Preview=2)
         if (tabWidget_) {
             tabWidget_->setCurrentIndex(2);
         }
         
         statusBar_->showMessage(QString::fromUtf8("\xE4\xBB\xBB\xE5\x8A\xA1\xE9\xA2\x84\xE8\xA7\x88\xE5\xB7\xB2\xE5\x8A\xA0\xE8\xBD\xBD: ") + QString::fromStdString(task.taskName));
    } else {
         QMessageBox::warning(this, QString::fromUtf8("\xE9\x94\x99\xE8\xAF\xAF"), QString::fromUtf8("\xE6\x97\xA0\xE6\xB3\x95\xE8\xAF\xBB\xE5\x8F\x96\xE6\x96\x87\xE4\xBB\xB6"));
         statusBar_->clearMessage();
    }
    QApplication::restoreOverrideCursor();
}

void ExcelProcessorGUI::executeSingleTask(int taskId) {
    auto task = processor_->getTask(taskId);
    if (task.id == 0) return; // Not found

    if (QMessageBox::question(this, QString::fromUtf8("\xE7\xA1\xAE\xE8\xAE\xA4"), 
        QString::fromUtf8("\xE7\xA1\xAE\xE5\xAE\x9A\xE8\xA6\x81\xE5\x8D\x95\xE7\x8B\xAC\xE6\x89\xA7\xE8\xA1\x8C\xE6\xAD\xA4\xE4\xBB\xBB\xE5\x8A\xA1\xE5\x90\x97\xEF\xBC\x9F\n\xE4\xBB\xBB\xE5\x8A\xA1: ") + QString::fromStdString(task.taskName)) 
        != QMessageBox::Yes) {
        return;
    }

    if (processButton_) processButton_->setEnabled(false);
    if (progressBar_) {
        progressBar_->setVisible(true);
        progressBar_->setRange(0, 0); // Indeterminate
    }
    if (logTextEdit_) logTextEdit_->clear();
    processor_->clearErrors();
    processor_->clearWarnings();

    std::thread([this, taskId]() {
        #ifdef _WIN32
        CoInitialize(nullptr);
        #endif
        
        // Log start
        QMetaObject::invokeMethod(this, [this, taskId]() {
             if (logTextEdit_) logTextEdit_->append(QString::fromUtf8("\xE5\xBC\x80\xE5\xA7\x8B\xE6\x89\xA7\xE8\xA1\x8C\xE4\xBB\xBB\xE5\x8A\xA1 ID: %1...").arg(taskId));
        });

        auto startTime = std::chrono::high_resolution_clock::now();

        ProcessingResult result = processor_->processTask(taskId);

        auto endTime = std::chrono::high_resolution_clock::now();
        double duration = std::chrono::duration<double, std::milli>(endTime - startTime).count() / 1000.0;
        result.processingTime = duration;

        #ifdef _WIN32
        CoUninitialize();
        #endif

        QMetaObject::invokeMethod(this, [this, result]() {
            if (processButton_) processButton_->setEnabled(true);
            if (progressBar_) {
                progressBar_->setVisible(false);
                progressBar_->setRange(0, 100);
            }

            if (logTextEdit_) {
                QString timestamp = QDateTime::currentDateTime().toString("[yyyy-MM-dd hh:mm:ss] ");
                logTextEdit_->append(timestamp + QString::fromUtf8("\xE5\x8D\x95\xE4\xBB\xBB\xE5\x8A\xA1\xE5\xA4\x84\xE7\x90\x86\xE5\xAE\x8C\xE6\x88\x90\xEF\xBC\x81\xE8\x80\x97\xE6\x97\xB6: %1\xE7\xA7\x92").arg(result.processingTime));
            }

            if (!result.errors.empty()) {
                QString errorList;
                for (const auto& err : result.errors) errorList += QString::fromStdString(err) + "\n";
                QMessageBox::warning(this, QString::fromUtf8("\xE9\x94\x99\xE8\xAF\xAF"), errorList);
            } else {
                 QString msg = QString::fromUtf8("\xE4\xBB\xBB\xE5\x8A\xA1\xE5\xA4\x84\xE7\x90\x86\xE5\xAE\x8C\xE6\x88\x90\xEF\xBC\x81\n\xE5\x8C\xB9\xE9\x85\x8D\xE8\xA1\x8C\xE6\x95\xB0: %1\n\xE5\xA4\x84\xE7\x90\x86\xE8\xA1\x8C\xE6\x95\xB0: %2").arg(result.matchedRows).arg(result.processedRows);
                 QMessageBox::information(this, QString::fromUtf8("\xE5\xAE\x8C\xE6\x88\x90"), msg);
            }
        });
    }).detach();
}

void ExcelProcessorGUI::editTreeItem() {
    QTreeWidgetItem* item = taskTree_->currentItem();
    if (!item) return;

    if (item->parent() == nullptr) {
        QString pattern = item->data(0, Qt::UserRole).toString();
        QString sheet = item->data(0, Qt::UserRole + 1).toString();
        QString oldPattern = pattern;
        QString oldSheet = sheet;
        
        if (showInputGroupDialog(pattern, sheet)) {
            auto tasks = processor_->getTasks();
            bool changed = false;
            for (auto& task : tasks) {
                if (QString::fromStdString(task.inputFilenamePattern) == oldPattern &&
                    QString::fromStdString(task.inputSheetName) == oldSheet) {
                    task.inputFilenamePattern = pattern.toStdString();
                    task.inputSheetName = sheet.toStdString();
                    processor_->updateTask(task);
                    changed = true;
                }
            }
            if (changed) updateTaskTree();
        }
    } else {
        int taskId = item->data(0, Qt::UserRole).toInt();
        showTaskEditDialog(taskId);
    }
}

void ExcelProcessorGUI::deleteTreeItem() {
    QTreeWidgetItem* item = taskTree_->currentItem();
    if (!item) return;

    if (item->parent() == nullptr) {
        auto reply = QMessageBox::question(this, QString::fromUtf8("\xE7\xA1\xAE\xE8\xAE\xA4\xE5\x88\xA0\xE9\x99\xA4"), 
            QString::fromUtf8("\xE7\xA1\xAE\xE5\xAE\x9A\xE8\xA6\x81\xE5\x88\xA0\xE9\x99\xA4\xE8\xAF\xA5\xE8\xBE\x93\xE5\x85\xA5\xE6\xBA\x90\xE5\x8F\x8A\xE5\x85\xB6\xE6\x89\x80\xE6\x9C\x89\xE4\xBB\xBB\xE5\x8A\xA1\xE5\x90\x97\xEF\xBC\x9F"),
            QMessageBox::Yes|QMessageBox::No);
            
        if (reply == QMessageBox::Yes) {
            QString pattern = item->data(0, Qt::UserRole).toString();
            QString sheet = item->data(0, Qt::UserRole + 1).toString();
            
            auto tasks = processor_->getTasks();
            std::vector<int> idsToRemove;
            for (const auto& task : tasks) {
                if (QString::fromStdString(task.inputFilenamePattern) == pattern &&
                    QString::fromStdString(task.inputSheetName) == sheet) {
                    idsToRemove.push_back(task.id);
                }
            }
            
            for (int id : idsToRemove) {
                processor_->removeTask(id);
            }
            updateTaskTree();
        }
    } else {
        int taskId = item->data(0, Qt::UserRole).toInt();
        processor_->removeTask(taskId);
        updateTaskTree();
    }
}

QWidget* ExcelProcessorGUI::createTaskTab() {
    auto widget = new QWidget();
    auto layout = new QVBoxLayout(widget);

    auto buttonLayout = new QHBoxLayout();
    auto addGroupBtn = new QPushButton(QString::fromUtf8("\xE6\xB7\xBB\xE5\x8A\xA0\xE8\xBE\x93\xE5\x85\xA5\xE6\xBA\x90")); // Add Input Source
    auto addTaskBtn = new QPushButton(QString::fromUtf8("\xE6\xB7\xBB\xE5\x8A\xA0\xE8\xBE\x93\xE5\x87\xBA\xE8\xA7\x84\xE5\x88\x99")); // Add Output Rule
    auto editBtn = new QPushButton(QString::fromUtf8("\xE7\xBC\x96\xE8\xBE\x91")); // Edit
    auto deleteBtn = new QPushButton(QString::fromUtf8("\xE5\x88\xA0\xE9\x99\xA4")); // Delete
    
    auto loadConfigBtn = new QPushButton(QString::fromUtf8("\xE5\x8A\xA0\xE8\xBD\xBD\xE9\x85\x8D\xE7\xBD\xAE")); // Load Config
    auto saveConfigBtn = new QPushButton(QString::fromUtf8("\xE4\xBF\x9D\xE5\xAD\x98\xE9\x85\x8D\xE7\xBD\xAE")); // Save Config

    int btnHeight = 35;
    addGroupBtn->setMinimumHeight(btnHeight);
    addTaskBtn->setMinimumHeight(btnHeight);
    editBtn->setMinimumHeight(btnHeight);
    deleteBtn->setMinimumHeight(btnHeight);
    loadConfigBtn->setMinimumHeight(btnHeight);
    saveConfigBtn->setMinimumHeight(btnHeight);

    addTaskBtn->setStyleSheet("font-weight: bold; color: green;");

    buttonLayout->addWidget(addGroupBtn);
    buttonLayout->addWidget(addTaskBtn);
    buttonLayout->addWidget(editBtn);
    buttonLayout->addWidget(deleteBtn);
    buttonLayout->addStretch();
    buttonLayout->addWidget(loadConfigBtn);
    buttonLayout->addWidget(saveConfigBtn);
    layout->addLayout(buttonLayout);

    taskTree_ = new QTreeWidget();
    setupTaskTree();
    layout->addWidget(taskTree_);

    connect(addGroupBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::addInputGroup);
    connect(addTaskBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::addTaskToGroup);
    connect(editBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::editTreeItem);
    connect(deleteBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::deleteTreeItem);
    connect(loadConfigBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::loadConfiguration);
    connect(saveConfigBtn, &QPushButton::clicked, this, &ExcelProcessorGUI::saveConfiguration);
    connect(taskTree_, &QTreeWidget::itemDoubleClicked, this, &ExcelProcessorGUI::editTreeItem);

    return widget;
}

void ExcelProcessorGUI::setupTaskTree() {
    QStringList headers;
    headers << QString::fromUtf8("\xE4\xBB\xBB\xE5\x8A\xA1/\xE9\x85\x8D\xE7\xBD\xAE")
            << QString::fromUtf8("\xE8\xAF\xA6\xE6\x83\x85")
            << QString::fromUtf8("\xE5\x90\xAF\xE7\x94\xA8")
            << QString::fromUtf8("\xE9\xA2\x84\xE8\xA7\x88"); // Add Preview Column
            
    taskTree_->setHeaderLabels(headers);
    taskTree_->setColumnWidth(0, 400);
    taskTree_->setColumnWidth(1, 400);
    taskTree_->setColumnWidth(2, 50);
    taskTree_->setColumnWidth(3, 80); // Width for button

    // Center align headers
    taskTree_->header()->setDefaultAlignment(Qt::AlignCenter);

    // Enable custom context menu
    taskTree_->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(taskTree_, &QTreeWidget::customContextMenuRequested, this, &ExcelProcessorGUI::onTaskTreeContextMenu);
}

void ExcelProcessorGUI::onTaskTreeContextMenu(const QPoint& pos) {
    QTreeWidgetItem* item = taskTree_->itemAt(pos);
    if (!item) return;

    QMenu menu(this);
    
    // Check if it's a root item (Input Group) or child item (Output Task)
    if (item->parent() == nullptr) {
        // Root Item: Input Source
        QAction* addAction = menu.addAction(QString::fromUtf8("\xE6\xB7\xBB\xE5\x8A\xA0\xE8\xBE\x93\xE5\x87\xBA\xE8\xA7\x84\xE5\x88\x99 (+)"));
        connect(addAction, &QAction::triggered, this, &ExcelProcessorGUI::addTaskToGroup);
        
        menu.addSeparator();
        
        QAction* editAction = menu.addAction(QString::fromUtf8("\xE7\xBC\x96\xE8\xBE\x91\xE8\xBE\x93\xE5\x85\xA5\xE6\xBA\x90"));
        connect(editAction, &QAction::triggered, this, &ExcelProcessorGUI::editTreeItem);
        
        QAction* deleteAction = menu.addAction(QString::fromUtf8("\xE5\x88\xA0\xE9\x99\xA4\xE8\xBE\x93\xE5\x85\xA5\xE6\xBA\x90 (-)"));
        connect(deleteAction, &QAction::triggered, this, &ExcelProcessorGUI::deleteTreeItem);
    } else {
        // Child Item: Output Task
        QAction* editAction = menu.addAction(QString::fromUtf8("\xE7\xBC\x96\xE8\xBE\x91\xE4\xBB\xBB\xE5\x8A\xA1"));
        connect(editAction, &QAction::triggered, this, &ExcelProcessorGUI::editTreeItem);
        
        QAction* deleteAction = menu.addAction(QString::fromUtf8("\xE5\x88\xA0\xE9\x99\xA4\xE4\xBB\xBB\xE5\x8A\xA1 (-)"));
        connect(deleteAction, &QAction::triggered, this, &ExcelProcessorGUI::deleteTreeItem);
    }

    menu.exec(taskTree_->viewport()->mapToGlobal(pos));
}

void ExcelProcessorGUI::updateTaskTree() {
    taskTree_->clear();
    auto tasks = processor_->getTasks();
    
    std::vector<std::string> groupOrder;
    std::map<std::string, std::vector<ProcessingTask>> groups;
    
    for (const auto& task : tasks) {
        std::string key = task.inputFilenamePattern + "|" + task.inputSheetName;
        if (groups.find(key) == groups.end()) {
            groupOrder.push_back(key);
        }
        groups[key].push_back(task);
    }
    
    for (const auto& key : groupOrder) {
        const auto& groupTasks = groups[key];
        if (groupTasks.empty()) continue;
        
        const auto& first = groupTasks[0];
        QTreeWidgetItem* root = new QTreeWidgetItem(taskTree_);
        QString inputName = QString::fromUtf8("\xE8\xBE\x93\xE5\x85\xA5: ") + QString::fromStdString(first.inputFilenamePattern);
        if (!first.inputSheetName.empty()) {
            inputName += QString::fromUtf8(" | \xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8: ") + QString::fromStdString(first.inputSheetName);
        } else {
             inputName += QString::fromUtf8(" | \xE6\x89\x80\xE6\x9C\x89\xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8");
        }
        
        root->setText(0, inputName);
        root->setText(1, QString::fromUtf8("\xE5\x8C\x85\xE5\x90\xAB %1 \xE4\xB8\xAA\xE8\xBE\x93\xE5\x87\xBA\xE4\xBB\xBB\xE5\x8A\xA1").arg(groupTasks.size()));
        root->setExpanded(true);
        root->setData(0, Qt::UserRole, QString::fromStdString(first.inputFilenamePattern));
        root->setData(0, Qt::UserRole + 1, QString::fromStdString(first.inputSheetName));
        
        for (const auto& task : groupTasks) {
            QTreeWidgetItem* child = new QTreeWidgetItem(root);
            QString outputInfo = QString::fromUtf8("\xE8\xBE\x93\xE5\x87\xBA: ") + QString::fromStdString(task.outputWorkbookName);
            child->setText(0, outputInfo);
            
            QString rulesInfo = QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99: %1 \xE4\xB8\xAA").arg(task.rules.size());
            if (task.ruleLogic == RuleLogic::AND) rulesInfo += " (AND)";
            else rulesInfo += " (OR)";
            
            child->setText(1, rulesInfo);
            child->setTextAlignment(1, Qt::AlignCenter); // Center align
            
            // Enable Checkbox (Column 2)
            QWidget* checkWidget = new QWidget();
            QHBoxLayout* checkLayout = new QHBoxLayout(checkWidget);
            checkLayout->setContentsMargins(0, 0, 0, 0);
            checkLayout->setAlignment(Qt::AlignCenter);
            QCheckBox* checkBox = new QCheckBox();
            checkBox->setChecked(task.enabled);
            
            // Connect signal to update task
            connect(checkBox, &QCheckBox::stateChanged, [this, taskId = task.id](int state) {
                auto t = processor_->getTask(taskId);
                if (t.id != 0) {
                   t.enabled = (state == Qt::Checked);
                   processor_->updateTask(t);
                   
                   // Auto-save if config is loaded to persist the enabled state immediately
                   if (!currentConfigFile_.isEmpty()) {
                       processor_->saveConfiguration(currentConfigFile_.toStdString());
                   }
                }
            });
            
            checkLayout->addWidget(checkBox);
            taskTree_->setItemWidget(child, 2, checkWidget);

            child->setData(0, Qt::UserRole, task.id);

            // Add Preview Button
            QWidget* btnWidget = new QWidget();
            QHBoxLayout* btnLayout = new QHBoxLayout(btnWidget);
            btnLayout->setContentsMargins(2, 2, 2, 2);
            btnLayout->setAlignment(Qt::AlignCenter);
            
            QPushButton* previewBtn = new QPushButton(QString::fromUtf8("\xE9\xA2\x84\xE8\xBE\x93\xE5\x85\xA5")); // Preview Input
            previewBtn->setToolTip(QString::fromUtf8("\xE9\xA2\x84\xE8\xA7\x88\xE8\xBE\x93\xE5\x85\xA5\xE6\x95\xB0\xE6\x8D\xAE"));
            previewBtn->setFixedSize(50, 20); // Smaller size
            previewBtn->setStyleSheet("background-color: #2196F3; color: white; border-radius: 4px; padding: 0px; font-size: 11px;");
            
            connect(previewBtn, &QPushButton::clicked, [this, task]() {
                this->previewTask(task.id);
            });
            
            btnLayout->addWidget(previewBtn);

            // Add Output Button
            QPushButton* outputBtn = new QPushButton(QString::fromUtf8("\xE8\xBE\x93\xE5\x87\xBA")); // Output
            outputBtn->setToolTip(QString::fromUtf8("\xE6\x89\xA7\xE8\xA1\x8C\xE6\xAD\xA4\xE4\xBB\xBB\xE5\x8A\xA1\xE5\xB9\xB6\xE8\xBE\x93\xE5\x87\xBA"));
            outputBtn->setFixedSize(50, 20);
            outputBtn->setStyleSheet("background-color: #4CAF50; color: white; border-radius: 4px; padding: 0px; font-size: 11px; margin-left: 5px;");
            
            connect(outputBtn, &QPushButton::clicked, [this, task]() {
                this->executeSingleTask(task.id);
            });
            
            btnLayout->addWidget(outputBtn);
            
            taskTree_->setItemWidget(child, 3, btnWidget);
        }
    }
}

int ExcelProcessorGUI::generateTaskId() {
    auto tasks = processor_->getTasks();
    int maxId = 0;
    for (const auto& task : tasks) {
        if (task.id > maxId) maxId = task.id;
    }
    return maxId + 1;
}

void ExcelProcessorGUI::showTaskEditDialog(int taskId, const QString& inputPattern, const QString& inputSheet) {
    ProcessingTask task;
    bool isNew = (taskId == -1);
    
    if (!isNew) {
        task = processor_->getTask(taskId);
    } else {
        task.id = generateTaskId();
        task.outputWorkbookName = "New_Output.xlsx";
        task.taskName = "Task " + std::to_string(task.id);
        if (!inputPattern.isEmpty()) task.inputFilenamePattern = inputPattern.toStdString();
        else task.inputFilenamePattern = "*.xlsx";
        
        if (!inputSheet.isEmpty()) task.inputSheetName = inputSheet.toStdString();
    }

    QDialog dialog(this);
    dialog.setWindowTitle(isNew ? QString::fromUtf8("\xE6\xB7\xBB\xE5\x8A\xA0\xE4\xBB\xBB\xE5\x8A\xA1") : QString::fromUtf8("\xE7\xBC\x96\xE8\xBE\x91\xE4\xBB\xBB\xE5\x8A\xA1"));
    dialog.resize(600, 500);
    
    QVBoxLayout* layout = new QVBoxLayout(&dialog);
    
    QGroupBox* basicGroup = new QGroupBox(QString::fromUtf8("\xE5\x9F\xBA\xE6\x9C\xAC\xE4\xBF\xA1\xE6\x81\xAF"));
    QGridLayout* basicLayout = new QGridLayout(basicGroup);
    basicLayout->addWidget(new QLabel(QString::fromUtf8("\xE4\xBB\xBB\xE5\x8A\xA1\xE5\x90\x8D\xE7\xA7\xB0:")), 0, 0);
    QLineEdit* nameEdit = new QLineEdit(QString::fromStdString(task.taskName));
    basicLayout->addWidget(nameEdit, 0, 1);
    layout->addWidget(basicGroup);
    
    // Input Group Removed - Input settings are managed via Input Group Dialog only
    
    QGroupBox* outputGroup = new QGroupBox(QString::fromUtf8("\xE8\xBE\x93\xE5\x87\xBA\xE8\xAE\xBE\xE7\xBD\xAE"));
    QGridLayout* outputLayout = new QGridLayout(outputGroup);
    outputLayout->addWidget(new QLabel(QString::fromUtf8("\xE8\xBE\x93\xE5\x87\xBA\xE6\xA8\xA1\xE5\xBC\x8F:")), 0, 0);
    QComboBox* modeCombo = new QComboBox();
    modeCombo->addItem(QString::fromUtf8("\xE6\x96\xB0\xE5\xBB\xBA\xE5\xB7\xA5\xE4\xBD\x9C\xE7\xB0\xBF"), static_cast<int>(OutputMode::NEW_WORKBOOK));
    modeCombo->addItem(QString::fromUtf8("\xE6\x96\xB0\xE5\xBB\xBA\xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8 (\xE8\xBF\xBD\xE5\x8A\xA0\xE5\x88\xB0\xE9\xBB\x98\xE8\xAE\xA4\xE6\x96\x87\xE4\xBB\xB6)"), static_cast<int>(OutputMode::NEW_SHEET));
    modeCombo->addItem(QString::fromUtf8("\xE4\xB8\x8D\xE8\xBE\x93\xE5\x87\xBA"), static_cast<int>(OutputMode::NONE));
    modeCombo->setCurrentIndex(modeCombo->findData(static_cast<int>(task.outputMode)));
    outputLayout->addWidget(modeCombo, 0, 1);
    outputLayout->addWidget(new QLabel(QString::fromUtf8("\xE8\xBE\x93\xE5\x87\xBA\xE5\x90\x8D\xE7\xA7\xB0:")), 1, 0);
    
    QHBoxLayout* outputNameLayout = new QHBoxLayout();
    QLineEdit* outputNameEdit = new QLineEdit(QString::fromStdString(task.outputWorkbookName));
    outputNameLayout->addWidget(outputNameEdit);
    
    QPushButton* outputBrowseBtn = new QPushButton("...");
    outputBrowseBtn->setFixedWidth(30);
    // Only enable browse button if Output Mode is NEW_WORKBOOK
    outputBrowseBtn->setEnabled(task.outputMode == OutputMode::NEW_WORKBOOK);
    outputNameLayout->addWidget(outputBrowseBtn);
    
    outputLayout->addLayout(outputNameLayout, 1, 1);
    
    connect(modeCombo, QOverload<int>::of(&QComboBox::currentIndexChanged), [outputBrowseBtn, modeCombo](int index) {
        OutputMode mode = static_cast<OutputMode>(modeCombo->itemData(index).toInt());
        outputBrowseBtn->setEnabled(mode == OutputMode::NEW_WORKBOOK);
    });
    
    connect(outputBrowseBtn, &QPushButton::clicked, [this, outputNameEdit, &dialog]() {
        QString filename = QFileDialog::getSaveFileName(
            &dialog, 
            QString::fromUtf8("\xE9\x80\x89\xE6\x8B\xA9\xE8\xBE\x93\xE5\x87\xBA\xE6\x96\x87\xE4\xBB\xB6"), 
            outputNameEdit->text(), 
            QString::fromUtf8("Excel\xE6\x96\x87\xE4\xBB\xB6 (*.xlsx);;CSV\xE6\x96\x87\xE4\xBB\xB6 (*.csv)")
        );
        if (!filename.isEmpty()) {
            outputNameEdit->setText(filename);
        }
    });
    QCheckBox* overwriteCheck = new QCheckBox(QString::fromUtf8("\xE8\xA6\x86\xE7\x9B\x96\xE5\x90\x8C\xE5\x90\x8D\xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8")); // Overwrite existing sheet
    overwriteCheck->setChecked(task.overwriteSheet);
    outputLayout->addWidget(overwriteCheck, 2, 0, 1, 2);

    QCheckBox* useHeaderCheck = new QCheckBox(QString::fromUtf8("\xE4\xBD\xBF\xE7\x94\xA8\xE8\xBE\x93\xE5\x85\xA5\xE6\xBA\x90\xE7\x9A\x84\xE8\xA1\xA8\xE5\xA4\xB4")); // Use input header
    useHeaderCheck->setChecked(task.useHeader);
    outputLayout->addWidget(useHeaderCheck, 3, 0, 1, 2);
    
    layout->addWidget(outputGroup);
    
    QGroupBox* rulesGroup = new QGroupBox(QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99\xE9\x80\x89\xE6\x8B\xA9"));
    QVBoxLayout* rulesLayout = new QVBoxLayout(rulesGroup);
    QHBoxLayout* logicLayout = new QHBoxLayout();
    logicLayout->addWidget(new QLabel(QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99\xE7\xBB\x84\xE5\x90\x88\xE9\x80\xBB\xE8\xBE\x91:")));
    QRadioButton* orRadio = new QRadioButton(QString::fromUtf8("\xE4\xBB\xBB\xE6\x84\x8F\xE6\xBB\xA1\xE8\xB6\xB3 (OR)"));
    QRadioButton* andRadio = new QRadioButton(QString::fromUtf8("\xE5\x85\xA8\xE9\x83\xA8\xE6\xBB\xA1\xE8\xB6\xB3 (AND)"));
    QButtonGroup* logicGroup = new QButtonGroup(&dialog);
    logicGroup->addButton(orRadio, static_cast<int>(RuleLogic::OR));
    logicGroup->addButton(andRadio, static_cast<int>(RuleLogic::AND));
    
    if (task.ruleLogic == RuleLogic::AND) andRadio->setChecked(true);
    else orRadio->setChecked(true);
    
    logicLayout->addWidget(orRadio);
    logicLayout->addWidget(andRadio);
    logicLayout->addStretch();
    rulesLayout->addLayout(logicLayout);

    QTableWidget* rulesTable = new QTableWidget();
    rulesTable->setColumnCount(3);
    rulesTable->setHorizontalHeaderLabels({QString::fromUtf8("\xE9\x80\x89\xE6\x8B\xA9"), QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99\xE5\x90\x8D\xE7\xA7\xB0"), QString::fromUtf8("\xE6\x8E\x92\xE9\x99\xA4\xE8\xA7\x84\xE5\x88\x99")});
    rulesTable->horizontalHeader()->setSectionResizeMode(0, QHeaderView::ResizeToContents);
    rulesTable->horizontalHeader()->setSectionResizeMode(1, QHeaderView::Stretch);
    rulesTable->horizontalHeader()->setSectionResizeMode(2, QHeaderView::ResizeToContents);
    
    auto allRules = processor_->getRules();
    
    // Helper to find entry
    auto findEntry = [&](int rid) -> const TaskRuleEntry* {
        for(const auto& e : task.rules) if(e.ruleId == rid) return &e;
        return nullptr;
    };

    for (int i = 0; i < allRules.size(); ++i) {
        const auto& rule = allRules[i];
        rulesTable->insertRow(i);
        
        // Checkbox Widget (Centered)
        QWidget* checkWidget = new QWidget();
        QHBoxLayout* checkLayout = new QHBoxLayout(checkWidget);
        checkLayout->setContentsMargins(0,0,0,0);
        checkLayout->setAlignment(Qt::AlignCenter);
        QCheckBox* checkBox = new QCheckBox();
        const TaskRuleEntry* entry = findEntry(rule.id);
        checkBox->setChecked(entry ? true : false);
        checkBox->setProperty("ruleId", rule.id); // Store Rule ID
        checkLayout->addWidget(checkBox);
        rulesTable->setCellWidget(i, 0, checkWidget);
        
        // Name
        rulesTable->setItem(i, 1, new QTableWidgetItem(QString::fromStdString(rule.name)));
        
        // Exclusions Button
        QPushButton* configBtn = new QPushButton(QString::fromUtf8("\xE8\xAE\xBE\xE7\xBD\xAE\xE6\x8E\x92\xE9\x99\xA4"));
        // Store current exclusions in property
        QList<QVariant> exList;
        if(entry) {
             for(int eid : entry->excludeRuleIds) exList.append(eid);
        }
        configBtn->setProperty("exIds", exList);
        
        // Initial text update
        if(!exList.isEmpty()) configBtn->setText(QString::fromUtf8("\xE6\x8E\x92\xE9\x99\xA4(%1)").arg(exList.size()));

        // Connect button
        // Capture allRules by reference is safe because dialog.exec() blocks
        connect(configBtn, &QPushButton::clicked, [this, configBtn, ruleId = rule.id, ruleName = rule.name, &allRules]() {
            QDialog exDialog(this);
            exDialog.setWindowTitle(QString::fromUtf8("\xE8\xAE\xBE\xE7\xBD\xAE\xE6\x8E\x92\xE9\x99\xA4\xE8\xA7\x84\xE5\x88\x99 - ") + QString::fromStdString(ruleName));
            exDialog.resize(400, 300);
            QVBoxLayout* exLayout = new QVBoxLayout(&exDialog);
            QListWidget* exListWidget = new QListWidget();
            
            QList<QVariant> currentEx = configBtn->property("exIds").toList();
            std::set<int> currentExSet;
            for(auto v : currentEx) currentExSet.insert(v.toInt());
            
            for(const auto& r : allRules) {
                if(r.id == ruleId) continue; // Can't exclude self
                QListWidgetItem* item = new QListWidgetItem(QString::fromStdString(r.name));
                item->setData(Qt::UserRole, r.id);
                item->setFlags(item->flags() | Qt::ItemIsUserCheckable);
                item->setCheckState(currentExSet.count(r.id) ? Qt::Checked : Qt::Unchecked);
                exListWidget->addItem(item);
            }
            exLayout->addWidget(exListWidget);
            
            QDialogButtonBox* exBtns = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel);
            exLayout->addWidget(exBtns);
            connect(exBtns, &QDialogButtonBox::accepted, &exDialog, &QDialog::accept);
            connect(exBtns, &QDialogButtonBox::rejected, &exDialog, &QDialog::reject);
            
            if(exDialog.exec() == QDialog::Accepted) {
                QList<QVariant> newEx;
                for(int j=0; j<exListWidget->count(); ++j) {
                     if(exListWidget->item(j)->checkState() == Qt::Checked) {
                         newEx.append(exListWidget->item(j)->data(Qt::UserRole));
                     }
                }
                configBtn->setProperty("exIds", newEx);
                if (!newEx.isEmpty())
                    configBtn->setText(QString::fromUtf8("\xE6\x8E\x92\xE9\x99\xA4(%1)").arg(newEx.size()));
                else
                    configBtn->setText(QString::fromUtf8("\xE8\xAE\xBE\xE7\xBD\xAE\xE6\x8E\x92\xE9\x99\xA4"));
            }
        });
        
        rulesTable->setCellWidget(i, 2, configBtn);
    }
    rulesLayout->addWidget(rulesTable);
    layout->addWidget(rulesGroup);
    
    QDialogButtonBox* buttons = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel);
    layout->addWidget(buttons);
    connect(buttons, &QDialogButtonBox::accepted, &dialog, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, &dialog, &QDialog::reject);
    
    if (dialog.exec() == QDialog::Accepted) {
        task.taskName = nameEdit->text().toStdString();
        task.outputMode = static_cast<OutputMode>(modeCombo->currentData().toInt());
        task.outputWorkbookName = outputNameEdit->text().toStdString();
        task.overwriteSheet = overwriteCheck->isChecked();
        task.useHeader = useHeaderCheck->isChecked();
        task.ruleLogic = static_cast<RuleLogic>(logicGroup->checkedId());
        
        task.rules.clear();
        for (int i = 0; i < rulesTable->rowCount(); ++i) {
            QWidget* w = rulesTable->cellWidget(i, 0);
            QCheckBox* cb = w ? w->findChild<QCheckBox*>() : nullptr;
            if (cb && cb->isChecked()) {
                int rid = cb->property("ruleId").toInt();
                TaskRuleEntry entry(rid);
                
                QPushButton* btn = qobject_cast<QPushButton*>(rulesTable->cellWidget(i, 2));
                if(btn) {
                    QList<QVariant> exList = btn->property("exIds").toList();
                    for(auto v : exList) entry.excludeRuleIds.push_back(v.toInt());
                }
                task.rules.push_back(entry);
            }
        }
        
        if (isNew) {
            processor_->addTask(task);
        } else {
            processor_->updateTask(task);
        }
        updateTaskTree();
    }
}

bool ExcelProcessorGUI::showInputGroupDialog(QString& pattern, QString& sheet) {
    QDialog dialog(this);
    dialog.setWindowTitle(QString::fromUtf8("\xE8\xBE\x93\xE5\x85\xA5\xE6\xBA\x90\xE9\x85\x8D\xE7\xBD\xAE"));
    QVBoxLayout* layout = new QVBoxLayout(&dialog);
    
    layout->addWidget(new QLabel(QString::fromUtf8("\xE6\x96\x87\xE4\xBB\xB6\xE5\x90\x8D\xE5\x8C\xB9\xE9\x85\x8D (\xE9\x80\x9A\xE9\x85\x8D\xE7\xAC\xA6 *):")));
    
    QHBoxLayout* fileLayout = new QHBoxLayout();
    QLineEdit* patternEdit = new QLineEdit(pattern);
    if (patternEdit->text().isEmpty()) patternEdit->setText("*.xlsx");
    fileLayout->addWidget(patternEdit);
    
    QPushButton* browseBtn = new QPushButton("...");
    browseBtn->setFixedWidth(30);
    fileLayout->addWidget(browseBtn);
    layout->addLayout(fileLayout);

    connect(browseBtn, &QPushButton::clicked, [&dialog, patternEdit]() {
        QString filename = QFileDialog::getOpenFileName(
            &dialog, QString::fromUtf8("\xE9\x80\x89\xE6\x8B\xA9\xE6\x96\x87\xE4\xBB\xB6"), "", QString::fromUtf8("Excel\xE6\x96\x87\xE4\xBB\xB6 (*.xlsx *.xls);;\xE6\x89\x80\xE6\x9C\x89\xE6\x96\x87\xE4\xBB\xB6 (*.*)")
        );
        if (!filename.isEmpty()) {
            patternEdit->setText(filename);
        }
    });
    
    layout->addWidget(new QLabel(QString::fromUtf8("\xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8\xE5\x90\x8D (\xE7\x95\x99\xE7\xA9\xBA\xE4\xBB\xA3\xE8\xA1\xA8\xE6\x89\x80\xE6\x9C\x89):")));
    QLineEdit* sheetEdit = new QLineEdit(sheet);
    layout->addWidget(sheetEdit);
    
    QDialogButtonBox* buttons = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel);
    layout->addWidget(buttons);
    connect(buttons, &QDialogButtonBox::accepted, &dialog, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, &dialog, &QDialog::reject);
    
    if (dialog.exec() == QDialog::Accepted) {
        pattern = patternEdit->text();
        sheet = sheetEdit->text();
        return true;
    }
    return false;
}
