#pragma once

#include <QMainWindow>
#include <memory>
#include <vector>
#include <string>
#include <atomic>
#include "ExcelProcessorCore.h"

class QGroupBox;
class QLabel;
class QLineEdit;
class QProgressBar;
class QPushButton;
class QStatusBar;
class QTableWidget;
class QTreeWidget;
class QTreeWidgetItem;
class QTabWidget;
class QComboBox;
class QTextEdit;

class ExcelProcessorGUI : public QMainWindow {
    Q_OBJECT

public:
    explicit ExcelProcessorGUI(QWidget *parent = nullptr);
    ~ExcelProcessorGUI();

private slots:
    void loadConfiguration();
    void saveConfiguration();
    // void loadInputFile(); // Removed
    // void selectOutputFile(); // Removed
    void addRule();
    void deleteRule();
    void editRule();
    void previewData();
    void previewProcessedData();
    void processFile();
    // void createRuleCombination(); // REMOVED
    void validateRules();
    void tabChanged(int index);
    // void updateProgress();
    
    void addInputGroup();
    void addTaskToGroup();
    void editTreeItem();
    void deleteTreeItem();
    void previewTask(int taskId);
    void executeSingleTask(int taskId);
    void onTaskTreeContextMenu(const QPoint& pos);

private:
    void setupUI();
    // QGroupBox* createFileSelectionGroup(); // Removed
    QWidget* createRuleTab();
    QWidget* createPreviewTab();
    // QWidget* createCombinationTab(); // REMOVED
    QWidget* createTaskTab();
    QWidget* createAboutTab();
    QGroupBox* createButtonGroup();
    
    void setupConnections();
    void loadSettings();
    void saveSettings();
    void setupRuleTable();
    // void setupCombinationTable(); // REMOVED
    void setupTaskTree();
    
    void updateRuleTable();
    void updatePreviewTable(const std::vector<DataRow>& data, const QStringList& customHeaders = QStringList());
    // void updateCombinationTable(); // REMOVED
    void updateTaskTree();
    
    std::vector<int> getSelectedRuleIds();
    int generateRuleId();
    // int generateComboId(); // REMOVED
    int generateTaskId();
    void showRuleEditDialog(int ruleId);
    void showTaskEditDialog(int taskId, const QString& inputPattern = "", const QString& inputSheet = "");
    bool showInputGroupDialog(QString& pattern, QString& sheet);
    std::string getRuleTypeString(RuleType type);
    QString generateRuleDescription(const Rule& rule);

    std::unique_ptr<ExcelProcessorCore> processor_;
    QStatusBar* statusBar_;
    QLabel* statusLabel_;
    QProgressBar* progressBar_;
    QTableWidget* ruleTable_;
    QTableWidget* previewTable_;
    // QTableWidget* combinationTable_; // REMOVED
    QTreeWidget* taskTree_; // Changed from QTableWidget
    // QLineEdit* inputFileEdit_; // Removed
    // QComboBox* sheetComboBox_; // Removed
    // QLineEdit* outputFileEdit_; // Removed
    QPushButton* processButton_;
    QTextEdit* logTextEdit_;
    std::atomic<bool> isLoading_{false};
    QString currentConfigFile_;
    QLabel* configNameLabel_;
    QLabel* previewInputLabel_;
    QLabel* previewOutputLabel_;
    QTabWidget* tabWidget_;
    int currentPreviewTaskId_ = -1; // Track current preview task for refresh
};
