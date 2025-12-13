#pragma once
#include <QDialog>
#include <vector>
#include "ExcelProcessorCore.h"

class QLineEdit;
class QComboBox;
class QTableWidget;
class QPushButton;
class QSpinBox;
class QListWidget;

class RuleEditDialog : public QDialog {
    Q_OBJECT

public:
    explicit RuleEditDialog(const Rule& rule, const std::vector<Rule>& allRules, QWidget* parent = nullptr);
    Rule getRule() const;

private slots:
    void addCondition();
    void removeCondition();
    void save();

private:
    void setupUI();
    void loadRule();
    void updateConditionRow(int row, const RuleCondition& condition);
    RuleCondition getConditionFromRow(int row) const;

    Rule rule_;
    std::vector<Rule> allRules_;
    
    QLineEdit* nameEdit_;
    QComboBox* typeCombo_;
    QLineEdit* targetSheetEdit_;
    QComboBox* enabledCombo_;
    QSpinBox* prioritySpin_;
    


    QComboBox* logicCombo_;
    QTableWidget* conditionsTable_;
    QPushButton* addConditionBtn_;
    QPushButton* removeConditionBtn_;
    QPushButton* okBtn_;
    QPushButton* cancelBtn_;
    
    QListWidget* excludedRulesList_;
};
