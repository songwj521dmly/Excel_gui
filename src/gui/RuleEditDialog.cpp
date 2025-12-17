#include "RuleEditDialog.h"
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QComboBox>
#include <QTableWidget>
#include <QPushButton>
#include <QHeaderView>
#include <QMessageBox>
#include <QChar>
#include <QSpinBox>
#include <QGroupBox>
#include <QListWidget>
#include <QListWidgetItem>

RuleEditDialog::RuleEditDialog(const Rule& rule, const std::vector<Rule>& allRules, QWidget* parent)
    : QDialog(parent), rule_(rule), allRules_(allRules) {
    setupUI();
    loadRule();
}

void RuleEditDialog::setupUI() {
    setWindowTitle("Edit Rule");
    resize(800, 600);

    auto mainLayout = new QVBoxLayout(this);

    // Basic Info Group
    auto basicGroup = new QGroupBox(QString::fromUtf8("\xE5\x9F\xBA\xE6\x9C\xAC" "\xE4\xBF\xA1\xE6\x81\xAF"));
    auto basicLayout = new QGridLayout(basicGroup);
    
    // Name
    basicLayout->addWidget(new QLabel(QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99" "\xE5\x90\x8D\xE7\xA7\xB0:")), 0, 0);
    nameEdit_ = new QLineEdit();
    basicLayout->addWidget(nameEdit_, 0, 1);

    // Type
    basicLayout->addWidget(new QLabel(QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99" "\xE7\xB1\xBB\xE5\x9E\x8B:")), 0, 2);
    typeCombo_ = new QComboBox();
    typeCombo_->addItem(QString::fromUtf8("\xE8\xBF\x87\xE6\xBB\xA4"), static_cast<int>(RuleType::FILTER));
    typeCombo_->addItem(QString::fromUtf8("\xE5\x88\xA0\xE9\x99\xA4"), static_cast<int>(RuleType::DELETE_ROW));
    typeCombo_->addItem(QString::fromUtf8("\xE6\x8B\x86\xE5\x88\x86"), static_cast<int>(RuleType::SPLIT));
    typeCombo_->addItem(QString::fromUtf8("\xE8\xBD\xAC\xE6\x8D\xA2"), static_cast<int>(RuleType::TRANSFORM));
    basicLayout->addWidget(typeCombo_, 0, 3);

    // Logic
    basicLayout->addWidget(new QLabel(QString::fromUtf8("\xE9\x80\xBB\xE8\xBE\x91" "\xE5\x85\xB3\xE7\xB3\xBB:")), 1, 0);
    logicCombo_ = new QComboBox();
    logicCombo_->addItem(QString::fromUtf8("\xE6\xBB\xA1\xE8\xB6\xB3\xE6\x89\x80\xE6\x9C\x89" "\xE6\x9D\xA1\xE4\xBB\xB6 (AND)"), static_cast<int>(RuleLogic::AND));
    logicCombo_->addItem(QString::fromUtf8("\xE6\xBB\xA1\xE8\xB6\xB3\xE4\xBB\xBB\xE4\xB8\x80" "\xE6\x9D\xA1\xE4\xBB\xB6 (OR)"), static_cast<int>(RuleLogic::OR));
    basicLayout->addWidget(logicCombo_, 1, 1);

    // Enabled
    basicLayout->addWidget(new QLabel(QString::fromUtf8("\xE6\x98\xAF\xE5\x90\xA6" "\xE5\x90\xAF\xE7\x94\xA8:")), 1, 2);
    enabledCombo_ = new QComboBox();
    enabledCombo_->addItem(QString::fromUtf8("\xE6\x98\xAF"), 1);
    enabledCombo_->addItem(QString::fromUtf8("\xE5\x90\xA6"), 0);
    basicLayout->addWidget(enabledCombo_, 1, 3);

    // Target Sheet (Visible only for Split)
    auto targetLabel = new QLabel(QString::fromUtf8("\xE7\x9B\xAE\xE6\xA0\x87" "\xE5\xB7\xA5\xE4\xBD\x9C\xE8\xA1\xA8:"));
    basicLayout->addWidget(targetLabel, 2, 0);
    targetSheetEdit_ = new QLineEdit();
    targetSheetEdit_->setPlaceholderText(QString::fromUtf8("\xE4\xBB\x85\xE7\x94\xA8\xE4\xBA\x8E" "\xE6\x8B\x86\xE5\x88\x86\xE8\xA7\x84\xE5\x88\x99"));
    basicLayout->addWidget(targetSheetEdit_, 2, 1);

    // Priority
    basicLayout->addWidget(new QLabel(QString::fromUtf8("\xE4\xBC\x98\xE5\x85\x88\xE7\xBA\xA7:")), 2, 2);
    prioritySpin_ = new QSpinBox();
    prioritySpin_->setRange(0, 100);
    basicLayout->addWidget(prioritySpin_, 2, 3);



    mainLayout->addWidget(basicGroup);
    


    /* Exclusion Logic REMOVED (Moved to Task Level)
    // Exclusion Group
    auto excludeGroup = new QGroupBox(QString::fromUtf8("\xE6\x8E\x92\xE9\x99\xA4\xE8\xA7\x84\xE5\x88\x99 (Exclude Rules)"));
    auto excludeLayout = new QVBoxLayout(excludeGroup);
    
    excludedRulesList_ = new QListWidget();
    excludeLayout->addWidget(excludedRulesList_);
    
    mainLayout->addWidget(excludeGroup);
    */

    // Conditions Table
    auto condGroup = new QGroupBox(QString::fromUtf8("\xE8\xA7\x84\xE5\x88\x99" "\xE6\x9D\xA1\xE4\xBB\xB6"));
    auto condLayout = new QVBoxLayout(condGroup);
    
    conditionsTable_ = new QTableWidget();
    conditionsTable_->setColumnCount(5);
    conditionsTable_->setHorizontalHeaderLabels({
        QString::fromUtf8("\xE5\x88\x97"),
        QString::fromUtf8("\xE6\x93\x8D\xE4\xBD\x9C\xE7\xAC\xA6"),
        QString::fromUtf8("\xE5\x80\xBC"),
        QString::fromUtf8("\xE5\x88\x86\xE9\x9A\x94\xE7\xAC\xA6"),
        QString::fromUtf8("\xE5\x88\x86\xE9\x9A\x94\xE9\x80\xBB\xE8\xBE\x91")
    });
    conditionsTable_->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    conditionsTable_->setSelectionBehavior(QAbstractItemView::SelectRows);
    condLayout->addWidget(conditionsTable_);

    // Condition Buttons
    auto condBtnLayout = new QHBoxLayout();
    addConditionBtn_ = new QPushButton(QString::fromUtf8("\xE6\xB7\xBB\xE5\x8A\xA0" "\xE6\x9D\xA1\xE4\xBB\xB6"));
    removeConditionBtn_ = new QPushButton(QString::fromUtf8("\xE5\x88\xA0\xE9\x99\xA4" "\xE6\x9D\xA1\xE4\xBB\xB6"));
    condBtnLayout->addWidget(addConditionBtn_);
    condBtnLayout->addWidget(removeConditionBtn_);
    condBtnLayout->addStretch();
    condLayout->addLayout(condBtnLayout);

    mainLayout->addWidget(condGroup);
    
    // Bottom Buttons
    auto btnLayout = new QHBoxLayout();
    okBtn_ = new QPushButton(QString::fromUtf8("\xE7\xA1\xAE\xE5\xAE\x9A"));
    cancelBtn_ = new QPushButton(QString::fromUtf8("\xE5\x8F\x96\xE6\xB6\x88"));
    btnLayout->addStretch();
    btnLayout->addWidget(okBtn_);
    btnLayout->addWidget(cancelBtn_);
    
    mainLayout->addLayout(btnLayout);

    // Connections
    connect(addConditionBtn_, &QPushButton::clicked, this, &RuleEditDialog::addCondition);
    connect(removeConditionBtn_, &QPushButton::clicked, this, &RuleEditDialog::removeCondition);
    connect(okBtn_, &QPushButton::clicked, this, &RuleEditDialog::save);
    connect(cancelBtn_, &QPushButton::clicked, this, &QDialog::reject);
    
    // Dynamic UI for Target Sheet
    connect(typeCombo_, QOverload<int>::of(&QComboBox::currentIndexChanged), [=](int index) {
        bool isSplit = (typeCombo_->itemData(index).toInt() == static_cast<int>(RuleType::SPLIT));
        targetSheetEdit_->setEnabled(isSplit);
        if (!isSplit) targetSheetEdit_->clear();
    });
}

void RuleEditDialog::loadRule() {
    nameEdit_->setText(QString::fromStdString(rule_.name));
    
    int typeIndex = typeCombo_->findData(static_cast<int>(rule_.type));
    if (typeIndex >= 0) typeCombo_->setCurrentIndex(typeIndex);

    int logicIndex = logicCombo_->findData(static_cast<int>(rule_.logic));
    if (logicIndex >= 0) logicCombo_->setCurrentIndex(logicIndex);
    
    enabledCombo_->setCurrentIndex(rule_.enabled ? 0 : 1);
    
    targetSheetEdit_->setText(QString::fromStdString(rule_.targetSheet));
    targetSheetEdit_->setEnabled(rule_.type == RuleType::SPLIT);
    
    prioritySpin_->setValue(rule_.priority);



    // Populate exclusion list (REMOVED)
    /*
    excludedRulesList_->clear();
    for (const auto& r : allRules_) {
        // Don't show self
        if (r.id == rule_.id && r.id != 0) continue;
        
        QListWidgetItem* item = new QListWidgetItem(QString::fromStdString(r.name) + " (ID: " + QString::number(r.id) + ")");
        item->setFlags(item->flags() | Qt::ItemIsUserCheckable);
        item->setData(Qt::UserRole, r.id);
        
        bool isExcluded = false;
        for (int exId : rule_.excludedRuleIds) {
            if (exId == r.id) {
                isExcluded = true;
                break;
            }
        }
        item->setCheckState(isExcluded ? Qt::Checked : Qt::Unchecked);
        
        excludedRulesList_->addItem(item);
    }
    */

    conditionsTable_->setRowCount(0);
    for (const auto& condition : rule_.conditions) {
        addCondition();
        updateConditionRow(conditionsTable_->rowCount() - 1, condition);
    }
}

void RuleEditDialog::addCondition() {
    int row = conditionsTable_->rowCount();
    conditionsTable_->insertRow(row);

    // Column Edit (Manual Input as requested)
    auto colEdit = new QLineEdit();
    colEdit->setPlaceholderText("A, B, or 1, 2");
    conditionsTable_->setCellWidget(row, 0, colEdit);

    // Operator Combo
    auto opCombo = new QComboBox();
    opCombo->addItem(QString::fromUtf8("\xE7\xAD\x89\xE4\xBA\x8E (=)"), static_cast<int>(Operator::EQUAL));
    opCombo->addItem(QString::fromUtf8("\xE4\xB8\x8D\xE7\xAD\x89\xE4\xBA\x8E (!=)"), static_cast<int>(Operator::NOT_EQUAL));
    opCombo->addItem(QString::fromUtf8("\xE5\xA4\xA7\xE4\xBA\x8E (>)"), static_cast<int>(Operator::GREATER));
    opCombo->addItem(QString::fromUtf8("\xE5\xB0\x8F\xE4\xBA\x8E (<)"), static_cast<int>(Operator::LESS));
    opCombo->addItem(QString::fromUtf8("\xE5\xA4\xA7\xE4\xBA\x8E\xE7\xAD\x89\xE4\xBA\x8E (>=)"), static_cast<int>(Operator::GREATER_EQUAL));
    opCombo->addItem(QString::fromUtf8("\xE5\xB0\x8F\xE4\xBA\x8E\xE7\xAD\x89\xE4\xBA\x8E (<=)"), static_cast<int>(Operator::LESS_EQUAL));
    opCombo->addItem(QString::fromUtf8("\xE5\x8C\x85\xE5\x90\xAB"), static_cast<int>(Operator::CONTAINS));
    opCombo->addItem(QString::fromUtf8("\xE4\xB8\x8D\xE5\x8C\x85\xE5\x90\xAB"), static_cast<int>(Operator::NOT_CONTAINS));
    opCombo->addItem(QString::fromUtf8("\xE4\xB8\xBA\xE7\xA9\xBA"), static_cast<int>(Operator::EMPTY));
    opCombo->addItem(QString::fromUtf8("\xE4\xB8\x8D\xE4\xB8\xBA\xE7\xA9\xBA"), static_cast<int>(Operator::NOT_EMPTY));
    opCombo->addItem(QString::fromUtf8("\xE4\xBB\xA5...\xE5\xBC\x80\xE5\xA4\xB4"), static_cast<int>(Operator::STARTS_WITH));
    opCombo->addItem(QString::fromUtf8("\xE4\xBB\xA5...\xE7\xBB\x93\xE5\xB0\xBE"), static_cast<int>(Operator::ENDS_WITH));
    conditionsTable_->setCellWidget(row, 1, opCombo);

    // Value Edit
    auto valEdit = new QLineEdit();
    conditionsTable_->setCellWidget(row, 2, valEdit);

    // Split Symbol Edit
    auto splitSymEdit = new QLineEdit();
    splitSymEdit->setPlaceholderText("*");
    conditionsTable_->setCellWidget(row, 3, splitSymEdit);

    // Split Logic Combo
    auto splitLogicCombo = new QComboBox();
    splitLogicCombo->addItem(QString::fromUtf8("\xE6\x97\xA0 (None)"), static_cast<int>(SplitTarget::NONE));
    splitLogicCombo->addItem(QString::fromUtf8("\xE7\xAC\xA6\xE5\x8F\xB7\xE5\x89\x8D (Before)"), static_cast<int>(SplitTarget::BEFORE));
    splitLogicCombo->addItem(QString::fromUtf8("\xE7\xAC\xA6\xE5\x8F\xB7\xE5\x90\x8E (After)"), static_cast<int>(SplitTarget::AFTER));
    splitLogicCombo->addItem(QString::fromUtf8("\xE7\xAC\xA6\xE5\x8F\xB7\xE5\x89\x8D\xE5\x90\x8E (Both)"), static_cast<int>(SplitTarget::BOTH));
    conditionsTable_->setCellWidget(row, 4, splitLogicCombo);
}

void RuleEditDialog::updateConditionRow(int row, const RuleCondition& condition) {
    auto colEdit = qobject_cast<QLineEdit*>(conditionsTable_->cellWidget(row, 0));
    if (colEdit) {
        // Convert int to string (e.g. 1 -> A)
        int col = condition.column;
        if (col <= 0) col = 1;
        QString s;
        int tempCol = col;
        while (tempCol > 0) {
            int rem = (tempCol - 1) % 26;
            s.prepend(QChar('A' + rem));
            tempCol = (tempCol - 1) / 26;
        }
        colEdit->setText(s);
    }

    auto opCombo = qobject_cast<QComboBox*>(conditionsTable_->cellWidget(row, 1));
    if (opCombo) {
        int idx = opCombo->findData(static_cast<int>(condition.oper));
        if (idx >= 0) opCombo->setCurrentIndex(idx);
    }

    auto valEdit = qobject_cast<QLineEdit*>(conditionsTable_->cellWidget(row, 2));
    if (valEdit) {
        if (auto strVal = std::get_if<std::string>(&condition.value)) {
            valEdit->setText(QString::fromStdString(*strVal));
        } else if (auto intVal = std::get_if<int>(&condition.value)) {
            valEdit->setText(QString::number(*intVal));
        } else if (auto doubleVal = std::get_if<double>(&condition.value)) {
            valEdit->setText(QString::number(*doubleVal));
        }
    }

    auto splitSymEdit = qobject_cast<QLineEdit*>(conditionsTable_->cellWidget(row, 3));
    if (splitSymEdit) {
        splitSymEdit->setText(QString::fromStdString(condition.splitSymbol));
    }

    auto splitLogicCombo = qobject_cast<QComboBox*>(conditionsTable_->cellWidget(row, 4));
    if (splitLogicCombo) {
        int idx = splitLogicCombo->findData(static_cast<int>(condition.splitTarget));
        if (idx >= 0) splitLogicCombo->setCurrentIndex(idx);
    }
}

void RuleEditDialog::removeCondition() {
    int row = conditionsTable_->currentRow();
    if (row >= 0) {
        conditionsTable_->removeRow(row);
    }
}

void RuleEditDialog::save() {
    rule_.name = nameEdit_->text().toStdString();
    rule_.type = static_cast<RuleType>(typeCombo_->currentData().toInt());
    rule_.logic = static_cast<RuleLogic>(logicCombo_->currentData().toInt());
    rule_.enabled = (enabledCombo_->currentData().toInt() == 1);
    rule_.targetSheet = targetSheetEdit_->text().toStdString();
    rule_.priority = prioritySpin_->value();
    


    // Save exclusions (REMOVED)
    /*
    rule_.excludedRuleIds.clear();
    for (int i = 0; i < excludedRulesList_->count(); ++i) {
        QListWidgetItem* item = excludedRulesList_->item(i);
        if (item->checkState() == Qt::Checked) {
            rule_.excludedRuleIds.push_back(item->data(Qt::UserRole).toInt());
        }
    }
    */

    rule_.conditions.clear();
    for (int i = 0; i < conditionsTable_->rowCount(); ++i) {
        rule_.conditions.push_back(getConditionFromRow(i));
    }
    
    accept();
}

RuleCondition RuleEditDialog::getConditionFromRow(int row) const {
    RuleCondition cond;
    
    auto colEdit = qobject_cast<QLineEdit*>(conditionsTable_->cellWidget(row, 0));
    if (colEdit) {
        QString s = colEdit->text().toUpper().trimmed();
        int col = 0;
        bool isNum = false;
        int num = s.toInt(&isNum);
        if (isNum) {
            col = num;
        } else {
            for (QChar c : s) {
                if (c >= 'A' && c <= 'Z') {
                    col = col * 26 + (c.unicode() - 'A' + 1);
                }
            }
        }
        cond.column = col > 0 ? col : 1;
    } else {
        cond.column = 1;
    }
    
    auto opCombo = qobject_cast<QComboBox*>(conditionsTable_->cellWidget(row, 1));
    cond.oper = opCombo ? static_cast<Operator>(opCombo->currentData().toInt()) : Operator::EQUAL;
    
    auto valEdit = qobject_cast<QLineEdit*>(conditionsTable_->cellWidget(row, 2));
    QString text = valEdit ? valEdit->text() : "";
    
    cond.value = text.toStdString();

    auto splitSymEdit = qobject_cast<QLineEdit*>(conditionsTable_->cellWidget(row, 3));
    cond.splitSymbol = splitSymEdit ? splitSymEdit->text().toStdString() : "";

    auto splitLogicCombo = qobject_cast<QComboBox*>(conditionsTable_->cellWidget(row, 4));
    cond.splitTarget = splitLogicCombo ? static_cast<SplitTarget>(splitLogicCombo->currentData().toInt()) : SplitTarget::NONE;
    
    return cond;
}

Rule RuleEditDialog::getRule() const {
    return rule_;
}
