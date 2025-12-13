#include "EditTaskDialog.h"
#include "ExclusionDialog.h"
#include <FL/Fl.H>
#include <FL/fl_ask.H>
#include <string>

EditTaskDialog::EditTaskDialog(int w, int h, const char* title)
    : Fl_Double_Window(w, h, title), result_(0) {
    setupUI();
}

EditTaskDialog::~EditTaskDialog() {
}

void EditTaskDialog::setupUI() {
    int y = 10;
    int inputH = 30;
    int labelW = 100;
    int inputW = w() - labelW - 20;
    int spacing = 10;

    // Task Name
    nameInput_ = new Fl_Input(labelW + 10, y, inputW, inputH, "\xE4\xBB\xBB\xE5\x8A\xA1\xE5\x90\x8D\xE7\xA7\xB0:"); // Task Name
    y += inputH + spacing;

    // Input Pattern
    inputPatternInput_ = new Fl_Input(labelW + 10, y, inputW, inputH, "\xE8\xBE\x93\xE5\x85\xA5\xE6\xA8\xA1\xE5\xBC\x8F:"); // Input Pattern
    y += inputH + spacing;

    // Input Sheet
    inputSheetInput_ = new Fl_Input(labelW + 10, y, inputW, inputH, "\xE8\xBE\x93\xE5\x85\xA5\x45\x78\x63\x65\x6C\xE8\xA1\xA8:"); // Input Sheet
    y += inputH + spacing;

    // Output Name
    outputNameInput_ = new Fl_Input(labelW + 10, y, inputW, inputH, "\xE8\xBE\x93\xE5\x87\xBA\xE5\x90\x8D\xE7\xA7\xB0:"); // Output Name
    y += inputH + spacing;

    // Output Mode
    outputModeChoice_ = new Fl_Choice(labelW + 10, y, inputW, inputH, "\xE8\xBE\x93\xE5\x87\xBA\xE6\xA8\xA1\xE5\xBC\x8F:"); // Output Mode
    outputModeChoice_->add("\xE6\x96\xB0\xE5\xB7\xA5\xE4\xBD\x9C\xE7\xB0\xBF"); // New Workbook
    outputModeChoice_->add("\xE6\x96\xB0\x45\x78\x63\x65\x6C\xE8\xA1\xA8"); // New Sheet
    outputModeChoice_->value(0);
    y += inputH + spacing;

    // Rule Logic
    ruleLogicChoice_ = new Fl_Choice(labelW + 10, y, inputW, inputH, "\xE8\xA7\x84\xE5\x88\x99\xE9\x80\xBB\xE8\xBE\x91:"); // Rule Logic
    ruleLogicChoice_->add("AND (\xE5\x85\xA8\xE9\x83\xA8\xE6\xBB\xA1\xE8\xB6\xB3)");
    ruleLogicChoice_->add("OR (\xE4\xBB\xBB\xE4\xB8\x80\xE6\xBB\xA1\xE8\xB6\xB3)");
    ruleLogicChoice_->value(1); // Default OR
    y += inputH + spacing;

    // Checkboxes
    enabledCheck_ = new Fl_Check_Button(labelW + 10, y, 100, inputH, "\xE5\x90\xAF\xE7\x94\xA8"); // Enabled
    overwriteCheck_ = new Fl_Check_Button(labelW + 120, y, 150, inputH, "\xE8\xA6\x86\xE7\x9B\x96\xE5\xB7\xB2\xE6\x9C\x89\x45\x78\x63\x65\x6C\xE8\xA1\xA8"); // Overwrite Sheet
    y += inputH + spacing;

    // Rules Browser
    new Fl_Box(10, y, 100, 20, "\xE9\x80\x89\xE6\x8B\xA9\xE8\xA7\x84\xE5\x88\x99:"); // Select Rules
    
    Fl_Button* configExBtn = new Fl_Button(w() - 150, y, 140, 20, "\xE9\x85\x8D\xE7\xBD\xAE\xE6\x8E\x92\xE9\x99\xA4\xE8\xA7\x84\xE5\x88\x99"); // Config Exclusions
    configExBtn->callback(cb_configExclusion, this);
    configExBtn->labelsize(12);
    
    y += 20;
    ruleBrowser_ = new Fl_Check_Browser(10, y, w() - 20, 150);
    y += 150 + spacing;

    // Buttons
    Fl_Button* saveBtn = new Fl_Button(w() - 180, h() - 40, 80, 30, "\xE4\xBF\x9D\xE5\xAD\x98"); // Save
    saveBtn->callback(cb_save, this);
    
    Fl_Button* cancelBtn = new Fl_Button(w() - 90, h() - 40, 80, 30, "\xE5\x8F\x96\xE6\xB6\x88"); // Cancel
    cancelBtn->callback(cb_cancel, this);

    end();
    set_modal();
}

void EditTaskDialog::setAvailableRules(const std::vector<Rule>& rules) {
    availableRules_ = rules;
    ruleBrowser_->clear();
    for (const auto& rule : rules) {
        std::string label = std::to_string(rule.id) + ": " + rule.name;
        ruleBrowser_->add(label.c_str());
    }
}

void EditTaskDialog::setTask(const ProcessingTask& task) {
    currentTask_ = task;
    nameInput_->value(task.taskName.c_str());
    inputPatternInput_->value(task.inputFilenamePattern.c_str());
    inputSheetInput_->value(task.inputSheetName.c_str());
    outputNameInput_->value(task.outputWorkbookName.c_str());
    
    if (task.outputMode == OutputMode::NEW_WORKBOOK) outputModeChoice_->value(0);
    else if (task.outputMode == OutputMode::NEW_SHEET) outputModeChoice_->value(1);
    
    if (task.ruleLogic == RuleLogic::AND) ruleLogicChoice_->value(0);
    else ruleLogicChoice_->value(1);

    enabledCheck_->value(task.enabled ? 1 : 0);
    overwriteCheck_->value(task.overwriteSheet ? 1 : 0);

    pendingExclusions_.clear();
    for (const auto& entry : task.rules) {
        pendingExclusions_[entry.ruleId] = entry.excludeRuleIds;
    }

    // Check rules
    for (int i = 1; i <= ruleBrowser_->nitems(); ++i) {
        // Find if this rule (by index or ID) is in task.rules
        // The browser items are in same order as availableRules_
        if (i-1 < availableRules_.size()) {
            int ruleId = availableRules_[i-1].id;
            bool found = false;
            for (const auto& entry : task.rules) {
                if (entry.ruleId == ruleId) {
                    found = true;
                    break;
                }
            }
            ruleBrowser_->checked(i, found ? 1 : 0);
        }
    }
}

ProcessingTask EditTaskDialog::getTask() const {
    ProcessingTask task = currentTask_;
    task.taskName = nameInput_->value();
    task.inputFilenamePattern = inputPatternInput_->value();
    task.inputSheetName = inputSheetInput_->value();
    task.outputWorkbookName = outputNameInput_->value();
    
    if (outputModeChoice_->value() == 0) task.outputMode = OutputMode::NEW_WORKBOOK;
    else task.outputMode = OutputMode::NEW_SHEET;

    if (ruleLogicChoice_->value() == 0) task.ruleLogic = RuleLogic::AND;
    else task.ruleLogic = RuleLogic::OR;

    task.enabled = enabledCheck_->value() != 0;
    task.overwriteSheet = overwriteCheck_->value() != 0;

    // Collect rules
    task.rules.clear();
    for (int i = 1; i <= ruleBrowser_->nitems(); ++i) {
        if (ruleBrowser_->checked(i)) {
            if (i-1 < availableRules_.size()) {
                TaskRuleEntry entry;
                entry.ruleId = availableRules_[i-1].id;
                auto it = pendingExclusions_.find(entry.ruleId);
                if (it != pendingExclusions_.end()) {
                    entry.excludeRuleIds = it->second;
                }
                task.rules.push_back(entry);
            }
        }
    }

    return task;
}

int EditTaskDialog::showModal() {
    show();
    while (shown()) {
        Fl::wait();
    }
    return result_;
}

void EditTaskDialog::cb_save(Fl_Widget* w, void* v) {
    EditTaskDialog* dlg = (EditTaskDialog*)v;
    dlg->result_ = 1;
    dlg->hide();
}

void EditTaskDialog::cb_cancel(Fl_Widget* w, void* v) {
    EditTaskDialog* dlg = (EditTaskDialog*)v;
    dlg->result_ = 0;
    dlg->hide();
}

void EditTaskDialog::cb_configExclusion(Fl_Widget* w, void* v) {
    ((EditTaskDialog*)v)->configExclusion();
}

void EditTaskDialog::configExclusion() {
    int selectedLine = ruleBrowser_->value(); // 1-based index
    if (selectedLine <= 0) {
        fl_alert("\xE8\xAF\xB7\xE5\x85\x88\xE9\x80\x89\xE4\xB8\xAD\xE4\xB8\x80\xE6\x9D\xA1\xE8\xA7\x84\xE5\x88\x99 (\xE9\xAB\x98\xE4\xBA\xAE)"); // Please select a rule first
        return;
    }
    
    if (selectedLine - 1 >= availableRules_.size()) return;
    
    int ruleId = availableRules_[selectedLine - 1].id;
    
    // Check if the rule is actually checked (enabled for this task)
    if (!ruleBrowser_->checked(selectedLine)) {
        if (fl_choice("\xE8\xAF\xA5\xE8\xA7\x84\xE5\x88\x99\xE6\x9C\xAA\xE8\xA2\xAB\xE5\x8B\xBE\xE9\x80\x89\xEF\xBC\x8C\xE6\x98\xAF\xE5\x90\xA6\xE8\xA6\x81\xE5\x8B\xBE\xE9\x80\x89\xEF\xBC\x9F", "\xE5\x90\xA6", "\xE6\x98\xAF", 0) == 1) {
            ruleBrowser_->checked(selectedLine, 1);
        } else {
            return;
        }
    }

    ExclusionDialog dlg(300, 400, "\xE9\x85\x8D\xE7\xBD\xAE\xE6\x8E\x92\xE9\x99\xA4\xE8\xA7\x84\xE5\x88\x99");
    dlg.setAvailableRules(availableRules_);
    
    if (pendingExclusions_.count(ruleId)) {
        dlg.setExcludedRuleIds(pendingExclusions_[ruleId]);
    }
    
    if (dlg.showModal() == 1) {
        pendingExclusions_[ruleId] = dlg.getExcludedRuleIds();
    }
}
