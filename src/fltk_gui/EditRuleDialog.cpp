#include "EditRuleDialog.h"
#include <FL/Fl.H>
#include <FL/fl_ask.H>
#include <string>
#include <iostream>

EditRuleDialog::EditRuleDialog(int w, int h, const char* title)
    : Fl_Double_Window(w, h, title), result_(0) {
    setupUI();
}

EditRuleDialog::~EditRuleDialog() {
}

void EditRuleDialog::setupUI() {
    int x = 10;
    int y = 10;
    int w = this->w();
    int h = this->h();
    int labelW = 70;
    int inputW = w - x * 2 - labelW;
    int rowH = 30;
    int spacing = 10;
    
    // Name
    nameInput_ = new Fl_Input(x + labelW, y, inputW, rowH, "\xE5\x90\x8D\xE7\xA7\xB0:"); // Name:
    y += rowH + spacing;
    
    // Type
    typeChoice_ = new Fl_Choice(x + labelW, y, inputW, rowH, "\xE7\xB1\xBB\xE5\x9E\x8B:"); // Type:
    typeChoice_->add("\xE7\xAD\x9B\xE9\x80\x89"); // Filter
    typeChoice_->add("\xE8\xBD\xAC\xE6\x8D\xA2"); // Transform
    typeChoice_->add("\xE5\x88\xA0\xE9\x99\xA4\xE8\xA1\x8C"); // Delete Row
    typeChoice_->add("\xE6\x8B\x86\xE5\x88\x86"); // Split
    typeChoice_->value(0);
    y += rowH + spacing;
    
    // Enabled
    enabledCheck_ = new Fl_Check_Button(x + labelW, y, inputW, rowH, "\xE5\x90\xAF\xE7\x94\xA8\xE8\xA7\x84\xE5\x88\x99"); // Enable Rule
    y += rowH + spacing;
    
    // Description
    descInput_ = new Fl_Input(x + labelW, y, inputW, rowH, "\xE6\x8F\x8F\xE8\xBF\xB0:"); // Description:
    y += rowH + spacing;
    
    // Priority
    priorityInput_ = new Fl_Int_Input(x + labelW, y, inputW, rowH, "\xE4\xBC\x98\xE5\x85\x88\xE7\xBA\xA7:"); // Priority:
    y += rowH + spacing;
    
    // Conditions Header
    Fl_Box* condLabel = new Fl_Box(x, y, w - 20, 20, "\xE8\xA7\x84\xE5\x88\x99\xE6\x9D\xA1\xE4\xBB\xB6"); // Rule Conditions
    condLabel->align(FL_ALIGN_LEFT | FL_ALIGN_INSIDE);
    condLabel->labelfont(FL_BOLD);
    y += 20 + 5;
    
    // Add Condition Button
    Fl_Button* addCondBtn = new Fl_Button(x, y, 100, 25, "\xE6\xB7\xBB\xE5\x8A\xA0\xE6\x9D\xA1\xE4\xBB\xB6"); // Add Condition
    addCondBtn->callback(cb_addCondition, this);
    addCondBtn->labelsize(12);
    y += 25 + 5;
    
    // Conditions Scroll
    int bottomH = 50;
    int scrollH = h - y - bottomH - 10;
    
    conditionsScroll_ = new Fl_Scroll(x, y, w - 20, scrollH);
    conditionsScroll_->type(Fl_Scroll::VERTICAL);
    conditionsScroll_->box(FL_DOWN_BOX);
    conditionsScroll_->end();
    
    y += scrollH + 10;
    
    // Buttons
    int btnW = 80;
    int btnSpacing = 20;
    int btnX = w / 2 - btnW - btnSpacing / 2;
    
    Fl_Button* saveBtn = new Fl_Button(btnX, y, btnW, 30, "\xE4\xBF\x9D\xE5\xAD\x98"); // Save
    saveBtn->callback(cb_save, this);
    saveBtn->color(FL_BLUE);
    saveBtn->labelcolor(FL_WHITE);
    
    Fl_Button* cancelBtn = new Fl_Button(btnX + btnW + btnSpacing, y, btnW, 30, "\xE5\x8F\x96\xE6\xB6\x88"); // Cancel
    cancelBtn->callback(cb_cancel, this);
    
    end();
    set_modal();
}

void EditRuleDialog::setRule(const Rule& rule) {
    currentRule_ = rule;
    
    nameInput_->value(rule.name.c_str());
    
    switch (rule.type) {
        case RuleType::FILTER: typeChoice_->value(0); break;
        case RuleType::TRANSFORM: typeChoice_->value(1); break;
        case RuleType::DELETE_ROW: typeChoice_->value(2); break;
        case RuleType::SPLIT: typeChoice_->value(3); break;
        default: typeChoice_->value(0); break;
    }
    
    enabledCheck_->value(rule.enabled ? 1 : 0);
    descInput_->value(rule.description.c_str());
    priorityInput_->value(std::to_string(rule.priority).c_str());
    
    refreshConditions();
}

Rule EditRuleDialog::getRule() const {
    return currentRule_;
}

int EditRuleDialog::showModal() {
    show();
    while (shown()) {
        Fl::wait();
    }
    return result_;
}

void EditRuleDialog::refreshConditions() {
    conditionsScroll_->clear();
    conditionWidgets_.clear();
    
    conditionsScroll_->begin();
    
    int y = 0;
    int h = 30;
    int spacing = 5;
    
    for (size_t i = 0; i < currentRule_.conditions.size(); ++i) {
        const auto& cond = currentRule_.conditions[i];
        
        // Col Input (Label: Col)
        Fl_Int_Input* colIn = new Fl_Int_Input(conditionsScroll_->x(), conditionsScroll_->y() + y, 40, h);
        colIn->value(std::to_string(cond.column).c_str());
        colIn->tooltip("\xE5\x88\x97\xE5\x8F\xB7"); // Column Number
        
        // Op Choice
        Fl_Choice* opCh = new Fl_Choice(conditionsScroll_->x() + 45, conditionsScroll_->y() + y, 100, h);
        opCh->add("=");
        opCh->add("<>");
        opCh->add(">");
        opCh->add("<");
        opCh->add(">=");
        opCh->add("<=");
        opCh->add("\xE5\x8C\x85\xE5\x90\xAB"); // Contains
        opCh->add("\xE4\xB8\x8D\xE5\x8C\x85\xE5\x90\xAB"); // Not Contains
        opCh->add("\xE4\xB8\xBA\xE7\xA9\xBA"); // Empty
        opCh->add("\xE4\xB8\x8D\xE4\xB8\xBA\xE7\xA9\xBA"); // Not Empty
        opCh->add("\xE5\xBC\x80\xE5\xA4\xB4\xE6\x98\xAF"); // Starts With
        opCh->add("\xE7\xBB\x93\xE5\xB0\xBE\xE6\x98\xAF"); // Ends With
        
        // Map Operator to index
        int opIdx = 0;
        switch (cond.oper) {
            case Operator::EQUAL: opIdx = 0; break;
            case Operator::NOT_EQUAL: opIdx = 1; break;
            case Operator::GREATER: opIdx = 2; break;
            case Operator::LESS: opIdx = 3; break;
            case Operator::GREATER_EQUAL: opIdx = 4; break;
            case Operator::LESS_EQUAL: opIdx = 5; break;
            case Operator::CONTAINS: opIdx = 6; break;
            case Operator::NOT_CONTAINS: opIdx = 7; break;
            case Operator::EMPTY: opIdx = 8; break;
            case Operator::NOT_EMPTY: opIdx = 9; break;
            case Operator::STARTS_WITH: opIdx = 10; break;
            case Operator::ENDS_WITH: opIdx = 11; break;
            default: opIdx = 0; break;
        }
        opCh->value(opIdx);
        
        // Value Input
        Fl_Input* valIn = new Fl_Input(conditionsScroll_->x() + 150, conditionsScroll_->y() + y, 150, h);
        if (std::holds_alternative<std::string>(cond.value)) {
            valIn->value(std::get<std::string>(cond.value).c_str());
        } else if (std::holds_alternative<int>(cond.value)) {
            valIn->value(std::to_string(std::get<int>(cond.value)).c_str());
        } else if (std::holds_alternative<double>(cond.value)) {
            valIn->value(std::to_string(std::get<double>(cond.value)).c_str());
        }
        
        ConditionWidgets cw;
        cw.colInput = colIn;
        cw.opChoice = opCh;
        cw.valInput = valIn;
        conditionWidgets_.push_back(cw);
        
        y += h + spacing;
    }
    
    conditionsScroll_->end();
    conditionsScroll_->redraw();
}

void EditRuleDialog::cb_save(Fl_Widget* w, void* v) {
    EditRuleDialog* dlg = (EditRuleDialog*)v;
    
    // Update currentRule_ from UI
    dlg->currentRule_.name = dlg->nameInput_->value();
    
    int typeIdx = dlg->typeChoice_->value();
    switch (typeIdx) {
        case 0: dlg->currentRule_.type = RuleType::FILTER; break;
        case 1: dlg->currentRule_.type = RuleType::TRANSFORM; break;
        case 2: dlg->currentRule_.type = RuleType::DELETE_ROW; break;
        case 3: dlg->currentRule_.type = RuleType::SPLIT; break;
    }
    
    dlg->currentRule_.enabled = dlg->enabledCheck_->value() != 0;
    dlg->currentRule_.description = dlg->descInput_->value();
    dlg->currentRule_.priority = std::atoi(dlg->priorityInput_->value());
    
    // Update conditions
    dlg->currentRule_.conditions.clear();
    for (const auto& cw : dlg->conditionWidgets_) {
        RuleCondition cond;
        cond.column = std::atoi(cw.colInput->value());
        
        int opIdx = cw.opChoice->value();
        switch (opIdx) {
            case 0: cond.oper = Operator::EQUAL; break;
            case 1: cond.oper = Operator::NOT_EQUAL; break;
            case 2: cond.oper = Operator::GREATER; break;
            case 3: cond.oper = Operator::LESS; break;
            case 4: cond.oper = Operator::GREATER_EQUAL; break;
            case 5: cond.oper = Operator::LESS_EQUAL; break;
            case 6: cond.oper = Operator::CONTAINS; break;
            case 7: cond.oper = Operator::NOT_CONTAINS; break;
            case 8: cond.oper = Operator::EMPTY; break;
            case 9: cond.oper = Operator::NOT_EMPTY; break;
            case 10: cond.oper = Operator::STARTS_WITH; break;
            case 11: cond.oper = Operator::ENDS_WITH; break;
        }
        
        cond.value = std::string(cw.valInput->value());
        dlg->currentRule_.conditions.push_back(cond);
    }
    
    dlg->result_ = 1;
    dlg->hide();
}

void EditRuleDialog::cb_cancel(Fl_Widget* w, void* v) {
    EditRuleDialog* dlg = (EditRuleDialog*)v;
    dlg->result_ = 0;
    dlg->hide();
}

void EditRuleDialog::cb_addCondition(Fl_Widget* w, void* v) {
    EditRuleDialog* dlg = (EditRuleDialog*)v;
    
    // Save current values first!
    // (Simplification: We reconstruct from currentRule_ in refresh, so we should update currentRule_ first)
    // But refreshConditions clears widgets, so we lose text if not saved.
    // Let's temporary save.
    
    // Quick hack: just push a new default condition and re-init from currentRule_ + saved edits?
    // Better: Update currentRule_ from widgets, then add, then refresh.
    
    dlg->currentRule_.conditions.clear();
    for (const auto& cw : dlg->conditionWidgets_) {
        RuleCondition cond;
        cond.column = std::atoi(cw.colInput->value());
        // ... (op mapping)
         int opIdx = cw.opChoice->value();
        switch (opIdx) {
            case 0: cond.oper = Operator::EQUAL; break;
            case 1: cond.oper = Operator::NOT_EQUAL; break;
            case 2: cond.oper = Operator::GREATER; break;
            case 3: cond.oper = Operator::LESS; break;
            case 4: cond.oper = Operator::GREATER_EQUAL; break;
            case 5: cond.oper = Operator::LESS_EQUAL; break;
            case 6: cond.oper = Operator::CONTAINS; break;
            case 7: cond.oper = Operator::NOT_CONTAINS; break;
            case 8: cond.oper = Operator::EMPTY; break;
            case 9: cond.oper = Operator::NOT_EMPTY; break;
            case 10: cond.oper = Operator::STARTS_WITH; break;
            case 11: cond.oper = Operator::ENDS_WITH; break;
        }
        cond.value = std::string(cw.valInput->value());
        dlg->currentRule_.conditions.push_back(cond);
    }
    
    RuleCondition newCond;
    newCond.column = 1;
    dlg->currentRule_.conditions.push_back(newCond);
    
    dlg->refreshConditions();
}
