#include "ExclusionDialog.h"
#include <FL/Fl.H>
#include <string>

ExclusionDialog::ExclusionDialog(int w, int h, const char* title)
    : Fl_Double_Window(w, h, title), result_(0) {
    setupUI();
}

ExclusionDialog::~ExclusionDialog() {
}

void ExclusionDialog::setupUI() {
    // Browser
    ruleBrowser_ = new Fl_Check_Browser(10, 10, w() - 20, h() - 60);
    
    // Buttons
    Fl_Button* saveBtn = new Fl_Button(w() - 180, h() - 40, 80, 30, "\xE7\xA1\xAE\xE5\xAE\x9A"); // Confirm
    saveBtn->callback(cb_save, this);
    
    Fl_Button* cancelBtn = new Fl_Button(w() - 90, h() - 40, 80, 30, "\xE5\x8F\x96\xE6\xB6\x88"); // Cancel
    cancelBtn->callback(cb_cancel, this);

    end();
    set_modal();
}

void ExclusionDialog::setAvailableRules(const std::vector<Rule>& rules) {
    availableRules_ = rules;
    ruleBrowser_->clear();
    for (const auto& rule : rules) {
        std::string label = std::to_string(rule.id) + ": " + rule.name;
        ruleBrowser_->add(label.c_str());
    }
}

void ExclusionDialog::setExcludedRuleIds(const std::vector<int>& ids) {
    for (int i = 1; i <= ruleBrowser_->nitems(); ++i) {
        if (i-1 < availableRules_.size()) {
            int ruleId = availableRules_[i-1].id;
            bool found = false;
            for (int id : ids) {
                if (id == ruleId) {
                    found = true;
                    break;
                }
            }
            ruleBrowser_->checked(i, found ? 1 : 0);
        }
    }
}

std::vector<int> ExclusionDialog::getExcludedRuleIds() const {
    std::vector<int> ids;
    for (int i = 1; i <= ruleBrowser_->nitems(); ++i) {
        if (ruleBrowser_->checked(i)) {
            if (i-1 < availableRules_.size()) {
                ids.push_back(availableRules_[i-1].id);
            }
        }
    }
    return ids;
}

int ExclusionDialog::showModal() {
    show();
    while (shown()) {
        Fl::wait();
    }
    return result_;
}

void ExclusionDialog::cb_save(Fl_Widget* w, void* v) {
    ExclusionDialog* dlg = (ExclusionDialog*)v;
    dlg->result_ = 1;
    dlg->hide();
}

void ExclusionDialog::cb_cancel(Fl_Widget* w, void* v) {
    ExclusionDialog* dlg = (ExclusionDialog*)v;
    dlg->result_ = 0;
    dlg->hide();
}
