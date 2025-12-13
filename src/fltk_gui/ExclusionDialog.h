#pragma once
#include <FL/Fl_Double_Window.H>
#include <FL/Fl_Check_Browser.H>
#include <FL/Fl_Button.H>
#include <vector>
#include "ExcelProcessorCore.h"

class ExclusionDialog : public Fl_Double_Window {
public:
    ExclusionDialog(int w, int h, const char* title = 0);
    ~ExclusionDialog();

    void setAvailableRules(const std::vector<Rule>& rules);
    void setExcludedRuleIds(const std::vector<int>& ids);
    std::vector<int> getExcludedRuleIds() const;

    int showModal();

private:
    void setupUI();
    static void cb_save(Fl_Widget*, void*);
    static void cb_cancel(Fl_Widget*, void*);

    Fl_Check_Browser* ruleBrowser_;
    std::vector<Rule> availableRules_;
    int result_;
};
