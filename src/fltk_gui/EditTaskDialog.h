#pragma once
#include <FL/Fl_Double_Window.H>
#include <FL/Fl_Input.H>
#include <FL/Fl_Check_Button.H>
#include <FL/Fl_Choice.H>
#include <FL/Fl_Button.H>
#include <FL/Fl_Check_Browser.H>
#include <FL/Fl_Box.H>
#include "ExcelProcessorCore.h"
#include <vector>

class EditTaskDialog : public Fl_Double_Window {
public:
    EditTaskDialog(int w, int h, const char* title = 0);
    ~EditTaskDialog();

    void setTask(const ProcessingTask& task);
    ProcessingTask getTask() const;
    void setAvailableRules(const std::vector<Rule>& rules);
    int showModal();
private:
    void setupUI();
    static void cb_save(Fl_Widget*, void*);
    static void cb_cancel(Fl_Widget*, void*);
    static void cb_configExclusion(Fl_Widget*, void*);

    void configExclusion();

    Fl_Input* nameInput_;
    Fl_Input* inputPatternInput_;
    Fl_Input* inputSheetInput_;
    Fl_Input* outputNameInput_;
    Fl_Choice* outputModeChoice_;
    Fl_Choice* ruleLogicChoice_;
    Fl_Check_Button* enabledCheck_;
    Fl_Check_Button* overwriteCheck_;
    Fl_Check_Browser* ruleBrowser_;
    
    ProcessingTask currentTask_;
    std::vector<Rule> availableRules_;
    std::map<int, std::vector<int>> pendingExclusions_; // RuleID -> ExcludedRuleIDs
    int result_;
};
