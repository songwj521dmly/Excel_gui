#pragma once
#include <FL/Fl_Double_Window.H>
#include <FL/Fl_Input.H>
#include <FL/Fl_Int_Input.H>
#include <FL/Fl_Check_Button.H>
#include <FL/Fl_Choice.H>
#include <FL/Fl_Button.H>
#include <FL/Fl_Scroll.H>
#include <FL/Fl_Box.H>
#include "ExcelProcessorCore.h"
#include <vector>

class EditRuleDialog : public Fl_Double_Window {
public:
    EditRuleDialog(int w, int h, const char* title = 0);
    ~EditRuleDialog();

    void setRule(const Rule& rule);
    Rule getRule() const;
    int showModal(); // Returns 1 if OK, 0 if Cancel

private:
    void setupUI();
    void refreshConditions();
    
    static void cb_save(Fl_Widget*, void*);
    static void cb_cancel(Fl_Widget*, void*);
    static void cb_addCondition(Fl_Widget*, void*);
    
    // Helper to create condition row
    void createConditionRow(int y, int index, const RuleCondition& cond);

    Fl_Input* nameInput_;
    Fl_Choice* typeChoice_;
    Fl_Check_Button* enabledCheck_;
    Fl_Input* descInput_;
    Fl_Int_Input* priorityInput_;
    
    Fl_Scroll* conditionsScroll_;
    
    Rule currentRule_;
    int result_; // 1=Save, 0=Cancel
    
    // Store widgets for conditions to read back values
    struct ConditionWidgets {
        Fl_Int_Input* colInput;
        Fl_Choice* opChoice;
        Fl_Input* valInput;
    };
    std::vector<ConditionWidgets> conditionWidgets_;
};
