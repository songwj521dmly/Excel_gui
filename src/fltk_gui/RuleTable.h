#pragma once
#include <FL/Fl_Table_Row.H>
#include "ExcelProcessorCore.h"
#include <vector>
#include <string>

class RuleTable : public Fl_Table_Row {
public:
    RuleTable(int x, int y, int w, int h, const char* l = 0);
    ~RuleTable();

    void setRules(const std::vector<Rule>& rules);
    int getSelectedRuleId(); // Returns -1 if no selection

protected:
    void draw_cell(TableContext context, int R, int C, int X, int Y, int W, int H) override;

private:
    std::vector<Rule> rules_;
    std::vector<std::string> headers_;
};
