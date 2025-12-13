#include "RuleTable.h"
#include <FL/fl_draw.H>
#include <string>

RuleTable::RuleTable(int x, int y, int w, int h, const char* l)
    : Fl_Table_Row(x, y, w, h, l) {
    headers_ = {"ID", "\xE5\x90\x8D\xE7\xA7\xB0", "\xE7\xB1\xBB\xE5\x9E\x8B", "\xE6\x9D\xA1\xE4\xBB\xB6\xE6\x95\xB0", "\xE5\x90\xAF\xE7\x94\xA8", "\xE6\x8F\x8F\xE8\xBF\xB0", "\xE4\xBC\x98\xE5\x85\x88\xE7\xBA\xA7"};
    
    rows(0);
    cols(7);
    col_header(1);
    col_resize(1);
    
    // Set column widths
    col_width(0, 50);  // ID
    col_width(1, 150); // Name
    col_width(2, 80);  // Type
    col_width(3, 60);  // Count
    col_width(4, 50);  // Enabled
    col_width(5, 200); // Desc
    col_width(6, 60);  // Priority
    
    end();
    type(SELECT_SINGLE);
}

RuleTable::~RuleTable() {
}

void RuleTable::setRules(const std::vector<Rule>& rules) {
    rules_ = rules;
    rows((int)rules_.size());
    redraw();
}

int RuleTable::getSelectedRuleId() {
    for (int i = 0; i < rows(); i++) {
        if (row_selected(i)) {
            if (i < rules_.size()) {
                return rules_[i].id;
            }
        }
    }
    return -1;
}

void RuleTable::draw_cell(TableContext context, int R, int C, int X, int Y, int W, int H) {
    switch (context) {
        case CONTEXT_STARTPAGE:
            fl_font(FL_HELVETICA, 14);
            return;
            
        case CONTEXT_COL_HEADER:
            fl_push_clip(X, Y, W, H);
            fl_draw_box(FL_THIN_UP_BOX, X, Y, W, H, FL_BACKGROUND_COLOR);
            fl_color(FL_BLACK);
            if (C < headers_.size()) {
                fl_draw(headers_[C].c_str(), X, Y, W, H, FL_ALIGN_CENTER);
            }
            fl_pop_clip();
            return;
            
        case CONTEXT_CELL:
            fl_push_clip(X, Y, W, H);
            // Background
            if (row_selected(R)) {
                fl_draw_box(FL_FLAT_BOX, X, Y, W, H, FL_SELECTION_COLOR);
                fl_color(FL_WHITE);
            } else {
                fl_draw_box(FL_FLAT_BOX, X, Y, W, H, FL_WHITE);
                fl_color(FL_BLACK);
            }
            
            // Text
            if (R < rules_.size()) {
                std::string text;
                const auto& rule = rules_[R];
                switch (C) {
                    case 0: text = std::to_string(rule.id); break;
                    case 1: text = rule.name; break;
                    case 2: 
                        switch(rule.type) {
                            case RuleType::FILTER: text = "\xE7\xAD\x9B\xE9\x80\x89"; break;
                            case RuleType::TRANSFORM: text = "\xE8\xBD\xAC\xE6\x8D\xA2"; break;
                            case RuleType::DELETE_ROW: text = "\xE5\x88\xA0\xE9\x99\xA4"; break;
                            case RuleType::SPLIT: text = "\xE6\x8B\x86\xE5\x88\x86"; break;
                            default: text = "\xE6\x9C\xAA\xE7\x9F\xA5"; break;
                        }
                        break;
                    case 3: text = std::to_string(rule.conditions.size()); break;
                    case 4: text = rule.enabled ? "\xE6\x98\xAF" : "\xE5\x90\xA6"; break;
                    case 5: text = rule.description; break;
                    case 6: text = std::to_string(rule.priority); break;
                }
                fl_draw(text.c_str(), X + 2, Y, W - 4, H, FL_ALIGN_LEFT); // Pad left
            }
            
            // Border
            fl_color(FL_LIGHT2);
            fl_rect(X, Y, W, H);
            
            fl_pop_clip();
            return;
            
        default:
            return;
    }
}
