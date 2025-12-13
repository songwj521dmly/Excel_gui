#include "TaskTable.h"
#include <FL/fl_draw.H>
#include <string>

TaskTable::TaskTable(int x, int y, int w, int h, const char* l)
    : Fl_Table_Row(x, y, w, h, l) {
    headers_ = { "ID", "\xE4\xBB\xBB\xE5\x8A\xA1\xE5\x90\x8D\xE7\xA7\xB0", "\xE8\xBE\x93\xE5\x85\xA5\xE6\xA8\xA1\xE5\xBC\x8F", "\xE8\xBE\x93\xE5\x87\xBA", "\xE7\x8A\xB6\xE6\x80\x81" }; // ID, Name, Input Pattern, Output, Status
    
    rows(0);
    cols((int)headers_.size());
    col_header(1);
    col_resize(1);
    row_header(0);
    row_resize(0);
    
    // Set column widths
    col_width(0, 40);  // ID
    col_width(1, 150); // Name
    col_width(2, 150); // Input Pattern
    col_width(3, 150); // Output
    col_width(4, 60);  // Status
    
    end();
    type(SELECT_SINGLE);
}

TaskTable::~TaskTable() {
}

void TaskTable::setTasks(const std::vector<ProcessingTask>& tasks) {
    tasks_ = tasks;
    rows((int)tasks_.size());
    redraw();
}

int TaskTable::getSelectedTaskId() {
    for (int i = 0; i < rows(); ++i) {
        if (row_selected(i)) {
            if (i >= 0 && i < tasks_.size()) {
                return tasks_[i].id;
            }
        }
    }
    return -1;
}

void TaskTable::draw_cell(TableContext context, int R, int C, int X, int Y, int W, int H) {
    switch (context) {
        case CONTEXT_STARTPAGE:
            fl_font(FL_HELVETICA, 14);
            return;
            
        case CONTEXT_COL_HEADER:
            fl_push_clip(X, Y, W, H);
            fl_draw_box(FL_THIN_UP_BOX, X, Y, W, H, FL_BACKGROUND_COLOR);
            fl_color(FL_BLACK);
            fl_draw(headers_[C].c_str(), X, Y, W, H, FL_ALIGN_CENTER);
            fl_pop_clip();
            return;
            
        case CONTEXT_CELL:
            fl_push_clip(X, Y, W, H);
            // Background
            fl_color(row_selected(R) ? FL_SELECTION_COLOR : FL_WHITE);
            fl_rectf(X, Y, W, H);
            
            // Text
            fl_color(row_selected(R) ? FL_WHITE : FL_BLACK);
            if (R < tasks_.size()) {
                std::string text;
                const auto& task = tasks_[R];
                switch (C) {
                    case 0: text = std::to_string(task.id); break;
                    case 1: text = task.taskName; break;
                    case 2: text = task.inputFilenamePattern; break;
                    case 3: text = task.outputWorkbookName; break;
                    case 4: text = task.enabled ? "\xE5\x90\xAF\xE7\x94\xA8" : "\xE7\xA6\x81\xE7\x94\xA8"; break; // Enabled/Disabled
                }
                fl_draw(text.c_str(), X + 2, Y, W - 4, H, FL_ALIGN_LEFT);
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
