#include "PreviewTable.h"
#include <FL/fl_draw.H>
#include <string>
#include <algorithm>

PreviewTable::PreviewTable(int x, int y, int w, int h, const char* l)
    : Fl_Table_Row(x, y, w, h, l) {
    rows(0);
    cols(0);
    col_header(1);
    col_resize(1);
    row_header(1);
    row_resize(1);
    
    end();
}

PreviewTable::~PreviewTable() {
}

void PreviewTable::setData(const std::vector<DataRow>& data) {
    data_ = data;
    rows((int)data_.size());
    
    int maxCols = 0;
    for (const auto& row : data_) {
        maxCols = (std::max)(maxCols, (int)row.data.size());
    }
    cols(maxCols);
    
    // Generate simple headers A, B, C...
    headers_.clear();
    for (int i = 0; i < maxCols; ++i) {
        std::string h = "";
        int n = i;
        while (n >= 0) {
            h.insert(0, 1, (char)('A' + n % 26));
            n = n / 26 - 1;
        }
        headers_.push_back(h);
    }
    
    // Default col width
    for (int i = 0; i < maxCols; ++i) {
        col_width(i, 100);
    }
    
    redraw();
}

void PreviewTable::clear() {
    data_.clear();
    rows(0);
    cols(0);
    redraw();
}

void PreviewTable::draw_cell(TableContext context, int R, int C, int X, int Y, int W, int H) {
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
            
        case CONTEXT_ROW_HEADER:
            fl_push_clip(X, Y, W, H);
            fl_draw_box(FL_THIN_UP_BOX, X, Y, W, H, FL_BACKGROUND_COLOR);
            fl_color(FL_BLACK);
            {
                std::string s = std::to_string(R + 1);
                fl_draw(s.c_str(), X, Y, W, H, FL_ALIGN_CENTER);
            }
            fl_pop_clip();
            return;
            
        case CONTEXT_CELL:
            fl_push_clip(X, Y, W, H);
            // Background
            fl_draw_box(FL_FLAT_BOX, X, Y, W, H, FL_WHITE);
            fl_color(FL_BLACK);
            
            // Text
            if (R < data_.size() && C < data_[R].data.size()) {
                std::string text;
                const auto& val = data_[R].data[C];
                
                if (std::holds_alternative<std::string>(val)) {
                    text = std::get<std::string>(val);
                } else if (std::holds_alternative<int>(val)) {
                    text = std::to_string(std::get<int>(val));
                } else if (std::holds_alternative<double>(val)) {
                    text = std::to_string(std::get<double>(val));
                } else if (std::holds_alternative<bool>(val)) {
                    text = std::get<bool>(val) ? "TRUE" : "FALSE";
                }
                // Date handling skipped for simplicity or add if needed
                
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
