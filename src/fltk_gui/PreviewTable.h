#pragma once
#include <FL/Fl_Table_Row.H>
#include "ExcelProcessorCore.h"
#include <vector>
#include <string>

class PreviewTable : public Fl_Table_Row {
public:
    PreviewTable(int x, int y, int w, int h, const char* l = 0);
    ~PreviewTable();

    void setData(const std::vector<DataRow>& data);
    void clear();

protected:
    void draw_cell(TableContext context, int R, int C, int X, int Y, int W, int H) override;

private:
    std::vector<DataRow> data_;
    std::vector<std::string> headers_; // We might need to deduce headers or just use Col A, Col B...
};
