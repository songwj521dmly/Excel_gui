#pragma once
#include <FL/Fl_Table_Row.H>
#include "ExcelProcessorCore.h"
#include <vector>
#include <string>

class TaskTable : public Fl_Table_Row {
public:
    TaskTable(int x, int y, int w, int h, const char* l = 0);
    ~TaskTable();

    void setTasks(const std::vector<ProcessingTask>& tasks);
    int getSelectedTaskId(); // Returns -1 if no selection

protected:
    void draw_cell(TableContext context, int R, int C, int X, int Y, int W, int H) override;

private:
    std::vector<ProcessingTask> tasks_;
    std::vector<std::string> headers_;
};
