#pragma once

#include <FL/Fl_Double_Window.H>
#include <FL/Fl_Tabs.H>
#include <FL/Fl_Group.H>
#include <FL/Fl_Button.H>
#include <FL/Fl_Text_Display.H>
#include <FL/Fl_Text_Buffer.H>
#include <FL/Fl_Box.H>
#include <FL/Fl_Progress.H>
#include "ExcelProcessorCore.h"
#include "RuleTable.h"
#include "TaskTable.h"
#include "PreviewTable.h"
#include <memory>
#include <vector>
#include <string>

class FltkMainWindow : public Fl_Double_Window {
public:
    FltkMainWindow(int w, int h, const char* title = 0);
    ~FltkMainWindow();

    void updateRuleTable();
    void updateTaskTable();
    void logMessage(const std::string& msg);
    static void awakeCallback(void* data);

private:
    void setupUI();
    Fl_Group* createTaskTab(int x, int y, int w, int h);
    Fl_Group* createRuleTab(int x, int y, int w, int h);
    Fl_Group* createPreviewTab(int x, int y, int w, int h);
    Fl_Group* createAboutTab(int x, int y, int w, int h);
    Fl_Group* createLogGroup(int x, int y, int w, int h);
    Fl_Group* createButtonGroup(int x, int y, int w, int h);

    // Callbacks
    static void cb_addRule(Fl_Widget*, void*);
    static void cb_editRule(Fl_Widget*, void*);
    static void cb_deleteRule(Fl_Widget*, void*);
    static void cb_addTask(Fl_Widget*, void*);
    static void cb_editTask(Fl_Widget*, void*);
    static void cb_deleteTask(Fl_Widget*, void*);
    static void cb_runTasks(Fl_Widget*, void*);
    static void cb_loadConfig(Fl_Widget*, void*);
    static void cb_saveConfig(Fl_Widget*, void*);
    static void cb_preview(Fl_Widget*, void*);
    static void cb_process(Fl_Widget*, void*);

    void addRule();
    void editRule();
    void deleteRule();
    void addTask();
    void editTask();
    void deleteTask();
    void runTasks();
    void loadConfig();
    void saveConfig();
    void preview();
    void process();

    std::unique_ptr<ExcelProcessorCore> processor_;
    
    Fl_Tabs* tabs_;
    RuleTable* ruleTable_;
    TaskTable* taskTable_;
    PreviewTable* previewTable_;
    Fl_Text_Buffer* logBuffer_;
    Fl_Text_Display* logDisplay_;
    Fl_Box* statusLabel_;
    Fl_Progress* progressBar_;
    Fl_Button* processButton_;
    Fl_Box* configNameLabel_;
    
    std::string currentConfigFile_;
    std::string currentInputFile_;
};
