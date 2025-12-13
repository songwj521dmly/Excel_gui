#include "FltkMainWindow.h"
#include "EditRuleDialog.h"
#include "EditTaskDialog.h"
#include <FL/fl_ask.H>
#include <FL/Fl_Native_File_Chooser.H>
#include <chrono>
#include <ctime>
#include <sstream>
#include <iomanip>
#include <iostream>

// Helper to get current timestamp string
std::string getTimestamp() {
    auto now = std::chrono::system_clock::now();
    auto time = std::chrono::system_clock::to_time_t(now);
    std::stringstream ss;
    ss << std::put_time(std::localtime(&time), "[%Y-%m-%d %H:%M:%S] ");
    return ss.str();
}

// Static awake callback wrapper
void FltkMainWindow::awakeCallback(void* data) {
    auto* func = static_cast<std::function<void()>*>(data);
    if (func) {
        (*func)();
        delete func;
    }
}

// Helper to schedule on main thread
void scheduleOnMainThread(std::function<void()> func) {
    auto* funcPtr = new std::function<void()>(func);
    Fl::awake(FltkMainWindow::awakeCallback, funcPtr);
}

FltkMainWindow::FltkMainWindow(int w, int h, const char* title) 
    : Fl_Double_Window(w, h, title) {
    
    processor_ = std::make_unique<ExcelProcessorCore>();
    
    // Setup Callbacks
    processor_->setLogger([this](const std::string& msg) {
        scheduleOnMainThread([this, msg]() {
            logMessage(msg);
        });
    });

    processor_->setProgressCallback([this](int progress, const std::string& msg) {
        scheduleOnMainThread([this, progress, msg]() {
            if (progressBar_) {
                progressBar_->value(progress);
                progressBar_->label(msg.c_str());
            }
            if (!msg.empty()) {
                logMessage(msg);
            }
        });
    });

    setupUI();
    
    // Resize behavior
    resizable(tabs_);
}

FltkMainWindow::~FltkMainWindow() {
}

void FltkMainWindow::setupUI() {
    int w = this->w();
    int h = this->h();
    
    int statusBarH = 30;
    int buttonGroupH = 60;
    int logGroupH = 120;
    int tabY = 10;
    int tabH = h - statusBarH - buttonGroupH - logGroupH - tabY - 10;
    
    // 1. Tabs
    tabs_ = new Fl_Tabs(10, tabY, w - 20, tabH);
    
    // Create tabs
    createTaskTab(10, tabY + 30, w - 20, tabH - 30);
    createRuleTab(10, tabY + 30, w - 20, tabH - 30);
    createPreviewTab(10, tabY + 30, w - 20, tabH - 30);
    createAboutTab(10, tabY + 30, w - 20, tabH - 30);
    
    tabs_->end();
    
    // 2. Log Group
    createLogGroup(10, tabY + tabH + 5, w - 20, logGroupH);
    
    // 3. Button Group
    createButtonGroup(10, tabY + tabH + logGroupH + 10, w - 20, buttonGroupH);
    
    // 4. Status Bar
    Fl_Group* statusGroup = new Fl_Group(0, h - statusBarH, w, statusBarH);
    statusGroup->box(FL_FLAT_BOX);
    
    statusLabel_ = new Fl_Box(5, h - statusBarH, 200, statusBarH, "\xE5\xB0\xB1\xE7\xBB\xAA"); // Ready
    statusLabel_->align(FL_ALIGN_LEFT | FL_ALIGN_INSIDE);
    
    progressBar_ = new Fl_Progress(210, h - statusBarH + 5, w - 220, statusBarH - 10);
    progressBar_->minimum(0);
    progressBar_->maximum(100);
    progressBar_->value(0);
    progressBar_->hide(); // Initially hidden
    
    statusGroup->end();
    
    updateRuleTable();
    updateTaskTable();
}

Fl_Group* FltkMainWindow::createTaskTab(int x, int y, int w, int h) {
    Fl_Group* group = new Fl_Group(x, y, w, h, "\xE5\xA4\x9A\xE5\xB7\xA5\xE4\xBD\x9C\xE7\xB0\xBF\xE4\xBB\xBB\xE5\x8A\xA1"); // Multi-workbook Tasks
    group->box(FL_FLAT_BOX);
    
    int btnY = y + 10;
    int btnH = 30;
    int btnW = 100;
    int spacing = 10;
    int curX = x + 10;
    
    Fl_Button* addBtn = new Fl_Button(curX, btnY, btnW, btnH, "\xE6\xB7\xBB\xE5\x8A\xA0\xE4\xBB\xBB\xE5\x8A\xA1"); // Add Task
    addBtn->callback(cb_addTask, this);
    addBtn->labelcolor(FL_BLUE);
    addBtn->labelfont(FL_BOLD);
    curX += btnW + spacing;
    
    Fl_Button* editBtn = new Fl_Button(curX, btnY, btnW, btnH, "\xE7\xBC\x96\xE8\xBE\x91\xE4\xBB\xBB\xE5\x8A\xA1"); // Edit Task
    editBtn->callback(cb_editTask, this);
    curX += btnW + spacing;
    
    Fl_Button* delBtn = new Fl_Button(curX, btnY, btnW, btnH, "\xE5\x88\xA0\xE9\x99\xA4\xE4\xBB\xBB\xE5\x8A\xA1"); // Delete Task
    delBtn->callback(cb_deleteTask, this);
    curX += btnW + spacing;

    Fl_Button* runBtn = new Fl_Button(curX, btnY, btnW, btnH, "\xE8\xBF\x90\xE8\xA1\x8C\xE4\xBB\xBB\xE5\x8A\xA1"); // Run Tasks
    runBtn->callback(cb_runTasks, this);
    runBtn->labelcolor(FL_RED);
    runBtn->labelfont(FL_BOLD);
    
    // Task Table
    int tableY = btnY + btnH + 10;
    int tableH = h - (tableY - y) - 10;
    
    taskTable_ = new TaskTable(x + 10, tableY, w - 20, tableH);
    
    group->end();
    return group;
}

Fl_Group* FltkMainWindow::createRuleTab(int x, int y, int w, int h) {
    Fl_Group* group = new Fl_Group(x, y, w, h, "\xE8\xA7\x84\xE5\x88\x99\xE9\x85\x8D\xE7\xBD\xAE"); // Rule Config
    group->box(FL_FLAT_BOX);
    
    int btnY = y + 10;
    int btnH = 30;
    int btnW = 100;
    int spacing = 10;
    int curX = x + 10;
    
    Fl_Button* addBtn = new Fl_Button(curX, btnY, btnW, btnH, "\xE6\xB7\xBB\xE5\x8A\xA0\xE8\xA7\x84\xE5\x88\x99"); // Add Rule
    addBtn->callback(cb_addRule, this);
    addBtn->labelcolor(FL_GREEN);
    addBtn->labelfont(FL_BOLD);
    curX += btnW + spacing;
    
    Fl_Button* editBtn = new Fl_Button(curX, btnY, btnW, btnH, "\xE7\xBC\x96\xE8\xBE\x91\xE8\xA7\x84\xE5\x88\x99"); // Edit Rule
    editBtn->callback(cb_editRule, this);
    curX += btnW + spacing;
    
    Fl_Button* delBtn = new Fl_Button(curX, btnY, btnW, btnH, "\xE5\x88\xA0\xE9\x99\xA4\xE8\xA7\x84\xE5\x88\x99"); // Delete Rule
    delBtn->callback(cb_deleteRule, this);
    curX += btnW + spacing;
    
    // Spacer
    curX += spacing * 2;
    
    configNameLabel_ = new Fl_Box(curX, btnY, 200, btnH, "\xE6\x9C\xAA\xE5\x8A\xA0\xE8\xBD\xBD\xE9\x85\x8D\xE7\xBD\xAE"); // No Config Loaded
    configNameLabel_->align(FL_ALIGN_LEFT | FL_ALIGN_INSIDE);
    configNameLabel_->labelcolor(FL_DARK3);
    curX += 200 + spacing;
    
    Fl_Button* loadBtn = new Fl_Button(curX, btnY, btnW, btnH, "\xE5\x8A\xA0\xE8\xBD\xBD\xE9\x85\x8D\xE7\xBD\xAE"); // Load Config
    loadBtn->callback(cb_loadConfig, this);
    curX += btnW + spacing;
    
    Fl_Button* saveBtn = new Fl_Button(curX, btnY, btnW, btnH, "\xE4\xBF\x9D\xE5\xAD\x98\xE9\x85\x8D\xE7\xBD\xAE"); // Save Config
    saveBtn->callback(cb_saveConfig, this);
    
    // Rule Table
    int tableY = btnY + btnH + 10;
    int tableH = h - (tableY - y) - 10;
    
    ruleTable_ = new RuleTable(x + 10, tableY, w - 20, tableH);
    
    group->end();
    return group;
}

Fl_Group* FltkMainWindow::createPreviewTab(int x, int y, int w, int h) {
    Fl_Group* group = new Fl_Group(x, y, w, h, "\xE6\x95\xB0\xE6\x8D\xAE\xE9\xA2\x84\xE8\xA7\x88"); // Data Preview
    group->box(FL_FLAT_BOX);
    
    previewTable_ = new PreviewTable(x + 10, y + 10, w - 20, h - 20);
    
    group->end();
    return group;
}

Fl_Group* FltkMainWindow::createAboutTab(int x, int y, int w, int h) {
    Fl_Group* group = new Fl_Group(x, y, w, h, "\xE5\x85\xB3\xE4\xBA\x8E"); // About
    group->box(FL_FLAT_BOX);
    
    int curY = y + 50;
    Fl_Box* title = new Fl_Box(x, curY, w, 40, "Excel\xE6\x95\xB0\xE6\x8D\xAE\xE5\xA4\x84\xE7\x90\x86\xE5\xB7\xA5\xE5\x85\xB7-\xE6\xA8\xAA\xE6\xA2\x81\xE4\xB8\x8B\xE6\x96\x99\xE4\xB8\x93\xE4\xB8\x9A\xE7\xBB\x84");
    title->labelsize(24);
    title->labelfont(FL_BOLD);
    curY += 50;
    
    Fl_Box* ver = new Fl_Box(x, curY, w, 30, "\xE7\x89\x88\xE6\x9C\xAC: v1.0.0 (FLTK\xE7\x89\x88)"); // Version
    ver->labelsize(16);
    ver->labelcolor(FL_DARK3);
    curY += 40;
    
    Fl_Box* desc = new Fl_Box(x, curY, w, 30, "\xE6\x8F\x90\xE4\xBE\x9B\xE9\xAB\x98\xE6\x95\x88\xE7\x9A\x84""Excel\xE6\x95\xB0\xE6\x8D\xAE\xE5\xA4\x84\xE7\x90\x86\xE3\x80\x81\xE8\xA7\x84\xE5\x88\x99\xE7\xAD\x9B\xE9\x80\x89\xE4\xB8\x8E\xE5\xA4\x9A\xE4\xBB\xBB\xE5\x8A\xA1\xE5\x88\x86\xE5\x8F\x91\xE5\x8A\x9F\xE8\x83\xBD\xE3\x80\x82");
    curY += 60;
    
    Fl_Box* dev = new Fl_Box(x, curY, w, 30, "\xE5\xBC\x80\xE5\x8F\x91\xE8\x80\x85\xE4\xBF\xA1\xE6\x81\xAF"); // Developer Info
    dev->labelsize(18);
    dev->labelfont(FL_BOLD);
    curY += 40;
    
    Fl_Box* email = new Fl_Box(x, curY, w, 30, "Email: songwj@yutong.com");
    email->labelcolor(FL_BLUE);
    
    group->end();
    return group;
}

Fl_Group* FltkMainWindow::createLogGroup(int x, int y, int w, int h) {
    Fl_Group* group = new Fl_Group(x, y, w, h);
    group->box(FL_NO_BOX);
    
    Fl_Box* title = new Fl_Box(x, y, w, 20, "\xE6\x97\xA5\xE5\xBF\x97\xE8\xBE\x93\xE5\x87\xBA"); // Log Output
    title->align(FL_ALIGN_LEFT | FL_ALIGN_INSIDE);
    
    logDisplay_ = new Fl_Text_Display(x, y + 20, w, h - 20);
    logBuffer_ = new Fl_Text_Buffer();
    logDisplay_->buffer(logBuffer_);
    
    group->end();
    return group;
}

Fl_Group* FltkMainWindow::createButtonGroup(int x, int y, int w, int h) {
    Fl_Group* group = new Fl_Group(x, y, w, h);
    group->box(FL_ENGRAVED_BOX); // Add border like GroupBox
    
    // Group Title
    Fl_Box* title = new Fl_Box(x + 10, y - 10, 50, 20, "\xE6\x93\x8D\xE4\xBD\x9C"); // Operation
    title->box(FL_FLAT_BOX); // Cover the border
    title->align(FL_ALIGN_CENTER);
    
    int btnH = 40;
    int btnW = 150;
    int startX = x + w / 2 - btnW - 10;
    int btnY = y + (h - btnH) / 2;
    
    Fl_Button* previewBtn = new Fl_Button(startX, btnY, btnW, btnH, "\xE9\xA2\x84\xE8\xA7\x88\xE6\x95\xB0\xE6\x8D\xAE"); // Preview Data
    previewBtn->callback(cb_preview, this);
    
    processButton_ = new Fl_Button(startX + btnW + 20, btnY, btnW, btnH, "\xE5\xBC\x80\xE5\xA7\x8B\xE5\xA4\x84\xE7\x90\x86"); // Start Process
    processButton_->callback(cb_process, this);
    processButton_->color(FL_GREEN);
    processButton_->labelcolor(FL_WHITE);
    processButton_->labelfont(FL_BOLD);
    
    group->end();
    return group;
}

void FltkMainWindow::logMessage(const std::string& msg) {
    if (logBuffer_) {
        std::string fullMsg = getTimestamp() + msg + "\n";
        logBuffer_->append(fullMsg.c_str());
        logDisplay_->scroll(logBuffer_->length(), 0);
    }
}

void FltkMainWindow::updateRuleTable() {
    if (ruleTable_ && processor_) {
        ruleTable_->setRules(processor_->getRules());
    }
}

void FltkMainWindow::updateTaskTable() {
    if (taskTable_ && processor_) {
        taskTable_->setTasks(processor_->getTasks());
    }
}

// Callbacks
void FltkMainWindow::cb_addRule(Fl_Widget* w, void* v) {
    ((FltkMainWindow*)v)->addRule();
}

void FltkMainWindow::cb_editRule(Fl_Widget* w, void* v) {
    ((FltkMainWindow*)v)->editRule();
}

void FltkMainWindow::cb_deleteRule(Fl_Widget* w, void* v) {
    ((FltkMainWindow*)v)->deleteRule();
}

void FltkMainWindow::cb_addTask(Fl_Widget* w, void* v) {
    ((FltkMainWindow*)v)->addTask();
}

void FltkMainWindow::cb_editTask(Fl_Widget* w, void* v) {
    ((FltkMainWindow*)v)->editTask();
}

void FltkMainWindow::cb_deleteTask(Fl_Widget* w, void* v) {
    ((FltkMainWindow*)v)->deleteTask();
}

void FltkMainWindow::cb_loadConfig(Fl_Widget* w, void* v) {
    ((FltkMainWindow*)v)->loadConfig();
}

void FltkMainWindow::cb_saveConfig(Fl_Widget* w, void* v) {
    ((FltkMainWindow*)v)->saveConfig();
}

void FltkMainWindow::cb_preview(Fl_Widget* w, void* v) {
    ((FltkMainWindow*)v)->preview();
}

void FltkMainWindow::cb_process(Fl_Widget* w, void* v) {
    ((FltkMainWindow*)v)->process();
}

// Implementations
void FltkMainWindow::addRule() {
    Rule rule;
    rule.id = processor_->getRules().size() + 1; // Simple ID generation
    rule.name = "\xE6\x96\xB0\xE8\xA7\x84\xE5\x88\x99 " + std::to_string(rule.id); // New Rule
    rule.type = RuleType::FILTER;
    rule.enabled = true;
    
    RuleCondition condition;
    condition.column = 1;
    condition.oper = Operator::EQUAL;
    condition.value = "";
    rule.conditions.push_back(condition);
    
    EditRuleDialog dlg(400, 500, "\xE6\xB7\xBB\xE5\x8A\xA0\xE8\xA7\x84\xE5\x88\x99"); // Add Rule
    dlg.setRule(rule);
    
    if (dlg.showModal() == 1) {
        Rule newRule = dlg.getRule();
        processor_->addRule(newRule);
        updateRuleTable();
        logMessage("\xE8\xA7\x84\xE5\x88\x99\xE6\xB7\xBB\xE5\x8A\xA0\xE6\x88\x90\xE5\x8A\x9F: " + newRule.name);
    }
}

void FltkMainWindow::editRule() {
    int ruleId = ruleTable_->getSelectedRuleId();
    if (ruleId != -1) {
        Rule rule = processor_->getRule(ruleId);
        
        EditRuleDialog dlg(400, 500, "\xE7\xBC\x96\xE8\xBE\x91\xE8\xA7\x84\xE5\x88\x99"); // Edit Rule
        dlg.setRule(rule);
        
        if (dlg.showModal() == 1) {
            Rule updatedRule = dlg.getRule();
            processor_->updateRule(updatedRule);
            updateRuleTable();
            logMessage("\xE8\xA7\x84\xE5\x88\x99\xE6\x9B\xB4\xE6\x96\xB0\xE6\x88\x90\xE5\x8A\x9F: " + updatedRule.name);
        }
    } else {
        fl_alert("\xE8\xAF\xB7\xE5\x85\x88\xE9\x80\x89\xE6\x8B\xA9\xE4\xB8\x80\xE6\x9D\xA1\xE8\xA7\x84\xE5\x88\x99");
    }
}

void FltkMainWindow::deleteRule() {
    int ruleId = ruleTable_->getSelectedRuleId();
    if (ruleId != -1) {
        if (fl_choice("\xE7\xA1\xAE\xE5\xAE\x9A\xE8\xA6\x81\xE5\x88\xA0\xE9\x99\xA4\xE8\xA7\x84\xE5\x88\x99 %d \xE5\x90\x97?", "\xE5\x8F\x96\xE6\xB6\x88", "\xE7\xA1\xAE\xE5\xAE\x9A", 0, ruleId) == 1) {
            processor_->removeRule(ruleId);
            updateRuleTable();
            logMessage("\xE8\xA7\x84\xE5\x88\x99\xE5\x88\xA0\xE9\x99\xA4\xE6\x88\x90\xE5\x8A\x9F");
        }
    } else {
        fl_alert("\xE8\xAF\xB7\xE5\x85\x88\xE9\x80\x89\xE6\x8B\xA9\xE4\xB8\x80\xE6\x9D\xA1\xE8\xA7\x84\xE5\x88\x99");
    }
}

void FltkMainWindow::addTask() {
    ProcessingTask task;
    task.id = processor_->getTasks().size() + 1; // Simple ID generation
    task.taskName = "\xE6\x96\xB0\xE4\xBB\xBB\xE5\x8A\xA1 " + std::to_string(task.id); // New Task
    task.outputMode = OutputMode::NEW_WORKBOOK;
    task.enabled = true;
    
    EditTaskDialog dlg(450, 500, "\xE6\xB7\xBB\xE5\x8A\xA0\xE4\xBB\xBB\xE5\x8A\xA1"); // Add Task
    dlg.setAvailableRules(processor_->getRules());
    dlg.setTask(task);
    
    if (dlg.showModal() == 1) {
        ProcessingTask newTask = dlg.getTask();
        processor_->addTask(newTask);
        updateTaskTable();
        logMessage("\xE4\xBB\xBB\xE5\x8A\xA1\xE6\xB7\xBB\xE5\x8A\xA0\xE6\x88\x90\xE5\x8A\x9F: " + newTask.taskName);
    }
}

void FltkMainWindow::editTask() {
    int taskId = taskTable_->getSelectedTaskId();
    if (taskId != -1) {
        ProcessingTask task = processor_->getTask(taskId);
        
        EditTaskDialog dlg(450, 500, "\xE7\xBC\x96\xE8\xBE\x91\xE4\xBB\xBB\xE5\x8A\xA1"); // Edit Task
        dlg.setAvailableRules(processor_->getRules());
        dlg.setTask(task);
        
        if (dlg.showModal() == 1) {
            ProcessingTask updatedTask = dlg.getTask();
            processor_->updateTask(updatedTask);
            updateTaskTable();
            logMessage("\xE4\xBB\xBB\xE5\x8A\xA1\xE6\x9B\xB4\xE6\x96\xB0\xE6\x88\x90\xE5\x8A\x9F: " + updatedTask.taskName);
        }
    } else {
        fl_alert("\xE8\xAF\xB7\xE5\x85\x88\xE9\x80\x89\xE6\x8B\xA9\xE4\xB8\x80\xE4\xB8\xAA\xE4\xBB\xBB\xE5\x8A\xA1"); // Please select a task first
    }
}

void FltkMainWindow::deleteTask() {
    int taskId = taskTable_->getSelectedTaskId();
    if (taskId != -1) {
        if (fl_choice("\xE7\xA1\xAE\xE5\xAE\x9A\xE8\xA6\x81\xE5\x88\xA0\xE9\x99\xA4\xE4\xBB\xBB\xE5\x8A\xA1 %d \xE5\x90\x97?", "\xE5\x8F\x96\xE6\xB6\x88", "\xE7\xA1\xAE\xE5\xAE\x9A", 0, taskId) == 1) {
            processor_->removeTask(taskId);
            updateTaskTable();
            logMessage("\xE4\xBB\xBB\xE5\x8A\xA1\xE5\x88\xA0\xE9\x99\xA4\xE6\x88\x90\xE5\x8A\x9F");
        }
    } else {
        fl_alert("\xE8\xAF\xB7\xE5\x85\x88\xE9\x80\x89\xE6\x8B\xA9\xE4\xB8\x80\xE4\xB8\xAA\xE4\xBB\xBB\xE5\x8A\xA1");
    }
}

void FltkMainWindow::cb_runTasks(Fl_Widget* w, void* v) {
    ((FltkMainWindow*)v)->runTasks();
}

void FltkMainWindow::runTasks() {
    if (processor_->getTasks().empty()) {
        fl_alert("\xE6\xB2\xA1\xE6\x9C\x89\xE5\x8F\xAF\xE8\xBF\x90\xE8\xA1\x8C\xE7\x9A\x84\xE4\xBB\xBB\xE5\x8A\xA1"); // No tasks to run
        return;
    }

    Fl_Native_File_Chooser chooser;
    chooser.title("\xE9\x80\x89\xE6\x8B\xA9\xE8\xBE\x93\xE5\x85\xA5\xE7\x9B\xAE\xE5\xBD\x95"); // Select Input Directory
    chooser.type(Fl_Native_File_Chooser::BROWSE_DIRECTORY);
    
    if (chooser.show() == 0) {
        std::string inputPath = chooser.filename();
        
        fl_cursor(FL_CURSOR_WAIT);
        
        // Use a simple progress dialog or just block for now
        // TODO: Move to thread for better UI responsiveness
        
        std::vector<ProcessingResult> results = processor_->processTasks(inputPath);
        
        fl_cursor(FL_CURSOR_DEFAULT);
        
        int successCount = 0;
        int failCount = 0;
        std::string failures = "";
        
        for (const auto& res : results) {
            if (res.errors.empty()) successCount++;
            else {
                failCount++;
                for (const auto& err : res.errors) {
                    failures += err + "\n";
                }
            }
        }
        
        std::string msg = "\xE5\xA4\x84\xE7\x90\x86\xE5\xAE\x8C\xE6\x88\x90!\n\n"; // Processing Complete!
        msg += "\xE6\x88\x90\xE5\x8A\x9F: " + std::to_string(successCount) + "\n";
        msg += "\xE5\xA4\xB1\xE8\xB4\xA5: " + std::to_string(failCount);
        
        if (failCount > 0) {
            msg += "\n\n\xE9\x94\x99\xE8\xAF\xAF\xE4\xBF\xA1\xE6\x81\xAF:\n" + failures;
        }
        
        if (msg.length() > 500) {
             msg = msg.substr(0, 500) + "... (\xE6\x9B\xB4\xE5\xA4\x9A\xE9\x94\x99\xE8\xAF\xAF\xE8\xAF\xB7\xE6\x9F\xA5\xE7\x9C\x8B\xE6\x97\xA5\xE5\xBF\x97)";
        }

        fl_message(msg.c_str());
    }
}

void FltkMainWindow::loadConfig() {
    Fl_Native_File_Chooser fnfc;
    fnfc.title("\xE5\x8A\xA0\xE8\xBD\xBD\xE9\x85\x8D\xE7\xBD\xAE"); // Load Config
    fnfc.type(Fl_Native_File_Chooser::BROWSE_FILE);
    fnfc.filter("Config Files\t*.cfg\nAll Files\t*.*");
    
    if (fnfc.show() == 0) {
        std::string filename = fnfc.filename();
        if (processor_->loadConfiguration(filename)) {
            currentConfigFile_ = filename;
            updateRuleTable();
            updateTaskTable();
            configNameLabel_->label(filename.c_str()); // Simplified to show full path or filename
            logMessage("\xE9\x85\x8D\xE7\xBD\xAE\xE5\x8A\xA0\xE8\xBD\xBD\xE6\x88\x90\xE5\x8A\x9F: " + filename);
        } else {
            fl_alert("\xE9\x85\x8D\xE7\xBD\xAE\xE6\x96\x87\xE4\xBB\xB6\xE5\x8A\xA0\xE8\xBD\xBD\xE5\xA4\xB1\xE8\xB4\xA5");
        }
    }
}

void FltkMainWindow::saveConfig() {
    Fl_Native_File_Chooser fnfc;
    fnfc.title("\xE4\xBF\x9D\xE5\xAD\x98\xE9\x85\x8D\xE7\xBD\xAE"); // Save Config
    fnfc.type(Fl_Native_File_Chooser::BROWSE_SAVE_FILE); // FIXED
    fnfc.filter("Config Files\t*.cfg\nAll Files\t*.*");
    
    if (fnfc.show() == 0) {
        std::string filename = fnfc.filename();
        // Add extension if missing
        if (filename.find(".cfg") == std::string::npos) {
            filename += ".cfg";
        }
        
        if (processor_->saveConfiguration(filename)) {
            currentConfigFile_ = filename;
            configNameLabel_->label(filename.c_str());
            logMessage("\xE9\x85\x8D\xE7\xBD\xAE\xE4\xBF\x9D\xE5\xAD\x98\xE6\x88\x90\xE5\x8A\x9F: " + filename);
        } else {
            fl_alert("\xE9\x85\x8D\xE7\xBD\xAE\xE6\x96\x87\xE4\xBB\xB6\xE4\xBF\x9D\xE5\xAD\x98\xE5\xA4\xB1\xE8\xB4\xA5");
        }
    }
}

void FltkMainWindow::preview() {
    Fl_Native_File_Chooser fnfc;
    fnfc.title("\xE9\x80\x89\xE6\x8B\xA9" "Excel\xE6\x96\x87\xE4\xBB\xB6"); // Select Excel File
    fnfc.type(Fl_Native_File_Chooser::BROWSE_FILE);
    fnfc.filter("Excel Files\t*.xlsx;*.xls\nAll Files\t*.*");
    
    if (fnfc.show() == 0) {
        std::string filename = fnfc.filename();
        currentInputFile_ = filename;
        logMessage("\xE6\xAD\xA3\xE5\x9C\xA8\xE5\x8A\xA0\xE8\xBD\xBD\xE6\x95\xB0\xE6\x8D\xAE: " + filename);
        
        // Show progress indicator (simple version)
        progressBar_->show();
        progressBar_->value(0);
        progressBar_->label("\xE5\x8A\xA0\xE8\xBD\xBD\xE4\xB8\xAD..."); // Loading...
        Fl::check(); // Process events to update UI
        
        // Load file in a separate thread to avoid freezing UI
        // But for simplicity and safety with FLTK, we'll just do it here for now
        // or use the processor's async capability if available.
        // processor_->loadFile is synchronous.
        
        if (processor_->loadFile(filename, "", 100)) { // Load up to 100 rows
            auto data = processor_->getPreviewData(100);
            previewTable_->setData(data);
            
            // Switch to preview tab
            tabs_->value(previewTable_->parent());
            
            logMessage("\xE6\x95\xB0\xE6\x8D\xAE\xE5\x8A\xA0\xE8\xBD\xBD\xE6\x88\x90\xE5\x8A\x9F, \xE5\x85\xB1 " + std::to_string(data.size()) + " \xE8\xA1\x8C");
        } else {
            fl_alert("\xE6\x95\xB0\xE6\x8D\xAE\xE5\x8A\xA0\xE8\xBD\xBD\xE5\xA4\xB1\xE8\xB4\xA5");
            logMessage("\xE6\x95\xB0\xE6\x8D\xAE\xE5\x8A\xA0\xE8\xBD\xBD\xE5\xA4\xB1\xE8\xB4\xA5: " + filename);
        }
        
        progressBar_->hide();
    }
}

void FltkMainWindow::process() {
    if (currentInputFile_.empty()) {
        fl_alert("\xE8\xAF\xB7\xE5\x85\x88\xE9\x80\x89\xE6\x8B\xA9" "Excel\xE6\x96\x87\xE4\xBB\xB6 (\xE8\xAF\xB7\xE4\xBD\xBF\xE7\x94\xA8\xE9\xA2\x84\xE8\xA7\x88\xE5\x8A\x9F\xE8\x83\xBD\xE9\x80\x89\xE6\x8B\xA9)");
        return;
    }

    Fl_Native_File_Chooser fnfc;
    fnfc.title("\xE4\xBF\x9D\xE5\xAD\x98\xE7\xBB\x93\xE6\x9E\x9C\xE6\x96\x87\xE4\xBB\xB6"); // Save Result File
    fnfc.type(Fl_Native_File_Chooser::BROWSE_SAVE_FILE);
    fnfc.filter("Excel Files\t*.xlsx\nAll Files\t*.*");
    
    if (fnfc.show() != 0) return;
    
    std::string outputFile = fnfc.filename();
    if (outputFile.find(".xlsx") == std::string::npos) {
        outputFile += ".xlsx";
    }

    logMessage("\xE5\xBC\x80\xE5\xA7\x8B\xE5\xA4\x84\xE7\x90\x86..."); // Start processing
    progressBar_->show();
    progressBar_->value(0);
    processButton_->deactivate();
    
    // Run in thread
    std::thread([this, outputFile]() {
        try {
            auto result = processor_->processExcelFile(currentInputFile_, outputFile);
            
            scheduleOnMainThread([this, result]() {
                progressBar_->value(100);
                progressBar_->hide();
                processButton_->activate();
                
                std::string msg = "\xE5\xA4\x84\xE7\x90\x86\xE5\xAE\x8C\xE6\x88\x90!\n"; // Processing Complete!
                msg += "\xE6\x80\xBB\xE8\xA1\x8C\xE6\x95\xB0: " + std::to_string(result.totalRows) + "\n";
                msg += "\xE5\x8C\xB9\xE9\x85\x8D\xE8\xA1\x8C\xE6\x95\xB0: " + std::to_string(result.matchedRows);
                
                logMessage(msg);
                fl_message("%s", msg.c_str());
            });
        } catch (const std::exception& e) {
            std::string errMsg = e.what();
            scheduleOnMainThread([this, errMsg]() {
                progressBar_->hide();
                processButton_->activate();
                logMessage("\xE5\xA4\x84\xE7\x90\x86\xE5\xA4\xB1\xE8\xB4\xA5: " + errMsg);
                fl_alert("\xE5\xA4\x84\xE7\x90\x86\xE5\xA4\xB1\xE8\xB4\xA5: %s", errMsg.c_str());
            });
        }
    }).detach();
}
