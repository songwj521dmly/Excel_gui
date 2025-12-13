#include "ExcelProcessorCore.h"
#include <iostream>
#include <fstream>
#include <cassert>
#include <cstdio>
#include <string>

void testConfigPersistence() {
    // Use a random-ish filename
    std::string testConfigFile = "test_config_persistence_" + std::to_string(std::time(nullptr)) + ".csv";
    
    std::cout << "Using config file: " << testConfigFile << std::endl;
    
    std::cout << "Creating processor..." << std::endl;
    ExcelProcessorCore processor;
    
    Rule ruleA;
    ruleA.id = 1;
    ruleA.name = "RuleA";
    ruleA.type = RuleType::FILTER;
    // ruleA.excludedRuleIds = {2, 3}; // Removed
    ruleA.enabled = true;
    // Add dummy condition for validation
    RuleCondition condA;
    condA.column = 1;
    condA.oper = Operator::NOT_EMPTY;
    ruleA.conditions.push_back(condA);
    
    Rule ruleB;
    ruleB.id = 2;
    ruleB.name = "RuleB";
    // Add dummy condition for validation
    RuleCondition condB;
    condB.column = 1;
    condB.oper = Operator::NOT_EMPTY;
    ruleB.conditions.push_back(condB);
    
    Rule ruleC;
    ruleC.id = 3;
    ruleC.name = "RuleC";
    // Add dummy condition for validation
    RuleCondition condC;
    condC.column = 1;
    condC.oper = Operator::NOT_EMPTY;
    ruleC.conditions.push_back(condC);
    
    processor.addRule(ruleA);
    processor.addRule(ruleB);
    processor.addRule(ruleC);
    
    std::cout << "Saving configuration..." << std::endl;
    if (!processor.saveConfiguration(testConfigFile)) {
        std::cerr << "Failed to save configuration" << std::endl;
        exit(1);
    }
    std::cout << "Configuration saved." << std::endl;
    
    std::cout << "Loading configuration..." << std::endl;
    ExcelProcessorCore processor2;
    if (!processor2.loadConfiguration(testConfigFile)) {
        std::cerr << "Failed to load configuration" << std::endl;
        auto errors = processor2.getErrors();
        for(const auto& err : errors) std::cerr << "Error: " << err << std::endl;
        exit(1);
    }
    std::cout << "Configuration loaded." << std::endl;
    
    Rule loadedRuleA = processor2.getRule(1);
    
    std::cout << "Loaded Rule A Name: " << loadedRuleA.name << std::endl;
    // std::cout << "Loaded Rule A Excluded IDs: ";
    // for (int id : loadedRuleA.excludedRuleIds) {
    //     std::cout << id << " ";
    // }
    std::cout << std::endl;
    
    if (loadedRuleA.name != "RuleA") {
        std::cerr << "FAIL: Rule name mismatch" << std::endl;
        exit(1);
    }
    /*
    if (loadedRuleA.excludedRuleIds.size() != 2) {
        std::cerr << "FAIL: Excluded rules size mismatch" << std::endl;
        exit(1);
    }
    if (loadedRuleA.excludedRuleIds[0] != 2 || loadedRuleA.excludedRuleIds[1] != 3) {
        std::cerr << "FAIL: Excluded rule IDs mismatch" << std::endl;
        exit(1);
    }
    */
    
    // Cleanup
    std::cout << "Cleaning up..." << std::endl;
    // std::remove(testConfigFile.c_str());
    
    std::cout << "Config persistence test passed!" << std::endl;
}

int main() {
    std::cout << "MAIN STARTING" << std::endl;
    try {
        testConfigPersistence();
    } catch (const std::exception& e) {
        std::cerr << "Test failed with exception: " << e.what() << std::endl;
        return 1;
    } catch (...) {
        std::cerr << "Test failed with unknown exception" << std::endl;
        return 1;
    }
    return 0;
}
