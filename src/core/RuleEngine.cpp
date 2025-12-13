#include "ExcelProcessorCore.h"
#include <algorithm>
#include <regex>
#include <sstream>
#include <iomanip>
#include <cmath>
#include <cstring>
#include <ctime>

// Rule engine implementation class
class DefaultRuleEngine : public RuleEngine {
public:
    bool evaluateRule(const Rule& rule, const DataRow& row, const std::vector<Rule>* allRules = nullptr) const override {
        if (!rule.enabled) return false;

        // Check exclusions (MOVED TO TASK LEVEL)
        /*
        if (!rule.excludedRuleIds.empty() && allRules) {
            for (int excludedId : rule.excludedRuleIds) {
                auto it = std::find_if(allRules->begin(), allRules->end(),
                                     [excludedId](const Rule& r) { return r.id == excludedId; });
                if (it != allRules->end()) {
                    // Check if the excluded rule matches. 
                    // Pass nullptr for allRules to prevent infinite recursion and cycles.
                    // This means we only support 1 level of exclusion depth (Rule A excludes Rule B, but Rule B's exclusions are ignored).
                    if (evaluateRule(*it, row, nullptr)) {
                        return false; 
                    }
                }
            }
        }
        */

        // Process based on rule type
        switch (rule.type) {
            case RuleType::FILTER:
                return evaluateFilterRule(rule, row);
            case RuleType::DELETE_ROW:
                return evaluateDeleteRule(rule, row);
            case RuleType::SPLIT:
                return evaluateSplitRule(rule, row);
            case RuleType::TRANSFORM:
                return evaluateTransformRule(rule, row);
            default:
                return false;
        }
    }

    bool evaluateCondition(const RuleCondition& condition,
                           const std::variant<std::string, int, double, bool, std::tm>& value) const override {
        // Helper to check if operator is string-only
        auto isStringOperator = [](Operator op) {
            return op == Operator::CONTAINS || 
                   op == Operator::NOT_CONTAINS || 
                   op == Operator::STARTS_WITH || 
                   op == Operator::ENDS_WITH || 
                   op == Operator::REGEX;
        };

        try {
            // Simple type-safe comparison
            if (auto strVal = std::get_if<std::string>(&value)) {
                return evaluateStringCondition(condition, *strVal);
            } else if (auto intVal = std::get_if<int>(&value)) {
                if (isStringOperator(condition.oper)) {
                    return evaluateStringCondition(condition, std::to_string(*intVal));
                }
                return evaluateNumberCondition(condition, static_cast<double>(*intVal));
            } else if (auto doubleVal = std::get_if<double>(&value)) {
                if (isStringOperator(condition.oper)) {
                    // Use to_string to keep decimals (e.g. 5.000000) so "contains 5.0" works
                    return evaluateStringCondition(condition, std::to_string(*doubleVal));
                }
                return evaluateNumberCondition(condition, *doubleVal);
            } else if (auto boolVal = std::get_if<bool>(&value)) {
                if (isStringOperator(condition.oper)) {
                    return evaluateStringCondition(condition, *boolVal ? "true" : "false");
                }
                return evaluateBooleanCondition(condition, *boolVal);
            } else if (auto tmVal = std::get_if<std::tm>(&value)) {
                // Convert tm to string for comparison
                std::ostringstream oss;
                oss << std::put_time(tmVal, "%Y-%m-%d");
                return evaluateStringCondition(condition, oss.str());
            }
            return false;
        } catch (const std::exception& e) {
            return false; // Condition evaluation failed
        }
    }

    std::vector<std::string> getValidationErrors(const Rule& rule) const override {
        std::vector<std::string> errors;

        // Validate rule ID
        if (rule.id <= 0) {
            errors.push_back("Rule ID must be greater than 0");
        }

        // Validate rule name
        if (rule.name.empty()) {
            errors.push_back("Rule name cannot be empty");
        }

        // Validate conditions
        if (rule.conditions.empty()) {
            errors.push_back("Rule must contain at least one condition");
        }

        for (const auto& condition : rule.conditions) {
            // Validate column number
            if (condition.column <= 0) {
                errors.push_back("Column number must be greater than 0");
            }
        }

        // Validate split rule requirements
        if (rule.type == RuleType::SPLIT && rule.targetSheet.empty()) {
            errors.push_back("Split rule must specify target sheet");
        }

        return errors;
    }

private:
    bool evaluateConditions(const Rule& rule, const DataRow& row) const {
        if (rule.conditions.empty()) return true;

        if (rule.logic == RuleLogic::AND) {
            for (const auto& condition : rule.conditions) {
                if (static_cast<size_t>(condition.column) > row.data.size()) {
                    return false; // Column out of range
                }

                if (!evaluateSingleCondition(condition, row.data[condition.column - 1])) {
                    return false;
                }
            }
            return true;
        } else { // OR
            for (const auto& condition : rule.conditions) {
                if (static_cast<size_t>(condition.column) <= row.data.size()) {
                    if (evaluateSingleCondition(condition, row.data[condition.column - 1])) {
                        return true;
                    }
                }
            }
            return false;
        }
    }

    bool evaluateFilterRule(const Rule& rule, const DataRow& row) const {
        return evaluateConditions(rule, row);
    }

    bool evaluateDeleteRule(const Rule& rule, const DataRow& row) const {
        return evaluateConditions(rule, row);
    }

    bool evaluateSplitRule(const Rule& rule, const DataRow& row) const {
        return evaluateConditions(rule, row);
    }

    bool evaluateTransformRule(const Rule& rule, const DataRow& row) const {
        return evaluateConditions(rule, row);
    }

    bool evaluateSingleCondition(const RuleCondition& condition, const std::variant<std::string, int, double, bool, std::tm>& value) const {
        // Delegate to type-specific evaluation
        return evaluateCondition(condition, value);
    }

    bool evaluateStringCondition(const RuleCondition& condition, const std::string& value) const {
        auto strConditionValue = std::get_if<std::string>(&condition.value);
        if (!strConditionValue) return false;

        // Trim whitespace to ensure accurate comparison
        auto trim = [](const std::string& s) {
            size_t first = s.find_first_not_of(" \t\r\n");
            if (std::string::npos == first) return std::string();
            size_t last = s.find_last_not_of(" \t\r\n");
            return s.substr(first, (last - first + 1));
        };

        std::string leftValue = trim(value);
        std::string rightValue = trim(*strConditionValue);

        if (!condition.case_sensitive) {
            std::transform(leftValue.begin(), leftValue.end(), leftValue.begin(), ::tolower);
            std::transform(rightValue.begin(), rightValue.end(), rightValue.begin(), ::tolower);
        }

        switch (condition.oper) {
            case Operator::EQUAL:
                if (leftValue == rightValue) return true;
                // Strict numeric check: only if BOTH strings are valid numbers
                try {
                    size_t idx1 = 0, idx2 = 0;
                    double leftNum = std::stod(leftValue, &idx1);
                    double rightNum = std::stod(rightValue, &idx2);
                    // Ensure the entire string was parsed as a number
                    if (idx1 == leftValue.size() && idx2 == rightValue.size()) {
                        return std::fabs(leftNum - rightNum) < 1e-10;
                    }
                } catch (...) {}
                return false;
            case Operator::NOT_EQUAL:
                if (leftValue != rightValue) {
                     // Check strict numeric equality to avoid false negatives
                     try {
                        size_t idx1 = 0, idx2 = 0;
                        double leftNum = std::stod(leftValue, &idx1);
                        double rightNum = std::stod(rightValue, &idx2);
                        if (idx1 == leftValue.size() && idx2 == rightValue.size()) {
                            return std::fabs(leftNum - rightNum) >= 1e-10;
                        }
                    } catch (...) {}
                    return true; 
                }
                return false; // Strings are equal
            case Operator::CONTAINS:
                return leftValue.find(rightValue) != std::string::npos;
            case Operator::NOT_CONTAINS:
                return leftValue.find(rightValue) == std::string::npos;
            case Operator::STARTS_WITH:
                return leftValue.rfind(rightValue, 0) == 0;
            case Operator::ENDS_WITH:
                return leftValue.size() >= rightValue.size() &&
                       leftValue.compare(leftValue.size() - rightValue.size(), rightValue.size(), rightValue) == 0;
            case Operator::GREATER:
            case Operator::LESS:
            case Operator::GREATER_EQUAL:
            case Operator::LESS_EQUAL:
                try {
                    size_t idx1 = 0, idx2 = 0;
                    double leftNum = std::stod(leftValue, &idx1);
                    double rightNum = std::stod(rightValue, &idx2);
                    // Strict numeric comparison only
                    if (idx1 == leftValue.size() && idx2 == rightValue.size()) {
                        if (condition.oper == Operator::GREATER) return leftNum > rightNum;
                        if (condition.oper == Operator::LESS) return leftNum < rightNum;
                        if (condition.oper == Operator::GREATER_EQUAL) return leftNum >= rightNum - 1e-10;
                        if (condition.oper == Operator::LESS_EQUAL) return leftNum <= rightNum + 1e-10;
                    }
                } catch (...) {
                    return false; // Not numbers
                }
                return false;
            case Operator::EMPTY:
                return leftValue.empty();
            case Operator::NOT_EMPTY:
                return !leftValue.empty();
            case Operator::REGEX:
                try {
                    std::regex pattern(rightValue);
                    // For regex, we use the original (trimmed?) value. 
                    // Usually regex should match against the trimmed value to be consistent, 
                    // but sometimes users want to match spaces. 
                    // Given the "EQUAL" issue was whitespace, let's use trimmed value for consistency 
                    // OR use original value? 
                    // Let's use leftValue (trimmed/lowercased if insensitive) for consistency with other operators.
                    // BUT regex usually expects case sensitivity handling via flags.
                    // Here we already lowercased if not case_sensitive.
                    return std::regex_match(leftValue, pattern);
                } catch (const std::regex_error&) {
                    return false;
                }
            default:
                return false;
        }
    }

    bool evaluateNumberCondition(const RuleCondition& condition, double value) const {
        double conditionValue = 0.0;

        // Extract condition value
        if (auto intVal = std::get_if<int>(&condition.value)) {
            conditionValue = static_cast<double>(*intVal);
        } else if (auto doubleVal = std::get_if<double>(&condition.value)) {
            conditionValue = *doubleVal;
        } else if (auto strVal = std::get_if<std::string>(&condition.value)) {
            try {
                conditionValue = std::stod(*strVal);
            } catch (...) {
                return false;
            }
        } else if (auto boolVal = std::get_if<bool>(&condition.value)) {
            conditionValue = *boolVal ? 1.0 : 0.0;
        } else {
            return false;
        }

        switch (condition.oper) {
            case Operator::EQUAL:
                return std::fabs(value - conditionValue) < 1e-10;
            case Operator::NOT_EQUAL:
                return std::fabs(value - conditionValue) >= 1e-10;
            case Operator::GREATER:
                return value > conditionValue;
            case Operator::LESS:
                return value < conditionValue;
            case Operator::GREATER_EQUAL:
                return value >= conditionValue - 1e-10;
            case Operator::LESS_EQUAL:
                return value <= conditionValue + 1e-10;
            case Operator::EMPTY:
                return false; // Numbers cannot be empty
            case Operator::NOT_EMPTY:
                return true;  // Numbers are always non-empty
            default:
                return false;
        }
    }

    bool evaluateBooleanCondition(const RuleCondition& condition, bool value) const {
        bool conditionValue = false;

        // Extract condition value
        if (auto boolVal = std::get_if<bool>(&condition.value)) {
            conditionValue = *boolVal;
        } else if (auto intVal = std::get_if<int>(&condition.value)) {
            conditionValue = *intVal != 0;
        } else if (auto doubleVal = std::get_if<double>(&condition.value)) {
            conditionValue = *doubleVal != 0.0;
        } else if (auto strVal = std::get_if<std::string>(&condition.value)) {
            std::string str = *strVal;
            std::transform(str.begin(), str.end(), str.begin(), ::tolower);
            conditionValue = (str == "true" || str == "1" || str == "yes" || str == "\xe6\x98\xaf");
        } else {
            return false;
        }

        switch (condition.oper) {
            case Operator::EQUAL:
                return value == conditionValue;
            case Operator::NOT_EQUAL:
                return value != conditionValue;
            case Operator::EMPTY:
                return false;
            case Operator::NOT_EMPTY:
                return true;
            default:
                return false;
        }
    }
};

// Factory function
std::unique_ptr<RuleEngine> createRuleEngine() {
    return std::make_unique<DefaultRuleEngine>();
}
