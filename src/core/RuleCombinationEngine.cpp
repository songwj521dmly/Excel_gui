#include "ExcelProcessorCore.h"
#include <algorithm>
#include <numeric>
#include <memory>
#include <future>
#include <thread>
#include <unordered_map>
#include <cmath>
#include <string>
#include <iostream>

// Rule Combination Engine
class RuleCombinationEngine {
public:
    struct CombinationStats {
        int combinationId;
        std::string name;
        int ruleCount;
        CombinationStrategy strategy;
        std::vector<RuleType> ruleTypes;
        int totalConditions;
        double estimatedTime;
    };

    explicit RuleCombinationEngine(std::shared_ptr<RuleEngine> ruleEngine)
        : ruleEngine_(ruleEngine) {}

    // Create combination
    std::vector<int> createCombination(int combinationId,
                                        const std::vector<int>& ruleIds,
                                        CombinationStrategy strategy,
                                        const std::vector<Rule>& rules) {
        CombinationInfo combo;
        combo.id = combinationId;
        combo.strategy = strategy;
        combo.ruleIds = ruleIds;
        combo.name = "Combination " + std::to_string(combinationId);

        // Verify rules
        for (int ruleId : ruleIds) {
            auto it = std::find_if(rules.begin(), rules.end(),
                [ruleId](const Rule& rule) { return rule.id == ruleId; });

            if (it == rules.end()) {
                throw std::runtime_error("Rule ID " + std::to_string(ruleId) + " does not exist");
            }

            if (!it->enabled) {
                throw std::runtime_error("Rule " + it->name + " is not enabled");
            }
        }

        combos_[combinationId] = combo;
        return ruleIds;
    }

    // Evaluate combination
    bool evaluateCombination(int combinationId, const DataRow& row,
                                const std::vector<Rule>& rules) const {
        auto it = combos_.find(combinationId);
        if (it == combos_.end()) {
            throw std::runtime_error("Combination ID " + std::to_string(combinationId) + " does not exist");
        }

        const auto& combo = it->second;
        return evaluateRules(combo.ruleIds, combo.strategy, row, rules);
    }

    // Optimize rule order
    std::vector<int> optimizeRuleOrder(const std::vector<int>& ruleIds,
                                       const std::vector<Rule>& rules,
                                       const std::vector<DataRow>& sampleData) {
        if (ruleIds.size() <= 1) {
            return ruleIds;
        }

        // Calculate cost
        std::unordered_map<int, double> ruleCosts;
        for (int ruleId : ruleIds) {
            ruleCosts[ruleId] = calculateRuleCost(ruleId, rules);
        }

        // Calculate selectivity
        std::unordered_map<int, double> ruleSelectivity;
        for (int ruleId : ruleIds) {
            ruleSelectivity[ruleId] = calculateRuleSelectivity(ruleId, rules, sampleData);
        }

        // Optimize
        std::vector<int> optimizedRules = ruleIds;
        std::sort(optimizedRules.begin(), optimizedRules.end(),
            [&](int a, int b) {
                double scoreA = ruleSelectivity[a] / ruleCosts[a];
                double scoreB = ruleSelectivity[b] / ruleCosts[b];
                return scoreA > scoreB;
            });

        return optimizedRules;
    }

    // Generate all combinations
    std::vector<std::vector<int>> generateAllCombinations(const std::vector<int>& ruleIds) {
        std::vector<std::vector<int>> result;
        for (int id : ruleIds) {
            result.push_back({id});
        }
        if (ruleIds.size() > 1) {
            result.push_back(ruleIds);
        }
        return result;
    }

    // Analyze dependencies
    std::vector<std::vector<int>> analyzeDependencies(const std::vector<Rule>& rules) {
        std::vector<std::vector<int>> dependencies;
        for (const auto& rule : rules) {
             std::vector<int> ruleDeps;
             dependencies.push_back(ruleDeps);
        }
        return dependencies;
    }

    // Predict execution time
    double predictExecutionTime(int combinationId,
                                 const std::vector<Rule>& rules,
                                 int estimatedDataSize) {
        auto it = combos_.find(combinationId);
        if (it == combos_.end()) {
            return 0.0;
        }

        const auto& combo = it->second;
        double totalTime = 0.0;

        for (int ruleId : combo.ruleIds) {
            auto ruleIt = std::find_if(rules.begin(), rules.end(),
                [ruleId](const Rule& rule) { return rule.id == ruleId; });

            if (ruleIt != rules.end()) {
                totalTime += predictRuleTime(*ruleIt, estimatedDataSize);
            }
        }

        return totalTime;
    }

    // Get stats
    CombinationStats getCombinationStats(int combinationId,
                                      const std::vector<Rule>& rules) const {
        CombinationStats stats;
        stats.combinationId = 0;
        stats.ruleCount = 0;
        stats.totalConditions = 0;
        stats.estimatedTime = 0.0;
        
        auto it = combos_.find(combinationId);
        if (it == combos_.end()) {
            return stats;
        }

        const auto& combo = it->second;
        stats.combinationId = combinationId;
        stats.ruleCount = (int)combo.ruleIds.size();
        stats.strategy = combo.strategy;
        stats.name = combo.name;

        for (int ruleId : combo.ruleIds) {
            auto ruleIt = std::find_if(rules.begin(), rules.end(),
                [ruleId](const Rule& rule) { return rule.id == ruleId; });

            if (ruleIt != rules.end()) {
                stats.ruleTypes.push_back(ruleIt->type);
                stats.totalConditions += (int)ruleIt->conditions.size();
            }
        }

        return stats;
    }

private:
    struct CombinationInfo {
        int id;
        std::string name;
        std::vector<int> ruleIds;
        CombinationStrategy strategy;
    };

    std::shared_ptr<RuleEngine> ruleEngine_;
    std::unordered_map<int, CombinationInfo> combos_;

    bool evaluateRules(const std::vector<int>& ruleIds,
                        CombinationStrategy strategy,
                        const DataRow& row,
                        const std::vector<Rule>& rules) const {
        if (ruleIds.empty()) {
            return false;
        }

        switch (strategy) {
            case CombinationStrategy::AND: {
                for (int ruleId : ruleIds) {
                    auto it = std::find_if(rules.begin(), rules.end(),
                        [ruleId](const Rule& rule) { return rule.id == ruleId; });

                    if (it != rules.end()) {
                        if (!ruleEngine_->evaluateRule(*it, row)) {
                            return false;
                        }
                    }
                }
                return true;
            }

            case CombinationStrategy::OR: {
                for (int ruleId : ruleIds) {
                    auto it = std::find_if(rules.begin(), rules.end(),
                        [ruleId](const Rule& rule) { return rule.id == ruleId; });

                    if (it != rules.end()) {
                        if (ruleEngine_->evaluateRule(*it, row)) {
                            return true;
                        }
                    }
                }
                return false;
            }

            default:
                return false;
        }
    }

    double calculateRuleCost(int ruleId, const std::vector<Rule>& rules) const {
        auto it = std::find_if(rules.begin(), rules.end(),
            [ruleId](const Rule& rule) { return rule.id == ruleId; });

        if (it == rules.end()) {
            return 1.0; 
        }

        double cost = 1.0;

        switch (it->type) {
            case RuleType::FILTER:
                cost = 1.0;
                break;
            case RuleType::TRANSFORM:
                cost = 2.0;
                break;
            case RuleType::DELETE_ROW:
                cost = 1.5;
                break;
            case RuleType::SPLIT:
                cost = 3.0;
                break;
        }

        cost += it->conditions.size() * 0.5;

        for (const auto& condition : it->conditions) {
            switch (condition.oper) {
                case Operator::REGEX:
                    cost += 2.0;
                    break;
                case Operator::CONTAINS:
                case Operator::NOT_CONTAINS:
                    cost += 0.5;
                    break;
                case Operator::STARTS_WITH:
                case Operator::ENDS_WITH:
                    cost += 0.3;
                    break;
                default:
                    cost += 0.1;
                    break;
            }
        }

        return cost;
    }

    double calculateRuleSelectivity(int ruleId, const std::vector<Rule>& rules,
                                   const std::vector<DataRow>& sampleData) const {
        if (sampleData.empty()) {
            return 0.5; 
        }

        auto it = std::find_if(rules.begin(), rules.end(),
            [ruleId](const Rule& rule) { return rule.id == ruleId; });

        if (it == rules.end()) {
            return 0.5;
        }

        int matchedCount = 0;
        for (const auto& row : sampleData) {
            if (ruleEngine_->evaluateRule(*it, row)) {
                matchedCount++;
            }
        }

        double selectivity = static_cast<double>(matchedCount) / sampleData.size();

        return 1.0 - std::abs(selectivity - 0.5);
    }

    double evaluateCombinationScore(const std::vector<int>& combination,
                                   const std::vector<Rule>& rules,
                                   const std::vector<DataRow>& sampleData) {
        if (sampleData.empty()) {
            return 0.0;
        }

        int matchedCount = 0;
        int totalConditions = 0;

        for (const auto& row : sampleData) {
            bool matched = true;

            for (int ruleId : combination) {
                auto it = std::find_if(rules.begin(), rules.end(),
                    [ruleId](const Rule& rule) { return rule.id == ruleId; });

                if (it != rules.end()) {
                    totalConditions += (int)it->conditions.size();
                    if (!ruleEngine_->evaluateRule(*it, row)) {
                        matched = false;
                        break;
                    }
                }
            }

            if (matched) {
                matchedCount++;
            }
        }

        double accuracy = static_cast<double>(matchedCount) / sampleData.size();
        double selectivityBalance = 1.0 - (static_cast<double>(totalConditions) / (combination.size() * 10.0));

        return accuracy * 0.7 + selectivityBalance * 0.3;
    }

    void generateCombinationsRecursive(const std::vector<int>& availableRules,
                                     int startIndex,
                                     int remainingSize,
                                     std::vector<int> current,
                                     std::vector<std::vector<int>>& result) {
        if (remainingSize == 0) {
            if (!current.empty()) {
                result.push_back(current);
            }
            return;
        }

        for (int i = startIndex; i <= (int)availableRules.size() - remainingSize; ++i) {
            current.push_back(availableRules[i]);
            generateCombinationsRecursive(availableRules, i + 1, remainingSize - 1,
                                       current, result);
            current.pop_back();
        }
    }

    bool requiresPreprocessing(const RuleCondition& condition) {
        switch (condition.oper) {
            case Operator::REGEX:
                return true;
            case Operator::STARTS_WITH:
            case Operator::ENDS_WITH:
                return true;
            default:
                return false;
        }
    }

    std::vector<int> findDependentRules(const RuleCondition& condition,
                                        const std::vector<Rule>& rules) {
        return {};
    }

    double predictRuleTime(const Rule& rule, int dataSize) {
        double baseTime = 0.001; 

        switch (rule.type) {
            case RuleType::FILTER:
                baseTime = 0.002;
                break;
            case RuleType::TRANSFORM:
                baseTime = 0.005;
                break;
            case RuleType::DELETE_ROW:
                baseTime = 0.003;
                break;
            case RuleType::SPLIT:
                baseTime = 0.008;
                break;
        }

        double complexityFactor = 1.0 + (rule.conditions.size() * 0.2);
        double dataFactor = std::log10(std::max(1.0, static_cast<double>(dataSize))) / 10.0;

        return baseTime * complexityFactor * dataFactor;
    }
};
