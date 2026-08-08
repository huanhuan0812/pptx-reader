#pragma once

#include "pptx/TextRun.hpp"

#include <string>
#include <vector>

namespace pptx {

// 文本框中的一个段落（对应 OOXML 中的 <a:p>）
class Paragraph {
public:
    Paragraph() = default;

    // 缩进层级（<a:pPr lvl="..."/>，从 0 开始）
    int level() const noexcept { return level_; }
    void setLevel(int l) noexcept { level_ = l; }

    // 对齐方式（l / ctr / r / just 等原始值）
    const std::string& alignment() const noexcept { return alignment_; }
    void setAlignment(const std::string& a) { alignment_ = a; }

    std::vector<TextRun>& runs() noexcept { return runs_; }
    const std::vector<TextRun>& runs() const noexcept { return runs_; }

    // 段落全文：按顺序拼接各 run 的文本；行内换行 <a:br/> 会变成 '\n'
    std::string text() const;

    bool empty() const noexcept { return runs_.empty(); }

    // 供范围 for 循环遍历本段内的 TextRun
    auto begin() noexcept { return runs_.begin(); }
    auto end() noexcept { return runs_.end(); }
    auto begin() const noexcept { return runs_.begin(); }
    auto end() const noexcept { return runs_.end(); }

private:
    int level_ = 0;
    std::string alignment_;
    std::vector<TextRun> runs_;
};

}  // namespace pptx
