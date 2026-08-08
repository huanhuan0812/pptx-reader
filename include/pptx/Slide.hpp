#pragma once

#include "pptx/Shape.hpp"

#include <string>
#include <vector>

namespace pptx {

// 一张幻灯片（对应 ppt/slides/slideN.xml）
class Slide {
public:
    Slide() = default;

    // 演示文稿内的页码（从 1 开始，按 presentation.xml 中 sldIdLst 的顺序）
    int slideNumber() const noexcept { return slideNumber_; }
    void setSlideNumber(int n) noexcept { slideNumber_ = n; }

    std::vector<Shape>& shapes() noexcept { return shapes_; }
    const std::vector<Shape>& shapes() const noexcept { return shapes_; }

    // 标题：第一个名称以 "Title"/"标题" 开头的形状文本；找不到则返回空串
    std::string title() const;

    // 整页全文：所有含文本形状的文本按顺序用 '\n' 拼接
    std::string text() const;

    bool empty() const noexcept { return shapes_.empty(); }

    // 供范围 for 循环遍历本页内的形状
    auto begin() noexcept { return shapes_.begin(); }
    auto end() noexcept { return shapes_.end(); }
    auto begin() const noexcept { return shapes_.begin(); }
    auto end() const noexcept { return shapes_.end(); }

private:
    int slideNumber_ = 0;
    std::vector<Shape> shapes_;
};

}  // namespace pptx
