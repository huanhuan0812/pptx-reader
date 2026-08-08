#pragma once

#include "pptx/Paragraph.hpp"

#include <string>
#include <vector>

namespace pptx {

// 形状 / 文本框（对应 OOXML 中的 <p:sp>、<p:cxnSp>、<p:pic>、<p:graphicFrame> 等）
class Shape {
public:
    Shape() = default;

    // 形状 id（<p:cNvPr id="..."/>）
    int id() const noexcept { return id_; }
    void setId(int i) noexcept { id_ = i; }

    // 形状名称（<p:cNvPr name="..."/>，如 "Title 1"、"TextBox 5"）
    const std::string& name() const noexcept { return name_; }
    void setName(const std::string& n) { name_ = n; }

    // 形状类型："sp" / "cxnSp" / "pic" / "grpSp" / "graphicFrame"
    const std::string& type() const noexcept { return type_; }
    void setType(const std::string& t) { type_ = t; }

    std::vector<Paragraph>& paragraphs() noexcept { return paragraphs_; }
    const std::vector<Paragraph>& paragraphs() const noexcept { return paragraphs_; }

    // 是否包含文本（存在至少一个段落）
    bool hasText() const noexcept { return !paragraphs_.empty(); }

    // 形状全文：各段落文本用 '\n' 分隔
    std::string text() const;

    bool empty() const noexcept { return paragraphs_.empty(); }

    // 供范围 for 循环遍历本形状内的段落
    auto begin() noexcept { return paragraphs_.begin(); }
    auto end() noexcept { return paragraphs_.end(); }
    auto begin() const noexcept { return paragraphs_.begin(); }
    auto end() const noexcept { return paragraphs_.end(); }

private:
    int id_ = 0;
    std::string name_;
    std::string type_;
    std::vector<Paragraph> paragraphs_;
};

}  // namespace pptx
