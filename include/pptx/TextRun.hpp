#pragma once

#include <string>
#include <utility>

namespace pptx {

// 一段具有相同格式的文本片段（对应 OOXML 中的 <a:r>）
class TextRun {
public:
    TextRun() = default;

    const std::string& text() const noexcept { return text_; }
    void setText(const std::string& t) { text_ = t; }
    void setText(std::string&& t) noexcept { text_ = std::move(t); }

    bool bold() const noexcept { return bold_; }
    void setBold(bool b) noexcept { bold_ = b; }

    bool italic() const noexcept { return italic_; }
    void setItalic(bool i) noexcept { italic_ = i; }

    bool underline() const noexcept { return underline_; }
    void setUnderline(bool u) noexcept { underline_ = u; }

    // 字号，单位：磅（pt）
    double fontSize() const noexcept { return fontSize_; }
    void setFontSize(double pt) noexcept { fontSize_ = pt; }

    // 西文字体名称（<a:latin typeface="..."/>）
    const std::string& fontName() const noexcept { return fontName_; }
    void setFontName(const std::string& n) { fontName_ = n; }

    // 颜色，形如 "RRGGBB"（十六进制）；未设置时为空串
    const std::string& color() const noexcept { return color_; }
    void setColor(const std::string& c) { color_ = c; }

    bool empty() const noexcept { return text_.empty(); }

private:
    std::string text_;
    bool bold_ = false;
    bool italic_ = false;
    bool underline_ = false;
    double fontSize_ = 0.0;
    std::string fontName_;
    std::string color_;
};

}  // namespace pptx
