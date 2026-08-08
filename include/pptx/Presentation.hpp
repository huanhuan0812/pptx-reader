#pragma once

#include "pptx/PptxArchive.hpp"
#include "pptx/Slide.hpp"

#include <string>
#include <vector>

namespace pptx {

// 一个 PPTX 演示文稿。解析时直接读取 zip 内的 XML（libzip + PugiXML），不解压。
class Presentation {
public:
    // 打开并解析 pptx 文件；失败时抛出 std::runtime_error
    static Presentation open(const std::string& filePath);

    Presentation(const Presentation&) = delete;
    Presentation& operator=(const Presentation&) = delete;
    Presentation(Presentation&&) noexcept = default;
    Presentation& operator=(Presentation&&) noexcept = default;

    int slideCount() const noexcept { return static_cast<int>(slides_.size()); }

    std::vector<Slide>& slides() noexcept { return slides_; }
    const std::vector<Slide>& slides() const noexcept { return slides_; }

    // 核心属性（来自 docProps/core.xml）
    const std::string& documentTitle() const noexcept { return documentTitle_; }
    const std::string& creator() const noexcept { return creator_; }
    const std::string& lastModifiedBy() const noexcept { return lastModifiedBy_; }
    const std::string& created() const noexcept { return created_; }

    // 全部幻灯片文本按顺序拼接
    std::string text() const;

    bool empty() const noexcept { return slides_.empty(); }

    // 供范围 for 循环遍历所有幻灯片
    auto begin() noexcept { return slides_.begin(); }
    auto end() noexcept { return slides_.end(); }
    auto begin() const noexcept { return slides_.begin(); }
    auto end() const noexcept { return slides_.end(); }

private:
    explicit Presentation(std::string filePath);

    void parseSlides();
    void parseCoreProperties();

    PptxArchive archive_;
    std::vector<Slide> slides_;
    std::string documentTitle_;
    std::string creator_;
    std::string lastModifiedBy_;
    std::string created_;
};

}  // namespace pptx
