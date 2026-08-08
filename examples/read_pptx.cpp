#include <pptx/pptx.hpp>

#include <exception>
#include <iostream>

namespace {

void printUsage(const char* argv0) {
    std::cerr << "用法: " << argv0 << " <file.pptx>\n";
}

}  // namespace

int main(int argc, char** argv) {
    if (argc < 2) {
        printUsage(argv[0]);
        return 1;
    }

    try {
        auto presentation = pptx::Presentation::open(argv[1]);

        std::cout << "演示文稿标题: " << presentation.documentTitle() << "\n"
                  << "作者: " << presentation.creator() << "\n"
                  << "幻灯片数量: " << presentation.slideCount() << "\n\n";

        // 嵌套 for 循环按顺序读取每一个文本框的文字
        for (const auto& slide : presentation) {           // 遍历幻灯片
            std::cout << "===== 第 " << slide.slideNumber() << " 页 =====";
            if (!slide.title().empty()) {
                std::cout << "  标题: " << slide.title();
            }
            std::cout << "\n";

            for (const auto& shape : slide) {              // 遍历文本框
                if (!shape.hasText()) continue;
                std::cout << "  [文本框] "
                          << (shape.name().empty() ? "(未命名)" : shape.name())
                          << " (类型: " << shape.type() << ")\n";

                for (const auto& paragraph : shape) {      // 遍历段落
                    if (paragraph.runs().empty()) continue;
                    std::cout << "    (层级 " << paragraph.level() << ") ";
                    for (const auto& run : paragraph) {    // 遍历文本片段
                        if (run.bold()) std::cout << "[B]";
                        if (run.italic()) std::cout << "[I]";
                        if (run.underline()) std::cout << "[U]";
                        std::cout << run.text();
                    }
                    std::cout << "\n";
                }
            }
            std::cout << "\n";
        }

        // 若只需要整篇纯文本：
        // std::cout << presentation.text();
    } catch (const std::exception& e) {
        std::cerr << "读取失败: " << e.what() << "\n";
        return 1;
    }
    return 0;
}
