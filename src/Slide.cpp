#include "pptx/Slide.hpp"

#include <cctype>

namespace pptx {

namespace {

bool startsWithIgnoreCase(const std::string& s, const char* prefix) {
    for (const char* p = prefix; *p != '\0'; ++p) {
        const std::size_t i = static_cast<std::size_t>(p - prefix);
        if (i >= s.size()) return false;
        if (std::tolower(static_cast<unsigned char>(s[i])) !=
            std::tolower(static_cast<unsigned char>(*p))) {
            return false;
        }
    }
    return true;
}

}  // namespace

std::string Slide::title() const {
    for (const auto& shape : shapes_) {
        const std::string& n = shape.name();
        const bool isTitle = (n.size() >= 5 && startsWithIgnoreCase(n, "Title")) ||
                             (n.find("标题") != std::string::npos);
        if (isTitle && shape.hasText()) {
            return shape.text();
        }
    }
    return {};
}

std::string Slide::text() const {
    std::string out;
    bool first = true;
    for (const auto& shape : shapes_) {
        if (!shape.hasText()) continue;
        if (!first) out += '\n';
        out += shape.text();
        first = false;
    }
    return out;
}

}  // namespace pptx
