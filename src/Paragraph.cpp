#include "pptx/Paragraph.hpp"

namespace pptx {

std::string Paragraph::text() const {
    std::string out;
    for (const auto& run : runs_) {
        out += run.text();
    }
    return out;
}

}  // namespace pptx
