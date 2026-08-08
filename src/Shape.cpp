#include "pptx/Shape.hpp"

namespace pptx {

std::string Shape::text() const {
    std::string out;
    bool first = true;
    for (const auto& p : paragraphs_) {
        if (!first) out += '\n';
        out += p.text();
        first = false;
    }
    return out;
}

}  // namespace pptx
