#include "pptx/PptxArchive.hpp"

#include <utility>

namespace pptx {

namespace {

std::string zipErrorMessage(int code) {
    zip_error_t e;
    zip_error_init_with_code(&e, code);
    const char* s = zip_error_strerror(&e);
    std::string msg = s ? s : "unknown zip error";
    zip_error_fini(&e);
    return msg;
}

}  // namespace

PptxArchive::PptxArchive(const std::string& filePath) : filePath_(filePath) {
    int err = 0;
    zip_ = zip_open(filePath_.c_str(), ZIP_RDONLY, &err);
    if (zip_ == nullptr) {
        throw std::runtime_error("pptx: cannot open '" + filePath_ + "': " + zipErrorMessage(err));
    }
}

PptxArchive::~PptxArchive() {
    if (zip_ != nullptr) {
        zip_close(zip_);
        zip_ = nullptr;
    }
}

PptxArchive::PptxArchive(PptxArchive&& other) noexcept
    : zip_(other.zip_), filePath_(std::move(other.filePath_)) {
    other.zip_ = nullptr;
}

PptxArchive& PptxArchive::operator=(PptxArchive&& other) noexcept {
    if (this != &other) {
        if (zip_ != nullptr) zip_close(zip_);
        zip_ = other.zip_;
        filePath_ = std::move(other.filePath_);
        other.zip_ = nullptr;
    }
    return *this;
}

bool PptxArchive::hasEntry(const std::string& path) const {
    zip_stat_t st;
    return zip_stat(zip_, path.c_str(), 0, &st) == 0;
}

std::string PptxArchive::readEntry(const std::string& path) const {
    zip_file_t* f = zip_fopen(zip_, path.c_str(), 0);
    if (f == nullptr) {
        throw std::runtime_error("pptx: cannot open entry '" + path + "' in '" + filePath_ + "'");
    }

    std::string out;
    out.reserve(4096);
    char buf[16384];
    for (;;) {
        zip_int64_t n = zip_fread(f, buf, sizeof(buf));
        if (n < 0) {
            zip_fclose(f);
            throw std::runtime_error("pptx: read error in entry '" + path + "'");
        }
        if (n == 0) break;  // EOF
        out.append(buf, static_cast<size_t>(n));
    }
    zip_fclose(f);
    return out;
}

}  // namespace pptx
