#pragma once

#include <zip.h>

#include <string>

namespace pptx {

// 基于 libzip 的 PPTX（zip）读取封装：直接从压缩包内读取文件，不解压到磁盘。
// 线程不安全，仅用于单线程读取。
class PptxArchive {
public:
    // 打开 pptx 文件；失败时抛出 std::runtime_error
    explicit PptxArchive(const std::string& filePath);
    ~PptxArchive();

    PptxArchive(const PptxArchive&) = delete;
    PptxArchive& operator=(const PptxArchive&) = delete;

    PptxArchive(PptxArchive&& other) noexcept;
    PptxArchive& operator=(PptxArchive&& other) noexcept;

    const std::string& filePath() const noexcept { return filePath_; }

    // 压缩包内是否存在该路径的条目
    bool hasEntry(const std::string& path) const;

    // 读取某个条目的原始内容（XML 为 UTF-8 文本，原样返回）。
    // 条目不存在或读取失败时抛出 std::runtime_error。
    std::string readEntry(const std::string& path) const;

private:
    zip_t* zip_ = nullptr;
    std::string filePath_;
};

}  // namespace pptx
