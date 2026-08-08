#include "pptx/Presentation.hpp"

#include <pugixml.hpp>

#include <map>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

namespace pptx {

namespace {

// ---------- 关系文件 ----------

// 解析 ppt/_rels/presentation.xml.rels，得到 关系Id -> 幻灯片 XML 路径。
// 只保留 slide 关系，排除 layout / master / notes / handout 等。
std::map<std::string, std::string> parseSlideRelationships(PptxArchive& archive) {
    std::map<std::string, std::string> result;
    const std::string relsPath = "ppt/_rels/presentation.xml.rels";
    if (!archive.hasEntry(relsPath)) return result;

    pugi::xml_document doc;
    if (!doc.load_string(archive.readEntry(relsPath).c_str())) return result;

    auto relationships = doc.child("Relationships");
    for (auto rel : relationships.children("Relationship")) {
        std::string type = rel.attribute("Type").as_string();
        if (type.find("/slide") == std::string::npos) continue;
        if (type.find("/slideLayout") != std::string::npos) continue;
        if (type.find("/slideMaster") != std::string::npos) continue;
        if (type.find("/notesSlide") != std::string::npos) continue;
        if (type.find("/handoutMaster") != std::string::npos) continue;
        if (type.find("/slideUpdateInfo") != std::string::npos) continue;

        std::string target = rel.attribute("Target").as_string();
        if (target.empty()) continue;
        if (target[0] == '/') {
            target = target.substr(1);
        } else {
            target = "ppt/" + target;  // 相对于 ppt/ 目录
        }
        result[rel.attribute("Id").as_string()] = std::move(target);
    }
    return result;
}

// ---------- 文本解析 ----------

// 从形状节点中找到 <p:cNvPr>（不同形状类型的容器名不同）
pugi::xml_node findCNvPr(const pugi::xml_node& shapeNode) {
    static const char* const containers[] = {
        "p:nvSpPr", "p:nvCxnSpPr", "p:nvPicPr", "p:nvGrpSpPr", "p:nvGraphicFramePr"};
    for (const char* c : containers) {
        pugi::xml_node n = shapeNode.child(c).child("p:cNvPr");
        if (n) return n;
    }
    return {};
}

// 解析 <a:rPr> 的格式属性
void parseRunProperties(const pugi::xml_node& rPr, TextRun& run) {
    run.setBold(rPr.attribute("b").as_bool());
    run.setItalic(rPr.attribute("i").as_bool());
    const char* u = rPr.attribute("u").as_string();
    if (u != nullptr && *u != '\0' && std::string(u) != "none") {
        run.setUnderline(true);
    }

    const int sz = rPr.attribute("sz").as_int(0);
    if (sz > 0) run.setFontSize(sz / 100.0);  // OOXML 中 sz 的单位为 1/100 磅

    auto latin = rPr.child("a:latin");
    if (latin) run.setFontName(latin.attribute("typeface").as_string());

    auto fill = rPr.child("a:solidFill");
    if (fill) {
        auto srgb = fill.child("a:srgbClr");
        if (srgb) run.setColor(srgb.attribute("val").as_string());
    }
}

// 解析 <p:txBody> 中的段落
void parseTextBody(const pugi::xml_node& txBody, std::vector<Paragraph>& paragraphs) {
    for (auto p : txBody.children("a:p")) {
        Paragraph para;
        auto pPr = p.child("a:pPr");
        if (pPr) {
            para.setLevel(pPr.attribute("lvl").as_int(0));
            para.setAlignment(pPr.attribute("algn").as_string());
        }

        for (auto child : p.children()) {
            const std::string name = child.name();
            if (name == "a:r" || name == "a:fld") {
                // 文本 run（<a:r>）或域（<a:fld>，如页码），二者都含 <a:t>
                TextRun run;
                auto rPr = child.child("a:rPr");
                if (rPr) parseRunProperties(rPr, run);
                run.setText(child.child_value("a:t"));
                para.runs().push_back(std::move(run));
            } else if (name == "a:br") {
                // 行内换行：附加到最后一个 run 末尾；若还没有 run 则新建一个
                if (!para.runs().empty()) {
                    TextRun& last = para.runs().back();
                    last.setText(last.text() + "\n");
                } else {
                    TextRun run;
                    run.setText("\n");
                    para.runs().push_back(std::move(run));
                }
            }
            // <a:endParaRPr> 等其它节点忽略
        }
        paragraphs.push_back(std::move(para));
    }
}

// ---------- 形状解析 ----------

Shape parseShape(const pugi::xml_node& node, const std::string& type) {
    Shape shape;
    shape.setType(type);
    if (auto cNvPr = findCNvPr(node); cNvPr) {
        shape.setId(cNvPr.attribute("id").as_int(0));
        shape.setName(cNvPr.attribute("name").as_string());
    }
    auto txBody = node.child("p:txBody");
    if (txBody) parseTextBody(txBody, shape.paragraphs());
    return shape;
}

// 图形框架：通常承载表格，逐行逐单元格读取文本
Shape parseGraphicFrame(const pugi::xml_node& node) {
    Shape shape;
    shape.setType("graphicFrame");
    if (auto cNvPr = findCNvPr(node); cNvPr) {
        shape.setId(cNvPr.attribute("id").as_int(0));
        shape.setName(cNvPr.attribute("name").as_string());
    }
    auto tbl = node.child("a:graphic").child("a:graphicData").child("a:tbl");
    if (tbl) {
        for (auto tr : tbl.children("a:tr")) {
            for (auto tc : tr.children("a:tc")) {
                auto txBody = tc.child("a:txBody");
                if (txBody) parseTextBody(txBody, shape.paragraphs());
            }
        }
    }
    return shape;
}

// 去掉 XML 前缀，返回本地名称（"p:sp" -> "sp"）
std::string localName(const std::string& full) {
    const auto pos = full.find(':');
    return (pos == std::string::npos) ? full : full.substr(pos + 1);
}

// 递归解析 spTree，保持文档顺序；组形状 (grpSp) 会被展开平铺
void parseShapeTree(const pugi::xml_node& spTree, std::vector<Shape>& shapes) {
    for (auto child : spTree.children()) {
        const std::string name = child.name();
        if (name == "p:sp" || name == "p:cxnSp" || name == "p:pic") {
            shapes.push_back(parseShape(child, localName(name)));
        } else if (name == "p:grpSp") {
            auto subTree = child.child("p:spTree");
            if (subTree) parseShapeTree(subTree, shapes);
        } else if (name == "p:graphicFrame") {
            shapes.push_back(parseGraphicFrame(child));
        }
    }
}

Slide parseSlide(const std::string& xml, int number) {
    Slide slide;
    slide.setSlideNumber(number);

    pugi::xml_document doc;
    if (!doc.load_string(xml.c_str())) {
        throw std::runtime_error("pptx: failed to parse slide XML");
    }
    auto root = doc.child("p:sld");
    if (!root) return slide;

    auto cSld = root.child("p:cSld");
    auto spTree = cSld.child("p:spTree");
    if (spTree) parseShapeTree(spTree, slide.shapes());
    return slide;
}

}  // namespace

Presentation::Presentation(std::string filePath) : archive_(std::move(filePath)) {}

Presentation Presentation::open(const std::string& filePath) {
    Presentation pres(filePath);
    pres.parseCoreProperties();
    pres.parseSlides();
    return pres;
}

void Presentation::parseSlides() {
    auto rels = parseSlideRelationships(archive_);
    if (rels.empty()) {
        throw std::runtime_error("pptx: no slide relationships found in '" + archive_.filePath() +
                                 "'");
    }

    pugi::xml_document doc;
    const std::string presentationXml = archive_.readEntry("ppt/presentation.xml");
    if (!doc.load_string(presentationXml.c_str())) {
        throw std::runtime_error("pptx: failed to parse ppt/presentation.xml");
    }

    auto sldIdLst = doc.child("p:presentation").child("p:sldIdLst");
    int number = 0;
    for (auto sldId : sldIdLst.children("p:sldId")) {
        const std::string rid = sldId.attribute("r:id").as_string();
        auto it = rels.find(rid);
        if (it == rels.end()) continue;  // 关系缺失则跳过
        ++number;
        slides_.push_back(parseSlide(archive_.readEntry(it->second), number));
    }

    // 若 sldIdLst 缺失或为空，退化为按关系表顺序读取
    if (slides_.empty() && !rels.empty()) {
        for (const auto& kv : rels) {
            ++number;
            slides_.push_back(parseSlide(archive_.readEntry(kv.second), number));
        }
    }
}

void Presentation::parseCoreProperties() {
    const std::string path = "docProps/core.xml";
    if (!archive_.hasEntry(path)) return;

    pugi::xml_document doc;
    if (!doc.load_string(archive_.readEntry(path).c_str())) return;
    auto root = doc.child("cp:coreProperties");
    if (!root) return;

    documentTitle_ = root.child_value("dc:title");
    creator_ = root.child_value("dc:creator");
    lastModifiedBy_ = root.child_value("cp:lastModifiedBy");
    created_ = root.child_value("dcterms:created");
}

std::string Presentation::text() const {
    std::string out;
    bool first = true;
    for (const auto& slide : slides_) {
        if (!first) out += '\n';
        out += slide.text();
        first = false;
    }
    return out;
}

}  // namespace pptx
