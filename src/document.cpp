#include "document.h"
#include "./opc/opcpackage.h"
#include "./parts/documentpart.h"
#include "package.h"
#include "table.h"
#include "text.h"
#include <QDebug>
#include <QFile>
#include <QLocale>

namespace Docx {

Document::Document()
{
    qDebug() << "construct docx document.";
    if (QLocale::system().name() == QStringLiteral("zh_CN"))
    {
        open(QStringLiteral("://default_zh_CN.docx"));
    }
    else { open(QStringLiteral("://default.docx")); }
}

Document::Document(const QString &name)
{
    Q_ASSERT_X(QFile::exists(name), "filed", "can not find the path!");
    open(name);
}

Document::Document(QIODevice *device) { open(device); }

void Document::open(const QString &name)
{
    m_package = Package::open(name);
    m_docPart = m_package->mainDocument();
}

void Document::open(QIODevice *device)
{
    m_package = Package::open(device);
    m_docPart = m_package->mainDocument();
}

Paragraph *Document::addParagraph(const QString &text, const QString &style)
{
    if (!m_docPart) return nullptr;

    return m_docPart->addParagraph(text, style);
}

Paragraph *Document::addHeading(const QString &text, int level)
{
    QString style;
    if (level == 0) style = "Title";
    else
        style = QString("Heading%1").arg(level);
    return addParagraph(text, style);
}

Table *Document::addTable(int rows, int cols, const QString &style)
{
    if (!m_docPart) return nullptr;

    return m_docPart->addTable(rows, cols, style);
}

void Document::addCloneTable(Table *pCloneTable, Paragraph *graph)
{
    if (!m_docPart) return;

    m_docPart->addCloneTable(pCloneTable, graph);
}

InlineShape *Document::addPicture(const QString &imgPath, const Length &width, const Length &height)
{
    // Q_ASSERT_X(QFile::exists(imgPath), "add image filed", "can not find the Image path!");

    if (QFile::exists(imgPath))
    {
        Run *run = addParagraph()->addRun();
        InlineShape *picture = run->addPicture(imgPath, width, height);
        return picture;
    }
    return nullptr;
}

InlineShape *Document::addPicture(const QImage &img, const Length &width, const Length &height)
{
    Run *run = addParagraph()->addRun();
    InlineShape *picture = run->addPicture(img, width, height);

    return picture;
}

Paragraph *Document::addPageBreak()
{
    Paragraph *p = addParagraph();
    Run *run = p->addRun();
    run->addBreak();
    return p;
}

QList<Paragraph *> Document::paragraphs()
{
    QList<Paragraph *> m_lstParagraphs;
    if (!m_docPart) return m_lstParagraphs;

    return m_docPart->paragraphs();
}

QList<Table *> Document::tables()
{
    QList<Table *> lstTables;
    if (!m_docPart) return lstTables;

    return m_docPart->tables();
}

Document::~Document()
{
    qDebug() << "delete Docx::Document.";
    delete m_docPart;
    delete m_package;
}

void Document::save(const QString &path)
{
    qDebug() << "save docx file: " << path;
    m_package->save(path);
}

Paragraph *Document::addTextByBookmark(const QString &bookmark, const QString &strText,
                                       WD_FONTSIZE fz, WD_PARAGRAPH_ALIGNMENT align,
                                       const QString &font)
{
    if (!m_docPart) return nullptr;

    return m_docPart->addTextByBookmark(bookmark, strText, fz, align, font);
}

Paragraph *Document::addPictureByBookmark(const QString &bookmark, const QString &strPath,
                                          const Length &width, const Length &height)
{
    if (!m_docPart) return nullptr;

    return m_docPart->addPictureByBookmark(bookmark, strPath, width, height);
}

Paragraph *Document::currentParagraph()
{
    if (!m_docPart) return nullptr;

    return m_docPart->currentParagraph();
}

void Document::setCurrentParagraph(Paragraph *graph)
{
    if (!m_docPart) return;

    return m_docPart->setCurrentParagraph(graph);
}

void Document::addBookmark(Paragraph *pParaph, const QString &bookmark)
{
    if (!m_docPart) return;

    return m_docPart->addBookmark(pParaph, bookmark);
}

Paragraph *Document::moveParagraph(Paragraph *curGraph, int nMove)
{
    int nBegin = 0, nCnt = 0, nMovePos = -1;
    QList<Paragraph *> graphs = paragraphs();
    nCnt = graphs.size();
    if (nCnt < 1) return nullptr;

    foreach (auto &graph, graphs)
    {
        nBegin++;
        if (graph == curGraph) break;
    }

    if (0 == nBegin) return nullptr;

    nMovePos = nBegin - 1 + nMove;
    if (nMovePos >= 0 && nMovePos < nCnt) { return graphs.at(nMovePos); }

    return nullptr;
}

Table *Document::getTableByBookmark(const QString &bookmark)
{
    if (!m_docPart) return NULL;

    return m_docPart->getTableByBookmark(bookmark);
}

void Document::setHeader(const QString &strText) { Q_UNUSED(strText) }

}// namespace Docx