#ifndef DOCUMENT_H
#define DOCUMENT_H

#include "docx_global.h"
#include "enumtext.h"
#include "length.h"
#include <QIODevice>
#include <QImage>
#include <QString>

namespace Docx {

class Paragraph;
class Table;
class DocumentPart;
class Package;
class InlineShape;

class DOCX_EXPORT Document
{
public:
    Document();
    explicit Document(const QString &name);
    explicit Document(QIODevice *device);
    ~Document();

    Paragraph *addParagraph(const QString &text = QString(), const QString &style = QString());
    Paragraph *addHeading(const QString &text = QString(), int level = 1);
    Table *addTable(int rows, int cols, const QString &style = QStringLiteral("TableGrid"));
    InlineShape *addPicture(const QString &imgPath, const Length &width = Length(),
                            const Length &height = Length());
    InlineShape *addPicture(const QImage &img, const Length &width = Length(),
                            const Length &height = Length());
    Paragraph *addPageBreak();
    QList<Paragraph *> paragraphs();
    QList<Table *> tables();

    Paragraph *addTextByBookmark(const QString &bookmark, const QString &text = QString(),
                                 WD_FONTSIZE fz = WD_FONTSIZE::SIZE5,
                                 WD_PARAGRAPH_ALIGNMENT align = WD_PARAGRAPH_ALIGNMENT::CENTER,
                                 const QString &font = QStringLiteral("Microsoft YaHei"));
    Paragraph *addPictureByBookmark(const QString &bookmark, const QString &path = QString(),
                                    const Length &width = Length(),
                                    const Length &height = Length());
    Paragraph *currentParagraph();
    Paragraph *moveParagraph(Paragraph *graph, int step);
    Table *getTableByBookmark(const QString &bookmark);

    void setCurrentParagraph(Paragraph *graph);
    void setHeader(const QString &text = QString());
    void addCloneTable(Table *pCloneTable, Paragraph *graph = nullptr);
    void addBookmark(Paragraph *pParaph, const QString &bookmark);

    void save(const QString &path);

private:
    void open(const QString &name);
    void open(QIODevice *device);

private:
    DocumentPart *m_docPart;
    Package *m_package;
};

}// namespace Docx

#endif// DOCUMENT_H
