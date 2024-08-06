#ifndef DOCUMENTPART_H
#define DOCUMENTPART_H

#include "length.h"
#include "../opc/part.h"
#include "docx_global.h"
#include "enumtext.h"

#include <QByteArray>
#include <QDomDocument>
#include <QImage>
#include <QPair>

namespace Docx {

class CT_Default;
class Paragraph;
class Run;
class Table;
class InlineShapes;
class ImagePart;
class ImageParts;
class InlineShape;

class DocumentPart : public Part
{
public:
    DocumentPart(const QString &partName, const QString &contentType,
                 const QByteArray &blob = QByteArray(), Package *package = nullptr);
    ~DocumentPart();

    Paragraph *addParagraph(const QString &text, const QString &style = QStringLiteral(""));
    static DocumentPart *load(const PackURI &partName, const QString &contentType,
                              const QByteArray &blob = QByteArray(), Package *package = nullptr);
    Table *addTable(int rows, int cols, const QString &style = QString::fromLatin1("TableGrid"));
    void afterUnmarshal();
    QDomDocument *element() const;
    QByteArray blob() const;
    QPair<ImagePart *, QString> getOrAddImagePart(const PackURI &imageDescriptor);
    QPair<ImagePart *, QString> getOrAddImagePart(const QImage &img);
    QPair<ImagePart *, QString> getOrAddImagePart(ImagePart *imagPart);
    QList<Paragraph *> paragraphs();
    QList<Table *> tables();
    int nextId();

    Paragraph *addTextByBookmark(const QString &strBookMark, const QString &strText = QString(),
                                 WD_FONTSIZE fz = WD_FONTSIZE::SIZE5,
                                 WD_PARAGRAPH_ALIGNMENT align = WD_PARAGRAPH_ALIGNMENT::CENTER,
                                 const QString &font = QStringLiteral("Microsoft YaHei"));
    Paragraph *addPictureByBookmark(const QString &strBookMark, const QString &strPath = QString(),
                                    const Length &width = Length(),
                                    const Length &height = Length());
    Paragraph *getParagraphByNode(QDomNode findNode);
    Paragraph *getNextParagraphByNode(QDomNode &findNode);
    Paragraph *currentParagraph();
    void setCurrentParagraph(Paragraph *graph);
    void setHeader(const QString &strText = QString());
    void addCloneTable(Table *pCloneTable, Paragraph *graph = nullptr);
    void addBookmark(Paragraph *pParaph, const QString &strBookmark);
    Table *getTableByBookmark(const QString &strBookMark);

private:
    void findAttributes(const QDomNodeList &elements, const QString &attr, QList<QString> *nums);
    QDomNode lastSectionPtr() const;

private:
    //Body *m_body;
    QDomDocument *m_dom;
    InlineShapes *m_inlineShapes;

    // add
    QList<Paragraph *> m_addParagraphs;
    QList<Table *> m_addTables;

    // load
    QList<Paragraph *> m_ps;
    QList<Table *> m_tables;

    friend class Paragraph;
    friend class Run;
    friend class InlineShapes;
    friend class Table;
};

class InlineShapes
{
public:
    InlineShapes(DocumentPart *part);
    ~InlineShapes();

    InlineShape *addPicture(const QString &imagePath, Run *run);
    InlineShape *addPicture(const QImage &img, Run *run);

private:
    InlineShape *addPicture(const QPair<ImagePart *, QString> &imagePartAndId, Run *run);

private:
    DocumentPart *m_part;
    QDomDocument *m_dom;
    friend class InlineShape;
};

}// namespace Docx

#endif// DOCUMENTPART_H
