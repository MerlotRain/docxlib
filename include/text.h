#ifndef DOCXTEXT_H
#define DOCXTEXT_H

#include "enumtext.h"
#include "docx_global.h"
#include "length.h"

#include <QDomElement>
#include <QImage>
#include <QList>
#include <QString>
#include <QDomDocument>

namespace Docx {
class Run;
class Text;
class CT_RPr;
class CT_PPr;
class DocumentPart;
class InlineShape;
class CT_Inline;

class DOCX_EXPORT Paragraph
{
public:
    Paragraph(DocumentPart *part, const QDomElement &element);
    ~Paragraph();

    Run *addRun(const QString &text = QString(), const QString &style = QString());
    Run *addRun(const QDomElement &beforeElement, const QString &text = QString(),
                const QString &style = QString());

    void addText(const QString &text);
    void setText(const QString &text);
    QString text() const;

    void setStyle(const QString &style);
    void setAlignment(WD_PARAGRAPH_ALIGNMENT align);

    Paragraph *insertParagraphBefore(const QString &text, const QString &style = QString());
    QDomNode currentNode() { return (QDomNode) *m_pEle; }
    Run *currentRun();

private:
    friend class DocumentPart; 
    DocumentPart *m_part;
    QDomDocument *m_dom;
    QDomElement *m_pEle;
    QList<Run *> m_runs;
    Run *m_pCurrentRun = nullptr;
    CT_PPr *m_style;
    friend class CT_PPr;
};

class DOCX_EXPORT Run
{
public:
    Run(DocumentPart *part, QDomElement *parent);
    Run(DocumentPart *part, QDomElement *parent, const QDomElement &ele, bool bInit = false);
    ~Run();
    void addTab();
    void addBreak(WD_BREAK breakType = WD_BREAK::PAGE);
    void addText(const QString &text);
    void setText(const QString &text);
    QString text() const;

    InlineShape *addPicture(const QString &path, const Length &width = Length(),
                            const Length &height = Length());
    InlineShape *addPicture(const QImage &img, const Length &width = Length(),
                            const Length &height = Length());

    void setStyle(const QString &style);
    void setAllcaps(bool isallcaps = true);
    void setSmallcaps(bool smallcpas = true);
    void setBold(bool isbold = true);
    void setItalic(bool isItalic = true);
    void setDoubleStrike(bool isDoubleStrike = true);
    void setShadow(bool shadow = true);
    void setUnderLine(WD_UNDERLINE underline);
    void addDrawing(CT_Inline *imline);
    void setFont(const QString &strFont);
    void setSize(WD_FONTSIZE size);
    void setColor(QColor &color);

private:
    InlineShape *scalePicture(InlineShape *picture, const Length &width, const Length &height);

private:
    QString m_text;
    DocumentPart *m_part;
    QDomDocument *m_dom;
    QDomElement m_rEle;
    QDomElement *m_parent;
    CT_RPr *m_style = nullptr;
    friend class CT_RPr;
};

}// namespace Docx

#endif// TEXT_H
