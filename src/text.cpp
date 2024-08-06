#include "text.h"
#include "oxml/oxmlshape.h"
#include "oxml/oxmltext.h"
#include "parts/documentpart.h"
#include "shape.h"
#include "shared.h"

#include <QDebug>
#include <QDomDocument>
#include <math.h>

namespace Docx {

Paragraph::Paragraph(DocumentPart *part, const QDomElement &element) : m_part(part)
{
    m_dom = part->m_dom;
    m_pEle = new QDomElement(element);
    m_style = new CT_PPr(this);

    QDomNodeList nodeList = m_pEle->elementsByTagName(QStringLiteral("w:r"));
    if (!nodeList.isEmpty())
    {
        for (int i = 0; i < nodeList.count(); i++)
        {
            QDomElement nodeEle = nodeList.at(i).toElement();
            Run *run = new Run(m_part, m_pEle, nodeEle);
            m_runs.append(run);
        }
    }
}

Run *Paragraph::addRun(const QString &text, const QString &style)
{
    Run *run = new Run(m_part, m_pEle);
    if (!text.isEmpty()) run->addText(text);
    if (!style.isEmpty()) setStyle(style);

    m_runs.append(run);
    m_pCurrentRun = run;
    return run;
}
Run *Paragraph::addRun(const QDomElement &beforeElement, const QString &text, const QString &style)
{
    Run *run = new Run(m_part, m_pEle, beforeElement, true);
    if (!text.isEmpty()) run->addText(text);
    if (!style.isEmpty()) setStyle(style);

    m_pCurrentRun = run;
    return run;
}

void Paragraph::addText(const QString &text) { addRun(text); }

void Paragraph::setText(const QString &text)
{
    Run *pRun = currentRun();
    if (!pRun) return;

    pRun->setText(text);
}

QString Paragraph::text() const
{
    QStringList str;
    for (const Run *r: m_runs) { str.append(r->text()); }
    return str.join("");
}

void Paragraph::setStyle(const QString &style) { m_style->setStyle(style); }

void Paragraph::setAlignment(WD_PARAGRAPH_ALIGNMENT align) { m_style->setAlignment(align); }

Paragraph *Paragraph::insertParagraphBefore(const QString &text, const QString &style)
{
    QDomElement pEle = m_dom->createElement(QStringLiteral("w:p"));
    Paragraph *p = new Paragraph(m_part, pEle);
    p->addRun(text, style);
    QDomNode parent = m_pEle->parentNode();

    parent.insertBefore(pEle, *m_pEle);
    return p;
}

Run *Paragraph::currentRun()
{
    if (m_pCurrentRun) return m_pCurrentRun;

    int nSize = m_runs.size();
    if (nSize > 0) { return m_runs.at(0); }
    return nullptr;
}

Paragraph::~Paragraph()
{
    qDeleteAll(m_runs);
    delete m_pEle;
}

Run::Run(DocumentPart *part, QDomElement *parent) : m_part(part), m_parent(parent)
{
    m_dom = part->m_dom;
    m_rEle = m_dom->createElement(QStringLiteral("w:r"));
    m_parent->appendChild(m_rEle);
    m_style = new CT_RPr(this);
}

Run::Run(DocumentPart *part, QDomElement *parent, const QDomElement &ele, bool bInit)
    : m_part(part), m_parent(parent)
{
    if (false == bInit)
    {
        m_dom = part->m_dom;
        m_rEle = QDomElement(ele);
        m_style = new CT_RPr(this);
    }
    else
    {
        m_dom = part->m_dom;
        m_rEle = m_dom->createElement(QStringLiteral("w:r"));
        m_parent->insertAfter(m_rEle, ele);
        m_style = new CT_RPr(this);
    }
}

void Run::addTab()
{
    QDomElement tabEle = m_dom->createElement(QStringLiteral("w:tab"));
    m_rEle.appendChild(tabEle);
}

void Run::addBreak(WD_BREAK breakType)
{
    QDomElement brele = m_dom->createElement(QStringLiteral("w:br"));
    QString elementType, elementClear;
    switch (breakType)
    {
        case WD_BREAK::LINE:
            break;
        case WD_BREAK::PAGE:
            elementType = QStringLiteral("page");
            break;
        case WD_BREAK::COLUMN:
            elementType = QStringLiteral("column");
            break;
        case WD_BREAK::LINE_CLEAR_LEFT:
            elementType = QStringLiteral("textWrapping");
            elementClear = QStringLiteral("left");
            break;
        case WD_BREAK::LINE_CLEAR_RIGHT:
            elementType = QStringLiteral("textWrapping");
            elementClear = QStringLiteral("right");
            break;
        case WD_BREAK::LINE_CLEAR_ALL:
            elementType = QStringLiteral("textWrapping");
            elementClear = QStringLiteral("all");
            break;
        default:
            break;
    }
    m_rEle.appendChild(brele);
    if (!elementType.isEmpty()) brele.setAttribute(QStringLiteral("w:type"), elementType);

    if (!elementClear.isEmpty()) brele.setAttribute(QStringLiteral("w:clear"), elementClear);
}

void Run::addText(const QString &text)
{
    QDomElement tEle = m_dom->createElement(QStringLiteral("w:t"));
    tEle.appendChild(m_dom->createTextNode(text));
    m_rEle.appendChild(tEle);
}

QString Run::text() const { return m_rEle.text(); }

void Run::setText(const QString &text)
{
    if (m_rEle.isNull()) return;

    QString strText = m_rEle.text();

    QDomNodeList nodeLst = m_rEle.elementsByTagName(QStringLiteral("w:t"));
    int nSize = nodeLst.size();
    for (int i = 0; i < nSize; i++)
    {
        QDomNode node = nodeLst.at(0);
        node.toElement().firstChild().setNodeValue(text);
    }
}

InlineShape *Run::addPicture(const QString &path, const Length &width, const Length &height)
{
    InlineShapes *ships = m_part->m_inlineShapes;
    InlineShape *picture = ships->addPicture(path, this);
    return scalePicture(picture, width, height);
}

InlineShape *Run::scalePicture(InlineShape *picture, const Length &width, const Length &height)
{
    if (!width.isEmpty() || !height.isEmpty())
    {
        int nWidth = width.emu();
        int nHeight = height.emu();

        int native_width = picture->width().emu();
        int native_height = picture->height().emu();
        if (width.isEmpty())
        {
            float scaling_factor = float(nHeight) / float(native_height);
            nWidth = int(round(native_width * scaling_factor));
        }
        if (height.isEmpty())
        {
            float scaling_factor = float(nWidth) / float(native_width);
            nHeight = int(round(native_height * scaling_factor));
        }
        picture->setWidth(nWidth);
        picture->setHeight(nHeight);
    }
    return picture;
}

InlineShape *Run::addPicture(const QImage &img, const Length &width, const Length &height)
{
    InlineShapes *ships = m_part->m_inlineShapes;
    InlineShape *picture = ships->addPicture(img, this);
    return scalePicture(picture, width, height);
}

void Run::setStyle(const QString &style) { m_style->setStyle(style); }

Run::~Run() { delete m_style; }

void Run::setAllcaps(bool isallcaps) { m_style->setAllcaps(isallcaps); }

void Run::setSmallcaps(bool smallcpas) { m_style->setSmallcaps(smallcpas); }

void Run::setBold(bool isbold) { m_style->setBold(isbold); }

void Run::setItalic(bool isItalic) { m_style->setItalic(isItalic); }

void Run::setDoubleStrike(bool isDoubleStrike) { m_style->setDoubleStrike(isDoubleStrike); }

void Run::setShadow(bool shadow) { m_style->setShadow(shadow); }

void Run::setUnderLine(WD_UNDERLINE underline) { m_style->setUnderLine(underline); }

void Run::addDrawing(CT_Inline *imline)
{
    QDomElement drawing = m_dom->createElement(QStringLiteral("w:drawing"));

    m_rEle.appendChild(drawing);
    drawing.appendChild(imline->element());
}

void Run::setFont(const QString &strFont) { m_style->setFont(strFont); }

void Run::setSize(WD_FONTSIZE size) { m_style->setSize(size); }

void Run::setColor(QColor &color) { m_style->setColor(color); }

}// namespace Docx
