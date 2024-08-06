#include "oxmltext.h"
#include "../shared.h"
#include "text.h"
#include <QString>
#include <QtXml>

namespace Docx {

CT_PPr::CT_PPr(Paragraph *paragraph) : m_paragraph(paragraph) { loadExistStyle(); }

void CT_PPr::loadExistStyle()
{
    QDomNodeList eleStyle = m_paragraph->m_pEle->elementsByTagName(QStringLiteral("w:pPr"));
    if (eleStyle.isEmpty()) return;

    m_style = eleStyle.at(0).toElement();

    // pStyle
    eleStyle = m_style.elementsByTagName(QStringLiteral("w:pStyle"));
    if (!eleStyle.isEmpty()) m_pStyle = eleStyle.at(0).toElement();
}

void CT_PPr::setStyle(const QString &style)
{
    addOrAssignStyle();
    if (m_pStyle.isNull())
        m_pStyle = addOrAssignElement(m_paragraph->m_dom, &m_style, QStringLiteral("w:pStyle"));
    if (!style.isEmpty()) m_pStyle.setAttribute(QStringLiteral("w:val"), style);
}

void CT_PPr::setAlignment(WD_PARAGRAPH_ALIGNMENT align)
{
    addOrAssignStyle();
    QDomElement alignStyle =
            addOrAssignElement(m_paragraph->m_dom, &m_style, QStringLiteral("w:jc"));
    QString str = paragraphAlignToString(align);
    alignStyle.setAttribute(QStringLiteral("w:val"), str);
}

void CT_PPr::addOrAssignStyle()
{
    if (m_style.isNull())
        m_style = addOrAssignElement(m_paragraph->m_dom, m_paragraph->m_pEle,
                                     QStringLiteral("w:pPr"), true);
}

CT_PPr::~CT_PPr() {}

CT_RPr::CT_RPr(Run *run) : m_run(run) { loadExistStyle(); }

void CT_RPr::setSize(WD_FONTSIZE &size)
{

    addOrAssignStyle();
    QDomElement szEle = addOrAssignElement(m_run->m_dom, &m_style, QStringLiteral("w:sz"));
    szEle.setAttribute("w:val", QString("%1").arg((int) size));
    szEle = addOrAssignElement(m_run->m_dom, &m_style, QStringLiteral("w:szCs"));
    szEle.setAttribute("w:val", QString("%1").arg((int) size));
}

void CT_RPr::setColor(QColor &color)
{
    if (!color.isValid()) return;

    addOrAssignStyle();
    QDomElement colorEle = addOrAssignElement(m_run->m_dom, &m_style, QStringLiteral("w:color"));
    colorEle.setAttribute("w:val", convertRGB16HexStr(color));
}

void CT_RPr::loadExistStyle()
{
    QDomNodeList eles = m_run->m_rEle.elementsByTagName(QStringLiteral("w:rPr"));
    if (eles.isEmpty()) return;
    m_style = eles.at(0).toElement();
}

void CT_RPr::setStyle(const QString &style)
{
    addOrAssignStyle();
    if (m_rStyle.isNull())
        m_rStyle = addOrAssignElement(m_run->m_dom, &m_style, QStringLiteral("w:rStyle"));
    if (!style.isEmpty()) m_rStyle.setAttribute(QStringLiteral("w:val"), style);
}

void CT_RPr::setBold(bool bold) { addOrAssignStyleChildElement(QStringLiteral("w:b"), bold); }

void CT_RPr::setAllcaps(bool isallcaps)
{
    addOrAssignStyleChildElement(QStringLiteral("w:caps"), isallcaps);
}

void CT_RPr::setSmallcaps(bool issmallcpas)
{
    addOrAssignStyleChildElement(QStringLiteral("w:smallCaps"), issmallcpas);
}

void CT_RPr::setItalic(bool italic) { addOrAssignStyleChildElement(QStringLiteral("w:i"), italic); }

void CT_RPr::setDoubleStrike(bool isDoubleStrike)
{
    addOrAssignStyleChildElement(QStringLiteral("w:dstrike"), isDoubleStrike);
}

void CT_RPr::setShadow(bool shadow)
{
    addOrAssignStyleChildElement(QStringLiteral("w:shadow"), shadow);
}

void CT_RPr::setUnderLine(WD_UNDERLINE underline)
{
    addOrAssignStyle();
    QDomElement line = addOrAssignElement(m_run->m_dom, &m_style, QStringLiteral("w:u"));

    line.setAttribute("w:val", underLineToString(underline));
}

void CT_RPr::setFont(const QString &strFont) {}


void CT_RPr::addOrAssignStyle()
{
    if (m_style.isNull())
        m_style = addOrAssignElement(m_run->m_dom, &m_run->m_rEle, QStringLiteral("w:rPr"), true);
}

QString CT_RPr::convertRGB16HexStr(QColor &color)
{
    QString redStr = QString("%1").arg(color.red(), 2, 16, QChar('0'));
    QString greenStr = QString("%1").arg(color.green(), 2, 16, QChar('0'));
    QString blueStr = QString("%1").arg(color.blue(), 2, 16, QChar('0'));
    QString hexStr = "#" + redStr + greenStr + blueStr;
    return hexStr;
}

void CT_RPr::addOrAssignStyleChildElement(const QString &eleName, bool enable)
{
    addOrAssignStyle();
    QDomElement m_bold = addOrAssignElement(m_run->m_dom, &m_style, eleName);
    if (!enable) m_bold.setAttribute("w:val", 0);
    else
        m_bold.removeAttribute("w:val");
}

CT_RPr::~CT_RPr() {}

}// namespace Docx