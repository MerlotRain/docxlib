#ifndef OXMLDOCXTEXT_H
#define OXMLDOCXTEXT_H

#include "enumtext.h"
#include "xmlchemy.h"
#include <QColor>

namespace Docx {
    
class Run;
class Paragraph;

class CT_PPr
{
public:
    CT_PPr(Paragraph *paragraph);
    void setStyle(const QString &style = QString());
    void setAlignment(WD_PARAGRAPH_ALIGNMENT align);
    void addOrAssignStyle();
    virtual ~CT_PPr();


    void loadExistStyle();

private:
    QDomElement m_style;
    QDomElement m_pStyle;
    Paragraph *m_paragraph;
};


class CT_RPr
{
public:
    CT_RPr(Run *run);
    ~CT_RPr();
    void setStyle(const QString &style = QString());
    void setBold(bool bold);
    void setAllcaps(bool isallcaps);
    void setSmallcaps(bool issmallcpas);
    void setItalic(bool italic);
    void setDoubleStrike(bool isDoubleStrike);
    void setShadow(bool shadow = true);
    void setUnderLine(WD_UNDERLINE underline);
    void setFont(const QString &strFont);
    void setSize(WD_FONTSIZE &size);
    void setColor(QColor &color);
    
    void loadExistStyle();

private:
    void addOrAssignStyle();
    QString convertRGB16HexStr(QColor &color);
    void addOrAssignStyleChildElement(const QString &eleName, bool enable);

private:
    Run *m_run;
    QDomElement m_style;
    QDomElement m_rStyle;
};

}// namespace Docx

#endif// TEXT_H
