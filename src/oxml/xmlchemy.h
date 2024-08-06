#ifndef XMLCHEMY_H
#define XMLCHEMY_H

#include <QDomElement>
#include <QString>

namespace Docx {

class BaseChildElement
{
public:
    BaseChildElement(const QString &nsptagname, const QString &successors = QString());

    virtual ~BaseChildElement();

private:
    QString m_nsptagname;
    QString m_successors;
};

class ZeroOrOne : public BaseChildElement
{
public:
    ZeroOrOne(const QString &nsptagname, const QString &successors = QString());

    virtual ~ZeroOrOne();

private:
};

class OxmlElementBase
{
public:
    OxmlElementBase();
    OxmlElementBase(QDomElement *x);
    void insertElementBefore(QDomElement *elm, const QString &tagname);

    virtual ~OxmlElementBase();

protected:
    QDomElement *m_element;
};

}// namespace Docx

#endif// XMLCHEMY_H
