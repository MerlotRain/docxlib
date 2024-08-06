#include "xmlchemy.h"

namespace Docx {

/* ---------------------------- BaseChildElement ---------------------------- */

BaseChildElement::BaseChildElement(const QString &nsptagname, const QString &successors)
    : m_nsptagname(nsptagname), m_successors(successors)
{
}

BaseChildElement::~BaseChildElement() {}


ZeroOrOne::ZeroOrOne(const QString &nsptagname, const QString &successors)
    : BaseChildElement(nsptagname, successors)
{
}

ZeroOrOne::~ZeroOrOne() {}

/* ----------------------------- OxmlElementBase ---------------------------- */

OxmlElementBase::OxmlElementBase() {}

OxmlElementBase::OxmlElementBase(QDomElement *x) : m_element(x) {}

void OxmlElementBase::insertElementBefore(QDomElement *elm, const QString &tagname) {}

OxmlElementBase::~OxmlElementBase() {}

}// namespace Docx
