#ifndef COREPROPERTIESPART_H
#define COREPROPERTIESPART_H

#include <QByteArray>
#include <QString>

#include "../../package.h"
#include "../part.h"

namespace Docx {

class CorePropertiesPart : public XmlPart
{
public:
    CorePropertiesPart(const QString &partName, const QString &contentType, const QByteArray &blob,
                       Package *package = nullptr);
    static void load(const QString &partName, const QString &contentType, const QByteArray &blob,
                     Package *package = nullptr);
    virtual ~CorePropertiesPart();

    void afterUnmarshal();
};

}// namespace Docx

#endif// COREPROPERTIESPART_H
