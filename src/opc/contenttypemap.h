#ifndef CONTENTTYPEMAP_H
#define CONTENTTYPEMAP_H

#include "packuri.h"

#include <QByteArray>
#include <QMap>
#include <QString>

namespace Docx {

class ContentTypeMap
{
public:
    ContentTypeMap();
    static ContentTypeMap *fromXml(const QByteArray &contentTypesXml);
    void addDefault(const QString &extension, const QString &contentType);
    void addOverride(const QString &partName, const QString &contentType);
    void addContentType(const PackURI &name, const QString &contentType);
    QString contentType(const PackURI &partname) const;
    QByteArray blob();
    ~ContentTypeMap();

private:
    QMap<QString, QString> m_overrides;
    QMap<QString, QString> m_defaults;
};

}// namespace Docx

#endif// CONTENTTYPEMAP_H
