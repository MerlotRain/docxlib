#ifndef PACKAGEREADER_H
#define PACKAGEREADER_H

#include "serializedrelationships.h"

#include <QIODevice>
#include <QList>
#include <QMap>
#include <QString>

namespace Docx {

class PhysPkgReader;
class SerializedRelationships;
class ContentTypeMap;

class PackageReader
{
public:
    PackageReader(SerializedRelationships *srels, ContentTypeMap *contentTypes,
                  const QList<SerializedPart> &sparts);
    static PackageReader *fromFile(const QString &pkgFile);
    static PackageReader *fromFile(QIODevice *device);
    static SerializedRelationships *srelsFrom(PhysPkgReader *physReader, const QString &sourceUri);
    static QList<SerializedPart> loadSerializedParts(PhysPkgReader *physReader,
                                                     const SerializedRelationships *srels,
                                                     const ContentTypeMap *contentTypes);
    QMap<QString, QVector<SerializedRelationship>> partRels() const;
    ContentTypeMap *contentTypes() const { return m_contentTypes; }
    QList<SerializedPart> sparts() const { return m_sparts; }

    ~PackageReader();

private:
    SerializedRelationships *m_srels;
    ContentTypeMap *m_contentTypes;
    QList<SerializedPart> m_sparts;
};

}// namespace Docx

#endif// PACKAGEREADER_H
