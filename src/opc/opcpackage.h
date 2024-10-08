#ifndef OPCPACKAGE_H
#define OPCPACKAGE_H

#include <QMap>
#include <QString>

namespace Docx {

class PackageReader;
class Package;
class Part;
class DocumentPart;
class Relationships;
class OpcPackage
{
public:
    OpcPackage();
    DocumentPart *mainDocument();
    Part *partByRelated(const QString &reltype);
    QList<Part *> parts() const;
    virtual ~OpcPackage();

protected:
    void partsbyRels(const Relationships *rels, QList<Part *> *parts) const;

protected:
    Relationships *m_rels;
};

class Unmarshaller
{
public:
    Unmarshaller();
    static void unmarshal(PackageReader *pkgReader, Package *package);
    static QMap<QString, Part *> unmarshalParts(PackageReader *pkgReader, Package *package);
    static void unmarshalRelationships(PackageReader *pkgReader, Package *package,
                                       const QMap<QString, Part *> &parts);
    virtual ~Unmarshaller();
};
}// namespace Docx
#endif// OPCPACKAGE_H
