#ifndef PHYSPKGREADER_H
#define PHYSPKGREADER_H

#include <QIODevice>
#include <QScopedPointer>
#include <QStringList>

class QByteArray;
class QZipReader;

namespace Docx {
class PackURI;
class PhysPkgReader
{
public:
    PhysPkgReader(const QString &filePath);
    PhysPkgReader(QIODevice *device);
    QByteArray contentTypesData();
    QByteArray blobForm(const QString &packuri);
    QByteArray relsFrom(const PackURI &sourceUri);

    ~PhysPkgReader();

private:
    QScopedPointer<QZipReader> m_reader;
};

}// namespace Docx

#endif// PHYSPKGREADER_H
