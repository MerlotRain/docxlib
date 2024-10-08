#include "physpkgreader.h"
#include "constants.h"
#include "packuri.h"

#include <QByteArray>
#include <private/qzipreader_p.h>

namespace Docx {


PhysPkgReader::PhysPkgReader(const QString &filePath) : m_reader(new QZipReader(filePath)) {}

PhysPkgReader::PhysPkgReader(QIODevice *device) : m_reader(new QZipReader(device)) {}

QByteArray PhysPkgReader::contentTypesData() { return blobForm(Constants::CONTENT_TYPES_URI); }

QByteArray PhysPkgReader::blobForm(const QString &packuri) { return m_reader->fileData(packuri); }

QByteArray PhysPkgReader::relsFrom(const PackURI &sourceUri)
{
    QString relstr = sourceUri.relsUriStr();
    return blobForm(relstr);
}

PhysPkgReader::~PhysPkgReader() {}

}// namespace Docx
