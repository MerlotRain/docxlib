#include "physpkgwriter.h"

#include <QIODevice>
#include <private/qzipwriter_p.h>

namespace Docx {

PhysPkgWriter::PhysPkgWriter(const QString &filePath)
{
    m_writer = new QZipWriter(filePath, QIODevice::WriteOnly);
    m_writer->setCompressionPolicy(QZipWriter::AutoCompress);
}

PhysPkgWriter::PhysPkgWriter(QIODevice *device)
{
    m_writer = new QZipWriter(device);
    m_writer->setCompressionPolicy(QZipWriter::AutoCompress);
}

void PhysPkgWriter::write(const QString &packUri, const QByteArray &blob)
{
    if (m_paths.contains(packUri)) return;
    m_paths.append(packUri);
    m_writer->addFile(packUri, blob);
}

void PhysPkgWriter::close() { m_writer->close(); }

PhysPkgWriter::~PhysPkgWriter() { delete m_writer; }

}// namespace Docx
