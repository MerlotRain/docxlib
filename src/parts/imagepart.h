#ifndef IMAGEPART_H
#define IMAGEPART_H

#include "../image/image.h"
#include "length.h"
#include "../opc/part.h"

namespace Docx {
class ImagePart : public Part
{
public:
    ImagePart(const PackURI &partName, const QString &contentType,
              const QByteArray &blob = QByteArray(), Image *image = nullptr,
              const QByteArray &hash = QByteArray());
    QString fileName() const;
    Length defaultCx() const;
    Length defaultCy() const;
    static ImagePart *load(const PackURI &partName, const QString &contentType,
                           const QByteArray &blob = QByteArray(), Image *image = nullptr);
    static ImagePart *fromImage(const PackURI &partName, Image *image,
                                const QByteArray &hash = QByteArray());
    void afterUnmarshal();
    Image *image() const;
    QByteArray blob() const;
    QByteArray hash() const;

    ~ImagePart();

private:
    Image *m_image;
    QByteArray m_hash;
};

}// namespace Docx

#endif// IMAGEPART_H
