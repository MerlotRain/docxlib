#ifndef IMAGE_H
#define IMAGE_H

#include "../opc/packuri.h"

#include <QByteArray>
#include <QImage>
#include <QString>

namespace Docx {

class Image
{
public:
    Image(const PackURI &path);
    Image(const QByteArray &data, const QString &format);
    QImage img() const;
    int pxWidth() const;
    int pxHeight() const;
    int horzDpi() const;
    int vertDpi() const;
    QString ext() const;
    QString contentType() const;
    QByteArray blob();
    ~Image();

private:
    void saveBlob();

private:
    QSize m_imgSize;
    QString m_ext;
    QString m_contentType;
    QImage m_img;
    QByteArray m_blob;
    int m_horzDpi;
    int m_vertDpi;
};
}// namespace Docx

#endif// IMAGE_H
