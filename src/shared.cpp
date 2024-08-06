#include "shared.h"

#include <QBuffer>
#include <QCryptographicHash>
#include <QDebug>
#include <QFile>

namespace Docx {

Parented::Parented() {}

Parented *Parented::part() { return nullptr; }

Parented::~Parented() {}

QDomElement addOrAssignElement(QDomDocument *dom, QDomElement *parent, const QString &eleName,
                               bool addToFirst)
{
    QDomNodeList elements = parent->elementsByTagName(eleName);

    if (elements.isEmpty())
    {
        QDomElement ele = dom->createElement(eleName);
        if (addToFirst)
        {
            QDomNode firstnode = parent->firstChild();
            parent->insertBefore(ele, firstnode);
        }
        else { parent->appendChild(ele); }
        return ele;
    }
    else
        return elements.at(0).toElement();
}

QByteArray getFileHash(const QString &fileName)
{
    QByteArray byteArray;
    QFile file(fileName);
    if (!file.exists()) return byteArray;
    if (file.open(QIODevice::ReadOnly))
    {
        byteArray = QCryptographicHash::hash(file.readAll(), QCryptographicHash::Md5);
        file.close();
    }

    return byteArray;
}

QByteArray imageHash(const QImage &img)
{
    QByteArray ba;
    QBuffer buffer(&ba);
    buffer.open(QIODevice::WriteOnly);
    img.save(&buffer);
    return QCryptographicHash::hash(ba, QCryptographicHash::Md5);
}

QByteArray byteHash(const QByteArray &bytes)
{
    return QCryptographicHash::hash(bytes, QCryptographicHash::Md5);
}

InvalidSpanError::InvalidSpanError(const QString &errorStr) : m_error(errorStr)
{
    qDebug() << m_error;
}

void InvalidSpanError::raise() const
{
    qDebug() << m_error;
    throw *this;
}

QException *InvalidSpanError::clone() const { return new InvalidSpanError(*this); }

}// namespace Docx
