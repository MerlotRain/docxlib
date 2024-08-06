#include "documentpart.h"
#include "../opc/oxml.h"
#include "../oxml/oxmlshape.h"
#include "../package.h"
#include "shape.h"
#include "table.h"
#include "text.h"
#include "imagepart.h"
#include <QDebug>
#include <QtXml>

namespace Docx {

DocumentPart::DocumentPart(const QString &partName, const QString &contentType,
                           const QByteArray &blob, Package *package)
    : Part(partName, contentType, QByteArray(), package)
{
    m_dom = new QDomDocument();
    m_dom->setContent(blob);
    m_inlineShapes = new InlineShapes(this);
}

Paragraph *DocumentPart::addParagraph(const QString &text, const QString &style)
{
    QDomNode n = lastSectionPtr();//nodes.at(nodes.count() - 1);
    QDomNode parentNode = n.parentNode();

    QDomElement pEle = m_dom->createElement(QStringLiteral("w:p"));
    Paragraph *p = new Paragraph(this, pEle);
    if (!text.isEmpty()) p->addRun(text, style);

    parentNode.insertBefore(pEle, n);
    m_addParagraphs.append(p);
    return p;
}

DocumentPart *DocumentPart::load(const PackURI &partName, const QString &contentType,
                                 const QByteArray &blob, Package *package)
{
    return new DocumentPart(partName, contentType, blob, package);
}

Table *DocumentPart::addTable(int rows, int cols, const QString &style)
{
    QDomNode n = lastSectionPtr();
    QDomNode parentNode = n.parentNode();

    QDomElement pEle = m_dom->createElement(QStringLiteral("w:tbl"));
    Table *table = new Table(this, pEle);

    parentNode.insertBefore(pEle, n);

    for (int i = 0; i < cols; i++) { table->addColumn(); }
    for (int i = 0; i < rows; i++) { table->addRow(); }
    table->setStyle(style);
    m_tables.append(table);
    return table;
}

void DocumentPart::addCloneTable(Table *pCloneTable, Paragraph *graph)
{
    QDomNode n = lastSectionPtr();
    QDomNode parentNode = n.parentNode();
    if (!graph) { parentNode.insertBefore(pCloneTable->getElement(), n); }
    else { parentNode.insertAfter(pCloneTable->getElement(), graph->currentNode()); }
    m_tables.append(pCloneTable);
}

void DocumentPart::afterUnmarshal() { qDebug() << "afetrUnmarshal"; }

QDomDocument *DocumentPart::element() const { return m_dom; }

QByteArray DocumentPart::blob() const { return m_dom->toByteArray(); }

QPair<ImagePart *, QString> DocumentPart::getOrAddImagePart(const PackURI &imageDescriptor)
{
    ImageParts *imgs = m_package->imageparts();
    ImagePart *imagPart = imgs->getOrAddImagePart(imageDescriptor);

    return getOrAddImagePart(imagPart);
}

QPair<ImagePart *, QString> DocumentPart::getOrAddImagePart(const QImage &img)
{
    ImageParts *imgs = m_package->imageparts();
    ImagePart *imagPart = imgs->getOrAddImagePart(img);

    return getOrAddImagePart(imagPart);
}

QPair<ImagePart *, QString> DocumentPart::getOrAddImagePart(ImagePart *imagPart)
{
    QString rId = relateTo(imagPart, Constants::IMAGE, QStringLiteral("word/"));
    return QPair<ImagePart *, QString>(imagPart, rId);
}

QList<Paragraph *> DocumentPart::paragraphs()
{
    qDeleteAll(m_ps);
    m_ps.clear();
    QDomNodeList pElements = m_dom->elementsByTagName(QStringLiteral("w:p"));
    if (pElements.isEmpty()) return m_ps;

    for (int i = 0; i < pElements.count(); i++)
    {
        QDomElement pele = pElements.at(i).toElement();
        Paragraph *p = new Paragraph(this, pele);
        m_ps.append(p);
    }
    return m_ps;
}

QList<Table *> DocumentPart::tables()
{
    qDeleteAll(m_tables);
    m_tables.clear();
    QDomNodeList pElements = m_dom->elementsByTagName(QStringLiteral("w:tbl"));
    if (pElements.isEmpty()) return m_tables;

    for (int i = 0; i < pElements.count(); i++)
    {
        QDomElement tblEle = pElements.at(i).toElement();
        Table *p = new Table(this, tblEle);
        m_tables.append(p);
    }
    return m_tables;
}

DocumentPart::~DocumentPart()
{
    delete m_inlineShapes;
    delete m_dom;
    qDeleteAll(m_addParagraphs);
    m_addParagraphs.clear();

    qDeleteAll(m_ps);
    m_ps.clear();

    qDeleteAll(m_tables);
    m_tables.clear();
}

int DocumentPart::nextId()
{
    QDomNodeList elements = m_dom->childNodes();
    QList<QString> numbers;
    findAttributes(elements, QStringLiteral("id"), &numbers);

    int size = numbers.count() + 2;
    for (int i = 1; i < size; i++)
    {
        if (!numbers.contains(QString::number(i))) return i;
    }
    return 1;
}

void DocumentPart::findAttributes(const QDomNodeList &elements, const QString &attr,
                                  QList<QString> *nums)
{
    for (int i = 0; i < elements.count(); i++)
    {
        QString num = elements.at(i).toElement().attribute(attr);
        if (!nums->contains(num)) nums->append(num);
        findAttributes(elements.at(i).childNodes(), attr, nums);
    }
}

QDomNode DocumentPart::lastSectionPtr() const
{
    QDomNodeList nodes = m_dom->elementsByTagName(QStringLiteral("w:sectPr"));

    QDomNode n = nodes.at(nodes.count() - 1);
    return n;
}

Paragraph *DocumentPart::addTextByBookmark(const QString &strBookMark, const QString &strText,
                                           WD_FONTSIZE fz, WD_PARAGRAPH_ALIGNMENT align,
                                           const QString &font)
{
    QDomNodeList nodeLst = m_dom->elementsByTagName(QStringLiteral("w:bookmarkStart"));
    int nCount = nodeLst.count();
    for (int i = 0; i < nCount; i++)
    {
        QDomNode nodeMark = nodeLst.at(i);
        QString strName = nodeMark.toElement().attribute(QStringLiteral("w:name"));
        if (strName == strBookMark)
        {
            Paragraph *pGraph = getParagraphByNode(nodeMark.parentNode());
            if (pGraph)
            {
                pGraph->addRun(nodeMark.nextSibling().toElement(), strText);
                pGraph->setAlignment(align);
                pGraph->m_pCurrentRun->setFont(font);
                pGraph->m_pCurrentRun->setSize(fz);
                return pGraph;
            }
        }
    }
    return nullptr;
}

Paragraph *DocumentPart::getParagraphByNode(QDomNode findNode)
{
    QList<Paragraph *> paraLst = paragraphs();
    long nSize = paraLst.count();
    for (long j = 0; j < nSize; j++)
    {
        Paragraph *pGraph = paraLst.at(j);
        QDomNode graphNode = pGraph->currentNode();
        if (findNode == graphNode) { return pGraph; }
    }
    return nullptr;
}

Paragraph *DocumentPart::getNextParagraphByNode(QDomNode &findNode)
{
    QList<Paragraph *> paraLst = paragraphs();
    long nSize = paraLst.count();
    for (long j = 0; j < nSize; j++)
    {
        Paragraph *pGraph = paraLst.at(j);
        QDomNode graphNode = pGraph->currentNode();
        if (findNode == graphNode)
        {
            if (j + 1 < nSize) { return paraLst.at(j + 1); }
        }
    }
    return nullptr;
}

Paragraph *DocumentPart::addPictureByBookmark(const QString &strBookMark, const QString &strPath,
                                              const Length &width, const Length &height)
{
    QDomNodeList nodeLst = m_dom->elementsByTagName(QStringLiteral("w:bookmarkStart"));
    int nCount = nodeLst.count();
    for (int i = 0; i < nCount; i++)
    {
        QDomNode nodeMark = nodeLst.at(i);
        QString strName = nodeMark.toElement().attribute(QStringLiteral("w:name"));
        if (strName == strBookMark)
        {
            Paragraph *pGraph = getParagraphByNode(nodeMark.parentNode());
            if (pGraph) { pGraph->addRun()->addPicture(strPath, width, height); }
        }
    }
    return nullptr;
}

Paragraph *DocumentPart::currentParagraph()
{
    QList<Paragraph *> graphList = paragraphs();
    foreach (auto graph, graphList)
    {
        QDomNode node = graph->currentNode();
        if (node.isNull()) continue;
        QDomNodeList nodeLst =
                node.ownerDocument().elementsByTagName(QStringLiteral("w:bookmarkStart"));
        int nCount = nodeLst.count();
        for (int i = 0; i < nCount; i++)
        {
            QDomNode nodeMark = nodeLst.at(i);
            QString strName = nodeMark.toElement().attribute(QStringLiteral("w:name"));
            if (strName == "_GoBack")
            {
                Paragraph *pGraph = getParagraphByNode(nodeMark.parentNode());
                return pGraph;
            }
        }
    }

    return nullptr;
}

void DocumentPart::setCurrentParagraph(Paragraph *pGraph)
{
    int nCount;
    QList<Paragraph *> graphList = paragraphs();
    foreach (auto graph, graphList)
    {
        QDomNode node = graph->currentNode();
        if (node.isNull()) continue;
        QDomNodeList nodeLst =
                node.ownerDocument().elementsByTagName(QStringLiteral("w:bookmarkStart"));
        nCount = nodeLst.count();
        for (int i = 0; i < nCount; i++)
        {
            QDomNode nodeMark = nodeLst.at(i);
            QString strName = nodeMark.toElement().attribute(QStringLiteral("w:name"));
            if (strName == "_GoBack") { nodeMark.parentNode().removeChild(nodeMark); }
        }
    }

    addBookmark(pGraph, "_GoBack");
}

void DocumentPart::addBookmark(Paragraph *pParaph, const QString &strBookmark)
{
    if (!pParaph || strBookmark.trimmed().isEmpty()) return;

    QDomNode node = pParaph->currentNode();
    if (node.isNull()) return;

    QDomNodeList nodeLst =
            node.ownerDocument().elementsByTagName(QStringLiteral("w:bookmarkStart"));
    int nCount = nodeLst.count();

    QDomElement pEle = m_dom->createElement(QStringLiteral("w:bookmarkStart"));
    pEle.setAttribute(QStringLiteral("w:id"), nCount);
    pEle.setAttribute(QStringLiteral("w:name"), strBookmark);
    node.appendChild(pEle);

    pEle = m_dom->createElement(QStringLiteral("w:bookmarkEnd"));
    pEle.setAttribute(QStringLiteral("w:id"), nCount);
    node.appendChild(pEle);
}

Table *DocumentPart::getTableByBookmark(const QString &strBookMark)
{
    int nSize = m_tables.size();
    if (nSize == 0) { nSize = tables().size(); }
    for (int i = 0; i < nSize; i++)
    {
        Table *pTable = m_tables.at(i);
        if (pTable)
        {
            QList<QString> lstMarks = pTable->getAllBookmarks();
            if (lstMarks.contains(strBookMark)) { return pTable; }
        }
    }

    return NULL;
}


void DocumentPart::setHeader(const QString &strText)
{
    // TODO
    //QDomNode n = lastSectionPtr();//nodes.at(nodes.count() - 1);

    //QDomElement pEle = m_dom->createElement(QStringLiteral("w:hdr"));
    //n.appendChild(pEle);

    //pEle.setAttribute(QStringLiteral("w:type"), QStringLiteral("odd"));

    //QDomElement pPEle = m_dom->createElement(QStringLiteral("w:p"));
    //pEle.appendChild(pPEle);

    //QDomElement pPrEle = m_dom->createElement(QStringLiteral("w:pPr"));
    //pPEle.appendChild(pPrEle);

    //QDomElement pStyleEle = m_dom->createElement(QStringLiteral("w:pStyle"));
    //pPrEle.appendChild(pStyleEle);
    //pStyleEle.setAttribute(QStringLiteral("w:val"), QStringLiteral("Header"));

    //QDomElement prEle = m_dom->createElement(QStringLiteral("w:r"));
    //pPEle.appendChild(prEle);

    //QDomElement ptEle = m_dom->createElement(QStringLiteral("w:t"));
    //prEle.appendChild(ptEle);
    //ptEle.appendChild(m_dom->createTextNode(strText));
}


InlineShapes::InlineShapes(DocumentPart *part) : m_part(part) { m_dom = part->m_dom; }

InlineShapes::~InlineShapes() {}

InlineShape *InlineShapes::addPicture(const QString &imagePath, Run *run)
{

    QPair<ImagePart *, QString> imagePartAndId = m_part->getOrAddImagePart(imagePath);
    return addPicture(imagePartAndId, run);
}

InlineShape *InlineShapes::addPicture(const QImage &img, Run *run)
{
    QPair<ImagePart *, QString> imagePartAndId = m_part->getOrAddImagePart(img);
    return addPicture(imagePartAndId, run);
}

InlineShape *InlineShapes::addPicture(const QPair<ImagePart *, QString> &imagePartAndId, Run *run)
{
    int shapeId = m_part->nextId();
    CT_Picture *pic =
            new CT_Picture(m_part->m_dom, QStringLiteral("0"), imagePartAndId.first->fileName(),
                           imagePartAndId.second, imagePartAndId.first->defaultCx().emu(),
                           imagePartAndId.first->defaultCy().emu());

    CT_Inline *lnLine =
            new CT_Inline(m_part->m_dom, imagePartAndId.first->defaultCx().emu(),
                          imagePartAndId.first->defaultCy().emu(), QString::number(shapeId), pic);
    run->addDrawing(lnLine);
    InlineShape *picture = new InlineShape(lnLine);
    return picture;
}

}// namespace Docx
