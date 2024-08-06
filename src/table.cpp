#include "table.h"
#include "parts/documentpart.h"
#include "text.h"
#include "oxml/oxmltable.h"

#include <QDebug>
#include <QVector>

namespace Docx {


const QString strtblRow = QStringLiteral("w:tr");
const QString strtblCell = QStringLiteral("w:tc");
const QString strBorders = QStringLiteral("w:tblBorders");
const QString strstyle = QStringLiteral("w:tblPr");


Table::Table(DocumentPart *part, const QDomElement &element) : m_part(part)
{
    m_dom = part->m_dom;
    m_element = element;
    m_ctTbl = new CT_Tbl(this, element);

    setBorders(element);
    loadExistRowElement();
}

//<w:tblBorders>
//	<w:top w : val = "single" w : sz = "4" w : space = "0" w : color = "auto" / >
//	<w:left w : val = "single" w : sz = "4" w : space = "0" w : color = "auto" / >
//	<w:bottom w : val = "single" w : sz = "4" w : space = "0" w : color = "auto" / >
//	<w:right w : val = "single" w : sz = "4" w : space = "0" w : color = "auto" / >
//	<w:insideH w : val = "single" w : sz = "4" w : space = "0" w : color = "auto" / >
//	<w:insideV w : val = "single" w : sz = "4" w : space = "0" w : color = "auto" / >
//< / w:tblBorders>
void Table::setBorders(const QDomElement &element)
{
    // tblPr
    QDomNodeList tblPrEle = element.elementsByTagName(strstyle);
    if (tblPrEle.isEmpty()) return;

    //get  tblPr
    QDomElement styleEle = tblPrEle.at(0).toElement();
    //create net node
    QDomElement BordersEle = m_dom->createElement(strBorders);
    styleEle.appendChild(BordersEle);
    for (int i = 0; i < 6; i++)
    {
        QString strVal;
        switch (i)
        {
            case 0:
                strVal = QStringLiteral("w:top");
                break;
            case 1:
                strVal = QStringLiteral("w:left");
                break;
            case 2:
                strVal = QStringLiteral("w:bottom");
                break;
            case 3:
                strVal = QStringLiteral("w:right");
                break;
            case 4:
                strVal = QStringLiteral("w:insideH");
                break;
            case 5:
                strVal = QStringLiteral("w:insideV");
                break;
            default:
                break;
        }
        QDomElement childEle = m_dom->createElement(strVal);
        BordersEle.appendChild(childEle);
        childEle.setAttribute(QStringLiteral("w:val"), QStringLiteral("single"));
        childEle.setAttribute(QStringLiteral("w:sz"), QStringLiteral("4"));
        childEle.setAttribute(QStringLiteral("w:space"), QStringLiteral("0"));
        childEle.setAttribute(QStringLiteral("w:color"), QStringLiteral("auto"));
    }
}

void Table::loadExistRowElement()
{

    QDomNodeList reles = m_ctTbl->m_tblEle.childNodes();
    if (!reles.isEmpty())
    {
        for (int i = 0; i < reles.count(); i++)
        {
            QDomElement rowEle = reles.at(i).toElement();
            if (rowEle.nodeName() != strtblRow) continue;
            Row *row = new Row(rowEle, this);
            m_rows.append(row);
            qDebug() << m_rows.count();
        }
    }
}

Cell *Table::cell(int rowIndex, int colIndex)
{
    if (rowIndex < 0 || colIndex < 0) return nullptr;

    if (m_rows.size() <= rowIndex) return nullptr;

    Row *row = m_rows.at(rowIndex);
    if (!row) return nullptr;
    if (row->cells().size() <= colIndex) return nullptr;

    return row->cells().at(colIndex);
}

Table::~Table() { delete m_ctTbl; }

Column *Table::addColumn()
{
    QDomElement gridCol = m_ctTbl->m_tblGrid->addGridCol();
    for (Row *row: m_rows) { row->addTc(); }
    return new Column(gridCol, m_ctTbl->m_tblGrid->count(), this);
}

Row *Table::addRow()
{
    QDomElement rowEle = m_dom->createElement(strtblRow);
    Row *row = new Row(rowEle, this);
    m_rows.append(row);
    for (int i = 0; i < m_ctTbl->m_tblGrid->count(); i++) { row->addTc(); }
    m_ctTbl->m_tblEle.appendChild(rowEle);
    return row;
}

Row *Table::addRowByPosition(int nRowPosition, bool bClearText)
{
    QDomNodeList lstNodes = m_ctTbl->m_tblEle.childNodes();
    int nSize = lstNodes.size();
    if (nSize <= nRowPosition || nRowPosition < 0) return nullptr;

    Row *pOldRow = m_rows.at(nRowPosition);
    if (!pOldRow) return nullptr;

    Row *pCloneRow = pOldRow->clone();
    if (!pCloneRow) return nullptr;
    m_rows.insert(nRowPosition, pCloneRow);

    if (bClearText)
    {
        for (int i = 0; i < pCloneRow->cells().size(); i++)
        {
            pCloneRow->cells().at(i)->setText(QString());
        }
    }
    m_ctTbl->m_tblEle.insertBefore(pCloneRow->domNode(), lstNodes.at(nRowPosition + 2));


    return pCloneRow;
}

QList<Cell *> Table::rowCells(int rowIndex)
{
    Row *row = m_rows.at(rowIndex);

    return row->cells();
}

QList<Row *> Table::rows() { return m_rows; }

void Table::setStyle(const QString &style) { m_ctTbl->setStyle(style); }

void Table::setAlignment(WD_TABLE_ALIGNMENT alignment) { m_ctTbl->setAlignment(alignment); }

void Table::clearBookmark()
{
    QDomNodeList nodeLst = m_element.elementsByTagName(QStringLiteral("w:bookmarkStart"));
    int nCount = nodeLst.count();
    for (int i = 0; i < nCount; i++)
    {
        QDomNode nodeMark = nodeLst.at(0);
        nodeMark.parentNode().removeChild(nodeMark);
    }

    nodeLst = m_element.elementsByTagName(QStringLiteral("w:bookmarkStart"));
    nCount = nodeLst.count();
}

QList<QString> Table::getAllBookmarks()
{
    QList<QString> lstBookMarks;
    QDomNodeList nodeLst = m_element.elementsByTagName(QStringLiteral("w:bookmarkStart"));
    int nCount = nodeLst.count();
    for (int i = 0; i < nCount; i++)
    {
        QDomNode nodeMark = nodeLst.at(i);
        QString strName = nodeMark.toElement().attribute(QStringLiteral("w:name"));
        lstBookMarks.append(strName);
    }

    return lstBookMarks;
}

Cell *Table::moveLeft(const QString &strBookmark, int nMove)
{
    int nRow = -1, nCol = -1;
    QDomNode nodeMark = getNodeByBookMark(strBookmark);
    getRowColByNode(nodeMark, nRow, nCol);
    if (-1 == nRow || -1 == nCol) return nullptr;

    int nCurColMove = 0;
    for (int i = 1; i <= nMove; i++)
    {
        Cell *pCell = nullptr;
        pCell = cell(nRow, nCol - nCurColMove - 1);

        if (pCell)
        {
            int nSpanCnt = pCell->span();
            if (nSpanCnt > 0) { nCurColMove += nSpanCnt; }
        }
    }

    nCol -= nCurColMove;

    return cell(nRow, nCol);
}

Cell *Table::moveRight(const QString &strBookmark, int nMove)
{
    int nRow = -1, nCol = -1;
    QDomNode nodeMark = getNodeByBookMark(strBookmark);
    getRowColByNode(nodeMark, nRow, nCol);
    if (-1 == nRow || -1 == nCol) return nullptr;

    int nCurColMove = 0;
    for (int i = 1; i <= nMove; i++)
    {
        Cell *pCell = nullptr;
        pCell = cell(nRow, nCol + nCurColMove + 1);

        if (pCell)
        {
            int nSpanCnt = pCell->span();
            if (nSpanCnt > 0) { nCurColMove += nSpanCnt; }
        }
    }

    nCol += nCurColMove;

    return cell(nRow, nCol);
}

Cell *Table::moveUp(const QString &strBookmark, int nMove)
{
    int nRow = -1, nCol = -1;
    QDomNode nodeMark = getNodeByBookMark(strBookmark);
    getRowColByNode(nodeMark, nRow, nCol);
    if (-1 == nRow || -1 == nCol) return nullptr;

    nRow -= nMove;

    return cell(nRow, nCol);
}

Cell *Table::moveDown(const QString &strBookmark, int nMove)
{
    int nRow = -1, nCol = -1;
    QDomNode nodeMark = getNodeByBookMark(strBookmark);
    getRowColByNode(nodeMark, nRow, nCol);
    if (-1 == nRow || -1 == nCol) return nullptr;

    nRow += nMove;

    return cell(nRow, nCol);
}


Cell *Table::movePosition(const QString &strBookmark, int nMoveRow, int nMoveCol)
{
    int nRow = -1, nCol = -1;
    QDomNode nodeMark = getNodeByBookMark(strBookmark);
    getRowColByNode(nodeMark, nRow, nCol);
    if (-1 == nRow || -1 == nCol) return nullptr;

    nRow += nMoveRow;
    nCol += nMoveCol;


    return cell(nRow, nCol);
}

QDomNode Table::getNodeByBookMark(const QString &strBookmark)
{
    QDomNodeList nodeLst =
            m_element.ownerDocument().elementsByTagName(QStringLiteral("w:bookmarkStart"));
    int nCount = nodeLst.count();
    for (int i = 0; i < nCount; i++)
    {
        QDomNode nodeMark = nodeLst.at(i);
        QString strName = nodeMark.toElement().attribute(QStringLiteral("w:name"));
        if (0 == strName.compare(strBookmark)) { return nodeMark; }
    }

    return QDomNode();
}

void Table::getRowColByNode(QDomNode node, int &nRow, int &nCol)
{
    nRow = nCol = -1;
    QDomNode parentColNode = node.parentNode().parentNode();
    QDomNode parentRowNode = parentColNode.parentNode();

    QList<Row *> lstRows = rows();
    foreach (auto pRow, lstRows)
    {
        QList<Cell *> lstCells = pRow->cells();
        nRow++;
        if (pRow->domNode() == parentRowNode) break;
    }
    if (-1 == nRow) return;

    QList<Cell *> lstCells = rowCells(nRow);
    foreach (auto pCell, lstCells)
    {
        nCol++;
        if (pCell->getElement() == parentColNode)
        {
            int nSpanCnt = pCell->span();
            if (nSpanCnt > 1) { nCol = nCol + nSpanCnt - 1; }
            break;
        }
    }
}

void Table::merge(int nBeginRow, int nBeginCol, int nEndRow, int nEndCol)
{
    Cell *pMergerCell = NULL;
    Cell *pBeforeCell = NULL;
    Cell *pCurCell = NULL;
    int nFirst = 0;

    int nRowCnt = rows().size();

    for (int i = nBeginRow; i <= nEndRow; i++)
    {
        for (int j = nBeginCol; j <= nEndCol; j++)
        {
            if (nFirst == 0) { pBeforeCell = cell(i, j); }
            else
            {
                pCurCell = cell(i, j);
                if (pCurCell)
                {
                    pMergerCell = pBeforeCell->merge(pCurCell);
                    pBeforeCell = pMergerCell;
                }
            }

            nFirst++;
        }
    }
}

Column::Column(const QDomElement &tlGrid, int gridIndex, Table *table)
    : m_grid(tlGrid), m_index(gridIndex), m_table(table)
{
}

Length Column::width() const { return Length(); }

void Column::setWidth() {}

Column::~Column() {}

Row::Row(const QDomElement &element, Table *table) : m_ele(element), m_table(table)
{
    m_dom = m_table->m_dom;
    m_part = m_table->m_part;

    loadExistElement();
}

void Row::loadExistElement()
{
    QDomNodeList eleList = m_ele.childNodes();
    if (eleList.isEmpty()) return;

    for (int i = 0; i < eleList.count(); i++)
    {
        QDomElement celEle = eleList.at(i).toElement();
        if (celEle.nodeName() != strtblCell) continue;

        Cell *cell = new Cell(celEle, this);
        m_cells.append(cell);

        QString strMerge = cell->m_tc->vMerge();
        if (!cell->m_tc->m_vMerge.isNull() && strMerge != QStringLiteral("restart"))
        {
            Row *uprow = m_table->m_rows.last();
            Cell *upcell = uprow->m_cells.at(m_cells.count() - 1);
            celEle = QDomElement(upcell->m_tc->ele());
            //
            m_cells.removeAll(cell);
            delete cell;
            cell = new Cell(celEle, this);
            m_cells.append(cell);
        }

        int spanCount = cell->m_tc->gridSpan();
        if (spanCount > 1)
            for (int i = 1; i < spanCount; i++)
            {
                Cell *hcell = new Cell(celEle, this);
                m_cells.append(hcell);
            }
    }
}

void Row::addTc()
{
    QDomElement celEle = m_dom->createElement(strtblCell);
    Cell *cell = new Cell(celEle, this);
    m_cells.append(cell);
    m_ele.appendChild(celEle);
}

Table *Row::table() const { return m_table; }

QList<Cell *> Row::cells() const { return m_cells; }

Table *Row::table() { return m_table; }

int Row::rowIndex() { return m_table->m_rows.indexOf(this); }

Row *Row::clone()
{
    QDomElement pEle = m_ele.cloneNode().toElement();
    Row *pRow = new Row(pEle, m_table);
    return pRow;
}

Row::~Row() {}

Cell::Cell(const QDomElement &element, Row *row) : m_row(row), m_ele(element)
{
    m_dom = row->m_dom;
    m_part = row->m_part;
    m_tc = QSharedPointer<CT_Tc>(new CT_Tc(this, element));
    if (!m_tc->m_isLoad) addParagraph();
    else
    {
        QDomNode node = element.lastChild();
        if (!node.isNull() && node.nodeName() == QStringLiteral("w:p"))
        {
            QDomElement pEle = node.toElement();
            m_currentPara = new Paragraph(m_part, pEle);
            m_paras.append(m_currentPara);
        }
    }
}

Paragraph *Cell::addParagraph(const QString &text, const QString &style)
{
    QDomElement pEle = m_dom->createElement(QStringLiteral("w:p"));

    m_currentPara = new Paragraph(m_part, pEle);

    if (!text.isEmpty()) m_currentPara->addRun(text, style);
    QDomNode node = m_tc->m_ele.lastChild();
    if (!node.isNull() && node.nodeName() == QStringLiteral("w:p"))
    {
        m_tc->m_ele.removeChild(node);
    }
    m_tc->m_ele.appendChild(pEle);
    m_paras.append(m_currentPara);
    return m_currentPara;
}

void Cell::addText(const QString &text) { m_currentPara->addRun(text); }

QString Cell::text()
{
    if (!m_currentPara) return QString();

    return m_currentPara->text();
}

void Cell::setText(const QString &text)
{
    if (!m_currentPara) return;

    return m_currentPara->setText(text);
}

int Cell::span()
{
    if (m_tc) { return m_tc->gridSpan(); }

    return 0;
}

void Cell::setAlignment(WD_PARAGRAPH_ALIGNMENT align) { m_currentPara->setAlignment(align); }

void Cell::setFont(const QString &strFont) { m_currentPara->currentRun()->setFont(strFont); }

Table *Cell::addTable(int rows, int cols, const QString &style)
{
    QDomElement pEle = m_dom->createElement(QStringLiteral("w:tbl"));
    Table *table = new Table(m_part, pEle);

    m_tc->m_ele.appendChild(pEle);

    for (int i = 0; i < cols; i++) { table->addColumn(); }
    for (int i = 0; i < rows; i++) { table->addRow(); }
    table->setStyle(style);
    addParagraph();
    return table;
}

Cell *Cell::merge(Cell *other)
{
    QSharedPointer<CT_Tc> tc = this->m_tc;
    QSharedPointer<CT_Tc> tc2 = other->m_tc;
    CT_Tc *tcReturn = tc->merge(tc2);
    return tcReturn->m_cell;
}

int Cell::cellIndex() { return m_row->m_cells.indexOf(this); }

int Cell::rowIndex() { return m_row->rowIndex(); }

Table *Cell::table() { return m_row->m_table; }

Cell::~Cell() { qDeleteAll(m_paras); }

}// namespace Docx
