#ifndef TABLE_H
#define TABLE_H

#include "docx_global.h"
#include "enumtext.h"
#include "length.h"
#include <QDomDocument>
#include <QSharedPointer>
#include <QString>
#include <QList>

namespace Docx {

class DocumentPart;
class Paragraph;
class Column;
class Row;
class Cell;
class CT_Tbl;
class CT_Tc;

class DOCX_EXPORT Table
{
public:
    Table(){};
    Table(DocumentPart *part, const QDomElement &element);
    ~Table();

    Cell *cell(int rowIndex, int colIndex);
    Row *addRow();
    Column *addColumn();
    QList<Cell *> rowCells(int rowIndex);
    QList<Row *> rows();
    void setStyle(const QString &style);
    void setAlignment(WD_TABLE_ALIGNMENT alignment);

    Row *addRowByPosition(int nRowPosition, bool bClearText = true);
    void merge(int nBeginRow, int nBeginCol, int nEndRow, int nEndCol);
    void setBorders(const QDomElement &element);
    void loadExistRowElement();
    QList<QString> getAllBookmarks();
    void clearBookmark();
    Table *clone()
    {
        QDomElement pEle = m_element.cloneNode().toElement();
        Table *table = new Table(m_part, pEle);
        if (table) { table->clearBookmark(); }

        return table;
    }

    Cell *moveLeft(const QString &strBookmark, int nMove);
    Cell *moveRight(const QString &strBookmark, int nMove);
    Cell *moveUp(const QString &strBookmark, int nMove);
    Cell *moveDown(const QString &strBookmark, int nMove);
    Cell *movePosition(const QString &strBookmark, int nMoveRow, int nMoveCol);
    void getRowColByNode(QDomNode node, int &nRow, int &nCol);

    QDomElement getElement() { return m_element; }
    QDomNode getNodeByBookMark(const QString &strBookmark);

private:
    QList<Row *> m_rows;
    DocumentPart *m_part = nullptr;
    QDomDocument *m_dom = nullptr;
    CT_Tbl *m_ctTbl = nullptr;

    QDomElement m_element;
    friend class Row;
    friend class CT_Tbl;
};

class DOCX_EXPORT Cell
{
public:
    Cell(const QDomElement &element, Row *row);
    ~Cell();
    Paragraph *addParagraph(const QString &text = QString(), const QString &style = QString());
    void addText(const QString &text);
    void setAlignment(WD_PARAGRAPH_ALIGNMENT align);
    void setFont(const QString &strFont);
    Table *addTable(int rows, int cols, const QString &style = QString::fromLatin1("TableGrid"));
    Cell *merge(Cell *other);
    int cellIndex();
    int rowIndex();
    Table *table();

    QString text();
    void setText(const QString &text);
    QDomElement getElement() { return m_ele; }
    Paragraph *getCurParagraph() { return m_currentPara; }

    int span();


private:
    QDomDocument *m_dom;
    DocumentPart *m_part;
    QDomElement m_ele;
    QList<Paragraph *> m_paras;
    Paragraph *m_currentPara;
    Row *m_row;
    QSharedPointer<CT_Tc> m_tc;
    friend class CT_Tc;
    friend class Row;
};


class DOCX_EXPORT Column
{
public:
    Column(const QDomElement &tlGrid, int gridIndex, Table *table);
    Length width() const;
    void setWidth();
    ~Column();

private:
    QDomElement m_grid;
    Table *m_table;
    int m_index;
};


class DOCX_EXPORT Row
{
public:
    Row(const QDomElement &element, Table *table);
    ~Row();

    void loadExistElement();
    Table *table() const;
    QList<Cell *> cells() const;
    Table *table();
    int rowIndex();

    Row *clone();

    QDomNode domNode() { return m_ele; }

private:
    void addTc();

private:
    QList<Cell *> m_cells;
    QDomElement m_ele;
    Table *m_table;
    DocumentPart *m_part;
    QDomDocument *m_dom;
    friend class Cell;
    friend class Table;
};

}// namespace Docx

#endif// TABLE_H
