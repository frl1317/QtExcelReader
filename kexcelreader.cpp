#include "kexcelreader.h"

#include <QAxObject>
#include <QDebug>

KExcelReader::KExcelReader(QObject *parent)
    : QObject(parent)
    , m_excel(0)
    , m_workbooks(0)
    , m_workbook(0)
    , m_sheets(0)
{
    init();
}

KExcelReader::~KExcelReader()
{
    if ( m_workbooks && !m_workbooks->isNull() )
    {
        m_workbooks->dynamicCall("Close()");
    }

    if ( m_excel && !m_excel->isNull() )
    {
        m_excel->dynamicCall("Quit()");
    }
}

bool KExcelReader::isExcelApplicationAvailable()
{
    QAxObject* excel = new QAxObject("Excel.Application", 0);
    excel->deleteLater();

    return !excel->isNull();
}

void KExcelReader::init()
{
    m_excel = new QAxObject("Excel.Application", this);

    Q_ASSERT(m_excel && !m_excel->isNull() && "Excel is not installed");

    if (m_excel && !m_excel->isNull())
    {
        m_workbooks = m_excel->querySubObject("Workbooks");
    }
}

bool KExcelReader::open( const QString& xlsFile )
{
    Q_ASSERT(m_excel && m_workbooks);

    if (m_workbooks && !m_workbooks->isNull())
    {
        m_workbook = m_workbooks->querySubObject("Open(const QString&)", xlsFile);
    }

    if (m_workbook && !m_workbook->isNull())
    {
        m_sheets = m_workbook->querySubObject("Worksheets");
    }

    return m_sheets != 0;
}

int KExcelReader::sheetCount() const
{
    Q_ASSERT(m_sheets);

    if ( m_sheets == 0 || m_sheets->isNull())
    {
        return 0;
    }

    return m_sheets->dynamicCall("Count()").toInt();
}

int KExcelReader::rowCount( QAxObject* sheet ) const
{
    Q_ASSERT(sheet);

    if (sheet == 0 || sheet->isNull())
    {
        return 0;
    }

    QAxObject* rows = sheet->querySubObject("Rows");

    if (rows && !rows->isNull())
    {
        return rows->dynamicCall("Count()").toInt(); //always returns 255
    }

    return 0;
}

int KExcelReader::columnCount( QAxObject* sheet ) const
{
    Q_ASSERT(sheet);

    if (sheet == 0 || sheet->isNull())
    {
        return 0;
    }

    QAxObject* columns = sheet->querySubObject("Columns");
    if (columns && !columns->isNull())
    {
        return columns->property("Count").toInt(); //always returns 65535
    }

    return 0;
}

QList<QVariantList> KExcelReader::values(int colsCount, int rowsCount, int sheetNumber) const
{
    Q_ASSERT(sheetNumber > 0);
    Q_ASSERT(colsCount > 0);
    Q_ASSERT(rowsCount > -2 && rowsCount != 0);

    //Data list from excel, each QVariantList is worksheet row
    QList<QVariantList> data;

    //sheet pointer
    QAxObject* sheet = m_sheets->querySubObject("Item( int )", sheetNumber);

    // rowsCountToRead
    // rowsCount == -1 -> read all
    int actualRowCount = rowCount(sheet);
    int rowsCountToRead = (actualRowCount < rowsCount || rowsCount == -1) ? actualRowCount:rowsCount;

    // columnsCountToRead
    int actualColCount = columnCount(sheet);
    int colsCountToRead = actualColCount < colsCount ? actualColCount:colsCount;

    for (int row=1; row <= rowsCountToRead; row++)
    {
        QVariantList dataRow;
        bool isEmpty = true; //When all the cells of row are empty, 
                             //it means that file is at end (of course, it maybe not right for different excel files.
                             //It's just criteria to calculate somehow row count for my file)

        for (int column=1; column <= colsCountToRead; column++)
        {
            QAxObject* cell = sheet->querySubObject("Cells( int, int )", row, column);
            QVariant value = cell->dynamicCall("Value()");

            bool tmp = value.isNull();
            if (!value.toString().isEmpty() && isEmpty)
            {
                isEmpty = false;
            }

            dataRow.append(value);
        }

        if (isEmpty)
        {
            return data;
        }
        else
        {
            data.push_back(dataRow);
        }
    }

    return data;
}
