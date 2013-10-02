#ifndef KEXCELREADER_H
#define KEXCELREADER_H

#include <QObject>
#include <QVariantList>

class QAxObject;
class KExcelReader : public QObject
{
    Q_OBJECT

public:
    KExcelReader(QObject *parent=0);
    ~KExcelReader();

public:
    static bool isExcelApplicationAvailable();

public:
    bool open(const QString& xlsFile);
    int sheetCount() const;
    int rowCount(QAxObject* sheet) const;
    int columnCount(QAxObject* sheet) const;
    QList<QVariantList> values(const int colsCount, const int rowsCount=-1, const int sheetNumber=1) const;

private:
    void init();

private:
    QAxObject *m_excel;
    QAxObject *m_workbooks;
    QAxObject *m_workbook;
    QAxObject *m_sheets;
};

#endif // KEXCELREADER_H
