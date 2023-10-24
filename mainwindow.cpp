#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)  : QMainWindow(parent) , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::on_pushButton_clicked()
{
    QAxWidget excel("Excel.Application");
    excel.setProperty("Visible", false);

    QAxObject* workbooks=excel.querySubObject("WorkBooks");
    workbooks->dynamicCall("Add"); // Add new workbook
    QAxObject* workbook=excel.querySubObject("ActiveWorkBook");
    QAxObject* sheets=workbook->querySubObject("WorkSheets");

    QAxObject * cell;
    int count= sheets->dynamicCall("Count()").toInt();
    bool isEmpty = true;

    for(int k=1;k<=count;k++)
    {
        //sheet pointer
        QAxObject* sheet = sheets->querySubObject( "Item( int )", k );
        for(int i=1; i<11/*row*/ ; i++)
        {
            for(int j=1; j<7/*column*/ ;j++)
            {

                QString text= "ventor";//qrandom(); // value you want to export

                //get cell
                QAxObject* cell = sheet->querySubObject( "Cells( int, int )", i,j);
                //set your value to the cell of ith row n jth column
                cell->dynamicCall("SetValue(QString)",text);
                // if you wish check your value set correctly or not by retrieving and printing it
                QVariant value = cell->dynamicCall( "Value()" );

            }
        }
    }
    QString fileName=QFileDialog::getSaveFileName(0,"save file","export_table", "XML Spreadhseet(*.xlsx);;eXceL Spreadsheet(*.xls);;Comma Seperated Value(*.csv)");
    fileName.replace("/","\\");
    workbook->dynamicCall("SaveAs(QString&)",fileName);
    workbook->dynamicCall("Close (Boolean)",false);
}

