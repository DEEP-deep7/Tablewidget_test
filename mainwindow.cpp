#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QDebug>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    QObject::connect(ui->pushButton,QPushButton::clicked,ui->tableWidget,[=](){

       test_excel(*(ui->tableWidget));
    });


}

void MainWindow::test_excel(QTableWidget &table)
{

    QString filename = QFileDialog::getSaveFileName(this,tr("Excle file"),QString("./data.xls"),tr("Excle Files(*.xel)"));
    qDebug()<<filename;
    if(filename != "")
    {
        qDebug()<<"进入逻辑";
        ui->progressBar->show();
        ui->progressBar->setValue(0);//进度条为0
        QAxObject *excel = new QAxObject;
        if(excel->setControl("Excel.Application"))
        {
            qDebug()<<"setcontrol";
            excel->dynamicCall("SetVisible (bool Visible)","false");
            excel->setProperty("DisplayAlerts",false);

            QAxObject *workbooks = excel->querySubObject("WorkBooks");          //获取工作簿集合
            qDebug()<<"获取工作集合完成";
            workbooks->dynamicCall("Add");                                      //创建一个工作簿
            qDebug()<<"创建工作簿完成";
            QAxObject *workbood = workbooks->querySubObject("ActiveWorkBook"); //获取当前工作簿
            qDebug()<<"获取当前工作簿完成";
            //QAxObject *worksheets = workbood->querySubObject("sheets");
            QAxObject *worksheet = workbood->querySubObject("Worksheets(int)",1);
            qDebug()<<"initover";

            QAxObject *cell;
            /*添加excel表头数据*/
            for(int i=1;i<=table.columnCount();i++)
            {
                cell=worksheet->querySubObject("Cells(int,int)",1,i);//???
                cell->setProperty("RowHeight",40);
                cell->dynamicCall("SetValue(const QString)",table.horizontalHeaderItem(i-1)->data(0).toString());
                if(ui->progressBar->value()<=50)
                {
                    ui->progressBar->setValue(10+i*5);
                }
                qDebug()<<"表头"<<i;
            }

            /*将form列表中的数据依次保存到Excel文件中*/
            for(int j=2;j<table.rowCount()+1;j++)
            {
                for(int k=1;k<table.columnCount();k++)
                {
                    cell=worksheet->querySubObject("Cells(int,int)",j,k);
                    cell->dynamicCall("SetValue(const QString&)",table.item(j-2,k-1)->text()+"\t");//m每行数据用table分隔
                }
                if(ui->progressBar->value()<80)
                {
                    ui->progressBar->setValue(50+j*5);
                }
                qDebug()<<"保存"<<j;
            }
            //将生成的Excel文件保存到指定目录下
            workbood->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filename));//保存到filename
            workbood->dynamicCall("Close()");
            excel->dynamicCall("Quit()");
            delete excel;
            excel=NULL;
            qDebug()<<"ok";
            ui->progressBar->setValue(100);
            if(QMessageBox::question(NULL,QString::fromUtf8("完成"),QString::fromUtf8("文件已导出，是否打开？"),QMessageBox::Yes|QMessageBox::No)==QMessageBox::Yes)
            {
                QDesktopServices::openUrl(QUrl("file:///"+QDir::toNativeSeparators(filename)));
            }
            ui->progressBar->setValue(0);
            ui->progressBar->hide();
        }

    }
}

void MainWindow::test_excel02(QTableWidget *table)
{

}



MainWindow::~MainWindow()
{
    delete ui;
}
