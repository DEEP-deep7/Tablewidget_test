#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QTableWidget>
#include <QString>
#include <QFileDialog>
#include <ActiveQt/QAxObject>
#include <QMessageBox>
#include <QDesktopServices>
#include <QPushButton>
#include <QDir>
#include <QVariant>


namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);

    void test_excel(QTableWidget& table);

    void test_excel02(QTableWidget* table);
    ~MainWindow();

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
