
import os
import sys
import datetime
from pathlib import Path
import sqlite3
from docxtpl import DocxTemplate

from home_panel import *
from PyQt6.QtWidgets import QMainWindow, QApplication, QMessageBox, QFileDialog
from PyQt6 import QtWidgets
from PyQt6.QtCore import QPropertyAnimation, QEasingCurve, QDir, QModelIndex
import pyodbc
server = 'mssql-157657-0.cloudclusters.net, 16555'
database = 'db_system'
username = 'root_db'
password = 'newUser09112001'
# Create a connection string
connection_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
connection = pyodbc.connect(connection_string)


class MainWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.ui = Ui_doc_panel()
        self.ui.setupUi(self)
        self.ui.stackedWidget.setCurrentWidget(self.ui.tbl_view)
        self.animation = QPropertyAnimation(self.ui.sidebar, b"minimumWidth")
        self.ui.btn_menu.clicked.connect(self.slideMenu)
        self.ui.btn_view.clicked.connect(lambda: self.ui.stackedWidget.setCurrentWidget(self.ui.tbl_view))
        self.ui.btn_accounts.clicked.connect(lambda: self.ui.stackedWidget.setCurrentWidget(self.ui.accounts))
        self.ui.btn_doccontents.clicked.connect(lambda: self.ui.stackedWidget.setCurrentWidget(self.ui.doc_contents))
        self.ui.btn_guide.clicked.connect(lambda: self.ui.stackedWidget.setCurrentWidget(self.ui.guide))
        self.ui.btn_prices.clicked.connect(lambda: self.ui.stackedWidget.setCurrentWidget(self.ui.doc_prices))
        self.ui.btn_tbledit.clicked.connect(lambda: self.ui.stackedWidget.setCurrentWidget(self.ui.edit_contents))
        self.ui.edit_accounts.clicked.connect(lambda: self.ui.stackedWidget.setCurrentWidget(self.ui.page))
        self.ui.user_delete.clicked.connect(self.deleteRow)
        self.ui.user_verify.clicked.connect(self.user_verif)
        self.ui.user_unverify.clicked.connect(self.user_unverif)
        self.ui.ref_table.clicked.connect(self.loadAccounts)
        self.ui.btn_filedir.clicked.connect(self.selectDir)
        self.ui.btn_create.clicked.connect(self.create_file)
        self.ui.btn_delete.clicked.connect(self.deleteRowDocs)
        self.ui.btn_editreq.clicked.connect(self.editReqDoc)
        self.ui.edit_row.clicked.connect(self.editRowDoc)
        self.ui.edit_user.clicked.connect(self.editUsers)
        self.ui.btn_refresh.clicked.connect(self.loadReq)
        self.ui.refresh_edittabel.clicked.connect(self.loadAcc)
        self.ui.ref_tabletransaction.clicked.connect(self.loadTransactions)
        self.ui.edit_doccontent.clicked.connect(self.editDocCont)
        self.loadTransactions()
        self.loadReq()
        self.loadAccounts()
        self.loadAcc()



    def editUsers(self):
        try:
            cur = connection.cursor()
            lemail = self.ui.edit_email.text()
            lepass = self.ui.edit_password.text()
            lename = self.ui.edit_uname.text()
            lesex = self.ui.edit_usex.text()
            leadd = self.ui.edit_uaddress()
            lebdate = self.ui.edit_ubday.text()
            lebplace = self.ui.edit_ubplace.text()
            lecs = self.ui.edit_ucstatus.text()
            lectzn = self.ui.edit_ucitizenship.text()


            query = "UPDATE tbl_users SET password='" + lepass + "' WHERE email ='" + lemail + "'"
            cur.execute(query)
            query = "UPDATE tbl_users SET name='" + lename + "' WHERE email ='" + lemail + "'"
            cur.execute(query)
            query = "UPDATE tbl_users SET sex='" + lesex + "' WHERE email ='" + lemail + "'"
            cur.execute(query)
            query = "UPDATE tbl_users SET address='" + leadd + "' WHERE email ='" + lemail + "'"
            cur.execute(query)
            query = "UPDATE tbl_users SET bday='" + lebdate + "' WHERE email ='" + lemail + "'"
            cur.execute(query)
            query = "UPDATE tbl_users SET bplace='" + lebplace + "' WHERE email ='" + lemail + "'"
            cur.execute(query)
            query = "UPDATE tbl_users SET cstatus='" + lecs + "' WHERE email ='" + lemail + "'"
            cur.execute(query)
            query = "UPDATE tbl_users SET citizenship='" + lectzn + "' WHERE email ='" + lemail + "'"
            cur.execute(query)
            connection.commit()
            msg = QMessageBox()
            msg.setWindowTitle("Notification")
            msg.setText("Successfully Edited! Refresh table to see the results.")
            msg.setIcon(QMessageBox.Icon.Information)
            msg.show()
            msg.exec()
        except Exception as e:
            print(str(e))
    def editRowDoc(self):
        try:
            cur = connection.cursor()
            uId = self.ui.line_editref.text()
            name = self.ui.line_editname.text()
            age = self.ui.line_editage.text()
            sex = self.ui.line_editsex.text()
            doct = self.ui.cBox_doctype.currentText()
            bplace = self.ui.line_editbirthplace.text()
            bdate = self.ui.line_editbirthdate.text()
            bname = self.ui.line_editbusinessname.text()
            baddress = self.ui.line_editbusinessaddress.text()
            ctzn = self.ui.line_editcitizenship.text()
            cstatus = self.ui.line_editcivilstatus.text()
            purpose = self.ui.cBox_purpose.currentText()
            dbquery = ("UPDATE doc_requests SET name='" + name + "', age='" + age + "', sex='" + sex + "', doc_req='" + doct + "', bday='" + bdate + "', bplace='"+
                       bplace + "', cstatus='" + cstatus + "', citizenship='" + ctzn + "', purpose='"+purpose+"', bname='"+bname+ "', baddress='"+baddress+"' WHERE reference = " + uId)
            cur.execute(dbquery)
            connection.commit()
            msg = QMessageBox()
            msg.setWindowTitle("Notification")
            msg.setText("Successfully Edited! Refresh to see the result")
            msg.setIcon(QMessageBox.Icon.Information)
            msg.show()
            msg.exec()
        except Exception as e:
            print(str(e))

    def loadAcc(self):
        c = connection.cursor()
        self.ui.tbl_accView.setRowCount(1000)
        c.execute("SELECT email, password, name, sex, address, bday, bplace, cstatus, citizenship FROM tbl_users")
        users = c.fetchall()
        tblrow = 0
        for row in users:
            self.ui.tbl_accView.setItem(tblrow, 0, QtWidgets.QTableWidgetItem(row[0]))
            self.ui.tbl_accView.setItem(tblrow, 1, QtWidgets.QTableWidgetItem(row[1]))
            self.ui.tbl_accView.setItem(tblrow, 2, QtWidgets.QTableWidgetItem(row[2]))
            self.ui.tbl_accView.setItem(tblrow, 3, QtWidgets.QTableWidgetItem(row[3]))
            self.ui.tbl_accView.setItem(tblrow, 4, QtWidgets.QTableWidgetItem(row[4]))
            self.ui.tbl_accView.setItem(tblrow, 5, QtWidgets.QTableWidgetItem(row[5]))
            self.ui.tbl_accView.setItem(tblrow, 6, QtWidgets.QTableWidgetItem(row[6]))
            self.ui.tbl_accView.setItem(tblrow, 7, QtWidgets.QTableWidgetItem(row[7]))
            self.ui.tbl_accView.setItem(tblrow, 8, QtWidgets.QTableWidgetItem(row[8]))
            tblrow += 1

        connection.commit()

    def slideMenu(self):
        width = self.ui.sidebar.width()
        if width == 0:
            newwidth = 260
        else:
            newwidth = 0

        self.animation.setDuration(250)
        self.animation.setStartValue(width)
        self.animation.setEndValue(newwidth)
        self.animation.setEasingCurve(QEasingCurve.Type.InOutQuart)
        self.animation.start()

    def editDocCont(self):
        print("test")
        try:

            cur = connection.cursor()
            bgy = self.ui.lineEdit_4.text()
            dbquery = "UPDATE barangay SET name = '"+bgy+"' WHERE position = year"
            cur.execute(dbquery)
            bgy2 = self.ui.line_captain.text()
            dbquery2 = "UPDATE barangay SET name = '" + bgy2 + "' WHERE position = captain"
            cur.execute(dbquery2)
            bgy3 = self.ui.line_ceb.text()
            dbquery3 = "UPDATE barangay SET name = '" + bgy3 + "' WHERE position = coeb"
            cur.execute(dbquery3)
            bgy4 = self.ui.line_cfcp.text()
            dbquery4 = "UPDATE barangay SET name = '" + bgy4 + "' WHERE position = cofcopw"
            cur.execute(dbquery4)
            bgy5 = self.ui.line_chsw.text()
            dbquery5 = "UPDATE barangay SET name = '" + bgy5 + "' WHERE position = cohsw"
            cur.execute(dbquery5)
            bgy6 = self.ui.line_livelihood.text()
            dbquery6 = "UPDATE barangay SET name = '" + bgy6 + "' WHERE position = col"
            cur.execute(dbquery6)
            bgy7 = self.ui.line_cpo.text()
            dbquery7 = "UPDATE barangay SET name = '" + bgy7 + "' WHERE position= copo"
            cur.execute(dbquery7)
            bgy8 = self.ui.line_css.text()
            dbquery8 = "UPDATE barangay SET name = '" + bgy8 + "' WHERE position = coss"
            cur.execute(dbquery8)
            bgy9 = self.ui.line_csyd.text()
            dbquery9 = "UPDATE barangay SET name = '" + bgy9 + "' WHERE position = cosyd"
            cur.execute(dbquery9)
            bgy10 = self.ui.line_skchairman.text()
            dbquery10 = "UPDATE barangay SET name = '" + bgy10 + "' WHERE position = sk"
            cur.execute(dbquery10)
            msg = QMessageBox()
            msg.setWindowTitle("Notification")
            msg.setText("Test")
            msg.setIcon(QMessageBox.Icon.Information)
            msg.show()
            msg.exec()
            connection.commit()
        except Exception as e:
            print(str(e))

    def editDocPrice(self):
        cur = connection.cursor()
        cftjs_price = "Price: " + self.ui.lineEdit.text()
        dbq = "UPDATE doc_price SET price = '"+cftjs_price+"' WHERE doc_type = certification"
        cur.execute(dbq)
        bp_price = "Price: " + self.ui.lineEdit_2.text()
        dbq2 = "UPDATE doc_price SET price = '"+bp_price+"' WHERE doc_type = permit"
        cur.execute(dbq2)
        clr_price = "Price: " + self.ui.lineEdit_3.text()
        dbq3 = "UPDATE doc_price SET price = '" + clr_price + "' WHERE doc_type = clearanceLoan"
        cur.execute(dbq3)
        connection.commit()

    def editReqDoc(self):
            try:
                    cur = connection.cursor()
                    uId = self.ui.line_ref.text()
                    cont = self.ui.line_req.text()
                    dbquery = "UPDATE doc_requests SET requirements = '"+cont+"' WHERE reference = " + uId
                    cur.execute(dbquery)
                    connection.commit()

            except Exception as e:
                    print(str(e))
    def deleteRowDocs(self):
        try:
            cur = connection.cursor()
            irow = self.ui.tbl_forview.currentRow()
            uid = self.ui.tbl_forview.item(irow, 0).text()
            cur.execute("DELETE FROM doc_requests WHERE reference = ?", (uid,))
            self.ui.tbl_forview.removeRow(irow)
            self.ui.tbl_forview.setCurrentIndex(QModelIndex())
            connection.commit()

        except:
            print('No Data')

    def loadAccounts(self):
        c = connection.cursor()
        self.ui.tbl_accounts.setRowCount(1000)
        c.execute("SELECT email, verification FROM tbl_users")
        users = c.fetchall()
        tblrow = 0
        for row in users:
            self.ui.tbl_accounts.setItem(tblrow, 0, QtWidgets.QTableWidgetItem(row[0]))
            self.ui.tbl_accounts.setItem(tblrow, 1, QtWidgets.QTableWidgetItem(row[1]))
            tblrow += 1
        connection.commit()

    def loadTransactions(self):
        c = connection.cursor()
        self.ui.tbl_accounts_2.setRowCount(1000)
        c.execute("SELECT * FROM db_transaction")
        users = c.fetchall()
        tblrow = 0
        for row in users:
            self.ui.tbl_accounts_2.setItem(tblrow, 0, QtWidgets.QTableWidgetItem(row[0]))
            self.ui.tbl_accounts_2.setItem(tblrow, 1, QtWidgets.QTableWidgetItem(row[1]))
            self.ui.tbl_accounts_2.setItem(tblrow, 2, QtWidgets.QTableWidgetItem(row[2]))
            tblrow += 1
        connection.commit()

    def deleteRow(self):
        try:
            cur = connection.cursor()
            irow = self.ui.tbl_accounts.currentRow()
            uid = self.ui.tbl_accounts.item(irow, 0).text()
            cur.execute("DELETE FROM tbl_users WHERE email = ?", (uid,))
            self.ui.tbl_accounts.removeRow(irow)
            self.ui.tbl_accounts.setCurrentIndex(QModelIndex())
            connection.commit()

        except:
            print('No Data')

    def loadReq(self):
        c = connection.cursor()
        self.ui.tbl_forview.setRowCount(1000)
        c.execute("SELECT * FROM doc_requests")
        requests = c.fetchall()
        tblrow = 1
        for row in requests:
            self.ui.tbl_forview.setItem(tblrow, 0, QtWidgets.QTableWidgetItem(row[0]))
            self.ui.tbl_forview.setItem(tblrow, 1, QtWidgets.QTableWidgetItem(row[1]))
            self.ui.tbl_forview.setItem(tblrow, 2, QtWidgets.QTableWidgetItem(row[2]))
            self.ui.tbl_forview.setItem(tblrow, 3, QtWidgets.QTableWidgetItem(row[3]))
            self.ui.tbl_forview.setItem(tblrow, 4, QtWidgets.QTableWidgetItem(row[4]))
            self.ui.tbl_forview.setItem(tblrow, 5, QtWidgets.QTableWidgetItem(row[5]))
            self.ui.tbl_forview.setItem(tblrow, 6, QtWidgets.QTableWidgetItem(row[6]))
            self.ui.tbl_forview.setItem(tblrow, 7, QtWidgets.QTableWidgetItem(row[7]))
            self.ui.tbl_forview.setItem(tblrow, 8, QtWidgets.QTableWidgetItem(row[8]))
            self.ui.tbl_forview.setItem(tblrow, 9, QtWidgets.QTableWidgetItem(row[9]))
            self.ui.tbl_forview.setItem(tblrow, 10, QtWidgets.QTableWidgetItem(row[10]))
            self.ui.tbl_forview.setItem(tblrow, 11, QtWidgets.QTableWidgetItem(row[11]))
            self.ui.tbl_forview.setItem(tblrow, 12, QtWidgets.QTableWidgetItem(row[12]))
            self.ui.tbl_forview.setItem(tblrow, 13, QtWidgets.QTableWidgetItem(row[13]))
            self.ui.tbl_forview.setItem(tblrow, 14, QtWidgets.QTableWidgetItem(row[14]))
            self.ui.tbl_forview.setItem(tblrow, 15, QtWidgets.QTableWidgetItem(row[15]))
            self.ui.tbl_forview.setItem(tblrow, 16, QtWidgets.QTableWidgetItem(row[16]))

            tblrow += 1
        connection.commit()

    def user_verif(self):
        try:
            cur = connection.cursor()
            irow = self.ui.tbl_accounts.currentRow()
            uId = self.ui.tbl_accounts.item(irow, 0).text()

            cur.execute("UPDATE tbl_users SET verification = 'Verified' WHERE email = '" + uId + "'")
            connection.commit()

        except Exception as e:
            print(str(e))

    def user_unverif(self):
        try:
            cur = connection.cursor()
            irow = self.ui.tbl_accounts.currentRow()
            uId = self.ui.tbl_accounts.item(irow, 0).text()
            cur.execute("UPDATE tbl_users SET verification = 'Not Yet Verified' WHERE email ='" + uId + "'")
            connection.commit()

        except Exception as e:
            print(str(e))

    def selectDir(self):
        dialog = QFileDialog()
        # dialog.setNameFilters(["Log files (*.log)"])
        c = sqlite3.connect("db_filedir.db")
        db = "CREATE TABLE IF NOT EXISTS tbl_dir(filedir TEXT)"
        con = c.cursor()
        con.execute(db)
        dialog.setOption(QFileDialog.Option.ShowDirsOnly, False)
        download_path = self.ui.line_searchbar.text()

        # open select folder dialog
        fname = QFileDialog.getExistingDirectory(dialog, 'Select a directory', download_path)

        if fname:
            fname = QDir.toNativeSeparators(fname)

        if os.path.isdir(fname):
            global fileD
            fileD = fname

            dba = "UPDATE tbl_dir SET filedir=" + "'" + fileD + "'" + " WHERE rowid = 1"

            con.execute(dba)

            dbq = "SELECT filedir from tbl_dir"
            con.execute(dbq)
            preferred = con.fetchone()[0]
            print(str(preferred))
            fileD = preferred
            dbc = "UPDATE tbl_dir SET filedir='" + str(preferred) + "' WHERE rowid = 1"
            con.execute(dbc)
            c.commit()

    def create_file(self):
        conn = sqlite3.connect("db_filedir.db")
        query = "SELECT filedir FROM tbl_dir"
        con = conn.cursor()
        con.execute(query)
        fileDir = con.fetchone()[0]
        c = connection.cursor()
        dbq = "SELECT name FROM barangay WHERE position LIKE 'captain'"
        c.execute(dbq)
        captain = c.fetchone()[0]
        dbq = "SELECT name FROM barangay WHERE position LIKE 'cofcopw'"
        c.execute(dbq)
        cofcopw = c.fetchone()[0]
        dbq = "SELECT name FROM barangay WHERE position LIKE 'cohsw'"
        c.execute(dbq)
        cohsw = c.fetchone()[0]
        dbq = "SELECT name FROM barangay WHERE position LIKE 'col'"
        c.execute(dbq)
        col = c.fetchone()[0]
        dbq = "SELECT name FROM barangay WHERE position LIKE 'copo'"
        c.execute(dbq)
        copo = c.fetchone()[0]
        dbq = "SELECT name FROM barangay WHERE position LIKE 'coeb'"
        c.execute(dbq)
        coeb = c.fetchone()[0]
        dbq = "SELECT name FROM barangay WHERE position LIKE 'coss'"
        c.execute(dbq)
        coss = c.fetchone()[0]
        dbq = "SELECT name FROM barangay WHERE position LIKE 'cosyd'"
        c.execute(dbq)
        cosyd = c.fetchone()[0]
        dbq = "SELECT name FROM barangay WHERE position LIKE 'sk'"
        c.execute(dbq)
        sk = c.fetchone()[0]
        dbq = "SELECT name FROM barangay WHERE position LIKE 'year'"
        c.execute(dbq)
        year = c.fetchone()[0]
        irow = self.ui.tbl_forview.currentRow()
        udoc = self.ui.tbl_forview.item(irow, 2).text()
        price = self.ui.tbl_forview.item(irow, 3).text()
        name = self.ui.tbl_forview.item(irow, 4).text()
        age = self.ui.tbl_forview.item(irow, 5).text()
        sex = self.ui.tbl_forview.item(irow, 6).text()
        address = self.ui.tbl_forview.item(irow, 7).text()
        requirements = self.ui.tbl_forview.item(irow, 8).text()
        bday = self.ui.tbl_forview.item(irow, 9).text()
        bplace = self.ui.tbl_forview.item(irow, 10).text()
        cstatus = self.ui.tbl_forview.item(irow, 11).text()
        cship = self.ui.tbl_forview.item(irow, 12).text()
        purpose = self.ui.tbl_forview.item(irow, 13).text()
        businame = self.ui.tbl_forview.item(irow, 14).text()
        busiaddress = self.ui.tbl_forview.item(irow, 15).text()

        print(requirements)
        years = self.ui.tbl_forview.item(irow, 16).text()
        try:
            if requirements == "Complete":

                if udoc == "Certification":
                    certif_name = "certification_{}.docx".format(
                        datetime.datetime.now().strftime("%Y-%m-%d"))
                    doc_path = Path(__file__).parent / "certification_template.docx"
                    doc = DocxTemplate(doc_path)
                    dname = name
                    daddress = address
                    context = {
                        "NAME": dname,
                        "ADDRESS": daddress
                    }
                    doc.render(context)
                    doc.save(Path(fileDir) / certif_name)
                    print('success')
                    msg = QMessageBox()
                    msg.setWindowTitle("Notification")
                    msg.setText("File Successfully Created")
                    msg.setIcon(QMessageBox.Icon.Information)
                    msg.show()
                    msg.exec()
                    c = connection.cursor()
                    query = "INSERT INTO db_transaction VALUES (?, ?, ?)"
                    c.execute(query, (name, udoc, price))
                    connection.commit()

                elif udoc == "Indigency":
                    idg_name = "idg_{}.docx".format(datetime.datetime.now().strftime("%Y-%m-%d"))
                    doc_path = Path(__file__).parent / "indigency_template.docx"
                    doc = DocxTemplate(doc_path)
                    context = {
                        "NAME": name,
                        "ADDRESS": address,
                        "PURPOSE": purpose,
                        "YEAR": year,
                        "CAPTAIN": captain,
                        "COFCOPW": cofcopw,
                        "COHSW": cohsw,
                        "COL": col,
                        "COPO": copo,
                        "COEB": coeb,
                        "COSS": coss,
                        "COSYD": cosyd,
                        "SK": sk
                    }
                    doc.render(context)
                    doc.save(Path(fileDir) / idg_name)
                    msg = QMessageBox()
                    msg.setWindowTitle("Notification")
                    msg.setText("File Successfully Created")
                    msg.setIcon(QMessageBox.Icon.Information)
                    msg.show()
                    msg.exec()
                    c = connection.cursor()
                    query = "INSERT INTO db_transaction VALUES (?, ?, ?)"
                    c.execute(query, (name, udoc, price))
                    connection.commit()
                elif udoc == "Business Permit":
                    permit_name = "permit_{}.docx".format(datetime.datetime.now().strftime("%Y-%m-%d"))
                    doc_path = Path(__file__).parent / "permit_template.docx"
                    doc = DocxTemplate(doc_path)
                    context = {
                        "NAME": name,
                        "ADDRESS": address,
                        "BUSINESSNAME": businame,
                        "BUSINESSADDRESS": busiaddress,
                        "YEAR": year,
                        "CAPTAIN": captain,
                        "COFCOPW": cofcopw,
                        "COHSW": cohsw,
                        "COL": col,
                        "COPO": copo,
                        "COEB": coeb,
                        "COSS": coss,
                        "COSYD": cosyd,
                        "SK": sk
                    }
                    doc.render(context)
                    doc.save(Path(fileDir) / permit_name)
                    msg = QMessageBox()
                    msg.setWindowTitle("Notification")
                    msg.setText("File Successfully Created")
                    msg.setIcon(QMessageBox.Icon.Information)
                    msg.show()
                    msg.exec()
                    c = connection.cursor()
                    query = "INSERT INTO db_transaction VALUES (?, ?, ?)"
                    c.execute(query, (name, udoc, price))
                    connection.commit()
                elif udoc == "First Time Job Seeker":
                    fjts_name = "job_seeker_{}.docx".format(
                        datetime.datetime.now().strftime("%Y-%m-%d"))
                    doc_path = Path(__file__).parent / "ftjs_template.docx"
                    doc = DocxTemplate(doc_path)
                    context = {
                        "NAME": name,
                        "ADDRESS": address,
                        "AGE": age,
                        "YEARS": years,
                        "YEAR": year,
                        "CAPTAIN": captain,
                        "COFCOPW": cofcopw,
                        "COHSW": cohsw,
                        "COL": col,
                        "COPO": copo,
                        "COEB": coeb,
                        "COSS": coss,
                        "COSYD": cosyd,
                        "SK": sk
                    }
                    doc.render(context)
                    doc.save(Path(fileDir) / fjts_name)
                    msg = QMessageBox()
                    msg.setWindowTitle("Notification")
                    msg.setText("File Successfully Created")
                    msg.setIcon(QMessageBox.Icon.Information)
                    msg.show()
                    msg.exec()
                    c = connection.cursor()
                    query = "INSERT INTO db_transaction VALUES (?, ?, ?)"
                    c.execute(query, (name, udoc, price))
                    connection.commit()
                elif udoc == "Barangay Clearance":
                    clearance_name = "clearance_{}.docx".format(
                        datetime.datetime.now().strftime("%Y-%m-%d"))
                    doc_path = Path(__file__).parent / "clearance_template.docx"
                    doc = DocxTemplate(doc_path)
                    context = {
                        "NAME": name,
                        "ADDRESS": address,
                        "BIRTHDAY": bday,
                        "BIRTHPLACE": bplace,
                        "SEX": sex,
                        "CIVILSTATUS": cstatus,
                        "CITIZENSHIP": cship,
                        "PURPOSE": purpose,
                        "YEAR": year,
                        "CAPTAIN": captain,
                        "COFCOPW": cofcopw,
                        "COHSW": cohsw,
                        "COL": col,
                        "COPO": copo,
                        "COEB": coeb,
                        "COSS": coss,
                        "COSYD": cosyd,
                        "SK": sk
                    }
                    doc.render(context)
                    doc.save(Path(fileDir) / clearance_name)
                    msg = QMessageBox()
                    msg.setWindowTitle("Notification")
                    msg.setText("File Successfully Created")
                    msg.setIcon(QMessageBox.Icon.Information)
                    msg.show()
                    msg.exec()
                    c = connection.cursor()
                    query = "INSERT INTO db_transaction VALUES (?, ?, ?)"
                    c.execute(query, (name, udoc, price))
                    connection.commit()
                else:
                    print("Something went Wrong")
            else:
                msg = QMessageBox()
                msg.setWindowTitle("Notification")
                msg.setText("Requirements Not Fulfilled")
                msg.setIcon(QMessageBox.Icon.Warning)
                msg.show()
                msg.exec()
        except:
            msg = QMessageBox()
            msg.setWindowTitle("Notification")
            msg.setText("Select A File Directory")
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.show()
            msg.exec()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.show()
    sys.exit(app.exec())
