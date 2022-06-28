# -*- coding: utf-8 -*-
"""
@author: Robin van Gyseghem, Ronny Friedrich
@date: 2022-06-28

This version fully uses fbs for packing
"""

import sys

import urllib.request
from fbs_runtime.application_context.PyQt5 import ApplicationContext
from PyQt5.QtWidgets import QSplitter, QSpacerItem, QSizePolicy, QGroupBox, QFrame, QWidget, QMainWindow, QGroupBox, \
    QTableView, QSpacerItem, QTextEdit, QGridLayout, QDialog, QCheckBox, QApplication, QPushButton, QVBoxLayout, \
    QLineEdit, QLabel, QHBoxLayout, QComboBox , QTabWidget, QProgressBar ,QMenu
from PyQt5 import QtGui
from PyQt5.QtCore import Qt ,pyqtSignal ,QObject, QThread

from datetime import timedelta, datetime
import sqlite3
from matplotlib.ticker import MaxNLocator
from PyQt5 import QtCore
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from pandas.plotting import register_matplotlib_converters
import matplotlib.pyplot as plt
import matplotlib.dates
from pandas import to_datetime, Series
from pandas import DataFrame, read_csv, read_excel, read_sql, DateOffset , ExcelWriter
import pandas
from mpldatacursor import datacursor
from itertools import cycle, islice
from numpy import nan, linspace, tan, arange, sin, pi, isfinite
import os.path
import random
import openpyxl
import xlsxwriter
from config.config_logger import logger

try:
    import xlrd
except ImportError:
    sys.exit("""You need xlrd!
                install it from https://pypi.org/project/xlrd/
                or run pip install xlrd.""")

# set logger name for this module
logger.name = __name__

# import seaborn as sns
global dpi
dpi = 100

myVersion = '2022-06-28'

class SearchWindow(QDialog):
    def __init__(self, parent=None):
        logger.debug('perform function')
        super(SearchWindow, self).__init__(parent)
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMinimizeButtonHint | QtCore.Qt.WindowMaximizeButtonHint )

        #self.setStyleSheet("background-color:lightyellow;")

        self.figure = plt.figure(dpi=dpi)
        #self.figure.set_facecolor("lightyellow")
        # this is the Canvas Widget that displays the `figure`
        # it takes the `figure` instance as a parameter to __init__
        self.canvas = FigureCanvas(self.figure)
        # this is the Navigation widget
        # it takes the Canvas widget and a parent
        self.toolbar = NavigationToolbar(self.canvas, self)
        self.setWindowTitle("Search-Resplor")
        self.editbox = QLineEdit()
        self.roweditX = QComboBox()
        #self.roweditX.setStyleSheet("background-color:lightgray;")
        self.roweditX.setFixedWidth(100)

        self.roweditY = QComboBox()
        #self.roweditY.setStyleSheet("background-color:lightgray;")
        # self.roweditY2 = QComboBox()
        # self.roweditY2.setStyleSheet("background-color:lightgray;")
        #self.roweditY.setFixedWidth(100)
        # self.roweditY2.setFixedWidth(100)
        self.searchcoledit = QComboBox()
        #self.searchcoledit.setStyleSheet("background-color:lightgray;")
        self.searchcoledit.setFixedWidth(100)
        self.nameedit = QLineEdit()
        #self.nameedit.setStyleSheet("background-color:lightgray;")
        # self.nameedit.setFixedWidth(100)
        self.tableedit = QComboBox()
        #self.tableedit.setStyleSheet("background-color:lightgray;")
        self.labelrowX = QLabel("Plotwerte X", self)
        self.labelrowY = QLabel("Plotwerte Y", self)
        # self.labelrowY2 = QLabel("Plotwerte Y", self)
        self.labelsearchcol = QLabel("in Suchspalte", self)
        self.labelname = QLabel("Suche", self)
        self.labeltabelle = QLabel("Tabelle", self)
        # Just some button connected to `plot` method
        self.button = QPushButton('Plot')
        self.button.clicked.connect(self.search)

        self.buttongetdata = QPushButton('Search')
        self.buttongetdata.clicked.connect(self.search)

        self.checkbx = QCheckBox("Plot Temp", self)
        self.checkbx.stateChanged.connect(self.clickBox)
        self.checkbox = False
        self.tableview = QTableView()
        # self.plottbutton = QPushButton('Plot')
        # set the layout
        self.hbox = QHBoxLayout()
        self.vbox = QVBoxLayout()
        self.hbox1 = QHBoxLayout()
        #self.hbox1.addStretch(1)
        self.hbox1.addWidget(self.labelrowX)
        self.hbox1.addWidget(self.roweditX)
        self.hbox1.addStretch(1)
        self.hbox1.addWidget(self.labelrowY)
        self.hbox1.addWidget(self.roweditY)
        # self.hbox1.addWidget(self.labelrowY2)
        # self.hbox1.addWidget(self.roweditY2)
        self.hbox.addLayout(self.vbox)
        #self.hbox.addStretch(4)
        self.hbox.addWidget(self.labeltabelle)
        self.hbox.addWidget(self.tableedit)
        self.hbox.addWidget(self.labelname)
        self.hbox.addWidget(self.nameedit)
        self.hbox.addWidget(self.labelsearchcol)
        self.hbox.addWidget(self.searchcoledit)
        #self.hbox.addStretch(1)

        layout = QVBoxLayout(self)

        layout.addLayout(self.hbox)
        layout.addLayout(self.hbox1)
        layout.addWidget(self.toolbar)
        # layout.addWidget(self.plottbutton)
        layout.addWidget(self.canvas)

        # layout.addWidget(self.editbox)
        layout.addWidget(self.button)
        layout.addWidget(self.checkbx)

        # self.eadialog = None
        # self.eadialog = None
        # self.cnmddialog = None
        # self.oprawdialog = None
        # self.opoddialog = None
        self.setLayout(layout)
        self.setAcceptDrops(True)
        # self.dialogopen = False
        # self.infodialogopen = False
        # self.roweditX.setText('finishdate')
        # self.roweditY.setText('o18gas')
        # self.nameedit.setText('Ag3PO4')
        # self.tableedit.setItem('opod')
        tablelist = self.tablelist()

        self.tableedit.addItems(tablelist)
        self.tableedit.currentIndexChanged.connect(lambda: self.tablechanged())
        self.tablechanged()
        self.tableedit.setCurrentText('cnod')

    def contextMenuEvent(self, event):
        logger.debug('perform function')
        contextMenu = QMenu(self)
        newAct = contextMenu.addAction("New")
        openAct = contextMenu.addAction("Open")
        quitAct = contextMenu.addAction("Quit")
        action = contextMenu.exec_(self.mapToGlobal(event.pos()))
        if action == quitAct:
            self.close()

    def search(self):
        logger.debug('perform function')
        # random data tabletoload = "cnmdtable"
        self.plotcol = self.roweditX.currentText()
        self.plotcolY1 = self.roweditY.currentText()
        # self.plotcolY2 = self.roweditY2.currentText()
        self.searchname = self.nameedit.text()
        self.searchcol = self.searchcoledit.currentText()
        self.tabletoload = self.tableedit.currentText()
        # noch auf geschützte eingabe umändern
        # if self.plotcolY2 != '':
        #   self.plotcolY1 = self.plotcolY1 +","+ self.plotcolY2
        plottable = "SELECT " + self.plotcol + ',' + self.plotcolY1 + ",filename FROM " + self.tabletoload + " WHERE " + self.searchcol + " LIKE '%" + self.searchname + "%';"

        # temphumidvalues = "SELECT temp,humid FROM temphumidtable "
        title = self.searchname
        self.plotprepare(plottable, title)

    def tablechanged(self):
        logger.debug('perform function')
        print('tablechanged to :' + self.tableedit.currentText())
        collist = self.collist()
        self.searchcoledit.clear()
        self.roweditX.clear()
        self.roweditY.clear()
        # self.roweditY2.clear()
        self.searchcoledit.addItems(collist)
        self.roweditX.addItems(collist)
        self.roweditY.addItems(collist)
        # self.roweditY2.addItems(collist)
        # self.roweditY2.addItem('')
        self.roweditY.setCurrentIndex(-1)
        # self.roweditY2.setCurrentIndex(-1)
        #
        self.roweditX.setCurrentText('finishdate')
        self.roweditY.setCurrentText('o18vsmowod')
        self.searchcoledit.setCurrentText('name')

    def collist(self):
        logger.debug('perform function')
        table = self.tableedit.currentText()
        conn = sqlite3.connect(database)
        c = conn.cursor()
        c.execute("PRAGMA table_info(" + table + " );")
        columnlist = c.fetchall()
        columnlist = [x[1] for x in columnlist]
        conn.commit()
        conn.close()
        return (columnlist)

    def tablelist(self):
        logger.debug('perform function')
        conn = sqlite3.connect(database)
        c = conn.cursor()
        c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';")
        tablelist = c.fetchall()
        print(tablelist)
        conn.commit()
        conn.close()
        tablelist = [x[0] for x in tablelist]
        print(tablelist)
        tablelist.remove('temphumidtable')
        # tablelist.remove('temptable')
        return (tablelist)

    def clickBox(self, state):
        logger.debug('perform function')
        if state == QtCore.Qt.Checked:
            print('Checked')
            self.checkbox = True
        else:
            print('Unchecked')
            self.checkbox = False

    def plotprepare(self, plottable, title):
        logger.debug('perform function')
        print(plottable)
        self.title = title
        try:
            connection = sqlite3.connect(database)
            self.plotdf = read_sql(plottable, con=connection, coerce_float=True, params=None, parse_dates=None,
                                   columns=None, chunksize=None)
            # self.temphumiddf = read_sql(temphumidvalues, con = connection, coerce_float=True, params=None, parse_dates=None, chunksize=None)
            self.plotdfcols = self.plotdf.columns
            # plotdf['finishdate'].replace(to_replace=[None], value = "2018.11.11 11:11:11", inplace=True)
            self.plotdf = self.plotdf.sort_values(by=self.plotdf.columns[0])

            if set(['date']).issubset(self.plotdf.columns):
                self.plotdf['date'] = to_datetime(self.plotdf.date, format='%Y/%m/%d %H:%M:%S')
                self.plotdf = self.plotdf.sort_values(by='date')
                self.plotdf = self.plotdf.iloc[1::20, :]

            if set(['finishdate']).issubset(self.plotdf.columns):
                self.plotdf['finishdate'] = to_datetime(self.plotdf.finishdate)
                self.plotdf = self.plotdf.sort_values(by='finishdate')  # sort_byvalues  inplace=True, ascending=False
            if set(['id']).issubset(self.plotdf.columns):
                self.plotdf = self.plotdf.sort_values(by=['id'])

            # self.plotdf = self.plotdf[pd.notnull(self.plotdf[self.plotdf.columns[1]])]  # get rid off nan
            # sortieren nach col0 also finsishdate oder id
            self.plotdf = self.plotdf.reset_index(drop=True)  # index wieder richtig setzen
            # plotdf = plotdf[(plotdf. != 0).any()]     'remove zeros'
            # print(len(plotdfcols))
            # instead of ax.hold(False)
            # self.plotdf.dropna(axis=0, how='any', inplace=True)    #get rid of na an zeros
            self.plotdf = self.plotdf.replace(0, nan)

            if self.checkbox:

                self.makeplot(title=title, plottemp=True)
            else:

                self.makeplot(title=title, plottemp=False)

        except Exception as e:
            message = 'Fehler bei der Ploterstellung: ' + str(e)
            main.open_infodialog(message)

    def makeplot(self, title='', plottemp=False):
        logger.debug('perform function')
        self.plottemp = plottemp
        plt.style.use('seaborn-darkgrid')
        model = PandasModel(self.plotdf)
        Fenstername = 'Plotdata: ' + title
        main.open_new_dialog(Fenstername, 'plot', model)

        self.figure.clear()
        ax = self.figure.add_subplot(111)
        xlabel = self.plotdfcols[0]
        ylabel = self.plotdfcols[1]

        self.figure.suptitle(title, fontsize=16, fontweight='bold')
        # ax.hold(True) # deprecated, see above
        ax.set_xlabel(xlabel, fontsize=16)
        ax.set_ylabel(ylabel, fontsize=16)

        if 'filename' in self.plotdf.columns:
            print("filename listed to plot")
            print(self.plotdf.filename[0])
            newlist = list(self.plotdf.filename)
            aktuell = self.plotdf.filename[0]
            farbe = 'k'
            farbliste = []

            for inhalt in newlist:
                if inhalt == aktuell:
                    print(farbe + inhalt)
                    farbliste.append(farbe)
                else:
                    if farbe == 'k':
                        farbe = 'r'
                    else:
                        farbe = 'k'
                    farbliste.append(farbe)
                aktuell = inhalt

            self.plotdf['farben'] = farbliste

            print(self.plotdf)
            grouped = self.plotdf.groupby('farben')
            blackdf = grouped.get_group('k')
            print(grouped.get_group('k'))

            try:                                                                #abfangen falls nur ein filename in der liste ist
                reddf = grouped.get_group('r')
            except:
                 reddf= blackdf
                 print('only one color')
                 pass

            start, end = ax.get_xlim()
            stepsize = 2
            if (self.plotdf[self.plotdfcols[0]].dtype == "datetime64[ns]"):
                plt.xticks(rotation=45, fontsize=9)

                color1 = ax.plot_date(blackdf[self.plotdfcols[0]], blackdf[self.plotdfcols[1]], fmt='o', color='sienna',
                                      picker=5)

                color2 = ax.plot_date(reddf[self.plotdfcols[0]], reddf[self.plotdfcols[1]], fmt='o', color='orange',
                                      picker=5, )

                # ax.xaxis.set_ticks(np.arange(start, end, stepsize)

            else:
                ax.scatter(self.plotdf[self.plotdfcols[0]], self.plotdf[self.plotdfcols[1]], color=farbliste)

        elif (self.plotdf[self.plotdfcols[0]].dtype == "datetime64[ns]"):
            plt.xticks(rotation=45, fontsize=9)
            # ax.xaxis.set_ticks(np.arange(start, end, stepsize))
            ax.plot_date(self.plotdf[self.plotdfcols[0]], self.plotdf[self.plotdfcols[1]], fmt='o', picker=5, )

        else:
            ax.scatter(self.plotdf[self.plotdfcols[0]], self.plotdf[self.plotdfcols[1]], color='g')

        if self.plottemp:
            ax2 = ax.twinx()
            tempdata = self.get_temp_values()
            ax2.plot(tempdata[tempdata.columns[0]], tempdata[tempdata.columns[1]], color='r', alpha=.2)
            ax2.set_ylabel('temp °C', fontsize=16)

        '''
        You
        can
        groupby and plot
        them
        separately
        for each color:

        import matplotlib.pyplot as plt

        fig, ax1 = plt.subplots(figsize=(30, 10))
        color = 'tab:red'
        for pcolor, gp in df.groupby('color'):
            ax1.plot_date(gp['time'], gp['distance'], marker='o', color=pcolor)
        '''
        # ax.set_position([0, 0, 0,0])
        self.figure.subplots_adjust(0.1, 0.2, 0.9, 0.9, )  # 0.2,0.3

        datacursor(formatter=self.myformatter, display='multiple', draggable=True)
        ax.legend(fontsize=12)
        ax.legend().set_visible(False)
        ax.grid(True)
        # (pdfile.index,pdfile.values,150,marker = ">")
        # ax.plot(plotdf.columns[1].value, plotdf.columns[1].value)
        self.canvas.draw()

    def myformatter(self, **kwarg):
        logger.debug('perform function')
        values = self.collist()
        items = []
        # for item in values
        #   value = searchvalue(item)

        xaxis = self.plotdfcols[0] + ':  '
        yaxis = self.plotdfcols[1] + ':  '
        if self.plotdfcols[0] == 'finishdate':
            val1 = matplotlib.dates.num2date(kwarg['x']).strftime('%Y-%m-%d %H:%M:%S')
        else:
            val1 = self.getfinishdate(kwarg['x'], self.plotdfcols[0])
            print(self.plotdfcols[0], int(kwarg['x']))
            print("retrived:", val1)

        valuesdf = self.getvalues(val1)
        pandas.set_option('display.max_columns', 30)
        # label = xaxis + val1 + '\n' + yaxis + ' {y:.3f}'.format(**kwarg) +str(values)

        label = str(valuesdf)
        label = label.split("\n", 1)[1]
        print(type(val1), val1)
        return label

    # if len(self.plotdfcols)>2:
    #    ax.scatter(self.plotdf[self.plotdfcols[0]], self.plotdf[self.plotdfcols[2]])
    # temperatur plotten
    # ax.scatter(self.temphumiddf[0],self.temphumiddf[1])

    # datacursor(ax)

    '''SELECT DISTINCT column_list
                FROM table_list
                JOIN table ON join_condition
                WHERE row_filter
                ORDER BY column
                LIMIT count OFFSET offset
                GROUP BY column
        HAVING group_filter '''

    def getvalues(self, finishdate):
        logger.debug('perform function')
        print(finishdate)
        print(type(finishdate))
        connection = sqlite3.connect(database)
        sqlsynt = "SELECT * FROM " + self.tabletoload + " WHERE finishdate LIKE '%" + str(finishdate) + "%';"
        values = read_sql(sqlsynt, con=connection, coerce_float=True, params=None, parse_dates=None, columns=None,
                          chunksize=None)
        connection.commit
        values = values.T

        print(values)
        return values

    def get_temp_values(self):
        logger.debug('perform function')
        connection = sqlite3.connect(database)
        tempvalues = "SELECT date,temp FROM temphumidtable"
        tempdata = read_sql(tempvalues, con=connection, coerce_float=True, params=None, parse_dates=None, columns=None,
                            chunksize=None)
        tempdata['date'] = to_datetime(tempdata['date'], format='%Y/%m/%d %H:%M:%S')
        tempdata = tempdata.sort_values(by='date')
        tempdata = tempdata.iloc[1::20, :]
        tempdata = tempdata.replace(0, 'nan')

        tempdata.dropna(axis=0, how='any', inplace=True)
        return tempdata


class MainWindow(QDialog):
    '''
    Main Class that defines the main window after program start with all the buttons
    '''
    register_matplotlib_converters()

    # set name of the database - using sqlite for not that is tored in the same folder as the code
    # --------------------------------------------------------
    global database
    database = "data1.db"
    logger.debug('set database to ' + database)
    # --------------------------------------------------------

    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setWindowTitle('Resplor v. ' + myVersion)
        self.setWindowFlags(
            self.windowFlags() | QtCore.Qt.WindowMinimizeButtonHint )
        # a figure instance to plot on
        # self.setStyleSheet("background-color:lightyellow;")

        self.resize(300,300)
        self.move(100,100)
        # Just some button connected to `plot` method

        self.buttonsql = QPushButton('SQL query')
        self.buttonsql.clicked.connect(self.sqlquery)
        self.buttonstd = QPushButton('Standards')
        self.buttonstd.clicked.connect(self.stdtab)
        self.buttonsamples = QPushButton('Samples')
        self.buttonsamples.clicked.connect(self.samplestaba)
        self.buttonsearch = QPushButton('Search')
        self.buttonsearch.clicked.connect(self.searchtab)
        self.buttondeleterun = QPushButton('Delete Run Data')
        self.buttondeleterun.clicked.connect(self.deltab)
        self.buttonoutput = QPushButton('Output')
        self.buttonoutput.clicked.connect(self.outputtab)

        self.buttonoptions = QPushButton('Options')
        self.buttonoptions.clicked.connect(self.optiontab)

        self.image = QLabel(self)
        #pixmap = QtGui.QPixmap("speichern.png")
        # when packaging with fbs, files are loaded like this using the ApplicationContext appctxt
        # the image is located in src/main/resources/base/
        pixmap = QtGui.QPixmap(appctxt.get_resource('speichern.png')) # when packaging with fbs, files are loaded like this
        #pixmap.fill(Qt.transparent)
        self.image.setPixmap(pixmap)
        self.image.setAlignment(Qt.AlignCenter)
        self.image.setToolTip('Drop Excel Files here')

        # generate Layout
        self.vbox = QVBoxLayout()
        self.vbox.addWidget(self.buttonsamples)
        self.vbox.addWidget(self.buttonstd)
        self.vbox.addWidget(self.buttonsearch)
        self.vbox.addWidget(self.buttonsql)
        self.vbox.addWidget(self.buttondeleterun)
        self.vbox.addWidget(self.buttonoutput)
        self.vbox.addWidget(self.buttonoptions)
        self.vbox.addWidget(self.image)
        self.setLayout(self.vbox)

        self.setAcceptDrops(True)
        # self.dialogopen = False
        # self.infodialogopen = False
        # self.roweditX.setText('finishdate')
        # self.roweditY.setText('o18gas')
        # self.nameedit.setText('Ag3PO4')
        # self.tableedit.setItem('opod')

    def getfinishdate(self, value, valuename):
        logger.debug('perform function')
        valuef = int(value)
        connection = sqlite3.connect(database)
        sqlsynt = "SELECT finishdate FROM " + self.tabletoload + " WHERE " + valuename + " LIKE " + str(valuef)
        datedf = read_sql(sqlsynt, con=connection, coerce_float=True, params=None, parse_dates=None, columns=None,
                          chunksize=None)
        connection.commit
        print(datedf.finishdate.values)
        return datedf.finishdate.values[0]

    def get_temp_values(self):
        logger.debug('perform function')
        connection = sqlite3.connect(database)
        tempvalues = "SELECT date,temp FROM temphumidtable"
        tempdata = read_sql(tempvalues, con=connection, coerce_float=True, params=None, parse_dates=None, columns=None,
                            chunksize=None)
        tempdata['date'] = to_datetime(tempdata['date'], format='%Y/%m/%d %H:%M:%S')
        tempdata = tempdata.sort_values(by='date')
        tempdata = tempdata.iloc[1::20, :]
        tempdata = tempdata.replace(0, 'nan')

        tempdata.dropna(axis=0, how='any', inplace=True)
        return tempdata

    def geklickt(self):
        logger.debug('perform function')
        print()

    def delzero(self):
        logger.debug('perform function')
        self.plotdf = self.plotdf.replace(0, nan)  # set zero to nan
        self.plotdialog.done(0)
        self.makeplot()

    def dropna(self):
        logger.debug('perform function')
        try:
            self.plotdf.dropna(axis=0, how='any', inplace=True)  #
            self.plotdialog.done(0)
            self.makeplot()
        except Exception as e:
            message = 'there are probs droping na values' + e
            self.open_infodialog(message)

    def dragEnterEvent(self, event):
        logger.debug('perform function')
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        logger.debug('perform function')
        for url in event.mimeData().urls():
            file = url.toLocalFile()
            logger.debug('File dropped: ' + file)
            if os.path.isfile(file):
                print(file)
               # self.editbox.setText(file)
                self.importfile = file
                self.sheetErzeugen()

    def sheetErzeugen(self):
        # runs when a excel sheet is dropped onto the drop field
        logger.debug('perform function')
        logger.debug('excel file has been dropped onto the drop field')
        self.samshe = Samplesheet() # instantiate samplesheet
        logger.debug('instantiate sample sheet for new data')
        print("methode sheeterzeugen - samplesheet erstellt")
        logger.debug('methode sheeterzeugen - samplesheet erstellt')

        file = self.importfile
        self.db_anlegen()
        try:
            self.samshe.load(file)
        except Exception as e:
            message = 'Fehler in der Formatierung ' + str(e)
            self.open_infodialog(message)

    def sqlquery(self):
        logger.debug('perform function')
        file = 'sql query'
        sqldialog = SqlDialog(self, database)
        # sqldialog.resize(700, 300)

        sqldialog.show()

    def stdtab(self):
        logger.debug('perform function')
        self.dialog = SampleDialog(self)
        self.dialog.changelayout2()
        self.dialog.show()

    def samplestaba(self):
        logger.debug('perform function')
        self.dialog = SampleDialog(self)
        self.dialog.changelayout()
        self.dialog.show()

    def searchtab(self):
        logger.debug('perform function')
        self.searchwindow = SearchWindow()
        self.searchwindow.show()

    def outputtab(self):
        logger.debug('perform function')
        self.dialog = SampleDialog(self)
        self.dialog.changelayout3()
        self.dialog.show()

    def deltab(self):
        logger.debug('perform function')
        self.dialog_del = SampleDialog(self)
        self.dialog_del.changelayout_delrun()
        self.dialog_del.show()

    def optiontab(self):
        logger.debug('perform function')
        self.dialog = QDialog()
        # self.dialog.setStyleSheet("background-color:lightyellow;")
        self.dialog.setWindowTitle('Options')
        self.layout = QVBoxLayout(self.dialog)
        self.Hlayout = QHBoxLayout()
        self.H2layout = QHBoxLayout()

        label = QLabel('Set DPI')

        self.dpiedit = QLineEdit('100')
        self.dpiedit.setAlignment(QtCore.Qt.AlignCenter)
        self.savebtn = QPushButton('Save')
        #self.delbtn = QPushButton('Delete Run')

        self.Hlayout.addWidget(label)
        self.Hlayout.addWidget(self.dpiedit)
        self.Hlayout.addWidget(self.savebtn)

        self.layout.addLayout(self.Hlayout)
        #self.H2layout.addWidget(self.delbtn)

        self.layout.addLayout(self.H2layout)

        self.savebtn.clicked.connect(self.save_dpi)
        #self.delbtn.clicked.connect(self.deltab)
        self.dialog.show()

    def save_dpi(self):
        logger.debug('perform function')
        global dpi
        dpi = int(self.dpiedit.text())
        print('new dpi value:' , dpi)

    def open_new_dialog(self, title, origin, model):
        logger.debug('perform function')

        if origin == 'standard':
            self.dialog = NewDialog(self, title)
            self.dialog.tableview.setModel(model)
            self.dialog.show()
        # or 'opea'
        if (origin == 'opea'):
            if self.eadialog is None:
                self.opeadialog = NewDialog(self, title)
                self.opeadialog.tableview.setModel(model)
                self.opeadialog.resize(900, 350)
                self.opeadialog.move(100, 100)
                self.opeadialog.show()
            if self.opeadialog is not None:
                self.opeadialog.tableview.setModel(model)
                self.opeadialog.show()

        if (origin == 'ea'):
            if self.eadialog is None:
                self.eadialog = NewDialog(self, title)
                self.eadialog.tableview.setModel(model)
                self.eadialog.resize(900, 350)
                self.eadialog.move(100, 100)
                self.eadialog.show()
            if self.eadialog is not None:
                self.eadialog.tableview.setModel(model)
                self.eadialog.show()

        if (origin == 'cnmd'):
            self.cnmddialog = NewDialog(self, title)
            self.cnmddialog.tableview.setModel(model)
            self.cnmddialog.show()

        if (origin == 'cnod'):
            self.cnmddialog = NewDialog(self, title)
            self.cnmddialog.tableview.setModel(model)
            self.cnmddialog.show()

        if (origin == "plot"):
            plotdialog = NewDialog(self, title)
            plotdialog.tableview.setModel(model)
            plotdialog.resize(220, 400)
            plotdialog.move(100, 100)
            delzerobutton = QPushButton('zero -> NaN')
            delzerobutton.clicked.connect(self.delzero)
            dropnabutton = QPushButton('drop NaN')
            dropnabutton.clicked.connect(self.dropna)
            hbox = QHBoxLayout()
            hbox.addWidget(delzerobutton)
            hbox.addWidget(dropnabutton)
            plotdialog.layout.addLayout(hbox)
            self.plotdialog = plotdialog
            self.plotdialog.show()

        if origin == "opmd":
            self.oprawdialog = NewDialog(self, title)
            self.oprawdialog.tableview.setModel(model)
            self.oprawdialog.show()
        # if (origin == "opraw" and self.oprawdialog is not None):
        # self.oprawdialog.tableview.setModel(model)
        # .oprawdialog.show()

        if origin == "opod":
            self.opoddialog = NewDialog(self, title)
            self.opoddialog.tableview.setModel(model)
            self.opoddialog.show()
        # if (origin == "opod" and self.opoddialog is not None):
        # self.opoddialog.tableview.setModel(model)
        # self.opoddialog.show()

    def open_infodialog(self, message):
        logger.debug('perform function')
        self.fehler = InfoDialog(self, message)
        self.fehler.show()

    def db_anlegen(self):
        logger.debug('perform function')
        if not os.path.exists(database):
            print("Datenbank data.db nicht vorhanden - Datenbank wird anglegt.")
            connection = sqlite3.connect(database)
            cursor = connection.cursor()
            # Tabelle erzeugen
            # sql = "CREATE TABLE rawtableco (id INT PRIMARY KEY,  name text NOT NULL, weight FLOAT,\
            #        finishdate TIMESTAMP DEFAULT CURRENT_TIMESTAMP,finishmax TIMESTAMP DEFAULT CURRENT_TIMESTAMP,beamheight FLOAT,dO18gas FLOAT,\
            #        dO18drift FLOAT,dO18 FLOAT,dN15 FLOAT,dN15drift FLOAT,dC13 FLOAT,dC13drift FLOAT,\
            #       c FLOAT, n FLOAT,cnratio FLOAT,cfact FLOAT,nfact FLOAT, UNIQUE(finishdate,name))"
            rawtableco = "CREATE TABLE rawtableco (idraw INT,superid text  PRIMARY KEY, name text NOT NULL,sampletype text,finishdate TIMESTAMP DEFAULT CURRENT_TIMESTAMP,peakid INT,time TIMESTAMP DEFAULT CURRENT_TIMESTAMP ,width TIMESTAMP DEFAULT CURRENT_TIMESTAMP,\
                       height FLOAT ,area FLOAT,ratio2928 FLOAT,ratio2928raw FLOAT,ratio3028 FLOAT,ratio3028raw FLOAT,c13gas FLOAT,std13C FLOAT,\
                       o18gas FLOAT,o18vsmowmd FLOAT,bsC13gas FLOAT,c13gasdrift FLOAT,c13 FLOAT,stddiffc13 FLOAT,bsO18gas FLOAT,o18gasdrift FLOAT,stddiffdO18 FLOAT,quality INT,final INT, comment text, extra INT, filename TEXT)"
            opodtable = "CREATE TABLE opodtable (id INT PRIMARY KEY, name text NOT NULL ,finishdate TIMESTAMP DEFAULT CURRENT_TIMESTAMP,opercent FLOAT, oarea FLOAT,height FLOAT,sampletype TEXT,o18gas FLOAT,o18vsmowod FLOAT,runid TEXT,filename TEXT,quality INT,final INT, comment text, extra INT)"


            #primary key auf superid und filename ausweiten
            opmd = "CREATE TABLE opmd (idraw INT,superid text , name text NOT NULL,sampletype text,notes text,finishdate TIMESTAMP DEFAULT CURRENT_TIMESTAMP,weight float,peakid INT,time TIMESTAMP DEFAULT CURRENT_TIMESTAMP ,width TIMESTAMP DEFAULT CURRENT_TIMESTAMP,\
                       height FLOAT ,area FLOAT,ratio2928 FLOAT,ratio2928raw FLOAT,ratio3028 FLOAT,ratio3028raw FLOAT,c13gas FLOAT,std13C FLOAT,\
                       o18gas FLOAT,o18vsmowmd FLOAT,bsC13gas FLOAT,c13gasdrift FLOAT,c13 FLOAT,stddiffc13 FLOAT,bsO18gas FLOAT,o18gasdrift FLOAT,stddiffdO18 FLOAT,quality INT,final INT, comment text, extra INT, filename TEXT, PRIMARY KEY (superid,filename)"

            #primary key auf id und filename ausweiten
            opod = "CREATE TABLE opod (id INT , name text NOT NULL ,notes TEXT,finishdate TIMESTAMP DEFAULT CURRENT_TIMESTAMP,weight FLOAT, opercent FLOAT, oarea FLOAT,height FLOAT,sampletype TEXT,o18gas FLOAT,o18vsmowod FLOAT,runid TEXT,filename TEXT,quality INT,final INT, comment text, extra INT, PRIMARY KEY (id,filename))"



            cnmd = "CREATE TABLE cnmd (idraw INT,superid text ,name text NOT NULL,sampletype text,notes text,finishdate TIMESTAMP CURRENT_TIMESTAMP,weight float,peakid INT,time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,width TIMESTAMP DEFAULT CURRENT_TIMESTAMP,height FLOAT,area FLOAT,ratio2928 FLOAT,ratio2928raw FLOAT,n15gas FLOAT,stdn15air FLOAT,n15gasdrift FLOAT,n15aircali FLOAT,ratio4544 FLOAT,ratio4544raw FLOAT,ratio4644 FLOAT,ratio4644raw FLOAT,c13gas FLOAT,stdc13vpdb FLOAT,c13gasdrift FLOAT,c13vpdbcali FLOAT,runid TEXT,filename TEXT,quality INT,final INT, comment text, extra INT, PRIMARY KEY(superid , filename))"

            cnod = "CREATE TABLE cnod (id INT ,sampletype TEXT, name text NOT NULL ,notes text, finishdate TIMESTAMP DEFAULT CURRENT_TIMESTAMP,weight float, areac INT,cpercent FLOAT,arean INT,npercent FLOAT,nheight FLOAT,nisoarea FLOAT, n15gas FLOAT, n15drift FLOAT,n15air FLOAT,cheight FLOAT,cisoarea FLOAT, c13gas FLOAT, c13drift FLOAT,c13vpdb FLOAT,runid TEXT,filename TEXT,quality INT,final INT, comment text, extra INT, PRIMARY KEY(id,filename))"

            rawtablecn = "CREATE TABLE rawtablecn (idraw INT,superid text PRIMARY KEY,name text NOT NULL,sampletype text,finishdate TIMESTAMP CURRENT_TIMESTAMP,peakid INT,time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,width TIMESTAMP DEFAULT CURRENT_TIMESTAMP,height FLOAT,area FLOAT,ratio2928 FLOAT,ratio2928raw FLOAT,n15gas FLOAT,stdn15air FLOAT,n15gasdrift FLOAT,n15aircali FLOAT,ratio4544 FLOAT,ratio4544raw FLOAT,ratio4644 FLOAT,ratio4644raw FLOAT,c13gas FLOAT,stdc13vpdb FLOAT,c13gasdrift FLOAT,c13vpdbcali FLOAT,runid TEXT,filename TEXT,quality INT,final INT, comment text, extra INT)"

            cnodtable = "CREATE TABLE cnodtable (id INT PRIMARY KEY,sampletype TEXT, name text NOT NULL ,finishdate TIMESTAMP DEFAULT CURRENT_TIMESTAMP,areac INT,cpercent FLOAT,arean INT,npercent FLOAT,nheight FLOAT,nisoarea FLOAT, n15gas FLOAT, n15drift FLOAT,n15air FLOAT,cheight FLOAT,cisoarea FLOAT, c13gas FLOAT, c13drift FLOAT,c13vpdb FLOAT,runid TEXT,filename TEXT,quality INT,final INT, comment text, extra INT)"

            temphumidtable = "CREATE TABLE temphumidtable (date TIMESTAMP DEFAULT CURRENT_TIMESTAMP PRIMARY KEY, seconds FLOAT, temp FLOAT, humid FLOAT )"

            metadata = "CREATE TABLE metadata (manr INT PRIMARY KEY, name TEXT , fundland TEXT, fundort TEXT, fundplatz TEXT, datierung TEXT,skelettelement TEXT, geschlecht TEXT , tierart TEXT, altermin INT , altermax INT, ausbeute FLOAT ,bemerkung TEXT, mams INT)"

            cursor.execute(opod)
            cursor.execute(opmd)
            cursor.execute(cnod)
            cursor.execute(cnmd)
            cursor.execute(temphumidtable)
            cursor.execute(metadata)
            connection.commit()

    def setruns(self):
        logger.debug('perform function')
        a = 0
        connection = sqlite3.connect(database)
        c = connection.cursor()
        # c.execute("SELECT finishdate FROM opodtable")
        # finishdates = c.fetchall()
        # for x in finishdates:
        # print(x)
        finishdates = "SELECT finishdate , id FROM opod"
        datedf = read_sql(finishdates, con=connection, coerce_float=True, params=None, parse_dates=None, columns=None,
                          chunksize=None)
        datedf = datedf.sort_values(by='finishdate')
        datedf['finishdate'] = to_datetime(datedf.finishdate)
        datedf['runid'] = 1
        #print(datedf)
        #print(datedf.dtypes)
        while a < 20:
            pos = datedf.iloc[a]
            print(pos)
            a += 1

    def setrunid(self, runid):
        logger.debug('perform function')
        connection = sqlite3.connect(database)
        c = connection.cursor()
        setrunid = "INSERT INTO opod(runid) VALUE " + runid
        c.execute(setrunid)
        c.close()
        connection.commit()
        print('runid inserted')

    def getrunid(self):
        logger.debug('perform function')
        connection = sqlite3.connect(database)
        c = connection.cursor()
        try:
            getlastidop = "SELECT MAX(runid) FROM opod"
            getlastidcn = "SELECT MAX(runid) FROM cnod"
            c.execute(getlastidop)
            lastidop = c.fetchall()
            c.execute(getlastidcn)
            lastidcn = c.fetchall()
            if lastidop > lastidcn:
                newid = lastidop + 1
            else:
                newid = lastidcn + 1

            print('actual runid:' + newid)
            return newid

        except:
            print('no runid in table, lastid automaticaly set to 0')
            return 0

    def get_triples(self, resultdf, gleiches='set', tocalculate='o18vsmowod', tocalculate2='o18vsmowmd',
                    tocalculate3='', tocalculate4=''):
        logger.debug('perform function')
        toignore = ['Blnk']     # ignore all blanks
        pos = 0
        anfangspos = 0
        self.resultdf = resultdf
        logger.debug(self.resultdf.head(5))
        self.resultdf['set'] = ''
        aktuell = self.resultdf.name[0]
        logger.debug('aktueller Name: ' + aktuell)
        self.triple = []

        for inhalt in self.resultdf.name:
        #ist der name der folgenden zeile gleich , notiere die id, ansonsten den namen der nächsten
        #ist der name der folgende zeile Blnk überspringe diese

            if inhalt == aktuell and inhalt != 'Blnk' :
                self.triple.append(self.resultdf.id[anfangspos])
                pos += 1

            else :
                self.triple.append(self.resultdf.id[pos])
                anfangspos = pos
                pos += 1
                if inhalt != 'Blnk':
                    aktuell = inhalt

        self.resultdf.set = self.triple

        groupedmean = self.resultdf.groupby(gleiches)[tocalculate].mean()
        groupedstd = self.resultdf.groupby(gleiches)[tocalculate].std()

        groupedmean2 = self.resultdf.groupby(gleiches)[tocalculate2].mean()
        groupedstd2 = self.resultdf.groupby(gleiches)[tocalculate2].std()
        if tocalculate3 != '':
            groupedmean3 = self.resultdf.groupby(gleiches)[tocalculate3].mean()
            groupedstd3 = self.resultdf.groupby(gleiches)[tocalculate3].std()

            mean3 = groupedmean3.to_frame()
            std3 = groupedstd3.to_frame()

            meanname3 = tocalculate3 + 'avg'
            stdname3 = tocalculate3 + 'std'

            mean3.columns = [meanname3]
            std3.columns = [stdname3]

            self.resultdf = pandas.merge(self.resultdf, mean3, on='set')
            self.resultdf = pandas.merge(self.resultdf, std3, on='set')

        if tocalculate4 !='':
            groupedmean4 = self.resultdf.groupby(gleiches)[tocalculate4].mean()
            groupedstd4 = self.resultdf.groupby(gleiches)[tocalculate4].std()
            mean4 = groupedmean4.to_frame()
            std4 = groupedstd4.to_frame()
            meanname4 = tocalculate4 + 'avg'
            stdname4 = tocalculate4 + 'std'
            mean4.columns = [meanname4]
            std4.columns = [stdname4]
            self.resultdf = pandas.merge(self.resultdf, mean4, on='set')
            self.resultdf = pandas.merge(self.resultdf, std4, on='set')

        mean = groupedmean.to_frame()
        std = groupedstd.to_frame()
        mean2 = groupedmean2.to_frame()
        std2 = groupedstd2.to_frame()
        meanname = tocalculate + 'avg'
        stdname = tocalculate + 'std'
        meanname2 = tocalculate2 + 'avg'
        stdname2 = tocalculate2 + 'std'

        mean.columns = [meanname]
        std.columns = [stdname]
        mean2.columns = [meanname2]
        std2.columns = [stdname2]

        self.resultdf = pandas.merge(self.resultdf, mean, on='set')
        self.resultdf = pandas.merge(self.resultdf, std, on='set')
        self.resultdf = pandas.merge(self.resultdf, mean2, on='set')
        self.resultdf = pandas.merge(self.resultdf, std2, on='set')
        #print('tripletdata:', self.resultdf)
        #print('hier', self.resultdf)
        realdataframe = DataFrame(self.resultdf,index=None)
        return (realdataframe)


class Samplesheet(object):
    def __init__(self):
        logger.debug('perform function')
        self.origin = ''

    # def speichern(self,fields):

    def addtodb(self):
        logger.debug('perform function')
        logger.debug('Inserting imported data into DB')
        origin = self.origin
        connection = sqlite3.connect(database)
        cursor = connection.cursor()

        self.tempdf.to_sql('temptable', index=False, con=connection, if_exists='replace')

        # depending on the type of file that was imported, decided what query to run
        # in order to INSERT those data into the correct table
        if origin == 'metadata':
            logger.debug('creating query for ' + origin)
            qry = "INSERT OR IGNORE INTO metadata (manr , name , fundland , fundort , fundplatz ,datierung, skelettelement ,tierart, geschlecht , bemerkung,mams,altermin )\
                            SELECT manr , name,fundland,fundort,fundplatz,datierung, skelettelement, tierart, geschlecht, bemerkung, mams,alterm FROM temptable"
            logger.debug(qry)
            tabletoload = 'metadata'

        if origin == 'templog':
            logger.debug('creating query for ' + origin)
            qry = "INSERT OR IGNORE INTO temphumidtable (date,seconds,temp,humid) \
                                           SELECT date,seconds,temp,humid FROM temptable "
            logger.debug(qry)
            tabletoload = "temphumidtable"

        if origin == 'opmd':
            logger.debug('creating query for ' + origin)
            qry = "INSERT OR IGNORE INTO opmd (idraw,superid ,name ,sampletype,finishdate, peakid ,time  ,width ,\
                                height,area,ratio2928,ratio2928raw,ratio3028,ratio3028raw,c13gas,std13C,o18gas,bsC13gas ,c13gasdrift,c13,\
                                stddiffc13,bsO18gas,o18gasdrift,o18vsmowmd,stddiffdO18,filename)\
                                SELECT Id,superid,Name,SampleType,finishdate,PeakID,Time,Width,Height,Area,ratio2928,ratio2928raw,ratio3028,ratio3028raw,C13gas,Std13C,\
                                O18gas,bsC13gas,C13gasdrift,C13vpdb,stddiffC13,bsO18gas,O18gasdrift,O18vsmowmd,stddiffdO18,filename FROM temptable"
            logger.debug(qry)
            tabletoload = "opmd"

        if origin == 'opod':
            logger.debug('creating query for ' + origin)
            qry = "INSERT OR IGNORE INTO opod(id, name, finishdate,sampletype,opercent,oarea, height, o18gas, o18vsmowod ,filename)\
                                SELECT id,name,finishdate,sampletype,opercent,oarea,height,o18gas,o18vsmow,filename FROM temptable"
            logger.debug(qry)
            tabletoload = "opod"

        if origin == 'cnmd':
            logger.debug('creating query for ' + origin)
            qry = 'INSERT OR IGNORE INTO cnmd (notes,weight,ratio4544,idraw,superid,name,sampletype,finishdate,peakid,time,width,height,area,ratio2928,ratio2928raw,n15gas,n15aircali,ratio4544raw,ratio4644,ratio4644raw,c13gas,stdc13vpdb,c13gasdrift,c13vpdbcali,filename)\
                                SELECT notes,weight,ratio4544,Id,superid,Name,SampleType,finishdate,PeakID,Time,Width,Height,IsoArea,ratio2928,ratio2928raw,N15gas,N15air,ratio4544raw,ratio4644,ratio4644raw,C13gas,stdC13,driftC13gas,C13vpdb ,filename FROM temptable'
            logger.debug(qry)
            tabletoload = "cnmd"

        if origin == 'cnod':
            logger.debug('creating query for ' + origin)
            qry = "INSERT OR IGNORE INTO cnod(notes,weight,id,sampletype,name,finishdate,areac,cpercent,arean,npercent,nheight,nisoarea,n15gas,n15drift,n15air,cheight,cisoarea,c13gas,c13drift,c13vpdb,filename)\
                                                    SELECT notes,weight, id,sampletype,name,finishdate,areac,cpercent,arean,npercent,nheight,nisoarea,n15gas,n15drift,n15air,cheight,cisoarea,c13gas,c13drift,c13vpdb,filename FROM temptable"
            logger.debug(qry)
            tabletoload = "cnod"
        
        # now run the query that was created above according to the type of excel file
        try:
            print('writing imported data to ' + tabletoload)
            logger.debug('writing imported data to ' + tabletoload)
            logger.debug('try executing query')
            cursor.execute(qry)
            logger.debug('try executing query...success')
        except Exception as e:
            logger.debug('try executing query...failed')
            errmessage = 'Error while writing to DB: ' + str(e)
            logger.exception(errmessage)
            main.open_infodialog(errmessage)

        # query that loads the whole table back into the dataframe
        # in order to have all the recent data available in the "resultsdf"
        loadtable = '''SELECT * FROM ''' + tabletoload
        resultdf = read_sql(loadtable, con=connection, index_col=None, coerce_float=True, params=None, parse_dates=None,
                            columns=None, chunksize=None)
        # resultdf['id']= resultdf['id'].astype(int)
        #print(resultdf)

        cursor.close()
        connection.commit()

        # display a table that shows the imported data 
        self.origin = origin
        model = PandasModel(self.tempdf)    # that's the model that is behind the table
        head, tail = os.path.split(self.file)
        main.open_new_dialog(tail, origin, model)

    def replacedata(self, data):
        logger.debug('perform function')
        self.tempdf.loc[self.tempdf['finishdate'].isnull(), 'finishdate'] = data
        self.tempdf['finishdate'] = to_datetime(self.tempdf.finishdate)
        #print(self.tempdf['finishdate'])
        self.finishdatencheck()

    def finishdatencheck(self):
        logger.debug('perform function')
        if not None in self.tempdf.finishdate.values:
            print('finishdates ok')
            # add samples to db
            logger.debug('adding imported excel data to the DB')
            self.addtodb()
        else:
            message = 'Es fehlt ein Finishdate'
            print(message)
            main.open_infodialog(message)

    def load(self, file):
        logger.debug('perform function')
        logger.debug('load data from excel into sample sheet')
        self.file = file
        head, tail = os.path.split(self.file)
        tail = tail.strip('.xlsx')
        print('tail=',tail)
        logger.debug('filename: ' + self.file)

        if tail.startswith('Charge',0):
            connection = sqlite3.connect(database)
            cursor = connection.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='metadata'")
            if cursor.fetchall() == 0:
                print("no databank table")
            else:
                print("metadatatable ok")

        if tail.startswith('Alle_Proben',0):
            print('metadaten aus bernds datenbank als excelfile importieren')
            self.origin = 'metadata'

            columns = ['name' , 'manr', 'name2','fundland','fundort' ,'fundplatz','datierung','objekt','material', 'skelettelement','tierart',
            'alterm','geschlecht', 'mams','leerespalte','bemerkung','laborleistungen']
            dataframe = read_excel(file, header=None, skiprows=1)
            dataframe.columns = columns
            #dataframe['alterm'] = dataframe['alterm'].fillna(0).astype(int)
            #dataframe['alterm']=dataframe['alterm'].astype(int)
            self.tempdf = dataframe
            print(dataframe)
            self.addtodb()
            return

        if file.endswith('logfile.txt', 0):
            # this is for the temperature logfiles
            # print('templogfile')
            file = open(self.file, mode='r', encoding='utf-8-sig')
            filecontent = file.read()
            filecontent = filecontent.split("\n")
            filecontent = [i.split(';') for i in filecontent]
            print(filecontent)
            cols = ['date', 'seconds', 'temp', 'humid']
            self.tempdf = DataFrame(filecontent, columns=cols)
            self.tempdf['date'] = to_datetime(self.tempdf['date'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
            self.tempdf = self.tempdf[:-1]
            #print(self.tempdf)
            self.origin = 'templog'
            self.addtodb()
            return

        if file.endswith('CN-oD.xlsx', 0):
            logger.debug('CN-oD.xlsx will be imported: ' + file)
            print('CN ohne Drift')
            origin = 'cnod'
            # use specific columns by column names of the excel file only
            # this is still not working since column names are dublicate
            usecolsCNRaw=['Id','Name', 'Sample Type', 'Notes', 'Finish Date', 'EA Sample Weight', 'C Area', 'C %', 'N Area', 'N %',
                            'Height (nA)', 'Area', 'δ¹⁵N (Gas)', 'D/C δ¹⁵N (Gas)', 'δ¹⁵N (Air)', 'Height (nA)', 'Area', 'δ¹³C (Gas)', 'D/C δ¹³C (Gas)', 'δ¹³C (VPDB)']
            # new column names within the dataframe
            columns = ['id', 'name', 'sampletype','notes', 'finishdate','weight', 'areac', 'cpercent', 'arean', 'npercent', 'nheight',
                       'nisoarea', 'n15gas', 'n15drift', 'n15air', 'cheight', 'cisoarea', 'c13gas', 'c13drift',
                       'c13vpdb']
            dataframe = read_excel(file, sheet_name='Batch Report', header=None, engine='openpyxl', skiprows=5)
            dataframe.columns = columns
            # dataframe = dataframe.drop(columns = ['N15gas','C13gas'])
            dataframe['id'] = dataframe['id'].fillna(0).astype(int)
            dataframe['id'] = dataframe['id'].astype(int)
            dataframe = dataframe[:-2]
            self.tempdf = dataframe

        if tail.startswith('OP-oD', 9):
            origin = 'opod'
            print('OP ohne Drift')
            columns = ['id', 'name', 'sampletype','notes', 'finishdate','weight', 'oarea', 'opercent', 'height', 'o18gas', 'o18vsmow']
            # dataframe = read_excel(file, sheet_name='Batch Report',header=None,skiprows=5)
            dataframe = read_excel(file, header=None, engine='openpyxl', skiprows=5)
            dataframe.columns = columns
            dataframe = dataframe[:-2]
            # dataframe['finishdate']=dataframe['finishdate'].astype(str)
            dataframe['id'] = dataframe['id'].fillna(0).astype(int)
            dataframe['id'] = dataframe['id'].astype(int)
            self.tempdf = dataframe

        if tail.startswith('OP-Ro', 9):
            origin = 'opmd'
            print('OP Rohdaten')
            cols = ['Id', 'Name', 'SampleType', 'notes','finishdate','weight', 'PeakID', 'Time', 'Width', 'Height', 'Area', 'ratio2928',
                    'ratio2928raw', 'ratio3028', 'ratio3028raw', 'C13gas', 'Std13C',
                    'O18gas', 'stdcorO18vsmow', 'bsC13gas', 'C13gasdrift', 'C13vpdb', 'stddiffC13', 'bsO18gas',
                    'O18gasdrift', 'O18vsmowmd', 'stddiffdO18']
            self.tempdf = read_excel(file, sheet_name='Batch Report', header=None,engine='openpyxl', skiprows=5)
            self.tempdf.columns = cols
            self.tempdf['superid'] = self.tempdf['Id'].map(str) + self.tempdf['PeakID']

        if file.endswith('CN-Rohdaten.xlsx', 0):
            logger.debug('CN-Rohdate.xlsx will be imported: ' + file)
            print('Import CN Rohdaten')
            origin = 'cnmd'
            # use specific columns by column names of the excel file only
            usecolsCNRaw=['Id','Name', 'Sample Type', 'Notes', 'Finish Date', 'EA Sample Weight', 'Peak Id', 'Retention Time', 'Width',
                    'Height (nA)', 'Area', '29/28', '29/28 Raw (raw)', 'δ¹⁵N (Gas)', 'Std δ¹⁵N (Air)', 'B/S δ¹⁵N (Gas)', 'D/C δ¹⁵N (Gas)',
                    'δ¹⁵N (Air)', 'Std Diff δ¹⁵N (Air)', '45/44', '45/44 Raw (raw)', '46/44', '46/44 Raw (raw)', 'δ¹³C (Gas)', 'Std δ¹³C (VPDB)',
                    'B/S δ¹³C (Gas)', 'D/C δ¹³C (Gas)', 'δ¹³C (VPDB)', 'Std Diff δ¹³C (VPDB)']
            # new column names within the dataframe
            columns = ['Id', 'Name', 'SampleType', 'notes','finishdate','weight', 'PeakID', 'Time', 'Width', 'Height', 'IsoArea',
                       'ratio2928', 'ratio2928raw', 'N15gas', 'stdN15',
                       'bsN15gas', 'N15gasdrift', 'N15air', 'stddiffdN15', 'ratio4544', 'ratio4544raw', 'ratio4644',
                       'ratio4644raw', 'C13gas', 'stdC13',
                       'bsC13gas', 'driftC13gas', 'C13vpdb', 'stddiffC13']
            # read excel file into dataframe
            logger.debug('start reading excel file into dataframe')

            # testing here
            # self.tempdfforheader = read_excel(file, sheet_name='Batch Report', header=0, usecols=usecolsCNRaw, skiprows=3)
            # logger.debug(self.tempdfforheader.head(5))
            # print(self.tempdfforheader.head(5))
            # self.tempdfforheader.columns = columns
            # print(self.tempdfforheader.head(5))

            # self.tempdf = read_excel(file, sheet_name='Batch Report', header=None, skiprows=4)
            self.tempdf = read_excel(file, sheet_name='Batch Report', header=0, usecols=usecolsCNRaw, skiprows=3)
            print(self.tempdf.head(3))
            self.tempdf.columns = columns
            self.tempdf['gastypeCO2'] = pandas.isnull(self.tempdf['N15gas'])

            self.tempdf['superid'] = self.tempdf['Id'].map(str) + self.tempdf['PeakID'].map(str) + self.tempdf[
                'gastypeCO2'].map(str)

            # self.tempdf['superid'] = self.tempdf['Id'].map(str) + self.tempdf['PeakID']
            # self.tempdf['superid'] = self.tempdf['superid'] + str(self.tempdf['gastypeCO2'])

        self.origin = origin
        head, tail = os.path.split(self.file)
        tail = tail.strip('.xlsx')
        self.tempdf['filename'] = tail
        #print(self.tempdf)

        self.finishdatencheck()


class InfoDialog(QDialog):
    newfont = QtGui.QFont("Times", 10, QtGui.QFont.Bold)

    def __init__(self, parent, message):
        logger.debug('perform function')
        super(InfoDialog, self).__init__(parent)
        self.setWindowTitle('Achtung!')
        # self.setFixedSize(180, 100)
        self.layout = QVBoxLayout()
        self.label = QLabel(message)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setFont(self.newfont)
        self.hbox = QHBoxLayout()
        self.closebutton = QPushButton('Abbrechen')
        self.hbox.addWidget(self.closebutton)
        # self.closebutton.setFont(self.newfont)
        self.layout.addWidget(self.label)
        self.closebutton.clicked.connect(self.close)

        if message == 'Es fehlt ein Finishdate':
            self.ignorebutton = QPushButton('Nachtragen')
            self.hbox.addWidget(self.ignorebutton)
            self.ignorebutton.clicked.connect(self.nachtragen)

        self.layout.addLayout(self.hbox)
        self.setLayout(self.layout)
        self.dialogopen = True

    def close(self):
        logger.debug('perform function')
        self.dialogopen = False
        self.done(0)

    def nachtragen(self):
        logger.debug('perform function')
        print(main.samshe.tempdf.finishdate.values)
        # main.samshe.tempdf['finishdate'].fillna('2019-02-28 15:05:10' , inplace=True )
        # main.samshe.tempdf['finishdate'].replace(to_replace= [None], value= data, inplace=True,)
        sum = main.samshe.tempdf.finishdate.isna().sum()
        print('number of na s : ', sum)

        print(main.samshe.tempdf['finishdate'])
        # main.samshe.addtodb()
        self.enterdata(sum)
        self.done(1)

    def enterdata(self, anzahl):
        logger.debug('perform function')
        dialog = EnterDialog(main, anzahl)
        dialog.show()

class DelFrame(object):
    def initUI(self, mainwindow):
        logger.debug('perform function')
        mainwindow.setWindowTitle(mainwindow.title)
        mainwindow.tableview2 = QTableView()

        #mainwindow.resize(1400, 500)
        mainwindow.vbox = QVBoxLayout()
        mainwindow.hbox = QHBoxLayout()
        self.centralwidget = QWidget(mainwindow)
        mainwindow.layout = QVBoxLayout(self.centralwidget)
        # mainwindow.openbtn = QPushButton('Open',self.centralwidget)
        mainwindow.filenamebox = QComboBox()
        mainwindow.items = mainwindow.get_filenames()
        mainwindow.filenamebox.addItems(mainwindow.items)
        mainwindow.delbtn = QPushButton('Delete Run')

        mainwindow.hbox.addWidget(mainwindow.filenamebox)
        mainwindow.vbox.addLayout(mainwindow.hbox)
        #mainwindow.vbox.addWidget(mainwindow.switchbtn)
        mainwindow.layout.addLayout(mainwindow.vbox)
        #mainwindow.delbtn.clicked.connect(lambda: mainwindow.update_combobox())
        mainwindow.filenamebox.currentIndexChanged.connect(lambda: mainwindow.loadsamples())

        mainwindow.delbtn.clicked.connect(lambda: mainwindow.delete_act())
        mainwindow.delbtn.clicked.connect(mainwindow.close)
        mainwindow.delbtn.clicked.connect(mainwindow.changelayout_delrun)
        mainwindow.layout.addLayout(mainwindow.vbox)
        mainwindow._main = QWidget()
        layout = QVBoxLayout(mainwindow._main)
        mainwindow.vbox.addWidget(mainwindow.tableview2)
        mainwindow.vbox.addWidget(mainwindow.delbtn)
        mainwindow.setCentralWidget(self.centralwidget)
        mainwindow.loadsamples()


class SamplesFrame(object):

    def initUI(self, mainwindow):
        logger.debug('perform function')
        mainwindow.setWindowTitle(mainwindow.title)
        mainwindow.tableview2 = QTableView()

        #mainwindow.resize(1400, 500)
        mainwindow.vbox = QVBoxLayout()
        mainwindow.hbox = QHBoxLayout()
        self.centralwidget = QWidget(mainwindow)
        mainwindow.layout = QVBoxLayout(self.centralwidget)
        # mainwindow.openbtn = QPushButton('Open',self.centralwidget)
        mainwindow.filenamebox = QComboBox()
        mainwindow.nextbtn = QPushButton('->',self.centralwidget)
        mainwindow.prevbtn = QPushButton('<-',self.centralwidget)
        #mainwindow.switchbtn = QPushButton('Check Standards', self.centralwidget)
        #mainwindow.switchbtn.clicked.connect(mainwindow.changelayout2)
        # mainwindow.switchbtn.clicked.connect(mainwindow.loadmoredata)
        mainwindow.items = mainwindow.get_filenames()
        mainwindow.filenamebox.addItems(mainwindow.items)

        mainwindow.hbox.addWidget(mainwindow.filenamebox)
        mainwindow.hbox.addStretch(1)
        mainwindow.hbox.addWidget(mainwindow.prevbtn)
        mainwindow.hbox.addWidget(mainwindow.nextbtn)
        mainwindow.vbox.addLayout(mainwindow.hbox)
        #mainwindow.vbox.addWidget(mainwindow.switchbtn)
        mainwindow.layout.addLayout(mainwindow.vbox)
        mainwindow.filenamebox.currentIndexChanged.connect(lambda: mainwindow.loadsamples())
        mainwindow.filenamebox.currentIndexChanged.connect(lambda: mainwindow.canvasupdate())
        mainwindow.prevbtn.clicked.connect(mainwindow.prev)
        mainwindow.nextbtn.clicked.connect(mainwindow.next)
        mainwindow.layout.addLayout(mainwindow.vbox)
        mainwindow._main = QWidget()
        layout = QVBoxLayout(mainwindow._main)
        self.sc = MyCanvas(self.centralwidget, width=5, height=4, dpi=dpi)
        toolbar = NavigationToolbar(self.sc, self.sc)
        mainwindow.splitter = QSplitter(QtCore.Qt.Vertical)
        mainwindow.splitter.addWidget(self.sc)
        mainwindow.splitter.addWidget(mainwindow.tableview2)
        mainwindow.splitter.setSizes([800, 0])
        mainwindow.vbox.addWidget(mainwindow.splitter)
        # mainwindow.vbox.addWidget(toolbar)
        mainwindow.setCentralWidget(self.centralwidget)
        mainwindow.loadsamples()
        mainwindow.canvasupdate()

    def open_list(self, model):
        logger.debug('perform function')
        self.dialog = NewDialog(self, 'list')
        self.dialog.tableview.setModel(model)
        #self.dialog.resize(900, 350)
        #self.dialog.move(100, 100)
        self.dialog.show()



class OutputFrame(object):

    def __init__(self):
        logger.debug('perform function')
        self.loadbar=LoadDialog()


    def initUI(self, mainwindow):
        logger.debug('perform function')

        mainwindow.setWindowTitle(mainwindow.title)
        #mainwindow.resize(1400, 900)
        mainwindow.filenameboxout = QComboBox()
        mainwindow.items = mainwindow.get_filenames()
        mainwindow.filenameboxout.addItems(mainwindow.items)

        mainwindow.centralwidget = QWidget()  # mainwindow
        mainwindow.setCentralWidget(mainwindow.centralwidget)

        mainwindow.savebtn = QPushButton('Save')
        mainwindow.saveinputline = QLineEdit()

        self.get_all_data(mainwindow)
        self.get_tabslist()
        mainwindow.tabs = MyTableWidget(self.tabslist)

        policy = mainwindow.filenameboxout.sizePolicy()
        policy.setHorizontalPolicy(QSizePolicy.Expanding)
        mainwindow.filenameboxout.setSizePolicy(policy)

        mainwindow.vbox = QVBoxLayout(mainwindow.centralwidget)
        mainwindow.hbox = QHBoxLayout()

        mainwindow.vbox.addWidget(mainwindow.filenameboxout)
        mainwindow.vbox.addLayout(mainwindow.hbox)
        mainwindow.vbox.addWidget(mainwindow.tabs)

        mainwindow.hbox.addWidget(mainwindow.saveinputline)
        mainwindow.hbox.addWidget(mainwindow.savebtn)

        self.init_tabs(mainwindow)

        mainwindow.filenameboxout.currentIndexChanged.connect(lambda: self.init_tabs(mainwindow))
        mainwindow.savebtn.clicked.connect(lambda: self.save_excel(mainwindow))


    def init_tabs(self,mainwindow):
        logger.debug('perform function')
        mainwindow.saveinputline.setText("output\\" + str(mainwindow.filenameboxout.currentText()[:-3]+".xls"))

        #self.get_all_data(mainwindow)
        self.get_tabslist()

        mainwindow.vbox.removeWidget(mainwindow.tabs)
        mainwindow.tabs.deleteLater()

        mainwindow.tabs = MyTableWidget(self.tabslist)
        mainwindow.vbox.addWidget(mainwindow.tabs)

        self.get_all_data(mainwindow)
        self.set_tabsdata(mainwindow)

    def get_tabslist(self):
        logger.debug('perform function')
        self.datadf['project'] = self.datadf['name'].astype(str).str[:4]      # namens-col auf die vier ersten zeichen reduzieren
        self.tabslist = self.datadf.drop_duplicates(subset = ['project'],inplace=True)

        self.tabslist = self.datadf['project']
        #print("suchnamen:", self.suchnamen1)
        print("tabslist:" , self.tabslist)

        print("typ: ",type(self.tabslist))

        self.tabslist = self.datadf['project'].to_list()
        try:
            for i in self.suchnamen1:
                if any(i in s for s in self.tabslist):
                    print("delete :" , i)
                    self.tabslist.remove(i)
        except:
            print("Not enough data available")

        self.tabslist.append('Standards')
        self.tabslist.append('Alle Daten')
        print("tabslist cleaned:" , self.tabslist)


    def set_tabsdata(self,mainwindow):
        logger.debug('perform function')

        for i in self.tabslist:

            self.tabdata = self.datadf.loc[self.datadf['name'].str.match(i)]
            if i == "Standards":
                model = self.stdmodel
            elif i == "Alle Daten":
                model = self.modelall
            else:
                model = PandasModel(self.tabdata)

            self.child = mainwindow.findChild(QWidget, i)
            self.child.tableview.setModel(model)             #model i
            #mainwindow.tabs.ruih.tableview.setModel(model)


    def get_all_data(self,mainwindow):
        logger.debug('perform function')
        connection = sqlite3.connect(database)
        cursor = connection.cursor()
        # alledaten = "select id,name,finishdate,o18gas,o18vsmowod from opodtable "
        #alledaten = "select finishdate,o18vsmowod,id,filename from opodtable where name like '%" + name + "%'"

        filename = mainwindow.filenameboxout.currentText()
        print('filename :'+filename)
        filename2 = filename[:11]
        filename3 = filename2[-2:]

        if filename3 != 'CN':
            alledaten = "select id as id , name as name,sampletype as 'Sampletype', peakid as ' Peak ID',time AS Time,width AS Width,height AS Height ,area AS Area,ratio2928 AS '29/28',ratio2928raw AS '29/28 \n (Raw)',ratio3028 AS '30/28',ratio3028raw AS '30/28 \n (Raw)',c13gas AS 'd13C (gas)' ,o18gas AS 'd18O (Gas)',o18gasdrift AS 'D/C d18O \n (Gas)',o18vsmowod AS 'd18O \n (V-SMOW)', o18vsmowmd as 'D/C d18O \n (V-SMOW)',stddiffdo18 AS  'Std Diff d18O \n (V-SMOW)',opercent AS 'O%'   from (select * from opmd where filename like '"+filename2+"-Rohdaten' and peakid = 'S1') table1\
                    LEFT JOIN (select id,name as name1,o18vsmowod, opercent,finishdate as finishdate1 from opod where filename like '" + filename +"') table2 ON table1.finishdate =table2.finishdate1"

        #

            #rawtablecn = "CREATE TABLE rawtablecn (idraw INT,superid text PRIMARY KEY,name text NOT NULL,sampletype text,finishdate TIMESTAMP CURRENT_TIMESTAMP,peakid INT,time TIMESTAMP DEFAULT CURRENT_TIMESTAMP,width TIMESTAMP DEFAULT CURRENT_TIMESTAMP,height FLOAT,area FLOAT,ratio2928 FLOAT,ratio2928raw FLOAT,n15gas FLOAT,stdn15air FLOAT,n15gasdrift FLOAT,n15aircali FLOAT,ratio4544 FLOAT,ratio4544raw FLOAT,ratio4644 FLOAT,ratio4644raw FLOAT,c13gas FLOAT,stdc13vpdb FLOAT,c13gasdrift FLOAT,c13vpdbcali FLOAT,runid TEXT,filename TEXT,quality INT,final INT, comment text, extra INT)"


            #cursor.execute("INSERT OR IGNORE INTO cnodtable(id,sampletype,name,finishdate,areac,cpercent,arean,npercent,nheight,nisoarea,n15gas,n15drift,n15air,cheight,cisoarea,c13gas,c13drift,c13vpdb,filename)\
                                                    #    SELECT id,sampletype,name,finishdate,areac,cpercent,arean,npercent,nheight,nisoarea,n15gas,n15drift,n15air,cheight,cisoarea,c13gas,c13drift,c13vpdb,filename FROM temptable")



            datadf = read_sql(alledaten, con=connection, index_col=None, coerce_float=True, params=None,parse_dates=None,columns=None, chunksize=None)
            cursor.close()
            connection.commit()
            #print('datadf: ',datadf)
            self.datadf = main.get_triples(datadf,tocalculate='O%',tocalculate2='d18O \n (V-SMOW)',tocalculate3='D/C d18O \n (V-SMOW)')

                                    # namen um tabs auszufiltern --> standards und blanks aus tablist entfernen
            self.suchnamen1 = ['Ag3PO4','Benzoic Acid','benzoic','IAEA-601','IAEA','nbs','iaea','NBS 120ci','HA-MAI','HA M', 'HaMA','Su-DAN','SU-D','NBS ','SU DAN','SU-DAN','Ag3P','Benz','pre_','Blnk','benz','Nbs']
                                    #namen um standards zu erkennen

            stdnamen = ['Ag3PO4','Benzoic Acid','benzoic','IAEA-601','IAEA','nbs','iaea','NBS 120ci','HA-MAI','HA M', 'HaMA','Su-DAN']

            stdcollection = DataFrame()

            for i in stdnamen:
                resultdf = self.datadf[self.datadf['name'].str.contains(i)]            #standards auslesen
                stdcollection = stdcollection.append(resultdf)

        else:
             alledaten = "select id as id, name as name, cpercent AS 'C%',npercent AS 'N%', time AS Time, width AS Width, height AS Height, area AS Area, ratio2928 AS '29/28',ratio2928raw AS '29/28 (Raw)' ,n15gas AS 'dN15 (gas)',n15gasdrift ,n15aircali,ratio4544 ,ratio4544raw,n15air,c13vpdb,c13vpdbcali from (select * from cnmd where filename like '"+filename2+"-Rohdaten' and peakid = 'S1') table1\
                    LEFT JOIN (select id , name as name1,finishdate AS finishdate1 ,areac,cpercent,arean,npercent,nheight,cheight,areac,cpercent,arean,npercent,nheight,nisoarea,n15gas AS n15gas2,n15drift,n15air,cheight \
                    ,cisoarea,c13gas,c13drift,c13vpdb from cnod where filename like '" + filename +"') table2 ON table1.finishdate =table2.finishdate1  LEFT JOIN (select name AS metaname,manr , fundland , fundort , fundplatz ,datierung, skelettelement ,tierart, geschlecht , bemerkung,mams,altermin from metadata)table3 ON table2.name1 = table3.manr"

             self.datadf = read_sql(alledaten, con=connection, index_col=None, coerce_float=True, params=None,parse_dates=None,columns=None, chunksize=None)





             self.datadf = main.get_triples(self.datadf,tocalculate='C%',tocalculate2='N%',tocalculate3='n15air',tocalculate4='n15aircali')
             self.datadf = main.get_triples(self.datadf,tocalculate='c13vpdb',tocalculate2='c13vpdbcali' )

             cols =['id','name','N%','N%avg','N%std','C%','C%avg','C%std','n15air','n15airavg','n15airstd',
             'n15aircali','n15aircaliavg','n15aircalistd','c13vpdb','c13vpdbavg','c13vpdbstd','c13vpdbcali','c13vpdbcaliavg','c13vpdbcalistd']

             contextcols =['manr' , 'fundland' , 'fundort' , 'fundplatz' ,'datierung', 'skelettelement' ,'tierart', 'geschlecht' , 'bemerkung','mams','altermin']
             self.datadf = self.datadf[cols]
             # c13vpdbcali und n15aircali auf eine row ziehen + doppelte row löschen
             self.datadf['c13vpdbcali'] = self.datadf['c13vpdbcali'].shift(-1)
             self.datadf.drop_duplicates(subset=['id','n15air','c13vpdb'] , inplace=True)

             progressfinish = len(self.datadf.index)
             print('zu verarbeiten:' , progressfinish)

             for i in self.datadf['name']:
                progressfinish -=1
                print('zu verarbeiten:', progressfinish)
                self.loadbar.state(progressfinish)
                self.loadbar.show()

                #print('ist name manr?' , manr)

                #print('name_i:' , i)
                if i.startswith('MA-'):
                   #print('name = manr')
                    try:

                       scrapedf = pandas.read_html('http://192.168.123.50/secure/DBCEZA/Suche/suche_proben_tab.php?Proben1='+i+'&Proben2=30&Save03=Suche+in+Tab+Proben',timeout=0.01)

                       print('SCRAPED:' ,scrapedf)
                       #print('SCRAPEDCOLS:')
                       #print('coltext',scrapedf[0].columns[9])

                       #Fundland-Spalte usw. Original Name von AM-81db holen, falls änderung safety
                       fundland= scrapedf[0].columns[9]
                       fundort=scrapedf[0].columns[10]
                       fundplatz=scrapedf[0].columns[11]
                       datierung=scrapedf[0].columns[12]
                       skelettelement=scrapedf[0].columns[15]
                       tierart=scrapedf[0].columns[14]
                       geschlecht=scrapedf[0].columns[9]
                       bemerkung=scrapedf[0].columns[13]
                       altermin=scrapedf[0].columns[13]
                       scrapedf=scrapedf[0]
                       #print('fundland: ' ,scrapedf.loc[0,fundland])
                       self.datadf.loc[self.datadf.name==i,fundland] =scrapedf.loc[0,fundland]
                       self.datadf.loc[self.datadf.name==i,fundort] =scrapedf.loc[0,fundort]
                       self.datadf.loc[self.datadf.name==i,fundplatz] =scrapedf.loc[0,fundplatz]
                       self.datadf.loc[self.datadf.name==i,datierung] =scrapedf.loc[0,datierung]
                       self.datadf.loc[self.datadf.name==i,skelettelement] =scrapedf.loc[0,skelettelement]
                       self.datadf.loc[self.datadf.name==i,tierart] =scrapedf.loc[0,tierart]

                       #folgend veraltet: aus excelfile einlesen

                       #self.datadf.loc[self.datadf.name == i, 'manr'] = i
                       #context = "select manr ,fundland,fundort,fundplatz, datierung,skelettelement,tierart,geschlecht,bemerkung,altermin ,mams from metadata where manr like '" + i + "'"

                       #contextdf = read_sql(context, con=connection, index_col=None, coerce_float=True, params=None,parse_dates=None,columns=None, chunksize=None)
                       #print('contextdf:' , contextdf['fundland'])
                       #self.datadf.loc[self.datadf.name == i,'Fundland'] = contextdf.fundland[0]
                       #self.datadf.loc[self.datadf.name == i,'Fundort'] = contextdf.fundort[0]
                       #self.datadf.loc[self.datadf.name == i,'Fundplatz'] = contextdf.fundplatz[0]
                       #self.datadf.loc[self.datadf.name == i,'Datierung'] = contextdf.datierung[0]
                       #self.datadf.loc[self.datadf.name == i,'Skelettelement'] = contextdf.skelettelement[0]
                       #self.datadf.loc[self.datadf.name == i,'Tierart'] = contextdf.tierart[0]
                       #self.datadf.loc[self.datadf.name == i,'Geschlecht'] = contextdf.geschlecht[0]
                       #self.datadf.loc[self.datadf.name == i,'bemerkung'] = contextdf.bemerkung[0]
                       #self.datadf.loc[self.datadf.name == i,'altermin'] = contextdf.altermin[0]
                    except:
                        print('keine connection zur Web DB - MA-NR nicht gefunden')
                        pass

                if i.startswith('MAMS'):
                   self.datadf.loc[self.datadf.name == i, 'mams'] = i[6:]
                   try:
                       scrapedf = pandas.read_html('http://192.168.123.50/secure/DBCEZA/Suche/suche_proben_tab.php?Proben1='+i[6:]+'&Proben2=40&Save03=Suche+in+Tab+Proben',timeout=0.01)


                       #print('SCRAPED:' ,scrapedf)
                       #print('SCRAPEDCOLS:')
                       #print('coltext',scrapedf[0].columns[9])

                       #Fundland-Spalte usw. Original Name von AM-81db holen, falls änderung safety
                       fundland= scrapedf[0].columns[9]
                       fundort=scrapedf[0].columns[10]
                       fundplatz=scrapedf[0].columns[11]
                       datierung=scrapedf[0].columns[12]
                       skelettelement=scrapedf[0].columns[15]
                       tierart=scrapedf[0].columns[14]
                       geschlecht=scrapedf[0].columns[9]
                       bemerkung=scrapedf[0].columns[13]
                       altermin=scrapedf[0].columns[13]
                       scrapedf=scrapedf[0]
                       #print('fundland: ' ,scrapedf.loc[0,fundland])
                       self.datadf.loc[self.datadf.name==i,fundland] =scrapedf.loc[0,fundland]
                       self.datadf.loc[self.datadf.name==i,fundort] =scrapedf.loc[0,fundort]
                       self.datadf.loc[self.datadf.name==i,fundplatz] =scrapedf.loc[0,fundplatz]
                       self.datadf.loc[self.datadf.name==i,datierung] =scrapedf.loc[0,datierung]
                       self.datadf.loc[self.datadf.name==i,skelettelement] =scrapedf.loc[0,skelettelement]
                       self.datadf.loc[self.datadf.name==i,tierart] =scrapedf.loc[0,tierart]
                   except:
                       print('keine connection zur Web DB MAMS Nr nicht gefunden')
                       pass
             #'manr' , 'fundland' , 'fundort' , 'fundplatz' ,'datierung', 'skelettelement' ,'tierart', 'geschlecht' , 'bemerkung','mams','altermin'

             self.suchnamen1 = ['USGS 40', 'USGS-40', 'usgs 40', 'usgs-40','USGS 41', 'USGS-41', 'usgs 41', 'usgs-41','USGS 41a', 'USGS-41a', 'usgs 41a', 'usgs-41a','USGS 41 a', 'USGS-41 a', 'usgs 41 a', 'usgs-41 a',
                                'USGS 43', 'USGS-43', 'usgs 43', 'usgs-43','IAEA CH6', 'IAEA-CH6', 'iaea ch6', 'iaea-ch6', 'iaea-ch-6', 'IAEA-CH-6', 'IAEA CH-6', 'iaea ch-6',
                                'IAEA CH7', 'IAEA-CH7', 'iaea ch7', 'iaea-ch7', 'iaea-ch-7', 'IAEA-CH-7', 'IAEA CH-7', 'iaea ch-7','IAEA N1', 'IAEA-N1', 'iaea n1', 'iaea-n1', 'iaea-n-1', 'IAEA-N-1', 'IAEA N-1', 'iaea n-1',
                                'IAEA N2', 'IAEA-N2', 'iaea n2', 'iaea-n2', 'iaea-n-2', 'IAEA-N-2', 'IAEA N-2', 'iaea n-2','sulfanilamid', 'sulfanilamide', 'Sulfanilamid', 'Sulfanilamide']
                                    #namen um standards zu erkennen

             stdnamen = ['USGS 40', 'USGS-40', 'usgs 40', 'usgs-40','USGS 41', 'USGS-41', 'usgs 41', 'usgs-41','USGS 41a', 'USGS-41a', 'usgs 41a', 'usgs-41a','USGS 41 a', 'USGS-41 a', 'usgs 41 a', 'usgs-41 a',
                                'USGS 43', 'USGS-43', 'usgs 43', 'usgs-43','IAEA CH6', 'IAEA-CH6', 'iaea ch6', 'iaea-ch6', 'iaea-ch-6', 'IAEA-CH-6', 'IAEA CH-6', 'iaea ch-6',
                                'IAEA CH7', 'IAEA-CH7', 'iaea ch7', 'iaea-ch7', 'iaea-ch-7', 'IAEA-CH-7', 'IAEA CH-7', 'iaea ch-7','IAEA N1', 'IAEA-N1', 'iaea n1', 'iaea-n1', 'iaea-n-1', 'IAEA-N-1', 'IAEA N-1', 'iaea n-1',
                                'IAEA N2', 'IAEA-N2', 'iaea n2', 'iaea-n2', 'iaea-n-2', 'IAEA-N-2', 'IAEA N-2', 'iaea n-2','sulfanilamid', 'sulfanilamide', 'Sulfanilamid', 'Sulfanilamide']

             stdcollection = DataFrame()

             for i in stdnamen:
                resultdf = self.datadf[self.datadf['name'].str.contains(i)]            #standards auslesen
                stdcollection = stdcollection.append(resultdf)

        stdcollection = stdcollection.sort_values(by=['name'])
        self.stdcollection = stdcollection.drop_duplicates()
        #print('collection:' , stdcollection)
        self.stdmodel = PandasModel(self.stdcollection)
        self.modelall = PandasModel(self.datadf)

    def get_triples(self,dataframe):
        logger.debug('perform function')

        pos = 0
        anfangspos = 0
        self.resultdf = resultdf
        self.resultdf['set'] = ''
        aktuell = self.resultdf.id[0]
        self.triple = []

        for number in self.resultdf.id:
            # ist die id der folgenden zeile actid+1 und finsihdateact - new < 15min, notiere die id, ansonsten die der nächsten
            # ist der name der folgende zeile Blnk überspringe diese

            if number == aktuell:
                self.triple.append(self.resultdf.id[anfangspos])
                pos += 1

            else:
                self.triple.append(self.resultdf.id[pos])
                anfangspos = pos
                pos += 1
                if inhalt != 'Blnk':
                    aktuell = inhalt

        self.resultdf.set = self.triple
        sort_values(by=['col1'])


    def save_excel(self,mainwindow):
       #todrop=['Time','Width','Height','Area','29/28']
       #self.datadf.drop(columns=todrop,inplace=True)
       #cols =['id','name','N%','N%avg','N%std','C%','C%avg','C%std','n15air','n15airavg','n15airstd',
       #'n15aircali','n15aircaliavg','n15aircalistd','c13vpdb','c13vpdbavg','c13vpdbstd','c13vpdbcali','c13vpdbcaliavg','c13vpdbcalistd']
       #self.datadf = self.datadf[cols]

       try:
            savename = mainwindow.saveinputline.text()

            writer = ExcelWriter(savename, engine='xlsxwriter')
            workbook = writer.book

            #length_list = [len(x) for x in cols]
           # for i, width in enumerate(length_list):
                #self.datadf.set_column(i, i, width)

            #text_fmt = workbook.add_format({'align': 'center'})
            deci_fmt = workbook.add_format({'num_format': '0.00' ,'align': 'center'})
                                                                                                    # Write each dataframe to a different worksheet.
            for i in self.tabslist:
                tabdata = self.datadf.loc[self.datadf['name'].str.match(i)]
                if i == "Standards":
                    self.stdcollection.to_excel(writer,index = False, sheet_name=i)

                elif i == "Alle Daten":
                    self.datadf.to_excel(writer,index = False, sheet_name=i)
                else:
                    tabdata.to_excel(writer,index = False,sheet_name= i)
                writer.sheets[i].set_column('B:B',22)
                writer.sheets[i].set_column('C:AE',20, deci_fmt )

            # Close the Pandas Excel writer and output the Excel file.
            writer.save()
            print('saved data as ' + savename + ' in programm folder')
       except:
            print("can't save..file open in EXCEL ?! ")


class MyTableWidget(QWidget):

    def __init__(self, tabslist, parent =None):
        logger.debug('perform function')
        super(QWidget, self).__init__(parent)
        self.layout = QVBoxLayout()
        self.tabswidget = QTabWidget()
        self.create_tabs(tabslist)

    def create_tabs(self,tabslist):
        logger.debug('perform function')

        for i in tabslist:
            # Initialize tab screen
            self.tab = QWidget()

            self.tabswidget.addTab(self.tab,i)  # Add tabs
            self.tab.layout = QVBoxLayout(self)
            self.tab.tableview = QTableView()

            self.tab.layout.addWidget(self.tab.tableview)


            self.tab.setLayout(self.tab.layout)
            self.tab.setObjectName(i)

        self.tabswidget.resize(300,200)

        self.layout.addWidget(self.tabswidget)
        self.setLayout(self.layout)

class StandardsFrame(object):
    '''
    Window that displays the standards in multiple plots
    '''
    def initUI(self, mainwindow):
        logger.debug('perform function')
        logger.debug('generating UI for Standards Window')
        mainwindow.setWindowTitle(mainwindow.title)
        #mainwindow.resize(1400, 900)

        self.centralwidget = QWidget()  # mainwindow
        mainwindow.tableview = QTableView()
        # self.tableview.setFixedWidth(500)

        # mainwindow.label = QLabel('Standards')
        # mainwindow.label.setAlignment(QtCore.Qt.AlignCenter)
        # mainwindow.label.setFont(QtGui.QFont('SansSerif', 13))
        mainwindow.layout = QVBoxLayout(self.centralwidget)
        # mainwindow.layout.addWidget(mainwindow.label)
        mainwindow.nextbtn = QPushButton('>>')

        mainwindow.vbox = QVBoxLayout()
        mainwindow.hbox = QHBoxLayout()
        mainwindow.grid = QGridLayout()

        mainwindow.filenamebox = QComboBox()
        mainwindow.switchbtn = QPushButton('Check Samples')
        mainwindow.switchbtn.clicked.connect(mainwindow.changelayout)

        mainwindow.items = mainwindow.get_filenames()
        mainwindow.filenamebox.addItems(mainwindow.items)

        mainwindow.figure = plt.figure()
        # mainwindow.figure.set_facecolor("lightyellow")

        mainwindow.hbox.addWidget(mainwindow.filenamebox)
        mainwindow.vbox.addLayout(mainwindow.hbox)
        # mainwindow.vbox.addWidget(mainwindow.switchbtn)

        mainwindow.cnframe = QFrame(mainwindow)     # frame that holds the CN-Data when a CN run is selected
        mainwindow.cnframe2 = QFrame(mainwindow)     # frame that holds the CN-Data when a CN run is selected
        mainwindow.opframe = QFrame(mainwindow)     # frame that holds the Oxygen-Data when a O run is selected
        mainwindow.stdzoomframe = QFrame(mainwindow)  # this is the fram that is shown at the bottom of the standards window holding detailed data

        # define the different layouts depending on the data and whether they are Cn or O Data
        self.ly = QGridLayout(mainwindow.cnframe)
        self.ly2 = QGridLayout(mainwindow.opframe)
        self.ly3 = QGridLayout(mainwindow.cnframe2)
        self.ly3.setColumnStretch(1, 1)
        self.ly3.setRowStretch(1, 1)
        self.ly4 = QVBoxLayout(mainwindow.stdzoomframe)  # layout of the detailed data

        # dropdown list that defines the number of past days to be plotted
        mainwindow.daylabel = QLabel('Days')
        mainwindow.dropdown = QComboBox(mainwindow.stdzoomframe)
        policy = mainwindow.dropdown.sizePolicy()
        policy.setHorizontalPolicy(QSizePolicy.Expanding)
        mainwindow.dropdown.setSizePolicy(policy)
        daylist =['60','90','180','360','720']  # days of data that should be displayed
        mainwindow.dropdown.addItems(daylist)
        mainwindow.dropdown.currentIndexChanged.connect(lambda: mainwindow.daychanged(mainwindow.dropdown.currentText()))
        
        # insert widgets into the layout of the window
        mainwindow.vbox.addWidget(mainwindow.cnframe)
        mainwindow.vbox.addWidget(mainwindow.cnframe2)
        mainwindow.vbox.addWidget(mainwindow.opframe)
        mainwindow.vbox.addWidget(mainwindow.stdzoomframe)
        mainwindow.hbox.addWidget(mainwindow.nextbtn)
        mainwindow.hbox.addStretch(1)
        mainwindow.hbox.addWidget(mainwindow.daylabel)
        mainwindow.hbox.addWidget(mainwindow.dropdown)
        mainwindow.nextbtn.clicked.connect(mainwindow.nextframe)
        mainwindow.dropdown.hide()
        mainwindow.daylabel.hide()

        # list of standards that will be filtered out and their various spellings
        ag3 = ['Ag3PO4']
        benz = ['Benzoic Acid', 'benzoic acid', 'benzoic acid end', 'Benzoic Acid End', 'benz', 'Benz']
        iaea601 = ['IAEA 601', 'IAEA-601', 'iaea 601', 'iaea-601']
        iaea602 = ['IAEA 602', 'IAEA-602', 'iaea 602', 'iaea-602']
        hama = ['HA MA ', 'HA-MA ']
        nbs = ['NBS 120c ', 'nbs 120c ', 'NBS-120c ', 'nbs-120c ', 'NBS 120 c ', 'nbs 120 c ', 'NBS-120 c ',
               'nbs-120 c ', ]
        sudan = ['SU-DAN', 'SU DAN', 'su-dan', 'su dan']
        hamai = ['HA MAI', 'HA-MAI', 'HA MAi', 'HA-MAi']
        nbsi = ['NBS 120ci', 'nbs 120ci', 'NBS-120ci', 'nbs-120ci', 'NBS 120 ci', 'nbs 120 ci', 'NBS-120 ci',
                'nbs-120 ci', ]
        sulfanil = ['sulfanilamid', 'sulfanilamide', 'Sulfanilamid', 'Sulfanilamide']
        usgs40 = ['USGS 40', 'USGS-40', 'usgs 40', 'usgs-40']
        usgs41 = ['USGS 41', 'USGS-41', 'usgs 41', 'usgs-41','USGS 41a', 'USGS-41a', 'usgs 41a', 'usgs-41a','USGS 41 a', 'USGS-41 a', 'usgs 41 a', 'usgs-41 a']
        usgs43 = ['USGS 43', 'USGS-43', 'usgs 43', 'usgs-43']
        iaeach6 = ['IAEA CH6', 'IAEA-CH6', 'iaea ch6', 'iaea-ch6', 'iaea-ch-6', 'IAEA-CH-6', 'IAEA CH-6', 'iaea ch-6']
        iaeach7 = ['IAEA CH7', 'IAEA-CH7', 'iaea ch7', 'iaea-ch7', 'iaea-ch-7', 'IAEA-CH-7', 'IAEA CH-7', 'iaea ch-7']
        iaean1 = ['IAEA N1', 'IAEA-N1', 'iaea n1', 'iaea-n1', 'iaea-n-1', 'IAEA-N-1', 'IAEA N-1', 'iaea n-1']
        iaean2 = ['IAEA N2', 'IAEA-N2', 'iaea n2', 'iaea-n2', 'iaea-n-2', 'IAEA-N-2', 'IAEA N-2', 'iaea n-2']
        iaeano3 = ['IAEA-NO-3','IAEA NO3', 'iaea no3', 'iaea-no3', 'IAEA-N-O3', 'IAEA N-O3', 'iaea-n-O3']

        # bulding groupboxes for each standard with the plot of the recent data and the sollwerte 
        mainwindow.ag3po4 = GroupBox(ag3, '21.7', typ='CO', parent=mainwindow.opframe)
        mainwindow.benzoic = GroupBox(benz, '23.37', typ='CO', parent=mainwindow.opframe)
        mainwindow.iaea601 = GroupBox(iaea601, '23.3', typ='CO', parent=mainwindow.opframe)
        #mainwindow.iaea602 = GroupBox(iaea602, '71.4', typ='CO', parent=mainwindow.opframe)
        mainwindow.hama = GroupBox(hama, '17.1', typ='CO', parent=mainwindow.opframe)
        mainwindow.nbs = GroupBox(nbs, '22.2', typ='CO', parent=mainwindow.opframe)
        mainwindow.sudan = GroupBox(sudan, '14.3', typ='CO', parent=mainwindow.opframe)
        mainwindow.hamai = GroupBox(hamai, '17.1', typ='CO', parent=mainwindow.opframe)
        mainwindow.nbsi = GroupBox(nbsi, '22.2', typ='CO', parent=mainwindow.opframe)
        mainwindow.usgs43co = GroupBox(usgs43, '11.50', typ='CO', parent=mainwindow.opframe)

        mainwindow.sulfanil = GroupBox(sulfanil, '-2.55', typ='CN', sollwert2='-28.17', parent=mainwindow.cnframe)
        mainwindow.usgs40 = GroupBox(usgs40, '-4.5', typ='CN', sollwert2='-26.389', parent=mainwindow.cnframe)
        mainwindow.usgs41 = GroupBox(usgs41, '47.6', typ='CN', sollwert2='37.626', parent=mainwindow.cnframe)
        mainwindow.usgs43 = GroupBox(usgs43, '8.44', typ='CN', sollwert2='-21.28', parent=mainwindow.cnframe)
        mainwindow.iaeach6 = GroupBox(iaeach6, '0', typ='CN', sollwert2='-10.45 ', parent=mainwindow.cnframe)
        mainwindow.iaeach7 = GroupBox(iaeach7, '0', typ='CN', sollwert2='-32.15 ', parent=mainwindow.cnframe)
        mainwindow.iaean1 = GroupBox(iaean1, '0.4', typ='CN', sollwert2='0', parent=mainwindow.cnframe2)
        mainwindow.iaean2 = GroupBox(iaean2, '20.3', typ='CN', sollwert2='0', parent=mainwindow.cnframe2)
        mainwindow.iaeano3 = GroupBox(iaeano3, '4.7', typ='CN', sollwert2='0', parent=mainwindow.cnframe2)

        mainwindow.canvas = MyCanvas(width=14, height=8, parent=mainwindow.stdzoomframe)
        # mainwindow.mousePressEvent = mainwindow.closestdzoom
        mainwindow.canvas.mouseDoubleClickEvent = mainwindow.close_zoomframe    # when double click on plot close plot
        mainwindow.canvas.setToolTip('Zum Schließen -Doppelklick')
        self.toolbar = NavigationToolbar(mainwindow.canvas, mainwindow.canvas)

        # layout for Oxygen mode
        self.ly2.addWidget(mainwindow.ag3po4, 0, 0)
        self.ly2.addWidget(mainwindow.benzoic, 0, 1)
        self.ly2.addWidget(mainwindow.sudan, 0, 2)
        self.ly2.addWidget(mainwindow.iaea601, 1, 0)
        self.ly2.addWidget(mainwindow.usgs43co, 2, 0)
        self.ly2.addWidget(mainwindow.hama, 1, 1)
        self.ly2.addWidget(mainwindow.hamai, 2, 1)
        self.ly2.addWidget(mainwindow.nbs, 1, 2)
        self.ly2.addWidget(mainwindow.nbsi, 2, 2)

        #layout for Cn mode
        self.ly.addWidget(mainwindow.sulfanil, 1, 2)
        self.ly.addWidget(mainwindow.usgs40, 0, 0)
        self.ly.addWidget(mainwindow.usgs41, 0, 1)
        self.ly.addWidget(mainwindow.usgs43, 0, 2)
        self.ly.addWidget(mainwindow.iaeach6, 1, 0)
        self.ly.addWidget(mainwindow.iaeach7, 1, 1)
        self.ly3.addWidget(mainwindow.iaean1, 0, 0)
        self.ly3.addWidget(mainwindow.iaean2, 0, 1)
        self.ly3.addWidget(mainwindow.iaeano3, 0, 2)

        # layout for the detial plot of one single standard
        self.ly4.addWidget(mainwindow.canvas)
        self.ly4.addWidget(mainwindow.dropdown)
        self.ly4.addWidget(mainwindow.daylabel)

        mainwindow.layout.addLayout(mainwindow.vbox)

        # hide the layouts that are not needed
        mainwindow.opframe.hide()
        mainwindow.cnframe.hide()
        mainwindow.cnframe2.hide()
        mainwindow.stdzoomframe.hide()

        # load data
        mainwindow.loadsheet()
        mainwindow.filenamebox.currentIndexChanged.connect(lambda: mainwindow.loadsheet())
        mainwindow.filenamebox.currentIndexChanged.connect(lambda: mainwindow.calculatesheet())
        # mainwindow.openbtn.clicked.connect(mainwindow.loadmoredata)
        mainwindow.loadsheet()
        mainwindow.calculatesheet()
        mainwindow.setCentralWidget(self.centralwidget)

class ProgressThread(QThread):
    """
    Runs a counter thread.
    """


    countChanged = pyqtSignal(int)
    count = 100

    def run1(self,count):
        logger.debug('perform function')
        self.countChanged.emit(count)


class LoadDialog(QDialog):
    def __init__(self,parent=None):
        logger.debug('perform function')
        super(LoadDialog, self).__init__(parent)
        self.title='loading...'
        self.progressbar = QProgressBar()

        self.layout = QHBoxLayout(self)
        self.layout.addWidget(self.progressbar)
        self.progressbar.setValue(100)
        self.show()
        QApplication.processEvents()
        self.thread= ProgressThread()
        self.thread.start()
        self.thread.countChanged.connect(self.setprogress)

    def setprogress(self,value):
        logger.debug('perform function')
        self.progressbar.setValue(value)

    def state(self,count):
        logger.debug('perform function')
        self.thread.run1(100-count)


class SampleDialog(QMainWindow):
    def __init__(self, parent):
        logger.debug('perform function')
        super(SampleDialog, self).__init__(parent)
        self.title = 'Standards & Samples'
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowMinimizeButtonHint )

    def closestdzoom(self, e):
        logger.debug('perform function')
        print('closestdzoomframe')
        self.stdzoomframe.hide()
        self.opframe.show()

    def changelayout(self):
        logger.debug('perform function')
        self.title = 'Samples'
        self.actframe = 'smplframe'
        self.SAF = SamplesFrame()
        self.SAF.initUI(self)
        self.show

    def changelayout2(self):
        logger.debug('perform function')
        self.title = 'Standards'
        self.actframe = 'stdframe'
        self.SF = StandardsFrame()
        self.SF.initUI(self)
        self.show

    def changelayout3(self):
        logger.debug('perform function')
        self.title = 'Output'
        self.actframe = 'outputframe'
        self.OF = OutputFrame()
        self.OF.initUI(self)
        self.show()

    def changelayout_delrun(self):
        logger.debug('perform function')
        self.title='delete runs'
        self.actframe = 'delruns'
        self.delframe = DelFrame()
        self.delframe.initUI(self)
        self.show()

    def changelayoutstd(self):
        logger.debug('perform function')
        self.title = 'Standard ZOOM'

    def update(self, dataframe, typ):
        logger.debug('perform function')
        if typ == 'CO':
            self.ag3po4.update(dataframe)
            self.benzoic.update(dataframe)
            self.iaea601.update(dataframe)
            self.usgs43co.update(dataframe)
            self.hama.update(dataframe)
            self.nbs.update(dataframe)
            self.hamai.update(dataframe)
            self.nbsi.update(dataframe)
            self.sudan.update(dataframe)
        if typ == 'CN':
            self.sulfanil.update(dataframe)
            self.usgs40.update(dataframe)
            self.usgs41.update(dataframe)
            self.usgs43.update(dataframe)
            self.iaeach6.update(dataframe)
            self.iaeach7.update(dataframe)
            self.iaean1.update(dataframe)
            self.iaean2.update(dataframe)
            self.iaeano3.update(dataframe)

        #print('updated:', dataframe)

    def update_combobox(self):
        logger.debug('perform function')
        self.filenamebox.clear()
        self.items = self.get_filenames()
        self.filenamebox.addItems(self.items)
        #self.filenamebox.setCurrentIndex(-1)

    def loadmoredata(self, name):
        logger.debug('perform function')
        connection = sqlite3.connect(database)
        cursor = connection.cursor()
        # alledaten = "select id,name,finishdate,o18gas,o18vsmowod from opodtable "
        alledaten = "select finishdate,o18vsmowod,id,filename from opod where name like '%" + name + "%'"
        moredatadf = read_sql(alledaten, con=connection, index_col=None, coerce_float=True, params=None,
                              parse_dates=None,
                              columns=None, chunksize=None)
        cursor.close()
        connection.commit()
        return moredatadf

    def loadsheet(self):
        '''
        main function that loads data from the database
        '''
        logger.debug('perform function')
        logger.debug('loading data that belong to the selected run')

        # figure out what run data to load according to the selected run in the window
        self.curtext = self.filenamebox.currentText()   #Name of the runfile selected in the window
        self.toload = self.curtext[:11]
        # self.toload = '20190512_OP-Rohdaten-2'
        logger.debug('run to load: ' + self.toload)
        print('toload2: ',self.toload)
        #self.mode = self.toload[-2:]
        
        #self.toloadraw = self.curtext.split('_')
        toloadraw1 = self.curtext.split('-')
        self.toloadraw = toloadraw1[0]+'-Rohdaten'
        logger.debug('raw data to load: ' + self.toloadraw)
        
        y=len(toloadraw1)
        if (y >2) :
            self.toloadraw2 = self.toloadraw +'-'+ toloadraw1[2]
        else:
            self.toloadraw2 = self.toloadraw
        logger.debug('toloadraweeend: ' + self.toloadraw2)
        print('toloadraweeend: ' ,self.toloadraw2)

        # find the mode CN or O
        self.mode = self.toload.split('_',1)
        self.mode = self.mode[1]
        #self.mode = self.mode[1:2]
        print('mmodeerscheinung ', self.mode)
        logger.debug('mmodeerscheinung: ' + self.mode)
        #print('toload3: ',toload)
        if self.mode == 'CN':
            print('CN')
            sqlselect = "select id,name,finishdate,c13vpdb,n15air,n15gas,c13gas from cnod where filename LIKE '%" + self.curtext + "%' order by id "
            driftdf = "select finishdate,c13vpdbcali,n15aircali, peakid, time, width from cnmd where filename LIKE '%" + self.toloadraw2 + "%'"
        else:
            sqlselect = "select id,name,finishdate,o18gas,o18vsmowod from opod where filename LIKE '%" + self.curtext + "%' order by id "
            driftdf = "select finishdate,o18vsmowmd, peakid, time, width,height, area, ratio2928, ratio2928raw, ratio3028, ratio3028raw, c13gas from opmd where filename LIKE '%" + self.toloadraw2 + "%'"
        
        logger.debug('SQL QUERY // sqlselect: ' + sqlselect)
        logger.debug('SQL QUERY // driftdf: ' + driftdf)
        connection = sqlite3.connect(database)
        cursor = connection.cursor()

        alledaten = "select id,name,finishdate,o18gas,o18vsmowod from opod "

        pandas.set_option('display.max_columns', 30)
        logger.debug('start reading sql into pandas')
        self.resultdf1 = read_sql(sqlselect, con=connection, index_col=None, coerce_float=True, params=None,
                                  parse_dates=None,
                                  columns=None, chunksize=None)
        logger.debug(self.resultdf1.head(5))
        self.driftdf = read_sql(driftdf, con=connection, index_col=None, coerce_float=True, params=None,
                                parse_dates=None,
                                columns=None, chunksize=None)
        logger.debug(self.driftdf.head(5))
        logger.debug('start reading sql into pandas...done')
        # self.driftdf = self.driftdf[self.driftdf.peakid != 'M1']
        # self.driftdf = self.driftdf[isfinite(self.driftdf['o18vsmowmd'])]
        logger.debug('merging dataframes resultsdf and driftdf using finishdate as merger')
        self.resultdf = pandas.merge(self.resultdf1, self.driftdf, on='finishdate')
        self.driftdf = self.driftdf[self.driftdf.peakid != 'M1']    # remove all M1 peaks
        self.resultdf = pandas.merge(self.resultdf1, self.driftdf, on='finishdate')
        #print("resultdf merged without M1: ", self.resultdf)
        cursor.close()
        connection.commit()

        # get the triplets of the selected run
        # self.resultdf['set']  = ''
        logger.debug('calculating triples')
        if self.mode == 'CN':
            self.resultdf = main.get_triples(self.resultdf, tocalculate='c13vpdb', tocalculate2='c13vpdbcali',
                                             tocalculate3='n15air', tocalculate4='n15aircali')
            print('toloadcn')
            logger.debug('toloadcn')
        if self.mode == 'OP':
            self.resultdf = main.get_triples(self.resultdf, tocalculate='o18vsmowod', tocalculate2='o18vsmowmd')
            print('toloadco')
            logger.debug('toloadco')

        # get_triples(self, resultdf, gleiches='set', tocalculate='o18vsmowod', tocalculate2='o18vsmowmd', )
        # self.driftdf =  main.get_triples(self.driftdf)
        #print("resultdf triples: ", self.resultdf)

        sulfanil = ['sulfanilamid', 'sulfanilamide', 'Sulfanilamid', 'Sulfanilamide']

        if self.actframe == 'stdframe':
            if self.mode == 'OP':
                self.cnframe.hide()
                self.cnframe2.hide()
                self.update(self.resultdf, typ='CO')
                self.opframe.show()
            if self.mode == 'CN':
                self.opframe.hide()
                self.update(self.resultdf, typ='CN')
                self.cnframe.show()

    def nextframe(self):
        logger.debug('perform function')
        if self.cnframe.isVisible():
            self.cnframe.hide()
            self.cnframe2.show()
            return
        if self.cnframe2.isVisible():
            self.cnframe2.hide()
            self.cnframe.show()
        if self.stdzoomframe.isVisible():
            self.stdzoomframe.hide()
            self.dropdown.hide()
            main.dialog.loadsheet()

    def close_zoomframe(self, event):
        logger.debug('perform function')
        if self.stdzoomframe.isVisible():
            self.stdzoomframe.hide()
            self.dropdown.hide()
            self.daylabel.hide()
            self.nextbtn.show()
            main.dialog.loadsheet()


    def zoomframe(self, name, typ, days):
        logger.debug('perform function')
        if self.cnframe.isVisible():
            self.cnframe.hide()
        if self.cnframe2.isVisible():
            self.cnframe2.hide()
        if self.opframe.isVisible():
            self.opframe.hide()

        self.canvas.updatestdzoom(name, typ,days)

        self.dropdown.show()
        self.daylabel.show()

        self.nextbtn.hide()
        self.stdzoomframe.show()

    def daychanged(self,days):
        logger.debug('perform function')
        self.canvas.changedays(days)
        self.stdzoomframe.show()

    def set_model(self, dataframe):
        logger.debug('perform function')
        model = PandasModel(dataframe)
        self.tableview.setModel(model)
        # self.show()

    def loadsamples(self):
        logger.debug('perform function')
        self.loadsheet()

        #print(type(self.resultdf))
        self.resultdfall = self.resultdf
        self.resultdf.drop_duplicates(subset=['set'], inplace=True)
        self.dataframe = self.resultdf.reset_index(drop=True)

        model = PandasModel(self.dataframe)
        self.tableview2.setModel(model)

    def delete_act(self):
        logger.debug('perform function')
        if self.mode == 'CN':
            print('CN-delete')
            sqlselect = "delete from cnod where filename LIKE '" + self.curtext + "'  "
            driftdf = "delete from cnmd where filename LIKE '" + self.toloadraw2 + "'"
        else:
            print('OP-delete')
            sqlselect = "delete from opod where filename LIKE '" + self.curtext + "'  "
            driftdf = "delete from opmd where filename LIKE '" + self.toloadraw2 + "'"
        connection = sqlite3.connect(database)
        cursor = connection.cursor()

        cursor.execute(sqlselect)
        cursor.execute(driftdf)
        cursor.close()
        connection.commit()
        print('delete')

    def calculatesheet(self):
        logger.debug('perform function')
        if self.mode == 'OP':
            self.resultdf = main.get_triples(self.resultdf)
        if self.mode == 'CN':
            self.resultdf = main.get_triples(self.resultdf, tocalculate='n15air', tocalculate2='n15aircali', tocalculate3='c13vpdb', tocalculate4='c13vpdbcali')

        self.set_model(self.resultdf)

        # self.canvas.draw()

    #  self.update(self.resultdf)

    def get_filenames(self):
        logger.debug('perform function')
        connection = sqlite3.connect(database)
        cursor = connection.cursor()
        filenames = "select distinct filename from opod union select distinct filename from cnod order by filename desc"
        cursor.execute(filenames)
        filenamelist = cursor.fetchall()
        filenamelist = [x[0] for x in filenamelist]
        connection.commit()
        connection.close()
        #print(filenamelist)
        return filenamelist

    def canvasupdate(self):
        logger.debug('perform function')
        print('bishier')

        #self.SAF.sc.update_samplecanvas(self.dataframe)
        self.SAF.sc.update_samplecanvas(self.resultdf)

    def prev(self):
        logger.debug('perform function')
        index = self.filenamebox.currentIndex()
        newindex = index +1
        self.filenamebox.setCurrentIndex(newindex)

    def next(self):
        logger.debug('perform function')
        index = self.filenamebox.currentIndex()
        if index >0:
            newindex = index - 1
        else:
            newindex = index
        self.filenamebox.setCurrentIndex(newindex)

class oQLabel(QLabel):
    def __init__(self, tresh=0.3, parent=None):
        logger.debug('perform function')
        super(oQLabel, self).__init__(parent)
        self.tresh = tresh

    def set_color(self):
        logger.debug('perform function')
        # val = abs(float(self.text()))

        if self.text() == '0% N2' or self.text() == '0% CO2' or abs(float(self.text())) < self.tresh:
            self.setStyleSheet("background-color:#00ff00;")
        else:

            self.setStyleSheet("background-color:#ff0000;")


class GroupBox(QGroupBox):
    def __init__(self, name, sollwert, typ, sollwert2=None, parent=None):
        logger.debug('perform function')
        logger.debug('generating standards groupbox with: name=' + name[0] + ', sollwert=' + sollwert + ', typ=' + typ)
        super(GroupBox, self).__init__(parent)
        self.typ = typ
        self.figure = plt.figure(dpi=dpi)
        self.canvas = FigureCanvas(self.figure)
        self.canvas.setToolTip('Achtung Gaswerte sind geplottet! Nicht mit normierten Werten (ohne/mit drift) verwechseln!!')

        self.sollwert = sollwert
        self.sollwert2 = sollwert2
        self.setTitle(name[0])
        self.suchnamen = name
        self.realname = name
        self.name = name[0]
        self.textbox0 = QLabel('Kalkulation:')
        self.textbox3 = QLabel('Sollwert: ')
        self.textbox2 = QLabel('Mittelwert: ')

        self.textbox1 = QLabel('Standardabw.:')
        self.textbox7 = QLabel('Differenz:')
        self.textbox8 = oQLabel()
        self.textbox6 = QLabel()
        self.textbox6b = QLabel()
        self.textbox5 = QLabel()
        self.textbox4 = oQLabel()
        self.textbox11 = QLabel()
        self.textbox12 = oQLabel()
        self.textbox13 = oQLabel()
        self.textbox9 = QLabel('Single Group ')
        self.textbox10 = QLabel('Drift ')
        self.textbox16 = oQLabel()
        self.textbox17 = QLabel()
        self.textbox18 = QLabel()
        self.textbox19 = oQLabel()
        self.textbox20 = oQLabel()
        self.textbox21 = QLabel()
        self.textbox22 = QLabel()
        self.textbox23 = oQLabel()

        self.textbox6.setText(sollwert)
        self.textbox6b.setText(sollwert)

        hbox1 = QHBoxLayout()
        hbox = QHBoxLayout()
        # hbox.addWidget(self.textbox4)
        vbox = QVBoxLayout()
        grid = QGridLayout()
        if self.typ == 'CN':
            self.textbox18.setText(sollwert2)
            self.textbox22.setText(sollwert2)
            self.textbox9 = QLabel('Single Group N2')
            self.textbox10 = QLabel('Drift N2')
            self.textbox14 = QLabel('Single Group  CO2')
            self.textbox15 = QLabel('Drift CO2')
            grid.addWidget(self.textbox14, 0, 3)
            grid.addWidget(self.textbox15, 0, 4)
            # C13 Werte
            grid.addWidget(self.textbox16, 1, 3)
            grid.addWidget(self.textbox17, 2, 3)
            grid.addWidget(self.textbox18, 3, 3)
            grid.addWidget(self.textbox19, 4, 3)
            grid.addWidget(self.textbox20, 1, 4)
            grid.addWidget(self.textbox21, 2, 4)
            grid.addWidget(self.textbox22, 3, 4)
            grid.addWidget(self.textbox23, 4, 4)

        # grid.addWidget(self.textbox0, 0, 0)
        grid.addWidget(self.textbox9, 0, 1)
        grid.addWidget(self.textbox10, 0, 2)

        grid.addWidget(self.textbox1, 1, 0)
        grid.addWidget(self.textbox2, 2, 0)
        grid.addWidget(self.textbox3, 3, 0)
        grid.addWidget(self.textbox7, 4, 0)
        # O18 werte oder N15 Werte
        grid.addWidget(self.textbox4, 1, 1)
        grid.addWidget(self.textbox5, 2, 1)
        grid.addWidget(self.textbox6, 3, 1)
        grid.addWidget(self.textbox8, 4, 1)
        grid.addWidget(self.textbox12, 1, 2)
        grid.addWidget(self.textbox11, 2, 2)
        grid.addWidget(self.textbox6b, 3, 2)
        grid.addWidget(self.textbox13, 4, 2)

        # grid.addWidget(self.box, 1, 1)
        # grid.addWidget(self.textbox6, 2, 1)
        hbox1.addWidget(self.canvas)
        vbox.addLayout(hbox1)
        vbox.addLayout(hbox)
        vbox.addLayout(grid)
        self.setLayout(vbox)

        self.canvas.mousePressEvent = self.open  # lambda event: print('released')

    def open(self, e):
        logger.debug('perform function')
        print('open', self.name)
        main.dialog.zoomframe(self.suchnamen, self.typ,main.dialog.dropdown.currentText())

    def plot(self):
        logger.debug('perform function')

        plotdf = self.get_mwabw(self.suchnamen, returndf=True)

        plt.style.use('seaborn-darkgrid')
        self.figure.clear()

        # ax2 = ax1.twinx()
        # ax3 = ax1.twinx()

        # self.figure.suptitle(self.name, fontsize=12, fontweight='bold')
        # ax.hold(True) # deprecated, see above

        if self.typ == 'CO':
            ax1 = self.figure.add_subplot(111)
            try:
                realname = plotdf['name'].values[0]
                print('realname ist:', realname)
                ax1.set_title(realname)
            except:
                #print('standard fehlt')
                ax1.set_title('No Data')
                #self.hide()
            ax1.set_xlabel('id')
            ax1.set_ylabel('o18gas')
            ax1.ticklabel_format(useOffset=False)
            ax1.scatter(plotdf['id'], plotdf['o18gas'])
            self.figure.subplots_adjust(0.12, 0.2, 0.95, 0.85, 0, 0)

        if self.typ == 'CN':
            ax1 = self.figure.add_subplot(212)
            ax2 = self.figure.add_subplot(211)
            try:
                self.realname = plotdf['name'].values[0]
                print('realname ist:', self.realname)
                ax2.set_title(self.realname)

                    #self.textbox6.setText(sollwert)
                    #self.textbox6b.setText(sollwert)

            except:
                #print('standard fehlt')
                ax2.set_title('No Data')

            ax2.xaxis.set_major_locator(MaxNLocator(integer=True))
            ax1.ticklabel_format(useOffset=False)
            ax2.ticklabel_format(useOffset=False)
            # ax2 = ax1.twinx()
            ax1.set_xlabel('id')
            ax1.set_ylabel('c13gas')
            ax2.set_xlabel('id')
            ax2.set_ylabel('n15gas')
            ax1.scatter(plotdf['id'], plotdf['c13gas'])
            ax2.scatter(plotdf['id'], plotdf['n15gas'], color='green')
            # ax2.xaxis.set_major_locator(MaxNLocator(integer=True))
            self.figure.subplots_adjust(left=0.15, bottom=None, right=0.95, top=None, wspace=None, hspace=0.35)
        # ax2.scatter(plotdf['id'],plotdf['o18vsmowod'])
        # ax2.scatter(plotdf['id'], plotdf['o18vsmowmd'])

        # self.figure.subplots_adjust(left=None, bottom=None, right=None, top=None, wspace=None, hspace=None)
        ax1.xaxis.set_major_locator(MaxNLocator(integer=True))
        self.canvas.draw()
        # self.update()

    def update(self, dataframe):
        logger.debug('perform function')
        self.data = dataframe
        values = self.get_mwabw(self.suchnamen)
        avg = str(round(values[0], 2))
        std = str(round(values[1], 2))
        avg2 = str(round(values[2], 2))
        std2 = str(round(values[3], 2))

        prediff = round((float(values[0]) - float(self.sollwert)), 2)
        diff = str(prediff)
        if prediff > 0:
            diff = '+' + str(prediff)

        prediff2 = round((float(values[2]) - float(self.sollwert)), 2)
        diff2 = str(prediff2)
        if prediff2 > 0:
            diff2 = '+' + str(prediff2)

        #self.figure.set_facecolor("lightyellow")
        if not self.name == 'IAEA CH6' and not self.name == 'IAEA CH7':
            if values[1] > 0.3:
                self.figure.set_facecolor("red")

        if self.typ == 'CO':
            self.textbox5.setText(avg)
            self.textbox4.setText(std)
            self.textbox8.setText(diff)
            self.textbox11.setText(avg2)
            self.textbox12.setText(std2)
            self.textbox13.setText(diff2)
            self.textbox8.set_color()
            self.textbox13.set_color()
            self.textbox4.set_color()
            self.textbox12.set_color()

        if self.typ == 'CN':

            if self.realname in ['USGS 41a', 'USGS-41a', 'usgs 41a', 'usgs-41a', 'USGS 41 a', 'USGS-41 a', 'usgs 41 a',
                            'usgs-41 a']:
                self.sollwert2 = '36.55'
                self.textbox18.setText(self.sollwert2)
                self.textbox22.setText(self.sollwert2)

            avg3 = str(round(values[4], 2))
            std3 = str(round(values[5], 2))
            avg4 = str(round(values[6], 2))
            std4 = str(round(values[7], 2))
            diff3 = str(round((float(values[4]) - float(self.sollwert2)), 2))
            diff4 = str(round((float(values[6]) - float(self.sollwert2)), 2))
            if self.name == "IAEA CH6" or self.name == "IAEA CH7":
                diff = '0% N2'
                diff2 = '0% N2'
            if self.name == "IAEA N1" or self.name == "IAEA N2":
                diff3 = '0% CO2'
                diff4 = '0% CO2'

            self.textbox5.setText(avg)
            self.textbox4.setText(std)
            self.textbox8.setText(diff)
            self.textbox11.setText(avg2)
            self.textbox12.setText(std2)
            self.textbox13.setText(diff2)

            self.textbox17.setText(avg3)
            self.textbox16.setText(std3),
            self.textbox19.setText(diff3)
            self.textbox20.setText(std4)
            self.textbox21.setText(avg4)
            self.textbox23.setText(diff4)
            self.textbox8.set_color()
            self.textbox13.set_color()
            self.textbox19.set_color()
            self.textbox23.set_color()
            self.textbox4.set_color()
            self.textbox12.set_color()
            self.textbox16.set_color()
            self.textbox20.set_color()

        self.plot()
        self.show

    def get_mwabw(self, suchnamen, returndf=False):
        logger.debug('perform function')

        if self.typ == 'CO':
            value1 = 'o18vsmowod'
            value2 = 'o18vsmowmd'
        if self.typ == 'CN':
            value3 = 'c13vpdb'
            value4 = 'c13vpdbcali'
            value1 = 'n15air'
            value2 = 'n15aircali'

        #print(name)
        # resultdf = self.data[self.data['name'].str.contains(name)]
        # resultdf1 =  self.data.loc[self.data['name'] == name]
        print('######self.data: ',self.data)

        resultdf = self.data.loc[self.data['name'].isin(suchnamen)]
        print('######Resultdf1: ', resultdf)
        if resultdf.empty:
            for i in suchnamen:
                    resultdf = self.data[self.data['name'].str.contains(i)]
                    print('######Resultdf2: ', resultdf)
                    if resultdf.empty == False:
                        break

        mean = resultdf[value1].mean()
        abw = resultdf[value1].std()
        mean2 = resultdf[value2].mean()
        abw2 = resultdf[value2].std()
        values = [mean, abw, mean2, abw2]
        if self.typ == 'CN':
            mean3 = resultdf[value3].mean()
            abw3 = resultdf[value3].std()
            mean4 = resultdf[value4].mean()
            abw4 = resultdf[value4].std()
            values = [mean, abw, mean2, abw2, mean3, abw3, mean4, abw4]

        if returndf:
            return (resultdf)

        #print(resultdf)

        return (values)


class SqlDialog(QDialog):
    newfont = QtGui.QFont("Times", 10, QtGui.QFont.Bold)

    def __init__(self, parent, database):
        logger.debug('perform function')
        super(SqlDialog, self).__init__(parent)
        self.label = QLabel('Insert SQL Syntax')
        self.inputline = QTextEdit()
        self.inputline.setFont(self.newfont)
        # self.inputline.setFixedHeight(50)
        self.outputline = QTableView()
        self.runquerybtn = QPushButton('>>>')
        self.purequerybtn = QPushButton('execute sql')
        self.plotbtn = QPushButton('Plot')
        self.savebtn = QPushButton('Save')
        self.saveinputline = QLineEdit()
        self.plotbtn.setToolTip(
            'Plot der ersten beiden Spalten als X,Y. Farben werden abwechseln wenn sich \'filename\' in den Ergebnissen befindet ')
        self.layout = QVBoxLayout(self)
        self.hbox = QHBoxLayout(self)
        self.hboxsave = QHBoxLayout(self)
        self.resize(800, 300)
        self.move(700, 100)
        self.layout.addWidget(self.label)
        # self.splitter = QSplitter(QtCore.Qt.Vertical)
        # self.splitter.addWidget(self.inputline)
        self.layout.addWidget(self.inputline)
        self.hbox.addWidget(self.runquerybtn)
        self.hbox.addWidget(self.purequerybtn)
        self.hboxsave.addWidget(self.saveinputline)
        self.hboxsave.addWidget(self.savebtn)
        self.layout.addLayout(self.hbox)
        self.layout.addWidget(self.outputline)
        self.layout.addWidget(self.plotbtn)
        self.layout.addLayout(self.hboxsave)
        self.database = database
        self.runquerybtn.clicked.connect(self.runquery)
        self.purequerybtn.clicked.connect(self.purequery)
        self.plotbtn.clicked.connect(self.plot)
        self.savebtn.clicked.connect(self.savetoexcel)
        self.inputline.setText(
            'select finishdate,o18gas,name,filename from opod where strftime(\'%Y-%m\' ,finishdate) like \'2019-03\' and name = "Ag3PO4" order by id')
        self.saveinputline.setText('Export/resplorExport.xlsx')

    def runquery(self):
        logger.debug('perform function')
        try:
            connection = sqlite3.connect(database)
            cursor = connection.cursor()
            sql = self.inputline.toPlainText()
            #print(sql)
            self.sql = sql
            self.resultdf = read_sql(sql, con=connection, index_col=None, coerce_float=True, params=None,
                                     parse_dates=None, columns=None, chunksize=None)
            # resultdf['id']= resultdf['id'].astype(int)
            # print(resultdf)
            cursor.close()
            connection.commit()
            model = PandasModel(self.resultdf)
            self.outputline.setModel(model)
        except Exception as e:
            message = 'bad sql sytax    : ' + str(e)
            main.open_infodialog(message)

    def purequery(self):
        logger.debug('perform function')
        try:
            connection = sqlite3.connect(database)
            cursor = connection.cursor()
            sql = self.inputline.toPlainText()
            print('execute : ' + sql)
            cursor.execute(sql)
            cursor.close()
            connection.commit()
        except Exception as e:
            message = 'probs with your query    : ' + str(e)
            main.open_infodialog(message)
            print("probs with your query")

    def plot(self):
        logger.debug('perform function')
        try:
            plotwindow = SearchWindow()
            plotwindow.plotprepare(self.sql, 'SQL Query results')
            plotwindow.show()
        except:
            message = 'Bitte zuerst SQL Query ausführen'
            main.open_infodialog(message)

    def savetoexcel(self):
        logger.debug('perform function')
        try:
            savename = self.saveinputline.text()

            writer = ExcelWriter(savename, engine='xlsxwriter')

            # Write each dataframe to a different worksheet.

            self.resultdf.to_excel(writer,index = False, sheet_name='Resplor Export')
            self.resultdf.to_excel(writer, sheet_name='Resplor Export2')

            # Close the Pandas Excel writer and output the Excel file.
            writer.save()
            print('saved data as ' + savename + ' in programm folder')
        except:
            print("can't save")


class EnterDialog(QDialog):
    logger.debug('perform function')
    def __init__(self, parent, anzahl):
        super(EnterDialog, self).__init__(parent)
        self.anzahl = anzahl
        self.setWindowTitle('Eingabe erforderlich')
        self.acceptbtn = QPushButton('Ok')
        self.acceptbtn.clicked.connect(self.enternewdate)
        self.layout = QVBoxLayout()
        hbox = QHBoxLayout()
        self.label = QLabel('Bitte Finishdates nachtragen: z.B. 2019-02-28 12:09:48')

        i = 0
        self.layout.addWidget(self.label)
        self.objektnamen = list()
        while i < self.anzahl:
            i += 1
            objname = 'input' + str(i)

            labeltext = 'Finishdate' + str(i)
            print('objektname = ', objname)
            label = QLabel(labeltext)
            self.inputline = QLineEdit()
            self.inputline.setObjectName(objname)
            self.objektnamen.append(self.inputline)
            hbox.addWidget(label)
            hbox.addWidget(self.inputline)
            self.layout.addLayout(hbox)
            self.layout.addWidget(self.acceptbtn)
        self.setLayout(self.layout)
        model = PandasModel(main.samshe.tempdf)
        main.open_new_dialog('Es fehlt ein finishdate', 'standard', model)

    def enternewdate(self):
        logger.debug('perform function')
        print(self.objektnamen)
        newdates = list()
        i = 0
        while i < len(self.objektnamen):
            newdate = self.objektnamen[i].text()
            newdates.append(newdate)
            i += 1
            print(newdate)
        main.samshe.replacedata(newdates)

        model = PandasModel(main.samshe.tempdf)
        main.dialog.tableview.setModel(model)
        main.dialog.show()
        self.done(1)


class NewDialog(QDialog):
    def __init__(self, parent, title):
        logger.debug('perform function')
        super(NewDialog, self).__init__(parent)
        self.setWindowTitle(title)
        self.delzerobtn = QPushButton()
        self.tableview = QTableView()
        self.layout = QVBoxLayout(self)
        self.layout.addWidget(self.tableview)
        self.setLayout(self.layout)
        self.dialogopen = True

    def closeEvent(self, event):
        logger.debug('perform function')
        self.dialogopen = False


class PandasModel(QtCore.QAbstractTableModel):
    def __init__(self, df=DataFrame(), parent=None):
        logger.debug('perform function')
        QtCore.QAbstractTableModel.__init__(self, parent=parent)
        self._df = df

    def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
        logger.debug('perform function')
        if role != QtCore.Qt.DisplayRole:
            return QtCore.QVariant()

        if orientation == QtCore.Qt.Horizontal:
            try:
                return self._df.columns.tolist()[section]
            except (IndexError,):
                return QtCore.QVariant()
        elif orientation == QtCore.Qt.Vertical:
            try:
                # return self.df.index.tolist()
                return self._df.index.tolist()[section]
            except (IndexError,):
                return QtCore.QVariant()

    def data(self, index, role=QtCore.Qt.DisplayRole):
        logger.debug('perform function')
        if role != QtCore.Qt.DisplayRole:
            return QtCore.QVariant()

        if not index.isValid():
            return QtCore.QVariant()

        return QtCore.QVariant(str(self._df.iloc[index.row(), index.column()]))

    def setData(self, index, value, role):
        logger.debug('perform function')
        row = self._df.index[index.row()]
        col = self._df.columns[index.column()]
        if hasattr(value, 'toPyObject'):
            # PyQt4 gets a QVariant
            value = value.toPyObject()
        else:
            # PySide gets an unicode  ..used
            dtype = self._df[col].dtype
            if dtype != object:
                value = None if value == '' else dtype.type(value)
        self._df.set_value(row, col, value)
        return True

    def rowCount(self, parent=QtCore.QModelIndex()):
        logger.debug('perform function')
        return len(self._df.index)

    def columnCount(self, parent=QtCore.QModelIndex()):
        logger.debug('perform function')
        return len(self._df.columns)

    def sort(self, column, order):
        logger.debug('perform function')
        colname = self._df.columns.tolist()[column]
        self.layoutAboutToBeChanged.emit()
        self._df.sort_values(colname, ascending=order == QtCore.Qt.AscendingOrder, inplace=True)
        self._df.reset_index(inplace=True, drop=True)
        self.layoutChanged.emit()


class MyCanvas(FigureCanvas):
    """Ultimately, this is a QWidget (as well as a FigureCanvasAgg, etc.)."""

    def __init__(self, parent=None, width=8, height=4, dpi=dpi):
        logger.debug('perform function')
        self.fig = plt.figure(figsize=(width, height), dpi=dpi)
        #self.fig.set_facecolor("lightyellow")
        plt.style.use('seaborn-darkgrid')

        self.axes = self.fig.add_subplot(111)

        # self.axes2 = self.axes.twinx()
        self.compute_initial_figure()
        self.fig.subplots_adjust(0.05, 0.1, 0.95, 0.9, )
        FigureCanvas.__init__(self, self.fig)
        self.setParent(parent)

        FigureCanvas.setSizePolicy(self,
                                   QSizePolicy.Expanding,
                                   QSizePolicy.Expanding)
        FigureCanvas.updateGeometry(self)



    def plotprepare(self, plottable, title, typ):
        logger.debug('perform function')
        if typ == 'CN' or typ =='CO':
            self.title = title
            try:
                self.plotdf = plottable
                # self.temphumiddf = read_sql(temphumidvalues, con = connection, coerce_float=True, params=None, parse_dates=None, chunksize=None)
                self.plotdfcols = self.plotdf.columns
                # plotdf['finishdate'].replace(to_replace=[None], value = "2018.11.11 11:11:11", inplace=True)
                self.plotdf = self.plotdf.sort_values(by=self.plotdf.columns[0])

                if set(['date']).issubset(self.plotdf.columns):
                    self.plotdf['date'] = to_datetime(self.plotdf.date, format='%Y/%m/%d %H:%M:%S')
                    self.plotdf = self.plotdf.sort_values(by='date')
                    self.plotdf = self.plotdf.iloc[1::20, :]

                if set(['finishdate']).issubset(self.plotdf.columns):
                    self.plotdf['finishdate'] = to_datetime(self.plotdf.finishdate)
                    self.plotdf = self.plotdf.sort_values(by='finishdate')  # sort_byvalues  inplace=True, ascending=False
                if set(['id']).issubset(self.plotdf.columns):
                    self.plotdf = self.plotdf.sort_values(by=['id'])

                # self.plotdf = self.plotdf[pd.notnull(self.plotdf[self.plotdf.columns[1]])]  # get rid off nan
                # sortieren nach col0 also finsishdate oder id
                self.plotdf = self.plotdf.reset_index(drop=True)  # index wieder richtig setzen
                # plotdf = plotdf[(plotdf. != 0).any()]     'remove zeros'
                # print(len(plotdfcols))
                # instead of ax.hold(False)
                # self.plotdf.dropna(axis=0, how='any', inplace=True)    #get rid of na an zeros
                self.plotdf = self.plotdf.replace(0, nan)

                #if main.checkbox:

                    #self.makeplot(typ, title=title, plottemp=True)
                #else:

                   # self.makeplot(typ, title=title, plottemp=False)
                self.makeplot(typ, title=title, plottemp=False)
            except Exception as e:
                message = 'Fehler bei der Ploterstellung: ' + str(e)
                main.open_infodialog(message)
        else:
            self.plotdf = plottable
            print('U:',self.plotdf)
            self.makeplot(typ, title=title, plottemp=False)

    def makeplot(self, typ, title='', plottemp=False):
        logger.debug('perform function')
        self.plottemp = plottemp
        plt.style.use('seaborn-darkgrid')
        # model = PandasModel(self.plotdf)
        # Fenstername = 'Plotdata: ' + title
        # main.open_new_dialog(Fenstername, 'plot', model)
        print(self.plotdf)
        self.figure.clear()
        if typ == 'CO':
            ax = self.figure.add_subplot(111)
            xlabel = self.plotdf.columns[0]
            ylabel = self.plotdf.columns[1]
            ax.set_xlabel(xlabel,fontsize=16)
            ax.set_ylabel(ylabel, fontsize=16)

        if typ == 'CN':
            ax = self.figure.add_subplot(211)
            ax2 = self.figure.add_subplot(212)
            ylabel = self.plotdf.columns[1]
            x2label = self.plotdf.columns[0]
            y2label = self.plotdf.columns[2]
            ax2.set_xlabel(x2label, fontsize=16)
            ax2.set_ylabel(y2label, fontsize=16)
            ax.set_ylabel(ylabel, fontsize=16)

        if typ == 'CN' or typ == 'CO':
            if 'filename' in self.plotdf.columns:
                print("filename listed to plot")
                print(self.plotdf.filename[0])
                newlist = list(self.plotdf.filename)
                aktuell = self.plotdf.filename[0]
                farbe = 'k'
                farbliste = []

                for inhalt in newlist:
                    if inhalt == aktuell:
                        print(farbe + inhalt)
                        farbliste.append(farbe)
                    else:
                        if farbe == 'k':
                            farbe = 'r'
                        else:
                            farbe = 'k'
                        farbliste.append(farbe)
                    aktuell = inhalt

                self.plotdf['farben'] = farbliste

                print(self.plotdf)
                grouped = self.plotdf.groupby('farben')
                blackdf = grouped.get_group('k')
                print(grouped.get_group('k'))
                reddf = grouped.get_group('r')

                start, end = ax.get_xlim()
                stepsize = 2
                if (self.plotdf[self.plotdf.columns[0]].dtype == "datetime64[ns]"):
                    plt.xticks(rotation=45, fontsize=9)

                    color1 = ax.plot_date(blackdf[self.plotdf.columns[0]], blackdf[self.plotdf.columns[1]], fmt='o',
                                          color='green',
                                          picker=5)

                    color2 = ax.plot_date(reddf[self.plotdf.columns[0]], reddf[self.plotdf.columns[1]], fmt='o',
                                          color='lime',
                                          picker=5, )

                    if typ == 'CN':
                        color1 = ax2.plot_date(blackdf[self.plotdf.columns[0]], blackdf[self.plotdf.columns[2]],
                                               fmt='o',
                                               color='blue', picker=5)

                        color2 = ax2.plot_date(reddf[self.plotdf.columns[0]], reddf[self.plotdf.columns[2]], fmt='o',
                                               color='navy',
                                               picker=5, )

                    # ax.xaxis.set_ticks(np.arange(start, end, stepsize))

                else:
                    ax.scatter(self.plotdf[self.plotdf.columns[0]], self.plotdf[self.plotdf.columns[1]],
                               color=farbliste)

            elif (self.plotdf[self.plotdf.columns[0]].dtype == "datetime64[ns]"):
                plt.xticks(rotation=45, fontsize=9)
                # ax.xaxis.set_ticks(np.arange(start, end, stepsize))
                ax.plot_date(self.plotdf[self.plotdf.columns[0]], self.plotdf[self.plotdf.columns[1]], fmt='o',
                             picker=5, )

            else:
                ax.scatter(self.plotdf[self.plotdf.columns[0]], self.plotdf[self.plotdf.columns[1]], color='g')

            if self.plottemp:
                ax2 = ax.twinx()
                tempdata = self.get_temp_values()
                ax2.plot(tempdata[tempdata.columns[0]], tempdata[tempdata.columns[1]], color='r', alpha=.2)
                ax2.set_ylabel('temp °C', fontsize=16)

            '''
            You
            can
            groupby and plot
            them
            separately
            for each color:

            import matplotlib.pyplot as plt

            fig, ax1 = plt.subplots(figsize=(30, 10))
            color = 'tab:red'
            for pcolor, gp in df.groupby('color'):
                ax1.plot_date(gp['time'], gp['distance'], marker='o', color=pcolor)
            '''
            # ax.set_position([0, 0, 0,0])
            self.figure.subplots_adjust(0.1, 0.2, 0.9, 0.9, )  # 0.2,0.3

            datacursor(formatter=self.myformatter2, display='multiple', draggable=True)
            ax.legend(fontsize=12)
            ax.legend().set_visible(False)
            ax.grid(True)
            # (pdfile.index,pdfile.values,150,marker = ">")
            # ax.plot(plotdf.columns[1].value, plotdf.columns[1].value)
            self.figure.suptitle(title, fontsize=16, fontweight='bold')
            self.draw()

        if typ == "CO_SA":
            self.ax = self.figure.add_subplot(111)
            xlabel = self.plotdf.columns[0]
            ylabel = self.plotdf.columns[1]

            self.ax.tick_params(labelright=True)
            Ag3PO4 = self.get_extraframe(self.plotdf, 'Ag3PO4')
            dfridofblnk = self.get_ridoff(self.plotdf, 'Blnk')

            self.x = dfridofblnk.id
            y = dfridofblnk.o18vsmowodavg
            e = dfridofblnk.o18vsmowodstd
            y2 = dfridofblnk.o18vsmowmdavg
            e2 = dfridofblnk.o18vsmowmdstd

            self.ax.cla()
            # self.axes2.cla()
            od = self.ax.errorbar(self.x, y, yerr=e, fmt='o', capsize=3, label='Single-Group-Calibration')

            datacursor(formatter=self.myformatter, display='multiple', draggable=True)

            # ax.hold(True) # deprecated, see above
            self.ax.set_xlabel('id', fontsize=16)
            self.ax.set_ylabel('δ 18O in ‰ ', fontsize=16, color='b')

            md = self.ax.errorbar(self.x, y2, yerr=e2, color='r', alpha=.5, fmt='o', capsize=3,
                                    label='Drift Calibration')
            # self.axes2.set_ylabel('‰ o18 vsmow drift corrected  ', fontsize=16, color='r')
            # self.axes2.grid(False)
            self.ax.plot(Ag3PO4.id, Ag3PO4.o18vsmowodavg, alpha=.2, color='orange', label='Ag3PO4')
            # Right Y-axis labels
            self.ax.legend()

            # dfridofblnk.reset_index(inplace=True,drop=True)
            for i in dfridofblnk.index:
                yt = dfridofblnk.o18vsmowodavg[i] + 2
                xt = dfridofblnk.id[i] - 0.4
                nt = dfridofblnk.name[i]
                self.ax.text(xt, yt, nt, fontsize=12,rotation=45)
            self.figure.suptitle(title, fontsize=16, fontweight='bold')
            self.draw()

        if typ == "CN_SA":

            self.ax = self.figure.add_subplot(211)
            self.ax2 = self.figure.add_subplot(212)
            xlabel = self.plotdf.columns[2]
            ylabel = self.plotdf.columns[4]
            ylabel2 = self.plotdf.columns[3]

            self.figure.suptitle(title, fontsize=16, fontweight='bold')
            # ax.hold(True) # deprecated, see above
            # ax.set_xlabel(xlabel, fontsize=16)
            self.ax.set_ylabel(ylabel, fontsize=16)
            self.ax2.set_ylabel(ylabel2, fontsize=16)

            USGS40 = self.get_extraframe(self.plotdf, 'USGS 40')
            USGS41 = self.get_extraframe(self.plotdf, 'USGS 41')
            USGS41a = self.get_extraframe(self.plotdf, 'USGS 41a')
            dfridofblnk = self.get_ridoff(self.plotdf, 'Blnk')

            self.x = dfridofblnk.id
            y = dfridofblnk.n15airavg
            e = dfridofblnk.n15airstd
            y2 = dfridofblnk.n15aircaliavg
            e2 = dfridofblnk.n15aircalistd

            y3 = dfridofblnk.c13vpdbavg
            e3 = dfridofblnk.c13vpdbstd                       # cali = drift sozusagen. da aber nach dem update von ionos 'drift ,calibrated und single group' möglich sind. als cali bezeichnet
            y4 = dfridofblnk.c13vpdbcaliavg
            e4 = dfridofblnk.c13vpdbcalistd

            odN15 = self.ax.errorbar(self.x, y, yerr=e, fmt='o', capsize=3,alpha=.5, label='Calibrated')

            mdN15 = self.ax.errorbar(self.x, y2, yerr=e2, color='g', alpha=.5, fmt='o', capsize=3,
                                  label='Drift calibrated')
            odC13 = self.ax2.errorbar(self.x, y3, yerr=e3, fmt='o', capsize=3,alpha=.5, label='Calibrated')

            mdC13 = self.ax2.errorbar(self.x, y4, yerr=e4, color='b', alpha=.5, fmt='o', capsize=3,
                                  label='Drift calibrated')

            datacursor(formatter=self.myformatter3, display='multiple', draggable=True)

            self.ax.grid(True)
            self.ax2.grid(True)

            for i in dfridofblnk.index:
                yt = dfridofblnk.n15airavg[i]
                xt = dfridofblnk.id[i]
                nt = dfridofblnk.name[i]
                self.ax.text(xt, yt, nt, fontsize=10 , rotation=65)

            for i in dfridofblnk.index:
                yt = dfridofblnk.c13vpdbavg[i]
                xt = dfridofblnk.id[i]
                nt = dfridofblnk.name[i]
                self.ax2.text(xt, yt, nt, fontsize=10,rotation = 65)

            self.ax.plot(USGS40.id, USGS40.n15airavg, alpha=.2, color='orange', label='USGS 40 calibrated')
            self.ax.plot(USGS41.id, USGS41.n15airavg, alpha=.2, color='orange', label='USGS 41 calibrated')
            self.ax.plot(USGS41a.id, USGS41a.n15airavg, alpha=.2, color='orange', label='USGS 41a calibrated')
            self.ax2.plot(USGS40.id, USGS40.c13vpdbavg, alpha=.2, color='orange', label='USGS 40 calibrated')
            self.ax2.plot(USGS41.id, USGS41.c13vpdbavg, alpha=.2, color='orange', label='USGS 41 calibrated')
            self.ax2.plot(USGS41a.id, USGS41a.c13vpdbavg, alpha=.2, color='orange', label='USGS 41a calibrated')

            self.ax.legend([odN15,mdN15],['Single-Group-Calibration','Drift Calibration'])
            self.ax2.legend([odC13,mdC13],['Single-Group-Calibration','Drift Calibration'] )
            self.draw()

    def compute_initial_figure(self):
        logger.debug('perform function')
        pass

    def get_extraframe(self, dataframe, name):
        logger.debug('perform function')
        booleandf_ = dataframe['name'] == name
        resultdf = dataframe[booleandf_]
        return resultdf

    def get_ridoff(self, dataframe, name):
        logger.debug('perform function')
        booleandf_ = dataframe['name'] != name
        resultdf = dataframe[booleandf_]
        return resultdf

    def update_samplecanvas(self, dataframe):
        logger.debug('perform function')
        self.plotdf = dataframe
        if dataframe.columns[4] == 'n15air':
            self.plotprepare(dataframe, 'Messlauf N15 C13 ', 'CN_SA')

        if dataframe.columns[4] != 'n15air':
            self.plotprepare(dataframe, 'Messlauf O18', 'CO_SA')

    def updatestdzoom(self, name, typ, days = '90'):
        logger.debug('perform function')
        self.typ = typ
        if self.typ == 'CO':
            getdata = "select finishdate,o18vsmowod,id,filename,height,oarea,name from opod where  finishdate between datetime('now', '-" + days + " days') AND datetime('now', 'localtime'); "
        if self.typ == 'CN':
            getdata = "select finishdate,n15air,c13vpdb,filename,id,nheight,cheight,name from cnod where  finishdate between datetime('now','-" + days + " days') AND datetime('now', 'localtime'); "
        self.titlename = name[0]
        self.name = name
        connection = sqlite3.connect(database)
        df = read_sql(getdata, con=connection, coerce_float=True, params=None, parse_dates=None,
                               columns=None, chunksize=None)

        self.get_name_frame(df)

    def changedays(self,days):
        logger.debug('perform function')
        if self.typ == 'CO':
            getdata = "select finishdate,o18vsmowod,id,filename,height,oarea,name from opod where  finishdate between datetime('now', '-" + days + " days') AND datetime('now', 'localtime'); "
        if self.typ == 'CN':
            getdata = "select finishdate,n15air,c13vpdb,filename,id,nheight,cheight,name from cnod where  finishdate between datetime('now','-" + days + " days') AND datetime('now', 'localtime');"

        connection = sqlite3.connect(database)
        df = read_sql(getdata, con=connection, coerce_float=True, params=None, parse_dates=None,
                      columns=None, chunksize=None)

        self.get_name_frame(df)

    def get_name_frame(self,df):
        logger.debug('perform function')
        dfend = df.loc[df['name'].isin(self.name)]
        if dfend.empty:
            for i in self.name:
                dfend = df[df['name'].str.contains(i)]
                if dfend.empty == False:
                    break

        self.plotprepare(dfend, self.titlename, self.typ)

    def getvalues(self, id, dataframe):
        logger.debug('perform function')
        df = dataframe
        print(id)
        dfrow = df.loc[df['id'] == id]
        print(type(id))
        if type(id) == str:
            print("erkannt")
            print(df)
            df.set_index('finishdate')
            dfrow = df.loc[df['finishdate'] == id]

        # dfrow = main.samplestab.resultdf.id[id]

        print('getvalues:', dfrow)
        return dfrow

    def myformatter(self, **kwarg):
        logger.debug('perform function')
        df = main.dialog.resultdf
        valuesdf = self.getvalues(kwarg['x'], df)

        pandas.set_option('display.max_columns', 30)
        print("Type: ", valuesdf)
        print(valuesdf)
        orderdf = valuesdf[
            ['id', 'name', 'o18vsmowodavg', 'o18vsmowmdavg', 'o18vsmowodstd', 'o18vsmowmdstd', 'finishdate', 'area',
             'height', 'ratio2928', 'ratio2928raw', 'ratio3028', 'ratio3028raw', 'c13gas']]
        label = orderdf.T
        label = label.to_string()
        # label = str(orderdf.T)
        label = label.split("\n", 1)[1]

        return label

    def myformatter2(self, **kwarg):
        logger.debug('perform function')
        df = self.plotdf

        val1 = matplotlib.dates.num2date(kwarg['x']).strftime('%Y-%m-%d %H:%M:%S.%f')
        valuesdf = self.getvalues(val1, df)
        pandas.set_option('display.max_columns', 30)

        # orderdf = valuesdf[['id','o18vsmowod','finishdate','filename','height','oarea']]
        label = valuesdf
        label = label.to_string()
        label = str(valuesdf.T)
        label = label.split("\n", 1)[1]

        return label

    def myformatter3(self, **kwarg):
        logger.debug('perform function')
        df = self.plotdf
        valuesdf = self.getvalues(kwarg['x'], df)
        label = valuesdf
        label = label.to_string()
        label = str(valuesdf.T)
        label = label.split("\n", 1)[1]

        return label

if __name__ == "__main__":
    logger.debug('running __main__')

    packaging = True

    if packaging == True:
        # use this when packaging with fbs
        appctxt = ApplicationContext()      
        #app = QApplication(sys.argv)
        main = MainWindow()
        main.show()
        #sys.exit(app.exec_())
        appctxt.app.setStyle("Fusion")
        sys.exit(appctxt.app.exec_())
    else:
        # use this when not packaging
        #appctxt = ApplicationContext()
        app = QApplication(sys.argv)
        main = MainWindow()
        main.show()
        sys.exit(app.exec_())
        #appctxt.app.setStyle("Fusion")
        #sys.exit(appctxt.app.exec_())
