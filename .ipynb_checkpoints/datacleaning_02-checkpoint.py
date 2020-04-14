{
 "cells": [
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Import libraries"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 1,
   "metadata": {},
   "outputs": [],
   "source": [
    "import pandas as pd\n",
    "import numpy as np\n",
    "import xlwings as xw"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Define global variables"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 3,
   "metadata": {},
   "outputs": [],
   "source": [
    "sheetNames = ['p', 'ey', 'fcfy', 'eps', 'ptbv', 'dy', 'mv']\n",
    "fileName = \"Pull 1.xlsx\""
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Read data from Excel"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 33,
   "metadata": {},
   "outputs": [],
   "source": [
    "# Requires file to be open\n",
    "# p = pd.DataFrame(xw.books(fileName).sheets(\"p\").range(\"A1:BVI287\").value)\n",
    "# ey = pd.DataFrame(xw.books(fileName).sheets(\"ey\").range(\"A1:BVI287\").value)\n",
    "# fcfy = pd.DataFrame(xw.books(fileName).sheets(\"fcfy\").range(\"A1:BVI287\").value)\n",
    "# eps = pd.DataFrame(xw.books(fileName).sheets(\"eps\").range(\"A1:BVI287\").value)\n",
    "# ptbv = pd.DataFrame(xw.books(fileName).sheets(\"ptbv\").range(\"A1:BVI287\").value)\n",
    "# dy = pd.DataFrame(xw.books(fileName).sheets(\"dy\").range(\"A1:BVI287\").value)\n",
    "# mv = pd.DataFrame(xw.books(fileName).sheets(\"mv\").range(\"A1:BVI287\").value)\n",
    "\n",
    "# fieldList = [p, ey, fcfy, eps, ptbv, dy, mv]\n",
    "\n",
    "# # Clean up the dataframes to have correct cols and rows\n",
    "# for field in fieldList:\n",
    "#     field.columns = field.iloc[0]\n",
    "#     field.index = field[field.columns[0]]\n",
    "#     del field[field.columns[0]]\n",
    "#     field.drop(field.index[0], inplace=True)\n",
    "#     field.replace(\"NA\", np.nan, inplace=True)\n",
    "#     field = field.apply(pd.to_numeric,errors='coerce')b  "
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 2,
   "metadata": {},
   "outputs": [],
   "source": [
    "p = pd.read_excel('..\\data\\Pull 1.xlsx', header=0, index_col=0, sheet_name='p')\n",
    "ey = pd.read_excel('..\\data\\Pull 1.xlsx', header=0, index_col=0, sheet_name='ey')\n",
    "fcfy = pd.read_excel('..\\data\\Pull 1.xlsx', header=0, index_col=0, sheet_name='fcfy')\n",
    "eps = pd.read_excel('..\\data\\Pull 1.xlsx', header=0, index_col=0, sheet_name='eps')\n",
    "ptbv = pd.read_excel('..\\data\\Pull 1.xlsx', header=0, index_col=0, sheet_name='ptbv')\n",
    "dy = pd.read_excel('..\\data\\Pull 1.xlsx', header=0, index_col=0, sheet_name='dy')\n",
    "mv = pd.read_excel('..\\data\\Pull 1.xlsx', header=0, index_col=0, sheet_name='mv')"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": [
    "\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Cleaning price series 'p'"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 5,
   "metadata": {},
   "outputs": [],
   "source": [
    "r = p/p.shift(1)-1"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": []
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Calculate Moving Averages"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 7,
   "metadata": {},
   "outputs": [],
   "source": [
    "p_ma2 = p.rolling(window=2).mean()\n",
    "p_ma3 = p.rolling(window=3).mean()\n",
    "p_ma6 = p.rolling(window=6).mean()\n",
    "p_ma12 = p.rolling(window=12).mean()\n",
    "p_ma18 = p.rolling(window=18).mean()"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.7.4"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 2
}
