using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Runtime.InteropServices;

namespace Specification
{
    public partial class MainForm : Form
    {
        List<Unit> Units;

        public MainForm()
        {
            InitializeComponent();
        }

        struct Unit
        {
            public string PosNum;
            public string Article;
            public string Designation;
            public int Num;
            public string NumName;
            public string Company;

            public Unit(string posNum, string article, string designation, int num, string numname, string company)
            {
                PosNum = posNum;
                Article = article;
                Designation = designation;
                Num = num;
                NumName = numname;
                Company = company;
            }

            public Unit(string article, string designation, int num, string numname, string company)
            {
                PosNum = "";
                Article = article;
                Designation = designation;
                Num = num;
                NumName = numname;
                Company = company;
            }
        }

        private void FormButton_Click(object sender, EventArgs e)
        {
            Units = new List<Unit>();
            string path = "";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "(*.xlsx); (*.xls)|*.xlsx; *.xls";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                path = ofd.FileName;
            if (path != "")
            {
                if (powerRB.Checked)
                {
                    Units = downloadData(path, true);
                    for (int i = 0; i < Units.Count; i++)
                        for (int j = i + 1; j < Units.Count; j++)
                        {
                            if (Units[i].Article == Units[j].Article)
                            {
                                Unit unit = Units[i];
                                unit.Num += Units[j].Num;
                                unit.PosNum += ", " + Units[j].PosNum;
                                Units.RemoveAt(j);
                                j--;
                                Units[i] = unit;
                            }
                        }
                    posGroup();
                    ofd = new OpenFileDialog();
                    ofd.Filter = "(*.xlsx); (*.xls)|*.xlsx; *.xls";
                    if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        uploadData(ofd.FileName, Units, true);
                }
                else
                    if (controlRB.Checked)
                    {
                        Units = downloadData(path, false);
                        for (int i = 0; i < Units.Count; i++)
                            for (int j = i + 1; j < Units.Count; j++)
                            {
                                if (Units[i].Article == Units[j].Article)
                                {
                                    Unit unit = Units[i];
                                    unit.Num += Units[j].Num;
                                    unit.PosNum += ", " + Units[j].PosNum;
                                    Units.RemoveAt(j);
                                    j--;
                                    Units[i] = unit;
                                }
                            }
                        posGroup();
                        ofd = new OpenFileDialog();
                        ofd.Filter = "(*.xlsx); (*.xls)|*.xlsx; *.xls";
                        if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            uploadData(ofd.FileName, Units, false);
                    }
            }
        }

        private static List<Unit> downloadData(string path, bool power)
        {
            List<Unit> units = new List<Unit>();
            DataSet dataSet = new DataSet("EXCEL");
            string connectionString;
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=NO;'";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
            adapter.Fill(dataSet);
            connection.Close();

            if (!power)
            {
                for (int row = 0; row < dataSet.Tables[0].Rows.Count; row++)
                {
                    if (dataSet.Tables[0].Rows[row][3].ToString() != "")
                    {
                        Unit unit = new Unit();
                        unit.PosNum = dataSet.Tables[0].Rows[row][2].ToString();
                        unit.PosNum = unit.PosNum.Replace("-", "");
                        unit.Article = dataSet.Tables[0].Rows[row][3].ToString();
                        unit.Designation = dataSet.Tables[0].Rows[row][4].ToString();
                        unit.Num = (int)Math.Round(double.Parse(dataSet.Tables[0].Rows[row][5].ToString()));
                        unit.NumName = dataSet.Tables[0].Rows[row][6].ToString();
                        unit.Company = dataSet.Tables[0].Rows[row][7].ToString();
                        units.Add(unit);
                    }
                }
            }
            else
            {
                for (int row = 2; row < dataSet.Tables[0].Rows.Count; row++)
                {
                    if (dataSet.Tables[0].Rows[row][1].ToString() != "")
                    {
                        Unit unit = new Unit();
                        unit.PosNum = dataSet.Tables[0].Rows[row][0].ToString();
                        unit.PosNum = unit.PosNum.Replace("-", "");
                        unit.Article = dataSet.Tables[0].Rows[row][1].ToString();
                        unit.Designation = dataSet.Tables[0].Rows[row][2].ToString();
                        unit.Company = dataSet.Tables[0].Rows[row][3].ToString();
                        unit.Num = (int)Math.Round(double.Parse(dataSet.Tables[0].Rows[row][4].ToString()));
                        units.Add(unit);
                    }
                }
            }
            return units;
        }

        private void uploadData(string fileName, List<Unit> units, bool power)
        {
            string ConnectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO;'");
            OleDbConnection myODCon = new OleDbConnection(ConnectionString);
            myODCon.Open();
            System.Data.DataTable schemaTable = myODCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string name = "СпецификацияNEW";
            int sub = 1;
            bool exist = true;
            while (exist)
            {
                for (int i = 0; i < schemaTable.Rows.Count; i++)
                    if (schemaTable.Rows[i].ItemArray[2].ToString().Contains(name))
                    {
                        name += sub.ToString();
                        sub++;
                        exist = true;
                        break;
                    }
                    else exist = false;
            }

            if (!power)
            {
                OleDbCommand myOleDbCommand = new OleDbCommand();
                myOleDbCommand.Connection = myODCon;
                myOleDbCommand.CommandText = String.Format("CREATE TABLE [{0}](1 string,2 string,3 string,4 string,5 string,6 string)", name);
                myOleDbCommand.ExecuteNonQuery();

                for (int i = 0; i < units.Count; i++)
                {
                    myOleDbCommand.CommandText = String.Format("INSERT INTO [{0}$] values('{1}','{2}','{3}',{4},'{5}','{6}')", name, units[i].PosNum, units[i].Article, units[i].Designation, units[i].Num, units[i].NumName, units[i].Company, myODCon);
                    myOleDbCommand.ExecuteNonQuery();
                }
            }
            else
            {
                OleDbCommand myOleDbCommand = new OleDbCommand();
                myOleDbCommand.Connection = myODCon;
                myOleDbCommand.CommandText = String.Format("CREATE TABLE [{0}](1 string,2 string,3 string,4 string,5 int)", name);
                myOleDbCommand.ExecuteNonQuery();

                for (int i = 0; i < units.Count; i++)
                {
                    myOleDbCommand.CommandText = String.Format("INSERT INTO [{0}$] values('{1}','{2}','{3}','{4}',{5})", name, units[i].PosNum, units[i].Article, units[i].Designation, units[i].Company, units[i].Num, myODCon);
                    myOleDbCommand.ExecuteNonQuery();
                }
            }
            myODCon.Close();
            MessageBox.Show("Файл сохранен");
        }

        private void posGroup()
        {
            for (int i=0; i<Units.Count; i++)
            {
                List<string> str = Units[i].PosNum.Split(',').ToList();
                for (int j = 0; j < str.Count; j++)
                {
                    str[j] = str[j].Trim();
                }
                if (str.Count > 0)
                {
                    str.Sort(new NaturalStringComparer());
                    List<posUnit> posUnits = new List<posUnit>();
                    if (str[0] != "")
                    {
                        str[0] = str[0].Replace(" ", "");
                        string numS;
                        string s = getPos(str[0], out numS);
                        int num = int.Parse(numS);
                        posUnits.Add(new posUnit(str[0], "", num));
                        for (int j = 1; j < str.Count; j++)
                        {
                            if (str[j] != "")
                            {
                                str[j] = str[j].Replace(" ", "");
                                if (str[j].Length > 2)
                                {
                                    string numNew;
                                    string posNew = getPos(str[j], out numNew);
                                    string numPrev;
                                    string posPrev = getPos(posUnits[posUnits.Count - 1].left, out numPrev);

                                    int numN = int.Parse(numNew);
                                    int numP = int.Parse(numPrev);
                                    if (posNew == posPrev && numN == posUnits[posUnits.Count - 1].lastNum + 1)
                                    {
                                        posUnit pUnit = posUnits[posUnits.Count - 1];
                                        pUnit.right = str[j];
                                        pUnit.lastNum = numN;
                                        posUnits[posUnits.Count - 1] = pUnit;
                                    }
                                    else posUnits.Add(new posUnit(str[j], "", numN));
                                }
                                else
                                {
                                    if (str[j] != posUnits[posUnits.Count - 1].left)
                                        posUnits.Add(new posUnit(str[j], "", 0));
                                }
                            }
                        }
                        Unit unit = Units[i];
                        if (posUnits[0].right != string.Empty)
                            unit.PosNum = posUnits[0].left + "-" + posUnits[0].right;
                        else
                            unit.PosNum = posUnits[0].left;
                        for (int j = 1; j < posUnits.Count; j++)
                        {
                            if (posUnits[j].right != string.Empty)
                                unit.PosNum += ", " + posUnits[j].left + "-" + posUnits[j].right;
                            else
                                unit.PosNum += ", " + posUnits[j].left;
                        }
                        Units[i] = unit;
                    }
                }
            }
        }

        struct posUnit
        {
            public string left;
            public string right;
            public int lastNum;
            public posUnit(string l, string r, int num)
            {
                left = l;
                right = r;
                lastNum = num;
            }
        }

        private string getPos(string str, out string number)
        {
            string pos = str[0].ToString();
            int index = 0;
            for (int i = 1; i < str.Length; i++ )
                if (Char.IsLetter(str[i]))
                {
                    pos += str[i];
                    index = i;
                }

            if (index == str.Length-1)
                number = "0";
            else
                number = str.Substring(index+1);
            return pos;
        }
    }

    class NaturalStringComparer : IComparer<String>
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        static extern int StrCmpLogicalW(string s1, string s2);
        public int Compare(string x, string y)
        {
            return StrCmpLogicalW(x, y);
        }
    }
}
