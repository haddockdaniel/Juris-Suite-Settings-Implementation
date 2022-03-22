using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            var count = groupBox1.Controls.OfType<CheckBox>().Count(x => x.Checked);
            int current = 1;
            if (checkBoxTeams.Checked)
            {
                setTeams();
                UpdateStatus("Updating.....", current, count);
                current++;
            }
            if (checkBoxGroups.Checked)
            {
                setGroups();
                UpdateStatus("Updating.....", current, count);
                current++;
            }
            if (checkBoxPermissions.Checked)
            {
                setPermissions();
                UpdateStatus("Updating.....", current, count);
                current++;
            }
            if (checkBoxFirmSettings.Checked)
            {
                setFsettings();
                UpdateStatus("Updating.....", current, count);
                current++;
            }
            if (checkBoxUserSettings.Checked)
            {
                setUsettings();
                UpdateStatus("Updating.....", current, count);
                current++;
            }

            UpdateStatus("Update Complete", count, count);


            MessageBox.Show("The process is complete.", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
        }

        private void setTeams()
        {
            this._jurisUtility.ExecuteNonQuery(0, " Insert into organizationalunit(name, code, type) select 'EVERYONE', 'EVERYONE', 5 where(select id from organizationalunit where type = 5 and name = 'EVERYONE') is null");
            this._jurisUtility.ExecuteNonQuery(0, "Insert into organizationalunitteam(organizationalunitid) select id from organizationalunit where type = 5 and id not in (select organizationalunitid from organizationalunitteam)");
            this._jurisUtility.ExecuteNonQuery(0, "Insert into organizationalunitteammember(organizationalunitid, parentid, status) select organizationalunit.id,(select id from organizationalunit where type = 5 and name = 'EVERYONE'),3 " +
                                                  " from organizationalunit where code = 'SMGR' and type = 2 and id not in (select OrganizationalUnitID from organizationalunitteammember inner join organizationalunit on " +
                                                   " parentid = organizationalunit.id where type = 5 and code = 'EVERYONE')");
        }

        private void setGroups()
        {
            this._jurisUtility.ExecuteNonQuery(0, " Insert into organizationalunit(name, code, type) select 'ADMIN', 'ADMIN', 4 where(select id from organizationalunit where type = 4 and name = 'ADMIN') is null");
            this._jurisUtility.ExecuteNonQuery(0, " Insert into organizationalunitgroup(organizationalunitid) select id from organizationalunit where type = 4 and id not in (select organizationalunitid from organizationalunitgroup)");
            this._jurisUtility.ExecuteNonQuery(0, " Insert into organizationalunitmember(organizationalunitid, parentid, sequencenumber) select organizationalunit.id,(select id from organizationalunit where type = 4 and name = 'ADMIN'),0 " +
                                                 " from organizationalunit where code = 'SMGR' and type = 2 and id not in (select OrganizationalUnitID from organizationalunitmember inner join organizationalunit on " +
                                                 " parentid = organizationalunit.id  where type = 4 and name = 'ADMIN')");
        }

        private void setPermissions()
        {
            this._jurisUtility.ExecuteNonQuery(0, "delete from ORGANIZATIONALUNITPERMISSION where organizationalunitid in (select id from organizationalunit where code = 'FIRM' and type = 2) or organizationalunitid in " +
                                                  "(select id from organizationalunit where code = 'ADMIN' and type = 4)");
            this._jurisUtility.ExecuteNonQuery(0, "Insert into ORGANIZATIONALUNITPERMISSION(organizationalunitid, permissionid, permissiontype) select(select id from organizationalunit where code = 'ADMIN' and type = 4), id, 1 " +
                                                  "from OrganizationalUnitPERMISSIONDefinition");
            this._jurisUtility.ExecuteNonQuery(0, " Insert into ORGANIZATIONALUNITPERMISSION(organizationalunitid, permissionid, permissiontype) select 2, id, 1 from OrganizationalUnitPERMISSIONDefinition  where id in (31, 41, 278,22,87,26,21,119,127,125,123,126,124,133, 28)");
            this._jurisUtility.ExecuteNonQuery(0, " Insert into ORGANIZATIONALUNITPERMISSION(organizationalunitid, permissionid, permissiontype) select 2, id, 0 from OrganizationalUnitPERMISSIONDefinition  where id not in (31, 41, 278,22,87,26,21,119,127,125,123,126,124,133, 28)");
        }

        private void setFsettings()
        {
            using (Stream manifestResourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("JurisUtilityBase.FirmSettings.txt"))
            {
                using (StreamReader streamReader = new StreamReader(manifestResourceStream))
                {
                    string[] strArray1 = streamReader.ReadLine().Split('\t');
                    DataTable dataTable = new DataTable();
                    foreach (string columnName in strArray1)
                        dataTable.Columns.Add(columnName);
                    string str1;
                    while ((str1 = streamReader.ReadLine()) != null)
                    {
                        DataRow row = dataTable.NewRow();
                        string[] strArray2 = str1.Split('\t');
                        for (int columnIndex = 0; columnIndex < 3; ++columnIndex)
                            row[columnIndex] = (object)strArray2[columnIndex];
                        dataTable.Rows.Add(row);
                    }
                    //this.dgvFirm.DataSource = (object)dataTable;
                    this._jurisUtility.ExecuteNonQuery(0, "Update SysParam set sptxtvalue='N,Y,Y,N,Y,N,N,N,N,N,N,Y,Y,Y,Y,Y,N,N,N,N,N,N,Y,Y,N,N,N,N,N,N,N,N,N,N,N,N,N,N,N,N,N,Y,Y' where spname='CfgConflict'");
                    foreach (DataRow row in (InternalDataCollectionBase)dataTable.Rows)
                    {
                        string str2 = row[0].ToString();
                        string str3 = row[1].ToString();
                        string str4 = row[2].ToString();
                        this._jurisUtility.ExecuteNonQuery(0, "Delete from organizationalunitsetting where organizationalunitid=" + str2 + " and settingid = " + str3);
                        this._jurisUtility.ExecuteNonQueryCommand(0, "Insert into OrganizationalUnitSetting(organizationalunitid, settingid, value)   Values(" + str2 + "," + str3 + "," + str4 + ")");
                    }
                }
            }
        }

        private void setUsettings()
        {
            using (Stream manifestResourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("JurisUtilityBase.UserSettings.txt"))
            {
                using (StreamReader streamReader = new StreamReader(manifestResourceStream))
                {
                    string[] strArray1 = streamReader.ReadLine().Split('\t');
                    DataTable dataTable = new DataTable();
                    foreach (string columnName in strArray1)
                        dataTable.Columns.Add(columnName);
                    string str1;
                    while ((str1 = streamReader.ReadLine()) != null)
                    {
                        DataRow row = dataTable.NewRow();
                        string[] strArray2 = str1.Split('\t');
                        for (int columnIndex = 0; columnIndex < 2; ++columnIndex)
                            row[columnIndex] = (object)strArray2[columnIndex];
                        dataTable.Rows.Add(row);
                    }
                    streamReader.Close();
                    //this.dgvUser.DataSource = (object)dataTable;
                    foreach (DataRow row in (InternalDataCollectionBase)dataTable.Rows)
                    {
                        string str2 = row[0].ToString();
                        string str3 = row[1].ToString();
                        this._jurisUtility.ExecuteNonQuery(0, "Delete from organizationalunitsetting where organizationalunitid in (select id from organizationalunit where type=2) and  settingid = " + str2);
                        this._jurisUtility.ExecuteNonQueryCommand(0, "Insert into OrganizationalUnitSetting(organizationalunitid, settingid, value)   select id," + str2 + "," + str3 + " from organizationalunit where type=2 ");
                    }
                }
            }
        }




        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }




    }
}
