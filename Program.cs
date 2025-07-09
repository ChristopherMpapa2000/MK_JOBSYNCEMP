using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Threading;
using System.IO;

namespace SyncEmp
{
    internal class Program
    {

        private static string filename
        {
            get
            {
                var filename = ConfigurationManager.AppSettings["filename"];
                if (!string.IsNullOrEmpty(filename))
                {
                    return filename;
                }
                return string.Empty;
            }
        }
        private static string Pathmove
        {
            get
            {
                var Pathmove = ConfigurationManager.AppSettings["Pathmove"];
                if (!string.IsNullOrEmpty(Pathmove))
                {
                    return Pathmove;
                }
                return string.Empty;
            }
        }
        private static string Connectionstring
        {
            get
            {
                var Connectionstring = ConfigurationManager.AppSettings["Connectionstring"];
                if (!string.IsNullOrEmpty(Connectionstring))
                {
                    return Connectionstring;
                }
                return string.Empty;
            }
        }
        static void Main(string[] args)
        {
            bool check = CheckFile();
            if (check)
            {
                GetData();
                MoveFile();
                EmployeeIsActive();
                Thread.Sleep(50000);
            }
            else
            {
                WriteLogFile("!!! <-- Not Have File Update --> !!! ");
            }
        }
        static void GetData()
        {
            try
            {
                string oledbConnectString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filename}; Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'";

                using (OleDbConnection connection = new OleDbConnection(oledbConnectString))
                {
                    connection.Open();
                    OleDbCommand command = new OleDbCommand("select * from [EmployeeISO]", connection);
                    OleDbDataAdapter oleda = new OleDbDataAdapter();

                    oleda.SelectCommand = command;
                    DataSet ds = new DataSet();

                    oleda.Fill(ds, "Employees");
                    List<string[]> lstData = new List<string[]>();
                    List<string[]> lstAllDepartment = new List<string[]>();
                    int update = 0;
                    int insert = 0;
                    for (int i = 1; i < ds.Tables[0].DefaultView.Count; i++)
                    {
                        string[] Data = new string[34];
                        try
                        {

                            Data[0] = ds.Tables[0].DefaultView[i][0].ToString().Trim();
                            Data[1] = ds.Tables[0].DefaultView[i][1].ToString().Trim();
                            Data[2] = ds.Tables[0].DefaultView[i][2].ToString().Trim();
                            Data[3] = ds.Tables[0].DefaultView[i][3].ToString().Trim();
                            Data[4] = ds.Tables[0].DefaultView[i][4].ToString().Trim();
                            Data[5] = ds.Tables[0].DefaultView[i][5].ToString().Trim();
                            Data[6] = ds.Tables[0].DefaultView[i][6].ToString().Trim();
                            Data[7] = ds.Tables[0].DefaultView[i][7].ToString().Trim();
                            Data[8] = ds.Tables[0].DefaultView[i][8].ToString().Trim();
                            Data[9] = ds.Tables[0].DefaultView[i][9].ToString().Trim();
                            Data[10] = ds.Tables[0].DefaultView[i][10].ToString().Trim();
                            Data[11] = ds.Tables[0].DefaultView[i][11].ToString().Trim();
                            Data[12] = ds.Tables[0].DefaultView[i][12].ToString().Trim();
                            Data[13] = ds.Tables[0].DefaultView[i][13].ToString().Trim();
                            Data[14] = ds.Tables[0].DefaultView[i][14].ToString().Trim();
                            Data[15] = ds.Tables[0].DefaultView[i][15].ToString().Trim();
                            Data[16] = ds.Tables[0].DefaultView[i][16].ToString().Trim();
                            Data[17] = ds.Tables[0].DefaultView[i][17].ToString().Trim();

                            Data[18] = ds.Tables[0].DefaultView[i][18].ToString().Trim();
                            Data[19] = ds.Tables[0].DefaultView[i][19].ToString().Trim();
                            if (!string.IsNullOrEmpty(Data[18].ToString()))
                            {
                                lstAllDepartment.Add(new string[] { Data[18].ToString(), Data[19].ToString() });
                            }
                            Data[20] = ds.Tables[0].DefaultView[i][20].ToString().Trim();
                            Data[21] = ds.Tables[0].DefaultView[i][21].ToString().Trim();
                            if (!string.IsNullOrEmpty(Data[20].ToString()))
                            {
                                lstAllDepartment.Add(new string[] { Data[20].ToString(), Data[21].ToString() });
                            }
                            Data[22] = ds.Tables[0].DefaultView[i][22].ToString().Trim();
                            Data[23] = ds.Tables[0].DefaultView[i][23].ToString().Trim();
                            if (!string.IsNullOrEmpty(Data[22].ToString()))
                            {
                                lstAllDepartment.Add(new string[] { Data[22].ToString(), Data[23].ToString() });
                            }
                            Data[24] = ds.Tables[0].DefaultView[i][24].ToString().Trim();
                            Data[25] = ds.Tables[0].DefaultView[i][25].ToString().Trim();
                            if (!string.IsNullOrEmpty(Data[24].ToString()))
                            {
                                lstAllDepartment.Add(new string[] { Data[24].ToString(), Data[25].ToString() });
                            }
                            Data[26] = ds.Tables[0].DefaultView[i][26].ToString().Trim();
                            Data[27] = ds.Tables[0].DefaultView[i][27].ToString().Trim();
                            if (!string.IsNullOrEmpty(Data[26].ToString()))
                            {
                                lstAllDepartment.Add(new string[] { Data[26].ToString(), Data[27].ToString() });
                            }
                            Data[28] = ds.Tables[0].DefaultView[i][28].ToString().Trim();
                            Data[29] = ds.Tables[0].DefaultView[i][29].ToString().Trim();
                            if (!string.IsNullOrEmpty(Data[28].ToString()))
                            {
                                lstAllDepartment.Add(new string[] { Data[28].ToString(), Data[29].ToString() });
                            }
                            Data[30] = ds.Tables[0].DefaultView[i][30].ToString().Trim();
                            Data[31] = ds.Tables[0].DefaultView[i][31].ToString().Trim();
                            if (!string.IsNullOrEmpty(Data[30].ToString()))
                            {
                                lstAllDepartment.Add(new string[] { Data[30].ToString(), Data[31].ToString() });
                            }

                            Data[32] = ds.Tables[0].DefaultView[i][32].ToString().Trim();
                            Data[33] = ds.Tables[0].DefaultView[i][33].ToString().Trim();
                            lstData.Add(Data);
                        }
                        catch (Exception ex)
                        {

                            Console.WriteLine("Error : " + ex.Message.ToString());
                            Console.ReadLine();
                        }
                    }
                    lstAllDepartment = lstAllDepartment.GroupBy(item => string.Join(",", item)).Select(group => group.First()).ToList();

                    DataClassesWolfDataContext db = new DataClassesWolfDataContext(Connectionstring);

                    if (db.Connection.State == ConnectionState.Open)
                    {
                        db.Connection.Close();
                        db.Connection.Open();
                    }
                    else
                    {
                        db.Connection.Open();

                        List<MSTDepartment> lstDepartment = db.MSTDepartments.ToList();
                        List<string[]> lstitem = new List<string[]>();
                        List<string[]> lstitem_update = new List<string[]>();
                        for (int i = 1; i < lstAllDepartment.Count(); i++)
                        {
                            string depCode = (lstAllDepartment[i][0].ToString());
                            if (!lstDepartment.Any(s => s.DepartmentCode == depCode && s.IsActive == true))
                            {
                                lstitem.Add(lstAllDepartment[i]);
                            }
                            else if (lstDepartment.Any(s => s.DepartmentCode == depCode && s.IsActive == true))
                            {
                                lstitem_update.Add(lstAllDepartment[i]);
                            }
                        }
                        #region Insert Department
                        WriteLogFile($"<!===================== Start Insert NewDepartment {lstitem.Count()} ===================== !>");
                        foreach (var insertdep in lstitem)
                        {
                            //เอาข้อมูลลงแบบไม่มี ParentId เดะเอาเข้าทีหลัง
                            MSTDepartment dep = new MSTDepartment();
                            dep.DepartmentCode = insertdep[0].ToString();
                            dep.NameEn = insertdep[1].ToString();
                            dep.NameTh = insertdep[1].ToString();
                            dep.ModifiedDate = DateTime.Now;
                            dep.ModifiedBy = "AdminC";
                            dep.CreatedDate = DateTime.Now;
                            dep.CreatedBy = "AdminC";
                            dep.IsActive = true;
                            dep.AccountId = 1;
                            db.MSTDepartments.InsertOnSubmit(dep);
                            db.SubmitChanges();
                            WriteLogFile($"DepartmentCode = {dep.DepartmentCode}|Name = {dep.NameEn}");
                            WriteLogFile("-------------------------------------------------------------------------------------------");
                        }
                        WriteLogFile("<!===================== End Insert NewDepartment ===================== !>");
                        #endregion
                        #region Update Department
                        WriteLogFile($"<!===================== Start Update Department {lstitem.Count()} ===================== !>");
                        foreach (var updatedep in lstitem_update)
                        {
                            var dep = db.MSTDepartments.Where(x => x.DepartmentCode == updatedep[0] && x.IsActive == true).FirstOrDefault();
                            dep.NameEn = updatedep[1].ToString();
                            dep.NameTh = updatedep[1].ToString();
                            dep.ModifiedDate = DateTime.Now;
                            dep.ModifiedBy = "AdminC";
                            db.SubmitChanges();
                            WriteLogFile($"DepartmentCode = {dep.DepartmentCode}|Name = {dep.NameEn}");
                            WriteLogFile("-------------------------------------------------------------------------------------------");
                        }
                        WriteLogFile("<!===================== End Update Department ===================== !>");
                        #endregion
                        #region Insert ParentId
                        WriteLogFile("<!===================== Start Insert ParentId ===================== !>");
                        List<MSTDepartment> lstDepartment_update = db.MSTDepartments.ToList();
                        foreach (var item in lstData)
                        {
                            string EmployeeCode = item[0].Trim();
                            if (EmployeeCode != "0000000")
                            {
                                List<string[]> lstbu = new List<string[]>();
                                Dictionary<string, string> parentMap = new Dictionary<string, string>();
                                int i = 30;
                                do
                                {
                                    string current_code = item[i].Trim();
                                    string current_name = item[i + 1].Trim();
                                    if (!string.IsNullOrEmpty(current_code) && !string.IsNullOrEmpty(current_name))
                                    {
                                        lstbu.Add(new string[] { current_code, current_name });

                                        if (i > 18)
                                        {
                                            int j = i - 2;
                                            string parent_code = "";
                                            while (j >= 18)
                                            {
                                                parent_code = item[j].Trim();
                                                if (!string.IsNullOrEmpty(parent_code))
                                                {
                                                    break;
                                                }
                                                j -= 2;
                                            }
                                            if (!string.IsNullOrEmpty(parent_code))
                                            {
                                                parentMap[current_code] = parent_code;
                                            }
                                        }
                                    }

                                    i -= 2;
                                } while (i >= 18);
                                if (lstbu.Count() > 0)
                                {
                                    foreach (var entry in parentMap)
                                    {
                                        string DepartmentCode = entry.Key;
                                        string Paren = entry.Value;
                                        if (lstbu.Count() > 0)
                                        {
                                            if (!string.IsNullOrEmpty(Paren))
                                            {
                                                var ParentId = lstDepartment_update.Where(x => x.DepartmentCode == Paren && x.IsActive == true).FirstOrDefault();
                                                List<MSTDepartment> dep = lstDepartment_update.Where(x => x.DepartmentCode == DepartmentCode && x.IsActive == true).ToList();
                                                foreach (var itemp in dep)
                                                {
                                                    itemp.ParentId = ParentId.DepartmentId;
                                                    itemp.ModifiedDate = DateTime.Now;
                                                    itemp.ModifiedBy = "AdminC";
                                                    db.SubmitChanges();
                                                    Console.WriteLine($"DepartmentCode = {DepartmentCode}|ParentId = {ParentId.DepartmentId}|ParentCode = {ParentId.DepartmentCode}|ParentName = {ParentId.NameEn}");
                                                    WriteLogFile($"DepartmentCode = {DepartmentCode}|ParentId = {ParentId.DepartmentId}|ParentCode = {ParentId.DepartmentCode}|ParentName = {ParentId.NameEn}");
                                                    WriteLogFile("-------------------------------------------------------------------------------------------");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        WriteLogFile("<!===================== End Insert ParentId ===================== !>");
                        #endregion

                        List<string> Re_ReportToEmpCode = new List<string>();
                        foreach (var item in lstData)
                        {
                            string EmployeeCode = item[0].Trim();
                            EmployeeCode = EmployeeCode.PadLeft(7, '0');
                            if (EmployeeCode != "0000000")
                            {
                                List<MSTEmployee> lstEmp = db.MSTEmployees.ToList();
                                MSTEmployee UpdateEmp = lstEmp.Where(a => a.EmployeeCode.Trim().PadLeft(7, '0') == EmployeeCode).FirstOrDefault();
                                if (UpdateEmp != null)
                                {
                                    update++;
                                    WriteLogFile("<!===================== Start UpdateEmp =====================!>");
                                    UpdateEmp.EmployeeCode = EmployeeCode;
                                    UpdateEmp.Username = item[1];
                                    UpdateEmp.NameTh = item[2];
                                    UpdateEmp.NameEn = item[5] + " " + item[6];
                                    UpdateEmp.Email = item[11];
                                    UpdateEmp.IsActive = true;
                                    UpdateEmp.AccountId = 1;
                                    UpdateEmp.ModifiedDate = DateTime.Now;
                                    UpdateEmp.ModifiedBy = "AdminC";
                                    List<MSTPosition> lstPos = db.MSTPositions.ToList();
                                    var CurrentPos = lstPos.FindAll(a => a.NameTh.Replace(" ", "") == item[8].Replace(" ", "") || a.NameEn.Replace(" ", "") == item[8].Replace(" ", "")).ToList();
                                    if (CurrentPos.Count > 0)
                                    {
                                        //InsertPosition
                                        UpdateEmp.PositionId = CurrentPos.First().PositionId;
                                    }
                                    else
                                    {
                                        //InsertNewPosition
                                        int Positionid = InsertNewPosition(item[8], item[9]);
                                        if (Positionid != 0)
                                        {
                                            UpdateEmp.PositionId = Positionid;
                                        }
                                    }
                                    var CurrentDep = lstDepartment_update.Where(a => a.DepartmentCode.Replace(" ", "") == item[30].Replace(" ", "") && a.IsActive == true).FirstOrDefault();
                                    if (CurrentDep != null)
                                    {
                                        //InsertDepartment
                                        UpdateEmp.DepartmentId = CurrentDep.DepartmentId;

                                    }
                                    var ReportTo = lstEmp.Where(a => a.EmployeeCode == item[12].Trim()).ToList();
                                    if (ReportTo.Count > 0)
                                    {
                                        UpdateEmp.ReportToEmpCode = ReportTo.First().EmployeeId.ToString();
                                    }
                                    else
                                    {
                                        ReportTo = lstEmp.Where(a => a.EmployeeCode == item[13].Trim()).ToList();
                                        if (ReportTo.Count > 0)
                                        {
                                            UpdateEmp.ReportToEmpCode = ReportTo.First().EmployeeId.ToString();
                                        }
                                        else
                                        {
                                            if (item[12] != "0000000")
                                            {
                                                Re_ReportToEmpCode.Add($"{EmployeeCode},{item[12]}");
                                            }
                                            else
                                            {
                                                Console.WriteLine($"Not have ReportToEmpCode|EmpID = {EmployeeCode}|ReportTo = {item[12]}");
                                                WriteLogFile($"Not have ReportToEmpCode|EmpID = {EmployeeCode}|ReportTo = {item[12]}");
                                            }
                                        }
                                    }
                                    db.SubmitChanges();
                                    WriteLogFile("EmployeeId : " + UpdateEmp.EmployeeId);
                                    WriteLogFile("PositionId : " + UpdateEmp.PositionId);
                                    WriteLogFile("DepartmentId : " + UpdateEmp.DepartmentId);
                                    WriteLogFile("ReportToEmpCode : " + UpdateEmp.ReportToEmpCode);
                                    WriteLogFile("<!====================== End UpdateEmp ======================!>");
                                }
                                else
                                {
                                    WriteLogFile("<!===================== Start Insert NewEmployee =====================!>");
                                    MSTEmployee NewEmployee = new MSTEmployee();
                                    insert++;
                                    NewEmployee.EmployeeCode = EmployeeCode;
                                    NewEmployee.Username = item[1];
                                    NewEmployee.NameTh = item[2];
                                    NewEmployee.NameEn = item[5] + " " + item[6];
                                    NewEmployee.Email = item[11];
                                    NewEmployee.IsActive = true;
                                    NewEmployee.AccountId = 1;
                                    NewEmployee.ModifiedDate = DateTime.Now;
                                    NewEmployee.ModifiedBy = "AdminC";
                                    NewEmployee.CreatedDate = DateTime.Now;
                                    NewEmployee.CreatedBy = "AdminC";
                                    List<MSTPosition> lstPos = db.MSTPositions.ToList();
                                    var CurrentPos = lstPos.FindAll(a => a.NameTh.Replace(" ", "") == item[8].Replace(" ", "") || a.NameEn.Replace(" ", "") == item[8].Replace(" ", "")).ToList();
                                    if (CurrentPos.Count > 0)
                                    {
                                        //InsertPosition
                                        NewEmployee.PositionId = CurrentPos.First().PositionId;
                                    }
                                    else
                                    {
                                        //InsertNewPosition
                                        int Positionid = InsertNewPosition(item[8], item[9]);
                                        if (Positionid != 0)
                                        {
                                            NewEmployee.PositionId = Positionid;
                                        }
                                    }
                                    var CurrentDep = lstDepartment_update.Where(a => a.DepartmentCode.Replace(" ", "") == item[30].Replace(" ", "") && a.IsActive == true).FirstOrDefault();
                                    if (CurrentDep != null)
                                    {
                                        //InsertDepartment
                                        NewEmployee.DepartmentId = CurrentDep.DepartmentId;
                                    }
                                    var ReportTo = lstEmp.Where(a => a.EmployeeCode == item[12].Trim()).ToList();
                                    if (ReportTo.Count > 0)
                                    {
                                        NewEmployee.ReportToEmpCode = ReportTo.First().EmployeeId.ToString();
                                    }
                                    else
                                    {
                                        ReportTo = lstEmp.Where(a => a.EmployeeCode == item[13].Trim()).ToList();
                                        if (ReportTo.Count > 0)
                                        {
                                            NewEmployee.ReportToEmpCode = ReportTo.First().EmployeeId.ToString();
                                        }
                                        else
                                        {
                                            if (item[12] != "0000000")
                                            {
                                                Re_ReportToEmpCode.Add($"{EmployeeCode},{item[12]}");
                                            }
                                            else
                                            {
                                                Console.WriteLine($"Not have ReportToEmpCode|EmpID = {EmployeeCode}|ReportTo = {item[12]}");
                                                WriteLogFile($"Not have ReportToEmpCode|EmpID = {EmployeeCode}|ReportTo = {item[12]}");
                                            }
                                        }
                                    }
                                    db.MSTEmployees.InsertOnSubmit(NewEmployee);
                                    db.SubmitChanges();
                                    WriteLogFile("EmployeeId : " + NewEmployee.EmployeeId);
                                    WriteLogFile("PositionId : " + NewEmployee.PositionId);
                                    WriteLogFile("DepartmentId : " + NewEmployee.DepartmentId);
                                    WriteLogFile("ReportToEmpCode : " + NewEmployee.ReportToEmpCode);
                                    WriteLogFile("<!====================== End Insert NewEmployee ======================!>");
                                }
                            }
                            else
                            {
                                //กรณีเป็นสาขา
                                List<MSTEmployee> lstEmp = db.MSTEmployees.ToList();
                                MSTEmployee UpdateEmp = lstEmp.Where(a => a.EmployeeCode.Trim() == item[1].Trim()).FirstOrDefault();
                                if (UpdateEmp != null)
                                {
                                    update++;
                                    WriteLogFile("<!===================== Start UpdateEmp Branch =====================!>");
                                    UpdateEmp.EmployeeCode = item[1].Trim();
                                    UpdateEmp.Username = item[1].Trim();
                                    UpdateEmp.NameTh = !string.IsNullOrEmpty(item[4]) ? $"{item[4]} {item[3]}" : item[3];
                                    UpdateEmp.NameEn = item[1].Trim() + " " + UpdateEmp.NameTh;
                                    UpdateEmp.Email = item[11];
                                    UpdateEmp.IsActive = true;
                                    UpdateEmp.AccountId = 1;
                                    UpdateEmp.ModifiedDate = DateTime.Now;
                                    UpdateEmp.ModifiedBy = "AdminC";
                                    if (!string.IsNullOrEmpty(item[30]))
                                    {
                                        var CurrentDep = lstDepartment_update.FindAll(a => a.DepartmentCode.Replace(" ", "") == item[30].Replace(" ", "")).ToList();
                                        if (CurrentDep.Count > 0)
                                        {
                                            //InsertDepartment
                                            UpdateEmp.DepartmentId = CurrentDep.First().DepartmentId;
                                            if (!string.IsNullOrEmpty(item[32]))
                                            {
                                                var ParentId = lstDepartment_update.FindAll(a => a.DepartmentCode.Replace(" ", "") == item[32].Replace(" ", "")).FirstOrDefault();
                                                if (ParentId != null)
                                                {
                                                    UpdateDepartment(CurrentDep.FirstOrDefault(), db, ParentId);
                                                    MSTDepartment department = lstDepartment_update.Where(x => x.DepartmentId == ParentId.DepartmentId).FirstOrDefault();
                                                    department.ModifiedDate = DateTime.Now;
                                                    department.IsActive = true;
                                                }
                                                else
                                                {
                                                    WriteLogFile("!! Not have ParentId: " + item[32]);
                                                }
                                            }
                                        }
                                    }
                                    UpdateEmp.PositionId = 705;
                                    db.SubmitChanges();
                                    WriteLogFile("EmployeeId : " + UpdateEmp.EmployeeId);
                                    WriteLogFile("PositionId : " + UpdateEmp.PositionId);
                                    WriteLogFile("DepartmentId : " + UpdateEmp.DepartmentId);
                                    WriteLogFile("ReportToEmpCode : " + UpdateEmp.ReportToEmpCode);
                                    WriteLogFile("<!====================== End UpdateEmp Branch ======================!>");
                                }
                                else
                                {
                                    WriteLogFile("<!===================== Start Insert NewEmployee Branch =====================!>");
                                    MSTEmployee NewEmployee = new MSTEmployee();
                                    insert++;
                                    NewEmployee.EmployeeCode = item[1].Trim();
                                    NewEmployee.Username = item[1].Trim();
                                    NewEmployee.NameTh = !string.IsNullOrEmpty(item[4]) ? $"{item[4]} {item[3]}" : item[3];
                                    NewEmployee.NameEn = item[1].Trim() + " " + NewEmployee.NameTh;
                                    NewEmployee.Email = item[11];
                                    NewEmployee.IsActive = true;
                                    NewEmployee.AccountId = 1;
                                    NewEmployee.ModifiedDate = DateTime.Now;
                                    NewEmployee.ModifiedBy = "AdminC";
                                    NewEmployee.CreatedDate = DateTime.Now;
                                    NewEmployee.CreatedBy = "AdminC";
                                    if (!string.IsNullOrEmpty(item[30]))
                                    {
                                        var CurrentDep = lstDepartment_update.FindAll(a => a.DepartmentCode.Replace(" ", "") == item[30].Replace(" ", "")).ToList();
                                        if (CurrentDep.Count > 0)
                                        {
                                            //InsertDepartment
                                            NewEmployee.DepartmentId = CurrentDep.First().DepartmentId;
                                            if (!string.IsNullOrEmpty(item[32]))
                                            {
                                                var ParentId = lstDepartment_update.FindAll(a => a.DepartmentCode.Replace(" ", "") == item[32].Replace(" ", "")).FirstOrDefault();
                                                if (ParentId != null)
                                                {
                                                    UpdateDepartment(CurrentDep.FirstOrDefault(), db, ParentId);
                                                    MSTDepartment department = lstDepartment_update.Where(x => x.DepartmentId == ParentId.DepartmentId).FirstOrDefault();
                                                    department.ModifiedDate = DateTime.Now;
                                                    department.IsActive = true;
                                                }
                                                else
                                                {
                                                    WriteLogFile("!! Not have ParentId: " + item[32]);
                                                }
                                            }
                                        }
                                    }
                                    NewEmployee.PositionId = 705;
                                    db.MSTEmployees.InsertOnSubmit(NewEmployee);
                                    db.SubmitChanges();
                                    WriteLogFile("EmployeeId : " + NewEmployee.EmployeeId);
                                    WriteLogFile("PositionId : " + NewEmployee.PositionId);
                                    WriteLogFile("DepartmentId : " + NewEmployee.DepartmentId);
                                    WriteLogFile("ReportToEmpCode : " + NewEmployee.ReportToEmpCode);
                                    WriteLogFile("<!====================== End Insert NewEmployee Branch ======================!>");
                                }
                            }
                        }
                        if (Re_ReportToEmpCode.Count > 0)
                        {
                            foreach (var item in Re_ReportToEmpCode)
                            {
                                WriteLogFile("<!===================== Start ReInsert ReportToEmpCode =====================!>");
                                string EmpCode = item.Split(',')[0];
                                string EmpReport_to = item.Split(',')[1];
                                var emp = db.MSTEmployees.Where(e => e.EmployeeCode == EmpCode).FirstOrDefault();
                                var empre = db.MSTEmployees.Where(e => e.EmployeeCode == EmpReport_to).FirstOrDefault();
                                if (emp != null && empre != null)
                                {
                                    emp.ReportToEmpCode = empre.EmployeeId.ToString();
                                    db.SubmitChanges();
                                    WriteLogFile("EmployeeId : " + emp.EmployeeId);
                                    WriteLogFile("PositionId : " + emp.PositionId);
                                    WriteLogFile("DepartmentId : " + emp.DepartmentId);
                                    WriteLogFile("ReportToEmpCode : " + emp.ReportToEmpCode);
                                }
                                else
                                {
                                    Console.WriteLine("Not have ReportToEmpCode : " + EmpCode + "," + EmpReport_to);
                                    WriteLogFile("Not have ReportToEmpCode : " + EmpCode + "," + EmpReport_to);
                                }
                                WriteLogFile("<!===================== End ReInsert ReportToEmpCode =====================!>");
                            }
                        }
                    }
                    WriteLogFile("UPDATE : " + update);
                    WriteLogFile("INSERT : " + insert);
                    Console.WriteLine("UPDATE : " + update);
                    Console.WriteLine("INSERT : " + insert);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("GetData : " + ex);
                WriteLogFile("GetData : " + ex);
            }
        }
        public static int InsertNewPosition(string Pname, string sPositionLevelld)
        {
            DataClassesWolfDataContext db = new DataClassesWolfDataContext(Connectionstring);
            MSTPosition NewPosition = new MSTPosition();
            try
            {
                if (db.Connection.State == ConnectionState.Open)
                {
                    db.Connection.Close();
                    db.Connection.Open();
                }
                else
                {
                    db.Connection.Open();
                    if (!string.IsNullOrEmpty(sPositionLevelld) && !string.IsNullOrEmpty(Pname))
                    {
                        string PositionLevelld = Regex.Replace(sPositionLevelld, @"\D", "");
                        NewPosition.NameTh = Pname.Trim();
                        NewPosition.NameEn = Pname.Trim();
                        NewPosition.PositionLevelId = Convert.ToInt32(PositionLevelld);
                        NewPosition.AccountId = 1;
                        NewPosition.IsActive = true;
                        NewPosition.ModifiedDate = DateTime.Now;
                        NewPosition.ModifiedBy = "AdminC";
                        NewPosition.CreatedDate = DateTime.Now;
                        NewPosition.CreatedBy = "AdminC";
                        db.MSTPositions.InsertOnSubmit(NewPosition);
                        db.SubmitChanges();

                    }
                    else
                    {
                        Console.WriteLine("Not have PositionLevel : " + Pname.Trim());
                        WriteLogFile("Not have PositionLevel : " + Pname.Trim());
                        return 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error InsertNewPosition: " + ex.Message);
                WriteLogFile("An error InsertNewPosition: " + ex.Message);
            }
            return NewPosition.PositionId;
        }
        public static void UpdateDepartment(MSTDepartment CurrentDep, DataClassesWolfDataContext db, MSTDepartment ParentId)
        {
            try
            {
                CurrentDep.ParentId = ParentId.DepartmentId;
                CurrentDep.ModifiedDate = DateTime.Now;
                CurrentDep.IsActive = true;
                db.SubmitChanges();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error UpdateDepartment: " + ex.Message);
                WriteLogFile("An error UpdateDepartment: " + ex.Message);
            }
        }
        public static string _LogFile = ConfigurationManager.AppSettings["LogFile"];
        public static void WriteLogFile(String iText)
        {

            String LogFilePath = String.Format("{0}{1}_Log.txt", _LogFile, DateTime.Now.ToString("yyyyMMdd"));

            try
            {
                using (System.IO.StreamWriter outfile = new System.IO.StreamWriter(LogFilePath, true))
                {
                    System.Text.StringBuilder sbLog = new System.Text.StringBuilder();

                    String[] ListText = iText.Split('|').ToArray();

                    foreach (String s in ListText)
                    {
                        sbLog.AppendLine($"{DateTime.Now.ToString("[HH:mm:ss]")} - {s}");
                    }

                    outfile.WriteLine(sbLog.ToString());
                }
            }
            catch { }
        }
        public static bool CheckFile()
        {
            if (File.Exists(filename))
            {
                return true;
            }
            return false;
        }
        public static void MoveFile()
        {
            if (!Directory.Exists(Pathmove))
            {
                Directory.CreateDirectory(Pathmove);
            }
            try
            {
                string destinationFilePath = string.Empty;
                destinationFilePath = Path.Combine(Pathmove, $"EmployeeISO-Completed-{DateTime.Now.ToString("dd-MM-yyyy")}.XLSX");
                File.Move(filename, destinationFilePath);
                WriteLogFile("MoveFile Success");

            }
            catch (Exception ex)
            {
                WriteLogFile("Error MoveFile : " + ex.Message);
            }
        }
        public static void EmployeeIsActive()
        {
            DateTime Date = DateTime.Now;
            DataClassesWolfDataContext db = new DataClassesWolfDataContext(Connectionstring);
            try
            {
                if (db.Connection.State == ConnectionState.Open)
                {
                    db.Connection.Close();
                    db.Connection.Open();
                }
                else
                {
                    db.Connection.Open();
                    WriteLogFile("<!===================== Start Employee IsActive =====================!>");
                    List<MSTEmployee> lstemp = db.MSTEmployees.ToList();
                    lstemp = lstemp.Where(x => x.ModifiedDate.GetValueOrDefault().ToString("dd/MM/yyyy") != Date.ToString("dd/MM/yyyy") && x.EmployeeId != 1).ToList();
                    foreach (var emp in lstemp)
                    {
                        emp.IsActive = false;
                        WriteLogFile("Employeeid,EmployeeCode: " + emp.EmployeeId + "," + emp.EmployeeCode);
                    }
                    db.SubmitChanges();
                    WriteLogFile("<!===================== End Employee IsActive =====================!>");
                    WriteLogFile("<!===================== Start MSTDepartment IsActive =====================!>");
                    List<MSTDepartment> lstdep = db.MSTDepartments.ToList();
                    lstdep = lstdep.Where(x => x.ModifiedDate.GetValueOrDefault().ToString("dd/MM/yyyy") != Date.ToString("dd/MM/yyyy") && x.DepartmentId != 2).ToList();
                    foreach (var dep in lstdep)
                    {
                        dep.IsActive = false;
                        WriteLogFile("DepartmentId,DepartmentCode: " + dep.DepartmentId + "," + dep.DepartmentCode);
                    }
                    db.SubmitChanges();
                    WriteLogFile("<!===================== End 11111 IsActive =====================!>");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error EmployeeIsActive: " + ex.Message);
                WriteLogFile("An error EmployeeIsActive: " + ex.Message);
            }
        }
    }
}
